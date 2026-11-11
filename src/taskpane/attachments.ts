import type Anthropic from "@anthropic-ai/sdk";
import { unzipSync } from "fflate";

type ContentBlockParam = Anthropic.Beta.Messages.BetaContentBlockParam;
type ImageMediaType = "image/png" | "image/jpeg" | "image/gif" | "image/webp";

export type AttachmentKind = "pdf" | "image" | "word" | "excel" | "text";

export interface Attachment {
  name: string;
  kind: AttachmentKind;
  /** PDF e imágenes: base64 sin el prefijo data:. Word, Excel y texto: el texto extraído. */
  data: string;
  mediaType?: ImageMediaType;
  size: number;
}

export const MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024;

export const ACCEPT_ATTRIBUTE =
  ".pdf,.docx,.xlsx,.xlsm,.txt,.md,.csv,.tsv,.png,.jpg,.jpeg,.gif,.webp,application/pdf," +
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document," +
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/plain,image/*";

const IMAGE_TYPES: Record<string, ImageMediaType> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
};

/** Clasifica un archivo por extensión (algunos sistemas no informan el tipo MIME). */
export function classifyFile(name: string, type: string): { kind: AttachmentKind; mediaType?: ImageMediaType } | null {
  const ext = name.split(".").pop()?.toLowerCase() ?? "";
  if (ext === "pdf" || type === "application/pdf") return { kind: "pdf" };
  if (ext === "docx") return { kind: "word" };
  if (ext === "xlsx" || ext === "xlsm") return { kind: "excel" };
  if (["txt", "md", "csv", "tsv"].includes(ext) || type === "text/plain") return { kind: "text" };
  const image = IMAGE_TYPES[ext] ?? (Object.values(IMAGE_TYPES) as string[]).find((t) => t === type);
  if (image) return { kind: "image", mediaType: image as ImageMediaType };
  return null;
}

function decodeEntities(text: string): string {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, h: string) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d: string) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, "&");
}

/** Convierte el XML de word/document.xml en texto plano (párrafos, tablas, tabulaciones). */
export function wordXmlToText(xml: string): string {
  const text = xml
    .replace(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/g, "")
    .replace(/<w:delText[^>]*>[\s\S]*?<\/w:delText>/g, "")
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<w:(br|cr)[^>]*\/>/g, "\n")
    .replace(/<\/w:tc>/g, " | ")
    .replace(/<\/w:tr>/g, "\n")
    .replace(/<\/w:p>/g, "\n")
    .replace(/<[^>]+>/g, "");
  return decodeEntities(text)
    .split("\n")
    .map((line) => line.replace(/\s*\|\s*$/, "").trimEnd())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export function docxToText(bytes: Uint8Array): string {
  const files = unzipSync(bytes, { filter: (f) => f.name === "word/document.xml" });
  const xml = files["word/document.xml"];
  if (!xml) throw new Error("No parece un documento de Word válido (.docx).");
  return wordXmlToText(new TextDecoder("utf-8").decode(xml));
}

/**
 * Texto de un archivo .txt/.csv: UTF-8 (con o sin BOM), UTF-16 con BOM o, si no es UTF-8 válido,
 * Windows-1252 (Bloc de notas antiguo, CSV de Excel en Windows), para no perder tildes ni eñes.
 */
export function decodeText(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le").decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder("utf-16be").decode(bytes);
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder("windows-1252").decode(bytes);
  }
}

const XLSX_MAX_CHARS = 300_000;
const DATE_FORMAT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

function columnIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function serialToIso(serial: number): string {
  const ms = Math.round((serial - 25569) * 86_400_000);
  const iso = new Date(ms).toISOString();
  return serial % 1 === 0 ? iso.slice(0, 10) : iso.slice(0, 16).replace("T", " ");
}

/** Índices de estilo (cellXfs) cuyo formato numérico es de fecha u hora. */
function dateStyles(stylesXml: string): Set<number> {
  const custom = new Map<number, string>();
  for (const m of stylesXml.matchAll(/<numFmt [^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g)) {
    custom.set(Number(m[1]), decodeEntities(m[2]));
  }
  const isDateCode = (code: string) => /[dmyhs]/i.test(code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, ""));
  const result = new Set<number>();
  const xfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? "";
  [...xfs.matchAll(/<xf [^>]*?numFmtId="(\d+)"/g)].forEach((m, i) => {
    const id = Number(m[1]);
    if (DATE_FORMAT_IDS.has(id) || (custom.has(id) && isDateCode(custom.get(id)!))) result.add(i);
  });
  return result;
}

/** Convierte un libro .xlsx en texto: una sección por hoja, filas numeradas y celdas separadas por tabulaciones. */
export function xlsxToText(bytes: Uint8Array): string {
  const files = unzipSync(bytes, {
    filter: (f) => /^xl\/(workbook\.xml|_rels\/workbook\.xml\.rels|sharedStrings\.xml|styles\.xml|worksheets\/[^/]+\.xml)$/.test(f.name),
  });
  const read = (name: string) => (files[name] ? new TextDecoder("utf-8").decode(files[name]) : "");
  const workbook = read("xl/workbook.xml");
  if (!workbook) throw new Error("No parece un libro de Excel válido (.xlsx).");
  const shared = [...read("xl/sharedStrings.xml").matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) =>
    decodeEntities([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join("")),
  );
  const dates = dateStyles(read("xl/styles.xml"));
  const rels = new Map<string, string>();
  for (const m of read("xl/_rels/workbook.xml.rels").matchAll(/<Relationship [^>]*>/g)) {
    const id = /Id="([^"]+)"/.exec(m[0])?.[1];
    const target = /Target="([^"]+)"/.exec(m[0])?.[1];
    if (id && target) rels.set(id, target.replace(/^\/?(xl\/)?/, "xl/"));
  }
  const out: string[] = [];
  for (const s of workbook.matchAll(/<sheet [^>]*>/g)) {
    const name = decodeEntities(/name="([^"]*)"/.exec(s[0])?.[1] ?? "");
    const rid = /r:id="([^"]+)"/.exec(s[0])?.[1] ?? "";
    const xml = read(rels.get(rid) ?? "");
    const lines: string[] = [];
    for (const row of xml.matchAll(/<row [^>]*?r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
      const cells: string[] = [];
      for (const c of row[2].matchAll(/<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const attrs = c[2];
        const inner = c[3] ?? "";
        const type = /t="(\w+)"/.exec(attrs)?.[1];
        const raw = /<v>([\s\S]*?)<\/v>/.exec(inner)?.[1];
        let value = "";
        if (type === "s" && raw !== undefined) value = shared[Number(raw)] ?? "";
        else if (type === "inlineStr") value = decodeEntities([...inner.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join(""));
        else if (type === "b") value = raw === "1" ? "VERDADERO" : "FALSO";
        else if (raw !== undefined) {
          const style = Number(/s="(\d+)"/.exec(attrs)?.[1] ?? -1);
          value = type === undefined || type === "n" ? (dates.has(style) ? serialToIso(Number(raw)) : raw) : decodeEntities(raw);
        }
        if (value !== "") cells[columnIndex(c[1])] = value.replace(/[\t\r\n]+/g, " ");
      }
      if (cells.length) lines.push(`${row[1]}\t${Array.from(cells, (v) => v ?? "").join("\t")}`);
    }
    out.push(`### Hoja «${name}» (${lines.length} filas con datos; la primera columna es el número de fila)`, ...lines, "");
  }
  const text = out.join("\n").trim();
  return text.length > XLSX_MAX_CHARS ? `${text.slice(0, XLSX_MAX_CHARS)}\n[… recortado: el libro es más grande]` : text;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

/** Lee un archivo y lo deja listo para enviarlo a Claude. */
export async function readAttachment(file: File): Promise<Attachment | null> {
  const info = classifyFile(file.name, file.type);
  if (!info) return null;
  const bytes = new Uint8Array(await file.arrayBuffer());
  const base = { name: file.name, kind: info.kind, size: file.size };
  switch (info.kind) {
    case "pdf":
      return { ...base, data: bytesToBase64(bytes) };
    case "image":
      return { ...base, data: bytesToBase64(bytes), mediaType: info.mediaType };
    case "word":
      return { ...base, data: docxToText(bytes) };
    case "excel":
      return { ...base, data: xlsxToText(bytes) };
    case "text":
      return { ...base, data: decodeText(bytes) };
  }
}

const TEXT_CONTEXT: Partial<Record<AttachmentKind, string>> = {
  word: "Texto extraído de un .docx; las imágenes y la numeración automática no se conservan.",
  excel:
    "Valores de un libro de Excel adjunto (no es el libro abierto): una sección por hoja, la primera columna es el " +
    "número de fila y las celdas van separadas por tabulaciones; las fechas están en formato ISO.",
};

export const defaultRequest = (doc = "libro"): string =>
  `Resuelve en este ${doc} todos los requerimientos del archivo adjunto.`;

/** Arma el contenido del mensaje: primero los adjuntos, luego el texto. */
export function buildUserContent(text: string, attachments: Attachment[], doc = "libro"): ContentBlockParam[] {
  const blocks: ContentBlockParam[] = attachments.flatMap((a): ContentBlockParam[] => {
    switch (a.kind) {
      case "pdf":
        return [{ type: "document", source: { type: "base64", media_type: "application/pdf", data: a.data }, title: a.name }];
      case "image":
        // El nombre permite insertar la imagen en el archivo (insertar_imagen).
        return [
          { type: "text", text: `Imagen adjunta «${a.name}»:` },
          { type: "image", source: { type: "base64", media_type: a.mediaType ?? "image/png", data: a.data } },
        ];
      case "word":
      case "excel":
      case "text":
        return [
          {
            type: "document",
            source: { type: "text", media_type: "text/plain", data: a.data || "(documento vacío)" },
            title: a.name,
            ...(TEXT_CONTEXT[a.kind] ? { context: TEXT_CONTEXT[a.kind] } : {}),
          },
        ];
    }
  });
  blocks.push({ type: "text", text: text.trim() || defaultRequest(doc) });
  return blocks;
}

export function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}
