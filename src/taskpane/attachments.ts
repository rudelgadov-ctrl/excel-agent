import type Anthropic from "@anthropic-ai/sdk";
import { unzipSync } from "fflate";

type ContentBlockParam = Anthropic.Beta.Messages.BetaContentBlockParam;
type ImageMediaType = "image/png" | "image/jpeg" | "image/gif" | "image/webp";

export type AttachmentKind = "pdf" | "image" | "word" | "text";

export interface Attachment {
  name: string;
  kind: AttachmentKind;
  /** PDF e imágenes: base64 sin el prefijo data:. Word y texto: el texto extraído. */
  data: string;
  mediaType?: ImageMediaType;
  size: number;
}

export const MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024;

export const ACCEPT_ATTRIBUTE =
  ".pdf,.docx,.txt,.md,.csv,.png,.jpg,.jpeg,.gif,.webp,application/pdf," +
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document,text/plain,image/*";

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
  if (["txt", "md", "csv"].includes(ext) || type === "text/plain") return { kind: "text" };
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
    case "text":
      return { ...base, data: new TextDecoder("utf-8").decode(bytes) };
  }
}

export const DEFAULT_REQUEST = "Resuelve en este libro todos los requerimientos del archivo adjunto.";

/** Arma el contenido del mensaje: primero los adjuntos, luego el texto. */
export function buildUserContent(text: string, attachments: Attachment[]): ContentBlockParam[] {
  const blocks: ContentBlockParam[] = attachments.map((a): ContentBlockParam => {
    switch (a.kind) {
      case "pdf":
        return { type: "document", source: { type: "base64", media_type: "application/pdf", data: a.data }, title: a.name };
      case "image":
        return { type: "image", source: { type: "base64", media_type: a.mediaType ?? "image/png", data: a.data } };
      case "word":
      case "text":
        return {
          type: "document",
          source: { type: "text", media_type: "text/plain", data: a.data || "(documento vacío)" },
          title: a.name,
          ...(a.kind === "word"
            ? { context: "Texto extraído de un .docx; las imágenes y la numeración automática no se conservan." }
            : {}),
        };
    }
  });
  blocks.push({ type: "text", text: text.trim() || DEFAULT_REQUEST });
  return blocks;
}

export function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}
