// Helpers de Office.js para Word compartidos por las herramientas.
import { builtInStyle, parseIndexList } from "./spec";

export type WordSet = "WordApi" | "WordApiDesktop";

const minorCache = new Map<WordSet, number>();

/** Versión menor más alta soportada de un requirement set (p. ej. 9 para WordApi 1.9). */
export function wordApiMinor(set: WordSet = "WordApi"): number {
  const cached = minorCache.get(set);
  if (cached !== undefined) return cached;
  let v = 0;
  try {
    for (let i = 1; i <= 30; i++) {
      if (!Office.context.requirements.isSetSupported(set, `1.${i}`)) break;
      v = i;
    }
  } catch {
    v = 0;
  }
  minorCache.set(set, v);
  return v;
}

export function wordSupports(minor: number, set: WordSet = "WordApi"): boolean {
  return wordApiMinor(set) >= minor;
}

/** Lanza un error explicativo si la versión de Word no tiene la función. */
export function requireWord(minor: number, set: WordSet, what: string): void {
  if (wordSupports(minor, set)) return;
  const desktop = set === "WordApiDesktop" ? " (Word de escritorio de Microsoft 365 actualizado)" : "";
  throw new Error(
    `${what} requiere ${set} 1.${minor}${desktop} y esta versión de Word tiene ${set} 1.${wordApiMinor(set)}. ` +
      "Resuélvelo con otra herramienta, con OOXML, con entregar_vba o con pasos manuales.",
  );
}

/** Convierte errores de Office.js en un mensaje útil para el agente. */
export async function describeWordError(err: unknown): Promise<string> {
  if (typeof OfficeExtension !== "undefined" && err instanceof OfficeExtension.Error) {
    let msg = `${err.code}: ${err.message}`;
    const loc = err.debugInfo?.errorLocation;
    if (loc) msg += ` (en ${loc})`;
    if (err.code === "ItemNotFound" || err.code === "InvalidArgument") {
      msg +=
        ". Vuelve a leer el documento con leer_documento: los números de párrafo y de tabla cambian al " +
        "insertar o borrar contenido.";
    }
    if (err.code === "ApiNotFound" || err.code === "NotImplemented") {
      msg += ". Esta versión de Word no tiene esa función: usa una alternativa (OOXML, entregar_vba o pasos manuales).";
    }
    return msg;
  }
  if (err instanceof Error) return err.message;
  return String(err);
}

export type Zona = "cuerpo" | "encabezado" | "pie";
export type HeaderType = "Primary" | "FirstPage" | "EvenPages";

/** Qué parte del documento usa una herramienta (ver el esquema `Objetivo`). */
export interface ObjetivoData {
  zona: Zona;
  seccion: number;
  tipo_encabezado: HeaderType;
  seleccion?: boolean;
  parrafos?: string;
  tabla?: number;
  texto?: string;
  ocurrencia?: number;
  mayusculas?: boolean;
  palabra_completa?: boolean;
  comodines?: boolean;
}

export interface Target {
  kind: "zona" | "parrafos" | "tabla" | "texto" | "seleccion";
  /** Cuerpo, encabezado o pie donde está el objetivo. */
  body: Word.Body;
  /** Cada unidad encontrada: el contenido de cada párrafo, la tabla, cada aparición del texto… */
  ranges: Word.Range[];
  paragraphs?: Word.Paragraph[];
  /** Números (1-based) de los párrafos cuando kind es "parrafos". */
  indices?: number[];
  table?: Word.Table;
}

const ZONE_LABEL: Record<Zona, string> = { cuerpo: "el cuerpo", encabezado: "el encabezado", pie: "el pie de página" };

export async function zoneBody(
  context: Word.RequestContext,
  obj: Pick<ObjetivoData, "zona" | "seccion" | "tipo_encabezado">,
): Promise<Word.Body> {
  if (obj.zona === "cuerpo") return context.document.body;
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();
  if (obj.seccion > sections.items.length) {
    throw new Error(`El documento tiene ${sections.items.length} sección(es); no existe la sección ${obj.seccion}.`);
  }
  const section = sections.items[obj.seccion - 1];
  return obj.zona === "encabezado" ? section.getHeader(obj.tipo_encabezado) : section.getFooter(obj.tipo_encabezado);
}

export async function resolveTarget(context: Word.RequestContext, obj: ObjetivoData): Promise<Target> {
  const body = await zoneBody(context, obj);
  const t: Target = { kind: "zona", body, ranges: [] };
  let scope: Word.Body | Word.Range = body;
  if (obj.seleccion) {
    const selection = context.document.getSelection();
    scope = selection;
    t.kind = "seleccion";
    t.ranges = [selection];
  }

  if (obj.tabla !== undefined) {
    const tables = scope.tables;
    tables.load("items/rowCount");
    await context.sync();
    if (obj.tabla > tables.items.length) {
      throw new Error(
        tables.items.length
          ? `Hay ${tables.items.length} tabla(s) en ${ZONE_LABEL[obj.zona]}; no existe la tabla ${obj.tabla}.`
          : `No hay tablas en ${ZONE_LABEL[obj.zona]}.`,
      );
    }
    t.table = tables.items[obj.tabla - 1];
    t.kind = "tabla";
    t.ranges = [t.table.getRange("Whole")];
  } else if (obj.parrafos) {
    const all = scope.paragraphs;
    all.load("items/text");
    await context.sync();
    t.indices = parseIndexList(obj.parrafos, all.items.length);
    t.paragraphs = t.indices.map((i) => all.items[i - 1]);
    t.kind = "parrafos";
    t.ranges = t.paragraphs.map((p) => p.getRange("Content"));
  }

  if (obj.texto) {
    const containers: (Word.Body | Word.Range | Word.Paragraph)[] =
      t.kind === "zona" ? [body] : (t.paragraphs ?? t.ranges);
    const options = {
      matchCase: obj.mayusculas ?? false,
      matchWholeWord: obj.palabra_completa ?? false,
      matchWildcards: obj.comodines ?? false,
    };
    const results = containers.map((c) => c.search(obj.texto as string, options));
    for (const r of results) r.load("items/text");
    await context.sync();
    let found = results.flatMap((r) => r.items);
    if (!found.length) {
      throw new Error(`No se encontró «${obj.texto}» en ${ZONE_LABEL[obj.zona]}. Revisa el texto exacto con leer_documento.`);
    }
    if (obj.ocurrencia) {
      if (obj.ocurrencia > found.length) {
        throw new Error(`«${obj.texto}» aparece ${found.length} vez/veces; no existe la ocurrencia ${obj.ocurrencia}.`);
      }
      found = [found[obj.ocurrencia - 1]];
    }
    return { kind: "texto", body, ranges: found };
  }

  if (t.kind === "zona") t.ranges = [body.getRange("Whole")];
  return t;
}

/** Párrafos que abarca el objetivo (sin repetir cuando la API permite identificarlos). */
export async function targetParagraphs(context: Word.RequestContext, t: Target): Promise<Word.Paragraph[]> {
  if (t.paragraphs) return t.paragraphs;
  const collections = t.kind === "zona" ? [t.body.paragraphs] : t.ranges.map((r) => r.paragraphs);
  const dedupe = wordSupports(6);
  for (const c of collections) c.load(dedupe ? "items/uniqueLocalId" : "items/text");
  await context.sync();
  const items = collections.flatMap((c) => c.items);
  if (!dedupe) return items;
  const seen = new Set<string>();
  return items.filter((p) => !seen.has(p.uniqueLocalId) && Boolean(seen.add(p.uniqueLocalId)));
}

/** Primer o último párrafo (o la tabla) del objetivo, para insertar contenido antes o después. */
export function edgeOf(t: Target, which: "first" | "last"): Word.Paragraph | Word.Table {
  if (t.table) return t.table;
  if (t.paragraphs) return which === "first" ? t.paragraphs[0] : t.paragraphs[t.paragraphs.length - 1];
  if (t.kind === "zona") return which === "first" ? t.body.paragraphs.getFirst() : t.body.paragraphs.getLast();
  const range = which === "first" ? t.ranges[0] : t.ranges[t.ranges.length - 1];
  return which === "first" ? range.paragraphs.getFirst() : range.paragraphs.getLast();
}

export type Ubicacion = "inicio" | "final" | "antes" | "despues" | "reemplazar";

export const isBefore = (u: Ubicacion): boolean => u === "inicio" || u === "antes";

/** Si la zona solo tiene un párrafo vacío (documento nuevo, encabezado sin contenido), lo devuelve. */
export async function isEmptyBody(context: Word.RequestContext, body: Word.Body): Promise<Word.Paragraph | null> {
  const paras = body.paragraphs;
  paras.load("items/text");
  await context.sync();
  if (paras.items.length !== 1 || paras.items[0].text !== "") return null;
  const pictures = paras.items[0].inlinePictures;
  pictures.load("items");
  await context.sync();
  return pictures.items.length === 0 ? paras.items[0] : null;
}

/**
 * Crea un párrafo vacío en la ubicación pedida (al inicio o final de la zona, o antes/después del objetivo).
 * En una zona vacía reutiliza su único párrafo para no dejar una línea en blanco.
 */
export async function newParagraph(
  context: Word.RequestContext,
  t: Target,
  ubicacion: Ubicacion,
): Promise<Word.Paragraph> {
  const before = isBefore(ubicacion);
  if (t.kind === "zona") {
    const empty = await isEmptyBody(context, t.body);
    if (empty) return empty;
    return t.body.insertParagraph("", before ? "Start" : "End");
  }
  return edgeOf(t, before ? "first" : "last").insertParagraph("", before ? "Before" : "After");
}

/** Aplica un estilo integrado (independiente del idioma) o, si no lo es, el estilo con ese nombre. */
export function applyStyle(target: { style: string }, name: string): void {
  const builtIn = builtInStyle(name);
  if (builtIn) (target as unknown as { styleBuiltIn: string }).styleBuiltIn = builtIn;
  else target.style = name;
}

/** Número (1-based) de una tabla dentro de su zona. */
export async function tableIndexOf(
  context: Word.RequestContext,
  body: Word.Body,
  table: Word.Table,
): Promise<number | undefined> {
  const tables = body.tables;
  tables.load("items/rowCount");
  await context.sync();
  const target = table.getRange("Whole");
  const relations = tables.items.map((t) => t.getRange("Whole").compareLocationWith(target));
  await context.sync();
  const i = relations.findIndex((r) => r.value === "Equal");
  return i >= 0 ? i + 1 : undefined;
}
