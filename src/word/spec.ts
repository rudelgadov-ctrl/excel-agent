// Helpers puros para las herramientas de Word (sin Office.js, testeables).

/**
 * Convierte una lista de números 1-based ("3", "2-5", "1, 4, 7-9") en índices ordenados y sin repetir.
 * Lanza un error si algún número está fuera de 1..max.
 */
export function parseIndexList(spec: string, max: number, what = "párrafo"): number[] {
  const out = new Set<number>();
  for (const raw of spec.split(",")) {
    const part = raw.trim();
    if (!part) continue;
    const m = part.match(/^(\d+)\s*(?:-\s*(\d+))?$/);
    if (!m) throw new Error(`Número de ${what} inválido: "${part}". Usa "3", "2-5" o "1,4,7-9".`);
    const from = Number(m[1]);
    const to = m[2] === undefined ? from : Number(m[2]);
    if (from < 1 || to < from) throw new Error(`Intervalo de ${what} inválido: "${part}".`);
    if (to > max) throw new Error(`El ${what} ${to} no existe: hay ${max}. Vuelve a leer el documento.`);
    for (let i = from; i <= to; i++) out.add(i);
  }
  if (out.size === 0) throw new Error(`No se indicó ningún ${what}.`);
  return [...out].sort((a, b) => a - b);
}

/** Agrupa índices ordenados en tramos consecutivos [inicio, cantidad], del último al primero (para borrar). */
export function runsForDeletion(indices: number[]): [number, number][] {
  const runs: [number, number][] = [];
  for (const i of [...indices].sort((a, b) => a - b)) {
    const last = runs[runs.length - 1];
    if (last && last[0] + last[1] === i) last[1]++;
    else runs.push([i, 1]);
  }
  return runs.reverse();
}

const PT_PER_CM = 72 / 2.54;

export function cmToPt(cm: number): number {
  return Math.round(cm * PT_PER_CM * 100) / 100;
}

export function ptToCm(pt: number): number {
  return Math.round((pt / PT_PER_CM) * 100) / 100;
}

/** En la API, el interlineado múltiple se expresa en puntos: sencillo = 12. */
export function lineSpacingPt(multiple: number): number {
  return Math.round(multiple * 12 * 100) / 100;
}

export function lineSpacingMultiple(pt: number): number {
  return Math.round((pt / 12) * 100) / 100;
}

const TABLE_STYLE_BASES = [
  "GridTable1Light",
  "GridTable2",
  "GridTable3",
  "GridTable4",
  "GridTable5Dark",
  "GridTable6Colorful",
  "GridTable7Colorful",
  "ListTable1Light",
  "ListTable2",
  "ListTable3",
  "ListTable4",
  "ListTable5Dark",
  "ListTable6Colorful",
  "ListTable7Colorful",
];

/** Valores de Word.BuiltInStyleName: funcionan igual en Word en cualquier idioma. */
export const BUILT_IN_STYLES: readonly string[] = [
  "Normal",
  ...Array.from({ length: 9 }, (_, i) => `Heading${i + 1}`),
  ...Array.from({ length: 9 }, (_, i) => `Toc${i + 1}`),
  "FootnoteText",
  "Header",
  "Footer",
  "Caption",
  "FootnoteReference",
  "EndnoteReference",
  "EndnoteText",
  "Title",
  "Subtitle",
  "Hyperlink",
  "Strong",
  "Emphasis",
  "NoSpacing",
  "ListParagraph",
  "Quote",
  "IntenseQuote",
  "SubtleEmphasis",
  "IntenseEmphasis",
  "SubtleReference",
  "IntenseReference",
  "BookTitle",
  "Bibliography",
  "TocHeading",
  "TableGrid",
  "PlainTable1",
  "PlainTable2",
  "PlainTable3",
  "PlainTable4",
  "PlainTable5",
  "TableGridLight",
  ...TABLE_STYLE_BASES.flatMap((b) => [b, ...Array.from({ length: 6 }, (_, i) => `${b}_Accent${i + 1}`)]),
];

// Nombres que muestra Word en español para los estilos integrados más usados.
const SPANISH_DISPLAY: Record<string, string> = {
  Normal: "Normal",
  Title: "Título",
  Subtitle: "Subtítulo",
  NoSpacing: "Sin espaciado",
  ListParagraph: "Párrafo de lista",
  Quote: "Cita",
  IntenseQuote: "Cita destacada",
  Emphasis: "Énfasis",
  SubtleEmphasis: "Énfasis sutil",
  IntenseEmphasis: "Énfasis intenso",
  Strong: "Texto en negrita",
  SubtleReference: "Referencia sutil",
  IntenseReference: "Referencia intensa",
  BookTitle: "Título del libro",
  Caption: "Descripción",
  Header: "Encabezado",
  Footer: "Pie de página",
  FootnoteText: "Texto nota pie",
  Hyperlink: "Hipervínculo",
  Bibliography: "Bibliografía",
  TocHeading: "Título de TDC",
  TableGrid: "Tabla con cuadrícula",
  ...Object.fromEntries(Array.from({ length: 9 }, (_, i) => [`Heading${i + 1}`, `Título ${i + 1}`])),
  ...Object.fromEntries(Array.from({ length: 9 }, (_, i) => [`Toc${i + 1}`, `TDC ${i + 1}`])),
};

function normalizeName(name: string): string {
  return name
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\s_.-]+/g, "")
    .toLowerCase();
}

const BUILT_IN_BY_KEY = new Map<string, string>([
  ...Object.entries(SPANISH_DISPLAY).map(([builtIn, local]): [string, string] => [normalizeName(local), builtIn]),
  ...BUILT_IN_STYLES.map((s): [string, string] => [normalizeName(s), s]),
]);

/**
 * Devuelve el nombre integrado (Word.BuiltInStyleName) si `name` corresponde a uno, ya sea en inglés
 * ("Heading 1", "heading1", "GridTable4 Accent1") o con el nombre que muestra Word en español ("Título 1").
 */
export function builtInStyle(name: string): string | undefined {
  return BUILT_IN_BY_KEY.get(normalizeName(name));
}

/** Nombres con los que se puede buscar un estilo en la colección de estilos del documento. */
export function styleNameCandidates(name: string): string[] {
  const builtIn = builtInStyle(name);
  if (!builtIn) return [name];
  const english = builtIn.replace(/_/g, " ").replace(/([a-z])([A-Z0-9])/g, "$1 $2");
  return [...new Set([name, english, SPANISH_DISPLAY[builtIn]].filter((n): n is string => Boolean(n)))];
}

/** Valores de Word.FieldType. */
export const FIELD_TYPES: readonly string[] = [
  "Addin", "AddressBlock", "Advance", "Ask", "Author", "AutoText", "AutoTextList", "BarCode", "Bibliography",
  "BidiOutline", "Citation", "Comments", "Compare", "CreateDate", "Data", "Database", "Date", "DisplayBarcode",
  "DocProperty", "DocVariable", "EditTime", "Embedded", "EQ", "FileName", "FileSize", "FillIn", "FormCheckbox",
  "FormDropdown", "FormText", "GotoButton", "GreetingLine", "Hyperlink", "If", "Import", "Include",
  "IncludePicture", "IncludeText", "Index", "Info", "Keywords", "LastSavedBy", "Link", "ListNum", "MacroButton",
  "MergeBarcode", "MergeField", "MergeRec", "MergeSeq", "Next", "NextIf", "NoteRef", "NumChars", "NumPages",
  "NumWords", "OCX", "Page", "PageRef", "Print", "PrintDate", "Private", "Quote", "RD", "Ref", "RevNum",
  "SaveDate", "Section", "SectionPages", "Seq", "Set", "Shape", "SkipIf", "StyleRef", "Subject", "Subscriber",
  "Symbol", "TA", "TC", "Template", "Time", "Title", "TOA", "TOC", "UserAddress", "UserInitials", "UserName", "XE",
];

const FIELD_BY_KEY = new Map(FIELD_TYPES.map((f) => [f.toUpperCase(), f]));

/**
 * Separa un código de campo ("TOC \o "1-3" \h") en el tipo de Word.FieldType y sus modificadores.
 * `type` queda indefinido si la palabra clave no es un tipo conocido (p. ej. "= SUM(ABOVE)").
 */
export function splitFieldCode(code: string): { type?: string; switches: string } {
  const trimmed = code.trim();
  const m = trimmed.match(/^([A-Za-z]+)\b\s*([\s\S]*)$/);
  if (!m) return { switches: trimmed };
  const type = FIELD_BY_KEY.get(m[1].toUpperCase());
  return type ? { type, switches: m[2].trim() } : { switches: trimmed };
}
