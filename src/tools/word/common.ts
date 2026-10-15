// Esquemas y helpers compartidos por las herramientas de Word.
import { z } from "zod";
import { wordSupports, type ObjetivoData } from "../../word/office";
import { cmToPt, lineSpacingPt } from "../../word/spec";

export const Objetivo = z
  .object({
    zona: z
      .enum(["cuerpo", "encabezado", "pie"])
      .default("cuerpo")
      .describe("Cuerpo del documento, encabezado o pie de página."),
    seccion: z.number().int().min(1).default(1).describe("Sección (1 = primera) cuando zona es encabezado o pie."),
    tipo_encabezado: z
      .enum(["Primary", "FirstPage", "EvenPages"])
      .default("Primary")
      .describe("Primary = todas las páginas (o las impares); FirstPage = primera página diferente; EvenPages = pares."),
    seleccion: z.boolean().optional().describe("true = lo que el usuario tiene seleccionado en Word."),
    parrafos: z
      .string()
      .optional()
      .describe('Números de párrafo según leer_documento (de la zona o de la selección): "5", "3-8" o "1,4,7-9".'),
    tabla: z.number().int().min(1).optional().describe("Número de tabla (1 = primera de la zona)."),
    texto: z
      .string()
      .min(1)
      .max(255)
      .optional()
      .describe("Texto a buscar dentro de lo anterior (o de toda la zona). Sin `ocurrencia` aplica a todas las apariciones."),
    ocurrencia: z.number().int().min(1).optional().describe("Solo la n-ésima aparición de `texto`."),
    mayusculas: z.boolean().optional().describe("Distinguir mayúsculas y minúsculas al buscar."),
    palabra_completa: z.boolean().optional(),
    comodines: z.boolean().optional().describe("Usar comodines de Word en `texto` (?, *, [ ], <, >, @)."),
  })
  .describe("Qué parte del documento. Sin parrafos, tabla, texto ni seleccion = toda la zona.");

export const DEFAULT_OBJETIVO: ObjetivoData = Objetivo.parse({});

export const Ubicacion = z
  .enum(["inicio", "final", "antes", "despues", "reemplazar"])
  .describe(
    "inicio/final = dentro del objetivo (al principio o al final de la zona o del párrafo); antes/despues = " +
      "fuera del objetivo, como párrafo o texto contiguo; reemplazar = en lugar del objetivo.",
  );

export const ALINEACION = { izquierda: "Left", centro: "Centered", derecha: "Right", justificado: "Justified" } as const;
export const ALINEACION_ES: Record<string, string> = {
  Left: "izquierda",
  Centered: "centro",
  Right: "derecha",
  Justified: "justificado",
};
export const Alineacion = z.enum(["izquierda", "centro", "derecha", "justificado"]);

const SUBRAYADO = {
  ninguno: "None",
  simple: "Single",
  doble: "Double",
  grueso: "Thick",
  punteado: "Dotted",
  discontinuo: "DashLine",
  ondulado: "Wave",
  palabras: "Word",
} as const;

/** Formato de fuente básico (también lo usan los estilos y las tablas). */
export const fuenteShape = {
  fuente: z.string().optional().describe('Nombre de la fuente, p. ej. "Arial".'),
  tamano: z.number().positive().optional().describe("Tamaño en puntos."),
  negrita: z.boolean().optional(),
  cursiva: z.boolean().optional(),
  subrayado: z.enum(["ninguno", "simple", "doble", "grueso", "punteado", "discontinuo", "ondulado", "palabras"]).optional(),
  color: z.string().optional().describe('Color del texto: "#1F4E79" o nombre en inglés ("red").'),
};

export const fuenteExtraShape = {
  resaltado: z
    .string()
    .nullable()
    .optional()
    .describe('Color de resaltado ("yellow", "#FFFF00", "turquoise"…); null lo quita.'),
  tachado: z.boolean().optional(),
  doble_tachado: z.boolean().optional(),
  superindice: z.boolean().optional(),
  subindice: z.boolean().optional(),
  mayusculas: z.boolean().optional().describe("Formato Todo mayúsculas (no cambia el texto)."),
  versalitas: z.boolean().optional(),
  espaciado_pt: z.number().optional().describe("Espaciado entre caracteres en puntos (positivo = expandido)."),
};

const FuenteBase = z.object(fuenteShape);
const FuenteCompleta = z.object({ ...fuenteShape, ...fuenteExtraShape });
type FuenteInput = z.infer<typeof FuenteBase> & Partial<z.infer<typeof FuenteCompleta>>;

/** Aplica el formato de fuente. Lo que esta versión de Word no admite se reporta en `warnings`. */
export function applyFont(font: Word.Font, p: FuenteInput, warnings: string[]): void {
  if (p.fuente !== undefined) font.name = p.fuente;
  if (p.tamano !== undefined) font.size = p.tamano;
  if (p.negrita !== undefined) font.bold = p.negrita;
  if (p.cursiva !== undefined) font.italic = p.cursiva;
  if (p.subrayado !== undefined) font.underline = SUBRAYADO[p.subrayado];
  if (p.color !== undefined) font.color = p.color;
  if (p.resaltado !== undefined) font.highlightColor = p.resaltado as string;
  if (p.tachado !== undefined) font.strikeThrough = p.tachado;
  if (p.doble_tachado !== undefined) font.doubleStrikeThrough = p.doble_tachado;
  if (p.superindice !== undefined) font.superscript = p.superindice;
  if (p.subindice !== undefined) font.subscript = p.subindice;
  const desktop = [
    ["mayusculas", "allCaps"],
    ["versalitas", "smallCaps"],
    ["espaciado_pt", "spacing"],
  ] as const;
  for (const [key, prop] of desktop) {
    if (p[key] === undefined) continue;
    if (wordSupports(3, "WordApiDesktop")) (font as unknown as Record<string, unknown>)[prop] = p[key];
    else warnings.push(`${key} requiere WordApiDesktop 1.3; aplícalo a mano (Inicio > Fuente) o con VBA.`);
  }
}

/** Formato de párrafo (también lo usan los estilos). */
export const parrafoShape = {
  alineacion: Alineacion.optional(),
  sangria_izquierda_cm: z.number().optional(),
  sangria_derecha_cm: z.number().optional(),
  primera_linea_cm: z
    .number()
    .optional()
    .describe("Sangría de primera línea; negativa = sangría francesa (pon también sangria_izquierda_cm con el mismo valor positivo)."),
  espacio_antes_pt: z.number().min(0).optional(),
  espacio_despues_pt: z.number().min(0).optional(),
  interlineado: z.number().positive().max(10).optional().describe("Múltiplo: 1 = sencillo, 1.15, 1.5, 2 = doble."),
};

const ParrafoBase = z.object(parrafoShape);
type ParrafoInput = z.infer<typeof ParrafoBase>;

interface ParagraphLike {
  alignment: unknown;
  leftIndent: number;
  rightIndent: number;
  firstLineIndent: number;
  spaceBefore: number;
  spaceAfter: number;
  lineSpacing: number;
}

export function applyParagraphFormat(target: ParagraphLike, p: ParrafoInput): void {
  if (p.alineacion !== undefined) target.alignment = ALINEACION[p.alineacion];
  if (p.sangria_izquierda_cm !== undefined) target.leftIndent = cmToPt(p.sangria_izquierda_cm);
  if (p.sangria_derecha_cm !== undefined) target.rightIndent = cmToPt(p.sangria_derecha_cm);
  if (p.primera_linea_cm !== undefined) target.firstLineIndent = cmToPt(p.primera_linea_cm);
  if (p.espacio_antes_pt !== undefined) target.spaceBefore = p.espacio_antes_pt;
  if (p.espacio_despues_pt !== undefined) target.spaceAfter = p.espacio_despues_pt;
  if (p.interlineado !== undefined) target.lineSpacing = lineSpacingPt(p.interlineado);
}

export const hasAny = (obj: Record<string, unknown>, keys: readonly string[]): boolean =>
  keys.some((k) => obj[k] !== undefined);

export const VERIFY_NOTE =
  "Revisa el resultado con leer_documento: Word puede dejar un párrafo vacío de más o heredar el estilo del " +
  "párrafo vecino (corrígelo con eliminar_contenido solo_vacios=true o formato_parrafo).";
