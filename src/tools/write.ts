import { z } from "zod";
import { fillMatrix } from "../excel/address";
import { cellAddress, getRange, supports } from "../excel/office";
import { dateSeries, linearSeries, repeatPattern, type DateUnit } from "../excel/series";
import { findErrors, splitRows } from "../excel/text";
import { defineTool } from "./types";

const CellValue = z.union([z.string(), z.number(), z.boolean(), z.null()]);
type Cell = z.infer<typeof CellValue>;

/** Si `range` es una sola celda, lo expande al tamaño de los datos. */
async function sizedTarget(
  context: Excel.RequestContext,
  ref: string,
  rows: number,
  cols: number,
): Promise<Excel.Range> {
  const range = getRange(context, ref);
  range.load("rowCount,columnCount");
  await context.sync();
  if (range.rowCount === 1 && range.columnCount === 1) {
    return range.getResizedRange(rows - 1, cols - 1);
  }
  if (range.rowCount !== rows || range.columnCount !== cols) {
    throw new Error(
      `El rango tiene ${range.rowCount}x${range.columnCount} celdas pero los datos son ${rows}x${cols}. ` +
        "Usa solo la celda superior izquierda o un rango del mismo tamaño.",
    );
  }
  return range;
}

/** Lee lo escrito y reporta celdas con error (#NAME?, #VALUE!, …). */
async function checkErrors(context: Excel.RequestContext, range: Excel.Range) {
  range.load("address,values,rowIndex,columnIndex");
  await context.sync();
  const errores = findErrors(range.values, range.rowIndex, range.columnIndex, cellAddress);
  return { rango: range.address, errores: errores.length ? errores.slice(0, 50) : undefined };
}

export const escribirRango = defineTool({
  name: "escribir_rango",
  description:
    "Escribe valores y/o fórmulas en un rango. `datos` es una matriz de filas. Los textos que empiezan con " +
    "\"=\" son fórmulas y deben ir en INGLÉS con coma como separador (=PMT(B3/12,B4,-B2)). null deja la " +
    "celda sin cambios; \"\" la borra. Si `rango` es una sola celda, se usa como esquina superior izquierda. " +
    "Devuelve las celdas que quedaron con error para que las corrijas.",
  schema: z.object({
    rango: z.string().describe("Celda inicial o rango destino, p. ej. \"'#1 - Amortización'!A8\"."),
    datos: z.array(z.array(CellValue)).min(1).describe("Matriz de filas x columnas."),
    formato_numero: z
      .string()
      .optional()
      .describe("Formato numérico opcional para todo el rango (código en-US, p. ej. \"0.00%\")."),
  }),
  writes: true,
  async run({ rango, datos, formato_numero }) {
    const cols = Math.max(...datos.map((r) => r.length));
    const matrix: Cell[][] = datos.map((r) => [...r, ...Array<Cell>(cols - r.length).fill(null)]);
    return Excel.run(async (context) => {
      const target = await sizedTarget(context, rango, matrix.length, cols);
      target.formulas = matrix as unknown as string[][];
      if (formato_numero) target.numberFormat = fillMatrix(matrix.length, cols, formato_numero);
      await context.sync();
      return checkErrors(context, target);
    });
  },
});

const COPY_TYPES = { todo: "All", valores: "Values", formulas: "Formulas", formatos: "Formats" } as const;

export const copiarRango = defineTool({
  name: "copiar_rango",
  description:
    "Copia un rango a otro lugar (también entre hojas), como Copiar/Pegar o Pegado especial. " +
    "`destino` puede ser solo la celda superior izquierda. Las referencias relativas de las fórmulas se ajustan.",
  schema: z.object({
    origen: z.string(),
    destino: z.string(),
    tipo: z.enum(["todo", "valores", "formulas", "formatos"]).default("todo"),
    omitir_vacias: z.boolean().default(false),
    transponer: z.boolean().default(false),
  }),
  writes: true,
  async run({ origen, destino, tipo, omitir_vacias, transponer }) {
    return Excel.run(async (context) => {
      const src = getRange(context, origen);
      const dest = getRange(context, destino);
      if (supports(9)) {
        dest.copyFrom(src, COPY_TYPES[tipo] as Excel.RangeCopyType, omitir_vacias, transponer);
        await context.sync();
        src.load("rowCount,columnCount");
        await context.sync();
        const rows = transponer ? src.columnCount : src.rowCount;
        const cols = transponer ? src.rowCount : src.columnCount;
        return checkErrors(context, dest.getCell(0, 0).getResizedRange(rows - 1, cols - 1));
      }
      // Alternativa para Excel sin ExcelApi 1.9.
      if (transponer || omitir_vacias) throw new Error("Transponer/omitir vacías requiere ExcelApi 1.9.");
      src.load("rowCount,columnCount,values,formulasR1C1,numberFormat");
      await context.sync();
      const target = dest.getCell(0, 0).getResizedRange(src.rowCount - 1, src.columnCount - 1);
      if (tipo === "todo" || tipo === "formulas") target.formulasR1C1 = src.formulasR1C1;
      if (tipo === "valores") target.values = src.values;
      if (tipo === "todo" || tipo === "formatos" || tipo === "valores") target.numberFormat = src.numberFormat;
      await context.sync();
      const result = await checkErrors(context, target);
      return { ...result, nota: "Copia parcial: solo valores/fórmulas y formato numérico (ExcelApi < 1.9)." };
    });
  },
});

export const dividirTexto = defineTool({
  name: "dividir_texto",
  description:
    "Texto en columnas: divide cada celda de una columna por un separador y escribe las partes en columnas " +
    "consecutivas a partir de `destino` (por defecto, sobre la misma columna de origen).",
  schema: z.object({
    origen: z.string().describe("Rango de UNA columna, p. ej. \"'#1 - Inventario'!B3:B49\"."),
    separador: z.string().min(1).describe("Separador, p. ej. \";\"."),
    destino: z.string().optional().describe("Celda superior izquierda del resultado. Por defecto, la primera del origen."),
    recortar_espacios: z.boolean().default(true),
  }),
  writes: true,
  async run({ origen, separador, destino, recortar_espacios }) {
    return Excel.run(async (context) => {
      const src = getRange(context, origen);
      src.load("values,columnCount");
      await context.sync();
      if (src.columnCount !== 1) throw new Error("El origen debe ser una sola columna.");
      const parts = splitRows(src.values, separador, recortar_espacios);
      const start = destino ? getRange(context, destino).getCell(0, 0) : src.getCell(0, 0);
      const target = start.getResizedRange(parts.length - 1, parts[0].length - 1);
      target.values = parts;
      await context.sync();
      target.load("address");
      await context.sync();
      return { rango: target.address, columnas_resultantes: parts[0].length };
    });
  },
});

const AUTOFILL = {
  serie: "FillSeries",
  meses: "FillMonths",
  dias: "FillDays",
  anios: "FillYears",
  copiar: "FillCopy",
  predeterminado: "FillDefault",
  formatos: "FillFormats",
} as const;

export const rellenarSerie = defineTool({
  name: "rellenar_serie",
  description:
    "Relleno automático de Excel (como arrastrar el controlador de relleno). `origen` son las celdas semilla " +
    "y `destino` el rango completo que debe quedar lleno (incluye el origen). Tipos: serie (1,2,3…), " +
    "meses/dias/anios (fechas), copiar, predeterminado (copia fórmulas ajustando referencias), formatos.",
  schema: z.object({
    origen: z.string().describe("Celdas semilla, p. ej. \"'#1 - Amortización'!A8\" o \"A8:A9\"."),
    destino: z.string().describe("Rango completo a rellenar, p. ej. \"'#1 - Amortización'!A8:A43\"."),
    tipo: z.enum(["serie", "meses", "dias", "anios", "copiar", "predeterminado", "formatos"]),
  }),
  writes: true,
  async run({ origen, destino, tipo }) {
    return Excel.run(async (context) => {
      const src = getRange(context, origen);
      const dest = getRange(context, destino);
      if (supports(9)) {
        src.autoFill(dest, AUTOFILL[tipo] as Excel.AutoFillType);
        await context.sync();
        return checkErrors(context, dest);
      }
      await fallbackFill(context, src, dest, tipo);
      const result = await checkErrors(context, dest);
      return { ...result, nota: "Relleno calculado por el complemento (ExcelApi < 1.9)." };
    });
  },
});

function transpose<T>(m: T[][]): T[][] {
  return m[0].map((_, c) => m.map((row) => row[c]));
}

async function fallbackFill(
  context: Excel.RequestContext,
  src: Excel.Range,
  dest: Excel.Range,
  tipo: keyof typeof AUTOFILL,
) {
  src.load("rowCount,columnCount,values,formulasR1C1,numberFormat");
  dest.load("rowCount,columnCount");
  await context.sync();
  if (tipo === "formatos") {
    dest.numberFormat = repeatPattern(src.numberFormat, dest.rowCount, dest.columnCount);
    await context.sync();
    return;
  }
  const vertical = dest.columnCount === src.columnCount;
  if (!vertical && dest.rowCount !== src.rowCount) {
    throw new Error("El destino debe extender el origen solo hacia abajo o hacia la derecha.");
  }
  // Trabajamos por columnas (vertical) o por filas (horizontal, transponiendo).
  const values = vertical ? src.values : transpose(src.values);
  const r1c1 = vertical ? src.formulasR1C1 : transpose(src.formulasR1C1);
  const length = vertical ? dest.rowCount : dest.columnCount;
  const lines = (vertical ? src.columnCount : src.rowCount);
  const out: (string | number | boolean)[][] = Array.from({ length }, () => []);
  for (let line = 0; line < lines; line++) {
    const seedVals = values.map((row) => row[line]);
    const seedF = r1c1.map((row) => row[line]);
    let column: (string | number | boolean)[];
    const hasFormula = seedF.some((f) => typeof f === "string" && f.startsWith("="));
    const numeric = seedVals.every((v) => typeof v === "number");
    if (hasFormula || tipo === "copiar" || tipo === "predeterminado" || !numeric) {
      column = repeatPattern(seedF.map((f) => [f]), length, 1).map((r) => r[0]);
    } else if (tipo === "serie") {
      column = linearSeries(seedVals as number[], length);
    } else {
      const seed = seedVals as number[];
      const unit = tipo as DateUnit;
      column = dateSeries(seed[0], length, unit);
    }
    column.forEach((v, i) => (out[i][line] = v));
  }
  dest.formulasR1C1 = (vertical ? out : transpose(out)) as unknown as string[][];
  await context.sync();
}
