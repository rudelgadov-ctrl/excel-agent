import { z } from "zod";
import { cellAddress, excelApiMinor, getRange, supports } from "../excel/office";
import { defineTool } from "./types";

type Cell = string | number | boolean | null;

function isEmpty(v: unknown): boolean {
  return v === "" || v === null || v === undefined;
}

export const leerLibro = defineTool({
  name: "leer_libro",
  description:
    "Resumen del libro abierto: hojas (con su rango usado, celdas combinadas, gráficos, tablas, tablas " +
    "dinámicas y texto de cuadros de texto), nombres definidos, hoja activa, versión de ExcelApi soportada y " +
    "una vista previa de las celdas no vacías de cada hoja (fórmulas incluidas). Úsala siempre al empezar: " +
    "a veces las instrucciones vienen dentro del propio libro.",
  schema: z.object({
    max_celdas_por_hoja: z
      .number()
      .int()
      .min(0)
      .max(1000)
      .default(250)
      .describe("Máximo de celdas no vacías a mostrar por hoja en la vista previa."),
  }),
  writes: false,
  async run({ max_celdas_por_hoja }) {
    return Excel.run(async (context) => {
      const wb = context.workbook;
      const sheets = wb.worksheets;
      sheets.load("items/name,items/position,items/visibility");
      const active = sheets.getActiveWorksheet();
      active.load("name");
      const names = wb.names;
      names.load("items/name,items/type,items/value,items/visible");
      await context.sync();

      const perSheet = sheets.items.map((ws) => {
        const used = ws.getUsedRangeOrNullObject(false);
        used.load("address,rowCount,columnCount,rowIndex,columnIndex");
        const charts = ws.charts;
        charts.load(supports(7) ? "items/name,items/chartType" : "items/name");
        const tables = ws.tables;
        tables.load("items/name");
        return { ws, used, charts, tables };
      });
      await context.sync();

      const previews = perSheet.map(({ used }) => {
        if (used.isNullObject) return null;
        // Limitar la vista previa a un bloque razonable.
        const rows = Math.min(used.rowCount, 200);
        const cols = Math.min(used.columnCount, 30);
        const block = used.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
        block.load("formulas");
        const merged = supports(13) ? used.getMergedAreasOrNullObject() : null;
        merged?.load("areas/items/address");
        return { block, merged };
      });
      await context.sync();

      const hojas = perSheet.map(({ ws, used, charts, tables }, i) => {
        const p = previews[i];
        const celdas: Record<string, Cell> = {};
        let total = 0;
        if (p && !used.isNullObject) {
          p.block.formulas.forEach((row: Cell[], r: number) =>
            row.forEach((v: Cell, c: number) => {
              if (isEmpty(v)) return;
              total++;
              if (Object.keys(celdas).length < max_celdas_por_hoja) {
                celdas[cellAddress(used.rowIndex + r, used.columnIndex + c)] = v;
              }
            }),
          );
        }
        const combinadas =
          p?.merged && !p.merged.isNullObject ? p.merged.areas.items.map((a) => a.address.split("!").pop()) : undefined;
        return {
          nombre: ws.name,
          posicion: ws.position,
          visibilidad: ws.visibility,
          rango_usado: used.isNullObject ? null : used.address.split("!").pop(),
          celdas_combinadas: combinadas,
          graficos: charts.items.map((c) => ({ nombre: c.name, tipo: supports(7) ? c.chartType : undefined })),
          tablas: tables.items.map((t) => t.name),
          celdas_no_vacias: total,
          vista_previa: celdas,
          vista_previa_recortada: total > Object.keys(celdas).length,
        };
      });

      const textos = await readShapeTexts(context, perSheet.map((p) => p.ws));
      const dinamicas = await readPivotNames(context, perSheet.map((p) => p.ws));
      hojas.forEach((h, i) => {
        if (textos[i]?.length) Object.assign(h, { cuadros_de_texto: textos[i] });
        if (dinamicas[i]?.length) Object.assign(h, { tablas_dinamicas: dinamicas[i] });
      });

      return {
        hoja_activa: active.name,
        excel_api: `1.${excelApiMinor()}`,
        nombres_definidos: names.items.map((n) => ({ nombre: n.name, tipo: n.type, referencia: n.value })),
        hojas,
      };
    });
  },
});

/**
 * Texto de las formas (cuadros de texto) de cada hoja: algunas plantillas traen ahí las
 * instrucciones. Es información opcional, así que cualquier fallo se ignora.
 */
async function readShapeTexts(
  context: Excel.RequestContext,
  sheets: Excel.Worksheet[],
): Promise<{ nombre: string; texto: string }[][]> {
  if (!supports(9)) return [];
  try {
    const shapes = sheets.map((ws) => {
      const s = ws.shapes;
      s.load("items/name,items/type");
      return s;
    });
    await context.sync();
    const withFrame = shapes.map((s) =>
      s.items
        .filter((shape) => shape.type === Excel.ShapeType.geometricShape)
        .map((shape) => {
          shape.textFrame.load("hasText");
          shape.textFrame.textRange.load("text");
          return shape;
        }),
    );
    await context.sync();
    return withFrame.map((list) =>
      list
        .filter((shape) => shape.textFrame.hasText)
        .map((shape) => ({ nombre: shape.name, texto: shape.textFrame.textRange.text.slice(0, 4000) })),
    );
  } catch {
    return [];
  }
}

async function readPivotNames(context: Excel.RequestContext, sheets: Excel.Worksheet[]): Promise<string[][]> {
  if (!supports(8)) return [];
  try {
    const pivots = sheets.map((ws) => {
      const p = ws.pivotTables;
      p.load("items/name");
      return p;
    });
    await context.sync();
    return pivots.map((p) => p.items.map((pt) => pt.name));
  } catch {
    return [];
  }
}

export const leerRango = defineTool({
  name: "leer_rango",
  description:
    "Lee un rango: valores, fórmulas y (opcional) formato numérico y texto mostrado. Úsala para inspeccionar " +
    "antes de escribir y para verificar el resultado después. Referencias: \"'Nombre de hoja'!A1:C10\", " +
    "\"A1\" (hoja activa) o un nombre definido.",
  schema: z.object({
    rango: z.string().describe("Rango a leer, p. ej. \"'Mi hoja'!A1:E20\"."),
    incluir_formatos: z
      .boolean()
      .default(false)
      .describe("Incluye el formato numérico y el texto tal como se ve en pantalla."),
    incluir_vacias: z.boolean().default(false).describe("Incluye las celdas vacías en el resultado."),
    max_celdas: z.number().int().min(1).max(5000).default(1500),
  }),
  writes: false,
  async run({ rango, incluir_formatos, incluir_vacias, max_celdas }) {
    return Excel.run(async (context) => {
      const range = getRange(context, rango);
      range.load("address,rowCount,columnCount");
      await context.sync();

      const cols = range.columnCount;
      const rows = Math.max(1, Math.min(range.rowCount, Math.floor(max_celdas / cols)));
      const target = range.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
      target.load(
        incluir_formatos
          ? "address,rowIndex,columnIndex,values,formulas,numberFormat,text"
          : "address,rowIndex,columnIndex,values,formulas",
      );
      await context.sync();

      const celdas: Record<string, Record<string, Cell>> = {};
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const value = target.values[r][c] as Cell;
          const formula = target.formulas[r][c] as Cell;
          const isFormula = typeof formula === "string" && formula.startsWith("=");
          if (!incluir_vacias && isEmpty(value) && !isFormula) continue;
          const cell: Record<string, Cell> = { v: value };
          if (isFormula) cell.f = formula;
          if (incluir_formatos) {
            const nf = target.numberFormat[r][c] as string;
            if (nf && nf !== "General") cell.formato = nf;
            cell.texto = target.text[r][c];
          }
          celdas[cellAddress(target.rowIndex + r, target.columnIndex + c)] = cell;
        }
      }
      return {
        rango: range.address,
        filas: range.rowCount,
        columnas: cols,
        recortado: rows < range.rowCount ? `Solo se leyeron ${rows} filas.` : undefined,
        celdas,
      };
    });
  },
});
