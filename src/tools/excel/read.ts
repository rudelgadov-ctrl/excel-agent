import { z } from "zod";
import { cellAddress, excelApiMinor, getRange, supports } from "../../excel/office";
import { defineTool } from "../types";

type Cell = string | number | boolean | null;

function isEmpty(v: unknown): boolean {
  return v === "" || v === null || v === undefined;
}

export const leerLibro = defineTool({
  name: "leer_libro",
  description:
    "Resumen del libro abierto: hojas (con su rango usado, celdas combinadas, gráficos, tablas con sus encabezados " +
    "exactos, tablas dinámicas con sus campos, segmentadores y formas con su texto, también dentro de grupos), " +
    "nombres definidos, hoja activa, versión de ExcelApi soportada y una vista previa de las celdas no vacías de cada " +
    "hoja (fórmulas incluidas). Úsala siempre al empezar (a veces las instrucciones vienen dentro del propio libro) " +
    "y para verificar al terminar.",
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
          tablas: tables.items.map((t) => t.name) as unknown[],
          celdas_no_vacias: total,
          vista_previa: celdas,
          vista_previa_recortada: total > Object.keys(celdas).length,
        };
      });

      const tablas = await readTables(context, perSheet.map((p) => p.tables));
      const formas = await readShapes(context, perSheet.map((p) => p.ws));
      const dinamicas = await readPivots(context, perSheet.map((p) => p.ws));
      const segmentadores = await readSlicers(context);
      hojas.forEach((h, i) => {
        if (tablas[i]?.length) h.tablas = tablas[i];
        if (formas[i]?.length) Object.assign(h, { formas: formas[i] });
        if (dinamicas[i]?.length) Object.assign(h, { tablas_dinamicas: dinamicas[i] });
        const sl = segmentadores.filter((s) => s.hoja === h.nombre).map(({ hoja: _, ...rest }) => rest);
        if (sl.length) Object.assign(h, { segmentadores: sl });
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

/** Tablas de Excel con su rango y encabezados (para usar los nombres exactos en tablas dinámicas y relaciones). */
async function readTables(
  context: Excel.RequestContext,
  collections: Excel.TableCollection[],
): Promise<{ nombre: string; rango: string; encabezados: Cell[] }[][]> {
  try {
    const parts = collections.map((c) =>
      c.items.map((t) => {
        const range = t.getRange();
        range.load("address");
        const header = t.getHeaderRowRange();
        header.load("values");
        return { t, range, header };
      }),
    );
    await context.sync();
    return parts.map((list) =>
      list.map(({ t, range, header }) => ({
        nombre: t.name,
        rango: range.address.split("!").pop() ?? range.address,
        encabezados: header.values[0] as Cell[],
      })),
    );
  } catch {
    return [];
  }
}

interface ShapeInfo {
  nombre: string;
  tipo: string;
  texto?: string;
  formas?: string[];
  posicion?: string;
}

/**
 * Formas de cada hoja con su texto (algunas plantillas traen ahí las instrucciones; en un dashboard son los
 * indicadores) y las formas de cada grupo. Es información opcional, así que cualquier fallo se ignora.
 */
async function readShapes(context: Excel.RequestContext, sheets: Excel.Worksheet[]): Promise<ShapeInfo[][]> {
  if (!supports(9)) return [];
  try {
    const shapes = sheets.map((ws) => {
      const s = ws.shapes;
      s.load("items/name,items/type,items/left,items/top,items/width,items/height");
      return s;
    });
    await context.sync();
    const details = shapes.map((s) =>
      s.items.map((shape) => {
        if (shape.type === Excel.ShapeType.geometricShape) {
          shape.textFrame.load("hasText");
          shape.textFrame.textRange.load("text");
        }
        const inner = shape.type === Excel.ShapeType.group ? shape.group.shapes : null;
        inner?.load("items/name");
        return { shape, inner };
      }),
    );
    await context.sync();
    return details.map((list) =>
      list.map(({ shape, inner }) => {
        const info: ShapeInfo = {
          nombre: shape.name,
          tipo: shape.type,
          posicion: `izq ${Math.round(shape.left)}, arriba ${Math.round(shape.top)}, ${Math.round(shape.width)}x${Math.round(shape.height)} pt`,
        };
        if (shape.type === Excel.ShapeType.geometricShape && shape.textFrame.hasText) {
          info.texto = shape.textFrame.textRange.text.slice(0, 4000);
        }
        if (inner) info.formas = inner.items.map((x) => x.name);
        return info;
      }),
    );
  } catch {
    return [];
  }
}

interface PivotInfo {
  nombre: string;
  rango?: string;
  filas?: string[];
  columnas?: string[];
  valores?: string[];
  filtros?: string[];
}

/** Tablas dinámicas con su rango y campos (las del modelo de datos pueden mostrarse solo con el nombre). */
async function readPivots(context: Excel.RequestContext, sheets: Excel.Worksheet[]): Promise<PivotInfo[][]> {
  if (!supports(8)) return [];
  try {
    const pivots = sheets.map((ws) => {
      const p = ws.pivotTables;
      p.load("items/name");
      return p;
    });
    await context.sync();
    const result: PivotInfo[][] = pivots.map((p) => p.items.map((pt) => ({ nombre: pt.name })));
    for (const [i, p] of pivots.entries()) {
      for (const [j, pt] of p.items.entries()) {
        try {
          const range = pt.layout.getRange();
          range.load("address");
          const axes = [pt.rowHierarchies, pt.columnHierarchies, pt.dataHierarchies, pt.filterHierarchies];
          axes.forEach((a) => a.load("items/name"));
          await context.sync();
          const [filas, columnas, valores, filtros] = axes.map((a) => a.items.map((h) => h.name));
          result[i][j] = { nombre: pt.name, rango: range.address.split("!").pop(), filas, columnas, valores, filtros };
        } catch {
          // Office.js no describe algunas tablas dinámicas (p. ej. del modelo de datos): queda solo el nombre.
        }
      }
    }
    return result;
  } catch {
    return [];
  }
}

async function readSlicers(context: Excel.RequestContext): Promise<{ hoja: string; nombre: string; titulo: string; seleccion?: string[] }[]> {
  if (!supports(10)) return [];
  try {
    const slicers = context.workbook.slicers;
    slicers.load("items/name,items/caption,items/isFilterCleared");
    await context.sync();
    const extra = slicers.items.map((s) => {
      s.worksheet.load("name");
      const selected = s.isFilterCleared ? null : s.getSelectedItems();
      return { s, selected };
    });
    await context.sync();
    return extra.map(({ s, selected }) => ({
      hoja: s.worksheet.name,
      nombre: s.name,
      titulo: s.caption,
      seleccion: selected ? selected.value : undefined,
    }));
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
