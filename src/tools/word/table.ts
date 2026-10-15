import { z } from "zod";
import { applyStyle, edgeOf, isBefore, isEmptyBody, requireWord, resolveTarget, tableIndexOf } from "../../word/office";
import { cmToPt, parseIndexList, runsForDeletion } from "../../word/spec";
import { defineTool } from "../types";
import { ALINEACION, Alineacion, applyFont, DEFAULT_OBJETIVO, fuenteShape, Objetivo } from "./common";

const Celda = z.union([z.string(), z.number()]);
type CeldaValue = z.infer<typeof Celda>;

const toStrings = (rows: CeldaValue[][]): string[][] => rows.map((r) => r.map(String));

const POSICION = { izquierda: "Left", centro: "Centered", derecha: "Right" } as const;
const Posicion = z.enum(["izquierda", "centro", "derecha"]).describe("Posición de la tabla entre los márgenes.");

const tablaShape = {
  tabla: z.number().int().min(1).describe("Número de tabla (1 = primera de la zona), según leer_documento."),
  zona: z.enum(["cuerpo", "encabezado", "pie"]).default("cuerpo"),
  seccion: z.number().int().min(1).default(1),
};

async function getTable(context: Word.RequestContext, tabla: number, zona: "cuerpo" | "encabezado" | "pie", seccion: number) {
  const t = await resolveTarget(context, { ...DEFAULT_OBJETIVO, zona, seccion, tabla });
  const table = t.table as Word.Table;
  table.load("rowCount,values");
  await context.sync();
  return { table, body: t.body, rows: table.rowCount, cols: Math.max(1, ...table.values.map((r) => r.length)) };
}

/** Celdas existentes en las filas y columnas indicadas (las combinadas pueden no existir). */
async function cellsIn(
  context: Word.RequestContext,
  table: Word.Table,
  size: { rows: number; cols: number },
  filas?: string,
  columnas?: string,
): Promise<{ cell: Word.TableCell; r: number; c: number }[]> {
  const rows = parseIndexList(filas ?? `1-${size.rows}`, size.rows, "fila");
  const cols = parseIndexList(columnas ?? `1-${size.cols}`, size.cols, "columna");
  const cells = rows.flatMap((r) =>
    cols.map((c) => {
      const cell = table.getCellOrNullObject(r - 1, c - 1);
      cell.load("rowIndex");
      return { cell, r, c };
    }),
  );
  await context.sync();
  return cells.filter((x) => !x.cell.isNullObject);
}

export const insertarTabla = defineTool({
  name: "insertar_tabla",
  description:
    "Inserta una tabla con datos (la primera fila suele ser el encabezado) al final del documento o antes/después " +
    "del objetivo, con un estilo de tabla integrado (TableGrid, GridTable4_Accent1, ListTable3_Accent2, " +
    "PlainTable1…) o el nombre local del estilo. Devuelve el número de tabla para editar_tabla y formato_tabla.",
  schema: z.object({
    objetivo: Objetivo.optional(),
    ubicacion: z.enum(["inicio", "final", "antes", "despues"]).default("final"),
    datos: z.array(z.array(Celda)).min(1).describe('Filas x columnas; "" para celdas vacías.'),
    estilo: z.string().default("TableGrid"),
    filas_encabezado: z.number().int().min(0).default(1).describe("Filas que se repiten como encabezado."),
    alineacion: Posicion.optional(),
    ajustar_a_ventana: z.boolean().default(false).describe("Ocupar todo el ancho entre márgenes."),
  }),
  writes: true,
  async run({ objetivo, ubicacion, datos, estilo, filas_encabezado, alineacion, ajustar_a_ventana }) {
    const cols = Math.max(...datos.map((r) => r.length));
    const values = datos.map((r) => Array.from({ length: cols }, (_, j) => (r[j] === undefined ? "" : String(r[j]))));
    return Word.run(async (context) => {
      const t = await resolveTarget(context, objetivo ?? DEFAULT_OBJETIVO);
      const before = isBefore(ubicacion);
      let table: Word.Table;
      if (t.kind === "zona") {
        // En una zona vacía, la tabla va antes del único párrafo (Word exige un párrafo después de la tabla).
        const empty = await isEmptyBody(context, t.body);
        table = empty
          ? empty.insertTable(values.length, cols, "Before", values)
          : t.body.insertTable(values.length, cols, before ? "Start" : "End", values);
      } else {
        table = edgeOf(t, before ? "first" : "last").insertTable(values.length, cols, before ? "Before" : "After", values);
      }
      applyStyle(table, estilo);
      table.headerRowCount = filas_encabezado;
      if (alineacion) table.alignment = POSICION[alineacion];
      if (ajustar_a_ventana) table.autoFitWindow();
      await context.sync();
      return { tabla: await tableIndexOf(context, t.body, table), filas: values.length, columnas: cols };
    });
  },
});

const Rango = z.object({
  fila_inicio: z.number().int().min(1),
  columna_inicio: z.number().int().min(1),
  fila_fin: z.number().int().min(1),
  columna_fin: z.number().int().min(1),
});

export const editarTabla = defineTool({
  name: "editar_tabla",
  description:
    "Cambia el contenido o la estructura de una tabla, en este orden: 1) eliminar filas/columnas, 2) agregar " +
    "filas/columnas (al inicio o al final, con datos opcionales), 3) escribir valores desde una celda, " +
    "4) combinar celdas (WordApi 1.4). Filas y columnas empiezan en 1. eliminar_tabla borra la tabla completa.",
  schema: z.object({
    ...tablaShape,
    eliminar_filas: z.string().optional().describe('"3", "2-4" o "1,5".'),
    eliminar_columnas: z.string().optional(),
    agregar_filas: z
      .object({
        cantidad: z.number().int().min(1),
        al_inicio: z.boolean().default(false),
        datos: z.array(z.array(Celda)).optional(),
      })
      .optional(),
    agregar_columnas: z
      .object({
        cantidad: z.number().int().min(1),
        al_inicio: z.boolean().default(false),
        datos: z.array(z.array(Celda)).optional().describe("Una fila por cada fila de la tabla."),
      })
      .optional(),
    valores: z
      .object({
        fila: z.number().int().min(1).default(1),
        columna: z.number().int().min(1).default(1),
        datos: z.array(z.array(z.union([Celda, z.null()]))).describe("null deja la celda como está."),
      })
      .optional(),
    combinar: z.array(Rango).optional(),
    eliminar_tabla: z.boolean().optional(),
  }),
  writes: true,
  async run(input) {
    if (input.combinar?.length) requireWord(4, "WordApi", "Combinar celdas");
    return Word.run(async (context) => {
      const { table, rows, cols } = await getTable(context, input.tabla, input.zona, input.seccion);
      if (input.eliminar_tabla) {
        table.delete();
        await context.sync();
        return { eliminada: true };
      }
      if (input.eliminar_filas) {
        for (const [start, count] of runsForDeletion(parseIndexList(input.eliminar_filas, rows, "fila"))) {
          table.deleteRows(start - 1, count);
        }
      }
      if (input.eliminar_columnas) {
        for (const [start, count] of runsForDeletion(parseIndexList(input.eliminar_columnas, cols, "columna"))) {
          table.deleteColumns(start - 1, count);
        }
      }
      const add = input.agregar_filas;
      if (add) table.addRows(add.al_inicio ? "Start" : "End", add.cantidad, add.datos && toStrings(add.datos));
      const addCols = input.agregar_columnas;
      if (addCols) {
        table.addColumns(addCols.al_inicio ? "Start" : "End", addCols.cantidad, addCols.datos && toStrings(addCols.datos));
      }
      await context.sync();

      let omitidas = 0;
      if (input.valores) {
        const { fila, columna, datos } = input.valores;
        const targets = datos.flatMap((row, i) =>
          row.flatMap((value, j) => {
            if (value === null) return [];
            const cell = table.getCellOrNullObject(fila - 1 + i, columna - 1 + j);
            cell.load("rowIndex");
            return [{ cell, value: String(value) }];
          }),
        );
        await context.sync();
        for (const { cell, value } of targets) {
          if (cell.isNullObject) omitidas++;
          else cell.value = value;
        }
      }
      // De abajo hacia arriba y de derecha a izquierda para que los índices no se corran.
      const merges = [...(input.combinar ?? [])].sort(
        (a, b) => b.fila_inicio - a.fila_inicio || b.columna_inicio - a.columna_inicio,
      );
      for (const m of merges) {
        table.mergeCells(m.fila_inicio - 1, m.columna_inicio - 1, m.fila_fin - 1, m.columna_fin - 1);
      }
      table.load("rowCount,values");
      await context.sync();
      return {
        filas: table.rowCount,
        columnas: Math.max(1, ...table.values.map((r) => r.length)),
        celdas_omitidas: omitidas || undefined,
      };
    });
  },
});

const BORDE_DONDE = {
  todos: "All",
  exterior: "Outside",
  interior: "Inside",
  superior: "Top",
  inferior: "Bottom",
  izquierdo: "Left",
  derecho: "Right",
  interior_horizontal: "InsideHorizontal",
  interior_vertical: "InsideVertical",
} as const;
const BORDE_TIPO = {
  ninguno: "None",
  simple: "Single",
  doble: "Double",
  punteado: "Dotted",
  discontinuo: "Dashed",
  grueso: "Single",
} as const;
const VERTICAL = { arriba: "Top", centro: "Center", abajo: "Bottom" } as const;

type BordeKey = keyof typeof BORDE_DONDE;
type TipoKey = keyof typeof BORDE_TIPO;

export const formatoTabla = defineTool({
  name: "formato_tabla",
  description:
    "Formato de una tabla: estilo, opciones de estilo (fila de encabezado, filas con bandas, primera/última " +
    "columna, fila de totales), posición, autoajuste, anchos de columna en cm, y por grupos de celdas " +
    "(filas/columnas): sombreado, alineación horizontal y vertical, fuente, alto de fila; y bordes de la tabla o " +
    "de un grupo de celdas.",
  schema: z.object({
    ...tablaShape,
    estilo: z.string().optional(),
    filas_encabezado: z.number().int().min(0).optional(),
    bandas_filas: z.boolean().optional(),
    bandas_columnas: z.boolean().optional(),
    primera_columna: z.boolean().optional(),
    ultima_columna: z.boolean().optional(),
    fila_total: z.boolean().optional(),
    alineacion: Posicion.optional(),
    ajuste: z.enum(["ventana", "distribuir_columnas"]).optional(),
    ancho_columnas_cm: z.array(z.number().positive()).optional().describe("Ancho de cada columna, en orden."),
    celdas: z
      .array(
        z.object({
          filas: z.string().optional().describe('"1", "2-5"; sin filas = todas.'),
          columnas: z.string().optional(),
          sombreado: z.string().optional().describe('Color de fondo, "#D9E2F3".'),
          alineacion_horizontal: Alineacion.optional(),
          alineacion_vertical: z.enum(["arriba", "centro", "abajo"]).optional(),
          alto_fila_cm: z.number().positive().optional(),
          ...fuenteShape,
        }),
      )
      .optional(),
    bordes: z
      .array(
        z.object({
          donde: z.enum(Object.keys(BORDE_DONDE) as [BordeKey, ...BordeKey[]]).default("todos"),
          tipo: z.enum(Object.keys(BORDE_TIPO) as [TipoKey, ...TipoKey[]]).default("simple"),
          color: z.string().optional(),
          grosor_pt: z.number().positive().optional(),
          filas: z.string().optional().describe("Solo esas filas (si no, toda la tabla)."),
          columnas: z.string().optional(),
        }),
      )
      .optional(),
  }),
  writes: true,
  async run(input) {
    return Word.run(async (context) => {
      const { table, rows, cols } = await getTable(context, input.tabla, input.zona, input.seccion);
      const size = { rows, cols };
      if (input.estilo) applyStyle(table, input.estilo);
      if (input.filas_encabezado !== undefined) table.headerRowCount = input.filas_encabezado;
      if (input.bandas_filas !== undefined) table.styleBandedRows = input.bandas_filas;
      if (input.bandas_columnas !== undefined) table.styleBandedColumns = input.bandas_columnas;
      if (input.primera_columna !== undefined) table.styleFirstColumn = input.primera_columna;
      if (input.ultima_columna !== undefined) table.styleLastColumn = input.ultima_columna;
      if (input.fila_total !== undefined) table.styleTotalRow = input.fila_total;
      if (input.alineacion) table.alignment = POSICION[input.alineacion];
      if (input.ajuste === "ventana") table.autoFitWindow();
      if (input.ajuste === "distribuir_columnas") table.distributeColumns();

      if (input.ancho_columnas_cm) {
        const widths = input.ancho_columnas_cm;
        const n = Math.min(widths.length, cols);
        for (const { cell, c } of await cellsIn(context, table, size, undefined, `1-${n}`)) {
          cell.columnWidth = cmToPt(widths[c - 1]);
        }
      }

      const warnings: string[] = [];
      const tableRows = input.celdas?.some((g) => g.alto_fila_cm) ? table.rows : null;
      tableRows?.load("items");
      for (const group of input.celdas ?? []) {
        for (const { cell } of await cellsIn(context, table, size, group.filas, group.columnas)) {
          if (group.sombreado) cell.shadingColor = group.sombreado;
          if (group.alineacion_horizontal) cell.horizontalAlignment = ALINEACION[group.alineacion_horizontal];
          if (group.alineacion_vertical) cell.verticalAlignment = VERTICAL[group.alineacion_vertical];
          applyFont(cell.body.font, group, warnings);
        }
        if (group.alto_fila_cm && tableRows) {
          for (const r of parseIndexList(group.filas ?? `1-${rows}`, rows, "fila")) {
            tableRows.items[r - 1].preferredHeight = cmToPt(group.alto_fila_cm);
          }
        }
      }

      for (const b of input.bordes ?? []) {
        const apply = (border: Word.TableBorder) => {
          border.type = BORDE_TIPO[b.tipo];
          if (b.color) border.color = b.color;
          if (b.grosor_pt || b.tipo === "grueso") border.width = b.grosor_pt ?? 2.25;
        };
        if (b.filas || b.columnas) {
          for (const { cell } of await cellsIn(context, table, size, b.filas, b.columnas)) apply(cell.getBorder(BORDE_DONDE[b.donde]));
        } else {
          apply(table.getBorder(BORDE_DONDE[b.donde]));
        }
      }
      await context.sync();
      return { tabla: input.tabla, avisos: warnings.length ? [...new Set(warnings)] : undefined };
    });
  },
});
