import { z } from "zod";
import { fillMatrix } from "../../excel/address";
import { cleanCell, excelWouldConvert } from "../../excel/clean";
import { cellAddress, getRange, supports } from "../../excel/office";
import { defineTool } from "../types";

const MAX_CELLS = 50_000;

export const limpiarDatos = defineTool({
  name: "limpiar_datos",
  description:
    "Depura los textos de un rango (también encabezados de tablas, lo que renombra sus columnas): quita espacios " +
    "sobrantes y caracteres no imprimibles (como ESPACIOS y LIMPIAR), cambia mayúsculas, y convierte números o " +
    "fechas guardados como texto en valores reales. No toca las fórmulas. Úsala antes de crear tablas dinámicas o " +
    "relaciones si los encabezados o claves traen espacios.",
  schema: z.object({
    rango: z.string(),
    recortar_espacios: z.boolean().default(true),
    quitar_no_imprimibles: z.boolean().default(true),
    texto_a_numero: z.boolean().default(false),
    separador_decimal: z.enum(["punto", "coma"]).default("punto").describe("Del texto original: 1,234.50 (punto) o 1.234,50 (coma)."),
    texto_a_fecha: z.boolean().default(false).describe("ISO (2024-03-15) o con barras (15/03/2024)."),
    orden_fecha: z.enum(["dia_mes_anio", "mes_dia_anio"]).default("dia_mes_anio"),
    formato_fecha: z.string().default("dd/mm/yyyy").describe("Formato para las fechas convertidas."),
    mayusculas: z.enum(["mayusculas", "minusculas", "nombre_propio"]).optional(),
  }),
  writes: true,
  async run(input) {
    return Excel.run(async (context) => {
      const range = getRange(context, input.rango);
      range.load("address,rowCount,columnCount,rowIndex,columnIndex");
      await context.sync();
      if (range.rowCount * range.columnCount > MAX_CELLS) {
        throw new Error(`El rango tiene más de ${MAX_CELLS} celdas; límpialo por partes.`);
      }
      range.load("values,formulas");
      await context.sync();
      const counts = { textos: 0, numeros: 0, fechas: 0 };
      const ejemplos: string[] = [];
      for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const f = range.formulas[r][c];
          if (typeof f === "string" && f.startsWith("=")) continue;
          const before = range.values[r][c];
          const res = cleanCell(before, input);
          if (!res) continue;
          const cell = range.getCell(r, c);
          if (res.kind === "text") {
            // Evita que Excel convierta un código como "00123" en número.
            if (excelWouldConvert(res.value)) cell.numberFormat = [["@"]];
            cell.values = [[res.value]];
            counts.textos++;
          } else if (res.kind === "number") {
            cell.values = [[res.value]];
            counts.numeros++;
          } else {
            cell.values = [[res.value]];
            cell.numberFormat = [[res.time ? `${input.formato_fecha} hh:mm` : input.formato_fecha]];
            counts.fechas++;
          }
          if (ejemplos.length < 8) {
            ejemplos.push(`${cellAddress(range.rowIndex + r, range.columnIndex + c)}: ${JSON.stringify(before)} -> ${JSON.stringify(res.value)}`);
          }
        }
      }
      await context.sync();
      return { rango: range.address, cambios: counts, ejemplos };
    });
  },
});

export const modificarTabla = defineTool({
  name: "modificar_tabla",
  description:
    "Modifica una tabla de Excel existente: renombrarla (el nombre que tendrá en el modelo de datos), estilo, fila de " +
    "totales, agregar columnas calculadas (fórmula en inglés con referencias estructuradas, p. ej. =YEAR([@Fecha]) o " +
    "=\"Trim \"&ROUNDUP(MONTH([@Fecha])/3,0), útiles para agrupar en tablas dinámicas), cambiar su rango o convertirla en " +
    "rango normal. Al renombrar avisa si un nombre definido ya usa ese nombre.",
  schema: z.object({
    tabla: z.string(),
    nuevo_nombre: z
      .string()
      .regex(/^[A-Za-z_\\À-ɏ][\w.À-ɏ]*$/, "Nombre de tabla inválido: sin espacios, empezando con letra")
      .optional(),
    quitar_nombre_en_conflicto: z
      .boolean()
      .default(false)
      .describe("Si un nombre definido se llama igual que nuevo_nombre, lo elimina (informa a qué apuntaba)."),
    estilo: z.string().optional().describe('P. ej. "TableStyleMedium2".'),
    fila_totales: z.boolean().optional(),
    agregar_columnas: z
      .array(
        z.object({
          nombre: z.string(),
          formula: z.string().optional().describe("Fórmula en inglés para toda la columna."),
          formato_numero: z.string().optional(),
        }),
      )
      .optional(),
    cambiar_rango: z.string().optional().describe("Nuevo rango de la tabla, encabezados incluidos."),
    convertir_a_rango: z.boolean().optional(),
  }),
  writes: true,
  async run(input) {
    return Excel.run(async (context) => {
      const wb = context.workbook;
      const table = wb.tables.getItemOrNullObject(input.tabla);
      wb.tables.load("items/name");
      await context.sync();
      if (table.isNullObject) {
        throw new Error(`No existe la tabla "${input.tabla}". Tablas: ${wb.tables.items.map((t) => t.name).join(", ") || "(ninguna)"}.`);
      }
      const hechos: string[] = [];
      const avisos: string[] = [];
      if (input.nuevo_nombre && input.nuevo_nombre.toLowerCase() !== input.tabla.toLowerCase()) {
        const nuevo = input.nuevo_nombre;
        if (wb.tables.items.some((t) => t.name.toLowerCase() === nuevo.toLowerCase())) {
          throw new Error(`Ya existe otra tabla llamada "${nuevo}".`);
        }
        const name = wb.names.getItemOrNullObject(nuevo);
        name.load("value");
        await context.sync();
        if (!name.isNullObject) {
          if (!input.quitar_nombre_en_conflicto) {
            throw new Error(
              `El nombre definido "${nuevo}" (=${name.value}) ya usa ese nombre. Usa quitar_nombre_en_conflicto=true ` +
                "si sobra (por ejemplo, un nombre viejo sobre el mismo rango) o elige otro nombre para la tabla.",
            );
          }
          name.delete();
          avisos.push(`Se eliminó el nombre definido "${nuevo}", que apuntaba a ${name.value}.`);
        }
        table.name = nuevo;
        hechos.push(`renombrada a ${nuevo}`);
      }
      if (input.estilo) {
        table.style = input.estilo;
        hechos.push("estilo");
      }
      if (input.fila_totales !== undefined) {
        table.showTotals = input.fila_totales;
        hechos.push("fila de totales");
      }
      for (const col of input.agregar_columnas ?? []) {
        const added = table.columns.add(undefined, undefined, col.nombre);
        const body = added.getDataBodyRange();
        body.load("rowCount");
        await context.sync();
        if (col.formula) body.formulas = fillMatrix(body.rowCount, 1, col.formula);
        if (col.formato_numero) body.numberFormat = fillMatrix(body.rowCount, 1, col.formato_numero);
        hechos.push(`columna ${col.nombre}`);
      }
      if (input.cambiar_rango) {
        if (!supports(13)) throw new Error("Cambiar el rango de una tabla requiere ExcelApi 1.13.");
        table.resize(input.cambiar_rango);
        hechos.push("nuevo rango");
      }
      await context.sync();
      if (input.convertir_a_rango) {
        const range = table.convertToRange();
        range.load("address");
        await context.sync();
        return { convertida_a_rango: range.address, cambios: hechos, avisos: avisos.length ? avisos : undefined };
      }
      table.load("name");
      const range = table.getRange();
      range.load("address");
      const header = table.getHeaderRowRange();
      header.load("values");
      const errors = table.getDataBodyRange();
      errors.load("values");
      await context.sync();
      const conError = (input.agregar_columnas ?? []).length
        ? errors.values.flat().filter((v) => typeof v === "string" && v.startsWith("#")).length
        : 0;
      return {
        tabla: table.name,
        rango: range.address,
        encabezados: header.values[0],
        cambios: hechos,
        celdas_con_error: conError || undefined,
        avisos: avisos.length ? avisos : undefined,
      };
    });
  },
});
