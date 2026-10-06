import { z } from "zod";
import { fillMatrix } from "../excel/address";
import { getRange } from "../excel/office";
import { defineTool } from "./types";

const BORDER_SIDES = {
  superior: ["EdgeTop"],
  inferior: ["EdgeBottom"],
  izquierdo: ["EdgeLeft"],
  derecho: ["EdgeRight"],
  contorno: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"],
  interiores: ["InsideHorizontal", "InsideVertical"],
  todos: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"],
} as const;

export const formatearRango = defineTool({
  name: "formatear_rango",
  description:
    "Aplica formato a un rango. Solo se cambian las propiedades que envíes. Formatos numéricos con códigos " +
    'en-US: moneda colones `"¢"#,##0.00`, fecha `dd/mm/yyyy`, porcentaje `0.00%`, moneda genérica ' +
    "`$#,##0.00`. Colores en hex (#RRGGBB). Anchos y altos en puntos (≈ 5.3 pt por carácter de ancho).",
  schema: z.object({
    rango: z.string(),
    formato_numero: z.string().optional(),
    alineacion_horizontal: z
      .enum(["General", "Left", "Center", "Right", "Fill", "Justify", "CenterAcrossSelection", "Distributed"])
      .optional(),
    alineacion_vertical: z.enum(["Top", "Center", "Bottom", "Justify", "Distributed"]).optional(),
    negrita: z.boolean().optional(),
    cursiva: z.boolean().optional(),
    subrayado: z.boolean().optional(),
    color_fuente: z.string().optional(),
    tamano_fuente: z.number().positive().optional(),
    fuente: z.string().optional().describe("Nombre de la fuente, p. ej. \"Calibri\"."),
    color_relleno: z.string().optional().describe("Hex, o \"ninguno\" para quitar el relleno."),
    bordes: z
      .object({
        lados: z.array(z.enum(["superior", "inferior", "izquierdo", "derecho", "contorno", "interiores", "todos"])).min(1),
        estilo: z
          .enum(["Continuous", "Dash", "DashDot", "DashDotDot", "Dot", "Double", "SlantDashDot", "None"])
          .default("Continuous"),
        grosor: z.enum(["Hairline", "Thin", "Medium", "Thick"]).default("Thin"),
        color: z.string().optional(),
      })
      .optional(),
    combinar: z.boolean().optional().describe("true combina las celdas; false las separa."),
    combinar_por_filas: z.boolean().optional().describe("Con combinar=true, combina cada fila por separado."),
    ajustar_texto: z.boolean().optional(),
    ancho_columna: z.number().positive().optional().describe("Ancho de las columnas del rango, en puntos."),
    autoajustar_columnas: z.boolean().optional(),
    alto_fila: z.number().positive().optional().describe("Alto de las filas del rango, en puntos."),
    autoajustar_filas: z.boolean().optional(),
  }),
  writes: true,
  async run(input) {
    return Excel.run(async (context) => {
      const range = getRange(context, input.rango);
      const f = range.format;
      if (input.formato_numero !== undefined) {
        range.load("rowCount,columnCount");
        await context.sync();
        range.numberFormat = fillMatrix(range.rowCount, range.columnCount, input.formato_numero);
      }
      if (input.alineacion_horizontal) f.horizontalAlignment = input.alineacion_horizontal;
      if (input.alineacion_vertical) f.verticalAlignment = input.alineacion_vertical;
      if (input.negrita !== undefined) f.font.bold = input.negrita;
      if (input.cursiva !== undefined) f.font.italic = input.cursiva;
      if (input.subrayado !== undefined) f.font.underline = input.subrayado ? "Single" : "None";
      if (input.color_fuente) f.font.color = input.color_fuente;
      if (input.tamano_fuente) f.font.size = input.tamano_fuente;
      if (input.fuente) f.font.name = input.fuente;
      if (input.color_relleno) {
        if (input.color_relleno.toLowerCase() === "ninguno") f.fill.clear();
        else f.fill.color = input.color_relleno;
      }
      if (input.bordes) {
        for (const lado of input.bordes.lados) {
          for (const idx of BORDER_SIDES[lado]) {
            const border = f.borders.getItem(idx as Excel.BorderIndex);
            border.style = input.bordes.estilo;
            if (input.bordes.estilo !== "None") {
              border.weight = input.bordes.grosor;
              if (input.bordes.color) border.color = input.bordes.color;
            }
          }
        }
      }
      if (input.combinar === true) range.merge(input.combinar_por_filas ?? false);
      if (input.combinar === false) range.unmerge();
      if (input.ajustar_texto !== undefined) f.wrapText = input.ajustar_texto;
      if (input.ancho_columna) f.columnWidth = input.ancho_columna;
      if (input.alto_fila) f.rowHeight = input.alto_fila;
      await context.sync();
      if (input.autoajustar_columnas) f.autofitColumns();
      if (input.autoajustar_filas) f.autofitRows();
      range.load("address");
      await context.sync();
      return { rango: range.address, ok: true };
    });
  },
});
