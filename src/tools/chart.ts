import { z } from "zod";
import { getRange, getSheet, supports } from "../excel/office";
import { defineTool } from "./types";

const LEGEND = { derecha: "Right", izquierda: "Left", arriba: "Top", abajo: "Bottom" } as const;

export const crearGrafico = defineTool({
  name: "crear_grafico",
  description:
    "Crea un gráfico. Dos modos: (a) `series` explícitas, cada una con su rango de valores, más `categorias` " +
    "(etiquetas del eje/radios); úsalo cuando solo deben incluirse ciertas columnas o filas; (b) `origen_datos` " +
    "con un bloque contiguo que incluye encabezados. Tipos comunes (Excel.ChartType): ColumnClustered, " +
    "3DColumnClustered, 3DColumn, BarClustered, Line, LineMarkers, Pie, Doughnut, Area, XYScatter, Radar, " +
    "RadarMarkers, RadarFilled.",
  schema: z.object({
    hoja: z.string().describe("Hoja donde se coloca el gráfico."),
    tipo: z.string().describe("Valor de Excel.ChartType, p. ej. \"RadarMarkers\"."),
    series: z
      .array(z.object({ nombre: z.string(), valores: z.string().describe("Rango de valores de la serie.") }))
      .optional(),
    categorias: z.string().optional().describe("Rango con las etiquetas de categoría (para modo series)."),
    origen_datos: z.string().optional().describe("Bloque contiguo con encabezados (modo b)."),
    series_por: z.enum(["Auto", "Columns", "Rows"]).default("Auto").describe("Solo para origen_datos."),
    titulo: z.string().optional(),
    posicion: z.string().optional().describe("Rango de celdas que ocupará el gráfico, p. ej. \"H2:P20\"."),
    leyenda: z.enum(["derecha", "izquierda", "arriba", "abajo", "ninguna"]).default("derecha"),
    titulo_eje_categorias: z.string().optional(),
    titulo_eje_valores: z.string().optional(),
    nombre: z.string().optional(),
  }),
  writes: true,
  async run(input) {
    if (!input.origen_datos && !input.series?.length) {
      throw new Error("Indica `series` o `origen_datos`.");
    }
    if (input.series && input.series.length > 1 && !supports(7)) {
      throw new Error("Series explícitas requieren ExcelApi 1.7; usa origen_datos.");
    }
    return Excel.run(async (context) => {
      const ws = getSheet(context, input.hoja);
      const type = input.tipo as Excel.ChartType;
      let chart: Excel.Chart;

      if (input.origen_datos) {
        chart = ws.charts.add(type, getRange(context, input.origen_datos), input.series_por);
      } else {
        const series = input.series!;
        const first = getRange(context, series[0].valores);
        first.load("rowCount,columnCount");
        await context.sync();
        const by = first.rowCount === 1 && first.columnCount > 1 ? "Rows" : "Columns";
        chart = ws.charts.add(type, first, by);
        const s0 = chart.series.getItemAt(0);
        s0.name = series[0].nombre;
        if (input.categorias) s0.setXAxisValues(getRange(context, input.categorias));
        for (const s of series.slice(1)) {
          const added = chart.series.add(s.nombre);
          added.setValues(getRange(context, s.valores));
          if (input.categorias) added.setXAxisValues(getRange(context, input.categorias));
        }
      }

      if (input.nombre) chart.name = input.nombre;
      if (input.titulo) {
        chart.title.text = input.titulo;
        chart.title.visible = true;
      }
      if (input.leyenda === "ninguna") {
        chart.legend.visible = false;
      } else {
        chart.legend.visible = true;
        chart.legend.position = LEGEND[input.leyenda];
      }
      if (input.posicion) {
        const [start, end] = input.posicion.split(":");
        chart.setPosition(start, end);
      }
      chart.load("name");
      await context.sync();

      // Los títulos de eje no existen en todos los tipos (p. ej. radial): se intentan aparte.
      const avisos: string[] = [];
      if (input.titulo_eje_categorias || input.titulo_eje_valores) {
        try {
          if (input.titulo_eje_categorias) chart.axes.categoryAxis.title.text = input.titulo_eje_categorias;
          if (input.titulo_eje_valores) chart.axes.valueAxis.title.text = input.titulo_eje_valores;
          await context.sync();
        } catch {
          avisos.push("Este tipo de gráfico no admite títulos de eje.");
        }
      }
      chart.series.load("items/name");
      await context.sync();
      return {
        grafico: chart.name,
        series: chart.series.items.map((s) => s.name),
        avisos: avisos.length ? avisos : undefined,
      };
    });
  },
});
