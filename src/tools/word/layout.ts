import { z } from "zod";
import { requireWord } from "../../word/office";
import { cmToPt } from "../../word/spec";
import { defineTool } from "../types";

const PAPEL = {
  carta: "Letter",
  legal: "Legal",
  a4: "A4",
  a5: "A5",
  a3: "A3",
  ejecutivo: "Executive",
  tabloide: "Tabloid",
} as const;
const INICIO = { continua: "Continuous", pagina_nueva: "NewPage", pagina_par: "EvenPage", pagina_impar: "OddPage" } as const;
const VERTICAL = { arriba: "Top", centro: "Center", justificada: "Justify", abajo: "Bottom" } as const;

type PapelKey = keyof typeof PAPEL;

export const configurarPagina = defineTool({
  name: "configurar_pagina",
  description:
    "Configura la página de una sección o de todo el documento: orientación, tamaño de papel, márgenes y " +
    "encuadernación en cm, columnas (cantidad, espacio, línea entre columnas), primera página diferente, " +
    "pares e impares diferentes, alineación vertical, distancia del encabezado y del pie, y cómo empieza la " +
    "sección. Requiere WordApiDesktop 1.3 (Word de escritorio de Microsoft 365); si no está disponible, entrega " +
    "los pasos manuales (Disposición > Configurar página) y el VBA equivalente.",
  schema: z.object({
    seccion: z.number().int().min(1).optional().describe("Sin sección = todas las secciones."),
    orientacion: z.enum(["vertical", "horizontal"]).optional(),
    tamano_papel: z
      .enum(Object.keys(PAPEL) as [PapelKey, ...PapelKey[]])
      .optional()
      .describe('"carta" = 21.59 x 27.94 cm; "legal" = 21.59 x 35.56 cm (en algunos países, oficio). Otros: ancho_cm y alto_cm.'),
    ancho_cm: z.number().positive().optional(),
    alto_cm: z.number().positive().optional(),
    margenes_cm: z
      .object({
        superior: z.number().min(0).optional(),
        inferior: z.number().min(0).optional(),
        izquierdo: z.number().min(0).optional(),
        derecho: z.number().min(0).optional(),
      })
      .optional(),
    encuadernacion_cm: z.number().min(0).optional(),
    columnas: z
      .object({
        cantidad: z.number().int().min(1).max(13),
        espacio_cm: z.number().min(0).optional(),
        linea_entre: z.boolean().optional(),
      })
      .optional(),
    primera_pagina_diferente: z.boolean().optional(),
    pares_impares_diferentes: z.boolean().optional(),
    alineacion_vertical: z.enum(["arriba", "centro", "justificada", "abajo"]).optional(),
    distancia_encabezado_cm: z.number().min(0).optional(),
    distancia_pie_cm: z.number().min(0).optional(),
    inicio_seccion: z.enum(["continua", "pagina_nueva", "pagina_par", "pagina_impar"]).optional(),
  }),
  writes: true,
  async run(input) {
    requireWord(3, "WordApiDesktop", "Configurar la página");
    return Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();
      if (input.seccion && input.seccion > sections.items.length) {
        throw new Error(`El documento tiene ${sections.items.length} sección(es); no existe la sección ${input.seccion}.`);
      }
      const targets = input.seccion ? [sections.items[input.seccion - 1]] : sections.items;
      for (const section of targets) {
        const ps = section.pageSetup;
        // El papel antes que la orientación: cambiar el papel restablece el ancho y el alto.
        if (input.tamano_papel) ps.paperSize = PAPEL[input.tamano_papel];
        if (input.ancho_cm !== undefined) ps.pageWidth = cmToPt(input.ancho_cm);
        if (input.alto_cm !== undefined) ps.pageHeight = cmToPt(input.alto_cm);
        if (input.orientacion) ps.orientation = input.orientacion === "horizontal" ? "Landscape" : "Portrait";
        const m = input.margenes_cm;
        if (m?.superior !== undefined) ps.topMargin = cmToPt(m.superior);
        if (m?.inferior !== undefined) ps.bottomMargin = cmToPt(m.inferior);
        if (m?.izquierdo !== undefined) ps.leftMargin = cmToPt(m.izquierdo);
        if (m?.derecho !== undefined) ps.rightMargin = cmToPt(m.derecho);
        if (input.encuadernacion_cm !== undefined) ps.gutter = cmToPt(input.encuadernacion_cm);
        if (input.columnas) {
          ps.textColumns.setCount(input.columnas.cantidad);
          if (input.columnas.linea_entre !== undefined) ps.textColumns.setHasLineBetween(input.columnas.linea_entre);
        }
        if (input.primera_pagina_diferente !== undefined) ps.differentFirstPageHeaderFooter = input.primera_pagina_diferente;
        if (input.pares_impares_diferentes !== undefined) ps.oddAndEvenPagesHeaderFooter = input.pares_impares_diferentes;
        if (input.alineacion_vertical) ps.verticalAlignment = VERTICAL[input.alineacion_vertical];
        if (input.distancia_encabezado_cm !== undefined) ps.headerDistance = cmToPt(input.distancia_encabezado_cm);
        if (input.distancia_pie_cm !== undefined) ps.footerDistance = cmToPt(input.distancia_pie_cm);
        if (input.inicio_seccion) ps.sectionStart = INICIO[input.inicio_seccion];
      }
      await context.sync();
      const spacing = input.columnas?.espacio_cm;
      if (spacing !== undefined) {
        // El espacio entre columnas se define en cada columna (espacio después de ella).
        const columns = targets.map((s) => s.pageSetup.textColumns);
        for (const c of columns) c.load("items");
        await context.sync();
        for (const c of columns) for (const col of c.items) col.spaceAfter = cmToPt(spacing);
        await context.sync();
      }
      return { secciones: targets.length, total_secciones: sections.items.length };
    });
  },
});
