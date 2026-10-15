import { z } from "zod";
import { truncate } from "../../excel/text";
import { resolveTarget, wordApiMinor, wordSupports, zoneBody } from "../../word/office";
import { lineSpacingMultiple, ptToCm } from "../../word/spec";
import { defineTool } from "../types";
import { ALINEACION_ES, DEFAULT_OBJETIVO } from "./common";

const HEADER_TYPES = ["Primary", "FirstPage", "EvenPages"] as const;

const zonaShape = {
  zona: z.enum(["cuerpo", "encabezado", "pie"]).default("cuerpo"),
  seccion: z.number().int().min(1).default(1).describe("Sección cuando zona es encabezado o pie."),
  tipo_encabezado: z.enum(["Primary", "FirstPage", "EvenPages"]).default("Primary"),
};

function apiVersion(): string {
  const desktop = wordApiMinor("WordApiDesktop");
  return `WordApi 1.${wordApiMinor()}${desktop ? ` · WordApiDesktop 1.${desktop}` : ""}`;
}

export const leerDocumento = defineTool({
  name: "leer_documento",
  description:
    "Resumen del documento abierto: párrafos numerados (texto, estilo, alineación, viñetas o numeración; con " +
    "detalle=true también fuente, tamaño, sangrías, espaciado e interlineado), tablas, secciones con encabezados, " +
    "pies y configuración de página, imágenes, campos, notas, propiedades y la selección actual. Los números de " +
    "párrafo son los que usan las demás herramientas (objetivo.parrafos) y cambian al insertar o borrar contenido. " +
    "Úsala siempre al empezar (a veces las instrucciones vienen dentro del propio documento) y para verificar.",
  schema: z.object({
    desde: z.number().int().min(1).default(1).describe("Primer párrafo a mostrar."),
    hasta: z.number().int().min(1).optional().describe("Último párrafo a mostrar (por defecto, 150 desde `desde`)."),
    detalle: z.boolean().default(false).describe("Incluir el formato de fuente y de párrafo de cada párrafo."),
    buscar: z
      .string()
      .optional()
      .describe("Texto a localizar: devuelve los números de párrafo donde aparece (sin distinguir mayúsculas)."),
    ...zonaShape,
  }),
  writes: false,
  async run({ desde, hasta, detalle, buscar, zona, seccion, tipo_encabezado }) {
    return Word.run(async (context) => {
      const doc = context.document;
      const body = await zoneBody(context, { zona, seccion, tipo_encabezado });
      const paras = body.paragraphs;
      const props = ["text", "style", "alignment", "tableNestingLevel", "isListItem"];
      if (detalle) {
        props.push(
          "leftIndent",
          "rightIndent",
          "firstLineIndent",
          "spaceBefore",
          "spaceAfter",
          "lineSpacing",
          "font/name",
          "font/size",
          "font/bold",
          "font/italic",
          "font/underline",
          "font/color",
          "font/highlightColor",
        );
      }
      paras.load(props.map((p) => `items/${p}`).join(","));
      const tables = body.tables;
      tables.load("items/rowCount,items/values,items/style,items/headerRowCount");
      const pictures = body.inlinePictures;
      pictures.load("items/width,items/height,items/altTextDescription");
      const shapes = wordSupports(2, "WordApiDesktop") ? body.shapes : null;
      shapes?.load("items/type,items/name");
      const fields = wordSupports(4) ? body.fields : null;
      fields?.load("items/code");
      const footnotes = wordSupports(5) ? body.footnotes : null;
      footnotes?.load("items/type");
      const sections = doc.sections;
      sections.load("items");
      const selection = doc.getSelection();
      selection.load("text");
      const properties = doc.properties;
      properties.load("title,author,subject");
      await context.sync();

      const total = paras.items.length;
      const start = Math.min(desde, Math.max(total, 1));
      const end = Math.min(hasta ?? start + 149, total);
      const window = paras.items.slice(start - 1, end);
      const listItems = window.map((p) => (p.isListItem ? p.listItem : null));
      for (const li of listItems) li?.load("level,listString");

      // Encabezados, pies y página de cada sección (solo al leer el cuerpo).
      const perSection =
        zona === "cuerpo"
          ? sections.items.slice(0, 10).map((s) => {
              const parts = HEADER_TYPES.map((type) => {
                const header = s.getHeader(type);
                const footer = s.getFooter(type);
                header.load("text");
                footer.load("text");
                return { type, header, footer };
              });
              const page = wordSupports(3, "WordApiDesktop") ? s.pageSetup : null;
              page?.load(
                "orientation,paperSize,pageWidth,pageHeight,topMargin,bottomMargin,leftMargin,rightMargin," +
                  "differentFirstPageHeaderFooter,oddAndEvenPagesHeaderFooter,sectionStart",
              );
              const columns = page ? page.textColumns : null;
              columns?.load("items/width");
              return { parts, page, columns };
            })
          : [];
      await context.sync();

      const parrafos = window.map((p, i) => {
        const row: Record<string, unknown> = { n: start + i, texto: truncate(p.text, 300) };
        if (p.tableNestingLevel > 0) {
          row.en_tabla = true;
          return row;
        }
        row.estilo = p.style;
        if (p.alignment !== "Left") row.alin = ALINEACION_ES[p.alignment] ?? p.alignment;
        const li = listItems[i];
        if (li) row.lista = `${li.listString} (nivel ${li.level + 1})`;
        if (detalle) {
          const f = p.font;
          row.fuente = `${f.name ?? "mixta"} ${f.size ?? "?"}pt`;
          if (f.bold !== false) row.negrita = f.bold ?? "mixto";
          if (f.italic !== false) row.cursiva = f.italic ?? "mixto";
          if (f.underline && f.underline !== "None") row.subrayado = f.underline;
          if (f.color && !/^#?000000$/i.test(f.color)) row.color = f.color;
          if (f.highlightColor) row.resaltado = f.highlightColor;
          if (p.leftIndent) row.sangria_izq_cm = ptToCm(p.leftIndent);
          if (p.rightIndent) row.sangria_der_cm = ptToCm(p.rightIndent);
          if (p.firstLineIndent) row.primera_linea_cm = ptToCm(p.firstLineIndent);
          if (p.spaceBefore) row.antes_pt = p.spaceBefore;
          if (p.spaceAfter) row.despues_pt = p.spaceAfter;
          row.interlineado = lineSpacingMultiple(p.lineSpacing);
        }
        return row;
      });

      const secciones = perSection.map(({ parts, page, columns }, i) => {
        const texts = (key: "header" | "footer") =>
          Object.fromEntries(parts.filter((x) => x[key].text.trim()).map((x) => [x.type, truncate(x[key].text, 300)]));
        return {
          n: i + 1,
          encabezado: texts("header"),
          pie: texts("footer"),
          pagina: page
            ? {
                orientacion: page.orientation,
                papel: page.paperSize,
                tamano_cm: `${ptToCm(page.pageWidth)} x ${ptToCm(page.pageHeight)}`,
                margenes_cm: {
                  superior: ptToCm(page.topMargin),
                  inferior: ptToCm(page.bottomMargin),
                  izquierdo: ptToCm(page.leftMargin),
                  derecho: ptToCm(page.rightMargin),
                },
                columnas: columns?.items.length || 1,
                primera_pagina_diferente: page.differentFirstPageHeaderFooter,
                pares_impares_diferentes: page.oddAndEvenPagesHeaderFooter,
                inicio: page.sectionStart,
              }
            : undefined,
        };
      });

      const query = buscar?.toLowerCase();
      return {
        zona,
        total_parrafos: total,
        mostrando: total ? `${start}-${end}` : "vacío",
        siguiente: end < total ? end + 1 : undefined,
        parrafos,
        tablas: tables.items.map((t, i) => ({
          n: i + 1,
          filas: t.rowCount,
          columnas: Math.max(0, ...t.values.map((r) => r.length)),
          estilo: t.style,
          filas_encabezado: t.headerRowCount,
          vista_previa: t.values.slice(0, 3).map((r) => r.map((c) => truncate(c, 60))),
        })),
        imagenes: pictures.items.length
          ? pictures.items.map((p) => ({
              ancho_cm: ptToCm(p.width),
              alto_cm: ptToCm(p.height),
              alt: p.altTextDescription || undefined,
            }))
          : undefined,
        formas: shapes?.items.length ? shapes.items.map((s) => `${s.type}: ${s.name}`) : undefined,
        campos: fields?.items.length ? fields.items.slice(0, 30).map((f) => truncate(f.code, 120)) : undefined,
        notas_al_pie: footnotes?.items.length || undefined,
        secciones: secciones.length ? secciones : undefined,
        total_secciones: sections.items.length,
        propiedades: { titulo: properties.title, autor: properties.author, asunto: properties.subject },
        seleccion: selection.text ? truncate(selection.text, 500) : undefined,
        coincidencias: query
          ? paras.items.flatMap((p, i) => (p.text.toLowerCase().includes(query) ? [i + 1] : [])).slice(0, 200)
          : undefined,
        api: apiVersion(),
      };
    });
  },
});

export const leerTabla = defineTool({
  name: "leer_tabla",
  description:
    "Lee una tabla completa: valores de todas las celdas, celdas por fila (las combinadas reducen la cuenta), " +
    "estilo y filas de encabezado; con formato=true también sombreado, alineación, anchos y fuente por fila.",
  schema: z.object({
    tabla: z.number().int().min(1).describe("Número de tabla (1 = primera de la zona)."),
    formato: z.boolean().default(false),
    ...zonaShape,
  }),
  writes: false,
  async run({ tabla, formato, zona, seccion, tipo_encabezado }) {
    return Word.run(async (context) => {
      const { table } = await resolveTarget(context, { ...DEFAULT_OBJETIVO, zona, seccion, tipo_encabezado, tabla });
      const t = table as Word.Table;
      t.load("rowCount,values,style,headerRowCount,alignment,width,isUniform,styleBandedRows,styleFirstColumn");
      const rows = t.rows;
      rows.load("items/cellCount,items/preferredHeight,items/shadingColor,items/isHeader");
      await context.sync();
      const detail = formato
        ? rows.items.map((row) => {
            row.cells.load("items/columnWidth,items/shadingColor,items/horizontalAlignment,items/verticalAlignment");
            row.font.load("name,size,bold,color");
            return row;
          })
        : [];
      if (formato) await context.sync();
      return {
        filas: t.rowCount,
        celdas_por_fila: t.isUniform ? undefined : rows.items.map((r) => r.cellCount),
        estilo: t.style,
        filas_encabezado: t.headerRowCount,
        alineacion: t.alignment,
        ancho_cm: ptToCm(t.width),
        bandas_filas: t.styleBandedRows,
        primera_columna: t.styleFirstColumn,
        valores: t.values.map((r) => r.map((c) => truncate(c, 300))),
        formato: formato
          ? detail.map((row, i) => ({
              fila: i + 1,
              alto_cm: row.preferredHeight ? ptToCm(row.preferredHeight) : undefined,
              fuente: `${row.font.name ?? "mixta"} ${row.font.size ?? "?"}pt${row.font.bold ? " negrita" : ""}`,
              color: row.font.color,
              celdas: row.cells.items.map((c, j) => ({
                c: j + 1,
                ancho_cm: ptToCm(c.columnWidth),
                sombreado: c.shadingColor && c.shadingColor !== "#FFFFFF" ? c.shadingColor : undefined,
                alin: c.horizontalAlignment,
                vert: c.verticalAlignment,
              })),
            }))
          : undefined,
      };
    });
  },
});
