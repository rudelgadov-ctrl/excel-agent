import { z } from "zod";
import {
  applyStyle,
  edgeOf,
  isBefore,
  isEmptyBody,
  newParagraph,
  requireWord,
  resolveTarget,
  targetParagraphs,
  wordSupports,
  type Target,
  type Ubicacion as UbicacionValue,
} from "../../word/office";
import { fieldParagraph, wrapOoxml } from "../../word/ooxml";
import { cmToPt, splitFieldCode } from "../../word/spec";
import { defineTool, type ImageFile } from "../types";
import { ALINEACION, Alineacion, DEFAULT_OBJETIVO, Objetivo, Ubicacion, VERIFY_NOTE } from "./common";

type InlineLocation = "Start" | "End" | "Before" | "After" | "Replace";

const LOC: Record<UbicacionValue, InlineLocation> = {
  inicio: "Start",
  final: "End",
  antes: "Before",
  despues: "After",
  reemplazar: "Replace",
};

/** En la zona completa no hay "antes" ni "después": equivalen al inicio y al final. */
function zoneLocation(u: UbicacionValue): "Start" | "End" | "Replace" {
  return u === "reemplazar" ? "Replace" : isBefore(u) ? "Start" : "End";
}

interface BlockTarget {
  insertHtml(html: string, location: "Start" | "End" | "Replace"): Word.Range;
  insertOoxml(ooxml: string, location: "Start" | "End" | "Replace"): Word.Range;
}

async function insertParagraphs(
  context: Word.RequestContext,
  t: Target,
  ubicacion: UbicacionValue,
  items: { texto: string; estilo?: string }[],
) {
  const created: Word.Paragraph[] = [];
  let first: number | undefined;
  if (t.kind === "zona") {
    const empty = ubicacion !== "reemplazar" ? await isEmptyBody(context, t.body) : null;
    if (ubicacion === "reemplazar" || empty) {
      if (ubicacion === "reemplazar") t.body.clear();
      const p0 = empty ?? t.body.paragraphs.getFirst();
      p0.insertText(items[0].texto, "Replace");
      created.push(p0);
      first = 1;
    } else if (isBefore(ubicacion)) {
      created.push(t.body.insertParagraph(items[0].texto, "Start"));
      first = 1;
    } else {
      const paras = t.body.paragraphs;
      paras.load("items/text");
      await context.sync();
      created.push(t.body.insertParagraph(items[0].texto, "End"));
      first = paras.items.length + 1;
    }
  } else {
    if (ubicacion === "reemplazar" && (t.kind === "texto" || t.kind === "seleccion")) {
      throw new Error("Para reemplazar un texto usa `texto` (o `html`); `parrafos` agrega párrafos completos.");
    }
    const before = isBefore(ubicacion) || ubicacion === "reemplazar";
    created.push(edgeOf(t, before ? "first" : "last").insertParagraph(items[0].texto, before ? "Before" : "After"));
    if (t.indices) first = before ? t.indices[0] : t.indices[t.indices.length - 1] + 1;
  }
  for (const item of items.slice(1)) created.push(created[created.length - 1].insertParagraph(item.texto, "After"));
  items.forEach((item, i) => {
    if (item.estilo) applyStyle(created[i], item.estilo);
  });
  if (ubicacion === "reemplazar") {
    t.paragraphs?.forEach((p) => p.delete());
    t.table?.delete();
  }
  await context.sync();
  return {
    insertados: items.length,
    parrafos_nuevos: first !== undefined ? `${first}-${first + items.length - 1}` : undefined,
    nota: "Los números de párrafo posteriores cambiaron; vuelve a leer antes de usar números viejos.",
  };
}

export const insertarContenido = defineTool({
  name: "insertar_contenido",
  description:
    "Escribe en el documento. Usa UNO de: `parrafos` (párrafos nuevos, cada uno con su estilo: lo más fiable " +
    "para redactar), `texto` (texto en línea dentro de un párrafo, o para reemplazar un texto buscado: buscar y " +
    "reemplazar), `html` (contenido con formato mixto: negritas parciales, listas, tablas sencillas, colores) u " +
    "`ooxml` (WordprocessingML: <w:p>, <w:tbl>…, para lo que no tiene herramienta: ecuaciones OMML, saltos de " +
    "sección con su configuración, tabulaciones, bordes de párrafo). Sin `objetivo` actúa sobre todo el cuerpo " +
    "(ubicacion final = al final del documento). En un documento vacío reutiliza el primer párrafo. Indica el " +
    "estilo de cada párrafo nuevo (p. ej. Heading1, Normal): si no, hereda el del párrafo vecino.",
  schema: z.object({
    objetivo: Objetivo.optional(),
    ubicacion: Ubicacion.default("final"),
    parrafos: z
      .array(
        z.object({
          texto: z.string().describe("Un párrafo, sin saltos de línea."),
          estilo: z
            .string()
            .optional()
            .describe('Estilo integrado ("Heading1", "Title", "ListParagraph", "Normal"…) o nombre exacto de un estilo del documento.'),
        }),
      )
      .min(1)
      .optional(),
    texto: z.string().optional(),
    html: z.string().optional(),
    ooxml: z.string().optional().describe("Contenido de <w:body> (o un paquete pkg:package completo)."),
  }),
  writes: true,
  async run(input) {
    const given = (["parrafos", "texto", "html", "ooxml"] as const).filter((k) => input[k] !== undefined);
    if (given.length !== 1) throw new Error("Indica exactamente uno de: parrafos, texto, html u ooxml.");
    const { ubicacion } = input;
    return Word.run(async (context) => {
      const t = await resolveTarget(context, input.objetivo ?? DEFAULT_OBJETIVO);

      if (input.texto !== undefined) {
        if (t.kind === "zona") t.body.insertText(input.texto, zoneLocation(ubicacion));
        else for (const r of t.ranges) r.insertText(input.texto, LOC[ubicacion]);
        await context.sync();
        return { insertado: true, lugares: t.kind === "zona" ? 1 : t.ranges.length };
      }
      if (input.parrafos) return insertParagraphs(context, t, ubicacion, input.parrafos);

      // HTML u OOXML: bloques completos.
      const anchors: { target: BlockTarget; loc: "Start" | "End" | "Replace" }[] = [];
      const inline = t.kind === "texto" || t.kind === "seleccion";
      if (t.kind === "zona") {
        const empty = ubicacion !== "reemplazar" ? await isEmptyBody(context, t.body) : null;
        anchors.push(empty ? { target: empty, loc: "Replace" } : { target: t.body, loc: zoneLocation(ubicacion) });
      } else if (inline && (ubicacion === "reemplazar" || ubicacion === "inicio" || ubicacion === "final")) {
        for (const r of t.ranges) anchors.push({ target: r, loc: LOC[ubicacion] as "Start" | "End" | "Replace" });
      } else {
        // Párrafo vacío como marcador antes o después del objetivo; al reemplazar se borra el original.
        const where = ubicacion === "reemplazar" ? "antes" : ubicacion;
        anchors.push({ target: await newParagraph(context, t, where), loc: "Replace" });
      }
      const content = input.html ?? wrapOoxml(input.ooxml as string);
      for (const a of anchors) {
        if (input.html !== undefined) a.target.insertHtml(content, a.loc);
        else a.target.insertOoxml(content, a.loc);
      }
      if (ubicacion === "reemplazar" && !inline) {
        t.paragraphs?.forEach((p) => p.delete());
        t.table?.delete();
      }
      await context.sync();
      return { insertado: true, nota: VERIFY_NOTE };
    });
  },
});

const PROTECTED_XML = [
  /<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*<w:sectPr/,
  /<w:br [^>]*w:type="page"/,
  /<w:(drawing|pict|object|fldChar|fldSimple|sdt|bookmarkStart)\b/,
];

export const eliminarContenido = defineTool({
  name: "eliminar_contenido",
  description:
    "Borra el objetivo: párrafos, una tabla, las apariciones de un texto, la selección o toda la zona. Con " +
    "solo_vacios=true borra solo los párrafos vacíos del objetivo (conserva saltos de página o sección, imágenes, " +
    "campos y celdas de tabla); úsalo para limpiar líneas en blanco sobrantes.",
  schema: z.object({
    objetivo: Objetivo,
    solo_vacios: z.boolean().default(false),
  }),
  writes: true,
  async run({ objetivo, solo_vacios }) {
    return Word.run(async (context) => {
      const t = await resolveTarget(context, objetivo);
      if (solo_vacios) {
        const paras = await targetParagraphs(context, t);
        for (const p of paras) p.load("text,tableNestingLevel,isLastParagraph");
        await context.sync();
        const candidates = paras
          .filter((p) => p.text.replace(/[ \t ]/g, "") === "" && p.tableNestingLevel === 0 && !p.isLastParagraph)
          .slice(0, 300);
        const xml = candidates.map((p) => p.getOoxml());
        await context.sync();
        const empties = candidates.filter((_, i) => !PROTECTED_XML.some((re) => re.test(xml[i].value)));
        for (const p of empties) p.delete();
        await context.sync();
        return { borrados: empties.length, conservados: candidates.length - empties.length };
      }
      if (t.kind === "zona") t.body.clear();
      else if (t.table) t.table.delete();
      else if (t.paragraphs) t.paragraphs.forEach((p) => p.delete());
      else t.ranges.forEach((r) => r.delete());
      await context.sync();
      return { borrado: t.kind, cantidad: t.paragraphs?.length ?? t.ranges.length };
    });
  },
});

const BREAKS = {
  pagina: "Page",
  seccion_pagina_siguiente: "SectionNext",
  seccion_continua: "SectionContinuous",
  seccion_pagina_par: "SectionEven",
  seccion_pagina_impar: "SectionOdd",
  linea: "Line",
} as const;

export const insertarSalto = defineTool({
  name: "insertar_salto",
  description:
    "Inserta un salto de página, de sección (página siguiente, continua, par, impar) o de línea antes o después " +
    "de cada párrafo del objetivo (o de cada aparición de un texto, o de la tabla). Los saltos de sección " +
    "permiten cambiar orientación, márgenes, columnas o encabezados solo en una parte del documento " +
    "(después usa configurar_pagina con esa sección).",
  schema: z.object({
    objetivo: Objetivo.optional(),
    ubicacion: z.enum(["antes", "despues"]).default("despues"),
    tipo: z.enum(Object.keys(BREAKS) as [keyof typeof BREAKS, ...(keyof typeof BREAKS)[]]),
  }),
  writes: true,
  async run({ objetivo, ubicacion, tipo }) {
    return Word.run(async (context) => {
      const t = await resolveTarget(context, objetivo ?? DEFAULT_OBJETIVO);
      const type = BREAKS[tipo];
      const loc = ubicacion === "antes" ? "Before" : "After";
      if (t.kind === "zona") t.body.insertBreak(type, ubicacion === "antes" ? "Start" : "End");
      else if (t.paragraphs) t.paragraphs.forEach((p) => p.insertBreak(type, loc));
      else t.ranges.forEach((r) => r.insertBreak(type, loc));
      await context.sync();
      return { insertados: t.kind === "zona" ? 1 : (t.paragraphs?.length ?? t.ranges.length), tipo };
    });
  },
});

export const insertarCampo = defineTool({
  name: "insertar_campo",
  description:
    "Inserta un campo de Word por su código: número de página (PAGE), total de páginas (NUMPAGES), fecha " +
    '(DATE \\@ "dd/MM/yyyy"), tabla de contenido (TOC \\o "1-3" \\h \\z \\u), tabla de ilustraciones ' +
    '(TOC \\h \\z \\c "Ilustración"), numeración de títulos y figuras (SEQ Ilustración \\* ARABIC), referencias ' +
    "(REF marcador \\h, PAGEREF marcador), índice (XE \"entrada\" e INDEX), autor, nombre de archivo, campos de " +
    "combinación de correspondencia (MERGEFIELD Nombre), etc. Para páginas en el pie usa objetivo.zona \"pie\". " +
    "Para una tabla de contenido o índice usa parrafo_propio=true.",
  schema: z.object({
    objetivo: Objetivo.optional(),
    ubicacion: Ubicacion.default("final"),
    codigo: z.string().min(1).describe("Código del campo sin llaves, en inglés (como en Alt+F9)."),
    parrafo_propio: z
      .boolean()
      .default(false)
      .describe("Crear un párrafo nuevo para el campo (antes/después del objetivo o al inicio/final de la zona)."),
  }),
  writes: true,
  async run({ objetivo, ubicacion, codigo, parrafo_propio }) {
    return Word.run(async (context) => {
      const t = await resolveTarget(context, objetivo ?? DEFAULT_OBJETIVO);
      let ranges: Word.Range[];
      let loc: InlineLocation;
      if (parrafo_propio) {
        ranges = [(await newParagraph(context, t, ubicacion === "reemplazar" ? "antes" : ubicacion)).getRange("Content")];
        loc = "Start";
      } else if (t.kind === "zona") {
        ranges = [t.body.getRange("Content")];
        loc = zoneLocation(ubicacion);
      } else {
        ranges = t.ranges;
        loc = LOC[ubicacion];
      }
      const { type, switches } = splitFieldCode(codigo);
      if (type && wordSupports(5)) {
        const inserted = ranges.map((r) => r.insertField(loc, type as Word.FieldType, switches || undefined));
        await context.sync();
        let actualizado = true;
        try {
          inserted.forEach((f) => f.updateResult());
          await context.sync();
        } catch {
          actualizado = false;
        }
        return { insertados: inserted.length, metodo: "API", actualizado };
      }
      // Sin API de campos (o tipo desconocido): campo en OOXML marcado para actualizarse.
      const xml = wrapOoxml(fieldParagraph(codigo));
      for (const r of ranges) r.insertOoxml(xml, loc);
      await context.sync();
      return {
        insertados: ranges.length,
        metodo: "OOXML",
        nota: "Si el campo muestra «Actualice el campo», el usuario debe seleccionarlo y presionar F9 (clic derecho > Actualizar campo).",
      };
    });
  },
});

export const insertarVinculoONota = defineTool({
  name: "insertar_vinculo_o_nota",
  description:
    "Sobre el objetivo (normalmente un texto buscado): hipervínculo (URL, mailto: o #marcador para un vínculo " +
    "interno), quitar hipervínculo, nota al pie, nota al final, comentario o marcador (para referencias cruzadas " +
    "con insertar_campo REF/PAGEREF). Las notas y comentarios se anclan al final del texto indicado.",
  schema: z.object({
    objetivo: Objetivo,
    tipo: z.enum(["hipervinculo", "quitar_hipervinculo", "nota_al_pie", "nota_al_final", "comentario", "marcador"]),
    valor: z
      .string()
      .optional()
      .describe("URL, texto de la nota o del comentario, o nombre del marcador (letra inicial, sin espacios, máx. 40)."),
  }),
  writes: true,
  async run({ objetivo, tipo, valor }) {
    if (tipo !== "quitar_hipervinculo" && !valor) throw new Error(`Falta \`valor\` para ${tipo}.`);
    if (tipo === "marcador" && !/^\p{L}[\p{L}\p{N}_]{0,39}$/u.test(valor as string)) {
      throw new Error("Nombre de marcador inválido: empieza con letra, sin espacios ni signos, máximo 40 caracteres.");
    }
    if (tipo === "comentario" || tipo === "marcador") requireWord(4, "WordApi", tipo);
    if (tipo === "nota_al_pie" || tipo === "nota_al_final") requireWord(5, "WordApi", "Insertar notas");
    return Word.run(async (context) => {
      const t = await resolveTarget(context, objetivo);
      const ranges = tipo === "marcador" ? t.ranges.slice(0, 1) : t.ranges;
      for (const r of ranges) {
        if (tipo === "hipervinculo") r.hyperlink = valor as string;
        else if (tipo === "quitar_hipervinculo") r.hyperlink = "";
        else if (tipo === "nota_al_pie") r.insertFootnote(valor);
        else if (tipo === "nota_al_final") r.insertEndnote(valor);
        else if (tipo === "comentario") r.insertComment(valor as string);
        else r.insertBookmark(valor as string);
      }
      await context.sync();
      return { aplicado: tipo, lugares: ranges.length };
    });
  },
});

const WRAP = {
  cuadrado: "Square",
  estrecho: "Tight",
  transparente: "Through",
  arriba_abajo: "TopBottom",
  detras: "Behind",
  delante: "Front",
} as const;

function findImage(images: ImageFile[], name: string): ImageFile {
  const exact = images.find((i) => i.name === name) ?? images.find((i) => i.name.toLowerCase() === name.toLowerCase());
  if (exact) return exact;
  if (!images.length) throw new Error("El usuario no adjuntó imágenes en esta conversación; pídele que adjunte el archivo.");
  throw new Error(`No hay una imagen adjunta llamada «${name}». Adjuntas: ${images.map((i) => `«${i.name}»`).join(", ")}.`);
}

export const insertarImagen = defineTool({
  name: "insertar_imagen",
  description:
    "Inserta una imagen que el usuario adjuntó en la conversación (PNG, JPG, GIF o BMP) en su propio párrafo " +
    "(antes o después del objetivo, o al inicio o final de la zona) o en lugar de un texto. Permite tamaño en cm " +
    "(con un solo lado se conserva la proporción), alineación, texto alternativo y ajuste del texto " +
    "(los ajustes distintos de en_linea requieren WordApiDesktop 1.2).",
  schema: z.object({
    adjunto: z.string().describe("Nombre del archivo de imagen adjunto, tal como aparece en la conversación."),
    objetivo: Objetivo.optional(),
    ubicacion: Ubicacion.default("final"),
    ancho_cm: z.number().positive().optional(),
    alto_cm: z.number().positive().optional(),
    alineacion: Alineacion.optional().describe("Alineación del párrafo de la imagen."),
    texto_alternativo: z.string().optional(),
    ajuste: z
      .enum(["en_linea", "cuadrado", "estrecho", "transparente", "arriba_abajo", "detras", "delante"])
      .default("en_linea"),
  }),
  writes: true,
  async run(input, env) {
    const image = findImage(env.images(), input.adjunto);
    if (!/png|jpe?g|gif|bmp/.test(image.mediaType)) {
      throw new Error(`Word no inserta imágenes ${image.mediaType}; pide la imagen en PNG o JPG.`);
    }
    if (input.ajuste !== "en_linea") requireWord(2, "WordApiDesktop", "El ajuste de texto de imágenes");
    return Word.run(async (context) => {
      const t = await resolveTarget(context, input.objetivo ?? DEFAULT_OBJETIVO);
      const width = input.ancho_cm !== undefined ? cmToPt(input.ancho_cm) : undefined;
      const height = input.alto_cm !== undefined ? cmToPt(input.alto_cm) : undefined;
      const inline = t.kind === "texto" || t.kind === "seleccion" || input.ubicacion === "reemplazar";

      let paragraph: Word.Paragraph | null = null;
      if (!inline) paragraph = await newParagraph(context, t, input.ubicacion);

      if (input.ajuste !== "en_linea") {
        const anchor = paragraph ? paragraph.getRange("Start") : t.ranges[0];
        const shape = anchor.insertPictureFromBase64(image.data, { width, height });
        shape.textWrap.type = WRAP[input.ajuste];
        if (input.texto_alternativo) shape.altTextDescription = input.texto_alternativo;
        if (paragraph && input.alineacion) paragraph.alignment = ALINEACION[input.alineacion];
        await context.sync();
        return { insertada: image.name, ajuste: input.ajuste };
      }

      const pictures = paragraph
        ? [paragraph.insertInlinePictureFromBase64(image.data, "Start")]
        : t.kind === "parrafos"
          ? (t.paragraphs as Word.Paragraph[]).map((p) => p.insertInlinePictureFromBase64(image.data, "Replace"))
          : t.ranges.map((r) => r.insertInlinePictureFromBase64(image.data, LOC[input.ubicacion]));
      for (const pic of pictures) {
        pic.lockAspectRatio = !(width && height);
        if (width) pic.width = width;
        if (height) pic.height = height;
        if (input.texto_alternativo) pic.altTextDescription = input.texto_alternativo;
        if (input.alineacion) pic.paragraph.alignment = ALINEACION[input.alineacion];
      }
      await context.sync();
      return { insertada: image.name, veces: pictures.length };
    });
  },
});
