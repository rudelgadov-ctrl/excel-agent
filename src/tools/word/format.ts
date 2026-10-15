import { z } from "zod";
import { applyStyle, requireWord, resolveTarget, targetParagraphs } from "../../word/office";
import { styleNameCandidates } from "../../word/spec";
import { defineTool } from "../types";
import {
  applyFont,
  applyParagraphFormat,
  fuenteExtraShape,
  fuenteShape,
  hasAny,
  Objetivo,
  parrafoShape,
} from "./common";

export const formatoTexto = defineTool({
  name: "formato_texto",
  description:
    "Formato de caracteres sobre el objetivo (palabras o frases buscadas, párrafos, tabla, selección o todo): " +
    "fuente, tamaño, negrita, cursiva, subrayado, color, resaltado, tachado, superíndice/subíndice, mayúsculas, " +
    "versalitas, espaciado; o un estilo (de carácter o de párrafo). Para resaltar todas las apariciones de una " +
    "palabra usa objetivo.texto sin ocurrencia.",
  schema: z.object({
    objetivo: Objetivo,
    ...fuenteShape,
    ...fuenteExtraShape,
    estilo: z.string().optional().describe("Estilo integrado (\"Strong\", \"Emphasis\", \"Heading2\"…) o nombre de un estilo."),
  }),
  writes: true,
  async run({ objetivo, estilo, ...font }) {
    return Word.run(async (context) => {
      const t = await resolveTarget(context, objetivo);
      const warnings: string[] = [];
      for (const r of t.ranges) {
        if (estilo) applyStyle(r, estilo);
        applyFont(r.font, font, warnings);
      }
      await context.sync();
      return { aplicado_en: t.ranges.length, avisos: warnings.length ? [...new Set(warnings)] : undefined };
    });
  },
});

export const formatoParrafo = defineTool({
  name: "formato_parrafo",
  description:
    "Formato de párrafo sobre los párrafos del objetivo: estilo (Heading1, Title, Normal… o el nombre local), " +
    "alineación, sangrías en cm (izquierda, derecha, primera línea o francesa), espaciado antes/después en pt, " +
    "interlineado múltiple (1, 1.15, 1.5, 2) y sombreado. Si el objetivo es un texto, aplica a los párrafos que lo contienen.",
  schema: z.object({
    objetivo: Objetivo,
    estilo: z.string().optional(),
    ...parrafoShape,
    sombreado: z.string().optional().describe('Color de fondo del párrafo ("#DEEAF6"); requiere WordApiDesktop 1.3.'),
  }),
  writes: true,
  async run({ objetivo, estilo, sombreado, ...format }) {
    if (sombreado !== undefined) requireWord(3, "WordApiDesktop", "El sombreado de párrafo");
    return Word.run(async (context) => {
      const t = await resolveTarget(context, objetivo);
      const paragraphs = await targetParagraphs(context, t);
      for (const p of paragraphs) {
        // El estilo primero: aplicarlo después borraría el formato directo.
        if (estilo) applyStyle(p, estilo);
        applyParagraphFormat(p, format);
        if (sombreado !== undefined) p.shading.backgroundPatternColor = sombreado;
      }
      await context.sync();
      return { parrafos: paragraphs.length };
    });
  },
});

const VINETAS = {
  solida: "Solid",
  hueca: "Hollow",
  cuadrado: "Square",
  rombos: "Diamonds",
  flecha: "Arrow",
  check: "Checkmark",
} as const;
const NUMERACION = { "1": "Arabic", I: "UpperRoman", i: "LowerRoman", A: "UpperLetter", a: "LowerLetter" } as const;

export const aplicarLista = defineTool({
  name: "aplicar_lista",
  description:
    "Convierte los párrafos del objetivo en una lista con viñetas o numerada (1, I, i, A, a; con \".\" o \")\"), " +
    "en el nivel indicado (2 = subnivel con sangría), o les quita la lista. Omite los párrafos vacíos. " +
    "Para listas de varios niveles, aplica primero la lista a todos y luego cambia el nivel de los subpuntos.",
  schema: z.object({
    objetivo: Objetivo,
    tipo: z.enum(["viñetas", "numeros", "nivel", "quitar"]).describe("nivel = solo cambiar el nivel de párrafos que ya son de una lista."),
    vineta: z.enum(["solida", "hueca", "cuadrado", "rombos", "flecha", "check"]).default("solida"),
    numeracion: z.enum(["1", "I", "i", "A", "a"]).default("1"),
    sufijo: z.enum([".", ")"]).default("."),
    nivel: z.number().int().min(1).max(9).default(1),
  }),
  writes: true,
  async run({ objetivo, tipo, vineta, numeracion, sufijo, nivel }) {
    return Word.run(async (context) => {
      const t = await resolveTarget(context, objetivo);
      const all = await targetParagraphs(context, t);
      for (const p of all) p.load("text,isListItem");
      await context.sync();
      const level = nivel - 1;

      if (tipo === "quitar" || tipo === "nivel") {
        const items = all.filter((p) => p.isListItem);
        for (const p of items) {
          if (tipo === "quitar") p.detachFromList();
          else p.listItem.level = level;
        }
        await context.sync();
        return { parrafos: items.length };
      }

      const items = all.filter((p) => p.text.trim() !== "");
      if (!items.length) throw new Error("El objetivo no tiene párrafos con texto.");
      for (const p of items) if (p.isListItem) p.detachFromList();
      const list = items[0].startNewList();
      list.load("id");
      await context.sync();
      for (const p of items.slice(1)) p.attachToList(list.id, level);
      if (level > 0) items[0].listItem.level = level;
      if (tipo === "viñetas") list.setLevelBullet(level, VINETAS[vineta]);
      else list.setLevelNumbering(level, NUMERACION[numeracion], [level, sufijo]);
      await context.sync();
      return { parrafos: items.length, lista: list.id };
    });
  },
});

const TIPO_ESTILO = { parrafo: "Paragraph", caracter: "Character", tabla: "Table" } as const;

export const modificarEstilo = defineTool({
  name: "modificar_estilo",
  description:
    "Modifica (o crea) un estilo del documento: fuente y formato de párrafo, estilo base y estilo siguiente. " +
    "Todos los párrafos con ese estilo cambian a la vez (p. ej. \"modifique el estilo Título 1\"). Requiere WordApi 1.5.",
  schema: z.object({
    nombre: z.string().describe("Nombre del estilo como lo muestra Word (\"Título 1\") o su nombre integrado (\"Heading1\")."),
    crear: z.enum(["parrafo", "caracter", "tabla"]).optional().describe("Crea el estilo con este tipo si no existe."),
    basado_en: z.string().optional(),
    siguiente: z.string().optional().describe("Estilo del párrafo siguiente al presionar Enter."),
    ...fuenteShape,
    ...parrafoShape,
  }),
  writes: true,
  async run({ nombre, crear, basado_en, siguiente, ...rest }) {
    requireWord(5, "WordApi", "Modificar estilos");
    return Word.run(async (context) => {
      const styles = context.document.getStyles();
      const candidates = styleNameCandidates(nombre).map((n) => {
        const s = styles.getByNameOrNullObject(n);
        s.load("nameLocal,type");
        return s;
      });
      await context.sync();
      let style = candidates.find((s) => !s.isNullObject);
      let creado = false;
      if (!style) {
        if (!crear) {
          styles.load("items/nameLocal,items/inUse");
          await context.sync();
          const inUse = styles.items.filter((s) => s.inUse).map((s) => s.nameLocal);
          throw new Error(`No existe el estilo «${nombre}». Estilos en uso: ${inUse.slice(0, 60).join(", ")}. Usa crear para crearlo.`);
        }
        style = context.document.addStyle(nombre, TIPO_ESTILO[crear]);
        creado = true;
      }
      if (basado_en) style.baseStyle = basado_en;
      if (siguiente) style.nextParagraphStyle = siguiente;
      const warnings: string[] = [];
      applyFont(style.font, rest, warnings);
      const paragraphKeys = Object.keys(parrafoShape);
      if (hasAny(rest, paragraphKeys)) applyParagraphFormat(style.paragraphFormat, rest);
      await context.sync();
      return { estilo: creado ? nombre : style.nameLocal, creado, avisos: warnings.length ? warnings : undefined };
    });
  },
});

