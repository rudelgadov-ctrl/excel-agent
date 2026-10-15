import { z } from "zod";
import { defineTool } from "../types";

// Constructor de funciones async para ejecutar el código generado por el agente.
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor as new (
  ...args: string[]
) => (context: Word.RequestContext, word: typeof Word) => Promise<unknown>;

export const ejecutarOfficeJs = defineTool({
  name: "ejecutar_office_js",
  description:
    "Último recurso: ejecuta código Office.js dentro de Word.run cuando ninguna otra herramienta sirve (p. ej. " +
    "propiedades del documento, control de cambios, controles de contenido, cuadros de texto o formas, " +
    "bordes de página, tabla de contenido con la API). El código es el CUERPO de una función async que recibe " +
    "`context` (Word.RequestContext) y `Word`; debe llamar `await context.sync()` y puede devolver datos simples " +
    "(JSON) con `return`. No devuelvas objetos proxy de Office; copia antes sus propiedades cargadas.",
  schema: z.object({
    descripcion: z.string().describe("Qué hace el código, para mostrarlo al usuario."),
    codigo: z.string(),
  }),
  writes: true,
  async run({ codigo }) {
    const fn = new AsyncFunction("context", "Word", codigo);
    return Word.run(async (context) => {
      const result = await fn(context, Word);
      await context.sync();
      return result === undefined ? { ok: true } : result;
    });
  },
});
