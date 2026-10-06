import { z } from "zod";
import { defineTool } from "./types";

export const entregarVba = defineTool({
  name: "entregar_vba",
  description:
    "Muestra al usuario un módulo VBA completo (con botón Copiar e instrucciones para Windows y Mac). " +
    "Úsala cuando el requerimiento pida grabar o crear una macro: el complemento no puede crear macros por sí mismo. " +
    "El código debe ser un módulo estándar completo (Option Explicit + Sub NombreMacro()) que use las celdas reales del libro.",
  schema: z.object({
    nombre_macro: z.string().describe("Nombre del Sub, p. ej. \"Facturar\"."),
    codigo: z.string().describe("Código VBA completo del módulo."),
    descripcion: z.string().optional().describe("Qué hace la macro, en una o dos frases."),
    como_ejecutar: z.string().optional().describe("Indicación extra (p. ej. asignarla a un botón)."),
  }),
  writes: false,
  async run({ nombre_macro, codigo, descripcion, como_ejecutar }, env) {
    env.showVba({ nombre: nombre_macro, codigo, descripcion, comoEjecutar: como_ejecutar });
    return { mostrado: true, nota: "El usuario ve el código con instrucciones para pegarlo y guardar como .xlsm." };
  },
});

// Constructor de funciones async para ejecutar el código generado por el agente.
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor as new (
  ...args: string[]
) => (context: Excel.RequestContext, excel: typeof Excel) => Promise<unknown>;

export const ejecutarOfficeJs = defineTool({
  name: "ejecutar_office_js",
  description:
    "Último recurso: ejecuta código Office.js dentro de Excel.run cuando ninguna otra herramienta sirve. " +
    "El código es el CUERPO de una función async que recibe `context` (Excel.RequestContext) y `Excel`; " +
    "debe llamar `await context.sync()` y puede devolver datos simples (JSON) con `return`. " +
    "No devuelvas objetos proxy de Office; copia antes sus propiedades cargadas.",
  schema: z.object({
    descripcion: z.string().describe("Qué hace el código, para mostrarlo al usuario."),
    codigo: z.string(),
  }),
  writes: true,
  async run({ codigo }) {
    const fn = new AsyncFunction("context", "Excel", codigo);
    return Excel.run(async (context) => {
      const result = await fn(context, Excel);
      await context.sync();
      return result === undefined ? { ok: true } : result;
    });
  },
});
