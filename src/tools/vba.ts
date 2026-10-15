import { z } from "zod";
import { defineTool } from "./types";

// Compartida por Excel y Word: el panel muestra los pasos de instalación de la aplicación abierta.
export const entregarVba = defineTool({
  name: "entregar_vba",
  description:
    "Muestra al usuario un módulo VBA completo (con botón Copiar e instrucciones para Windows y Mac). " +
    "Úsala cuando el requerimiento pida grabar o crear una macro, o algo que la API del complemento no permite: " +
    "el complemento no puede crear macros por sí mismo. El código debe ser un módulo estándar completo " +
    "(Option Explicit + Sub NombreMacro()) que use los objetos reales del archivo abierto (hojas y celdas, o " +
    "párrafos, tablas y secciones).",
  schema: z.object({
    nombre_macro: z.string().describe("Nombre del Sub, el que pida la instrucción, p. ej. \"ActualizarTotales\"."),
    codigo: z.string().describe("Código VBA completo del módulo."),
    descripcion: z.string().optional().describe("Qué hace la macro, en una o dos frases."),
    como_ejecutar: z.string().optional().describe("Indicación extra (p. ej. asignarla a un botón)."),
  }),
  writes: false,
  async run({ nombre_macro, codigo, descripcion, como_ejecutar }, env) {
    env.showVba({ nombre: nombre_macro, codigo, descripcion, comoEjecutar: como_ejecutar });
    return { mostrado: true, nota: "El usuario ve el código con instrucciones para pegarlo y guardar el archivo con macros." };
  },
});
