import { z } from "zod";
import { columnToIndex } from "../excel/address";
import { getRange, supports } from "../excel/office";
import { defineTool } from "./types";

export const crearNombre = defineTool({
  name: "crear_nombre",
  description:
    "Asigna un nombre definido (alcance de libro) a un rango para usarlo en fórmulas (=PMT(Tasa/12,Plazo,-Monto)). " +
    "Si el nombre ya existe, lo reemplaza. Nombres sin espacios ni acentos raros, que no parezcan una celda (no \"A1\").",
  schema: z.object({
    nombre: z.string().regex(/^[A-Za-z_À-ſ][\w.À-ſ]*$/, "Nombre inválido para Excel"),
    rango: z.string(),
    comentario: z.string().optional(),
  }),
  writes: true,
  async run({ nombre, rango, comentario }) {
    return Excel.run(async (context) => {
      const existing = context.workbook.names.getItemOrNullObject(nombre);
      await context.sync();
      if (!existing.isNullObject) existing.delete();
      const item = context.workbook.names.add(nombre, getRange(context, rango), comentario);
      item.load("name,value");
      await context.sync();
      return { nombre: item.name, referencia: item.value };
    });
  },
});

export const gestionarHojas = defineTool({
  name: "gestionar_hojas",
  description: "Crea, copia, renombra, elimina o activa hojas del libro.",
  schema: z.object({
    accion: z.enum(["crear", "copiar", "renombrar", "eliminar", "activar"]),
    hoja: z.string().describe("Hoja sobre la que se actúa (o nombre de la nueva hoja al crear)."),
    nuevo_nombre: z.string().optional().describe("Para copiar o renombrar."),
  }),
  writes: true,
  async run({ accion, hoja, nuevo_nombre }) {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      switch (accion) {
        case "crear": {
          const ws = sheets.add(hoja);
          ws.load("name");
          await context.sync();
          return { creada: ws.name };
        }
        case "copiar": {
          if (!supports(10)) throw new Error("Copiar hojas requiere ExcelApi 1.10; usa copiar_rango sobre una hoja nueva.");
          const copy = sheets.getItem(hoja).copy("End");
          if (nuevo_nombre) copy.name = nuevo_nombre;
          copy.load("name");
          await context.sync();
          return { copia: copy.name };
        }
        case "renombrar": {
          if (!nuevo_nombre) throw new Error("Falta nuevo_nombre.");
          sheets.getItem(hoja).name = nuevo_nombre;
          await context.sync();
          return { renombrada: nuevo_nombre };
        }
        case "eliminar":
          sheets.getItem(hoja).delete();
          await context.sync();
          return { eliminada: hoja };
        case "activar":
          sheets.getItem(hoja).activate();
          await context.sync();
          return { activa: hoja };
      }
    });
  },
});

const Operator = z.enum([
  "Between",
  "NotBetween",
  "EqualTo",
  "NotEqualTo",
  "GreaterThan",
  "LessThan",
  "GreaterThanOrEqualTo",
  "LessThanOrEqualTo",
]);

export const validacionDatos = defineTool({
  name: "validacion_datos",
  description:
    "Configura la validación de datos de un rango. Para listas, `lista` puede ser una referencia con '=' " +
    "(p. ej. \"='Catálogo'!$A$2:$A$50\") o valores separados por comas (\"Sí,No\"). " +
    "`permitir_vacio` deja la celda en blanco como valor válido. tipo=\"ninguna\" quita la validación.",
  schema: z.object({
    rango: z.string(),
    tipo: z.enum(["lista", "entero", "decimal", "fecha", "longitud_texto", "personalizada", "ninguna"]),
    lista: z.string().optional(),
    operador: Operator.optional(),
    valor1: z.union([z.string(), z.number()]).optional(),
    valor2: z.union([z.string(), z.number()]).optional(),
    formula: z.string().optional().describe("Para tipo personalizada, en inglés: =ISNUMBER(A1)"),
    permitir_vacio: z.boolean().default(true),
    mensaje_error: z.string().optional(),
    titulo_error: z.string().optional(),
    mensaje_entrada: z.string().optional(),
  }),
  writes: true,
  async run(input) {
    if (!supports(8)) throw new Error("La validación de datos requiere ExcelApi 1.8.");
    return Excel.run(async (context) => {
      const range = getRange(context, input.rango);
      const dv = range.dataValidation;
      dv.clear();
      if (input.tipo === "ninguna") {
        await context.sync();
        return { ok: true, validacion: "eliminada" };
      }
      const basic = () => ({
        formula1: input.valor1 ?? "",
        formula2: input.valor2,
        operator: input.operador ?? "Between",
      });
      let rule: Excel.DataValidationRule;
      switch (input.tipo) {
        case "lista":
          if (!input.lista) throw new Error("Falta `lista`.");
          rule = { list: { inCellDropDown: true, source: input.lista } };
          break;
        case "entero":
          rule = { wholeNumber: basic() };
          break;
        case "decimal":
          rule = { decimal: basic() };
          break;
        case "fecha": {
          const { formula1, formula2, operator } = basic();
          rule = {
            date: { formula1: String(formula1), formula2: formula2 === undefined ? undefined : String(formula2), operator },
          };
          break;
        }
        case "longitud_texto":
          rule = { textLength: basic() };
          break;
        case "personalizada":
          if (!input.formula) throw new Error("Falta `formula`.");
          rule = { custom: { formula: input.formula } };
          break;
      }
      dv.rule = rule;
      dv.ignoreBlanks = input.permitir_vacio;
      if (input.mensaje_error) {
        dv.errorAlert = {
          message: input.mensaje_error,
          showAlert: true,
          style: "Stop",
          title: input.titulo_error ?? "Valor no válido",
        };
      }
      if (input.mensaje_entrada) {
        dv.prompt = { message: input.mensaje_entrada, showPrompt: true, title: "" };
      }
      range.load("address");
      await context.sync();
      return { rango: range.address, ok: true };
    });
  },
});

export const ordenar = defineTool({
  name: "ordenar",
  description:
    "Ordena las filas de un rango por una o más columnas (letras de columna de la hoja, p. ej. \"C\"). " +
    "Incluye en el rango todas las columnas que deben moverse juntas.",
  schema: z.object({
    rango: z.string().describe("Rango de datos completo, p. ej. \"'Mi hoja'!A1:D50\"."),
    claves: z
      .array(z.object({ columna: z.string().regex(/^[A-Za-z]{1,3}$/), ascendente: z.boolean().default(true) }))
      .min(1),
    tiene_encabezados: z.boolean().default(false).describe("true si la primera fila del rango es de títulos."),
  }),
  writes: true,
  async run({ rango, claves, tiene_encabezados }) {
    return Excel.run(async (context) => {
      const range = getRange(context, rango);
      range.load("address,columnIndex,columnCount");
      await context.sync();
      const fields = claves.map((k) => {
        const key = columnToIndex(k.columna) - range.columnIndex;
        if (key < 0 || key >= range.columnCount) throw new Error(`La columna ${k.columna} no está dentro de ${range.address}.`);
        return { key, ascending: k.ascendente };
      });
      range.sort.apply(fields, false, tiene_encabezados, "Rows");
      await context.sync();
      return { rango: range.address, ok: true };
    });
  },
});
