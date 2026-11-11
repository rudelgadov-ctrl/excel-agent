import { z } from "zod";
import { parseRef } from "../../excel/address";
import { getRange, getSheet, supports } from "../../excel/office";
import { defineTool } from "../types";

function requireApi(minor: number, feature: string): void {
  if (!supports(minor)) throw new Error(`${feature} requiere ExcelApi 1.${minor}; esta versión de Excel no la tiene.`);
}

/** Rango de posición: admite "B6:D13" (en `hoja`) o "'Otra'!B6:D13". */
function positionRange(context: Excel.RequestContext, hoja: string, rango: string): Excel.Range {
  return parseRef(rango).sheet ? getRange(context, rango) : getSheet(context, hoja).getRange(rango);
}

async function placeAt(
  context: Excel.RequestContext,
  target: { left: number; top: number; width: number; height: number },
  range: Excel.Range,
): Promise<void> {
  range.load("left,top,width,height");
  await context.sync();
  target.left = range.left;
  target.top = range.top;
  target.width = range.width;
  target.height = range.height;
}

// ---------- segmentadores ----------

export const segmentacion = defineTool({
  name: "segmentacion",
  description:
    "Segmentadores (segmentación de datos) con Office.js: crear sobre una tabla dinámica o una tabla de Excel, " +
    "ubicarlo sobre un rango de celdas, título, estilo y orden; seleccionar elementos, quitar el filtro o eliminar. " +
    "Office.js no permite varias columnas, conexiones de informe (un segmentador para varias tablas dinámicas), " +
    "escalas de tiempo ni segmentadores de tablas del modelo de datos: para eso usa macro_modelo_datos.",
  schema: z.object({
    accion: z.enum(["crear", "seleccionar", "quitar_filtro", "eliminar"]),
    nombre: z.string().optional().describe("Nombre del segmentador (obligatorio salvo al crear)."),
    origen: z.string().optional().describe("Al crear: nombre de la tabla dinámica o de la tabla de Excel."),
    campo: z.string().optional().describe("Al crear: encabezado del campo que filtra."),
    hoja: z.string().optional().describe("Al crear: hoja donde se coloca (por defecto la del origen)."),
    rango: z.string().optional().describe("Celdas que ocupará, p. ej. \"B6:D13\"."),
    titulo: z.string().optional(),
    estilo: z.string().optional().describe('P. ej. "SlicerStyleLight1"…"SlicerStyleLight6", "SlicerStyleDark1".'),
    orden: z.enum(["DataSourceOrder", "Ascending", "Descending"]).optional(),
    elementos: z.array(z.string()).optional().describe("Para seleccionar: elementos que quedan activos."),
  }),
  writes: true,
  async run(input) {
    requireApi(10, "Los segmentadores");
    return Excel.run(async (context) => {
      const wb = context.workbook;
      let slicer: Excel.Slicer;
      if (input.accion === "crear") {
        if (!input.origen || !input.campo) throw new Error("Para crear indica origen y campo.");
        const pivot = wb.pivotTables.getItemOrNullObject(input.origen);
        const table = wb.tables.getItemOrNullObject(input.origen);
        await context.sync();
        if (pivot.isNullObject && table.isNullObject) throw new Error(`No existe la tabla dinámica ni la tabla "${input.origen}".`);
        const source = pivot.isNullObject ? table : pivot;
        let sheetName = input.hoja;
        if (!sheetName) {
          const ws = pivot.isNullObject ? table.worksheet : pivot.worksheet;
          ws.load("name");
          await context.sync();
          sheetName = ws.name;
        }
        slicer = wb.slicers.add(source, input.campo, sheetName);
        if (input.nombre) slicer.name = input.nombre;
        if (input.rango) {
          const pos = { left: 0, top: 0, width: 0, height: 0 };
          await placeAt(context, pos, positionRange(context, sheetName, input.rango));
          slicer.set(pos);
        }
      } else {
        if (!input.nombre) throw new Error("Indica el nombre del segmentador.");
        slicer = wb.slicers.getItem(input.nombre);
        if (input.accion === "eliminar") {
          slicer.delete();
          await context.sync();
          return { eliminado: input.nombre };
        }
        if (input.accion === "quitar_filtro") slicer.clearFilters();
        if (input.accion === "seleccionar") {
          if (!input.elementos?.length) throw new Error("Indica los elementos a seleccionar.");
          slicer.selectItems(input.elementos);
        }
      }
      if (input.titulo) slicer.caption = input.titulo;
      if (input.estilo) slicer.style = input.estilo;
      if (input.orden) slicer.sortBy = input.orden;
      slicer.load("name,caption,left,top,width,height");
      await context.sync();
      return {
        segmentador: slicer.name,
        titulo: slicer.caption,
        posicion: { izquierda: slicer.left, arriba: slicer.top, ancho: slicer.width, alto: slicer.height },
      };
    });
  },
});

// ---------- formas ----------

const SHAPE_TYPES: Record<string, string> = {
  rectangulo: "Rectangle",
  rectangulo_redondeado: "RoundRectangle",
  elipse: "Ellipse",
  triangulo: "Triangle",
  flecha_derecha: "RightArrow",
  flecha_izquierda: "LeftArrow",
  pentagono: "Pentagon",
  chevron: "Chevron",
};

const Font = z
  .object({
    nombre: z.string().optional(),
    tamano: z.number().positive().optional(),
    color: z.string().optional(),
    negrita: z.boolean().optional(),
    cursiva: z.boolean().optional(),
  })
  .optional();

/** Busca una forma por nombre en la hoja, también dentro de grupos. */
async function findShape(context: Excel.RequestContext, ws: Excel.Worksheet, name: string): Promise<Excel.Shape> {
  const direct = ws.shapes.getItemOrNullObject(name);
  ws.shapes.load("items/name,items/type");
  await context.sync();
  if (!direct.isNullObject) return direct;
  const groups = ws.shapes.items.filter((s) => s.type === Excel.ShapeType.group);
  const inner = groups.map((g) => {
    const c = g.group.shapes;
    c.load("items/name");
    return c;
  });
  await context.sync();
  for (const c of inner) {
    const found = c.items.find((s) => s.name.toLowerCase() === name.toLowerCase());
    if (found) return found;
  }
  throw new Error(`No hay una forma llamada "${name}". Formas: ${ws.shapes.items.map((s) => s.name).join(", ") || "(ninguna)"}.`);
}

export const gestionarFormas = defineTool({
  name: "gestionar_formas",
  description:
    "Formas y cuadros de texto (Insertar > Formas): crear, modificar, agrupar o eliminar. Sirve para indicadores " +
    "(KPI) de un dashboard, títulos y fondos: relleno, transparencia, borde, texto, fuente y alineación, ubicada " +
    "sobre un rango de celdas. Ponle nombre para poder referirte a ella. Office.js no puede vincular el texto de una " +
    "forma a una celda (=Hoja!B5): para eso usa macro_modelo_datos con vinculos_formas; mientras tanto escribe el valor actual.",
  schema: z.object({
    accion: z.enum(["crear", "modificar", "agrupar", "eliminar"]),
    hoja: z.string(),
    nombre: z.string().optional().describe("Nombre de la forma (o del grupo al agrupar). Obligatorio salvo al crear."),
    tipo: z
      .string()
      .default("rectangulo")
      .describe(
        `Al crear: cuadro_texto, ${Object.keys(SHAPE_TYPES).join(", ")}, u otro valor de Excel.GeometricShapeType.`,
      ),
    rango: z.string().optional().describe("Celdas que ocupará, p. ej. \"K5:M10\"."),
    texto: z.string().optional(),
    relleno: z.string().optional().describe('Color hex o "ninguno".'),
    transparencia: z.number().min(0).max(1).optional(),
    borde: z.object({ color: z.string().optional(), grosor: z.number().min(0).optional(), ninguno: z.boolean().optional() }).optional(),
    fuente: Font,
    alineacion_horizontal: z.enum(["Left", "Center", "Right", "Justify"]).optional(),
    alineacion_vertical: z.enum(["Top", "Middle", "Bottom"]).optional(),
    formas: z.array(z.string()).optional().describe("Para agrupar: nombres de las formas."),
    al_frente: z.boolean().optional().describe("Traer al frente."),
  }),
  writes: true,
  async run(input) {
    requireApi(9, "Las formas");
    return Excel.run(async (context) => {
      const ws = getSheet(context, input.hoja);
      let shape: Excel.Shape;
      switch (input.accion) {
        case "eliminar": {
          if (!input.nombre) throw new Error("Indica el nombre.");
          (await findShape(context, ws, input.nombre)).delete();
          await context.sync();
          return { eliminada: input.nombre };
        }
        case "agrupar": {
          if (!input.formas || input.formas.length < 2) throw new Error("Indica al menos dos formas para agrupar.");
          shape = ws.shapes.addGroup(input.formas);
          if (input.nombre) shape.name = input.nombre;
          shape.load("name");
          await context.sync();
          return { grupo: shape.name, formas: input.formas };
        }
        case "crear": {
          if (input.tipo === "cuadro_texto") {
            shape = ws.shapes.addTextBox(input.texto ?? "");
          } else {
            const type = SHAPE_TYPES[input.tipo] ?? input.tipo;
            shape = ws.shapes.addGeometricShape(type as Excel.GeometricShapeType);
            if (input.texto !== undefined) shape.textFrame.textRange.text = input.texto;
          }
          if (input.nombre) shape.name = input.nombre;
          break;
        }
        case "modificar": {
          if (!input.nombre) throw new Error("Indica el nombre.");
          shape = await findShape(context, ws, input.nombre);
          if (input.texto !== undefined) shape.textFrame.textRange.text = input.texto;
          break;
        }
      }
      if (input.rango) {
        const pos = { left: 0, top: 0, width: 0, height: 0 };
        await placeAt(context, pos, positionRange(context, input.hoja, input.rango));
        shape.set(pos);
      }
      if (input.relleno) {
        if (input.relleno.toLowerCase() === "ninguno") shape.fill.clear();
        else shape.fill.setSolidColor(input.relleno);
      }
      if (input.transparencia !== undefined) shape.fill.transparency = input.transparencia;
      if (input.borde) {
        if (input.borde.ninguno) shape.lineFormat.visible = false;
        else {
          shape.lineFormat.visible = true;
          if (input.borde.color) shape.lineFormat.color = input.borde.color;
          if (input.borde.grosor !== undefined) shape.lineFormat.weight = input.borde.grosor;
        }
      }
      if (input.fuente) {
        const font = shape.textFrame.textRange.font;
        if (input.fuente.nombre) font.name = input.fuente.nombre;
        if (input.fuente.tamano) font.size = input.fuente.tamano;
        if (input.fuente.color) font.color = input.fuente.color;
        if (input.fuente.negrita !== undefined) font.bold = input.fuente.negrita;
        if (input.fuente.cursiva !== undefined) font.italic = input.fuente.cursiva;
      }
      if (input.alineacion_horizontal) shape.textFrame.horizontalAlignment = input.alineacion_horizontal;
      if (input.alineacion_vertical) shape.textFrame.verticalAlignment = input.alineacion_vertical;
      if (input.al_frente) shape.setZOrder(Excel.ShapeZOrder.bringToFront);
      shape.load("name,left,top,width,height");
      await context.sync();
      return {
        forma: shape.name,
        posicion: { izquierda: shape.left, arriba: shape.top, ancho: shape.width, alto: shape.height },
      };
    });
  },
});
