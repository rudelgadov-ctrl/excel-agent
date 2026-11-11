import { z } from "zod";
import { columnToIndex } from "../../excel/address";
import { getRange, getSheet, supports } from "../../excel/office";
import { MOSTRAR_COMO, MOSTRAR_COMO_KEYS } from "../../excel/pivot";
import { defineTool } from "../types";

function requireApi(minor: number, feature: string): void {
  if (!supports(minor)) throw new Error(`${feature} requiere ExcelApi 1.${minor}; esta versión de Excel no la tiene.`);
}

/** Índice (base 0) de una columna dentro de un rango, a partir de su letra en la hoja. */
async function columnOffset(context: Excel.RequestContext, range: Excel.Range, letter: string): Promise<number> {
  range.load("address,columnIndex,columnCount");
  await context.sync();
  const offset = columnToIndex(letter) - range.columnIndex;
  if (offset < 0 || offset >= range.columnCount) throw new Error(`La columna ${letter} no está dentro de ${range.address}.`);
  return offset;
}

const Letter = z.string().regex(/^[A-Za-z]{1,3}$/, "Letra de columna, p. ej. \"C\"");

// ---------- formato condicional ----------

const CfFormat = z
  .object({
    color_fuente: z.string().optional(),
    color_relleno: z.string().optional(),
    negrita: z.boolean().optional(),
    cursiva: z.boolean().optional(),
    formato_numero: z.string().optional(),
  })
  .describe("Formato que se aplica cuando se cumple la regla. Colores en hex.");

type CfFormatInput = z.infer<typeof CfFormat>;

function applyCfFormat(format: Excel.ConditionalRangeFormat, f: CfFormatInput | undefined): void {
  const fmt = f ?? { color_fuente: "#9C0006", color_relleno: "#FFC7CE" };
  if (fmt.color_fuente) format.font.color = fmt.color_fuente;
  if (fmt.color_relleno) format.fill.color = fmt.color_relleno;
  if (fmt.negrita !== undefined) format.font.bold = fmt.negrita;
  if (fmt.cursiva !== undefined) format.font.italic = fmt.cursiva;
  if (fmt.formato_numero) format.numberFormat = fmt.formato_numero;
}

export const formatoCondicional = defineTool({
  name: "formato_condicional",
  description:
    "Agrega una regla de formato condicional a un rango (o borra todas con tipo=borrar). Tipos: valor_celda " +
    "(operador + valor1/valor2), texto (contiene/empieza/termina), top_bottom (10 mejores, 10 % inferior…), " +
    "formula (personalizada, en inglés y relativa a la celda superior izquierda: =$D2>100), predefinido " +
    "(duplicados, únicos, sobre el promedio, vacías, hoy…), escala_color, barra_datos, iconos.",
  schema: z.object({
    rango: z.string(),
    tipo: z.enum(["valor_celda", "texto", "top_bottom", "formula", "predefinido", "escala_color", "barra_datos", "iconos", "borrar"]),
    operador: z
      .enum(["Between", "NotBetween", "EqualTo", "NotEqualTo", "GreaterThan", "LessThan", "GreaterThanOrEqual", "LessThanOrEqual"])
      .optional()
      .describe("Para valor_celda."),
    valor1: z.union([z.string(), z.number()]).optional().describe("Número o fórmula (\"=$B$1\")."),
    valor2: z.union([z.string(), z.number()]).optional(),
    texto: z.string().optional(),
    operador_texto: z.enum(["Contains", "NotContains", "BeginsWith", "EndsWith"]).default("Contains"),
    top_bottom: z
      .object({ tipo: z.enum(["TopItems", "TopPercent", "BottomItems", "BottomPercent"]), cantidad: z.number().int().positive() })
      .optional(),
    formula: z.string().optional(),
    criterio: z
      .enum([
        "DuplicateValues",
        "UniqueValues",
        "AboveAverage",
        "BelowAverage",
        "EqualOrAboveAverage",
        "EqualOrBelowAverage",
        "Blanks",
        "NonBlanks",
        "Errors",
        "NonErrors",
        "Yesterday",
        "Today",
        "Tomorrow",
        "LastSevenDays",
        "LastWeek",
        "ThisWeek",
        "NextWeek",
        "LastMonth",
        "ThisMonth",
        "NextMonth",
      ])
      .optional()
      .describe("Para predefinido."),
    colores: z.array(z.string()).min(2).max(3).optional().describe("Escala de color: [mínimo, (medio), máximo]."),
    color_barra: z.string().optional(),
    estilo_iconos: z
      .enum([
        "ThreeArrows",
        "ThreeArrowsGray",
        "ThreeFlags",
        "ThreeTrafficLights1",
        "ThreeTrafficLights2",
        "ThreeSigns",
        "ThreeSymbols",
        "ThreeSymbols2",
        "ThreeStars",
        "ThreeTriangles",
        "FourArrows",
        "FourArrowsGray",
        "FourRedToBlack",
        "FourRating",
        "FourTrafficLights",
        "FiveArrows",
        "FiveArrowsGray",
        "FiveRating",
        "FiveQuarters",
        "FiveBoxes",
      ])
      .optional(),
    formato: CfFormat.optional(),
  }),
  writes: true,
  async run(input) {
    requireApi(6, "El formato condicional");
    return Excel.run(async (context) => {
      const range = getRange(context, input.rango);
      const cfs = range.conditionalFormats;
      const str = (v: string | number | undefined) => (v === undefined ? undefined : String(v));
      switch (input.tipo) {
        case "borrar":
          cfs.clearAll();
          break;
        case "valor_celda": {
          if (!input.operador || input.valor1 === undefined) throw new Error("Faltan operador y valor1.");
          const cf = cfs.add(Excel.ConditionalFormatType.cellValue);
          applyCfFormat(cf.cellValue.format, input.formato);
          cf.cellValue.rule = { formula1: String(input.valor1), formula2: str(input.valor2), operator: input.operador };
          break;
        }
        case "texto": {
          if (!input.texto) throw new Error("Falta texto.");
          const cf = cfs.add(Excel.ConditionalFormatType.containsText);
          applyCfFormat(cf.textComparison.format, input.formato);
          cf.textComparison.rule = { operator: input.operador_texto, text: input.texto };
          break;
        }
        case "top_bottom": {
          if (!input.top_bottom) throw new Error("Falta top_bottom.");
          const cf = cfs.add(Excel.ConditionalFormatType.topBottom);
          applyCfFormat(cf.topBottom.format, input.formato);
          cf.topBottom.rule = { type: input.top_bottom.tipo, rank: input.top_bottom.cantidad };
          break;
        }
        case "formula": {
          if (!input.formula) throw new Error("Falta formula.");
          const cf = cfs.add(Excel.ConditionalFormatType.custom);
          applyCfFormat(cf.custom.format, input.formato);
          cf.custom.rule.formula = input.formula;
          break;
        }
        case "predefinido": {
          if (!input.criterio) throw new Error("Falta criterio.");
          const cf = cfs.add(Excel.ConditionalFormatType.presetCriteria);
          applyCfFormat(cf.preset.format, input.formato);
          cf.preset.rule = { criterion: input.criterio };
          break;
        }
        case "escala_color": {
          const [min, a, b] = input.colores ?? ["#F8696B", "#FFEB84", "#63BE7B"];
          const max = b ?? a;
          const cf = cfs.add(Excel.ConditionalFormatType.colorScale);
          cf.colorScale.criteria = {
            minimum: { type: "LowestValue", formula: undefined, color: min },
            ...(b ? { midpoint: { type: "Percentile", formula: "50", color: a } } : {}),
            maximum: { type: "HighestValue", formula: undefined, color: max },
          };
          break;
        }
        case "barra_datos": {
          const cf = cfs.add(Excel.ConditionalFormatType.dataBar);
          if (input.color_barra) cf.dataBar.positiveFormat.fillColor = input.color_barra;
          break;
        }
        case "iconos": {
          const cf = cfs.add(Excel.ConditionalFormatType.iconSet);
          cf.iconSet.style = input.estilo_iconos ?? "ThreeTrafficLights1";
          break;
        }
      }
      range.load("address");
      await context.sync();
      return { rango: range.address, ok: true };
    });
  },
});

// ---------- tablas de Excel ----------

export const crearTabla = defineTool({
  name: "crear_tabla",
  description:
    "Convierte un rango en una tabla de Excel (Insertar > Tabla / Dar formato como tabla), con estilo y fila de totales opcionales.",
  schema: z.object({
    rango: z.string().describe("Rango con los datos, incluidos los encabezados."),
    tiene_encabezados: z.boolean().default(true),
    nombre: z.string().optional().describe("Nombre de la tabla, sin espacios."),
    estilo: z.string().optional().describe("Estilo, p. ej. \"TableStyleMedium2\", \"TableStyleLight9\"."),
    fila_totales: z.boolean().optional(),
  }),
  writes: true,
  async run({ rango, tiene_encabezados, nombre, estilo, fila_totales }) {
    return Excel.run(async (context) => {
      const range = getRange(context, rango);
      range.load("address");
      await context.sync();
      const table = context.workbook.tables.add(range.address, tiene_encabezados);
      if (nombre) table.name = nombre;
      if (estilo) table.style = estilo;
      if (fila_totales !== undefined) table.showTotals = fila_totales;
      table.load("name");
      const body = table.getRange();
      body.load("address");
      await context.sync();
      return { tabla: table.name, rango: body.address };
    });
  },
});

// ---------- tablas dinámicas ----------

const Aggregation = z.enum([
  "Sum",
  "Count",
  "Average",
  "Max",
  "Min",
  "Product",
  "CountNumbers",
  "StandardDeviation",
  "StandardDeviationP",
  "Variance",
  "VarianceP",
]);

const TOP_TIPOS = {
  superiores: { condition: "TopN", selection: "Items" },
  inferiores: { condition: "BottomN", selection: "Items" },
  porcentaje_superior: { condition: "TopN", selection: "Percent" },
  porcentaje_inferior: { condition: "BottomN", selection: "Percent" },
  suma_superior: { condition: "TopN", selection: "Sum" },
  suma_inferior: { condition: "BottomN", selection: "Sum" },
} as const;

export const tablaDinamica = defineTool({
  name: "tabla_dinamica",
  description:
    "Crea una tabla dinámica desde una tabla de Excel o un rango (los campos van por el texto exacto del encabezado). " +
    "Incluye mostrar valores como (% del total general, de la fila, acumulado, clasificación…), subtotales arriba/abajo, " +
    "totales generales, filtro Top N, filtrar elementos y ordenar. El destino puede estar en otra hoja (créala antes " +
    "con gestionar_hojas). Varias en la misma hoja: déjales espacio, porque crecen al filtrar. Para el modelo de datos " +
    "(campos de varias tablas relacionadas), agrupar fechas, estilos o recuento distinto usa macro_modelo_datos.",
  schema: z.object({
    nombre: z.string().describe("Nombre de la tabla dinámica, sin espacios."),
    origen: z.string().describe("Rango de datos con encabezados, o nombre de una tabla de Excel."),
    destino: z.string().describe("Celda superior izquierda, p. ej. \"'Resumen'!A3\"."),
    filas: z.array(z.string()).default([]),
    columnas: z.array(z.string()).default([]),
    valores: z
      .array(
        z.object({
          campo: z.string(),
          resumen: Aggregation.default("Sum"),
          nombre: z.string().optional().describe("Título del campo de valores (no puede ser igual a un encabezado)."),
          formato_numero: z.string().optional(),
          mostrar_como: z.enum(MOSTRAR_COMO_KEYS).optional(),
          campo_base: z.string().optional().describe("Para total acumulado y clasificar (por defecto, el primer campo de filas)."),
        }),
      )
      .default([]),
    filtros: z.array(z.string()).default([]),
    diseno: z.enum(["Compact", "Tabular", "Outline"]).optional(),
    subtotales: z.enum(["arriba", "abajo", "ninguno"]).optional(),
    total_general_abajo: z.boolean().optional().describe("Fila «Total general» al pie."),
    total_general_derecha: z.boolean().optional().describe("Columna «Total general» a la derecha."),
    top: z
      .object({
        campo: z.string().describe("Campo de filas o columnas."),
        cantidad: z.number().positive(),
        tipo: z.enum(Object.keys(TOP_TIPOS) as [keyof typeof TOP_TIPOS]).default("superiores"),
        valor: z.string().optional().describe("Título del valor por el que se clasifica (por defecto el primero)."),
      })
      .optional(),
    filtrar_elementos: z
      .array(z.object({ campo: z.string(), elementos: z.array(z.string()).min(1).describe("Elementos que quedan visibles.") }))
      .optional(),
    ordenar: z
      .array(
        z.object({
          campo: z.string(),
          descendente: z.boolean().default(true),
          por_valor: z.string().optional().describe("Título del valor; sin él ordena por etiqueta."),
        }),
      )
      .optional(),
    reemplazar: z.boolean().default(false).describe("Si ya existe una tabla dinámica con ese nombre, la borra antes."),
  }),
  writes: true,
  async run(input) {
    requireApi(8, "Las tablas dinámicas");
    return Excel.run(async (context) => {
      const avisos: string[] = [];
      if (input.reemplazar) {
        const old = context.workbook.pivotTables.getItemOrNullObject(input.nombre);
        await context.sync();
        if (!old.isNullObject) {
          old.delete();
          await context.sync();
        }
      }
      const dest = getRange(context, input.destino);
      const table = context.workbook.tables.getItemOrNullObject(input.origen);
      await context.sync();
      const source: Excel.Range | Excel.Table = table.isNullObject ? getRange(context, input.origen) : table;
      const pivot = dest.worksheet.pivotTables.add(input.nombre, source, dest);
      const fieldOf = (name: string) => pivot.hierarchies.getItem(name).fields.getItem(name);
      for (const f of input.filas) pivot.rowHierarchies.add(pivot.hierarchies.getItem(f));
      for (const f of input.columnas) pivot.columnHierarchies.add(pivot.hierarchies.getItem(f));
      const data: Excel.DataPivotHierarchy[] = [];
      for (const v of input.valores) {
        const dh = pivot.dataHierarchies.add(pivot.hierarchies.getItem(v.campo));
        dh.summarizeBy = v.resumen;
        if (v.nombre) dh.name = v.nombre;
        if (v.formato_numero) dh.numberFormat = v.formato_numero;
        if (v.mostrar_como && v.mostrar_como !== "sin_calculo") {
          const rule = MOSTRAR_COMO[v.mostrar_como];
          const base = v.campo_base ?? input.filas[0] ?? input.columnas[0];
          if (rule.base && !base) throw new Error(`"${v.mostrar_como}" necesita campo_base o un campo en filas.`);
          dh.showAs = {
            calculation: rule.officeJs as Excel.ShowAsCalculation,
            ...(rule.base ? { baseField: fieldOf(base!) } : {}),
          };
        }
        data.push(dh);
      }
      for (const f of input.filtros) pivot.filterHierarchies.add(pivot.hierarchies.getItem(f));
      if (input.diseno) pivot.layout.layoutType = input.diseno;
      if (input.subtotales) pivot.layout.subtotalLocation = { arriba: "AtTop", abajo: "AtBottom", ninguno: "Off" }[input.subtotales] as Excel.SubtotalLocationType;
      if (input.total_general_abajo !== undefined) pivot.layout.showColumnGrandTotals = input.total_general_abajo;
      if (input.total_general_derecha !== undefined) pivot.layout.showRowGrandTotals = input.total_general_derecha;
      await context.sync();
      data.forEach((dh) => dh.load("name"));
      await context.sync();
      const valueNames = data.map((dh) => dh.name);
      const valueByName = (name: string | undefined, what: string) => {
        if (!data.length) throw new Error(`${what} necesita al menos un campo de valores.`);
        if (!name) return data[0];
        const i = valueNames.findIndex((n) => n.toLowerCase() === name.toLowerCase());
        if (i < 0) throw new Error(`${what}: no hay un valor "${name}". Valores: ${valueNames.join(", ")}.`);
        return data[i];
      };

      if (input.top || input.filtrar_elementos?.length) {
        if (!supports(12)) {
          avisos.push("Los filtros de tabla dinámica requieren ExcelApi 1.12: usa macro_modelo_datos (top) o filtra a mano.");
        } else {
          if (input.top) {
            const t = TOP_TIPOS[input.top.tipo];
            const dh = valueByName(input.top.valor, "El filtro top");
            fieldOf(input.top.campo).applyFilter({
              valueFilter: {
                condition: t.condition as Excel.ValueFilterCondition,
                comparator: input.top.cantidad,
                value: dh.name,
                selectionType: t.selection as Excel.TopBottomSelectionType,
              },
            });
          }
          for (const f of input.filtrar_elementos ?? []) {
            fieldOf(f.campo).applyFilter({ manualFilter: { selectedItems: f.elementos } });
          }
        }
      }
      for (const o of input.ordenar ?? []) {
        const order = o.descendente ? Excel.SortBy.descending : Excel.SortBy.ascending;
        if (o.por_valor) {
          if (!supports(9)) avisos.push("Ordenar por valor requiere ExcelApi 1.9.");
          else fieldOf(o.campo).sortByValues(order, valueByName(o.por_valor, "Ordenar"));
        } else {
          fieldOf(o.campo).sortByLabels(order);
        }
      }
      await context.sync();
      const area = pivot.layout.getRange();
      area.load("address");
      await context.sync();
      return {
        tabla_dinamica: input.nombre,
        rango: area.address,
        valores: valueNames,
        avisos: avisos.length ? avisos : undefined,
      };
    });
  },
});

// ---------- filtros ----------

export const filtrar = defineTool({
  name: "filtrar",
  description:
    "Autofiltro sobre un rango con encabezados: filtra una columna por valores, por criterio personalizado " +
    "(\">100\", \"=Activo\", \"*texto*\"), o por los N mayores/menores. tipo=quitar elimina el autofiltro de la hoja.",
  schema: z.object({
    rango: z.string().describe("Rango de datos con encabezados."),
    columna: Letter.optional().describe("Columna a filtrar (letra de la hoja)."),
    tipo: z.enum(["valores", "personalizado", "top", "quitar"]),
    valores: z.array(z.string()).optional(),
    criterio1: z.string().optional(),
    criterio2: z.string().optional(),
    operador: z.enum(["And", "Or"]).optional(),
    top: z
      .object({ tipo: z.enum(["TopItems", "TopPercent", "BottomItems", "BottomPercent"]), cantidad: z.number().int().positive() })
      .optional(),
  }),
  writes: true,
  async run(input) {
    requireApi(9, "El autofiltro");
    return Excel.run(async (context) => {
      const range = getRange(context, input.rango);
      const autoFilter = range.worksheet.autoFilter;
      if (input.tipo === "quitar") {
        autoFilter.remove();
        await context.sync();
        return { ok: true, filtro: "quitado" };
      }
      if (!input.columna) throw new Error("Falta la columna.");
      const index = await columnOffset(context, range, input.columna);
      let criteria: Excel.FilterCriteria;
      if (input.tipo === "valores") {
        if (!input.valores?.length) throw new Error("Faltan valores.");
        criteria = { filterOn: "Values", values: input.valores };
      } else if (input.tipo === "personalizado") {
        if (!input.criterio1) throw new Error("Falta criterio1.");
        criteria = { filterOn: "Custom", criterion1: input.criterio1, criterion2: input.criterio2, operator: input.operador };
      } else {
        if (!input.top) throw new Error("Falta top.");
        criteria = { filterOn: input.top.tipo, criterion1: String(input.top.cantidad) };
      }
      autoFilter.apply(range, index, criteria);
      await context.sync();
      return { rango: range.address, ok: true };
    });
  },
});

// ---------- filas, columnas y duplicados ----------

export const insertarEliminar = defineTool({
  name: "insertar_eliminar",
  description:
    "Inserta o elimina celdas, filas o columnas. Filas completas: \"'Hoja'!5:7\"; columnas completas: \"'Hoja'!C:D\". " +
    "Al insertar, lo existente se desplaza hacia abajo (filas) o a la derecha (columnas).",
  schema: z.object({
    accion: z.enum(["insertar", "eliminar"]),
    rango: z.string(),
    desplazar: z
      .enum(["Down", "Right", "Up", "Left"])
      .optional()
      .describe("Solo para celdas sueltas. Insertar: Down/Right. Eliminar: Up/Left."),
  }),
  writes: true,
  async run({ accion, rango, desplazar }) {
    return Excel.run(async (context) => {
      const range = getRange(context, rango);
      const address = rango.split("!").pop() ?? rango;
      const isColumns = /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/.test(address);
      if (accion === "insertar") {
        const shift = (desplazar ?? (isColumns ? "Right" : "Down")) as Excel.InsertShiftDirection;
        const inserted = range.insert(shift);
        inserted.load("address");
        await context.sync();
        return { insertado: inserted.address };
      }
      const shift = (desplazar ?? (isColumns ? "Left" : "Up")) as Excel.DeleteShiftDirection;
      range.load("address");
      await context.sync();
      const address2 = range.address;
      range.delete(shift);
      await context.sync();
      return { eliminado: address2 };
    });
  },
});

export const quitarDuplicados = defineTool({
  name: "quitar_duplicados",
  description: "Quita filas duplicadas de un rango comparando las columnas indicadas (Datos > Quitar duplicados).",
  schema: z.object({
    rango: z.string(),
    columnas: z.array(Letter).min(1).describe("Columnas que definen un duplicado (letras de la hoja)."),
    tiene_encabezados: z.boolean().default(true),
  }),
  writes: true,
  async run({ rango, columnas, tiene_encabezados }) {
    requireApi(9, "Quitar duplicados");
    return Excel.run(async (context) => {
      const range = getRange(context, rango);
      const indexes: number[] = [];
      for (const c of columnas) indexes.push(await columnOffset(context, range, c));
      const result = range.removeDuplicates(indexes, tiene_encabezados);
      result.load("removed,uniqueRemaining");
      await context.sync();
      return { eliminadas: result.removed, filas_unicas: result.uniqueRemaining };
    });
  },
});

// ---------- configuración de hoja ----------

export const configurarHoja = defineTool({
  name: "configurar_hoja",
  description:
    "Ajustes de una hoja: inmovilizar paneles, color de pestaña, líneas de cuadrícula, ocultar/mostrar, " +
    "proteger/desproteger y configuración de impresión (orientación, área de impresión, ajustar a una página, centrar).",
  schema: z.object({
    hoja: z.string(),
    inmovilizar_filas: z.number().int().min(0).optional().describe("Filas superiores fijas (0 quita)."),
    inmovilizar_columnas: z.number().int().min(0).optional().describe("Columnas izquierdas fijas (0 quita)."),
    color_pestana: z.string().optional().describe("Hex, o \"\" para quitar."),
    cuadricula: z.boolean().optional(),
    visible: z.boolean().optional(),
    proteger: z.boolean().optional().describe("true protege la hoja; false la desprotege."),
    contrasena: z.string().optional(),
    orientacion: z.enum(["Portrait", "Landscape"]).optional(),
    area_impresion: z.string().optional().describe("Rango, p. ej. \"A1:F40\"."),
    ajustar_una_pagina: z.boolean().optional(),
    centrar_horizontal: z.boolean().optional(),
  }),
  writes: true,
  async run(input) {
    return Excel.run(async (context) => {
      const ws = getSheet(context, input.hoja);
      const hechos: string[] = [];
      if (input.inmovilizar_filas !== undefined || input.inmovilizar_columnas !== undefined) {
        requireApi(7, "Inmovilizar paneles");
        const rows = input.inmovilizar_filas ?? 0;
        const cols = input.inmovilizar_columnas ?? 0;
        ws.freezePanes.unfreeze();
        if (rows > 0 && cols > 0) ws.freezePanes.freezeAt(ws.getRangeByIndexes(0, 0, rows, cols));
        else if (rows > 0) ws.freezePanes.freezeRows(rows);
        else if (cols > 0) ws.freezePanes.freezeColumns(cols);
        hechos.push("paneles");
      }
      if (input.color_pestana !== undefined) {
        requireApi(7, "El color de pestaña");
        ws.tabColor = input.color_pestana;
        hechos.push("color de pestaña");
      }
      if (input.cuadricula !== undefined) {
        requireApi(8, "Ocultar la cuadrícula");
        ws.showGridlines = input.cuadricula;
        hechos.push("cuadrícula");
      }
      if (input.visible !== undefined) {
        ws.visibility = input.visible ? "Visible" : "Hidden";
        hechos.push("visibilidad");
      }
      if (input.orientacion || input.area_impresion || input.ajustar_una_pagina !== undefined || input.centrar_horizontal !== undefined) {
        requireApi(9, "La configuración de impresión");
        const page = ws.pageLayout;
        if (input.orientacion) page.orientation = input.orientacion;
        if (input.area_impresion) page.setPrintArea(input.area_impresion);
        if (input.ajustar_una_pagina) page.zoom = { horizontalFitToPages: 1, verticalFitToPages: 1 };
        if (input.centrar_horizontal !== undefined) page.centerHorizontally = input.centrar_horizontal;
        hechos.push("impresión");
      }
      // La protección va al final: una hoja protegida no acepta los demás cambios.
      if (input.proteger === false) {
        ws.protection.unprotect(input.contrasena);
        hechos.push("desprotegida");
      }
      await context.sync();
      if (input.proteger === true) {
        ws.protection.protect(undefined, input.contrasena);
        await context.sync();
        hechos.push("protegida");
      }
      return { hoja: input.hoja, cambios: hechos };
    });
  },
});
