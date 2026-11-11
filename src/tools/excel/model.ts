import { z } from "zod";
import { parseRef } from "../../excel/address";
import { buildMacro, FILTROS_TOP, type MacroSpec, PERIODOS, TIPOS_GRAFICO_DINAMICO } from "../../excel/macro";
import { supports } from "../../excel/office";
import { MOSTRAR_COMO_KEYS, parseFieldRef, RESUMENES } from "../../excel/pivot";
import { defineTool } from "../types";

const Campo = z.string().describe('Modelo: "Tabla[Columna]". Tabla dinámica desde una tabla o rango: "Columna".');

const Valor = z.object({
  campo: Campo,
  resumen: z.enum(RESUMENES).default("Sum").describe("DistinctCount (recuento distinto) solo en el modelo."),
  nombre: z.string().optional().describe('Título del valor; por defecto "Suma de X", "Recuento de X"…'),
  formato_numero: z.string().optional().describe('Código en-US, p. ej. "$#,##0.00" o "0.00%".'),
  mostrar_como: z.enum(MOSTRAR_COMO_KEYS).optional().describe("Mostrar valores como."),
  campo_base: Campo.optional().describe("Para total acumulado y clasificar (por defecto, el primer campo de filas)."),
});

const TablaDinamica = z.object({
  nombre: z.string().regex(/^[^[\]'"!]{1,60}$/, "Nombre sin corchetes, comillas ni !"),
  destino: z.string().describe("Celda superior izquierda con hoja, p. ej. \"'Resumen'!A3\". Si la hoja no existe, la macro la crea."),
  origen: z
    .string()
    .default("modelo")
    .describe('"modelo" (modelo de datos) o el nombre de una tabla de Excel o rango con encabezados.'),
  filas: z.array(Campo).default([]),
  columnas: z.array(Campo).default([]),
  filtros: z.array(Campo).default([]),
  valores: z.array(Valor).default([]),
  agrupar_fechas: z
    .array(z.object({ campo: Campo, periodos: z.array(z.enum(PERIODOS)).min(1) }))
    .optional()
    .describe("Solo origen tabla/rango (en el modelo usa columnas calculadas). El campo debe estar en filas o columnas."),
  subtotales: z
    .enum(["arriba", "abajo", "ninguno"])
    .optional()
    .describe("Con 2 o más campos de fila el valor por defecto es arriba, como en la interfaz."),
  total_general_abajo: z.boolean().optional().describe("Fila «Total general» al pie (totales de columnas)."),
  total_general_derecha: z.boolean().optional().describe("Columna «Total general» a la derecha (totales de filas)."),
  diseno: z.enum(["compacto", "esquema", "tabular"]).optional(),
  encabezado_filas: z.string().optional().describe("Texto en lugar de «Etiquetas de fila» (diseño compacto)."),
  estilo: z.string().optional().describe('Estilo de tabla dinámica, p. ej. "PivotStyleMedium9".'),
  top: z
    .object({
      campo: Campo.describe("Campo de filas o columnas a filtrar."),
      cantidad: z.number().int().positive(),
      tipo: z.enum(Object.keys(FILTROS_TOP) as [keyof typeof FILTROS_TOP]).default("superiores"),
      valor: z.string().optional().describe("Título del valor por el que se ordena (por defecto el primero)."),
    })
    .optional()
    .describe("Filtro «10 mejores» (o inferiores, porcentaje, suma)."),
  ordenar: z
    .array(
      z.object({
        campo: Campo,
        descendente: z.boolean().default(true),
        por_valor: z.string().optional().describe("Título del valor; sin él ordena por etiqueta."),
      }),
    )
    .optional(),
});

export const macroModeloDatos = defineTool({
  name: "macro_modelo_datos",
  description:
    "Genera y muestra una macro VBA probada, para que el usuario la ejecute una vez, que hace lo que Office.js no permite: " +
    "agregar tablas al modelo de datos (Power Pivot), crear relaciones, tablas dinámicas desde el modelo (campos de " +
    "varias tablas relacionadas, recuento distinto), agrupar fechas (años, trimestres, meses) en tablas dinámicas, " +
    "filtros Top N, mostrar valores como, subtotales arriba, segmentadores con varias columnas o conectados a varias " +
    "tablas dinámicas (conexiones de informe), escalas de tiempo, gráficos dinámicos y vincular el texto de una forma a " +
    "una celda. Las tablas del modelo toman el nombre de la tabla de Excel: renómbralas antes con modificar_tabla. " +
    "Valida contra el libro abierto que existan tablas, columnas y formas. Prepara antes con Office.js todo lo demás " +
    "(datos, limpieza, columnas calculadas, hojas, formas con nombre).",
  schema: z.object({
    nombre_macro: z.string().default("CrearModeloYTablasDinamicas"),
    descripcion: z.string().optional().describe("Qué hace la macro, en una o dos frases, para el usuario."),
    tablas_modelo: z.array(z.string()).default([]).describe("Tablas de Excel que se agregan al modelo de datos."),
    relaciones: z
      .array(
        z.object({
          foranea: z.string().describe('Lado "varios": "Tabla[Columna]" que repite valores (clave foránea).'),
          primaria: z.string().describe('Lado "uno": "Tabla[Columna]" con valores únicos (clave primaria).'),
        }),
      )
      .default([]),
    tablas_dinamicas: z.array(TablaDinamica).default([]),
    segmentadores: z
      .array(
        z.object({
          campo: Campo,
          tablas_dinamicas: z.array(z.string()).min(1).describe("La primera es el origen; las demás quedan conectadas."),
          rango: z.string().describe("Celdas que ocupará, con hoja: \"'Panel'!B6:D13\"."),
          titulo: z.string().optional(),
          columnas: z.number().int().min(1).max(20).optional(),
          estilo: z.string().optional().describe('P. ej. "SlicerStyleLight1".'),
          escala_tiempo: z.boolean().default(false).describe("Escala de tiempo (cronología) en vez de segmentador; campo de fecha."),
          nombre: z.string().optional(),
        }),
      )
      .default([]),
    graficos: z
      .array(
        z.object({
          tabla_dinamica: z.string(),
          rango: z.string().describe("Celdas que ocupará, con hoja (puede ser otra hoja que la de la tabla dinámica)."),
          tipo: z.enum(TIPOS_GRAFICO_DINAMICO).default("ColumnClustered"),
          titulo: z.string().optional(),
          leyenda: z.boolean().default(true),
          etiquetas_datos: z.boolean().default(false),
          nombre: z.string().optional(),
        }),
      )
      .default([])
      .describe("Gráficos dinámicos de tablas dinámicas creadas por esta macro o ya existentes."),
    vinculos_formas: z
      .array(
        z.object({
          hoja: z.string(),
          forma: z.string().describe("Nombre de la forma o cuadro de texto (también dentro de un grupo)."),
          celda: z.string().describe("Celda con hoja cuyo valor mostrará, p. ej. \"'Calculos'!E5\"."),
        }),
      )
      .default([]),
    codigo_extra: z
      .string()
      .optional()
      .describe("Instrucciones VBA adicionales (sin Sub/Function) para lo que no cubren los campos; puede usar TD(nombre) y Hoja(nombre)."),
  }),
  writes: false,
  async run(input, env) {
    const spec = input as MacroSpec;
    const avisos = await checkAgainstWorkbook(spec);
    const result = buildMacro(spec);
    env.showVba({
      nombre: spec.nombre_macro,
      codigo: result.codigo,
      descripcion: input.descripcion,
      soloWindows: result.usaModelo,
      comoEjecutar:
        "Ejecútela una sola vez, con este libro abierto. Al terminar muestra un resumen; si algún paso falla, copie el " +
        "mensaje (Ctrl+C) y péguelo en el chat. Después puede guardar como .xlsx: Excel avisa que quitará la macro y " +
        "el modelo, las tablas dinámicas, los segmentadores y los gráficos se conservan.",
    });
    return {
      mostrado: true,
      pasos: result.pasos,
      valores_por_tabla_dinamica: result.valores,
      solo_windows: result.usaModelo,
      avisos: avisos.length ? avisos : undefined,
      siguiente:
        "Pide al usuario que ejecute la macro y te avise. Luego verifica con leer_libro (tablas dinámicas, " +
        "segmentadores, gráficos) y completa lo que dependa de su resultado.",
    };
  },
});

interface WorkbookInfo {
  tables: Map<string, { name: string; headers: string[] }>;
  sheets: string[];
  pivots: Set<string>;
  shapes: Map<string, Set<string>>;
}

async function loadWorkbookInfo(): Promise<WorkbookInfo> {
  return Excel.run(async (context) => {
    const wb = context.workbook;
    const tables = wb.tables;
    tables.load("items/name");
    const sheets = wb.worksheets;
    sheets.load("items/name");
    await context.sync();
    const headers = tables.items.map((t) => {
      const h = t.getHeaderRowRange();
      h.load("values");
      return h;
    });
    const pivots = supports(8) ? wb.pivotTables : null;
    pivots?.load("items/name");
    const shapes = supports(9)
      ? sheets.items.map((ws) => {
          const s = ws.shapes;
          s.load("items/name,items/type");
          return s;
        })
      : [];
    await context.sync();
    const groups: { sheet: string; shapes: Excel.GroupShapeCollection }[] = [];
    shapes.forEach((s, i) =>
      s.items
        .filter((shape) => shape.type === Excel.ShapeType.group)
        .forEach((shape) => {
          const inner = shape.group.shapes;
          inner.load("items/name");
          groups.push({ sheet: sheets.items[i].name, shapes: inner });
        }),
    );
    if (groups.length) await context.sync();
    const shapeMap = new Map<string, Set<string>>();
    shapes.forEach((s, i) => shapeMap.set(sheets.items[i].name.toLowerCase(), new Set(s.items.map((x) => x.name.toLowerCase()))));
    for (const g of groups) for (const x of g.shapes.items) shapeMap.get(g.sheet.toLowerCase())?.add(x.name.toLowerCase());
    return {
      tables: new Map(
        tables.items.map((t, i) => [t.name.toLowerCase(), { name: t.name, headers: (headers[i].values[0] ?? []).map(String) }]),
      ),
      sheets: sheets.items.map((s) => s.name),
      pivots: new Set((pivots?.items ?? []).map((p) => p.name.toLowerCase())),
      shapes: shapeMap,
    };
  });
}

/** Comprueba tablas, columnas, tablas dinámicas y formas. Lanza un error si algo no existe; devuelve avisos. */
async function checkAgainstWorkbook(spec: MacroSpec): Promise<string[]> {
  let info: WorkbookInfo;
  try {
    info = await loadWorkbookInfo();
  } catch {
    return ["No se pudo revisar el libro antes de generar la macro; verifica los nombres de tablas y columnas."];
  }
  const errors: string[] = [];
  const avisos: string[] = [];
  const tableList = () => [...info.tables.values()].map((t) => t.name).join(", ") || "(ninguna)";
  const modelTables = new Set(spec.tablas_modelo.map((t) => t.toLowerCase()));

  for (const t of spec.tablas_modelo) {
    if (!info.tables.has(t.toLowerCase())) errors.push(`No existe la tabla de Excel "${t}". Tablas: ${tableList()}.`);
  }
  const checkColumn = (table: string, column: string, where: string) => {
    const t = info.tables.get(table.toLowerCase());
    if (!t) {
      if (!modelTables.has(table.toLowerCase())) {
        avisos.push(`${where}: "${table}" no es una tabla de Excel del libro; se supone que ya está en el modelo de datos.`);
      }
      return;
    }
    if (t.headers.includes(column)) return;
    const trimmed = t.headers.find((h) => h.trim().toLowerCase() === column.trim().toLowerCase());
    errors.push(
      trimmed !== undefined
        ? `${where}: la columna de "${t.name}" se llama ${JSON.stringify(trimmed)} (con espacios o mayúsculas distintas); ` +
            "limpia los encabezados con limpiar_datos o usa el nombre exacto."
        : `${where}: "${t.name}" no tiene la columna "${column}". Columnas: ${t.headers.map((h) => JSON.stringify(h)).join(", ")}.`,
    );
  };
  const checkModelRef = (ref: string, where: string) => {
    const f = parseFieldRef(ref);
    if (f.tabla) checkColumn(f.tabla, f.columna, where);
  };

  for (const r of spec.relaciones) {
    checkModelRef(r.foranea, "Relación");
    checkModelRef(r.primaria, "Relación");
  }
  const specPivots = new Set(spec.tablas_dinamicas.map((p) => p.nombre.toLowerCase()));
  for (const pt of spec.tablas_dinamicas) {
    const where = `Tabla dinámica "${pt.nombre}"`;
    const refs = [
      ...pt.filas,
      ...pt.columnas,
      ...pt.filtros,
      ...pt.valores.map((v) => v.campo),
      ...pt.valores.flatMap((v) => (v.campo_base ? [v.campo_base] : [])),
    ];
    if (pt.origen.trim().toLowerCase() === "modelo") {
      for (const ref of refs) checkModelRef(ref, where);
    } else if (!parseRef(pt.origen).sheet) {
      const source = info.tables.get(pt.origen.toLowerCase());
      if (!source) errors.push(`${where}: no existe la tabla de origen "${pt.origen}". Tablas: ${tableList()}.`);
      else for (const ref of refs) checkColumn(source.name, parseFieldRef(ref).columna, where);
    }
    if (info.pivots.has(pt.nombre.toLowerCase())) avisos.push(`${where} ya existe: la macro la reemplazará.`);
  }
  const pivotExists = (name: string) => specPivots.has(name.toLowerCase()) || info.pivots.has(name.toLowerCase());
  for (const s of spec.segmentadores) {
    for (const p of s.tablas_dinamicas) if (!pivotExists(p)) errors.push(`Segmentador de "${s.campo}": no existe la tabla dinámica "${p}".`);
  }
  for (const g of spec.graficos) {
    if (!pivotExists(g.tabla_dinamica)) errors.push(`Gráfico: no existe la tabla dinámica "${g.tabla_dinamica}".`);
  }
  for (const v of spec.vinculos_formas) {
    const shapes = info.shapes.get(v.hoja.toLowerCase());
    if (!info.sheets.some((s) => s.toLowerCase() === v.hoja.toLowerCase())) errors.push(`Vínculo: no existe la hoja "${v.hoja}".`);
    else if (shapes && !shapes.has(v.forma.toLowerCase())) {
      errors.push(`Vínculo: no hay una forma llamada "${v.forma}" en "${v.hoja}". Créala antes con gestionar_formas (y su nombre).`);
    }
  }
  if (errors.length) throw new Error(`No se generó la macro:\n- ${errors.join("\n- ")}`);
  return avisos;
}
