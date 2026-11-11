// Generador de la macro VBA para lo que Office.js no permite: modelo de datos (Power Pivot), relaciones,
// tablas dinámicas desde el modelo, agrupar fechas, segmentadores con columnas o conectados a varias tablas
// dinámicas, escalas de tiempo, gráficos dinámicos y formas vinculadas a celdas.
// Es puro (sin Office.js) para poder probarlo; las llamadas al modelo de objetos se validaron en Excel 365.
import { parseRef } from "./address";
import { defaultCaption, MOSTRAR_COMO, type MostrarComo, parseFieldRef, type Resumen, uniqueCaptions } from "./pivot";

export const PERIODOS = ["Dias", "Meses", "Trimestres", "Anios"] as const;
export type Periodo = (typeof PERIODOS)[number];

export const TIPOS_GRAFICO_DINAMICO = [
  "ColumnClustered",
  "ColumnStacked",
  "ColumnStacked100",
  "3DColumnClustered",
  "3DColumn",
  "BarClustered",
  "BarStacked",
  "BarStacked100",
  "3DBarClustered",
  "Line",
  "LineMarkers",
  "LineStacked",
  "Pie",
  "3DPie",
  "PieExploded",
  "Doughnut",
  "Area",
  "AreaStacked",
  "Radar",
  "RadarMarkers",
  "RadarFilled",
] as const;
export type TipoGrafico = (typeof TIPOS_GRAFICO_DINAMICO)[number];

export const FILTROS_TOP = {
  superiores: "xlTopCount",
  inferiores: "xlBottomCount",
  porcentaje_superior: "xlTopPercent",
  porcentaje_inferior: "xlBottomPercent",
  suma_superior: "xlTopSum",
  suma_inferior: "xlBottomSum",
} as const;
export type FiltroTop = keyof typeof FILTROS_TOP;

export interface ValorSpec {
  campo: string;
  resumen: Resumen;
  nombre?: string;
  formato_numero?: string;
  mostrar_como?: MostrarComo;
  campo_base?: string;
}

export interface TablaDinamicaSpec {
  nombre: string;
  destino: string;
  /** "modelo" o el nombre de una tabla de Excel / rango con encabezados. */
  origen: string;
  filas: string[];
  columnas: string[];
  filtros: string[];
  valores: ValorSpec[];
  agrupar_fechas?: { campo: string; periodos: Periodo[] }[];
  subtotales?: "arriba" | "abajo" | "ninguno";
  total_general_abajo?: boolean;
  total_general_derecha?: boolean;
  diseno?: "compacto" | "esquema" | "tabular";
  encabezado_filas?: string;
  estilo?: string;
  top?: { campo: string; cantidad: number; tipo: FiltroTop; valor?: string };
  ordenar?: { campo: string; descendente: boolean; por_valor?: string }[];
}

export interface SegmentadorSpec {
  campo: string;
  tablas_dinamicas: string[];
  rango: string;
  titulo?: string;
  columnas?: number;
  estilo?: string;
  escala_tiempo?: boolean;
  nombre?: string;
}

export interface GraficoSpec {
  tabla_dinamica: string;
  rango: string;
  tipo: TipoGrafico;
  titulo?: string;
  leyenda: boolean;
  etiquetas_datos: boolean;
  nombre?: string;
}

export interface VinculoSpec {
  hoja: string;
  forma: string;
  celda: string;
}

export interface MacroSpec {
  nombre_macro: string;
  tablas_modelo: string[];
  relaciones: { foranea: string; primaria: string }[];
  tablas_dinamicas: TablaDinamicaSpec[];
  segmentadores: SegmentadorSpec[];
  graficos: GraficoSpec[];
  vinculos_formas: VinculoSpec[];
  codigo_extra?: string;
}

export interface MacroResult {
  codigo: string;
  /** Títulos de los pasos, en orden. */
  pasos: string[];
  /** Títulos de los campos de valores de cada tabla dinámica (para referirse a ellos). */
  valores: Record<string, string[]>;
  /** true si usa el modelo de datos (solo Excel para Windows). */
  usaModelo: boolean;
}

// ---------- utilidades de texto VBA ----------

/** Literal de cadena VBA. Los caracteres fuera de Latin-1 van con ChrW() para que el editor no los dañe. */
export function vbaString(text: string): string {
  const parts: string[] = [];
  let lit = "";
  const flush = () => {
    if (lit) parts.push(`"${lit}"`);
    lit = "";
  };
  for (const ch of text) {
    const code = ch.codePointAt(0)!;
    if ((code >= 0x20 && code < 0x7f) || (code >= 0xa0 && code <= 0xff)) {
      lit += ch === '"' ? '""' : ch;
      continue;
    }
    flush();
    if (ch === "\n") parts.push("vbLf");
    else if (ch === "\t") parts.push("vbTab");
    else if (ch === "\r") continue;
    else for (let i = 0; i < ch.length; i++) parts.push(`ChrW(${ch.charCodeAt(i)})`);
  }
  flush();
  return parts.length ? parts.join(" & ") : '""';
}

/** Comentario VBA de una línea (sin saltos). */
function comment(text: string): string {
  return `' ${text.replace(/\r?\n/g, " ")}`;
}

/** Nombre de objeto válido (segmentadores, gráficos): letras, números y _. */
export function objectName(prefix: string, base: string): string {
  const clean = base
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^A-Za-z0-9_]+/g, "_")
    .replace(/^_+|_+$/g, "");
  return `${prefix}_${clean || "1"}`.slice(0, 60);
}

const mdx = (name: string) => `[${name.replace(/]/g, "]]")}]`;

// ---------- referencias ----------

interface CellRef {
  sheet: string;
  address: string;
}

function cellRef(ref: string, what: string): CellRef {
  const { sheet, address } = parseRef(ref);
  if (!sheet) throw new Error(`${what}: indica la hoja, p. ej. "'Resumen'!B4" (recibido "${ref}").`);
  return { sheet, address };
}

const rangeExpr = (r: CellRef) => `Hoja(${vbaString(r.sheet)}).Range(${vbaString(r.address)})`;

function absoluteFormula(ref: string): string {
  const { sheet, address } = cellRef(ref, "celda");
  const abs = address
    .split(":")
    .map((part) => part.replace(/^([A-Za-z]{1,3})(\d+)$/, "$$$1$$$2"))
    .join(":");
  return `='${sheet.replace(/'/g, "''")}'!${abs}`;
}

const RESERVED = new Set(
  ["Hoja", "TD", "BuscarTD", "BorrarTD", "Anotar", "AgregarAlModelo", "CrearRelacion", "ExisteTablaModelo", "BuscarTabla"]
    .concat(["CrearTDClasica", "Subtotales", "BorrarSegmentacion", "BuscarForma", "BorrarForma"])
    .map((s) => s.toLowerCase()),
);

// ---------- pasos ----------

interface Step {
  titulo: string;
  dims: string[];
  cuerpo: string[];
}

interface PivotInfo {
  modelo: boolean;
  captions: string[];
  /** Columna -> nombre del campo en VBA (PivotFields). */
  field: (ref: string) => string;
}

function pivotFieldName(modelo: boolean, ref: string, origen: string, where: string): string {
  const f = parseFieldRef(ref);
  if (modelo) {
    if (!f.tabla) throw new Error(`${where}: en una tabla dinámica del modelo el campo va como "Tabla[Columna]" (recibido "${ref}").`);
    return `${mdx(f.tabla)}.${mdx(f.columna)}.${mdx(f.columna)}`;
  }
  if (f.tabla && origen.toLowerCase() !== f.tabla.toLowerCase()) {
    throw new Error(`${where}: el campo "${ref}" no es de la tabla de origen "${origen}".`);
  }
  return f.columna;
}

function cubeField(ref: string, where: string): string {
  const f = parseFieldRef(ref);
  if (!f.tabla) throw new Error(`${where}: en el modelo el campo va como "Tabla[Columna]" (recibido "${ref}").`);
  return `${mdx(f.tabla)}.${mdx(f.columna)}`;
}

const PERIOD_INDEX: Record<Periodo, number> = { Dias: 3, Meses: 4, Trimestres: 5, Anios: 6 };

function pivotStep(pt: TablaDinamicaSpec, sharedWith: string | undefined): { step: Step; info: PivotInfo } {
  const where = `Tabla dinámica "${pt.nombre}"`;
  const modelo = pt.origen.trim().toLowerCase() === "modelo";
  const dest = cellRef(pt.destino, `${where} (destino)`);
  const field = (ref: string) => pivotFieldName(modelo, ref, pt.origen, where);
  const body: string[] = [`BorrarTD ${vbaString(pt.nombre)}`];

  if (modelo) {
    body.push(
      "Set pt = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, " +
        'SourceData:=ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _',
      `    CreatePivotTable(TableDestination:=${rangeExpr(dest)}, TableName:=${vbaString(pt.nombre)}, DefaultVersion:=6)`,
    );
  } else {
    const { sheet, address } = parseRef(pt.origen);
    const source = sheet ? rangeExpr({ sheet, address }) : vbaString(pt.origen);
    body.push(
      `Set pt = CrearTDClasica(${source}, ${vbaString(pt.nombre)}, ${rangeExpr(dest)}, ${vbaString(sharedWith ?? "")})`,
    );
  }

  const axis = (refs: string[], orientation: string) => {
    for (const ref of refs) {
      body.push(
        modelo
          ? `pt.CubeFields(${vbaString(cubeField(ref, where))}).Orientation = ${orientation}`
          : `pt.PivotFields(${vbaString(field(ref))}).Orientation = ${orientation}`,
      );
    }
  };
  axis(pt.filas, "xlRowField");
  axis(pt.columnas, "xlColumnField");
  axis(pt.filtros, "xlPageField");

  for (const g of pt.agrupar_fechas ?? []) {
    if (modelo) {
      throw new Error(
        `${where}: en el modelo de datos no se pueden agrupar fechas; agrega antes a la tabla columnas calculadas ` +
          `(por ejemplo Año =YEAR([@Fecha]) y Trimestre ="Trim "&ROUNDUP(MONTH([@Fecha])/3,0)) y úsalas como filas.`,
      );
    }
    const onAxis = [...pt.filas, ...pt.columnas].some((r) => field(r) === field(g.campo));
    if (!onAxis) throw new Error(`${where}: para agrupar "${g.campo}" debe estar en filas o columnas.`);
    const periods = Array<boolean>(7).fill(false);
    for (const p of g.periodos) periods[PERIOD_INDEX[p]] = true;
    body.push(
      `pt.PivotFields(${vbaString(field(g.campo))}).DataRange.Cells(1).Group Start:=True, End:=True, ` +
        `Periods:=Array(${periods.map((b) => (b ? "True" : "False")).join(", ")})`,
    );
  }

  const captions = uniqueCaptions(
    pt.valores.map((v) => v.nombre ?? defaultCaption(v.resumen, parseFieldRef(v.campo).columna, modelo)),
  );
  pt.valores.forEach((v, i) => {
    const caption = captions[i];
    if (modelo) {
      const fn = v.resumen === "DistinctCount" ? "xlDistinctCount" : `xl${v.resumen}`;
      body.push(
        `Set df = pt.AddDataField(pt.CubeFields.GetMeasure(${vbaString(cubeField(v.campo, where))}, ${fn}, ` +
          `${vbaString(caption)}), ${vbaString(caption)})`,
      );
    } else {
      if (v.resumen === "DistinctCount") {
        throw new Error(`${where}: el recuento distinto solo existe en tablas dinámicas del modelo de datos (origen "modelo").`);
      }
      body.push(
        `Set df = pt.AddDataField(pt.PivotFields(${vbaString(field(v.campo))}), ${vbaString(caption)}, xl${v.resumen})`,
      );
    }
    if (v.mostrar_como && v.mostrar_como !== "sin_calculo") {
      const calc = MOSTRAR_COMO[v.mostrar_como];
      body.push(`df.Calculation = ${calc.vba}`);
      if (calc.base) {
        const base = v.campo_base ?? pt.filas[0] ?? pt.columnas[0];
        if (!base) throw new Error(`${where}: "${v.mostrar_como}" necesita campo_base o un campo en filas.`);
        body.push(`df.BaseField = ${vbaString(field(base))}`);
      }
    }
    if (v.formato_numero) body.push(`df.NumberFormat = ${vbaString(v.formato_numero)}`);
  });

  if (pt.diseno) {
    const layout = { compacto: "xlCompactRow", esquema: "xlOutlineRow", tabular: "xlTabularRow" }[pt.diseno];
    body.push(`pt.RowAxisLayout ${layout}`);
  }
  // Excel crea las tablas del modelo sin subtotales; en la interfaz aparecen arriba.
  const subtotales = pt.subtotales ?? (pt.filas.length > 1 ? "arriba" : undefined);
  if (subtotales) {
    const loc = { arriba: "xlAtTop", abajo: "xlAtBottom", ninguno: "0" }[subtotales];
    body.push(`Subtotales pt, ${loc}`);
  }
  if (pt.total_general_abajo !== undefined) body.push(`pt.ColumnGrand = ${pt.total_general_abajo ? "True" : "False"}`);
  if (pt.total_general_derecha !== undefined) body.push(`pt.RowGrand = ${pt.total_general_derecha ? "True" : "False"}`);
  if (pt.encabezado_filas) body.push(`pt.CompactLayoutRowHeader = ${vbaString(pt.encabezado_filas)}`);
  if (pt.estilo) body.push(`pt.TableStyle2 = ${vbaString(pt.estilo)}`);

  // Los campos de valores se referencian por posición: en el modelo, GetMeasure reutiliza la medida implícita
  // que ya exista (p. ej. "Suma de X") y el título pedido queda solo como rótulo, así que el nombre no sirve.
  const valueIndex = (name: string | undefined, what: string) => {
    if (!captions.length) throw new Error(`${where}: ${what} necesita al menos un campo de valores.`);
    if (!name) return 1;
    const i = captions.findIndex((c) => c.toLowerCase() === name.toLowerCase());
    if (i < 0) throw new Error(`${where}: ${what} usa el valor "${name}", pero los valores son: ${captions.join(", ")}.`);
    return i + 1;
  };

  if (pt.top) {
    const onRows = [...pt.filas, ...pt.columnas].some((r) => field(r) === field(pt.top!.campo));
    if (!onRows) throw new Error(`${where}: el filtro top necesita que "${pt.top.campo}" esté en filas o columnas.`);
    const k = valueIndex(pt.top.valor, "el filtro top");
    body.push(
      `pt.PivotFields(${vbaString(field(pt.top.campo))}).PivotFilters.Add2 Type:=${FILTROS_TOP[pt.top.tipo]}, ` +
        `DataField:=pt.DataFields(${k}), Value1:=${pt.top.cantidad}`,
    );
  }
  for (const o of pt.ordenar ?? []) {
    const order = o.descendente ? "xlDescending" : "xlAscending";
    const key = o.por_valor ? `pt.DataFields(${valueIndex(o.por_valor, "ordenar")}).Name` : vbaString(field(o.campo));
    body.push(`pt.PivotFields(${vbaString(field(o.campo))}).AutoSort ${order}, ${key}`);
  }

  const origen = modelo ? "modelo de datos" : pt.origen;
  return {
    step: {
      titulo: `Tabla dinámica «${pt.nombre}» en ${dest.sheet}!${dest.address} (${origen})`,
      dims: ["Dim pt As PivotTable", "Dim df As PivotField"],
      cuerpo: body,
    },
    info: { modelo, captions, field },
  };
}

function slicerStep(s: SegmentadorSpec, name: string, pivots: Map<string, PivotInfo>): Step {
  const where = `Segmentador de "${s.campo}"`;
  if (!s.tablas_dinamicas.length) throw new Error(`${where}: indica al menos una tabla dinámica.`);
  const r = cellRef(s.rango, `${where} (rango)`);
  const first = pivots.get(s.tablas_dinamicas[0].toLowerCase());
  const modelo = first ? first.modelo : parseFieldRef(s.campo).tabla !== undefined;
  const f = parseFieldRef(s.campo);
  const source = modelo ? cubeField(s.campo, where) : f.columna;
  const body = [`BorrarSegmentacion ${vbaString(name)}`, `Set r = ${rangeExpr(r)}`];
  const caption = s.titulo ?? f.columna;
  const geometry = "r.Top, r.Left, r.Width, r.Height";
  if (s.escala_tiempo) {
    body.push(
      `Set sc = ActiveWorkbook.SlicerCaches.Add2(TD(${vbaString(s.tablas_dinamicas[0])}), ${vbaString(source)}, , xlTimeline)`,
      `Set sl = sc.Slicers.Add(r.Worksheet, , ${vbaString(name)}, ${vbaString(caption)}, ${geometry})`,
    );
  } else {
    body.push(`Set sc = ActiveWorkbook.SlicerCaches.Add2(TD(${vbaString(s.tablas_dinamicas[0])}), ${vbaString(source)})`);
    const level = modelo ? `${vbaString(`${source}.${mdx(f.columna)}`)}` : "";
    body.push(`Set sl = sc.Slicers.Add(r.Worksheet, ${level}, ${vbaString(name)}, ${vbaString(caption)}, ${geometry})`);
    if (s.columnas) body.push(`sl.NumberOfColumns = ${s.columnas}`);
  }
  if (s.estilo) body.push(`sl.Style = ${vbaString(s.estilo)}`);
  for (const other of s.tablas_dinamicas.slice(1)) body.push(`sc.PivotTables.AddPivotTable TD(${vbaString(other)})`);
  const extra = s.tablas_dinamicas.length > 1 ? `, conectado a ${s.tablas_dinamicas.join(", ")}` : "";
  return {
    titulo: `${s.escala_tiempo ? "Escala de tiempo" : "Segmentador"} «${caption}» en ${r.sheet}!${r.address}${extra}`,
    dims: ["Dim sc As SlicerCache", "Dim sl As Slicer", "Dim r As Range"],
    cuerpo: body,
  };
}

function chartStep(g: GraficoSpec, name: string): Step {
  const r = cellRef(g.rango, `Gráfico de "${g.tabla_dinamica}" (rango)`);
  const body = [
    `Set r = ${rangeExpr(r)}`,
    `BorrarForma r.Worksheet, ${vbaString(name)}`,
    `Set co = r.Worksheet.Shapes.AddChart2(-1, xl${g.tipo}, r.Left, r.Top, r.Width, r.Height)`,
    `co.Name = ${vbaString(name)}`,
    `co.Chart.SetSourceData Source:=TD(${vbaString(g.tabla_dinamica)}).TableRange1`,
    `co.Chart.HasLegend = ${g.leyenda ? "True" : "False"}`,
  ];
  if (g.etiquetas_datos) {
    body.push("For i = 1 To co.Chart.SeriesCollection.Count", "    co.Chart.SeriesCollection(i).HasDataLabels = True", "Next i");
  }
  if (g.titulo) body.push("co.Chart.HasTitle = True", `co.Chart.ChartTitle.Text = ${vbaString(g.titulo)}`);
  return {
    titulo: `Gráfico dinámico «${name}» de «${g.tabla_dinamica}» en ${r.sheet}!${r.address}`,
    dims: ["Dim co As Shape", "Dim r As Range", "Dim i As Long"],
    cuerpo: body,
  };
}

function linkStep(v: VinculoSpec): Step {
  const formula = absoluteFormula(v.celda);
  return {
    titulo: `Vincular la forma «${v.forma}» (${v.hoja}) a ${formula.slice(1)}`,
    dims: [],
    cuerpo: [`BuscarForma(Hoja(${vbaString(v.hoja)}), ${vbaString(v.forma)}).DrawingObject.Formula = ${vbaString(formula)}`],
  };
}

// ---------- módulo ----------

const HELPERS = `
Private Sub Anotar(ByVal ok As Boolean, ByVal paso As String, Optional ByVal detalle As String = "")
    If ok Then
        correctos = correctos + 1
        Debug.Print "OK    " & paso
    Else
        fallos = fallos + 1
        resumen = resumen & vbLf & "- " & paso & ": " & detalle
        Debug.Print "ERROR " & paso & ": " & detalle
    End If
End Sub

' Devuelve la hoja; si no existe, la crea al final.
Private Function Hoja(ByVal nombre As String) As Worksheet
    On Error Resume Next
    Set Hoja = ActiveWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If Hoja Is Nothing Then
        Set Hoja = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
        Hoja.Name = nombre
    End If
End Function

Private Function BuscarTabla(ByVal nombre As String) As ListObject
    Dim h As Worksheet, lo As ListObject
    For Each h In ActiveWorkbook.Worksheets
        For Each lo In h.ListObjects
            If StrComp(lo.Name, nombre, vbTextCompare) = 0 Then
                Set BuscarTabla = lo
                Exit Function
            End If
        Next lo
    Next h
End Function

Private Function ExisteTablaModelo(ByVal nombre As String) As Boolean
    Dim t As ModelTable
    For Each t In ActiveWorkbook.Model.ModelTables
        If StrComp(t.Name, nombre, vbTextCompare) = 0 Then ExisteTablaModelo = True
    Next t
End Function

' Agrega una tabla de Excel al modelo de datos (como Power Pivot > Agregar al modelo de datos).
' El nombre de la tabla en el modelo es el de la tabla de Excel.
Private Sub AgregarAlModelo(ByVal tabla As String)
    Dim h As Worksheet, original As String, temporal As String, numero As Long, texto As String
    If ExisteTablaModelo(tabla) Then Exit Sub
    If BuscarTabla(tabla) Is Nothing Then Err.Raise vbObjectError + 1, , "No existe la tabla de Excel " & tabla & "."
    ' Si una hoja se llama igual que la tabla, Excel no encuentra la tabla: se renombra la hoja un momento.
    For Each h In ActiveWorkbook.Worksheets
        If StrComp(h.Name, tabla, vbTextCompare) = 0 Then
            original = h.Name
            temporal = Left(original, 24) & "_tmp"
            h.Name = temporal
        End If
    Next h
    On Error Resume Next
    ActiveWorkbook.Connections.Add2 "WorksheetConnection_" & tabla, "", "WORKSHEET;" & ActiveWorkbook.FullName, _
        ActiveWorkbook.Name & "!" & tabla, 7, True, False
    numero = Err.Number
    texto = Err.Description
    On Error GoTo 0
    If original <> "" Then ActiveWorkbook.Worksheets(temporal).Name = original
    If numero <> 0 Then Err.Raise numero, , texto
End Sub

' Relación del modelo: la clave foránea (lado "varios") apunta a la clave primaria (lado "uno", valores únicos).
Private Sub CrearRelacion(ByVal tablaF As String, ByVal columnaF As String, ByVal tablaP As String, ByVal columnaP As String)
    Dim rel As ModelRelationship
    If Not ExisteTablaModelo(tablaF) Then Err.Raise vbObjectError + 2, , "La tabla " & tablaF & " no está en el modelo de datos."
    If Not ExisteTablaModelo(tablaP) Then Err.Raise vbObjectError + 2, , "La tabla " & tablaP & " no está en el modelo de datos."
    For Each rel In ActiveWorkbook.Model.ModelRelationships
        If StrComp(rel.ForeignKeyTable.Name, tablaF, vbTextCompare) = 0 And StrComp(rel.ForeignKeyColumn.Name, columnaF, vbTextCompare) = 0 And _
           StrComp(rel.PrimaryKeyTable.Name, tablaP, vbTextCompare) = 0 And StrComp(rel.PrimaryKeyColumn.Name, columnaP, vbTextCompare) = 0 Then Exit Sub
    Next rel
    With ActiveWorkbook.Model
        .ModelRelationships.Add .ModelTables(tablaF).ModelTableColumns(columnaF), .ModelTables(tablaP).ModelTableColumns(columnaP)
    End With
End Sub

Private Function BuscarTD(ByVal nombre As String) As PivotTable
    Dim h As Worksheet, p As PivotTable
    For Each h In ActiveWorkbook.Worksheets
        For Each p In h.PivotTables
            If StrComp(p.Name, nombre, vbTextCompare) = 0 Then
                Set BuscarTD = p
                Exit Function
            End If
        Next p
    Next h
End Function

Private Function TD(ByVal nombre As String) As PivotTable
    Set TD = BuscarTD(nombre)
    If TD Is Nothing Then Err.Raise vbObjectError + 3, , "No existe la tabla dinámica " & nombre & "."
End Function

' Si la macro se vuelve a ejecutar, reemplaza la tabla dinámica con ese nombre.
Private Sub BorrarTD(ByVal nombre As String)
    Dim p As PivotTable
    Set p = BuscarTD(nombre)
    If Not p Is Nothing Then p.TableRange2.Clear
End Sub

' Tabla dinámica desde una tabla o rango. Las que comparten origen comparten la caché, como en la interfaz,
' para que un mismo segmentador pueda filtrarlas a todas.
Private Function CrearTDClasica(ByVal origen As Variant, ByVal nombre As String, ByVal destino As Range, ByVal compartirCon As String) As PivotTable
    Dim pc As PivotCache, base As PivotTable
    If compartirCon <> "" Then Set base = BuscarTD(compartirCon)
    If base Is Nothing Then
        Set pc = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=origen, Version:=6)
    Else
        Set pc = base.PivotCache
    End If
    Set CrearTDClasica = pc.CreatePivotTable(TableDestination:=destino, TableName:=nombre, DefaultVersion:=6)
End Function

' Subtotales de los campos de fila: arriba (xlAtTop), abajo (xlAtBottom) o ninguno (0).
Private Sub Subtotales(ByVal pt As PivotTable, ByVal ubicacion As Long)
    Dim i As Long, valores As String
    valores = pt.DataPivotField.Name
    For i = 1 To pt.RowFields.Count
        If pt.RowFields(i).Name <> valores Then pt.RowFields(i).Subtotals(1) = (ubicacion <> 0)
    Next i
    If ubicacion <> 0 Then pt.SubtotalLocation ubicacion
End Sub

Private Sub BorrarSegmentacion(ByVal nombre As String)
    Dim sc As SlicerCache, sl As Slicer
    For Each sc In ActiveWorkbook.SlicerCaches
        For Each sl In sc.Slicers
            If StrComp(sl.Name, nombre, vbTextCompare) = 0 Then
                sc.Delete
                Exit Sub
            End If
        Next sl
    Next sc
End Sub

' Busca una forma por nombre, también dentro de grupos.
Private Function BuscarForma(ByVal h As Worksheet, ByVal nombre As String) As Shape
    Dim s As Shape, g As Shape
    For Each s In h.Shapes
        If StrComp(s.Name, nombre, vbTextCompare) = 0 Then
            Set BuscarForma = s
            Exit Function
        End If
        If s.Type = msoGroup Then
            For Each g In s.GroupItems
                If StrComp(g.Name, nombre, vbTextCompare) = 0 Then
                    Set BuscarForma = g
                    Exit Function
                End If
            Next g
        End If
    Next s
    Err.Raise vbObjectError + 4, , "No existe la forma " & nombre & " en la hoja " & h.Name & "."
End Function

Private Sub BorrarForma(ByVal h As Worksheet, ByVal nombre As String)
    Dim s As Shape
    For Each s In h.Shapes
        If StrComp(s.Name, nombre, vbTextCompare) = 0 Then
            s.Delete
            Exit Sub
        End If
    Next s
End Sub
`;

function validateSpec(spec: MacroSpec): void {
  if (!/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(spec.nombre_macro)) {
    throw new Error(`nombre_macro "${spec.nombre_macro}" no es válido: solo letras sin tilde, números y _, empezando con letra.`);
  }
  if (RESERVED.has(spec.nombre_macro.toLowerCase()) || /^paso\d+$/i.test(spec.nombre_macro)) {
    throw new Error(`nombre_macro "${spec.nombre_macro}" coincide con un procedimiento interno; usa otro nombre.`);
  }
  const total =
    spec.tablas_modelo.length +
    spec.relaciones.length +
    spec.tablas_dinamicas.length +
    spec.segmentadores.length +
    spec.graficos.length +
    spec.vinculos_formas.length +
    (spec.codigo_extra?.trim() ? 1 : 0);
  if (!total) throw new Error("La macro no tiene ningún paso.");
  const names = new Set<string>();
  for (const pt of spec.tablas_dinamicas) {
    const key = pt.nombre.toLowerCase();
    if (names.has(key)) throw new Error(`Hay dos tablas dinámicas llamadas "${pt.nombre}".`);
    names.add(key);
  }
  if (spec.codigo_extra && /^\s*(public\s+|private\s+)?(sub|function)\s/im.test(spec.codigo_extra)) {
    throw new Error("codigo_extra debe tener solo instrucciones (sin Sub ni Function): se ejecutan dentro de su propio paso.");
  }
}

/** Genera el módulo VBA completo. Lanza un error con un mensaje útil si la especificación no es coherente. */
export function buildMacro(spec: MacroSpec): MacroResult {
  validateSpec(spec);
  const steps: Step[] = [];

  for (const t of spec.tablas_modelo) {
    steps.push({ titulo: `Agregar la tabla «${t}» al modelo de datos`, dims: [], cuerpo: [`AgregarAlModelo ${vbaString(t)}`] });
  }
  for (const r of spec.relaciones) {
    const f = parseFieldRef(r.foranea);
    const p = parseFieldRef(r.primaria);
    if (!f.tabla || !p.tabla) {
      throw new Error(`Relación ${r.foranea} -> ${r.primaria}: usa "Tabla[Columna]" en ambos lados.`);
    }
    steps.push({
      titulo: `Relación ${p.tabla}[${p.columna}] (uno) -> ${f.tabla}[${f.columna}] (varios)`,
      dims: [],
      cuerpo: [`CrearRelacion ${[f.tabla, f.columna, p.tabla, p.columna].map(vbaString).join(", ")}`],
    });
  }

  const pivots = new Map<string, PivotInfo>();
  const firstBySource = new Map<string, string>();
  const valores: Record<string, string[]> = {};
  for (const pt of spec.tablas_dinamicas) {
    const modelo = pt.origen.trim().toLowerCase() === "modelo";
    const sourceKey = pt.origen.trim().toLowerCase();
    const shared = modelo ? undefined : firstBySource.get(sourceKey);
    const { step, info } = pivotStep(pt, shared);
    if (!modelo && !shared) firstBySource.set(sourceKey, pt.nombre);
    pivots.set(pt.nombre.toLowerCase(), info);
    valores[pt.nombre] = info.captions;
    steps.push(step);
  }

  const used = new Set<string>();
  const unique = (base: string) => {
    let name = base;
    for (let n = 2; used.has(name.toLowerCase()); n++) name = `${base}_${n}`;
    used.add(name.toLowerCase());
    return name;
  };
  for (const s of spec.segmentadores) {
    const name = unique(s.nombre ?? objectName(s.escala_tiempo ? "Cron" : "Seg", parseFieldRef(s.campo).columna));
    steps.push(slicerStep(s, name, pivots));
  }
  for (const g of spec.graficos) {
    steps.push(chartStep(g, unique(g.nombre ?? objectName("Graf", g.tabla_dinamica))));
  }
  for (const v of spec.vinculos_formas) steps.push(linkStep(v));
  if (spec.codigo_extra?.trim()) {
    steps.push({ titulo: "Instrucciones adicionales", dims: [], cuerpo: spec.codigo_extra.trim().split(/\r?\n/) });
  }

  const usaModelo =
    spec.tablas_modelo.length > 0 ||
    spec.relaciones.length > 0 ||
    [...pivots.values()].some((p) => p.modelo);
  const title = vbaString(spec.nombre_macro);
  const pad = (n: number) => String(n).padStart(2, "0");
  const lines: string[] = [
    "Option Explicit",
    "",
    comment(`Macro generada por Agente Excel. Ejecútela una vez con este libro abierto (Alt+F8 > ${spec.nombre_macro} > Ejecutar).`),
    comment("Si la vuelve a ejecutar, reemplaza las tablas dinámicas, segmentadores y gráficos con los mismos nombres."),
    ...(usaModelo ? [comment("Usa el modelo de datos (Power Pivot): requiere Excel 2013 o posterior para Windows.")] : []),
    comment("Pasos:"),
    ...steps.map((s, i) => comment(`  ${i + 1}. ${s.titulo}`)),
    "",
    "Private resumen As String",
    "Private correctos As Long",
    "Private fallos As Long",
    "",
    `Public Sub ${spec.nombre_macro}()`,
  ];
  if (usaModelo) {
    lines.push(
      "#If Mac Then",
      `    MsgBox "El modelo de datos (Power Pivot) solo existe en Excel para Windows.", vbExclamation, ${title}`,
      "    Exit Sub",
      "#End If",
    );
  }
  lines.push('    resumen = ""', "    correctos = 0", "    fallos = 0", "    Application.ScreenUpdating = False");
  steps.forEach((_, i) => lines.push(`    Paso${pad(i + 1)}`));
  lines.push(
    "    Application.ScreenUpdating = True",
    "    If fallos = 0 Then",
    `        MsgBox "Listo: " & correctos & " pasos completados.", vbInformation, ${title}`,
    "    Else",
    `        MsgBox correctos & " pasos completados y " & fallos & " con error:" & resumen & vbLf & vbLf & ` +
      `${vbaString("Copie este mensaje (Ctrl+C) y péguelo en el chat del agente.")}, vbExclamation, ${title}`,
    "    End If",
    "End Sub",
  );
  steps.forEach((s, i) => {
    lines.push(
      "",
      comment(`Paso ${i + 1}: ${s.titulo}`),
      `Private Sub Paso${pad(i + 1)}()`,
      ...s.dims.map((d) => `    ${d}`),
      "    On Error GoTo Fallo",
      ...s.cuerpo.map((l) => `    ${l}`),
      `    Anotar True, ${vbaString(s.titulo)}`,
      "    Exit Sub",
      "Fallo:",
      `    Anotar False, ${vbaString(s.titulo)}, Err.Description`,
      "End Sub",
    );
  });
  lines.push(HELPERS.trimEnd());
  return { codigo: lines.join("\n") + "\n", pasos: steps.map((s) => s.titulo), valores, usaModelo };
}
