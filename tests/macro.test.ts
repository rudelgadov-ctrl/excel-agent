import { describe, expect, it } from "vitest";
import { cleanCell, cleanText, excelSerial, excelWouldConvert, parseDate, parseNumber, properCase } from "../src/excel/clean";
import { buildMacro, type MacroSpec, objectName, vbaString } from "../src/excel/macro";
import { defaultCaption, parseFieldRef, uniqueCaptions } from "../src/excel/pivot";

const base: MacroSpec = {
  nombre_macro: "CrearModelo",
  tablas_modelo: [],
  relaciones: [],
  tablas_dinamicas: [],
  segmentadores: [],
  graficos: [],
  vinculos_formas: [],
};

const full: MacroSpec = {
  ...base,
  tablas_modelo: ["Ventas", "Vendedores", "Regiones"],
  relaciones: [
    { foranea: "Ventas[IdVendedor]", primaria: "Vendedores[IdVendedor]" },
    { foranea: "Vendedores[CodRegion]", primaria: "Regiones[Codigo]" },
  ],
  tablas_dinamicas: [
    {
      nombre: "TD_Trimestre",
      destino: "'Análisis'!A2",
      origen: "modelo",
      filas: ["Ventas[Trimestre]"],
      columnas: [],
      filtros: [],
      valores: [
        { campo: "Ventas[Importe]", resumen: "Sum", formato_numero: '"₡"#,##0.00' },
        { campo: "Ventas[Importe]", resumen: "Sum", mostrar_como: "porcentaje_total_general", formato_numero: "0.00%" },
      ],
    },
    {
      nombre: "TD_Top",
      destino: "'Panel'!A18",
      origen: "modelo",
      filas: ["Vendedores[Nombre]"],
      columnas: [],
      filtros: [],
      valores: [{ campo: "Ventas[Importe]", resumen: "Sum", nombre: "Total vendido" }],
      top: { campo: "Vendedores[Nombre]", cantidad: 10, tipo: "superiores" },
      encabezado_filas: "Nombre",
    },
    {
      nombre: "TD_AnioTrim",
      destino: "'Análisis'!K2",
      origen: "modelo",
      filas: ["Ventas[Año]", "Ventas[Trimestre]"],
      columnas: [],
      filtros: [],
      valores: [{ campo: "Ventas[IdVenta]", resumen: "Count" }],
      ordenar: [{ campo: "Ventas[Año]", descendente: true, por_valor: "Recuento de IdVenta" }],
    },
  ],
  segmentadores: [
    { campo: "Regiones[Region]", tablas_dinamicas: ["TD_Top", "TD_Trimestre"], rango: "'Panel'!A6:B13" },
    { campo: "Ventas[Trimestre]", tablas_dinamicas: ["TD_Top"], rango: "'Panel'!E6:G13", columnas: 2, titulo: "Fecha (Trimestre)" },
    { campo: "Ventas[Fecha]", tablas_dinamicas: ["TD_Top"], rango: "'Panel'!H6:M10", escala_tiempo: true },
  ],
  graficos: [{ tabla_dinamica: "TD_Trimestre", rango: "'Panel'!D17:J30", tipo: "ColumnClustered", leyenda: false, etiquetas_datos: true }],
  vinculos_formas: [{ hoja: "Panel", forma: "KPI_Valor", celda: "'Análisis'!E5" }],
};

describe("vbaString", () => {
  it("duplica comillas, conserva Latin-1 y usa ChrW para el resto", () => {
    expect(vbaString('Dijo "hola"')).toBe('"Dijo ""hola"""');
    expect(vbaString("Año ñandú")).toBe('"Año ñandú"');
    expect(vbaString('"₡"#,##0')).toBe('"""" & ChrW(8353) & """#,##0"');
    expect(vbaString("a\nb")).toBe('"a" & vbLf & "b"');
    expect(vbaString("")).toBe('""');
  });

  it("nombres de objeto sin tildes ni espacios", () => {
    expect(objectName("Seg", "Región de venta")).toBe("Seg_Region_de_venta");
  });
});

describe("definiciones de tablas dinámicas", () => {
  it("lee referencias Tabla[Columna]", () => {
    expect(parseFieldRef("Ventas[Importe]")).toEqual({ tabla: "Ventas", columna: "Importe" });
    expect(parseFieldRef("'Mis ventas'[Fecha de pago]")).toEqual({ tabla: "Mis ventas", columna: "Fecha de pago" });
    expect(parseFieldRef("Importe")).toEqual({ columna: "Importe" });
  });

  it("títulos por defecto como Excel en español y únicos", () => {
    expect(defaultCaption("Count", "Id", true)).toBe("Recuento de Id");
    expect(defaultCaption("Count", "Id", false)).toBe("Cuenta de Id");
    expect(uniqueCaptions(["Suma de X", "Suma de X", "Suma de Y"])).toEqual(["Suma de X", "Suma de X2", "Suma de Y"]);
  });
});

describe("buildMacro", () => {
  const { codigo, pasos, valores, usaModelo } = buildMacro(full);

  it("arma un módulo completo con un paso por acción", () => {
    expect(codigo.startsWith("Option Explicit\n")).toBe(true);
    expect(codigo).toContain("Public Sub CrearModelo()");
    expect(pasos).toHaveLength(3 + 2 + 3 + 3 + 1 + 1);
    expect(codigo.match(/^Private Sub Paso\d\d\(\)$/gm)).toHaveLength(pasos.length);
    for (let i = 1; i <= pasos.length; i++) expect(codigo).toContain(`    Paso${String(i).padStart(2, "0")}\n`);
    expect(usaModelo).toBe(true);
    expect(codigo).toContain("#If Mac Then");
  });

  it("es VBA bien formado: bloques balanceados, líneas cortas y solo caracteres Latin-1", () => {
    const lines = codigo.split("\n");
    const count = (re: RegExp) => lines.filter((l) => re.test(l.trim())).length;
    expect(count(/^(Public |Private )?Sub /)).toBe(count(/^End Sub$/));
    expect(count(/^(Public |Private )?Function /)).toBe(count(/^End Function$/));
    expect(count(/^For (Each )?/)).toBe(count(/^Next\b/));
    expect(count(/^If .* Then$/)).toBe(count(/^End If$/));
    expect(count(/^With /)).toBe(count(/^End With$/));
    for (const l of lines) expect(l.length).toBeLessThan(1000);
    expect([...codigo].every((c) => c.codePointAt(0)! <= 0xff)).toBe(true);
  });

  it("agrega tablas al modelo y crea las relaciones", () => {
    expect(codigo).toContain('AgregarAlModelo "Ventas"');
    expect(codigo).toContain('CrearRelacion "Ventas", "IdVendedor", "Vendedores", "IdVendedor"');
    expect(codigo).toContain('CrearRelacion "Vendedores", "CodRegion", "Regiones", "Codigo"');
  });

  it("crea tablas dinámicas desde el modelo con medidas, % del total, formato, top 10, orden y subtotales", () => {
    expect(codigo).toContain('ActiveWorkbook.Connections("ThisWorkbookDataModel")');
    expect(codigo).toContain('CreatePivotTable(TableDestination:=Hoja("Análisis").Range("A2"), TableName:="TD_Trimestre"');
    expect(codigo).toContain('pt.CubeFields("[Ventas].[Trimestre]").Orientation = xlRowField');
    expect(codigo).toContain(
      'Set df = pt.AddDataField(pt.CubeFields.GetMeasure("[Ventas].[Importe]", xlSum, "Suma de Importe2"), "Suma de Importe2")',
    );
    expect(codigo).toContain("df.Calculation = xlPercentOfTotal");
    expect(codigo).toContain('df.NumberFormat = """" & ChrW(8353) & """#,##0.00"');
    expect(codigo).toContain(
      'pt.PivotFields("[Vendedores].[Nombre].[Nombre]").PivotFilters.Add2 Type:=xlTopCount, DataField:=pt.DataFields(1), Value1:=10',
    );
    expect(codigo).toContain('pt.CompactLayoutRowHeader = "Nombre"');
    expect(codigo).toContain('pt.PivotFields("[Ventas].[Año].[Año]").AutoSort xlDescending, pt.DataFields(1).Name');
    expect(codigo).toContain("Subtotales pt, xlAtTop");
    expect(valores).toEqual({
      TD_Trimestre: ["Suma de Importe", "Suma de Importe2"],
      TD_Top: ["Total vendido"],
      TD_AnioTrim: ["Recuento de IdVenta"],
    });
  });

  it("crea segmentadores conectados, con columnas, y escalas de tiempo", () => {
    expect(codigo).toContain('Set sc = ActiveWorkbook.SlicerCaches.Add2(TD("TD_Top"), "[Regiones].[Region]")');
    expect(codigo).toContain(
      'Set sl = sc.Slicers.Add(r.Worksheet, "[Regiones].[Region].[Region]", "Seg_Region", "Region", r.Top, r.Left, r.Width, r.Height)',
    );
    expect(codigo).toContain('sc.PivotTables.AddPivotTable TD("TD_Trimestre")');
    expect(codigo).toContain("sl.NumberOfColumns = 2");
    expect(codigo).toContain('"Fecha (Trimestre)"');
    expect(codigo).toContain('ActiveWorkbook.SlicerCaches.Add2(TD("TD_Top"), "[Ventas].[Fecha]", , xlTimeline)');
  });

  it("crea gráficos dinámicos y vincula formas a celdas", () => {
    expect(codigo).toContain("r.Worksheet.Shapes.AddChart2(-1, xlColumnClustered, r.Left, r.Top, r.Width, r.Height)");
    expect(codigo).toContain('co.Chart.SetSourceData Source:=TD("TD_Trimestre").TableRange1');
    expect(codigo).toContain("co.Chart.HasLegend = False");
    expect(codigo).toContain("co.Chart.SeriesCollection(i).HasDataLabels = True");
    expect(codigo).toContain(`BuscarForma(Hoja("Panel"), "KPI_Valor").DrawingObject.Formula = "='Análisis'!$E$5"`);
  });

  it("tablas dinámicas clásicas: agrupa fechas y comparte la caché del mismo origen", () => {
    const r = buildMacro({
      ...base,
      tablas_dinamicas: [
        {
          nombre: "TD1",
          destino: "'R'!A1",
          origen: "Ventas",
          filas: ["Fecha"],
          columnas: [],
          filtros: [],
          valores: [{ campo: "Ventas[Importe]", resumen: "Sum" }],
          agrupar_fechas: [{ campo: "Fecha", periodos: ["Anios", "Trimestres"] }],
          subtotales: "arriba",
        },
        {
          nombre: "TD2",
          destino: "'R'!H1",
          origen: "Ventas",
          filas: ["Vendedor"],
          columnas: [],
          filtros: [],
          valores: [{ campo: "Importe", resumen: "Count" }],
          ordenar: [{ campo: "Vendedor", descendente: false }],
        },
      ],
      segmentadores: [{ campo: "Vendedor", tablas_dinamicas: ["TD1", "TD2"], rango: "'R'!N1:P8" }],
    });
    expect(r.usaModelo).toBe(false);
    expect(r.codigo).not.toContain("#If Mac");
    expect(r.codigo).toContain('Set pt = CrearTDClasica("Ventas", "TD1", Hoja("R").Range("A1"), "")');
    expect(r.codigo).toContain('Set pt = CrearTDClasica("Ventas", "TD2", Hoja("R").Range("H1"), "TD1")');
    expect(r.codigo).toContain(
      'pt.PivotFields("Fecha").DataRange.Cells(1).Group Start:=True, End:=True, Periods:=Array(False, False, False, False, False, True, True)',
    );
    expect(r.codigo).toContain('Set df = pt.AddDataField(pt.PivotFields("Importe"), "Suma de Importe", xlSum)');
    expect(r.codigo).toContain('Set df = pt.AddDataField(pt.PivotFields("Importe"), "Cuenta de Importe", xlCount)');
    expect(r.codigo).toContain('pt.PivotFields("Vendedor").AutoSort xlAscending, "Vendedor"');
    expect(r.codigo).toContain('Set sl = sc.Slicers.Add(r.Worksheet, , "Seg_Vendedor", "Vendedor"');
  });

  it("rechaza especificaciones incoherentes con mensajes útiles", () => {
    const pivot = (extra: object) => ({
      ...base,
      tablas_dinamicas: [
        {
          nombre: "TD",
          destino: "'R'!A1",
          origen: "modelo",
          filas: ["Ventas[Fecha]"],
          columnas: [],
          filtros: [],
          valores: [{ campo: "Ventas[Importe]", resumen: "Sum" as const }],
          ...extra,
        },
      ],
    });
    expect(() => buildMacro(base)).toThrow(/ningún paso/);
    expect(() => buildMacro({ ...full, nombre_macro: "Mi macro" })).toThrow(/nombre_macro/);
    expect(() => buildMacro({ ...full, nombre_macro: "Hoja" })).toThrow(/interno/);
    expect(() => buildMacro(pivot({ agrupar_fechas: [{ campo: "Ventas[Fecha]", periodos: ["Meses"] }] }))).toThrow(
      /columnas calculadas/,
    );
    expect(() => buildMacro(pivot({ filas: ["Fecha"] }))).toThrow(/Tabla\[Columna\]/);
    expect(() => buildMacro(pivot({ top: { campo: "Ventas[Otro]", cantidad: 5, tipo: "superiores" } }))).toThrow(/filas o columnas/);
    expect(() => buildMacro(pivot({ ordenar: [{ campo: "Ventas[Fecha]", descendente: true, por_valor: "Nada" }] }))).toThrow(
      /Suma de Importe/,
    );
    expect(() => buildMacro(pivot({ destino: "A1" }))).toThrow(/indica la hoja/);
    expect(() =>
      buildMacro({ ...base, tablas_dinamicas: [{ ...pivot({}).tablas_dinamicas[0], origen: "Ventas", valores: [{ campo: "Importe", resumen: "DistinctCount" }] }] }),
    ).toThrow(/recuento distinto/);
    expect(() => buildMacro({ ...base, codigo_extra: "Sub X()\nEnd Sub" })).toThrow(/sin Sub/);
    expect(() => buildMacro({ ...base, relaciones: [{ foranea: "IdVendedor", primaria: "Vendedores[IdVendedor]" }] })).toThrow(
      /Tabla\[Columna\]/,
    );
  });
});

describe("limpieza de datos", () => {
  const opts = {
    recortar_espacios: true,
    quitar_no_imprimibles: true,
    texto_a_numero: false,
    separador_decimal: "punto" as const,
    texto_a_fecha: false,
    orden_fecha: "dia_mes_anio" as const,
  };

  it("recorta espacios (también no separables) y cambia mayúsculas", () => {
    expect(cleanText("  Código  de   cliente ", opts)).toBe("Código de cliente");
    expect(cleanText("​hola\u0007", opts)).toBe("hola");
    expect(properCase("maría DE los ángeles o'neil")).toBe("María De Los Ángeles O'Neil");
    expect(cleanText("san josé", { ...opts, mayusculas: "mayusculas" })).toBe("SAN JOSÉ");
  });

  it("convierte números escritos como texto según el separador decimal", () => {
    expect(parseNumber("1,234.50", "punto")).toBe(1234.5);
    expect(parseNumber("1.234,50", "coma")).toBe(1234.5);
    expect(parseNumber("2407984,02", "coma")).toBe(2407984.02);
    expect(parseNumber("₡ 2 500", "punto")).toBe(2500);
    expect(parseNumber("15%", "punto")).toBe(0.15);
    expect(parseNumber("ABC001", "punto")).toBeNull();
    expect(parseNumber("1,23", "punto")).toBeNull();
  });

  it("convierte fechas ISO o con barras en números de serie de Excel", () => {
    expect(excelSerial(2024, 3, 15)).toBe(45366);
    expect(excelSerial(2023, 2, 29)).toBeNull();
    expect(parseDate("2024-03-15", "dia_mes_anio")).toEqual({ serial: 45366, time: false });
    expect(parseDate("2024-03-15 00:00:00", "dia_mes_anio")).toEqual({ serial: 45366, time: false });
    expect(parseDate("15/03/2024", "dia_mes_anio")).toEqual({ serial: 45366, time: false });
    expect(parseDate("03/15/2024", "mes_dia_anio")).toEqual({ serial: 45366, time: false });
    expect(parseDate("2024-03-15 12:00", "dia_mes_anio")).toEqual({ serial: 45366.5, time: true });
    expect(parseDate("31/02/2024", "dia_mes_anio")).toBeNull();
  });

  it("limpia solo textos e informa qué tipo de valor quedó", () => {
    expect(cleanCell(" ABC001  ", opts)).toEqual({ kind: "text", value: "ABC001" });
    expect(cleanCell("Listo", opts)).toBeNull();
    expect(cleanCell(42, opts)).toBeNull();
    expect(cleanCell(" 1,500.00 ", { ...opts, texto_a_numero: true })).toEqual({ kind: "number", value: 1500 });
    expect(cleanCell("2020-03-15 00:00:00", { ...opts, texto_a_fecha: true })).toEqual({ kind: "date", value: 43905, time: false });
    expect(excelWouldConvert("00123")).toBe(true);
    expect(excelWouldConvert("ABC001")).toBe(false);
  });
});
