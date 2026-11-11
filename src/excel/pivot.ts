// Definiciones puras de tablas dinámicas compartidas por la herramienta de Office.js (tabla_dinamica)
// y por el generador de la macro (macro_modelo_datos).

export const RESUMENES = ["Sum", "Count", "Average", "Max", "Min", "DistinctCount"] as const;
export type Resumen = (typeof RESUMENES)[number];

/** "Mostrar valores como": nombre en español -> constante de VBA y valor de Excel.ShowAsCalculation. */
export const MOSTRAR_COMO = {
  sin_calculo: { vba: "xlNoAdditionalCalculation", officeJs: "None", base: false },
  porcentaje_total_general: { vba: "xlPercentOfTotal", officeJs: "PercentOfGrandTotal", base: false },
  porcentaje_total_columna: { vba: "xlPercentOfColumn", officeJs: "PercentOfColumnTotal", base: false },
  porcentaje_total_fila: { vba: "xlPercentOfRow", officeJs: "PercentOfRowTotal", base: false },
  porcentaje_total_fila_principal: { vba: "xlPercentOfParentRow", officeJs: "PercentOfParentRowTotal", base: false },
  porcentaje_total_columna_principal: {
    vba: "xlPercentOfParentColumn",
    officeJs: "PercentOfParentColumnTotal",
    base: false,
  },
  indice: { vba: "xlIndex", officeJs: "Index", base: false },
  total_acumulado: { vba: "xlRunningTotal", officeJs: "RunningTotal", base: true },
  porcentaje_total_acumulado: { vba: "xlPercentRunningTotal", officeJs: "PercentRunningTotal", base: true },
  clasificar_mayor_a_menor: { vba: "xlRankDecending", officeJs: "RankDecending", base: true },
  clasificar_menor_a_mayor: { vba: "xlRankAscending", officeJs: "RankAscending", base: true },
} as const;
export type MostrarComo = keyof typeof MOSTRAR_COMO;
export const MOSTRAR_COMO_KEYS = Object.keys(MOSTRAR_COMO) as [MostrarComo, ...MostrarComo[]];

const PREFIJOS: Record<Resumen, string> = {
  Sum: "Suma de",
  Count: "Cuenta de",
  Average: "Promedio de",
  Max: "Máx. de",
  Min: "Mín. de",
  DistinctCount: "Recuento distinto de",
};

/** Título por defecto de un campo de valores, como lo pone Excel en español. */
export function defaultCaption(resumen: Resumen, columna: string, modelo: boolean): string {
  const prefijo = resumen === "Count" && modelo ? "Recuento de" : PREFIJOS[resumen];
  return `${prefijo} ${columna}`;
}

/** Títulos únicos: si se repite uno, Excel agrega 2, 3… al final. */
export function uniqueCaptions(captions: string[]): string[] {
  const seen = new Map<string, number>();
  return captions.map((c) => {
    const key = c.toLowerCase();
    const n = (seen.get(key) ?? 0) + 1;
    seen.set(key, n);
    return n === 1 ? c : `${c}${n}`;
  });
}

export interface FieldRef {
  /** Tabla de Excel o del modelo, si la referencia la incluye ("Tabla[Columna]"). */
  tabla?: string;
  columna: string;
}

/** "Ventas[Monto]" -> { tabla: "Ventas", columna: "Monto" }; "Monto" -> { columna: "Monto" }. */
export function parseFieldRef(ref: string): FieldRef {
  const text = ref.trim();
  const m = /^(.+?)\[(.+)\]$/.exec(text);
  if (m) return { tabla: m[1].trim().replace(/^'(.*)'$/, "$1"), columna: m[2] };
  return { columna: text };
}

export function formatFieldRef(f: FieldRef): string {
  return f.tabla ? `${f.tabla}[${f.columna}]` : f.columna;
}
