// Helpers puros de texto para herramientas como "texto en columnas".

/**
 * Divide cada fila (primera columna) por el separador y devuelve una matriz
 * rectangular, rellenando con "" las filas que tienen menos partes.
 */
export function splitRows(values: unknown[][], separator: string, trim = true): string[][] {
  if (separator === "") throw new Error("El separador no puede estar vacío");
  const rows = values.map((row) => {
    const cell = row[0];
    if (cell === null || cell === undefined || cell === "") return [""];
    const parts = String(cell).split(separator);
    return trim ? parts.map((p) => p.trim()) : parts;
  });
  const width = Math.max(1, ...rows.map((r) => r.length));
  return rows.map((r) => [...r, ...Array<string>(width - r.length).fill("")]);
}

/** Recorta un texto largo para devolverlo a Claude sin pasarse de tamaño. */
export function truncate(text: string, max: number): string {
  if (text.length <= max) return text;
  return `${text.slice(0, max)}\n… [recortado: ${text.length - max} caracteres más]`;
}

const ERROR_VALUES = new Set([
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#GETTING_DATA",
  "#SPILL!",
  "#CALC!",
  "#FIELD!",
  "#BLOCKED!",
  "#CONNECT!",
  "#BUSY!",
  "#UNKNOWN!",
  "#¡NULO!",
  "#¡DIV/0!",
  "#¡VALOR!",
  "#¡REF!",
  "#¿NOMBRE?",
  "#¡NUM!",
  "#N/D",
  "#¡DESBORDAMIENTO!",
  "#¡CALC!",
]);

/** Devuelve las celdas con error (#NAME?, #VALUE!, …) para que el agente las corrija. */
export function findErrors(
  values: unknown[][],
  firstRow: number,
  firstCol: number,
  toAddress: (row: number, col: number) => string,
): { celda: string; error: string }[] {
  const out: { celda: string; error: string }[] = [];
  values.forEach((row, r) =>
    row.forEach((v, c) => {
      if (typeof v === "string" && ERROR_VALUES.has(v)) {
        out.push({ celda: toAddress(firstRow + r, firstCol + c), error: v });
      }
    }),
  );
  return out;
}
