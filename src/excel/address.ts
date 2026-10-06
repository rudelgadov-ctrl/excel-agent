// Helpers puros para referencias tipo Excel ("'#1 - Hoja'!A1:B2", "C3", "MiNombre").

export interface ParsedRef {
  /** Nombre de la hoja sin comillas, si la referencia lo incluye. */
  sheet?: string;
  /** Dirección dentro de la hoja ("A1:B2") o nombre definido. */
  address: string;
}

/** Separa "Hoja!A1" / "'Hoja con espacios'!A1:B2" en hoja y dirección. */
export function parseRef(ref: string): ParsedRef {
  const text = ref.trim();
  if (text.startsWith("'")) {
    // Nombre entre comillas simples; '' representa una comilla literal.
    let i = 1;
    let sheet = "";
    while (i < text.length) {
      if (text[i] === "'") {
        if (text[i + 1] === "'") {
          sheet += "'";
          i += 2;
          continue;
        }
        break;
      }
      sheet += text[i];
      i++;
    }
    const rest = text.slice(i + 1);
    if (!rest.startsWith("!")) throw new Error(`Referencia inválida: ${ref}`);
    return { sheet, address: stripDollar(rest.slice(1)) };
  }
  const bang = text.lastIndexOf("!");
  if (bang >= 0) {
    return { sheet: text.slice(0, bang), address: stripDollar(text.slice(bang + 1)) };
  }
  return { address: stripDollar(text) };
}

function stripDollar(address: string): string {
  return address.replace(/\$/g, "").trim();
}

const CELL_OR_RANGE = /^[A-Z]{1,3}\d+(:[A-Z]{1,3}\d+)?$|^[A-Z]{1,3}:[A-Z]{1,3}$|^\d+:\d+$/i;

/** true si el texto es una dirección A1 (no un nombre definido). */
export function isA1Address(address: string): boolean {
  return CELL_OR_RANGE.test(address.trim());
}

/** "A" -> 0, "Z" -> 25, "AA" -> 26 */
export function columnToIndex(column: string): number {
  let n = 0;
  for (const ch of column.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) throw new Error(`Columna inválida: ${column}`);
    n = n * 26 + (code - 64);
  }
  return n - 1;
}

/** 0 -> "A", 25 -> "Z", 26 -> "AA" */
export function indexToColumn(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Pone comillas al nombre de hoja cuando hace falta para usarlo en una fórmula. */
export function quoteSheet(sheet: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
}

/** Crea una matriz rows x cols rellena con el mismo valor. */
export function fillMatrix<T>(rows: number, cols: number, value: T): T[][] {
  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => value));
}
