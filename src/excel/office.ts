// Helpers de Office.js compartidos por las herramientas.
import { indexToColumn, isA1Address, parseRef } from "./address";

/** Obtiene un rango a partir de "Hoja!A1:B2", "A1" (hoja activa) o un nombre definido. */
export function getRange(context: Excel.RequestContext, ref: string): Excel.Range {
  const { sheet, address } = parseRef(ref);
  const wb = context.workbook;
  if (sheet !== undefined) return wb.worksheets.getItem(sheet).getRange(address);
  if (isA1Address(address)) return wb.worksheets.getActiveWorksheet().getRange(address);
  return wb.names.getItem(address).getRange();
}

export function getSheet(context: Excel.RequestContext, name?: string): Excel.Worksheet {
  return name ? context.workbook.worksheets.getItem(name) : context.workbook.worksheets.getActiveWorksheet();
}

/** Dirección A1 de una celda a partir de índices (base 0). */
export function cellAddress(row: number, col: number): string {
  return `${indexToColumn(col)}${row + 1}`;
}

let apiMinor: number | undefined;

/** Versión menor más alta de ExcelApi soportada (p. ej. 17 para ExcelApi 1.17). */
export function excelApiMinor(): number {
  if (apiMinor !== undefined) return apiMinor;
  let v = 0;
  try {
    for (let i = 1; i <= 30; i++) {
      if (!Office.context.requirements.isSetSupported("ExcelApi", `1.${i}`)) break;
      v = i;
    }
  } catch {
    v = 0;
  }
  apiMinor = v;
  return v;
}

export function supports(minor: number): boolean {
  return excelApiMinor() >= minor;
}

/** Convierte errores de Office.js en un mensaje útil para el agente. */
export async function describeOfficeError(err: unknown): Promise<string> {
  if (err instanceof OfficeExtension.Error) {
    let msg = `${err.code}: ${err.message}`;
    const loc = err.debugInfo?.errorLocation;
    if (loc) msg += ` (en ${loc})`;
    if (err.code === "ItemNotFound") {
      try {
        const names = await Excel.run(async (context) => {
          const sheets = context.workbook.worksheets;
          sheets.load("items/name");
          await context.sync();
          return sheets.items.map((s) => s.name);
        });
        msg += `. Hojas existentes: ${names.map((n) => JSON.stringify(n)).join(", ")}`;
      } catch {
        // sin información adicional
      }
    }
    return msg;
  }
  if (err instanceof Error) return err.message;
  return String(err);
}
