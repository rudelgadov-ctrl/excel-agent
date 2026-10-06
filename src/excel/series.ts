// Helpers puros para el relleno automático cuando la API autoFill (ExcelApi 1.9) no existe.

/** Continúa una serie lineal a partir de 1 o más valores semilla. */
export function linearSeries(seed: number[], count: number): number[] {
  if (seed.length === 0) throw new Error("La serie necesita al menos un valor semilla");
  const step = seed.length >= 2 ? seed[seed.length - 1] - seed[seed.length - 2] : 1;
  const out = seed.slice(0, count);
  while (out.length < count) out.push(out[out.length - 1] + step);
  return out;
}

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;

export function serialToDate(serial: number): Date {
  return new Date(EXCEL_EPOCH_UTC + Math.round(serial) * DAY_MS);
}

export function dateToSerial(date: Date): number {
  return Math.round((date.getTime() - EXCEL_EPOCH_UTC) / DAY_MS);
}

/** Suma meses a un número de serie de Excel, igual que FECHA.MES (EDATE). */
export function addMonthsSerial(serial: number, months: number): number {
  const d = serialToDate(serial);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + months;
  const day = d.getUTCDate();
  const lastDay = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return dateToSerial(new Date(Date.UTC(y, m, Math.min(day, lastDay))));
}

export type DateUnit = "dias" | "meses" | "anios";

/** Serie de fechas (números de serie) a partir de una fecha inicial. */
export function dateSeries(start: number, count: number, unit: DateUnit, step = 1): number[] {
  return Array.from({ length: count }, (_, i) => {
    if (unit === "dias") return start + i * step;
    if (unit === "meses") return addMonthsSerial(start, i * step);
    return addMonthsSerial(start, i * step * 12);
  });
}

/** Repite un patrón 2D hasta llenar rows x cols (para copiar fórmulas R1C1). */
export function repeatPattern<T>(pattern: T[][], rows: number, cols: number): T[][] {
  const pr = pattern.length;
  const pc = pattern[0]?.length ?? 0;
  if (pr === 0 || pc === 0) throw new Error("Patrón vacío");
  return Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, c) => pattern[r % pr][c % pc]),
  );
}
