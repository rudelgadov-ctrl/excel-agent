// Limpieza de datos (depuración): helpers puros para limpiar_datos.

export interface CleanOptions {
  recortar_espacios: boolean;
  quitar_no_imprimibles: boolean;
  texto_a_numero: boolean;
  separador_decimal: "punto" | "coma";
  texto_a_fecha: boolean;
  orden_fecha: "dia_mes_anio" | "mes_dia_anio";
  mayusculas?: "mayusculas" | "minusculas" | "nombre_propio";
}

export type Cleaned =
  | { kind: "text"; value: string }
  | { kind: "number"; value: number }
  | { kind: "date"; value: number; time: boolean };

/** Como NOMPROPIO: mayúscula al inicio de cada palabra y el resto en minúscula. */
export function properCase(text: string): string {
  let out = "";
  let prevLetter = false;
  for (const ch of text.toLocaleLowerCase("es")) {
    const isLetter = /\p{L}/u.test(ch);
    out += isLetter && !prevLetter ? ch.toLocaleUpperCase("es") : ch;
    prevLetter = isLetter;
  }
  return out;
}

export function cleanText(text: string, opts: Pick<CleanOptions, "recortar_espacios" | "quitar_no_imprimibles" | "mayusculas">): string {
  let s = text;
  // eslint-disable-next-line no-control-regex
  if (opts.quitar_no_imprimibles) s = s.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f​-‍﻿]/g, "");
  if (opts.recortar_espacios) s = s.replace(/[\s   ]+/g, " ").trim();
  if (opts.mayusculas === "mayusculas") s = s.toLocaleUpperCase("es");
  else if (opts.mayusculas === "minusculas") s = s.toLocaleLowerCase("es");
  else if (opts.mayusculas === "nombre_propio") s = properCase(s);
  return s;
}

/** Número escrito como texto ("1,234.50", "1.234,50", "₡ 2 500", "15%"); null si no lo es. */
export function parseNumber(text: string, decimal: "punto" | "coma"): number | null {
  let s = text.replace(/[\s   ]/g, "").replace(/^[$€₡£¢]|[$€₡£¢]$/g, "");
  let percent = false;
  if (s.endsWith("%")) {
    percent = true;
    s = s.slice(0, -1);
  }
  const pattern = decimal === "punto" ? /^[-+]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/ : /^[-+]?(\d{1,3}(\.\d{3})+|\d+)(,\d+)?$/;
  if (!pattern.test(s)) return null;
  s = decimal === "punto" ? s.replace(/,/g, "") : s.replace(/\./g, "").replace(",", ".");
  const n = Number(s);
  if (!Number.isFinite(n)) return null;
  return percent ? n / 100 : n;
}

const EPOCH = Date.UTC(1899, 11, 30);

/** Número de serie de Excel para una fecha (y hora) válida; null si no existe. */
export function excelSerial(y: number, m: number, d: number, hh = 0, mm = 0, ss = 0): number | null {
  if (m < 1 || m > 12 || d < 1 || d > 31 || hh > 23 || mm > 59 || ss > 59) return null;
  const t = Date.UTC(y, m - 1, d);
  const date = new Date(t);
  if (date.getUTCFullYear() !== y || date.getUTCMonth() !== m - 1 || date.getUTCDate() !== d) return null;
  return (t - EPOCH) / 86_400_000 + (hh * 3600 + mm * 60 + ss) / 86_400;
}

/** Fecha escrita como texto: ISO (2024-03-15 [10:30:00]) o con barras (15/03/2024). */
export function parseDate(text: string, order: CleanOptions["orden_fecha"]): { serial: number; time: boolean } | null {
  const s = text.trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
  let parts: number[] | null = null;
  if (m) parts = [Number(m[1]), Number(m[2]), Number(m[3]), Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0)];
  else {
    m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
    if (m) {
      const a = Number(m[1]);
      const b = Number(m[2]);
      let y = Number(m[3]);
      if (m[3].length === 2) y += y < 30 ? 2000 : 1900;
      const [d, mo] = order === "dia_mes_anio" ? [a, b] : [b, a];
      parts = [y, mo, d, Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0)];
    }
  }
  if (!parts) return null;
  const serial = excelSerial(parts[0], parts[1], parts[2], parts[3], parts[4], parts[5]);
  if (serial === null) return null;
  return { serial, time: parts[3] + parts[4] + parts[5] > 0 };
}

/** true si Excel convertiría este texto en número, fecha u hora al escribirlo. */
export function excelWouldConvert(text: string): boolean {
  const s = text.trim();
  return (
    /^[-+]?[$€₡£¢]?\s?[\d.,]+\s?%?$/.test(s) ||
    /^\d{1,4}[/-]\d{1,2}([/-]\d{1,4})?/.test(s) ||
    /^\d{1,2}:\d{2}/.test(s) ||
    /^(true|false|verdadero|falso)$/i.test(s)
  );
}

/** Limpia el valor de una celda. Devuelve null si no cambia. Solo procesa textos. */
export function cleanCell(value: unknown, opts: CleanOptions): Cleaned | null {
  if (typeof value !== "string" || value === "") return null;
  const text = cleanText(value, opts);
  if (opts.texto_a_fecha) {
    const d = parseDate(text, opts.orden_fecha);
    if (d) return { kind: "date", value: d.serial, time: d.time };
  }
  if (opts.texto_a_numero) {
    const n = parseNumber(text, opts.separador_decimal);
    if (n !== null) return { kind: "number", value: n };
  }
  return text === value ? null : { kind: "text", value: text };
}
