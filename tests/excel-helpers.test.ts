import { describe, expect, it } from "vitest";
import { columnToIndex, indexToColumn, isA1Address, parseRef, quoteSheet } from "../src/excel/address";
import { addMonthsSerial, dateSeries, dateToSerial, linearSeries, repeatPattern } from "../src/excel/series";
import { findErrors, splitRows, truncate } from "../src/excel/text";

describe("parseRef", () => {
  it("separa hojas con comillas, espacios y símbolos", () => {
    expect(parseRef("'#1 - Amortización'!A1:E7")).toEqual({ sheet: "#1 - Amortización", address: "A1:E7" });
    expect(parseRef("'#2 - Prec. por Mes'!$B$3:$F$14")).toEqual({ sheet: "#2 - Prec. por Mes", address: "B3:F14" });
  });
  it("respeta comillas escapadas en el nombre", () => {
    expect(parseRef("'Hoja de ''Ana'''!C3")).toEqual({ sheet: "Hoja de 'Ana'", address: "C3" });
  });
  it("acepta hojas sin comillas y referencias sin hoja", () => {
    expect(parseRef("Hoja1!B2")).toEqual({ sheet: "Hoja1", address: "B2" });
    expect(parseRef("B2:C4")).toEqual({ address: "B2:C4" });
    expect(parseRef("MontoPrestamo")).toEqual({ address: "MontoPrestamo" });
  });
  it("rechaza referencias mal cerradas", () => {
    expect(() => parseRef("'Hoja sin cierre!A1")).toThrow();
  });
});

describe("direcciones", () => {
  it("distingue direcciones de nombres definidos", () => {
    expect(isA1Address("A1")).toBe(true);
    expect(isA1Address("aa10:ab20")).toBe(true);
    expect(isA1Address("C:C")).toBe(true);
    expect(isA1Address("3:5")).toBe(true);
    expect(isA1Address("Tasa")).toBe(false);
    expect(isA1Address("Cuota_Mensual")).toBe(false);
  });
  it("convierte columnas en ambos sentidos", () => {
    for (const [letter, index] of [["A", 0], ["Z", 25], ["AA", 26], ["AZ", 51], ["XFD", 16383]] as const) {
      expect(columnToIndex(letter)).toBe(index);
      expect(indexToColumn(index)).toBe(letter);
    }
  });
  it("pone comillas al nombre de hoja solo cuando hace falta", () => {
    expect(quoteSheet("Hoja1")).toBe("Hoja1");
    expect(quoteSheet("#1 - Factura")).toBe("'#1 - Factura'");
  });
});

describe("series", () => {
  it("continúa series lineales", () => {
    expect(linearSeries([1], 5)).toEqual([1, 2, 3, 4, 5]);
    expect(linearSeries([2, 4], 4)).toEqual([2, 4, 6, 8]);
  });
  it("suma meses como FECHA.MES, ajustando fin de mes", () => {
    const jan31 = dateToSerial(new Date(Date.UTC(2025, 0, 31)));
    const feb28 = dateToSerial(new Date(Date.UTC(2025, 1, 28)));
    expect(addMonthsSerial(jan31, 1)).toBe(feb28);
    // 45658 = 01/01/2025 en Excel
    expect(dateToSerial(new Date(Date.UTC(2025, 0, 1)))).toBe(45658);
  });
  it("genera fechas mensuales consecutivas", () => {
    const start = dateToSerial(new Date(Date.UTC(2025, 9, 6)));
    const months = dateSeries(start, 3, "meses");
    expect(months[1]).toBe(dateToSerial(new Date(Date.UTC(2025, 10, 6))));
    expect(months[2]).toBe(dateToSerial(new Date(Date.UTC(2025, 11, 6))));
  });
  it("repite patrones para copiar fórmulas", () => {
    expect(repeatPattern([["=R[-1]C+1"]], 3, 1)).toEqual([["=R[-1]C+1"], ["=R[-1]C+1"], ["=R[-1]C+1"]]);
  });
});

describe("texto", () => {
  it("divide texto en columnas con relleno", () => {
    expect(
      splitRows([["Alejandro Dumas;Los Tres Mosqueteros"], ["Julio Verne;De la Tierra a la Luna "], [""]], ";"),
    ).toEqual([
      ["Alejandro Dumas", "Los Tres Mosqueteros"],
      ["Julio Verne", "De la Tierra a la Luna"],
      ["", ""],
    ]);
  });
  it("encuentra celdas con error en inglés y español", () => {
    const errors = findErrors([[1, "#NAME?"], ["#¡VALOR!", "ok"]], 4, 1, (r, c) => `${"ABCDEF"[c]}${r + 1}`);
    expect(errors).toEqual([
      { celda: "C5", error: "#NAME?" },
      { celda: "B6", error: "#¡VALOR!" },
    ]);
  });
  it("recorta textos largos", () => {
    expect(truncate("abcdef", 3)).toContain("recortado: 3");
    expect(truncate("abc", 3)).toBe("abc");
  });
});
