import { describe, expect, it } from "vitest";
import { fieldParagraph, wrapOoxml } from "../src/word/ooxml";
import {
  BUILT_IN_STYLES,
  builtInStyle,
  cmToPt,
  lineSpacingMultiple,
  lineSpacingPt,
  parseIndexList,
  ptToCm,
  runsForDeletion,
  splitFieldCode,
  styleNameCandidates,
} from "../src/word/spec";

describe("parseIndexList", () => {
  it("acepta números, intervalos y listas", () => {
    expect(parseIndexList("3", 10)).toEqual([3]);
    expect(parseIndexList("2-5", 10)).toEqual([2, 3, 4, 5]);
    expect(parseIndexList(" 7-8, 1 ,4,8 ", 10)).toEqual([1, 4, 7, 8]);
  });
  it("rechaza números fuera del documento o mal escritos", () => {
    expect(() => parseIndexList("11", 10)).toThrow(/hay 10/);
    expect(() => parseIndexList("0", 10)).toThrow();
    expect(() => parseIndexList("5-3", 10)).toThrow();
    expect(() => parseIndexList("a", 10)).toThrow(/inválido/);
    expect(() => parseIndexList(" , ", 10)).toThrow();
    expect(() => parseIndexList("4", 3, "fila")).toThrow(/fila 4/);
  });
});

describe("unidades", () => {
  it("convierte cm y puntos", () => {
    expect(cmToPt(2.54)).toBe(72);
    expect(ptToCm(72)).toBe(2.54);
    expect(cmToPt(1.25)).toBeCloseTo(35.43, 2);
  });
  it("expresa el interlineado múltiple en puntos", () => {
    expect(lineSpacingPt(1.5)).toBe(18);
    expect(lineSpacingPt(2)).toBe(24);
    expect(lineSpacingMultiple(13.8)).toBe(1.15);
  });
});

describe("estilos", () => {
  it("reconoce estilos integrados en inglés y en español", () => {
    expect(builtInStyle("Heading 1")).toBe("Heading1");
    expect(builtInStyle("heading1")).toBe("Heading1");
    expect(builtInStyle("Título 2")).toBe("Heading2");
    expect(builtInStyle("TITULO")).toBe("Title");
    expect(builtInStyle("Sin espaciado")).toBe("NoSpacing");
    expect(builtInStyle("Párrafo de lista")).toBe("ListParagraph");
    expect(builtInStyle("GridTable4 Accent1")).toBe("GridTable4_Accent1");
    expect(builtInStyle("Tabla con cuadrícula")).toBe("TableGrid");
    expect(builtInStyle("Mi estilo")).toBeUndefined();
  });
  it("incluye los estilos de tabla con sus énfasis", () => {
    expect(BUILT_IN_STYLES).toContain("ListTable7Colorful_Accent6");
    expect(new Set(BUILT_IN_STYLES).size).toBe(BUILT_IN_STYLES.length);
  });
  it("propone nombres alternativos para buscar el estilo en el documento", () => {
    expect(styleNameCandidates("Heading1")).toEqual(["Heading1", "Heading 1", "Título 1"]);
    expect(styleNameCandidates("Título")).toEqual(["Título", "Title"]);
    expect(styleNameCandidates("Cita personal")).toEqual(["Cita personal"]);
  });
});

describe("campos", () => {
  it("separa el tipo de campo de sus modificadores", () => {
    expect(splitFieldCode('TOC \\o "1-3" \\h \\z \\u')).toEqual({ type: "TOC", switches: '\\o "1-3" \\h \\z \\u' });
    expect(splitFieldCode(" page ")).toEqual({ type: "Page", switches: "" });
    expect(splitFieldCode("NUMPAGES")).toEqual({ type: "NumPages", switches: "" });
    expect(splitFieldCode("SEQ Figura \\* ARABIC")).toEqual({ type: "Seq", switches: "Figura \\* ARABIC" });
    expect(splitFieldCode("= SUM(ABOVE)")).toEqual({ switches: "= SUM(ABOVE)" });
    expect(splitFieldCode("NOEXISTE x")).toEqual({ switches: "NOEXISTE x" });
  });
  it("arma un campo complejo marcado para actualizarse", () => {
    const xml = fieldParagraph('DATE \\@ "dd/MM/yyyy"');
    expect(xml).toContain('w:fldCharType="begin" w:dirty="true"');
    expect(xml).toContain("DATE \\@ &quot;dd/MM/yyyy&quot;");
    expect(xml.match(/<w:fldChar /g)).toHaveLength(3);
  });
});

describe("OOXML", () => {
  it("envuelve contenido del cuerpo en un paquete", () => {
    const pkg = wrapOoxml('<?xml version="1.0"?><w:p><w:r><w:t>Hola</w:t></w:r></w:p>');
    expect(pkg.startsWith("<pkg:package")).toBe(true);
    expect(pkg).toContain("<w:body><w:p><w:r><w:t>Hola</w:t></w:r></w:p></w:body>");
    expect(pkg).toContain('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
    expect(pkg).not.toContain("<?xml");
  });
  it("respeta paquetes y documentos completos", () => {
    const full = '<pkg:package xmlns:pkg="x"><pkg:part/></pkg:package>';
    expect(wrapOoxml(full)).toBe(full);
    const doc = '<w:document xmlns:w="w"><w:body><w:p/></w:body></w:document>';
    expect(wrapOoxml(doc)).toContain(`<pkg:xmlData>${doc}</pkg:xmlData>`);
    expect(wrapOoxml("<w:body><w:p/></w:body>")).toContain("<w:body><w:p/></w:body></w:document>");
  });
});

describe("tablas", () => {
  it("agrupa filas consecutivas para borrarlas de abajo hacia arriba", () => {
    expect(runsForDeletion([2, 3, 4, 7, 9, 10])).toEqual([
      [9, 2],
      [7, 1],
      [2, 3],
    ]);
    expect(runsForDeletion([5])).toEqual([[5, 1]]);
  });
});
