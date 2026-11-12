import type Anthropic from "@anthropic-ai/sdk";
import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { appendUserContent, prepareAssistantTurn } from "../src/agent/content";
import { HOSTS } from "../src/agent/hosts";
import { EXCEL_SYSTEM_PROMPT, WORD_SYSTEM_PROMPT } from "../src/agent/systemPrompt";
import { EXCEL_TOOLSET, WORD_TOOLSET, type ToolEnv, type VbaCard } from "../src/tools";
import { buildUserContent, classifyFile, decodeText, docxToText, xlsxToText } from "../src/taskpane/attachments";
import { renderMarkdown } from "../src/taskpane/markdown";

type Block = Anthropic.Beta.Messages.BetaContentBlock;
type MessageParam = Anthropic.Beta.Messages.BetaMessageParam;

const env: ToolEnv = { showVba: () => undefined, images: () => [] };

describe.each([
  ["Excel", EXCEL_TOOLSET],
  ["Word", WORD_TOOLSET],
])("herramientas de %s", (_, toolset) => {
  it("genera esquemas de entrada válidos para la API", () => {
    expect(toolset.api).toHaveLength(toolset.tools.length);
    const names = new Set(toolset.api.map((t) => t.name));
    expect(names.size).toBe(toolset.api.length);
    for (const tool of toolset.api) {
      expect(tool.input_schema.type).toBe("object");
      expect(tool.input_schema).not.toHaveProperty("$schema");
      expect(tool.eager_input_streaming).toBe(true);
      expect(tool.description?.length).toBeGreaterThan(20);
    }
  });

  it("los nombres de propiedades son válidos para la API (sin tildes ni ñ)", () => {
    const keys: string[] = [];
    const walk = (node: unknown): void => {
      if (Array.isArray(node)) return node.forEach(walk);
      if (!node || typeof node !== "object") return;
      const o = node as Record<string, unknown>;
      if (o.properties && typeof o.properties === "object") keys.push(...Object.keys(o.properties));
      Object.values(o).forEach(walk);
    };
    for (const tool of toolset.api) {
      expect(tool.name).toMatch(/^[a-zA-Z0-9_-]{1,64}$/);
      walk(tool.input_schema);
    }
    expect(keys.length).toBeGreaterThan(0);
    for (const key of keys) expect(key).toMatch(/^[a-zA-Z0-9_.-]{1,64}$/);
  });

  it("entregar_vba no requiere aprobación y llama al panel", async () => {
    const shown: string[] = [];
    const out = await toolset.execute(
      "entregar_vba",
      { nombre_macro: "Totales", codigo: "Sub Totales()\nEnd Sub" },
      { allowWrites: false, env: { ...env, showVba: (c) => shown.push(c.nombre) } },
    );
    expect(out.isError).toBe(false);
    expect(shown).toEqual(["Totales"]);
  });

  it("reporta herramientas desconocidas", async () => {
    const out = await toolset.execute("no_existe", {}, { allowWrites: true, env });
    expect(out).toEqual({ content: "Herramienta desconocida: no_existe", isError: true });
  });

  it("solo las herramientas de lectura y las que muestran VBA funcionan sin aprobación", () => {
    for (const tool of toolset.tools) {
      const readOnly = tool.name.startsWith("leer_") || ["entregar_vba", "macro_modelo_datos"].includes(tool.name);
      expect(toolset.isWrite(tool.name), tool.name).toBe(!readOnly);
    }
  });
});

describe("herramientas de Excel", () => {
  it("los campos con valor por defecto son opcionales", () => {
    const leerRango = EXCEL_TOOLSET.api.find((t) => t.name === "leer_rango")!;
    expect(leerRango.input_schema.required).toEqual(["rango"]);
  });

  it("bloquea herramientas de escritura sin aprobación", async () => {
    const out = await EXCEL_TOOLSET.execute("escribir_rango", { rango: "A1", datos: [[1]] }, { allowWrites: false, env });
    expect(out.isError).toBe(true);
    expect(out.content).toMatch(/^BLOQUEADO/);
  });

  it("rechaza entradas inválidas con un mensaje útil", async () => {
    const out = await EXCEL_TOOLSET.execute("escribir_rango", { rango: "A1" }, { allowWrites: true, env });
    expect(out.isError).toBe(true);
    expect(out.content).toContain("datos");
  });

  it("macro_modelo_datos muestra la macro (solo Windows si usa el modelo) y devuelve los pasos", async () => {
    const cards: VbaCard[] = [];
    const out = await EXCEL_TOOLSET.execute(
      "macro_modelo_datos",
      {
        nombre_macro: "ArmarModelo",
        tablas_modelo: ["Ventas", "Vendedores"],
        relaciones: [{ foranea: "Ventas[IdVendedor]", primaria: "Vendedores[IdVendedor]" }],
        tablas_dinamicas: [
          {
            nombre: "TD_Region",
            destino: "'Resumen'!A3",
            filas: ["Vendedores[Region]"],
            valores: [{ campo: "Ventas[Importe]" }],
          },
        ],
        graficos: [{ tabla_dinamica: "TD_Region", rango: "'Panel'!B2:H15", leyenda: false, etiquetas_datos: true }],
      },
      { allowWrites: false, env: { ...env, showVba: (c) => cards.push(c) } },
    );
    expect(out.isError).toBe(false);
    expect(cards).toHaveLength(1);
    expect(cards[0]).toMatchObject({ nombre: "ArmarModelo", soloWindows: true });
    expect(cards[0].codigo).toContain("Public Sub ArmarModelo()");
    const result = JSON.parse(out.content);
    expect(result.pasos).toHaveLength(5);
    expect(result.valores_por_tabla_dinamica).toEqual({ TD_Region: ["Suma de Importe"] });
  });

  it("macro_modelo_datos explica los errores de la especificación", async () => {
    const out = await EXCEL_TOOLSET.execute(
      "macro_modelo_datos",
      {
        tablas_dinamicas: [
          {
            nombre: "TD",
            destino: "'Resumen'!A3",
            filas: ["Ventas[Fecha]"],
            valores: [{ campo: "Ventas[Importe]" }],
            agrupar_fechas: [{ campo: "Ventas[Fecha]", periodos: ["Trimestres"] }],
          },
        ],
      },
      { allowWrites: false, env },
    );
    expect(out).toMatchObject({ isError: true, content: expect.stringContaining("columnas calculadas") });
  });
});

describe("herramientas de Word", () => {
  it("el objetivo es opcional y sus campos tienen valores por defecto", () => {
    const insertar = WORD_TOOLSET.api.find((t) => t.name === "insertar_contenido")!;
    expect(insertar.input_schema.required ?? []).toEqual([]);
    const formato = WORD_TOOLSET.api.find((t) => t.name === "formato_texto")!;
    expect(formato.input_schema.required).toEqual(["objetivo"]);
  });

  it("bloquea herramientas de escritura sin aprobación", async () => {
    const out = await WORD_TOOLSET.execute(
      "insertar_contenido",
      { parrafos: [{ texto: "Hola" }] },
      { allowWrites: false, env },
    );
    expect(out.isError).toBe(true);
    expect(out.content).toMatch(/^BLOQUEADO/);
  });

  it("exige exactamente un tipo de contenido al insertar", async () => {
    const out = await WORD_TOOLSET.execute("insertar_contenido", { texto: "a", html: "<p>b</p>" }, { allowWrites: true, env });
    expect(out).toMatchObject({ isError: true, content: expect.stringContaining("exactamente uno") });
  });

  it("valida el objetivo antes de tocar el documento", async () => {
    const out = await WORD_TOOLSET.execute(
      "formato_texto",
      { objetivo: { zona: "margen" }, negrita: true },
      { allowWrites: true, env },
    );
    expect(out.isError).toBe(true);
    expect(out.content).toContain("zona");
  });

  it("insertar_imagen avisa si no hay imágenes adjuntas", async () => {
    const out = await WORD_TOOLSET.execute("insertar_imagen", { adjunto: "logo.png" }, { allowWrites: true, env });
    expect(out).toMatchObject({ isError: true, content: expect.stringContaining("no adjuntó imágenes") });
    const withImage: ToolEnv = { ...env, images: () => [{ name: "Foto.PNG", data: "", mediaType: "image/webp" }] };
    const other = await WORD_TOOLSET.execute("insertar_imagen", { adjunto: "otra.png" }, { allowWrites: true, env: withImage });
    expect(other.content).toContain("«Foto.PNG»");
  });
});

describe("historial", () => {
  const text = (t: string) => ({ type: "text", text: t, citations: null }) as Block;
  const toolUse = (id: string) =>
    ({ type: "tool_use", id, name: "leer_libro", input: {}, caller: undefined }) as unknown as Block;
  const thinking = { type: "thinking", thinking: "", signature: "x" } as Block;
  const fallback = { type: "fallback", from: { model: "a" }, to: { model: "b" } } as unknown as Block;

  it("sin fallback devuelve el contenido sin cambios", () => {
    const content = [thinking, text("hola"), toolUse("t1")];
    const { echo, toolUses } = prepareAssistantTurn(content);
    expect(echo).toEqual(content);
    expect(toolUses.map((t) => t.id)).toEqual(["t1"]);
  });

  it("con fallback descarta thinking y tool_use anteriores al último bloque fallback", () => {
    const content = [thinking, text("parcial"), toolUse("viejo"), fallback, thinking, toolUse("nuevo")];
    const { echo, toolUses } = prepareAssistantTurn(content);
    expect(echo).toEqual([text("parcial"), thinking, toolUse("nuevo")]);
    expect(toolUses.map((t) => t.id)).toEqual(["nuevo"]);
  });

  it("fusiona mensajes de usuario consecutivos", () => {
    const messages: MessageParam[] = [{ role: "user", content: "primero" }];
    appendUserContent(messages, [{ type: "text", text: "segundo" }]);
    expect(messages).toEqual([
      {
        role: "user",
        content: [
          { type: "text", text: "primero" },
          { type: "text", text: "segundo" },
        ],
      },
    ]);
    messages.push({ role: "assistant", content: "ok" });
    appendUserContent(messages, [{ type: "text", text: "tercero" }]);
    expect(messages).toHaveLength(3);
  });
});

describe("adjuntos", () => {
  it("clasifica por extensión aunque el sistema no informe el tipo", () => {
    expect(classifyFile("instrucciones.PDF", "")).toEqual({ kind: "pdf" });
    expect(classifyFile("practica.docx", "")).toEqual({ kind: "word" });
    expect(classifyFile("notas.txt", "")).toEqual({ kind: "text" });
    expect(classifyFile("foto.jpg", "")).toEqual({ kind: "image", mediaType: "image/jpeg" });
    expect(classifyFile("datos.XLSX", "")).toEqual({ kind: "excel" });
    expect(classifyFile("datos.csv", "")).toEqual({ kind: "text" });
    expect(classifyFile("viejo.xls", "application/vnd.ms-excel")).toBeNull();
  });

  it("lee texto en UTF-8, UTF-16 o Windows-1252 sin perder tildes", () => {
    const text = "Región\tAño\nSan Ramón\t2024";
    expect(decodeText(new TextEncoder().encode(text))).toBe(text);
    expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode(text)]))).toBe(text);
    const latin1 = Uint8Array.from([...text].map((c) => c.charCodeAt(0)));
    expect(decodeText(latin1)).toBe(text);
    const utf16 = new Uint8Array([0xff, 0xfe, ...[...text].flatMap((c) => [c.charCodeAt(0), 0])]);
    expect(decodeText(utf16)).toBe(text);
  });

  it("convierte un .xlsx adjunto en texto con fechas ISO", () => {
    const sheet =
      '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>' +
      '<row r="2"><c r="A2" t="inlineStr"><is><t>Ana &amp; Luis</t></is></c><c r="B2"><v>12.5</v></c>' +
      '<c r="C2" s="1"><v>45366</v></c></row></sheetData></worksheet>';
    const xlsx = zipSync({
      "xl/workbook.xml": strToU8('<workbook><sheets><sheet name="Datos" sheetId="1" r:id="rId1"/></sheets></workbook>'),
      "xl/_rels/workbook.xml.rels": strToU8('<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'),
      "xl/sharedStrings.xml": strToU8("<sst><si><t>Nombre</t></si><si><r><t>Fe</t></r><r><t>cha</t></r></si></sst>"),
      "xl/styles.xml": strToU8('<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>'),
      "xl/worksheets/sheet1.xml": strToU8(sheet),
    });
    const text = xlsxToText(xlsx);
    expect(text).toContain("### Hoja «Datos»");
    expect(text).toContain("1\tNombre\t\tFecha");
    expect(text).toContain("2\tAna & Luis\t12.5\t2024-03-15");
  });

  it("nombra las imágenes adjuntas y usa la solicitud por defecto de la aplicación", () => {
    const content = buildUserContent("", [{ name: "logo.png", kind: "image", data: "QUJD", mediaType: "image/png", size: 3 }], "documento");
    expect(content.map((b) => b.type)).toEqual(["text", "image", "text"]);
    expect(content[0]).toMatchObject({ text: expect.stringContaining("«logo.png»") });
    expect(content[2]).toMatchObject({ text: expect.stringContaining("este documento") });
  });

  it("pone los adjuntos antes del texto y usa documentos de texto para Word", () => {
    const content = buildUserContent("", [
      { name: "p.pdf", kind: "pdf", data: "QUJD", size: 3 },
      { name: "p.docx", kind: "word", data: "a) Sume la columna B", size: 10 },
    ]);
    expect(content.map((b) => b.type)).toEqual(["document", "document", "text"]);
    expect(content[1]).toMatchObject({ source: { type: "text", data: "a) Sume la columna B" } });
    expect(content[2]).toMatchObject({ type: "text", text: expect.stringContaining("adjunto") });
  });

  it("extrae el texto de un .docx (párrafos, tablas, entidades)", () => {
    const xml =
      '<w:document><w:body><w:p><w:r><w:t>Ejercicio #1</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t xml:space="preserve">a) Calcule el </w:t></w:r><w:r><w:t>total &amp; el IVA</w:t></w:r></w:p>' +
      "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Mes</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Ventas</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
      '<w:p><w:r><w:instrText> PAGE </w:instrText></w:r></w:p></w:body></w:document>';
    const docx = zipSync({ "word/document.xml": strToU8(xml), "[Content_Types].xml": strToU8("<Types/>") });
    const text = docxToText(docx);
    expect(text).toContain("Ejercicio #1");
    expect(text).toContain("a) Calcule el total & el IVA");
    expect(text).toMatch(/Mes\s*\n?\s*\|\s*\n?\s*Ventas/);
    expect(text).not.toContain("PAGE");
  });
});

describe("generalidad", () => {
  // El agente debe servir para cualquier práctica: nada en el prompt ni en las
  // herramientas puede depender de las prácticas usadas como referencia.
  const specific =
    /Amortizaci|Inventario|Facturar|Librer[ií]a|Prec\. por|San Isidro|DN0102|Quiz|Cr[eé]ditos?\b|Clientes|Provincia|Cant[oó]n|AnalisisDinamico|IdCredito|Britannic|XZY|Clasificaciones|Tri1/i;

  it.each(Object.values(HOSTS))("el prompt de $app no menciona prácticas concretas", ({ systemPrompt, app }) => {
    expect(systemPrompt).not.toMatch(specific);
    expect(systemPrompt).toMatch(/CUALQUIER práctica/);
    expect(systemPrompt).toContain(`Agente ${app}`);
  });

  it.each(Object.values(HOSTS))("el prompt de $app no pide aprobación del plan", ({ systemPrompt }) => {
    expect(systemPrompt).toContain("No hay paso de aprobación");
    expect(systemPrompt).not.toMatch(/esperando aprobación|BLOQUEADO|plan aprobado/i);
  });

  it.each(Object.values(HOSTS))("las herramientas de $app no mencionan prácticas concretas", ({ toolset }) => {
    expect(JSON.stringify(toolset.api)).not.toMatch(specific);
  });

  it("cada prompt menciona todas sus herramientas y ninguna de la otra aplicación", () => {
    for (const host of Object.values(HOSTS)) {
      const own = new Set(host.toolset.tools.map((t) => t.name));
      for (const name of own) expect(host.systemPrompt, `${host.app}: ${name}`).toContain(name);
      const other = Object.values(HOSTS).find((h) => h !== host)!;
      for (const t of other.toolset.tools) {
        if (!own.has(t.name)) expect(host.systemPrompt, `${host.app}: ${t.name}`).not.toContain(t.name);
      }
    }
    expect(EXCEL_SYSTEM_PROMPT).not.toBe(WORD_SYSTEM_PROMPT);
  });

  it("cubre los temas habituales de un curso de Word", () => {
    const names = WORD_TOOLSET.api.map((t) => t.name);
    for (const name of [
      "leer_documento",
      "insertar_contenido",
      "formato_texto",
      "formato_parrafo",
      "aplicar_lista",
      "modificar_estilo",
      "insertar_tabla",
      "editar_tabla",
      "formato_tabla",
      "configurar_pagina",
      "insertar_salto",
      "insertar_campo",
      "insertar_vinculo_o_nota",
      "insertar_imagen",
      "entregar_vba",
    ]) {
      expect(names).toContain(name);
    }
  });

  it("cubre los temas habituales de un curso de Excel", () => {
    const names = EXCEL_TOOLSET.api.map((t) => t.name);
    for (const name of [
      "formato_condicional",
      "tabla_dinamica",
      "crear_tabla",
      "filtrar",
      "ordenar",
      "validacion_datos",
      "crear_grafico",
      "crear_nombre",
      "rellenar_serie",
      "insertar_eliminar",
      "quitar_duplicados",
      "configurar_hoja",
      "entregar_vba",
      "limpiar_datos",
      "modificar_tabla",
      "segmentacion",
      "gestionar_formas",
      "macro_modelo_datos",
    ]) {
      expect(names).toContain(name);
    }
  });
});

describe("markdown", () => {
  it("escapa HTML", () => {
    expect(renderMarkdown("<script>alert(1)</script>")).toBe("<p>&lt;script&gt;alert(1)&lt;/script&gt;</p>");
  });

  it("renderiza listas, negritas, código y tablas", () => {
    const html = renderMarkdown(
      "## Plan\n1. **Formato** de `B2`\n2. Fórmulas\n   - PMT\n\n| Celda | Valor |\n|---|---|\n| B2 | 5 |",
    );
    expect(html).toContain("<h4>Plan</h4>");
    expect(html).toContain("<ol><li><strong>Formato</strong> de <code>B2</code></li><li>Fórmulas<ul><li>PMT</li></ul></li></ol>");
    expect(html).toContain("<table><thead><tr><th>Celda</th><th>Valor</th></tr></thead><tbody><tr><td>B2</td><td>5</td></tr></tbody></table>");
  });

  it("no aplica formato dentro de bloques de código", () => {
    expect(renderMarkdown("```\n**x** <b>\n```")).toBe("<pre><code>**x** &lt;b&gt;</code></pre>");
  });
});
