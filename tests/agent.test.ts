import type Anthropic from "@anthropic-ai/sdk";
import { describe, expect, it } from "vitest";
import { appendUserContent, prepareAssistantTurn } from "../src/agent/content";
import { API_TOOLS, TOOLS, executeTool } from "../src/tools";
import { buildUserContent, normalizeMediaType } from "../src/taskpane/attachments";
import { renderMarkdown } from "../src/taskpane/markdown";

type Block = Anthropic.Beta.Messages.BetaContentBlock;
type MessageParam = Anthropic.Beta.Messages.BetaMessageParam;

const env = { showVba: () => undefined };

describe("herramientas", () => {
  it("genera esquemas de entrada válidos para la API", () => {
    expect(API_TOOLS).toHaveLength(TOOLS.length);
    const names = new Set(API_TOOLS.map((t) => t.name));
    expect(names.size).toBe(API_TOOLS.length);
    for (const tool of API_TOOLS) {
      expect(tool.input_schema.type).toBe("object");
      expect(tool.input_schema).not.toHaveProperty("$schema");
      expect(tool.eager_input_streaming).toBe(true);
      expect(tool.description?.length).toBeGreaterThan(20);
    }
  });

  it("los campos con valor por defecto son opcionales", () => {
    const leerRango = API_TOOLS.find((t) => t.name === "leer_rango")!;
    expect(leerRango.input_schema.required).toEqual(["rango"]);
  });

  it("bloquea herramientas de escritura sin aprobación", async () => {
    const out = await executeTool("escribir_rango", { rango: "A1", datos: [[1]] }, { allowWrites: false, env });
    expect(out.isError).toBe(true);
    expect(out.content).toMatch(/^BLOQUEADO/);
  });

  it("rechaza entradas inválidas con un mensaje útil", async () => {
    const out = await executeTool("escribir_rango", { rango: "A1" }, { allowWrites: true, env });
    expect(out.isError).toBe(true);
    expect(out.content).toContain("datos");
  });

  it("entregar_vba no requiere aprobación y llama al panel", async () => {
    const shown: string[] = [];
    const out = await executeTool(
      "entregar_vba",
      { nombre_macro: "Facturar", codigo: "Sub Facturar()\nEnd Sub" },
      { allowWrites: false, env: { showVba: (c) => shown.push(c.nombre) } },
    );
    expect(out.isError).toBe(false);
    expect(shown).toEqual(["Facturar"]);
  });

  it("reporta herramientas desconocidas", async () => {
    const out = await executeTool("no_existe", {}, { allowWrites: true, env });
    expect(out).toEqual({ content: "Herramienta desconocida: no_existe", isError: true });
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
  it("detecta el tipo por extensión si el sistema no lo informa", () => {
    expect(normalizeMediaType("instrucciones.PDF", "")).toBe("application/pdf");
    expect(normalizeMediaType("foto.jpg", "")).toBe("image/jpeg");
    expect(normalizeMediaType("hoja.xlsx", "application/vnd.ms-excel")).toBeNull();
  });

  it("pone los adjuntos antes del texto", () => {
    const content = buildUserContent("", [
      { name: "p.pdf", mediaType: "application/pdf", data: "QUJD", size: 3 },
    ]);
    expect(content.map((b) => b.type)).toEqual(["document", "text"]);
    expect(content[1]).toMatchObject({ type: "text", text: expect.stringContaining("adjunto") });
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
