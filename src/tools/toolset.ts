import type Anthropic from "@anthropic-ai/sdk";
import { z } from "zod";
import { truncate } from "../excel/text";
import type { AgentTool, ToolEnv } from "./types";

type BetaTool = Anthropic.Beta.Messages.BetaTool;

export interface ToolOutcome {
  content: string;
  isError: boolean;
}

/** Herramientas de una aplicación (Excel o Word) con sus definiciones para la API. */
export interface Toolset {
  tools: AgentTool[];
  /** Definiciones para la API. Se calculan una vez para que el prefijo cacheado no cambie. */
  api: BetaTool[];
  isWrite(name: string): boolean;
  execute(name: string, input: unknown, opts: { allowWrites: boolean; env: ToolEnv }): Promise<ToolOutcome>;
}

function toApiTool(tool: AgentTool): BetaTool {
  const schema = z.toJSONSchema(tool.schema, { io: "input" }) as Record<string, unknown>;
  delete schema.$schema;
  return {
    name: tool.name,
    description: tool.description,
    input_schema: schema as BetaTool["input_schema"],
    // Las entradas grandes (código VBA, matrices de datos, HTML) llegan en streaming;
    // por eso se validan con Zod antes de ejecutarlas.
    eager_input_streaming: true,
  };
}

const fail = (content: string): ToolOutcome => ({ content, isError: true });

/** `describeError` convierte los errores de Office.js de esa aplicación en un mensaje útil para el agente. */
export function createToolset(tools: AgentTool[], describeError: (err: unknown) => Promise<string>): Toolset {
  const byName = new Map(tools.map((t) => [t.name, t]));
  return {
    tools,
    api: tools.map(toApiTool),
    isWrite: (name) => byName.get(name)?.writes ?? false,
    async execute(name, input, opts) {
      const tool = byName.get(name);
      if (!tool) return fail(`Herramienta desconocida: ${name}`);
      if (tool.writes && !opts.allowWrites) {
        return fail(
          "BLOQUEADO: el usuario todavía no aprobó el plan. Termina tu turno presentando el plan y espera " +
            "su aprobación; no vuelvas a intentar herramientas de escritura en este turno.",
        );
      }
      const parsed = tool.schema.safeParse(input);
      if (!parsed.success) {
        return fail(
          `Entrada inválida para ${name}:\n${z.prettifyError(parsed.error)}\n` +
            `Entrada recibida: ${truncate(JSON.stringify(input), 2000)}`,
        );
      }
      try {
        const result = await tool.run(parsed.data, opts.env);
        return { content: truncate(JSON.stringify(result), 40_000), isError: false };
      } catch (err) {
        return fail(await describeError(err));
      }
    },
  };
}
