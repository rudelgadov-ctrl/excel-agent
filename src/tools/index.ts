import type Anthropic from "@anthropic-ai/sdk";
import { z } from "zod";
import { describeOfficeError } from "../excel/office";
import { truncate } from "../excel/text";
import { crearGrafico } from "./chart";
import { formatearRango } from "./format";
import { ejecutarOfficeJs, entregarVba } from "./misc";
import { leerLibro, leerRango } from "./read";
import { crearNombre, gestionarHojas, ordenar, validacionDatos } from "./structure";
import type { AgentTool, ToolEnv } from "./types";
import { copiarRango, dividirTexto, escribirRango, rellenarSerie } from "./write";

export type { ToolEnv, VbaCard } from "./types";

export const TOOLS: AgentTool[] = [
  leerLibro,
  leerRango,
  escribirRango,
  formatearRango,
  crearNombre,
  rellenarSerie,
  dividirTexto,
  copiarRango,
  gestionarHojas,
  validacionDatos,
  ordenar,
  crearGrafico,
  entregarVba,
  ejecutarOfficeJs,
];

type BetaTool = Anthropic.Beta.Messages.BetaTool;

function toApiTool(tool: AgentTool): BetaTool {
  const schema = z.toJSONSchema(tool.schema, { io: "input" }) as Record<string, unknown>;
  delete schema.$schema;
  return {
    name: tool.name,
    description: tool.description,
    input_schema: schema as BetaTool["input_schema"],
    // Las entradas grandes (código VBA, matrices de datos) llegan en streaming;
    // por eso se validan con Zod antes de ejecutarlas.
    eager_input_streaming: true,
  };
}

/** Definiciones para la API. Se calculan una vez para que el prefijo cacheado no cambie. */
export const API_TOOLS: BetaTool[] = TOOLS.map(toApiTool);

export function isWriteTool(name: string): boolean {
  return TOOLS.find((t) => t.name === name)?.writes ?? false;
}

export interface ToolOutcome {
  content: string;
  isError: boolean;
}

const fail = (content: string): ToolOutcome => ({ content, isError: true });

export async function executeTool(
  name: string,
  input: unknown,
  opts: { allowWrites: boolean; env: ToolEnv },
): Promise<ToolOutcome> {
  const tool = TOOLS.find((t) => t.name === name);
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
    return fail(await describeOfficeError(err));
  }
}
