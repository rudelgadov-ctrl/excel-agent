import Anthropic from "@anthropic-ai/sdk";
import { API_TOOLS, executeTool, type ToolEnv } from "../tools";
import { appendUserContent, prepareAssistantTurn } from "./content";
import { modelInfo, type ModelInfo, type Settings } from "./settings";
import { STATUS, SYSTEM_PROMPT } from "./systemPrompt";

type MessageParam = Anthropic.Beta.Messages.BetaMessageParam;
type ContentBlockParam = Anthropic.Beta.Messages.BetaContentBlockParam;
type ToolResult = Anthropic.Beta.Messages.BetaToolResultBlockParam;
type Usage = Anthropic.Beta.Messages.BetaUsage;

const MAX_STEPS = 80;
const BETAS = ["server-side-fallback-2026-07-01", "thinking-display-updates-2026-08-18"];

export interface UsageTotals {
  inputTokens: number;
  outputTokens: number;
  cacheReadTokens: number;
  cacheWriteTokens: number;
  costUsd: number;
}

export interface AgentEvents {
  /** Texto visible de la respuesta (streaming). */
  onText(delta: string): void;
  /** Notas de progreso entre llamadas a herramientas (thinking display "updates"). */
  onProgress(delta: string): void;
  onToolStart(id: string, name: string, input: unknown): void;
  onToolEnd(id: string, ok: boolean, output: string): void;
  onUsage(totals: UsageTotals): void;
  onNotice(text: string, kind: "info" | "error"): void;
}

export type TurnEnd = "done" | "refusal" | "aborted" | "error" | "limit";

/** reset: mensaje nuevo del usuario (pide plan); grant: aprobó el plan; keep: continuar igual. */
export type Approval = "reset" | "grant" | "keep";

const emptyUsage = (): UsageTotals => ({
  inputTokens: 0,
  outputTokens: 0,
  cacheReadTokens: 0,
  cacheWriteTokens: 0,
  costUsd: 0,
});

export class ExcelAgent {
  private messages: MessageParam[] = [];
  private controller: AbortController | null = null;
  approved = false;
  usage: UsageTotals = emptyUsage();

  constructor(private readonly env: ToolEnv) {}

  get busy(): boolean {
    return this.controller !== null;
  }

  stop(): void {
    this.controller?.abort();
  }

  /** Borra la conversación. Solo con el agente detenido: el bucle en curso escribe en el historial. */
  reset(): void {
    if (this.busy) throw new Error("Detén el agente antes de reiniciar la conversación.");
    this.messages = [];
    this.approved = false;
    this.usage = emptyUsage();
  }

  async send(
    content: ContentBlockParam[],
    settings: Settings,
    events: AgentEvents,
    approval: Approval = "reset",
  ): Promise<TurnEnd> {
    if (this.busy) throw new Error("El agente está ocupado.");
    if (approval === "grant") this.approved = true;
    if (approval === "reset") this.approved = false;
    const allowWrites = settings.autoMode || this.approved;
    const status = settings.autoMode ? STATUS.auto : this.approved ? STATUS.approved : STATUS.awaitingApproval;
    appendUserContent(this.messages, [...content, { type: "text", text: status }]);

    const controller = new AbortController();
    this.controller = controller;
    const client = new Anthropic({ apiKey: settings.apiKey, dangerouslyAllowBrowser: true, maxRetries: 3 });
    const model = modelInfo(settings.model);

    try {
      let jsonRetries = 0;
      for (let step = 0; step < MAX_STEPS; step++) {
        const stream = client.beta.messages.stream(
          {
            model: model.id,
            max_tokens: 64_000,
            system: SYSTEM_PROMPT,
            tools: API_TOOLS,
            messages: this.messages,
            thinking: { type: "adaptive", display: "updates" },
            output_config: { effort: settings.effort },
            cache_control: { type: "ephemeral" },
            fallbacks: "default",
            betas: BETAS,
          },
          { signal: controller.signal },
        );
        stream.on("text", (delta) => events.onText(delta));
        stream.on("thinking", (delta) => {
          if (delta) events.onProgress(delta);
        });

        let message: Anthropic.Beta.Messages.BetaMessage;
        try {
          message = await stream.finalMessage();
          jsonRetries = 0;
        } catch (err) {
          // Solo se reintenta cuando la entrada de una herramienta no se pudo interpretar.
          if (err instanceof Anthropic.APIError || controller.signal.aborted || jsonRetries++ >= 2) throw err;
          events.onNotice("La respuesta llegó incompleta; reintentando…", "info");
          continue;
        }
        this.addUsage(message.usage, model);
        events.onUsage(this.usage);

        if (message.stop_reason === "refusal") {
          events.onNotice("Claude declinó continuar con esta solicitud. Reformúlala o inicia una conversación nueva.", "error");
          return "refusal";
        }
        const { echo, toolUses } = prepareAssistantTurn(message.content);
        if (message.stop_reason === "pause_turn") {
          this.messages.push({ role: "assistant", content: echo });
          continue;
        }
        if (toolUses.length === 0) {
          if (echo.length) this.messages.push({ role: "assistant", content: echo });
          if (message.stop_reason === "max_tokens") events.onNotice("La respuesta se cortó por longitud.", "info");
          return "done";
        }
        if (message.stop_reason === "max_tokens") {
          events.onNotice("Una acción quedó incompleta por longitud. Pide el trabajo en partes más pequeñas.", "error");
          return "error";
        }

        this.messages.push({ role: "assistant", content: echo });
        const results: ToolResult[] = [];
        for (const toolUse of toolUses) {
          if (controller.signal.aborted) {
            results.push({ type: "tool_result", tool_use_id: toolUse.id, is_error: true, content: "Cancelado por el usuario." });
            continue;
          }
          events.onToolStart(toolUse.id, toolUse.name, toolUse.input);
          const out = await executeTool(toolUse.name, toolUse.input, { allowWrites, env: this.env });
          events.onToolEnd(toolUse.id, !out.isError, out.content);
          results.push({
            type: "tool_result",
            tool_use_id: toolUse.id,
            content: out.content,
            ...(out.isError ? { is_error: true } : {}),
          });
        }
        // Todas las respuestas de herramientas van en un solo mensaje.
        this.messages.push({ role: "user", content: results });
        if (controller.signal.aborted) return "aborted";
      }
      events.onNotice(`Se alcanzó el límite de ${MAX_STEPS} pasos. Presiona "Continuar" para seguir.`, "info");
      return "limit";
    } catch (err) {
      if (controller.signal.aborted || err instanceof Anthropic.APIUserAbortError) return "aborted";
      events.onNotice(describeApiError(err), "error");
      return "error";
    } finally {
      this.controller = null;
    }
  }

  private addUsage(usage: Usage, model: ModelInfo): void {
    const u = this.usage;
    const cacheRead = usage.cache_read_input_tokens ?? 0;
    const cacheWrite = usage.cache_creation_input_tokens ?? 0;
    u.inputTokens += usage.input_tokens;
    u.outputTokens += usage.output_tokens;
    u.cacheReadTokens += cacheRead;
    u.cacheWriteTokens += cacheWrite;
    const p = model.price;
    u.costUsd +=
      (usage.input_tokens * p.input +
        usage.output_tokens * p.output +
        cacheRead * p.cacheRead +
        cacheWrite * p.cacheWrite) /
      1_000_000;
  }
}

function describeApiError(err: unknown): string {
  if (err instanceof Anthropic.AuthenticationError) return "La API key no es válida o fue revocada. Revísala en Ajustes.";
  if (err instanceof Anthropic.PermissionDeniedError) return "Tu API key no tiene permiso para usar este modelo.";
  if (err instanceof Anthropic.RateLimitError) return "Se alcanzó el límite de uso de la API. Espera un momento y reintenta.";
  if (err instanceof Anthropic.BadRequestError) return `La API rechazó la solicitud: ${err.message}`;
  if (err instanceof Anthropic.InternalServerError) return "El servicio de Claude está saturado o con errores. Reintenta en unos minutos.";
  if (err instanceof Anthropic.APIConnectionError) return "No hay conexión con la API de Anthropic. Revisa internet o el firewall.";
  if (err instanceof Anthropic.APIError) return `Error de la API: ${err.message}`;
  return err instanceof Error ? err.message : String(err);
}
