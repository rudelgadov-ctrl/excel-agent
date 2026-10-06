import type Anthropic from "@anthropic-ai/sdk";

type Block = Anthropic.Beta.Messages.BetaContentBlock;
type BlockParam = Anthropic.Beta.Messages.BetaContentBlockParam;
type ToolUse = Anthropic.Beta.Messages.BetaToolUseBlock;

/**
 * Prepara la respuesta del modelo para devolverla al historial.
 *
 * Con `fallbacks`, si el modelo declina a mitad de la respuesta y otro modelo continúa,
 * aparece un bloque `fallback`. Antes del último de esos bloques solo se conservan los
 * textos (thinking y tool_use previos no se reenvían ni se ejecutan). Todo lo demás se
 * devuelve sin cambios, como exige el pensamiento preservado.
 */
export function prepareAssistantTurn(content: Block[]): { echo: BlockParam[]; toolUses: ToolUse[] } {
  let lastFallback = -1;
  content.forEach((block, i) => {
    if (block.type === "fallback") lastFallback = i;
  });
  const kept = content.filter((block, i) => i > lastFallback || block.type === "text");
  const toolUses = kept.filter((block): block is ToolUse => block.type === "tool_use");
  return { echo: kept as BlockParam[], toolUses };
}

/** Agrega contenido de usuario, fusionándolo si el último mensaje ya es del usuario. */
export function appendUserContent(
  messages: Anthropic.Beta.Messages.BetaMessageParam[],
  content: BlockParam[],
): void {
  const last = messages[messages.length - 1];
  if (last?.role === "user") {
    const previous: BlockParam[] =
      typeof last.content === "string" ? [{ type: "text", text: last.content }] : last.content;
    last.content = [...previous, ...content];
    return;
  }
  messages.push({ role: "user", content });
}
