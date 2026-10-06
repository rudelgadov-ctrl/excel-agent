import type { z } from "zod";

/** Tarjeta con código VBA que el panel muestra al usuario. */
export interface VbaCard {
  nombre: string;
  codigo: string;
  descripcion?: string;
  comoEjecutar?: string;
}

/** Lo que el panel expone a las herramientas. */
export interface ToolEnv {
  showVba(card: VbaCard): void;
}

export interface AgentTool<S extends z.ZodType = z.ZodType> {
  name: string;
  description: string;
  schema: S;
  /** Modifica el libro: requiere que el usuario haya aprobado el plan. */
  writes: boolean;
  run(input: z.infer<S>, env: ToolEnv): Promise<unknown>;
}

export function defineTool<S extends z.ZodType>(tool: AgentTool<S>): AgentTool<S> {
  return tool;
}
