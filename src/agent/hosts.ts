// Lo que cambia según la aplicación donde se abre el complemento.
import { EXCEL_TOOLSET, WORD_TOOLSET, type Toolset } from "../tools";
import { EXCEL_SYSTEM_PROMPT, WORD_SYSTEM_PROMPT } from "./systemPrompt";

export type HostKind = "excel" | "word";

export interface HostProfile {
  kind: HostKind;
  /** Nombre de la aplicación. */
  app: "Excel" | "Word";
  /** Cómo se llama el archivo abierto: "libro" o "documento". */
  doc: string;
  systemPrompt: string;
  toolset: Toolset;
}

export const HOSTS: Record<HostKind, HostProfile> = {
  excel: { kind: "excel", app: "Excel", doc: "libro", systemPrompt: EXCEL_SYSTEM_PROMPT, toolset: EXCEL_TOOLSET },
  word: { kind: "word", app: "Word", doc: "documento", systemPrompt: WORD_SYSTEM_PROMPT, toolset: WORD_TOOLSET },
};
