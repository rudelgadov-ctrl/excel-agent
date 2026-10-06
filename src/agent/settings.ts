// Ajustes por computadora. Se guardan en el localStorage del complemento (nunca en el repo).

export type Effort = "low" | "medium" | "high" | "xhigh" | "max";

export interface Settings {
  apiKey: string;
  model: string;
  effort: Effort;
  /** Ejecuta sin pedir aprobación del plan. */
  autoMode: boolean;
}

export interface ModelInfo {
  id: string;
  label: string;
  /** USD por millón de tokens. */
  price: { input: number; output: number; cacheRead: number; cacheWrite: number };
}

export const MODELS: ModelInfo[] = [
  {
    id: "claude-opus-5-5",
    label: "Claude Opus 5.5 (recomendado)",
    price: { input: 4, output: 20, cacheRead: 0.2, cacheWrite: 5 },
  },
  {
    id: "claude-sonnet-5-5",
    label: "Claude Sonnet 5.5 (más barato)",
    price: { input: 2, output: 10, cacheRead: 0.2, cacheWrite: 2.5 },
  },
];

export const EFFORTS: { id: Effort; label: string }[] = [
  { id: "medium", label: "Medio (más rápido)" },
  { id: "high", label: "Alto (recomendado)" },
  { id: "xhigh", label: "Muy alto" },
];

export const DEFAULT_SETTINGS: Settings = {
  apiKey: "",
  model: "claude-opus-5-5",
  effort: "high",
  autoMode: false,
};

const STORAGE_KEY = "excel-agent.settings.v1";

export function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    const saved = JSON.parse(raw) as Partial<Settings>;
    const merged = { ...DEFAULT_SETTINGS, ...saved };
    if (!MODELS.some((m) => m.id === merged.model)) merged.model = DEFAULT_SETTINGS.model;
    if (!EFFORTS.some((e) => e.id === merged.effort)) merged.effort = DEFAULT_SETTINGS.effort;
    return merged;
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(settings: Settings): boolean {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
    return true;
  } catch {
    return false;
  }
}

export function modelInfo(id: string): ModelInfo {
  return MODELS.find((m) => m.id === id) ?? MODELS[0];
}
