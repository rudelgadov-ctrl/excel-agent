import type Anthropic from "@anthropic-ai/sdk";

type ContentBlockParam = Anthropic.Beta.Messages.BetaContentBlockParam;
type ImageMediaType = "image/png" | "image/jpeg" | "image/gif" | "image/webp";

export interface Attachment {
  name: string;
  mediaType: "application/pdf" | ImageMediaType;
  /** Contenido en base64 (sin el prefijo data:). */
  data: string;
  size: number;
}

export const MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024;

const BY_EXTENSION: Record<string, Attachment["mediaType"]> = {
  pdf: "application/pdf",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
};

/** Tipo MIME admitido por la API, o null. Algunos sistemas no informan el tipo del archivo. */
export function normalizeMediaType(name: string, type: string): Attachment["mediaType"] | null {
  const known = Object.values(BY_EXTENSION) as string[];
  if (known.includes(type)) return type as Attachment["mediaType"];
  const ext = name.split(".").pop()?.toLowerCase() ?? "";
  return BY_EXTENSION[ext] ?? null;
}

export const DEFAULT_REQUEST = "Ejecuta en este libro los requerimientos del archivo adjunto.";

/** Arma el contenido del mensaje: primero los adjuntos, luego el texto. */
export function buildUserContent(text: string, attachments: Attachment[]): ContentBlockParam[] {
  const blocks: ContentBlockParam[] = attachments.map((a) =>
    a.mediaType === "application/pdf"
      ? {
          type: "document",
          source: { type: "base64", media_type: "application/pdf", data: a.data },
          title: a.name,
        }
      : { type: "image", source: { type: "base64", media_type: a.mediaType, data: a.data } },
  );
  blocks.push({ type: "text", text: text.trim() || DEFAULT_REQUEST });
  return blocks;
}

export function readAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1] ?? "");
    reader.onerror = () => reject(reader.error ?? new Error("No se pudo leer el archivo"));
    reader.readAsDataURL(file);
  });
}

export function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}
