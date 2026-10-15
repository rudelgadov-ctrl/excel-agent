import "./styles.css";
import type Anthropic from "@anthropic-ai/sdk";
import { HOSTS, type HostKind, type HostProfile } from "../agent/hosts";
import { OfficeAgent, type AgentEvents, type Approval, type TurnEnd } from "../agent/loop";
import { EFFORTS, MODELS, loadSettings, modelInfo, saveSettings, type Effort, type Settings } from "../agent/settings";
import { excelApiMinor } from "../excel/office";
import { truncate } from "../excel/text";
import type { ImageFile, ToolEnv, VbaCard } from "../tools";
import { wordApiMinor } from "../word/office";
import {
  ACCEPT_ATTRIBUTE,
  MAX_ATTACHMENT_BYTES,
  buildUserContent,
  classifyFile,
  formatSize,
  readAttachment,
  type Attachment,
} from "./attachments";
import { renderMarkdown } from "./markdown";

type ContentBlockParam = Anthropic.Beta.Messages.BetaContentBlockParam;

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

const ui = {
  log: $("log"),
  empty: $("empty"),
  brandName: $("brand-name"),
  emptyTitle: $("empty-title"),
  emptyScope: $("empty-scope"),
  emptyOpen: $("empty-open"),
  hostWarning: $("host-warning"),
  settings: $("settings"),
  setKey: $<HTMLInputElement>("set-key"),
  setModel: $<HTMLSelectElement>("set-model"),
  setEffort: $<HTMLSelectElement>("set-effort"),
  setAuto: $<HTMLInputElement>("set-auto"),
  settingsMsg: $("settings-msg"),
  actionBar: $("action-bar"),
  btnApprove: $<HTMLButtonElement>("btn-approve"),
  btnContinue: $<HTMLButtonElement>("btn-continue"),
  actionHint: $("action-hint"),
  composer: $("composer"),
  attachments: $("attachments"),
  input: $<HTMLTextAreaElement>("input"),
  file: $<HTMLInputElement>("file"),
  status: $("status"),
  btnSend: $<HTMLButtonElement>("btn-send"),
  btnStop: $<HTMLButtonElement>("btn-stop"),
  meta: $("meta"),
};

const TOOL_LABELS: Record<string, string> = {
  leer_libro: "Leyendo el libro",
  leer_rango: "Leyendo",
  escribir_rango: "Escribiendo",
  formatear_rango: "Aplicando formato",
  crear_nombre: "Nombre definido",
  rellenar_serie: "Relleno automático",
  dividir_texto: "Texto en columnas",
  copiar_rango: "Copiando",
  gestionar_hojas: "Hojas",
  validacion_datos: "Validación de datos",
  ordenar: "Ordenando",
  crear_grafico: "Gráfico",
  insertar_eliminar: "Insertar o eliminar",
  quitar_duplicados: "Quitando duplicados",
  formato_condicional: "Formato condicional",
  crear_tabla: "Tabla",
  tabla_dinamica: "Tabla dinámica",
  filtrar: "Filtrando",
  configurar_hoja: "Configurando la hoja",
  leer_documento: "Leyendo el documento",
  leer_tabla: "Leyendo la tabla",
  insertar_contenido: "Escribiendo",
  eliminar_contenido: "Eliminando",
  formato_texto: "Formato de texto",
  formato_parrafo: "Formato de párrafo",
  aplicar_lista: "Lista",
  modificar_estilo: "Estilo",
  insertar_tabla: "Insertando tabla",
  editar_tabla: "Editando tabla",
  formato_tabla: "Formato de tabla",
  configurar_pagina: "Configurando la página",
  insertar_salto: "Salto",
  insertar_campo: "Campo",
  insertar_vinculo_o_nota: "Vínculo o nota",
  insertar_imagen: "Imagen",
  entregar_vba: "Macro VBA",
  ejecutar_office_js: "Código Office.js",
};

let settings: Settings = loadSettings();
let attachments: Attachment[] = [];
let currentTurn: TurnView | null = null;
/** El panel está dentro de Excel o Word (y no en un navegador u otra aplicación). */
let inHost = false;
let resetWhenIdle = false;
/** Imágenes adjuntadas en esta conversación, para que el agente pueda insertarlas en el archivo. */
let sessionImages: ImageFile[] = [];

const env: ToolEnv = {
  showVba(card) {
    const turn = currentTurn ?? new TurnView();
    turn.addBlock(vbaCard(card, agent.host.kind));
  },
  images: () => sessionImages,
};

// Se reemplaza al arrancar si el complemento se abrió en Word.
let agent = new OfficeAgent(HOSTS.excel, env);

// ---------- utilidades de DOM ----------

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function isNearBottom(): boolean {
  const { scrollTop, scrollHeight, clientHeight } = ui.log;
  return scrollHeight - scrollTop - clientHeight < 80;
}

function appendToLog(node: HTMLElement): void {
  const stick = isNearBottom();
  ui.empty.hidden = true;
  ui.log.append(node);
  if (stick) ui.log.scrollTop = ui.log.scrollHeight;
}

function keepScrolled(stick: boolean): void {
  if (stick) ui.log.scrollTop = ui.log.scrollHeight;
}

async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // Algunos visores de Office bloquean la API del portapapeles.
  }
  const area = el("textarea");
  area.value = text;
  area.style.position = "fixed";
  area.style.opacity = "0";
  document.body.append(area);
  area.select();
  let ok = false;
  try {
    ok = document.execCommand("copy");
  } catch {
    ok = false;
  }
  area.remove();
  return ok;
}

function prettyJson(raw: unknown): string {
  const text = typeof raw === "string" ? raw : JSON.stringify(raw);
  try {
    return truncate(JSON.stringify(JSON.parse(text), null, 2), 4000);
  } catch {
    return truncate(text, 4000);
  }
}

/** Resumen corto del objetivo de una herramienta de Word ({ parrafos, tabla, texto, zona… }). */
function describeObjetivo(o: Record<string, unknown>): string {
  const parts: string[] = [];
  if (o.zona === "encabezado" || o.zona === "pie") parts.push(o.zona === "pie" ? "pie de página" : "encabezado");
  if (o.seleccion) parts.push("selección");
  if (typeof o.parrafos === "string") parts.push(`párrafos ${o.parrafos}`);
  if (typeof o.tabla === "number") parts.push(`tabla ${o.tabla}`);
  if (typeof o.texto === "string") parts.push(`«${truncate(o.texto, 40)}»`);
  return parts.join(" · ") || "documento";
}

function toolTarget(input: unknown): string {
  if (!input || typeof input !== "object") return "";
  const o = input as Record<string, unknown>;
  for (const key of ["rango", "destino", "origen", "hoja", "nombre", "nombre_macro", "adjunto", "codigo", "descripcion"]) {
    if (typeof o[key] === "string") return truncate(o[key] as string, 60);
  }
  if (o.objetivo && typeof o.objetivo === "object") return describeObjetivo(o.objetivo as Record<string, unknown>);
  if (typeof o.tabla === "number") return `tabla ${o.tabla}`;
  return "";
}

// ---------- vista de un turno del agente ----------

class TurnView {
  readonly root = el("div", "msg assistant");
  private text: { node: HTMLDivElement; raw: string } | null = null;
  private progress: HTMLDivElement | null = null;
  private tools = new Map<string, HTMLDetailsElement>();
  private renderQueued = false;

  constructor() {
    appendToLog(this.root);
  }

  appendText(delta: string): void {
    this.progress = null;
    if (!this.text) {
      const node = el("div", "md");
      this.root.append(node);
      this.text = { node, raw: "" };
    }
    this.text.raw += delta;
    if (!this.renderQueued) {
      this.renderQueued = true;
      requestAnimationFrame(() => {
        this.renderQueued = false;
        this.renderText();
      });
    }
  }

  appendProgress(delta: string): void {
    this.closeText();
    const stick = isNearBottom();
    if (!this.progress) {
      this.progress = el("div", "progress");
      this.root.append(this.progress);
    }
    this.progress.textContent += delta;
    keepScrolled(stick);
  }

  toolStart(id: string, name: string, input: unknown): void {
    this.closeText();
    this.progress = null;
    const stick = isNearBottom();
    const details = el("details", "tool running");
    const summary = el("summary");
    summary.append(el("span", "icon", "…"), el("span", "label", TOOL_LABELS[name] ?? name));
    const target = toolTarget(input);
    if (target) summary.append(el("span", "target", target));
    details.append(summary, el("pre", undefined, prettyJson(input)));
    this.root.append(details);
    this.tools.set(id, details);
    keepScrolled(stick);
  }

  toolEnd(id: string, ok: boolean, output: string): void {
    const details = this.tools.get(id);
    if (!details) return;
    details.classList.remove("running");
    details.classList.add(ok ? "ok" : "fail");
    const icon = details.querySelector(".icon");
    if (icon) icon.textContent = ok ? "✓" : "✗";
    details.append(el("pre", undefined, prettyJson(output)));
  }

  addBlock(node: HTMLElement): void {
    this.closeText();
    this.progress = null;
    const stick = isNearBottom();
    this.root.append(node);
    keepScrolled(stick);
  }

  notice(text: string, kind: "info" | "error"): void {
    this.addBlock(el("div", `notice ${kind}`, text));
  }

  finish(): void {
    this.closeText();
    // Las herramientas que no terminaron (p. ej. al detener) quedan marcadas.
    for (const details of this.tools.values()) {
      if (details.classList.contains("running")) {
        details.classList.replace("running", "fail");
        const icon = details.querySelector(".icon");
        if (icon) icon.textContent = "✗";
      }
    }
  }

  private renderText(): void {
    if (!this.text) return;
    const stick = isNearBottom();
    this.text.node.innerHTML = renderMarkdown(this.text.raw);
    keepScrolled(stick);
  }

  private closeText(): void {
    if (this.text) this.renderText();
    this.text = null;
  }
}

const VBA_STEPS: Record<HostKind, { windows: string[]; mac: string[] }> = {
  excel: {
    windows: [
      "Archivo > Guardar como > «Libro de Excel habilitado para macros (*.xlsm)».",
      "Alt+F11 abre el editor de VBA. Menú Insertar > Módulo.",
      "Pega el código y cierra el editor.",
      "Ejecuta con Alt+F8 > {macro} > Ejecutar (o asígnala a un botón: Programador > Insertar > Botón).",
    ],
    mac: [
      "Archivo > Guardar como > formato «Libro de Excel habilitado para macros (.xlsm)».",
      "Herramientas > Macro > Editor de Visual Basic. Menú Insert > Module.",
      "Pega el código y cierra el editor.",
      "Herramientas > Macro > Macros… > {macro} > Ejecutar.",
    ],
  },
  word: {
    windows: [
      "Archivo > Guardar como > «Documento habilitado con macros de Word (*.docm)».",
      "Alt+F11 abre el editor de VBA. En el panel de la izquierda selecciona tu documento y luego Insertar > Módulo.",
      "Pega el código y cierra el editor.",
      "Ejecuta con Alt+F8 > {macro} > Ejecutar (o agrégala a la barra de acceso rápido).",
    ],
    mac: [
      "Archivo > Guardar como > formato «Documento de Word habilitado para macros (.docm)».",
      "Herramientas > Macro > Editor de Visual Basic. Selecciona tu documento y luego Insert > Module.",
      "Pega el código y cierra el editor.",
      "Herramientas > Macro > Macros… > {macro} > Ejecutar.",
    ],
  },
};

function vbaCard(card: VbaCard, host: HostKind): HTMLElement {
  const box = el("div", "vba");
  box.append(el("h4", undefined, `Macro VBA: ${card.nombre}`));
  if (card.descripcion) box.append(el("p", undefined, card.descripcion));
  box.append(el("pre", "code", card.codigo));

  const copy = el("button", "primary small", "Copiar código");
  copy.type = "button";
  copy.addEventListener("click", async () => {
    const ok = await copyText(card.codigo);
    copy.textContent = ok ? "Copiado ✓" : "No se pudo copiar: selecciona el código";
    setTimeout(() => (copy.textContent = "Copiar código"), 2500);
  });
  const row = el("div", "row");
  row.append(copy);
  box.append(row);

  const steps = el("details");
  steps.append(el("summary", undefined, "Cómo instalarla"));
  const list = (title: string, items: string[]) => {
    const wrap = el("div");
    wrap.append(el("strong", undefined, title));
    const ol = el("ol");
    for (const item of items) ol.append(el("li", undefined, item));
    wrap.append(ol);
    return wrap;
  };
  const fill = (items: string[]) => items.map((item) => item.replace("{macro}", card.nombre));
  steps.append(list("Windows", fill(VBA_STEPS[host].windows)), list("Mac", fill(VBA_STEPS[host].mac)));
  box.append(steps);
  if (card.comoEjecutar) box.append(el("p", "muted", card.comoEjecutar));
  return box;
}

// ---------- estado de la interfaz ----------

function setBusy(busy: boolean): void {
  ui.btnSend.disabled = busy || !inHost;
  ui.btnStop.hidden = !busy;
  ui.btnApprove.disabled = busy;
  ui.btnContinue.disabled = busy;
}

function setStatus(text: string): void {
  ui.status.textContent = text;
}

function updateMeta(): void {
  const model = modelInfo(settings.model).label.replace(/\s*\(.*\)$/, "");
  const api = !inHost ? "" : agent.host.kind === "word" ? ` · WordApi 1.${wordApiMinor()}` : ` · ExcelApi 1.${excelApiMinor()}`;
  const cost = agent.usage.costUsd > 0 ? ` · ≈ US$${agent.usage.costUsd.toFixed(2)} en esta conversación` : "";
  const mode = settings.autoMode ? " · modo automático" : "";
  ui.meta.textContent = `${model}${api}${mode}${cost}`;
}

function showActionBar(end: TurnEnd | null): void {
  const planPending = end === "done" && !settings.autoMode && !agent.approved;
  const canContinue = end === "limit";
  ui.btnApprove.hidden = !planPending;
  ui.actionHint.hidden = !planPending;
  ui.btnContinue.hidden = !canContinue;
  ui.actionBar.hidden = !(planPending || canContinue);
}

function renderUserMessage(text: string, files: Attachment[] = []): void {
  const msg = el("div", "msg user", text);
  if (files.length) {
    const chips = el("div", "chips");
    for (const f of files) chips.append(el("span", "chip", `📎 ${f.name}`));
    msg.append(chips);
  }
  appendToLog(msg);
}

function renderAttachments(): void {
  ui.attachments.replaceChildren(
    ...attachments.map((a, i) => {
      const chip = el("span", "chip", `📎 ${a.name} (${formatSize(a.size)})`);
      const remove = el("button", undefined, "×");
      remove.type = "button";
      remove.title = "Quitar";
      remove.addEventListener("click", () => {
        attachments.splice(i, 1);
        renderAttachments();
      });
      chip.append(remove);
      return chip;
    }),
  );
}

function flashLogNotice(text: string, kind: "info" | "error" = "error"): void {
  appendToLog(el("div", `notice ${kind}`, text));
}

async function addFiles(files: Iterable<File>): Promise<void> {
  for (const file of files) {
    if (!classifyFile(file.name, file.type)) {
      flashLogNotice(`No se puede adjuntar «${file.name}»: usa PDF, Word (.docx), texto o imágenes.`);
      continue;
    }
    const total = attachments.reduce((sum, a) => sum + a.size, 0) + file.size;
    if (total > MAX_ATTACHMENT_BYTES) {
      flashLogNotice(`«${file.name}» supera el límite de ${formatSize(MAX_ATTACHMENT_BYTES)} en adjuntos.`);
      continue;
    }
    try {
      const attachment = await readAttachment(file);
      if (attachment) attachments.push(attachment);
    } catch (err) {
      flashLogNotice(`No se pudo leer «${file.name}»: ${err instanceof Error ? err.message : String(err)}`);
    }
  }
  renderAttachments();
}

// ---------- flujo de conversación ----------

async function runTurn(
  content: ContentBlockParam[],
  approval: Approval,
  label: string,
  files: Attachment[] = [],
): Promise<void> {
  renderUserMessage(label, files);
  showActionBar(null);
  const turn = new TurnView();
  currentTurn = turn;
  setBusy(true);
  setStatus("Pensando…");

  const events: AgentEvents = {
    onText: (delta) => {
      turn.appendText(delta);
      setStatus("Escribiendo…");
    },
    onProgress: (delta) => turn.appendProgress(delta),
    onToolStart: (id, name, input) => {
      turn.toolStart(id, name, input);
      setStatus(`${TOOL_LABELS[name] ?? name}…`);
    },
    onToolEnd: (id, ok, output) => {
      turn.toolEnd(id, ok, output);
      setStatus("Pensando…");
    },
    onUsage: () => updateMeta(),
    onNotice: (text, kind) => turn.notice(text, kind),
  };

  let end: TurnEnd;
  try {
    end = await agent.send(content, settings, events, approval);
  } catch (err) {
    turn.notice(err instanceof Error ? err.message : String(err), "error");
    end = "error";
  }
  if (end === "aborted") turn.notice("Detenido por el usuario.", "info");
  turn.finish();
  currentTurn = null;
  setBusy(false);
  setStatus("");
  updateMeta();
  showActionBar(end);
  if (resetWhenIdle) {
    resetWhenIdle = false;
    onNewConversation();
  }
}

function requireApiKey(): boolean {
  if (settings.apiKey.trim()) return true;
  openSettings(true);
  ui.settingsMsg.textContent = "Primero guarda tu API key.";
  ui.setKey.focus();
  return false;
}

async function onSend(): Promise<void> {
  if (agent.busy || !inHost) return;
  const text = ui.input.value.trim();
  if (!text && attachments.length === 0) return;
  if (!requireApiKey()) return;
  const files = attachments;
  const content = buildUserContent(text, files, agent.host.doc);
  for (const f of files) {
    if (f.kind !== "image") continue;
    sessionImages = [...sessionImages.filter((i) => i.name !== f.name), { name: f.name, data: f.data, mediaType: f.mediaType ?? "image/png" }];
  }
  ui.input.value = "";
  attachments = [];
  renderAttachments();
  await runTurn(content, "reset", text || "(instrucciones en el adjunto)", files);
}

async function onApprove(): Promise<void> {
  if (agent.busy || !requireApiKey()) return;
  await runTurn([{ type: "text", text: "Plan aprobado. Ejecútalo completo." }], "grant", "✅ Plan aprobado");
}

async function onContinue(): Promise<void> {
  if (agent.busy || !requireApiKey()) return;
  await runTurn([{ type: "text", text: "Continúa donde te quedaste." }], "keep", "Continuar");
}

function onNewConversation(): void {
  if (agent.busy) {
    // Se reinicia cuando el turno en curso termine de detenerse.
    resetWhenIdle = true;
    agent.stop();
    return;
  }
  agent.reset();
  sessionImages = [];
  currentTurn = null;
  ui.log.replaceChildren(ui.empty);
  ui.empty.hidden = false;
  showActionBar(null);
  setStatus("");
  updateMeta();
  ui.input.focus();
}

// ---------- ajustes ----------

function openSettings(open = ui.settings.hidden): void {
  ui.settings.hidden = !open;
  if (open) {
    ui.setKey.value = settings.apiKey;
    ui.setModel.value = settings.model;
    ui.setEffort.value = settings.effort;
    ui.setAuto.checked = settings.autoMode;
    ui.settingsMsg.textContent = "";
  }
}

function onSaveSettings(): void {
  settings = {
    apiKey: ui.setKey.value.trim(),
    model: ui.setModel.value,
    effort: ui.setEffort.value as Effort,
    autoMode: ui.setAuto.checked,
  };
  const stored = saveSettings(settings);
  ui.settingsMsg.textContent = stored ? "Guardado." : "No se pudo guardar en este equipo; se usará solo en esta sesión.";
  updateMeta();
  if (stored && settings.apiKey) setTimeout(() => openSettings(false), 600);
}

function initSettingsForm(): void {
  for (const m of MODELS) ui.setModel.append(new Option(m.label, m.id));
  for (const e of EFFORTS) ui.setEffort.append(new Option(e.label, e.id));
}

// ---------- arranque ----------

function wireEvents(): void {
  ui.btnSend.addEventListener("click", () => void onSend());
  ui.btnStop.addEventListener("click", () => agent.stop());
  ui.btnApprove.addEventListener("click", () => void onApprove());
  ui.btnContinue.addEventListener("click", () => void onContinue());
  $("btn-new").addEventListener("click", onNewConversation);
  $("btn-settings").addEventListener("click", () => openSettings());
  $("btn-save-settings").addEventListener("click", onSaveSettings);
  $("btn-toggle-key").addEventListener("click", (e) => {
    const show = ui.setKey.type === "password";
    ui.setKey.type = show ? "text" : "password";
    (e.currentTarget as HTMLButtonElement).textContent = show ? "Ocultar" : "Ver";
  });
  ui.input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      void onSend();
    }
  });
  ui.file.addEventListener("change", () => {
    if (ui.file.files) void addFiles(Array.from(ui.file.files));
    ui.file.value = "";
  });
  ui.composer.addEventListener("dragover", (e) => {
    e.preventDefault();
    ui.composer.classList.add("dragging");
  });
  ui.composer.addEventListener("dragleave", () => ui.composer.classList.remove("dragging"));
  ui.composer.addEventListener("drop", (e) => {
    e.preventDefault();
    ui.composer.classList.remove("dragging");
    if (e.dataTransfer?.files.length) void addFiles(Array.from(e.dataTransfer.files));
  });
}

/** Requisitos mínimos de la API (el manifiesto ya no los exige para poder cargarse en ambas aplicaciones). */
const MIN_API: Record<HostKind, [string, string]> = { excel: ["ExcelApi", "1.7"], word: ["WordApi", "1.3"] };

function applyHost(host: HostProfile): void {
  const name = `Agente ${host.app}`;
  document.title = name;
  ui.brandName.textContent = name;
  ui.emptyTitle.textContent = `¿Qué hay que hacer en este ${host.doc}?`;
  ui.emptyScope.textContent = `Sirve para cualquier práctica, tarea o requerimiento de ${host.app}.`;
  ui.emptyOpen.textContent =
    host.kind === "word" ? "Abre la plantilla o el documento en Word." : "Abre la plantilla o el libro en Excel.";
}

function start(officeHost: Office.HostType | null): void {
  const kind: HostKind | null =
    officeHost === Office.HostType.Excel ? "excel" : officeHost === Office.HostType.Word ? "word" : null;
  inHost = kind !== null;
  if (kind) {
    agent = new OfficeAgent(HOSTS[kind], env);
    applyHost(HOSTS[kind]);
    const [set, version] = MIN_API[kind];
    if (!Office.context.requirements.isSetSupported(set, version)) {
      ui.hostWarning.textContent =
        `Esta versión de ${HOSTS[kind].app} es demasiado antigua (requiere ${set} ${version}): usa Microsoft 365 o ` +
        `${HOSTS[kind].app} 2019 o posterior. Algunas funciones pueden fallar.`;
      ui.hostWarning.hidden = false;
    } else {
      ui.hostWarning.hidden = true;
    }
  } else {
    ui.hostWarning.hidden = false;
  }
  ui.file.accept = ACCEPT_ATTRIBUTE;
  initSettingsForm();
  wireEvents();
  setBusy(false);
  updateMeta();
  if (!settings.apiKey) openSettings(true);
}

if (typeof Office === "undefined") {
  // office.js no cargó (sin internet o abierto fuera de Excel).
  ui.hostWarning.hidden = false;
  ui.btnSend.disabled = true;
} else {
  Office.onReady((info) => start(info.host));
}
