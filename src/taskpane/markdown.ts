// Renderizador de Markdown mínimo para las respuestas del agente.
// Escapa todo el HTML primero; solo genera etiquetas propias.

export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function inline(raw: string): string {
  // El contenido entre comillas invertidas no recibe más formato.
  return raw
    .split(/(`[^`]+`)/)
    .map((part) => {
      if (part.startsWith("`") && part.endsWith("`") && part.length > 1) {
        return `<code>${escapeHtml(part.slice(1, -1))}</code>`;
      }
      return escapeHtml(part)
        .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
        .replace(/(^|[^*\w])\*([^*\s][^*]*?)\*(?!\w)/g, "$1<em>$2</em>");
    })
    .join("");
}

function tableCells(line: string): string[] {
  return line
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((c) => inline(c.trim()));
}

const TABLE_SEPARATOR = /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/;
const BULLET = /^(\s*)[-*•]\s+(.*)$/;
const NUMBERED = /^(\s*)\d+[.)]\s+(.*)$/;
const HEADING = /^(#{1,6})\s+(.*)$/;

export function renderMarkdown(source: string): string {
  const lines = source.replace(/\r\n?/g, "\n").split("\n");
  const out: string[] = [];
  let paragraph: string[] = [];
  const lists: { type: "ul" | "ol"; indent: number }[] = [];

  const flushParagraph = () => {
    if (paragraph.length) out.push(`<p>${paragraph.map(inline).join("<br>")}</p>`);
    paragraph = [];
  };
  const closeLists = (downTo = -1) => {
    while (lists.length && lists[lists.length - 1].indent > downTo) {
      out.push(`</li></${lists.pop()!.type}>`);
    }
  };
  const listItem = (type: "ul" | "ol", indent: number, text: string) => {
    flushParagraph();
    closeLists(indent);
    let top = lists[lists.length - 1];
    if (top && top.indent === indent && top.type !== type) {
      out.push(`</li></${lists.pop()!.type}>`);
      top = lists[lists.length - 1];
    }
    if (!top || top.indent < indent) {
      out.push(`<${type}><li>${inline(text)}`);
      lists.push({ type, indent });
    } else {
      out.push(`</li><li>${inline(text)}`);
    }
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (line.trim().startsWith("```")) {
      flushParagraph();
      closeLists();
      const code: string[] = [];
      i++;
      while (i < lines.length && !lines[i].trim().startsWith("```")) code.push(lines[i++]);
      out.push(`<pre><code>${escapeHtml(code.join("\n"))}</code></pre>`);
      continue;
    }

    if (line.trim().startsWith("|") && i + 1 < lines.length && TABLE_SEPARATOR.test(lines[i + 1])) {
      flushParagraph();
      closeLists();
      const header = tableCells(line);
      const rows: string[][] = [];
      i += 2;
      while (i < lines.length && lines[i].trim().startsWith("|")) rows.push(tableCells(lines[i++]));
      i--;
      out.push(
        `<table><thead><tr>${header.map((h) => `<th>${h}</th>`).join("")}</tr></thead><tbody>` +
          rows.map((r) => `<tr>${r.map((c) => `<td>${c}</td>`).join("")}</tr>`).join("") +
          "</tbody></table>",
      );
      continue;
    }

    const heading = HEADING.exec(line);
    if (heading) {
      flushParagraph();
      closeLists();
      const level = Math.min(6, heading[1].length + 2);
      out.push(`<h${level}>${inline(heading[2])}</h${level}>`);
      continue;
    }

    const bullet = BULLET.exec(line);
    if (bullet) {
      listItem("ul", bullet[1].length, bullet[2]);
      continue;
    }
    const numbered = NUMBERED.exec(line);
    if (numbered) {
      listItem("ol", numbered[1].length, numbered[2]);
      continue;
    }

    if (line.trim() === "") {
      flushParagraph();
      // Una línea en blanco no cierra la lista si el siguiente renglón sigue en ella.
      const next = lines[i + 1] ?? "";
      if (!BULLET.test(next) && !NUMBERED.test(next) && !/^\s+\S/.test(next)) closeLists();
      continue;
    }

    if (lists.length && /^\s+\S/.test(line)) {
      out.push(`<br>${inline(line.trim())}`);
      continue;
    }

    closeLists();
    paragraph.push(line);
  }
  flushParagraph();
  closeLists();
  return out.join("");
}
