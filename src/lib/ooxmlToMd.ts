import { MERMAID_TITLE } from "./ooxmlImage";

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";

export interface ExportedDiagram {
  code: string;
  fileName: string;
}

export interface ExportedImage {
  fileName: string;
  base64: string;
}

export interface MdExport {
  markdown: string;
  diagrams: ExportedDiagram[];
  images: ExportedImage[];
}

// ---------- accesso al DOM XML ----------

function kids(el: Element, name: string): Element[] {
  return Array.from(el.children).filter((c) => c.localName === name);
}

function kid(el: Element | null | undefined, name: string): Element | null {
  return el ? kids(el, name)[0] ?? null : null;
}

function wAttr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(NS_W, name) ?? el.getAttribute("w:" + name);
}

function rAttr(el: Element, name: string): string | null {
  return el.getAttributeNS(NS_R, name) ?? el.getAttribute("r:" + name);
}

/** Proprietà on/off di Word: presente senza valore = acceso; "0"/"false"/"off" = spento. */
function isOn(el: Element | null | undefined): boolean {
  if (!el) return false;
  const v = wAttr(el, "val");
  return v === null || !/^(0|false|off)$/i.test(v);
}

function descendants(el: Element, name: string): Element[] {
  return Array.from(el.getElementsByTagNameNS("*", name));
}

// ---------- pacchetto Flat OPC ----------

interface Part {
  contentType: string;
  xml?: Element;
  base64?: string;
}

function readPackage(flat: string): Map<string, Part> {
  const doc = new DOMParser().parseFromString(flat, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("OOXML del documento non valido");
  const parts = new Map<string, Part>();
  for (const p of Array.from(doc.getElementsByTagNameNS(NS_PKG, "part"))) {
    const name = p.getAttributeNS(NS_PKG, "name") ?? p.getAttribute("pkg:name") ?? "";
    const contentType = p.getAttributeNS(NS_PKG, "contentType") ?? p.getAttribute("pkg:contentType") ?? "";
    const xmlData = kid(p, "xmlData");
    const bin = kid(p, "binaryData");
    parts.set(name, {
      contentType,
      xml: xmlData?.firstElementChild ?? undefined,
      base64: bin?.textContent?.replace(/\s+/g, ""),
    });
  }
  return parts;
}

function readFootnotes(pkg: Map<string, Part>): Map<string, Element> {
  const map = new Map<string, Element>();
  const root = pkg.get("/word/footnotes.xml")?.xml;
  if (root) for (const f of kids(root, "footnote")) map.set(wAttr(f, "id") ?? "", f);
  return map;
}

// ---------- stili e numerazione ----------

interface StyleInfo {
  name: string;
  basedOn: string | null;
  numId: string | null;
  ilvl: number;
}

function readStyles(pkg: Map<string, Part>): Map<string, StyleInfo> {
  const map = new Map<string, StyleInfo>();
  const root = pkg.get("/word/styles.xml")?.xml;
  if (!root) return map;
  for (const s of kids(root, "style")) {
    const id = wAttr(s, "styleId");
    if (!id) continue;
    const numPr = kid(kid(s, "pPr"), "numPr");
    map.set(id, {
      name: (wAttr(kid(s, "name"), "val") ?? id).toLowerCase(),
      basedOn: wAttr(kid(s, "basedOn"), "val"),
      numId: wAttr(kid(numPr, "numId"), "val"),
      ilvl: Number(wAttr(kid(numPr, "ilvl"), "val") ?? 0),
    });
  }
  return map;
}

/** Nomi (minuscoli) dello stile e di quelli da cui deriva, dal più specifico al più generico. */
function styleChain(styles: Map<string, StyleInfo>, id: string | null): string[] {
  const names: string[] = [];
  const seen = new Set<string>();
  while (id && !seen.has(id)) {
    seen.add(id);
    const s = styles.get(id);
    if (!s) break;
    names.push(s.name);
    id = s.basedOn;
  }
  return names;
}

class Numbering {
  private abstractOf = new Map<string, string>();
  private formats = new Map<string, string>(); // `${abstractId}:${ilvl}` -> numFmt
  private counters = new Map<string, number[]>();

  constructor(pkg: Map<string, Part>) {
    const root = pkg.get("/word/numbering.xml")?.xml;
    if (!root) return;
    for (const a of kids(root, "abstractNum")) {
      const aid = wAttr(a, "abstractNumId");
      for (const l of kids(a, "lvl")) {
        this.formats.set(`${aid}:${wAttr(l, "ilvl")}`, wAttr(kid(l, "numFmt"), "val") ?? "decimal");
      }
    }
    for (const n of kids(root, "num")) {
      this.abstractOf.set(wAttr(n, "numId") ?? "", wAttr(kid(n, "abstractNumId"), "val") ?? "");
    }
  }

  isBullet(numId: string, ilvl: number): boolean {
    const fmt = this.formats.get(`${this.abstractOf.get(numId)}:${ilvl}`);
    return fmt === undefined ? true : fmt === "bullet" || fmt === "none";
  }

  next(numId: string, ilvl: number): number {
    const c = this.counters.get(numId) ?? [];
    c[ilvl] = (c[ilvl] ?? 0) + 1;
    c.length = ilvl + 1;
    this.counters.set(numId, c);
    return c[ilvl];
  }
}

// ---------- contenuto in linea ----------

interface Seg {
  text: string;
  b: boolean;
  i: boolean;
  s: boolean;
  code: boolean;
  sub: boolean;
  sup: boolean;
  mark: boolean;
  u: boolean;
  href: string | null;
}

type Item = Seg | { diagram: string } | { image: string } | { pagebreak: true } | { footnote: string };

const MONO_FONT = /consolas|courier|monospace|menlo|lucida console|fira code|source code|cascadia/i;
const CODE_STYLE = /(^|\s)(html preformatted|html code|html typewriter|source code|code block|code|plain text|macro text)$/;

interface Ctx {
  pkg: Map<string, Part>;
  styles: Map<string, StyleInfo>;
  numbering: Numbering;
  rels: Map<string, { target: string; external: boolean }>;
  baseName: string;
  diagrams: ExportedDiagram[];
  images: ExportedImage[];
  /** Id Word delle note citate, nell'ordine di comparsa (l'indice + 1 è il numero nel Markdown). */
  footnoteRefs: string[];
  footnotes: Map<string, Element>;
}

function readRels(pkg: Map<string, Part>): Ctx["rels"] {
  const rels: Ctx["rels"] = new Map();
  const root = pkg.get("/word/_rels/document.xml.rels")?.xml;
  if (root) {
    for (const r of Array.from(root.children)) {
      rels.set(r.getAttribute("Id") ?? "", {
        target: r.getAttribute("Target") ?? "",
        external: r.getAttribute("TargetMode") === "External",
      });
    }
  }
  return rels;
}

/** Evidenziazione del testo (evidenziatore o ombreggiatura colorata del run). */
function isHighlighted(rPr: Element | null): boolean {
  const hl = wAttr(kid(rPr, "highlight"), "val");
  if (hl && hl !== "none") return true;
  const fill = wAttr(kid(rPr, "shd"), "fill");
  return !!fill && !/^(auto|FFFFFF)$/i.test(fill);
}

function collectItems(p: Element, ctx: Ctx): Item[] {
  const items: Item[] = [];
  let href: string | null = null;
  const fields: { instr: string; inResult: boolean; href: string | null }[] = [];

  const run = (r: Element, link: string | null): void => {
    const rPr = kid(r, "rPr");
    if (isOn(kid(rPr, "vanish"))) return;
    const styleChainNames = styleChain(ctx.styles, wAttr(kid(rPr, "rStyle"), "val"));
    const fonts = kid(rPr, "rFonts");
    const fontName = wAttr(fonts, "ascii") ?? wAttr(fonts, "hAnsi") ?? "";
    const fmt = {
      b: isOn(kid(rPr, "b")) || styleChainNames.includes("strong"),
      i: isOn(kid(rPr, "i")) || styleChainNames.includes("emphasis"),
      s: isOn(kid(rPr, "strike")) || isOn(kid(rPr, "dstrike")),
      code: MONO_FONT.test(fontName) || styleChainNames.some((n) => CODE_STYLE.test(n)),
      sub: wAttr(kid(rPr, "vertAlign"), "val") === "subscript",
      sup: wAttr(kid(rPr, "vertAlign"), "val") === "superscript",
      mark: isHighlighted(rPr),
      u: !!kid(rPr, "u") && wAttr(kid(rPr, "u"), "val") !== "none",
      href: link,
    };
    if (fmt.code) fmt.mark = false;
    const push = (text: string) => items.push({ text, ...fmt });

    for (const c of Array.from(r.children)) {
      switch (c.localName) {
        case "t":
          push(c.textContent ?? "");
          break;
        case "tab":
          push(" ");
          break;
        case "noBreakHyphen":
          push("-");
          break;
        case "br":
          if (wAttr(c, "type") === "page") items.push({ pagebreak: true });
          else if (wAttr(c, "type") !== "column") push("\n");
          break;
        case "instrText": {
          const f = fields[fields.length - 1];
          if (f) f.instr += c.textContent ?? "";
          break;
        }
        case "fldChar": {
          const type = wAttr(c, "fldCharType");
          if (type === "begin") fields.push({ instr: "", inResult: false, href: null });
          else if (type === "separate") {
            const f = fields[fields.length - 1];
            if (f) {
              f.inResult = true;
              f.href = /HYPERLINK\s+"([^"]+)"/i.exec(f.instr)?.[1] ?? null;
              if (f.href) href = f.href;
            }
          } else if (type === "end") {
            const f = fields.pop();
            if (f?.href) href = null;
          }
          break;
        }
        case "drawing":
          drawing(c);
          break;
        case "footnoteReference":
          items.push({ footnote: wAttr(c, "id") ?? "" });
          break;
      }
    }
  };

  const drawing = (d: Element): void => {
    const docPr = descendants(d, "docPr")[0];
    const title = docPr?.getAttribute("title") ?? "";
    const descr = docPr?.getAttribute("descr") ?? "";
    if (title === MERMAID_TITLE) {
      items.push({ diagram: descr });
      return;
    }
    const blip = descendants(d, "blip")[0];
    const rid = blip ? rAttr(blip, "embed") : null;
    const rel = rid ? ctx.rels.get(rid) : undefined;
    const part = rel ? ctx.pkg.get("/word/" + rel.target.replace(/^\/?(word\/)?/, "")) : undefined;
    if (!part?.base64) return;
    const ext = part.contentType.split("/")[1]?.replace("jpeg", "jpg").replace(/\+.*$/, "") ?? "png";
    const fileName = `${ctx.baseName}-immagine-${ctx.images.length + 1}.${ext}`;
    ctx.images.push({ fileName, base64: part.base64 });
    const alt = (descr || docPr?.getAttribute("name") || "").replace(/[\[\]\n]/g, " ").trim();
    items.push({ image: `![${alt}](${encodeURI(fileName)})` });
  };

  const walk = (el: Element, link: string | null): void => {
    for (const c of Array.from(el.children)) {
      switch (c.localName) {
        case "r":
          run(c, link ?? href);
          break;
        case "hyperlink": {
          const rid = rAttr(c, "id");
          const anchor = wAttr(c, "anchor");
          const target = rid ? ctx.rels.get(rid)?.target : anchor ? "#" + anchor.replace(/^h_/, "").replace(/_/g, "-") : null;
          walk(c, target ?? link);
          break;
        }
        case "fldSimple": {
          const url = /HYPERLINK\s+"([^"]+)"/i.exec(wAttr(c, "instr") ?? "")?.[1];
          walk(c, url ?? link);
          break;
        }
        case "ins":
        case "smartTag":
        case "sdtContent":
        case "sdt":
        case "customXml":
          walk(c, link);
          break;
      }
    }
  };

  walk(p, null);
  return items;
}

// ---------- rendering Markdown ----------

function escapeText(t: string): string {
  return t.replace(/([\\`*_\[\]<>])/g, "\\$1");
}

function codeSpan(text: string): string {
  const longest = Math.max(0, ...(text.match(/`+/g) ?? []).map((m) => m.length));
  const fence = "`".repeat(longest + 1);
  const pad = text.startsWith("`") || text.endsWith("`") ? " " : "";
  return fence + pad + text + pad + fence;
}

/** Applica i marcatori lasciando fuori gli spazi iniziali e finali (richiesto da Markdown). */
function wrapSegment(seg: Seg): string {
  const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(seg.text)!;
  if (!m[2]) return seg.text;
  let inner = seg.code ? codeSpan(m[2]) : escapeText(m[2]);
  if (seg.mark) inner = `<mark>${inner}</mark>`;
  if (seg.sub) inner = `<sub>${inner}</sub>`;
  if (seg.sup) inner = `<sup>${inner}</sup>`;
  if (seg.u && !seg.href) inner = `<u>${inner}</u>`;
  if (seg.s) inner = `~~${inner}~~`;
  if (seg.b && seg.i) inner = `***${inner}***`;
  else if (seg.b) inner = `**${inner}**`;
  else if (seg.i) inner = `*${inner}*`;
  return m[1] + inner + m[3];
}

function sameFormat(a: Seg, b: Seg): boolean {
  return a.b === b.b && a.i === b.i && a.s === b.s && a.code === b.code && a.href === b.href &&
    a.sub === b.sub && a.sup === b.sup && a.mark === b.mark && a.u === b.u;
}

function renderSegments(segs: Seg[], lineBreak: string): string {
  const merged: Seg[] = [];
  for (const s of segs) {
    const last = merged[merged.length - 1];
    if (last && sameFormat(last, s)) last.text += s.text;
    else merged.push({ ...s });
  }
  let out = "";
  for (let i = 0; i < merged.length; ) {
    const href = merged[i].href;
    let j = i;
    let group = "";
    while (j < merged.length && merged[j].href === href) group += wrapSegment(merged[j++]);
    group = group.replace(/\n/g, lineBreak);
    out += href && group.trim() ? `[${group}](${href.replace(/ /g, "%20").replace(/\)/g, "%29")})` : group;
    i = j;
  }
  return out;
}

/** Evita che un paragrafo normale venga riletto come titolo, elenco o citazione. */
function protectLineStart(text: string): string {
  return text.replace(/^(\s*)(#{1,6}(?=\s)|[-+](?=\s)|\d+(?=[.)]\s)|(?=>))/, "$1\\$2");
}

// ---------- blocchi ----------

type Kind = "para" | "heading" | "list" | "quote" | "code" | "table" | "mermaid" | "footnote";
interface Block {
  kind: Kind;
  text: string;
}

function paragraphBlocks(p: Element, ctx: Ctx): Block[] {
  const pPr = kid(p, "pPr");
  const styleId = wAttr(kid(pPr, "pStyle"), "val");
  const chain = styleChain(ctx.styles, styleId);
  const items = collectItems(p, ctx);
  const blocks: Block[] = [];

  let pending: Item[] = [];
  const flush = (): void => {
    if (pending.length) blocks.push(...textBlock(pending, p, chain, ctx));
    pending = [];
  };
  for (const it of items) {
    if ("pagebreak" in it) {
      flush();
      blocks.push({ kind: "para", text: "<!-- pagebreak -->" });
    } else if ("diagram" in it) {
      flush();
      const n = ctx.diagrams.length + 1;
      const fileName = `${ctx.baseName}-diagramma-${n}.svg`;
      ctx.diagrams.push({ code: it.diagram, fileName });
      blocks.push({ kind: "mermaid", text: "```mermaid\n" + it.diagram.replace(/\n+$/, "") + "\n```" });
    } else pending.push(it);
  }
  flush();
  return blocks;
}

function textBlock(items: Item[], p: Element, chain: string[], ctx: Ctx): Block[] {
  const pPr = kid(p, "pPr");
  const segs = items.filter((i): i is Seg => "text" in i);
  const plain = segs.map((s) => s.text).join("");
  const hasImage = items.some((i) => "image" in i);

  if (!plain.trim() && !hasImage) return [];

  const textSegs = segs.filter((s) => s.text.trim());
  const isCode =
    !hasImage &&
    (chain.some((n) => CODE_STYLE.test(n)) || (textSegs.length > 0 && textSegs.every((s) => s.code)));
  if (isCode) return [{ kind: "code", text: plain.replace(/\n$/, "") }];

  const text = renderItems(items, ctx);

  const heading = chain.map((n) => /(?:^|\s)heading (\d)$/.exec(n)?.[1]).find(Boolean);
  if (heading) return [{ kind: "heading", text: "#".repeat(Math.min(6, Number(heading))) + " " + text.replace(/\\\n/g, " ") }];
  if (chain.some((n) => /(^|\s)title$/.test(n))) return [{ kind: "heading", text: "# " + text.replace(/\\\n/g, " ") }];

  if (chain.some((n) => /(^|\s)(quote|intense quote|block text)$/.test(n))) {
    return [{ kind: "quote", text: text.split("\n").map((l) => "> " + l).join("\n") }];
  }

  let list = listInfo(pPr, chain, ctx);
  if (!list && /^[☐☑] /.test(plain)) {
    // voce con casella (elenco di attività): nessuna numerazione, il livello viene dal rientro
    const left = Number(wAttr(kid(pPr, "ind"), "left") ?? 0);
    list = { numId: "task", ilvl: Math.max(0, Math.round(left / 360) - 1), bullet: true };
  }
  if (list) {
    const marker = list.bullet ? "-" : `${ctx.numbering.next(list.numId, list.ilvl)}.`;
    const task = text.replace(/^☐ /, "[ ] ").replace(/^☑ /, "[x] ");
    return [{ kind: "list", text: "    ".repeat(list.ilvl) + marker + " " + task.replace(/\\\n/g, " ") }];
  }
  return [{ kind: "para", text: protectLineStart(text) }];
}

/** Unisce segmenti consecutivi e rende anche le immagini in linea. */
function renderItems(items: Item[], ctx: Ctx): string {
  let out = "";
  let run: Seg[] = [];
  const flush = (): void => {
    if (run.length) out += renderSegments(run, "\\\n");
    run = [];
  };
  for (const it of items) {
    if ("text" in it) run.push(it);
    else if ("image" in it) {
      flush();
      out += it.image;
    } else if ("footnote" in it) {
      flush();
      let n = ctx.footnoteRefs.indexOf(it.footnote) + 1;
      if (!n) n = ctx.footnoteRefs.push(it.footnote);
      out += `[^${n}]`;
    }
  }
  flush();
  return out;
}

function listInfo(
  pPr: Element | null,
  chain: string[],
  ctx: Ctx,
): { numId: string; ilvl: number; bullet: boolean } | null {
  const numPr = kid(pPr, "numPr");
  let numId = wAttr(kid(numPr, "numId"), "val");
  let ilvl = numPr ? Number(wAttr(kid(numPr, "ilvl"), "val") ?? 0) : 0;
  if (!numId) {
    const styleId = wAttr(kid(pPr, "pStyle"), "val");
    for (let id = styleId, guard = 0; id && guard < 10; guard++) {
      const s = ctx.styles.get(id);
      if (s?.numId) {
        numId = s.numId;
        ilvl = numPr ? ilvl : s.ilvl;
        break;
      }
      id = s?.basedOn ?? null;
    }
  }
  if (numId === "0") return null; // numId 0 = "nessun elenco"
  if (numId) return { numId, ilvl, bullet: ctx.numbering.isBullet(numId, ilvl) };

  const named = chain.map((n) => /^list (bullet|number)(?: (\d))?$/.exec(n)).find(Boolean);
  if (named) {
    const level = Math.max(0, Number(named[2] ?? 1) - 1);
    return { numId: "style:" + named[1], ilvl: level, bullet: named[1] === "bullet" };
  }
  return null;
}

function tableBlock(tbl: Element, ctx: Ctx): Block | null {
  const rows: string[][] = [];
  const aligns: string[] = [];
  for (const tr of kids(tbl, "tr")) {
    const row: string[] = [];
    for (const tc of kids(tr, "tc")) {
      const cellParas = kids(tc, "p").map((p) => {
        const items = collectItems(p, ctx)
          .filter((i) => !("diagram" in i) && !("pagebreak" in i))
          .map((i) => (rows.length === 0 && "text" in i ? { ...i, b: false } : i));
        return renderItems(items, ctx).replace(/\\\n/g, "<br>").replace(/\|/g, "\\|").trim();
      });
      row.push(cellParas.filter(Boolean).join("<br>"));
      if (rows.length === 0) {
        const jc = wAttr(kid(kid(kids(tc, "p")[0] ?? null, "pPr"), "jc"), "val");
        aligns.push(jc === "center" ? ":---:" : jc === "right" || jc === "end" ? "---:" : "---");
      }
      const span = Number(wAttr(kid(kid(tc, "tcPr"), "gridSpan"), "val") ?? 1);
      for (let k = 1; k < span; k++) {
        row.push("");
        if (rows.length === 0) aligns.push("---");
      }
    }
    rows.push(row);
  }
  if (!rows.length) return null;
  const cols = Math.max(...rows.map((r) => r.length));
  const pad = (r: string[]) => Array.from({ length: cols }, (_, i) => r[i] ?? "");
  const line = (r: string[]) => "| " + pad(r).join(" | ") + " |";
  const sep = pad(aligns.length ? aligns : []).map((a) => a || "---");
  return {
    kind: "table",
    text: [line(rows[0]), "| " + sep.join(" | ") + " |", ...rows.slice(1).map(line)].join("\n"),
  };
}

function bodyBlocks(parent: Element, ctx: Ctx, out: Block[]): void {
  for (const el of Array.from(parent.children)) {
    switch (el.localName) {
      case "p":
        out.push(...paragraphBlocks(el, ctx));
        break;
      case "tbl": {
        const t = tableBlock(el, ctx);
        if (t) out.push(t);
        break;
      }
      case "sdt":
        {
          const content = kid(el, "sdtContent");
          if (content) bodyBlocks(content, ctx, out);
        }
        break;
    }
  }
}

/** Unisce righe di codice o di citazione consecutive e separa i blocchi con una riga vuota. */
function joinBlocks(blocks: Block[]): string {
  const merged: Block[] = [];
  for (const b of blocks) {
    const last = merged[merged.length - 1];
    if (last && b.kind === "code" && last.kind === "code") last.text += "\n" + b.text;
    else if (last && b.kind === "quote" && last.kind === "quote") last.text += "\n>\n" + b.text;
    else merged.push({ ...b });
  }
  let out = "";
  merged.forEach((b, i) => {
    const text = b.kind === "code" ? "```\n" + b.text + "\n```" : b.text;
    if (i > 0) out += b.kind === "list" && merged[i - 1].kind === "list" ? "\n" : "\n\n";
    out += text;
  });
  return out ? out + "\n" : "";
}

/** Converte l'OOXML (Flat OPC) restituito da body.getOoxml() in Markdown. */
export function ooxmlToMarkdown(flatOpc: string, baseName: string): MdExport {
  const pkg = readPackage(flatOpc);
  const root = pkg.get("/word/document.xml")?.xml;
  const body = root ? kid(root, "body") : null;
  if (!body) throw new Error("Il documento non contiene un corpo (word/document.xml)");

  const ctx: Ctx = {
    pkg,
    styles: readStyles(pkg),
    numbering: new Numbering(pkg),
    rels: readRels(pkg),
    baseName,
    diagrams: [],
    images: [],
    footnoteRefs: [],
    footnotes: readFootnotes(pkg),
  };
  const blocks: Block[] = [];
  bodyBlocks(body, ctx, blocks);
  for (let i = 0; i < ctx.footnoteRefs.length; i++) {
    const note = ctx.footnotes.get(ctx.footnoteRefs[i]);
    const text = note
      ? kids(note, "p").map((p) => renderItems(collectItems(p, ctx), ctx).replace(/\\\n/g, " ").trim()).filter(Boolean).join(" ")
      : "";
    blocks.push({ kind: "footnote", text: `[^${i + 1}]: ${text}` });
  }
  return { markdown: joinBlocks(blocks), diagrams: ctx.diagrams, images: ctx.images };
}
