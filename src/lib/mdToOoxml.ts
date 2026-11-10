import MarkdownIt, { type Token } from "markdown-it";
import footnote from "markdown-it-footnote";
import deflist from "markdown-it-deflist";
import sub from "markdown-it-sub";
import sup from "markdown-it-sup";
import mark from "markdown-it-mark";
import type { ImageRes } from "./imageInfo";
import type { RenderedDiagram } from "./mermaidRender";
import { diagramParagraph, diagramRun } from "./ooxmlImage";
import { PackageBuilder, escapeAttr, escapeXml, inlineDrawing, utf8ToBase64 } from "./opc";
import { ELEMENTS, type Align, type Border, type ElementFmt, type Profile } from "./profile";


const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const JC: Record<Align, string> = { left: "left", center: "center", right: "right", justify: "both" };
const TEXT_PROPS = ["font", "sizePt", "color", "bold", "italic", "underline", "strike", "background"] as const;

export interface Prepared {
  tokens: Token[];
  /** Codice dei blocchi ```mermaid, nell'ordine in cui compaiono. */
  mermaid: string[];
  /** Indirizzi (src) delle immagini da risolvere, senza duplicati. */
  imageSrcs: string[];
}

function newParser() {
  return new MarkdownIt({ html: true, linkify: true, typographer: false })
    .use(footnote).use(deflist).use(sub).use(sup).use(mark);
}

const at = (t: Token, name: string): string => {
  const v = t.attrGet(name);
  return v === null || v === undefined ? "" : String(v);
};

export function isMermaidFence(t: Token): boolean {
  const lang = t.info.trim().split(/\s+/)[0].toLowerCase();
  return t.type === "fence" && (lang === "mermaid" || lang === "mmd");
}

/** Il front matter YAML iniziale diventa un blocco di codice invece di essere letto come linea e titolo. */
function protectFrontMatter(md: string): string {
  return md.replace(/^---\r?\n((?:[A-Za-z_][\w-]*:[^\n]*\n(?:[ \t]+[^\n]*\n|-[^\n]*\n)*)+)---\r?\n/, "```yaml\n$1```\n");
}

export function prepare(markdown: string): Prepared {
  const tokens = newParser().parse(protectFrontMatter(markdown.replace(/^﻿/, "")), {});
  const mermaid: string[] = [];
  const srcs = new Set<string>();
  const scan = (list: Token[]) => {
    for (const t of list) {
      if (isMermaidFence(t)) mermaid.push(t.content.replace(/\n$/, ""));
      if (t.type === "image") srcs.add(at(t, "src"));
      if (t.children) scan(t.children);
    }
  };
  scan(tokens);
  srcs.delete("");
  return { tokens, mermaid, imageSrcs: [...srcs] };
}

// ---------- XML delle proprietà ----------

const eighths = (pt: number) => Math.max(2, Math.round(pt * 8));
const twips = (pt: number) => Math.round(pt * 20);

function borderXml(tag: string, b: Border, space: number): string {
  return `<w:${tag} w:val="${b.style ?? "single"}" w:sz="${eighths(b.widthPt)}" w:space="${space}" w:color="${b.color}"/>`;
}

function pPrInner(f: ElementFmt): string {
  let x = "";
  if (f.keepNext) x += "<w:keepNext/>";
  if (f.keepLines) x += "<w:keepLines/>";
  if (f.pageBreakBefore) x += "<w:pageBreakBefore/>";
  if (f.borderTop || f.borderLeft || f.borderBottom) {
    x += "<w:pBdr>" +
      (f.borderTop ? borderXml("top", f.borderTop, 1) : "") +
      (f.borderLeft ? borderXml("left", f.borderLeft, 6) : "") +
      (f.borderBottom ? borderXml("bottom", f.borderBottom, 1) : "") + "</w:pBdr>";
  }
  if (f.background) x += `<w:shd w:val="clear" w:color="auto" w:fill="${f.background}"/>`;
  if (f.beforePt !== undefined || f.afterPt !== undefined || f.lineSpacing !== undefined) {
    x += "<w:spacing" +
      (f.beforePt !== undefined ? ` w:before="${twips(f.beforePt)}"` : "") +
      (f.afterPt !== undefined ? ` w:after="${twips(f.afterPt)}"` : "") +
      (f.lineSpacing !== undefined ? ` w:line="${Math.round(f.lineSpacing * 240)}" w:lineRule="auto"` : "") + "/>";
  }
  if (f.indentPt) x += `<w:ind w:left="${twips(f.indentPt)}"/>`;
  if (f.align) x += `<w:jc w:val="${JC[f.align]}"/>`;
  return x;
}

function rPrInner(f: ElementFmt, rStyle?: string, vertAlign?: string): string {
  let x = "";
  if (rStyle) x += `<w:rStyle w:val="${rStyle}"/>`;
  if (f.font) x += `<w:rFonts w:ascii="${escapeAttr(f.font)}" w:hAnsi="${escapeAttr(f.font)}" w:cs="${escapeAttr(f.font)}" w:eastAsia="${escapeAttr(f.font)}"/>`;
  if (f.bold !== undefined) x += f.bold ? "<w:b/><w:bCs/>" : '<w:b w:val="0"/><w:bCs w:val="0"/>';
  if (f.italic !== undefined) x += f.italic ? "<w:i/><w:iCs/>" : '<w:i w:val="0"/><w:iCs w:val="0"/>';
  if (f.strike) x += "<w:strike/>";
  if (f.color) x += `<w:color w:val="${f.color}"/>`;
  if (f.sizePt) x += `<w:sz w:val="${Math.round(f.sizePt * 2)}"/><w:szCs w:val="${Math.round(f.sizePt * 2)}"/>`;
  if (f.underline) x += '<w:u w:val="single"/>';
  else if (f.underline === false) x += '<w:u w:val="none"/>';
  if (f.background) x += `<w:shd w:val="clear" w:color="auto" w:fill="${f.background}"/>`;
  if (vertAlign) x += `<w:vertAlign w:val="${vertAlign}"/>`;
  return x;
}

const wrapTag = (tag: string, inner: string) => (inner ? `<w:${tag}>${inner}</w:${tag}>` : "");

/** Nome canonico → styleId ("heading 1" → "Heading1", "footnote text" → "FootnoteText"). */
export function styleIdFor(name: string): string {
  return name.replace(/[^A-Za-z0-9 ]/g, "").split(" ").filter(Boolean).map((w) => w[0].toUpperCase() + w.slice(1)).join("") || "Normal";
}

const safe = (s: string) => s.replace(/[^A-Za-z0-9]/g, "");

/** Nome con cui uno stile viene creato nel documento in modalità "nuovo stile". */
export function newStyleName(profileName: string, en: string): string {
  return `WordMD ${profileName} - ${en}`;
}

/** Nomi degli stili che l'importazione creerà (da cancellare prima di un nuovo import, per aggiornarli). */
export function createdStyleNames(profile: Profile): string[] {
  if (profile.mode !== "new") return [];
  return ELEMENTS.filter((e) => e.kind === "p").map((e) => newStyleName(profile.name, e.en));
}

// ---------- costruttore ----------

interface InlineState {
  strong: boolean; em: boolean; del: boolean; mark: boolean; sub: boolean; sup: boolean; u: boolean; kbd: boolean;
  link: { rid?: string; anchor?: string } | null;
  /** Formattazione aggiuntiva del contesto (intestazione di tabella, ecc.). */
  base?: ElementFmt;
}

const blankState = (): InlineState => ({
  strong: false, em: false, del: false, mark: false, sub: false, sup: false, u: false, kbd: false, link: null,
});

interface Cell { inline: Token | null; align: Align | null; header: boolean }

class Builder {
  private pkg = new PackageBuilder();
  private body: string[] = [];
  private styles = new Map<string, string>();
  private nums: string[] = [];
  private nextNumId = 2;
  private footnotes: string[] = [];
  private bookmarkId = 1;
  private slugs = new Map<string, number>();
  private diagramIdx = 0;
  private quoteDepth = 0;
  private lists: { ordered: boolean; numId: number }[] = [];
  private inDefinition = false;

  constructor(
    private p: Profile,
    private rendered: (RenderedDiagram | null)[],
    private images: Map<string, ImageRes>,
  ) {}

  // ----- stili -----

  private el(key: string): ElementFmt {
    return this.p.elements[key] ?? {};
  }

  /** styleId dello stile paragrafo di un elemento; lo registra per includerlo nel pacchetto. */
  private style(key: string): string {
    const def = ELEMENTS.find((e) => e.key === key)!;
    const fmt = this.el(key);
    if (this.p.mode === "existing") {
      const name = fmt.wordStyle || def.builtIn || "Normal";
      const id = styleIdFor(name);
      if (!this.styles.has(id)) this.styles.set(id, `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${escapeAttr(name)}"/></w:style>`);
      return id;
    }
    const id = `WordMD${safe(this.p.name)}${safe(def.en)}`;
    if (!this.styles.has(id)) {
      const heading = /^h([1-6])$/.exec(key);
      const outline = heading ? `<w:outlineLvl w:val="${Number(heading[1]) - 1}"/>` : "";
      this.styles.set(id,
        `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${id}"><w:name w:val="${escapeAttr(newStyleName(this.p.name, def.en))}"/>` +
        `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
        wrapTag("pPr", pPrInner(fmt) + outline) + wrapTag("rPr", rPrInner(fmt)) + "</w:style>");
    }
    return id;
  }

  private charStyle(name: string): string {
    const id = styleIdFor(name);
    if (!this.styles.has(id)) this.styles.set(id, `<w:style w:type="character" w:styleId="${id}"><w:name w:val="${escapeAttr(name)}"/></w:style>`);
    return id;
  }

  private rPr(s: InlineState, extra?: ElementFmt): string {
    const f: ElementFmt = {};
    let rStyle: string | undefined;
    const apply = (key: string) => {
      const e = this.el(key);
      if (this.p.mode === "existing" && e.wordStyle) rStyle ??= this.charStyle(e.wordStyle);
      else for (const k of TEXT_PROPS) if (e[k] !== undefined) (f as Record<string, unknown>)[k] = e[k];
    };
    if (s.base) for (const k of TEXT_PROPS) if (s.base[k] !== undefined) (f as Record<string, unknown>)[k] = s.base[k];
    if (s.strong) apply("strong");
    if (s.em) apply("em");
    if (s.del) apply("del");
    if (s.mark) apply("mark");
    if (s.kbd) apply("codeInline");
    if (s.link) { apply("link"); if (this.p.mode === "new" && f.underline === undefined) f.underline = false; }
    if (s.u) f.underline = true;
    if (extra) Object.assign(f, extra);
    return wrapTag("rPr", rPrInner(f, rStyle, s.sub ? "subscript" : s.sup ? "superscript" : undefined));
  }

  // ----- paragrafi -----

  private para(key: string, inner: string, o: { numPr?: string; ind?: number; hanging?: number; jc?: Align; keepNext?: boolean; border?: string } = {}): string {
    let pPr = `<w:pStyle w:val="${this.style(key)}"/>`;
    if (o.keepNext) pPr += "<w:keepNext/>";
    if (o.numPr) pPr += o.numPr;
    if (o.border) pPr += o.border;
    if (o.ind !== undefined) pPr += `<w:ind w:left="${o.ind}"${o.hanging ? ` w:hanging="${o.hanging}"` : ""}/>`;
    if (o.jc) pPr += `<w:jc w:val="${JC[o.jc]}"/>`;
    return `<w:p><w:pPr>${pPr}</w:pPr>${inner}</w:p>`;
  }

  private text(t: string, s: InlineState, extra?: ElementFmt): string {
    if (!t) return "";
    return `<w:r>${this.rPr(s, extra)}<w:t xml:space="preserve">${escapeXml(t)}</w:t></w:r>`;
  }

  private slug(text: string): string {
    let s = text.toLowerCase().trim().replace(/[^\p{L}\p{N}\s_-]/gu, "").replace(/\s/g, "-");
    const n = this.slugs.get(s) ?? 0;
    this.slugs.set(s, n + 1);
    if (n) s += `-${n}`;
    return s;
  }

  private bookmarkName(slug: string): string {
    return ("h_" + slug.replace(/[^\p{L}\p{N}_]/gu, "_")).slice(0, 40);
  }

  // ----- contenuto in linea -----

  private image(t: Token, s: InlineState): string {
    const src = at(t, "src");
    const alt = (t.children ?? []).map((c) => c.content).join("") || t.content;
    const res = this.images.get(src);
    const png = res ? (res.contentType === "image/svg+xml" ? res.pngBase64 : res.base64) : undefined;
    if (!res || !png) {
      return this.text(`[immagine non disponibile: ${alt || src}]`, s, { italic: true, color: "888888" });
    }
    const pngRid = this.pkg.addMedia(png, res.contentType === "image/svg+xml" ? "image/png" : res.contentType, res.contentType === "image/svg+xml" ? "png" : res.ext);
    const svgRid = res.svgText ? this.pkg.addMedia(utf8ToBase64(res.svgText), "image/svg+xml", "svg") : undefined;
    const name = src.startsWith("data:") ? "immagine" : decodeURIComponent(src.split(/[?#]/)[0].split("/").pop() || "immagine");
    const { xml } = inlineDrawing({
      pngRid, svgRid, widthPx: res.widthPx, heightPx: res.heightPx, maxWidthCm: this.p.image.maxWidthCm,
      id: this.pkg.newDocPrId(), name, descr: alt || at(t, "title") || "",
    });
    return `<w:r>${xml}</w:r>`;
  }

  private htmlInline(raw: string, s: InlineState): string {
    const m = /^<\s*(\/)?\s*([a-zA-Z0-9]+)/.exec(raw);
    if (!m) return "";
    const close = !!m[1];
    switch (m[2].toLowerCase()) {
      case "br": return "<w:r><w:br/></w:r>";
      case "sub": s.sub = !close; break;
      case "sup": s.sup = !close; break;
      case "u": case "ins": s.u = !close; break;
      case "kbd": case "code": s.kbd = !close; break;
      case "b": case "strong": s.strong = !close; break;
      case "i": case "em": s.em = !close; break;
      case "s": case "del": case "strike": s.del = !close; break;
      case "mark": s.mark = !close; break;
    }
    return "";
  }

  private inline(tok: Token | null, base?: ElementFmt): string {
    if (!tok?.children) return "";
    const s = blankState();
    s.base = base;
    let out = "";
    const linkStack: InlineState["link"][] = [];
    for (const c of tok.children) {
      switch (c.type) {
        case "text": out += this.text(c.content, s); break;
        case "code_inline": out += this.text(c.content, { ...s, kbd: true }); break;
        case "softbreak": out += this.text(" ", s); break;
        case "hardbreak": out += "<w:r><w:br/></w:r>"; break;
        case "em_open": s.em = true; break;
        case "em_close": s.em = false; break;
        case "strong_open": s.strong = true; break;
        case "strong_close": s.strong = false; break;
        case "s_open": s.del = true; break;
        case "s_close": s.del = false; break;
        case "mark_open": s.mark = true; break;
        case "mark_close": s.mark = false; break;
        case "sub_open": s.sub = true; break;
        case "sub_close": s.sub = false; break;
        case "sup_open": s.sup = true; break;
        case "sup_close": s.sup = false; break;
        case "html_inline": out += this.htmlInline(c.content, s); break;
        case "image": out += this.image(c, s); break;
        case "footnote_ref": out += `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${Number(c.meta?.id ?? 0) + 1}"/></w:r>`; break;
        case "link_open": {
          const href = at(c, "href");
          linkStack.push(s.link);
          s.link = href.startsWith("#") ? { anchor: this.bookmarkName(decodeURIComponent(href.slice(1))) } : { rid: this.pkg.addHyperlink(href) };
          out += `\u0000LINK:${JSON.stringify(s.link)}\u0000`;
          break;
        }
        case "link_close":
          out += "\u0000ENDLINK\u0000";
          s.link = linkStack.pop() ?? null;
          break;
      }
    }
    return this.closeLinks(out);
  }

  /** Racchiude le sequenze di run tra i marcatori di collegamento in <w:hyperlink>. */
  private closeLinks(out: string): string {
    return out
      .replace(/\u0000LINK:(\{.*?\})\u0000/g, (_, json: string) => {
        const l = JSON.parse(json) as { rid?: string; anchor?: string };
        return l.rid ? `<w:hyperlink r:id="${l.rid}" w:history="1">` : `<w:hyperlink w:anchor="${l.anchor}" w:history="1">`;
      })
      .replace(/\u0000ENDLINK\u0000/g, "</w:hyperlink>");
  }

  private plain(tok: Token | null): string {
    return (tok?.children ?? []).map((c) => (c.type === "text" || c.type === "code_inline" ? c.content : "")).join("");
  }

  // ----- blocchi -----

  private add(xml: string): void {
    this.body.push(xml);
  }

  private imageOnly(tok: Token | null): Token | null {
    const kids = (tok?.children ?? []).filter((c) => !(c.type === "text" && !c.content.trim()) && c.type !== "softbreak");
    const imgs = kids.filter((c) => c.type === "image");
    const others = kids.filter((c) => c.type !== "image" && c.type !== "link_open" && c.type !== "link_close");
    return imgs.length === 1 && others.length === 0 ? imgs[0] : null;
  }

  private paragraph(tok: Token | null, listItemStart: boolean): void {
    const list = this.lists[this.lists.length - 1];
    const lvl = Math.min(8, this.lists.length - 1);

    // un'immagine da sola in un paragrafo: centrata, con didascalia dal testo alternativo
    const img = !list ? this.imageOnly(tok) : null;
    if (img) {
      const alt = (img.children ?? []).map((c) => c.content).join("") || img.content;
      const caption = this.p.image.captionFromAlt && alt;
      this.add(this.para("p", this.inline(tok), { jc: this.p.image.align, keepNext: !!caption && this.p.image.keepWithCaption }));
      if (caption) this.add(this.para("caption", this.text(alt, blankState())));
      return;
    }

    if (list) {
      let inner = this.inline(tok);
      let task: string | null = null;
      const first = tok?.children?.[0];
      const m = first?.type === "text" ? /^\[([ xX])\]\s+/.exec(first.content) : null;
      if (listItemStart && m && first) {
        task = m[1] === " " ? "☐ " : "☑ ";
        first.content = first.content.slice(m[0].length);
        inner = this.text(task, blankState()) + this.inline(tok);
      }
      const key = list.ordered ? "olist" : "ulist";
      if (listItemStart && !task) {
        this.add(this.para(key, inner, { numPr: `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${list.numId}"/></w:numPr>` }));
      } else if (listItemStart) {
        this.add(this.para(key, inner, { numPr: `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>`, ind: 360 * (lvl + 1), hanging: 360 }));
      } else {
        this.add(this.para(key, inner, { numPr: `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>`, ind: 360 * (lvl + 1) }));
      }
      return;
    }

    if (this.inDefinition) return void this.add(this.para("dd", this.inline(tok)));
    if (this.quoteDepth > 0) {
      const extra = this.quoteDepth > 1 ? { ind: twips((this.el("quote").indentPt ?? 0) + 18 * (this.quoteDepth - 1)) } : {};
      return void this.add(this.para("quote", this.inline(tok), extra));
    }
    this.add(this.para("p", this.inline(tok)));
  }

  private heading(level: number, tok: Token | null): void {
    const key = `h${level}`;
    const slug = this.slug(this.plain(tok));
    const id = this.bookmarkId++;
    const name = this.bookmarkName(slug);
    this.add(this.para(key, `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${this.inline(tok)}<w:bookmarkEnd w:id="${id}"/>`));
  }

  private code(content: string): void {
    const lines = content.replace(/\n$/, "").split("\n");
    const runs = lines.map((l, i) => (i ? "<w:r><w:br/></w:r>" : "") + this.text(l, blankState()));
    this.add(this.para("codeBlock", runs.join("")));
  }

  private diagram(): void {
    const code = this.currentMermaid[this.diagramIdx];
    const d = this.rendered[this.diagramIdx++];
    if (!d) {
      this.add(this.para("codeBlock", this.text("[Diagramma Mermaid non valido]", blankState(), { bold: true }) + "<w:r><w:br/></w:r>" +
        code.split("\n").map((l, i) => (i ? "<w:r><w:br/></w:r>" : "") + this.text(l, blankState())).join("")));
      return;
    }
    const m = this.p.mermaid;
    const run = diagramRun(this.pkg, code, d, `Diagramma ${this.diagramIdx}`, m.maxWidthCm);
    this.add(diagramParagraph(run, m.align, m.keepWithNext));
  }

  private hr(): void {
    const border = this.p.mode === "existing" ? '<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="A6A6A6"/></w:pBdr>' : undefined;
    this.add(this.para("hr", "", { border }));
  }

  private htmlBlock(raw: string): void {
    if (/page-?break|\\newpage|break-(?:after|before)\s*:\s*page/i.test(raw)) {
      this.add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
      return;
    }
    const text = raw.replace(/<!--[\s\S]*?-->/g, "").replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, "").trim();
    if (text) this.add(this.para("p", this.text(text, blankState())));
  }

  private startList(ordered: boolean, start: number): void {
    let numId = 1;
    if (ordered) {
      numId = this.nextNumId++;
      this.nums.push(`<w:num w:numId="${numId}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="${start}"/></w:lvlOverride></w:num>`);
    }
    this.lists.push({ ordered, numId });
  }

  // ----- tabelle -----

  private table(rows: Cell[][]): void {
    const t = this.p.table;
    const cols = Math.max(...rows.map((r) => r.length));
    const weights = Array.from({ length: cols }, (_, c) =>
      Math.min(40, Math.max(6, ...rows.map((r) => this.plain(r[c]?.inline ?? null).length))));
    const sum = weights.reduce((a, b) => a + b, 0);
    const totalTw = Math.round(((this.p.image.maxWidthCm * 567) * t.widthPercent) / 100);
    const grid = weights.map((w) => Math.round((totalTw * w) / sum));

    const line = (tag: string, on: boolean) =>
      on ? `<w:${tag} w:val="${t.borderStyle}" w:sz="${eighths(t.borderWidthPt)}" w:space="0" w:color="${t.borderColor}"/>` : `<w:${tag} w:val="nil"/>`;
    const b = t.borders;
    const borders = `<w:tblBorders>${line("top", b !== "none")}${line("left", b === "all" || b === "outer")}${line("bottom", b !== "none")}` +
      `${line("right", b === "all" || b === "outer")}${line("insideH", b === "all" || b === "horizontal")}${line("insideV", b === "all")}</w:tblBorders>`;
    const pad = twips(t.cellPaddingPt);
    const tblPr = `<w:tblPr><w:tblW w:w="${Math.round(t.widthPercent * 50)}" w:type="pct"/><w:jc w:val="${t.align === "justify" ? "left" : t.align}"/>${borders}` +
      `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="${Math.round(pad / 2)}" w:type="dxa"/><w:left w:w="${Math.round(pad * 1.5)}" w:type="dxa"/>` +
      `<w:bottom w:w="${Math.round(pad / 2)}" w:type="dxa"/><w:right w:w="${Math.round(pad * 1.5)}" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0420" w:firstRow="1"/></w:tblPr>`;

    const body = rows.map((r, ri) => {
      const header = ri === 0 && r.some((c) => c.header);
      const cells = Array.from({ length: cols }, (_, ci) => {
        const c = r[ci] ?? { inline: null, align: null, header };
        const fill = header ? t.headerBackground : t.zebraBackground && ri > 0 && ri % 2 === 0 ? t.zebraBackground : "";
        const base: ElementFmt | undefined = header ? { color: t.headerColor, bold: t.headerBold } : undefined;
        const content = this.inline(c.inline, base);
        const key = header ? "tableHeader" : "tableCell";
        const p = this.para(key, content, { jc: c.align ?? undefined });
        return `<w:tc><w:tcPr><w:tcW w:w="${grid[ci]}" w:type="dxa"/>${fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` : ""}</w:tcPr>${p}</w:tc>`;
      });
      const trPr = (t.keepRowsTogether ? "<w:cantSplit/>" : "") + (header && t.repeatHeader ? "<w:tblHeader/>" : "");
      return `<w:tr>${wrapTag("trPr", trPr)}${cells.join("")}</w:tr>`;
    });
    this.add(`<w:tbl>${tblPr}<w:tblGrid>${grid.map((g) => `<w:gridCol w:w="${g}"/>`).join("")}</w:tblGrid>${body.join("")}</w:tbl>`);
    // Word richiede un paragrafo dopo la tabella; uno vuoto minimo evita che si fonda con quella successiva.
    this.add('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/></w:pPr></w:p>');
  }

  // ----- note a piè -----

  private collectFootnotes(tokens: Token[], from: number): number {
    let i = from + 1;
    while (i < tokens.length && tokens[i].type !== "footnote_block_close") {
      if (tokens[i].type === "footnote_open") {
        const id = Number(tokens[i].meta?.id ?? 0) + 1;
        const paras: string[] = [];
        for (i++; tokens[i].type !== "footnote_close"; i++) {
          if (tokens[i].type === "paragraph_open") {
            const body = this.inline(tokens[i + 1]);
            const lead = paras.length === 0
              ? '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>' + this.text(" ", blankState())
              : "";
            paras.push(`<w:p><w:pPr><w:pStyle w:val="${this.style("footnote")}"/></w:pPr>${lead}${body}</w:p>`);
          }
        }
        this.footnotes.push(`<w:footnote w:id="${id}">${paras.join("")}</w:footnote>`);
      }
      i++;
    }
    return i;
  }

  // ----- percorso dei token -----

  private currentMermaid: string[] = [];

  walk(prep: Prepared): void {
    this.currentMermaid = prep.mermaid;
    const tokens = prep.tokens;
    let listItemStart = false;
    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i];
      switch (t.type) {
        case "heading_open": this.heading(Number(t.tag.slice(1)), tokens[i + 1]); i += 2; break;
        case "paragraph_open": this.paragraph(tokens[i + 1], listItemStart); listItemStart = false; i += 2; break;
        case "fence":
          if (isMermaidFence(t)) this.diagram();
          else this.code(t.content);
          break;
        case "code_block": this.code(t.content); break;
        case "hr": this.hr(); break;
        case "html_block": this.htmlBlock(t.content); break;
        case "blockquote_open": this.quoteDepth++; break;
        case "blockquote_close": this.quoteDepth--; break;
        case "bullet_list_open": this.startList(false, 1); break;
        case "ordered_list_open": this.startList(true, Number(at(t, "start") || 1)); break;
        case "bullet_list_close": case "ordered_list_close": this.lists.pop(); break;
        case "list_item_open": listItemStart = true; break;
        case "dt_open": this.add(this.para("dt", this.inline(tokens[i + 1]))); i += 2; break;
        case "dd_open": this.inDefinition = true; break;
        case "dd_close": this.inDefinition = false; break;
        case "footnote_block_open": i = this.collectFootnotes(tokens, i); break;
        case "table_open": {
          const end = tokens.findIndex((x, k) => k > i && x.type === "table_close");
          const rows: Cell[][] = [];
          for (let k = i; k < end; k++) {
            if (tokens[k].type === "tr_open") rows.push([]);
            if (tokens[k].type === "th_open" || tokens[k].type === "td_open") {
              const style = at(tokens[k], "style");
              const al = /text-align:\s*(left|center|right)/.exec(style)?.[1] as Align | undefined;
              rows[rows.length - 1].push({ inline: tokens[k + 1], align: al ?? null, header: tokens[k].type === "th_open" });
            }
          }
          if (rows.length) this.table(rows);
          i = end;
          break;
        }
      }
    }
  }

  // ----- parti del pacchetto -----

  numberingXml(): string {
    const bullets = ["•", "◦", "▪"];
    const ordered = ["decimal", "lowerLetter", "lowerRoman"];
    const lvls = (kind: "bullet" | "ordered") =>
      Array.from({ length: 9 }, (_, l) => {
        const fmt = kind === "bullet" ? "bullet" : ordered[l % 3];
        const text = kind === "bullet" ? bullets[l % 3] : `%${l + 1}.`;
        const font = kind === "bullet" ? '<w:rPr><w:rFonts w:ascii="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol" w:hint="default"/></w:rPr>' : "";
        return `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/>` +
          `<w:pPr><w:ind w:left="${360 + l * 360}" w:hanging="360"/></w:pPr>${font}</w:lvl>`;
      }).join("");
    return `<w:numbering xmlns:w="${NS_W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvls("bullet")}</w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvls("ordered")}</w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>${this.nums.join("")}</w:numbering>`;
  }

  stylesXml(): string {
    // Normal va definito una sola volta: due stili con lo stesso id confondono Word.
    const normal = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>';
    const others = [...this.styles.entries()].filter(([id]) => id !== "Normal").map(([, xml]) => xml);
    return `<w:styles xmlns:w="${NS_W}">${normal}${others.join("")}</w:styles>`;
  }

  footnotesXml(): string | undefined {
    if (!this.footnotes.length) return undefined;
    return `<w:footnotes xmlns:w="${NS_W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
      `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>${this.footnotes.join("")}</w:footnotes>`;
  }

  finish(): string {
    return this.pkg.toFlatOpc({
      // Word non importa il segno di paragrafo finale: senza un paragrafo sacrificale l'ultimo paragrafo perderebbe lo stile.
      body: this.body.join("") + "<w:p/>",
      styles: this.stylesXml(),
      numbering: this.numberingXml(),
      footnotes: this.footnotesXml(),
    });
  }
}

/**
 * Genera il pacchetto OOXML del documento.
 * `rendered[i]` è il diagramma i-esimo già disegnato (null se non valido); `images` associa src → immagine risolta.
 */
export function buildPackage(
  prep: Prepared,
  profile: Profile,
  rendered: (RenderedDiagram | null)[],
  images: Map<string, ImageRes>,
): string {
  const b = new Builder(profile, rendered, images);
  b.walk(prep);
  return b.finish();
}
