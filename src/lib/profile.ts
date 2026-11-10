/**
 * Profilo di importazione: come ogni elemento Markdown diventa formattazione Word
 * e come vengono disegnati i diagrammi Mermaid. I colori sono esadecimali RRGGBB senza "#".
 */

export type Align = "left" | "center" | "right" | "justify";
export type BorderStyle = "single" | "double" | "dashed" | "dotted";

export interface Border {
  color: string;
  widthPt: number;
  style?: BorderStyle;
}

export interface ElementFmt {
  /** Modalità "stili esistenti": nome canonico dello stile Word da applicare (es. "heading 1"). */
  wordStyle?: string;
  font?: string;
  sizePt?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Sfondo (ombreggiatura) del paragrafo, o evidenziazione per gli elementi in linea. */
  background?: string;
  align?: Align;
  beforePt?: number;
  afterPt?: number;
  /** Interlinea come multiplo (1 = singola, 1.5 = una e mezza). */
  lineSpacing?: number;
  indentPt?: number;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  borderTop?: Border;
  borderBottom?: Border;
  borderLeft?: Border;
}

export interface TableFmt {
  widthPercent: number;
  align: Align;
  borderColor: string;
  borderWidthPt: number;
  borderStyle: BorderStyle;
  /** "all" = griglia completa, "horizontal" = solo linee orizzontali, "outer" = solo contorno, "none". */
  borders: "all" | "horizontal" | "outer" | "none";
  headerBackground: string;
  headerColor: string;
  headerBold: boolean;
  /** Sfondo delle righe pari; vuoto = nessuna zebratura. */
  zebraBackground: string;
  cellPaddingPt: number;
  repeatHeader: boolean;
  /** Impedisce che una riga venga spezzata tra due pagine. */
  keepRowsTogether: boolean;
}

export interface ImageFmt {
  maxWidthCm: number;
  align: Align;
  /** Scrive il testo alternativo come didascalia sotto l'immagine. */
  captionFromAlt: boolean;
  keepWithCaption: boolean;
}

export interface MermaidFmt {
  theme: "base" | "default" | "neutral" | "dark" | "forest";
  /** Variabili di tema di Mermaid (usate dal tema "base"). */
  variables: Record<string, string>;
  fontFamily: string;
  fontSizePx: number;
  /** "transparent" oppure RRGGBB. */
  background: string;
  lineWidthPx: number;
  nodeBorderPx: number;
  curve: "basis" | "linear" | "step" | "cardinal" | "monotoneX";
  maxWidthCm: number;
  align: Align;
  /** Evita che il diagramma finisca su una pagina diversa dalla sua didascalia. */
  keepWithNext: boolean;
}

export interface Profile {
  name: string;
  mode: "existing" | "new";
  elements: Record<string, ElementFmt>;
  table: TableFmt;
  image: ImageFmt;
  mermaid: MermaidFmt;
}

export interface ElementDef {
  key: string;
  label: string;
  /** "p" = paragrafo, "c" = formato in linea. */
  kind: "p" | "c";
  /** Nome canonico dello stile Word predefinito (modalità "stili esistenti"). */
  builtIn: string;
  /** Suffisso del nome dello stile creato in modalità "nuovo stile". */
  en: string;
}

export const ELEMENTS: ElementDef[] = [
  { key: "h1", label: "Titolo 1", kind: "p", builtIn: "heading 1", en: "Heading 1" },
  { key: "h2", label: "Titolo 2", kind: "p", builtIn: "heading 2", en: "Heading 2" },
  { key: "h3", label: "Titolo 3", kind: "p", builtIn: "heading 3", en: "Heading 3" },
  { key: "h4", label: "Titolo 4", kind: "p", builtIn: "heading 4", en: "Heading 4" },
  { key: "h5", label: "Titolo 5", kind: "p", builtIn: "heading 5", en: "Heading 5" },
  { key: "h6", label: "Titolo 6", kind: "p", builtIn: "heading 6", en: "Heading 6" },
  { key: "p", label: "Paragrafo", kind: "p", builtIn: "Normal", en: "Paragraph" },
  { key: "quote", label: "Citazione", kind: "p", builtIn: "Quote", en: "Quote" },
  { key: "codeBlock", label: "Blocco di codice", kind: "p", builtIn: "HTML Preformatted", en: "Code Block" },
  { key: "ulist", label: "Elenco puntato", kind: "p", builtIn: "List Paragraph", en: "Bullet List" },
  { key: "olist", label: "Elenco numerato", kind: "p", builtIn: "List Paragraph", en: "Numbered List" },
  { key: "dt", label: "Termine di definizione", kind: "p", builtIn: "Normal", en: "Definition Term" },
  { key: "dd", label: "Definizione", kind: "p", builtIn: "Normal", en: "Definition" },
  { key: "tableHeader", label: "Intestazione di tabella", kind: "p", builtIn: "Normal", en: "Table Header" },
  { key: "tableCell", label: "Cella di tabella", kind: "p", builtIn: "Normal", en: "Table Cell" },
  { key: "caption", label: "Didascalia", kind: "p", builtIn: "caption", en: "Caption" },
  { key: "footnote", label: "Testo della nota", kind: "p", builtIn: "footnote text", en: "Footnote" },
  { key: "hr", label: "Linea orizzontale", kind: "p", builtIn: "Normal", en: "Horizontal Rule" },
  { key: "link", label: "Collegamento", kind: "c", builtIn: "Hyperlink", en: "Link" },
  { key: "strong", label: "Grassetto", kind: "c", builtIn: "Strong", en: "Strong" },
  { key: "em", label: "Corsivo", kind: "c", builtIn: "Emphasis", en: "Emphasis" },
  { key: "del", label: "Barrato", kind: "c", builtIn: "", en: "Strikethrough" },
  { key: "mark", label: "Evidenziato", kind: "c", builtIn: "", en: "Highlight" },
  { key: "codeInline", label: "Codice in linea", kind: "c", builtIn: "", en: "Inline Code" },
];

export const ALL_MERMAID_VARIABLES: { key: string; label: string }[] = [
  { key: "primaryColor", label: "Sfondo dei nodi" },
  { key: "primaryTextColor", label: "Testo dei nodi" },
  { key: "primaryBorderColor", label: "Bordo dei nodi" },
  { key: "secondaryColor", label: "Colore secondario" },
  { key: "tertiaryColor", label: "Colore terziario" },
  { key: "lineColor", label: "Colore delle linee" },
  { key: "textColor", label: "Testo generale" },
  { key: "noteBkgColor", label: "Sfondo delle note" },
  { key: "noteTextColor", label: "Testo delle note" },
  { key: "noteBorderColor", label: "Bordo delle note" },
  { key: "clusterBkg", label: "Sfondo dei gruppi" },
  { key: "clusterBorder", label: "Bordo dei gruppi" },
];

// ---------- preset ----------

const rule = (color: string, widthPt: number): Border => ({ color, widthPt });

function base(name: string, o: Partial<Profile> & { elements: Record<string, ElementFmt> }): Profile {
  return {
    name,
    mode: "new",
    table: {
      widthPercent: 100, align: "left", borderColor: "BFBFBF", borderWidthPt: 0.75, borderStyle: "single",
      borders: "all", headerBackground: "F2F2F2", headerColor: "000000", headerBold: true, zebraBackground: "",
      cellPaddingPt: 4, repeatHeader: true, keepRowsTogether: true,
    },
    image: { maxWidthCm: 16, align: "center", captionFromAlt: true, keepWithCaption: true },
    mermaid: {
      theme: "default", variables: {}, fontFamily: "Segoe UI", fontSizePx: 14, background: "transparent",
      lineWidthPx: 1.5, nodeBorderPx: 1.5, curve: "basis", maxWidthCm: 16, align: "center", keepWithNext: true,
    },
    ...o,
  };
}

function github(): Profile {
  const f = "Segoe UI";
  const head = (sizePt: number, border = false, color = "1F2328"): ElementFmt => ({
    font: f, sizePt, bold: true, color, beforePt: 18, afterPt: 10, keepNext: true, keepLines: true,
    ...(border ? { borderBottom: rule("D1D9E0", 0.75) } : {}),
  });
  return base("GitHub", {
    elements: {
      h1: head(24, true), h2: head(18, true), h3: head(14), h4: head(12), h5: head(11), h6: head(10.5, false, "59636E"),
      p: { font: f, sizePt: 11, color: "1F2328", beforePt: 0, afterPt: 8, lineSpacing: 1.35, align: "left" },
      quote: { font: f, sizePt: 11, color: "59636E", indentPt: 14, afterPt: 8, borderLeft: rule("D1D9E0", 3) },
      codeBlock: { font: "Consolas", sizePt: 9.5, color: "1F2328", background: "F6F8FA", afterPt: 10, lineSpacing: 1.2, keepLines: true },
      ulist: { font: f, sizePt: 11, color: "1F2328", afterPt: 3 },
      olist: { font: f, sizePt: 11, color: "1F2328", afterPt: 3 },
      dt: { font: f, sizePt: 11, bold: true, beforePt: 6, afterPt: 2 },
      dd: { font: f, sizePt: 11, indentPt: 20, afterPt: 4 },
      tableHeader: { font: f, sizePt: 10.5, bold: true },
      tableCell: { font: f, sizePt: 10.5 },
      caption: { font: f, sizePt: 9.5, italic: true, color: "59636E", align: "center", afterPt: 10 },
      footnote: { font: f, sizePt: 9, color: "59636E" },
      hr: { afterPt: 8, borderBottom: rule("D1D9E0", 1.5) },
      link: { color: "0969DA" },
      strong: { bold: true }, em: { italic: true }, del: { strike: true }, mark: { background: "FFF8C5" },
      codeInline: { font: "Consolas", sizePt: 10, background: "EFF1F3" },
    },
    table: {
      widthPercent: 100, align: "left", borderColor: "D1D9E0", borderWidthPt: 0.75, borderStyle: "single", borders: "all",
      headerBackground: "F6F8FA", headerColor: "1F2328", headerBold: true, zebraBackground: "F6F8FA",
      cellPaddingPt: 4, repeatHeader: true, keepRowsTogether: true,
    },
    mermaid: {
      theme: "base", fontFamily: f, fontSizePx: 14, background: "transparent", lineWidthPx: 1.5, nodeBorderPx: 1.2,
      curve: "basis", maxWidthCm: 16, align: "center", keepWithNext: true,
      variables: {
        primaryColor: "DDF4FF", primaryTextColor: "1F2328", primaryBorderColor: "54AEFF", lineColor: "59636E",
        secondaryColor: "FFF8C5", tertiaryColor: "F6F8FA", noteBkgColor: "FFF8C5", noteTextColor: "1F2328",
        noteBorderColor: "D4A72C", clusterBkg: "F6F8FA", clusterBorder: "D1D9E0", textColor: "1F2328",
      },
    },
  });
}

function academic(): Profile {
  const serif = "Cambria";
  const navy = "1F2A44";
  return base("Accademico elegante", {
    elements: {
      h1: { font: serif, sizePt: 20, bold: true, color: navy, align: "center", beforePt: 24, afterPt: 14, keepNext: true },
      h2: { font: serif, sizePt: 15, bold: true, color: navy, beforePt: 18, afterPt: 8, keepNext: true, borderBottom: rule(navy, 0.5) },
      h3: { font: serif, sizePt: 12.5, bold: true, italic: true, color: navy, beforePt: 14, afterPt: 6, keepNext: true },
      h4: { font: serif, sizePt: 11.5, bold: true, color: "333333", beforePt: 10, afterPt: 4, keepNext: true },
      h5: { font: serif, sizePt: 11.5, italic: true, color: "333333", beforePt: 8, afterPt: 3, keepNext: true },
      h6: { font: serif, sizePt: 11, italic: true, color: "555555", beforePt: 8, afterPt: 3, keepNext: true },
      p: { font: serif, sizePt: 11.5, color: "1A1A1A", align: "justify", afterPt: 6, lineSpacing: 1.25 },
      quote: { font: serif, sizePt: 11, italic: true, color: "444444", indentPt: 28, align: "justify", afterPt: 8 },
      codeBlock: { font: "Consolas", sizePt: 9.5, background: "F4F4F2", borderLeft: rule("999999", 2), afterPt: 8, keepLines: true },
      ulist: { font: serif, sizePt: 11.5, align: "justify", afterPt: 2 },
      olist: { font: serif, sizePt: 11.5, align: "justify", afterPt: 2 },
      dt: { font: serif, sizePt: 11.5, bold: true, beforePt: 6, afterPt: 1 },
      dd: { font: serif, sizePt: 11.5, indentPt: 24, afterPt: 4 },
      tableHeader: { font: serif, sizePt: 10.5, bold: true },
      tableCell: { font: serif, sizePt: 10.5 },
      caption: { font: serif, sizePt: 10, italic: true, align: "center", afterPt: 12 },
      footnote: { font: serif, sizePt: 9.5 },
      hr: { afterPt: 10, borderBottom: rule("999999", 0.5) },
      link: { color: "1F3A8A", underline: true },
      strong: { bold: true }, em: { italic: true }, del: { strike: true }, mark: { background: "FFF2B3" },
      codeInline: { font: "Consolas", sizePt: 10, background: "F4F4F2" },
    },
    table: {
      widthPercent: 100, align: "center", borderColor: "1A1A1A", borderWidthPt: 1, borderStyle: "single", borders: "horizontal",
      headerBackground: "", headerColor: "1A1A1A", headerBold: true, zebraBackground: "", cellPaddingPt: 4,
      repeatHeader: true, keepRowsTogether: true,
    },
    image: { maxWidthCm: 14, align: "center", captionFromAlt: true, keepWithCaption: true },
    mermaid: {
      theme: "base", fontFamily: serif, fontSizePx: 13, background: "FFFFFF", lineWidthPx: 1.2, nodeBorderPx: 1,
      curve: "linear", maxWidthCm: 14, align: "center", keepWithNext: true,
      variables: {
        primaryColor: "F2F2F2", primaryTextColor: navy, primaryBorderColor: navy, lineColor: navy,
        secondaryColor: "E6E6E6", tertiaryColor: "FAFAFA", noteBkgColor: "FFFBE6", noteTextColor: "1A1A1A",
        noteBorderColor: "999999", clusterBkg: "FAFAFA", clusterBorder: "999999", textColor: "1A1A1A",
      },
    },
  });
}

function wordClassic(): Profile {
  const head = (sizePt: number, color: string): ElementFmt => ({
    font: "Calibri Light", sizePt, color, beforePt: 12, afterPt: 0, keepNext: true, keepLines: true,
  });
  return base("Word classico", {
    elements: {
      h1: head(16, "2F5496"), h2: head(13, "2F5496"), h3: head(12, "1F3763"),
      h4: { ...head(11, "2F5496"), italic: true }, h5: head(11, "2F5496"), h6: { ...head(11, "1F3763"), italic: true },
      p: { font: "Calibri", sizePt: 11, afterPt: 8, lineSpacing: 1.08 },
      quote: { font: "Calibri", sizePt: 11, italic: true, color: "404040", indentPt: 36, align: "center", afterPt: 10 },
      codeBlock: { font: "Consolas", sizePt: 10, background: "F2F2F2", afterPt: 8, keepLines: true },
      ulist: { font: "Calibri", sizePt: 11, afterPt: 0 },
      olist: { font: "Calibri", sizePt: 11, afterPt: 0 },
      dt: { font: "Calibri", sizePt: 11, bold: true, afterPt: 0 },
      dd: { font: "Calibri", sizePt: 11, indentPt: 24, afterPt: 4 },
      tableHeader: { font: "Calibri", sizePt: 11, bold: true },
      tableCell: { font: "Calibri", sizePt: 11 },
      caption: { font: "Calibri", sizePt: 9, italic: true, color: "44546A", afterPt: 10 },
      footnote: { font: "Calibri", sizePt: 10 },
      hr: { afterPt: 8, borderBottom: rule("A6A6A6", 0.75) },
      link: { color: "0563C1", underline: true },
      strong: { bold: true }, em: { italic: true }, del: { strike: true }, mark: { background: "FFFF00" },
      codeInline: { font: "Consolas", sizePt: 10 },
    },
    table: {
      widthPercent: 100, align: "left", borderColor: "BFBFBF", borderWidthPt: 0.75, borderStyle: "single", borders: "all",
      headerBackground: "D9E2F3", headerColor: "000000", headerBold: true, zebraBackground: "", cellPaddingPt: 4,
      repeatHeader: true, keepRowsTogether: true,
    },
    mermaid: {
      theme: "default", variables: {}, fontFamily: "Calibri", fontSizePx: 14, background: "transparent",
      lineWidthPx: 1.5, nodeBorderPx: 1.5, curve: "basis", maxWidthCm: 16, align: "center", keepWithNext: true,
    },
  });
}

function corporate(): Profile {
  const accent = "0F4C81";
  const f = "Segoe UI";
  const head = (sizePt: number, extra: ElementFmt = {}): ElementFmt => ({
    font: f, sizePt, bold: true, color: accent, beforePt: 16, afterPt: 8, keepNext: true, keepLines: true, ...extra,
  });
  return base("Report aziendale", {
    elements: {
      h1: head(22, { pageBreakBefore: true, beforePt: 0, afterPt: 14, borderBottom: rule(accent, 2) }),
      h2: head(15), h3: head(12.5), h4: head(11, { color: "333333" }), h5: head(10.5, { color: "333333" }),
      h6: head(10.5, { color: "666666", bold: false, italic: true }),
      p: { font: f, sizePt: 10.5, color: "222222", afterPt: 8, lineSpacing: 1.3 },
      quote: { font: f, sizePt: 10.5, italic: true, color: "333333", background: "EEF3F8", indentPt: 12, afterPt: 10, borderLeft: rule(accent, 3) },
      codeBlock: { font: "Consolas", sizePt: 9.5, background: "F3F4F6", afterPt: 10, keepLines: true },
      ulist: { font: f, sizePt: 10.5, afterPt: 3 },
      olist: { font: f, sizePt: 10.5, afterPt: 3 },
      dt: { font: f, sizePt: 10.5, bold: true, color: accent, beforePt: 6, afterPt: 1 },
      dd: { font: f, sizePt: 10.5, indentPt: 20, afterPt: 4 },
      tableHeader: { font: f, sizePt: 10, bold: true },
      tableCell: { font: f, sizePt: 10 },
      caption: { font: f, sizePt: 9, italic: true, color: "666666", align: "center", afterPt: 12 },
      footnote: { font: f, sizePt: 8.5, color: "555555" },
      hr: { afterPt: 8, borderBottom: rule(accent, 1) },
      link: { color: accent, underline: true },
      strong: { bold: true }, em: { italic: true }, del: { strike: true }, mark: { background: "FFE9A8" },
      codeInline: { font: "Consolas", sizePt: 9.5, background: "F3F4F6" },
    },
    table: {
      widthPercent: 100, align: "left", borderColor: "C9D3DE", borderWidthPt: 0.75, borderStyle: "single", borders: "all",
      headerBackground: accent, headerColor: "FFFFFF", headerBold: true, zebraBackground: "F2F6FA", cellPaddingPt: 5,
      repeatHeader: true, keepRowsTogether: true,
    },
    mermaid: {
      theme: "base", fontFamily: f, fontSizePx: 14, background: "transparent", lineWidthPx: 1.6, nodeBorderPx: 1.4,
      curve: "basis", maxWidthCm: 16, align: "center", keepWithNext: true,
      variables: {
        primaryColor: "E3EEF8", primaryTextColor: "0B2A47", primaryBorderColor: accent, lineColor: accent,
        secondaryColor: "FDF1D8", tertiaryColor: "F2F6FA", noteBkgColor: "FDF1D8", noteTextColor: "0B2A47",
        noteBorderColor: "C99A2E", clusterBkg: "F2F6FA", clusterBorder: "C9D3DE", textColor: "222222",
      },
    },
  });
}

/** Modalità "stili esistenti": ogni elemento punta a uno stile predefinito di Word, senza formattazione propria. */
export function existingStylesProfile(): Profile {
  const elements: Record<string, ElementFmt> = {};
  for (const e of ELEMENTS) elements[e.key] = { wordStyle: e.builtIn };
  elements.codeInline = { font: "Consolas" };
  elements.del = { strike: true };
  elements.mark = { background: "FFFF00" };
  const p = base("Stili esistenti di Word", { elements });
  p.mode = "existing";
  return p;
}

export const PRESETS: { id: string; make: () => Profile }[] = [
  { id: "Word classico", make: wordClassic },
  { id: "GitHub", make: github },
  { id: "Accademico elegante", make: academic },
  { id: "Report aziendale", make: corporate },
  { id: "Stili esistenti di Word", make: existingStylesProfile },
];

export function defaultProfile(): Profile {
  return wordClassic();
}

/** Completa un profilo letto da disco con i valori mancanti (profili salvati da versioni precedenti). */
export function normalizeProfile(raw: Partial<Profile> | null | undefined): Profile {
  const d = defaultProfile();
  if (!raw) return d;
  return {
    name: raw.name || d.name,
    mode: raw.mode === "existing" ? "existing" : "new",
    elements: { ...d.elements, ...(raw.elements ?? {}) },
    table: { ...d.table, ...(raw.table ?? {}) },
    image: { ...d.image, ...(raw.image ?? {}) },
    mermaid: { ...d.mermaid, ...(raw.mermaid ?? {}), variables: { ...(raw.mermaid?.variables ?? {}) } },
  };
}
