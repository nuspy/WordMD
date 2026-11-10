import mermaid from "mermaid";
import { defaultProfile, type MermaidFmt } from "./profile";

export interface RenderedDiagram {
  /** SVG autonomo: stili inline, niente foreignObject, width/height espliciti. */
  svg: string;
  /** PNG (base64, senza prefisso data:) usato da Word come immagine di ripiego. */
  pngBase64: string;
  /** Dimensioni in px CSS (96 dpi) del diagramma. */
  width: number;
  height: number;
}

const PNG_SCALE = 2;
const SVG_NS = "http://www.w3.org/2000/svg";

/** Proprietà CSS da copiare come attributi: gli editor che non leggono <style> le vedono comunque. */
const STYLE_PROPS = [
  "fill", "fill-opacity", "stroke", "stroke-width", "stroke-dasharray", "stroke-opacity", "stroke-linecap",
  "stroke-linejoin", "opacity", "font-family", "font-size", "font-weight", "font-style", "text-anchor",
  "dominant-baseline",
];
const STYLED_TAGS = new Set([
  "path", "rect", "circle", "ellipse", "polygon", "polyline", "line", "text", "tspan",
]);

let counter = 0;

const hex = (c: string): string => (c.startsWith("#") || c === "transparent" ? c : "#" + c);

function configFor(fmt: MermaidFmt): Parameters<typeof mermaid.initialize>[0] {
  const vars: Record<string, string> = { fontFamily: fmt.fontFamily, fontSize: `${fmt.fontSizePx}px` };
  if (fmt.theme === "base") {
    for (const [k, v] of Object.entries(fmt.variables)) if (v) vars[k] = hex(v);
  }
  const themeCSS =
    `.flowchart-link, .edgePath .path, .messageLine0, .messageLine1, .relation, .transition, .edge-thickness-normal { stroke-width: ${fmt.lineWidthPx}px !important; }\n` +
    `.node rect, .node circle, .node ellipse, .node polygon, .node path, .actor, .classBox, .cluster rect { stroke-width: ${fmt.nodeBorderPx}px !important; }`;
  const noMaxWidth = { useMaxWidth: false };
  return {
    startOnLoad: false,
    securityLevel: "strict",
    theme: fmt.theme,
    themeVariables: vars,
    themeCSS,
    fontFamily: fmt.fontFamily,
    // Word non disegna <foreignObject>: le etichette devono essere testo SVG.
    htmlLabels: false,
    flowchart: { htmlLabels: false, useMaxWidth: false, curve: fmt.curve },
    sequence: noMaxWidth, gantt: noMaxWidth, journey: noMaxWidth, timeline: noMaxWidth, pie: noMaxWidth,
    class: { htmlLabels: false, useMaxWidth: false }, state: noMaxWidth, er: noMaxWidth, gitGraph: noMaxWidth,
    requirement: noMaxWidth, mindmap: noMaxWidth, sankey: noMaxWidth, xyChart: noMaxWidth, quadrantChart: noMaxWidth,
    block: noMaxWidth, c4: noMaxWidth, architecture: noMaxWidth, packet: noMaxWidth, kanban: noMaxWidth,
  } as Parameters<typeof mermaid.initialize>[0];
}

/** Verifica la sintassi; restituisce il messaggio d'errore oppure null se il codice è valido. */
export async function checkMermaid(code: string): Promise<string | null> {
  mermaid.initialize(configFor(defaultProfile().mermaid));
  try {
    await mermaid.parse(code);
    return null;
  } catch (e) {
    return e instanceof Error ? e.message : String(e);
  }
}

export async function renderMermaid(code: string, fmt: MermaidFmt = defaultProfile().mermaid): Promise<RenderedDiagram> {
  mermaid.initialize(configFor(fmt));
  const id = `wordmd-mermaid-${++counter}`;
  let raw: string;
  try {
    raw = (await mermaid.render(id, code)).svg;
  } finally {
    // In caso di errore mermaid lascia un nodo di servizio nel documento.
    document.getElementById(`d${id}`)?.remove();
  }
  const { svg, width, height } = makeStandalone(raw, fmt.background);
  const pngBase64 = await svgToPng(svg, width, height);
  return { svg, pngBase64, width, height };
}

/** Sostituisce ogni <foreignObject> (etichette HTML) con testo SVG, l'unica forma che Word sa disegnare. */
function flattenForeignObjects(svg: SVGSVGElement): void {
  for (const fo of Array.from(svg.querySelectorAll("foreignObject"))) {
    const host = (fo.firstElementChild as HTMLElement | null) ?? (fo as unknown as HTMLElement);
    host.querySelectorAll("br").forEach((br) => br.replaceWith("\n"));
    const lines = (host.textContent ?? "").split("\n").map((l) => l.trim()).filter(Boolean);
    if (!lines.length) {
      fo.remove();
      continue;
    }
    const x = Number(fo.getAttribute("x") ?? 0);
    const y = Number(fo.getAttribute("y") ?? 0);
    const w = Number(fo.getAttribute("width") ?? 0);
    const h = Number(fo.getAttribute("height") ?? 0);
    const cs = getComputedStyle(host);
    const size = parseFloat(cs.fontSize) || 14;
    const text = document.createElementNS(SVG_NS, "text");
    text.setAttribute("x", String(x + w / 2));
    text.setAttribute("text-anchor", "middle");
    text.setAttribute("fill", cs.color || "#000");
    text.setAttribute("font-family", cs.fontFamily);
    text.setAttribute("font-size", String(size));
    text.setAttribute("font-weight", cs.fontWeight);
    const top = y + h / 2 - ((lines.length - 1) * size * 1.2) / 2;
    lines.forEach((line, i) => {
      const span = document.createElementNS(SVG_NS, "tspan");
      span.setAttribute("x", String(x + w / 2));
      span.setAttribute("y", String(top + i * size * 1.2 + size * 0.35));
      span.textContent = line;
      text.appendChild(span);
    });
    fo.replaceWith(text);
  }
}

/** Rende l'SVG leggibile fuori dalla pagina: copia gli stili calcolati negli attributi e fissa le dimensioni. */
export function makeStandalone(rawSvg: string, background = "transparent"): { svg: string; width: number; height: number } {
  const holder = document.createElement("div");
  holder.style.cssText = "position:absolute;left:-99999px;top:0;visibility:hidden";
  holder.innerHTML = rawSvg;
  document.body.appendChild(holder);
  try {
    const svg = holder.querySelector("svg");
    if (!svg) throw new Error("Mermaid non ha prodotto un SVG");

    flattenForeignObjects(svg);
    for (const el of Array.from(svg.querySelectorAll("*"))) {
      if (!STYLED_TAGS.has(el.localName)) continue;
      const cs = getComputedStyle(el);
      for (const prop of STYLE_PROPS) {
        const value = cs.getPropertyValue(prop);
        if (value) el.setAttribute(prop, value);
      }
    }
    svg.querySelectorAll("style").forEach((s) => s.remove());

    const vb = (svg.getAttribute("viewBox") ?? "").split(/[\s,]+/).map(Number);
    const hasVb = vb.length === 4 && vb[2] > 0 && vb[3] > 0;
    const width = Math.ceil(hasVb ? vb[2] : svg.getBoundingClientRect().width || 600);
    const height = Math.ceil(hasVb ? vb[3] : svg.getBoundingClientRect().height || 400);
    svg.setAttribute("xmlns", SVG_NS);
    svg.setAttribute("width", String(width));
    svg.setAttribute("height", String(height));
    svg.removeAttribute("style");

    if (background && background !== "transparent") {
      const rect = document.createElementNS(SVG_NS, "rect");
      rect.setAttribute("x", String(hasVb ? vb[0] : 0));
      rect.setAttribute("y", String(hasVb ? vb[1] : 0));
      rect.setAttribute("width", String(width));
      rect.setAttribute("height", String(height));
      rect.setAttribute("fill", hex(background));
      svg.insertBefore(rect, svg.firstChild);
    }
    return { svg: new XMLSerializer().serializeToString(svg), width, height };
  } finally {
    holder.remove();
  }
}

export function svgToPng(svg: string, width: number, height: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round(width * PNG_SCALE));
      canvas.height = Math.max(1, Math.round(height * PNG_SCALE));
      const ctx = canvas.getContext("2d");
      if (!ctx) return reject(new Error("Canvas 2D non disponibile"));
      ctx.scale(PNG_SCALE, PNG_SCALE);
      ctx.drawImage(img, 0, 0, width, height);
      try {
        resolve(canvas.toDataURL("image/png").split(",")[1]);
      } catch (e) {
        reject(e);
      }
    };
    img.onerror = () => reject(new Error("Impossibile convertire l'SVG in PNG"));
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
  });
}
