import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { MERMAID_TITLE } from "./opc";

const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export interface DocxDiagram {
  /** Sorgente Mermaid salvato nel testo alternativo. */
  code: string;
  /** SVG incorporato nel documento, se presente. */
  svg?: string;
}

function parse(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("XML del documento non valido");
  return doc;
}

function diagramsOf(doc: Document): Element[] {
  return Array.from(doc.getElementsByTagNameNS(NS_WP, "docPr")).filter((e) => e.getAttribute("title") === MERMAID_TITLE);
}

/** Legge i diagrammi Mermaid del .docx / .mdocx con il loro SVG incorporato. */
export function listDiagrams(bytes: Uint8Array): DocxDiagram[] {
  const files = unzipSync(bytes);
  const docXml = files["word/document.xml"];
  if (!docXml) throw new Error("Il file non è un documento Word (manca word/document.xml)");
  const doc = parse(strFromU8(docXml));
  const relsXml = files["word/_rels/document.xml.rels"];
  const targets = new Map<string, string>();
  if (relsXml) {
    for (const r of Array.from(parse(strFromU8(relsXml)).documentElement.children)) {
      targets.set(r.getAttribute("Id") ?? "", r.getAttribute("Target") ?? "");
    }
  }
  return diagramsOf(doc).map((pr) => {
    const code = pr.getAttribute("descr") ?? "";
    // docPr → ... → asvg:svgBlip r:embed (è nello stesso wp:inline)
    const inline = pr.parentElement;
    const svgBlip = inline ? Array.from(inline.getElementsByTagNameNS("*", "svgBlip"))[0] : undefined;
    const rid = svgBlip?.getAttributeNS(NS_R, "embed") ?? svgBlip?.getAttribute("r:embed");
    const target = rid ? targets.get(rid) : undefined;
    const part = target ? files["word/" + target.replace(/^\/?(word\/)?/, "")] : undefined;
    return { code, svg: part ? strFromU8(part) : undefined };
  });
}

/**
 * Copia "piatta" per la condivisione: le immagini restano (PNG + SVG) ma perdono il sorgente Mermaid,
 * quindi non sono più modificabili come diagrammi.
 */
export function flattenDocx(bytes: Uint8Array): Uint8Array {
  const files = unzipSync(bytes);
  const docXml = files["word/document.xml"];
  if (!docXml) throw new Error("Il file non è un documento Word (manca word/document.xml)");
  const doc = parse(strFromU8(docXml));
  diagramsOf(doc).forEach((pr, i) => {
    pr.removeAttribute("title");
    pr.setAttribute("descr", `Diagramma ${i + 1}`);
  });
  files["word/document.xml"] = strToU8(new XMLSerializer().serializeToString(doc));
  return zipSync(files, { level: 6 });
}

/** True se il pacchetto contiene almeno un diagramma Mermaid modificabile. */
export function hasEditableDiagrams(bytes: Uint8Array): boolean {
  return listDiagrams(bytes).length > 0;
}
