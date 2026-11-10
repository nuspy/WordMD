/** Pacchetto Flat OPC (il formato di insertOoxml / getOoxml di Word) e costruzione delle immagini in linea. */

export const MERMAID_TITLE = "WordMD Mermaid";

const NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EMU_PER_PX = 9525;
export const EMU_PER_CM = 360000;

export function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Per i valori di attributo: preserva anche gli a capo (il testo alternativo contiene il sorgente Mermaid). */
export function escapeAttr(s: string): string {
  return escapeXml(s).replace(/"/g, "&quot;").replace(/\r?\n/g, "&#10;").replace(/\t/g, "&#9;");
}

export function utf8ToBase64(text: string): string {
  const bytes = new TextEncoder().encode(text);
  return bytesToBase64(bytes);
}

export function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(bin);
}

export interface DrawingOptions {
  pngRid: string;
  svgRid?: string;
  /** Dimensioni in px CSS del contenuto; vengono ridotte per stare entro maxWidthCm. */
  widthPx: number;
  heightPx: number;
  maxWidthCm: number;
  id: number;
  name: string;
  descr: string;
  title?: string;
}

/** Corpo di un paragrafo con un'immagine: PNG come immagine principale, SVG come estensione svgBlip. */
export function inlineDrawing(o: DrawingOptions): { xml: string; cx: number; cy: number } {
  const scale = Math.min(1, (o.maxWidthCm * EMU_PER_CM) / (o.widthPx * EMU_PER_PX));
  const cx = Math.max(1, Math.round(o.widthPx * EMU_PER_PX * scale));
  const cy = Math.max(1, Math.round(o.heightPx * EMU_PER_PX * scale));
  const name = escapeAttr(o.name);
  const svg = o.svgRid
    ? `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${o.svgRid}"/></a:ext></a:extLst>`
    : "";
  const xml = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${o.id}" name="${name}" descr="${escapeAttr(o.descr)}"${o.title ? ` title="${escapeAttr(o.title)}"` : ""}/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="0" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${o.pngRid}">${svg}</a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  return { xml, cx, cy };
}

interface MediaPart {
  name: string;
  contentType: string;
  base64: string;
}

export class PackageBuilder {
  private rels: string[] = [];
  private media: MediaPart[] = [];
  private hyperlinks = new Map<string, string>();
  private nextRid = 1;
  private nextDocPr = 1;

  newDocPrId(): number {
    return this.nextDocPr++;
  }

  addMedia(base64: string, contentType: string, ext: string): string {
    const rid = `rId${this.nextRid++}`;
    const name = `media/wordmd${this.media.length + 1}.${ext}`;
    this.media.push({ name, contentType, base64 });
    this.rels.push(`<Relationship Id="${rid}" Type="${REL}/image" Target="${name}"/>`);
    return rid;
  }

  addHyperlink(url: string): string {
    let rid = this.hyperlinks.get(url);
    if (!rid) {
      rid = `rId${this.nextRid++}`;
      this.hyperlinks.set(url, rid);
      this.rels.push(`<Relationship Id="${rid}" Type="${REL}/hyperlink" Target="${escapeAttr(url)}" TargetMode="External"/>`);
    }
    return rid;
  }

  toFlatOpc(parts: { body: string; styles?: string; numbering?: string; footnotes?: string }): string {
    const rels = [...this.rels];
    if (parts.styles) rels.push(`<Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/>`);
    if (parts.numbering) rels.push(`<Relationship Id="rIdNumbering" Type="${REL}/numbering" Target="numbering.xml"/>`);
    if (parts.footnotes) rels.push(`<Relationship Id="rIdFootnotes" Type="${REL}/footnotes" Target="footnotes.xml"/>`);

    const part = (name: string, type: string, xml: string) =>
      `<pkg:part pkg:name="${name}" pkg:contentType="${type}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`;
    const wml = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const nsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    const out = [
      `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">`,
      part("/_rels/.rels", "application/vnd.openxmlformats-package.relationships+xml",
        `<Relationships xmlns="${NS_REL}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`),
      part("/word/_rels/document.xml.rels", "application/vnd.openxmlformats-package.relationships+xml",
        `<Relationships xmlns="${NS_REL}">${rels.join("")}</Relationships>`),
      part("/word/document.xml", `${wml}.document.main+xml`,
        `<w:document xmlns:w="${nsW}" xmlns:r="${REL}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${parts.body}</w:body></w:document>`),
    ];
    if (parts.styles) out.push(part("/word/styles.xml", `${wml}.styles+xml`, parts.styles));
    if (parts.numbering) out.push(part("/word/numbering.xml", `${wml}.numbering+xml`, parts.numbering));
    if (parts.footnotes) out.push(part("/word/footnotes.xml", `${wml}.footnotes+xml`, parts.footnotes));
    for (const m of this.media) {
      out.push(`<pkg:part pkg:name="/word/${m.name}" pkg:contentType="${m.contentType}" pkg:compression="store"><pkg:binaryData>${m.base64}</pkg:binaryData></pkg:part>`);
    }
    out.push("</pkg:package>");
    return out.join("\n");
  }
}
