import type { RenderedDiagram } from "./mermaidRender";
import type { Align } from "./profile";
import { MERMAID_TITLE, PackageBuilder, inlineDrawing, utf8ToBase64 } from "./opc";

export { MERMAID_TITLE };

const JC: Record<Align, string> = { left: "left", center: "center", right: "right", justify: "both" };

/** Esegue il corpo `<w:r>` di un diagramma Mermaid: l'immagine porta il sorgente nel testo alternativo. */
export function diagramRun(pkg: PackageBuilder, code: string, d: RenderedDiagram, label: string, maxWidthCm: number): string {
  const pngRid = pkg.addMedia(d.pngBase64, "image/png", "png");
  const svgRid = pkg.addMedia(utf8ToBase64(d.svg), "image/svg+xml", "svg");
  const { xml } = inlineDrawing({
    pngRid, svgRid, widthPx: d.width, heightPx: d.height, maxWidthCm,
    id: pkg.newDocPrId(), name: label, descr: code, title: MERMAID_TITLE,
  });
  return `<w:r>${xml}</w:r>`;
}

export function diagramParagraph(run: string, align: Align, keepNext: boolean): string {
  return `<w:p><w:pPr>${keepNext ? "<w:keepNext/>" : ""}<w:jc w:val="${JC[align]}"/></w:pPr>${run}</w:p>`;
}

/** Pacchetto Flat OPC con un solo paragrafo-diagramma (inserimento dal pannello "Diagramma"). */
export function buildDiagramPackage(
  code: string,
  d: RenderedDiagram,
  label: string,
  opts: { maxWidthCm: number; align: Align },
): string {
  const pkg = new PackageBuilder();
  const run = diagramRun(pkg, code, d, label, opts.maxWidthCm);
  return pkg.toFlatOpc({ body: diagramParagraph(run, opts.align, false) });
}
