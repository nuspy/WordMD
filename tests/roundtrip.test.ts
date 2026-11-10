import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { flattenDocx, listDiagrams } from "../src/lib/docxTools";
import { rasterSize, svgSize } from "../src/lib/imageInfo";
import { buildPackage, createdStyleNames, prepare } from "../src/lib/mdToOoxml";
import type { RenderedDiagram } from "../src/lib/mermaidRender";
import { ooxmlToMarkdown } from "../src/lib/ooxmlToMd";
import { PRESETS, existingStylesProfile } from "../src/lib/profile";
import { resolveImages } from "../src/lib/resources";

const PNG_1x1 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const diagram = (code: string): RenderedDiagram => ({
  svg: `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"><text>${code.length}</text></svg>`,
  pngBase64: PNG_1x1,
  width: 200,
  height: 100,
});

const MD = `---
title: Prova
---

# Titolo principale

Un paragrafo con **grassetto**, *corsivo*, ~~barrato~~, \`codice\` e un [collegamento](https://example.com/a?b=1&c=2).
Vai alla [sezione](#secondo-livello). Nota[^1] e H~2~O e x^2^ e ==evidenziato==.

## Secondo livello

- uno
- due
    - annidato
- [ ] da fare
- [x] fatto

1. primo
2. secondo

> una citazione
> su due righe

\`\`\`js
const a = 1;
if (a) { console.log("<ok>"); }
\`\`\`

| Nome | Valore | Note |
|:-----|:------:|-----:|
| a    | 1      | x \\| y |
| b    | 2      | z    |

---

\`\`\`mermaid
graph LR
  A[Inizio] --> B{Scelta}
  B -->|sì| C[Fine]
\`\`\`

![Logo piccolo](data:image/png;base64,${PNG_1x1})

Termine
: Definizione del termine

<!-- pagebreak -->

Ultimo paragrafo.

[^1]: Testo della nota.
`;

async function build(profile = PRESETS[0].make(), md = MD) {
  const prep = prepare(md);
  const rendered = prep.mermaid.map(diagram);
  const { images, missing } = await resolveImages(prep.imageSrcs, new Map());
  return { prep, flat: buildPackage(prep, profile, rendered, images), missing, images };
}

const xmlOk = (xml: string) => {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  expect(doc.getElementsByTagName("parsererror").length, doc.documentElement.textContent?.slice(0, 200)).toBe(0);
  return doc;
};

describe("Markdown → OOXML", () => {
  it.each(PRESETS.map((p) => p.id))("produce un pacchetto XML valido con il profilo %s", async (id) => {
    const profile = PRESETS.find((p) => p.id === id)!.make();
    const { flat, missing } = await build(profile);
    expect(missing).toEqual([]);
    const doc = xmlOk(flat);
    const names = Array.from(doc.getElementsByTagNameNS("http://schemas.microsoft.com/office/2006/xmlPackage", "part"))
      .map((p) => p.getAttribute("pkg:name"));
    expect(names).toEqual(expect.arrayContaining([
      "/word/document.xml", "/word/styles.xml", "/word/numbering.xml", "/word/footnotes.xml",
    ]));
    // ogni parte XML interna deve essere ben formata e ogni r:id deve avere la sua relazione
    const body = flat.slice(flat.indexOf("<w:body>"), flat.indexOf("</w:body>"));
    const rels = flat.slice(flat.indexOf("document.xml.rels"));
    for (const [, rid] of body.matchAll(/r:(?:id|embed)="(rId\d+)"/g)) expect(rels).toContain(`Id="${rid}"`);
  });

  it("riconosce tutti gli elementi", async () => {
    const { flat, prep } = await build();
    expect(prep.mermaid).toHaveLength(1);
    expect(flat).toContain("<w:tbl>");
    expect(flat).toContain("<w:hyperlink r:id=");
    expect(flat).toContain('w:anchor="h_secondo_livello"');
    expect(flat).toContain('<w:bookmarkStart w:id="2" w:name="h_secondo_livello"/>');
    expect(flat).toContain('<w:footnoteReference w:id="1"/>');
    expect(flat).toContain("Testo della nota.");
    expect(flat).toContain('w:vertAlign w:val="subscript"');
    expect(flat).toContain('w:vertAlign w:val="superscript"');
    expect(flat).toContain('<w:br w:type="page"/>');
    expect(flat).toContain("☐ </w:t>");
    expect(flat).toContain("☑ </w:t>");
    expect(flat).toContain('title="WordMD Mermaid"');
    expect(flat).toMatch(/descr="graph LR&#10;/);
    expect(flat).toContain("svgBlip");
    expect(flat).toContain("&lt;ok&gt;");
    expect(flat).toContain('<w:tblHeader/>');
    expect(flat).toContain('w:val="center"'); // colonna centrata
  });

  it("modalità nuovo stile: crea stili con il nome del profilo e le opzioni di pagina", async () => {
    const { flat } = await build(PRESETS.find((p) => p.id === "Report aziendale")!.make());
    expect(flat).toContain("WordMD Report aziendale - Heading 1");
    expect(flat).toMatch(/<w:pageBreakBefore\/>/);
    expect(flat).toContain('<w:outlineLvl w:val="0"/>');
    expect(createdStyleNames(PRESETS[3].make())).toContain("WordMD Report aziendale - Code Block");
  });

  it("modalità stili esistenti: solo riferimenti per nome, nessuna formattazione propria", async () => {
    const { flat } = await build(existingStylesProfile());
    expect(flat).toContain('<w:name w:val="heading 1"/>');
    expect(flat).toContain('w:styleId="Heading1"');
    expect(flat).not.toContain("WordMD Stili esistenti");
    expect(flat).toContain('w:styleId="Hyperlink"');
    expect(createdStyleNames(existingStylesProfile())).toEqual([]);
    // nessuno styleId duplicato nel pacchetto (Normal compariva due volte)
    const ids = [...flat.matchAll(/<w:style [^>]*w:styleId="([^"]+)"/g)].map((m) => m[1]);
    expect(ids.length).toBeGreaterThan(3);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it("un diagramma non valido diventa un blocco di codice con il sorgente", () => {
    const prep = prepare("```mermaid\ngraph LR\n  A-->\n```");
    const flat = buildPackage(prep, PRESETS[0].make(), [null], new Map());
    xmlOk(flat);
    expect(flat).toContain("Diagramma Mermaid non valido");
    expect(flat).not.toContain("svgBlip");
  });

  it("immagini mancanti non rompono il documento", async () => {
    const prep = prepare("![Mappa](immagini/mappa.png)");
    const { images, missing } = await resolveImages(prep.imageSrcs, new Map());
    expect(missing).toHaveLength(1);
    const flat = buildPackage(prep, PRESETS[0].make(), [], images);
    xmlOk(flat);
    expect(flat).toContain("immagine non disponibile");
  });
});

describe("OOXML → Markdown (andata e ritorno)", () => {
  it("ricostruisce il Markdown con diagrammi, tabelle, elenchi, link e immagini", async () => {
    for (const preset of PRESETS) {
      const { flat } = await build(preset.make());
      const out = ooxmlToMarkdown(flat, "doc");
      const md = out.markdown;
      expect(md, preset.id).toContain("# Titolo principale");
      expect(md).toContain("## Secondo livello");
      expect(md).toContain("**grassetto**");
      expect(md).toContain("*corsivo*");
      expect(md).toContain("~~barrato~~");
      expect(md).toContain("`codice`");
      expect(md).toContain("[collegamento](https://example.com/a?b=1&c=2)");
      expect(md).toContain("[sezione](#secondo-livello)");
      expect(md).toContain("- uno\n- due\n    - annidato");
      expect(md).toContain("- [ ] da fare");
      expect(md).toContain("- [x] fatto");
      expect(md).toContain("1. primo\n2. secondo");
      expect(md).toContain("> una citazione");
      expect(md).toContain("```\nconst a = 1;");
      expect(md).toContain("| Nome | Valore | Note |");
      expect(md).toContain(":---:");
      expect(md).toContain("x \\| y");
      expect(md).toContain("```mermaid\ngraph LR\n  A[Inizio] --> B{Scelta}\n  B -->|sì| C[Fine]\n```");
      expect(md).toContain("![Logo piccolo](doc-immagine-1.png)");
      expect(md).toContain("<!-- pagebreak -->");
      expect(out.diagrams).toEqual([{ code: expect.stringContaining("graph LR"), fileName: "doc-diagramma-1.svg" }]);
      expect(out.images).toHaveLength(1);
    }
  });
});

describe("pacchetti .docx / .mdocx", () => {
  const docx = () => {
    const doc = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:body><w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="D" descr="graph LR&#10;  A-->B" title="WordMD Mermaid"/><a:graphic><a:blip r:embed="rId1"><a:extLst><a:ext uri="x"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId2"/></a:ext></a:extLst></a:blip></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`;
    const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="media/a.png"/><Relationship Id="rId2" Target="media/a.svg"/></Relationships>`;
    return zipSync({
      "word/document.xml": strToU8(doc),
      "word/_rels/document.xml.rels": strToU8(rels),
      "word/media/a.png": new Uint8Array([1, 2, 3]),
      "word/media/a.svg": strToU8("<svg/>"),
    });
  };

  it("legge il sorgente e l'SVG dei diagrammi", () => {
    expect(listDiagrams(docx())).toEqual([{ code: "graph LR\n  A-->B", svg: "<svg/>" }]);
  });

  it(".docx piatto: l'immagine resta ma il sorgente Mermaid non è più recuperabile", () => {
    const flat = flattenDocx(docx());
    expect(listDiagrams(flat)).toEqual([]);
    const files = unzipSync(flat);
    expect(strFromU8(files["word/document.xml"])).not.toContain("graph LR");
    expect(files["word/media/a.svg"]).toBeDefined();
  });
});

describe("dimensioni delle immagini", () => {
  it("PNG e SVG", () => {
    expect(rasterSize(Uint8Array.from(atob(PNG_1x1), (c) => c.charCodeAt(0)))).toEqual({ width: 1, height: 1 });
    expect(svgSize('<svg viewBox="0 0 120 60"></svg>')).toEqual({ width: 120, height: 60 });
    expect(svgSize('<svg width="30px" height="20"></svg>')).toEqual({ width: 30, height: 20 });
  });
});
