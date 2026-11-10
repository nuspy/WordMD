import { flattenDocx, listDiagrams } from "./docxTools";
import { buildPackage, createdStyleNames, prepare } from "./mdToOoxml";
import { renderMermaid } from "./mermaidRender";
import { buildDiagramPackage } from "./ooxmlImage";
import { MERMAID_TITLE, bytesToBase64 } from "./opc";
import { ooxmlToMarkdown } from "./ooxmlToMd";
import type { MermaidFmt, Profile } from "./profile";
import { resolveImages, type LocalFiles } from "./resources";
import { base64ToBytes, type OutFile } from "./saveFiles";

export interface ImportReport {
  diagrams: number;
  failed: { index: number; message: string }[];
  missingImages: { src: string; reason: string }[];
  images: number;
}

const errText = (e: unknown) => (e instanceof Error ? e.message : String(e));

/**
 * Importa il Markdown nel documento applicando il profilo di stile.
 * Il rendering di diagrammi e immagini avviene prima di toccare il documento: se fallisce, nulla viene modificato.
 */
export async function importMarkdown(
  markdown: string,
  profile: Profile,
  local: LocalFiles,
  replaceAll: boolean,
): Promise<ImportReport> {
  const prep = prepare(markdown);
  const failed: ImportReport["failed"] = [];
  const rendered = [];
  for (const [i, code] of prep.mermaid.entries()) {
    try {
      rendered.push(await renderMermaid(code, profile.mermaid));
    } catch (e) {
      rendered.push(null);
      failed.push({ index: i + 1, message: errText(e) });
    }
  }
  const { images, missing } = await resolveImages(prep.imageSrcs, local);
  const pkg = buildPackage(prep, profile, rendered, images);

  await Word.run(async (ctx) => {
    // Gli stili con lo stesso nome non vengono sovrascritti da Word: li rimuovo per applicare le impostazioni attuali.
    for (const name of createdStyleNames(profile)) {
      const style = ctx.document.getStyles().getByNameOrNullObject(name);
      style.load("isNullObject");
      await ctx.sync();
      if (!style.isNullObject) {
        try {
          style.delete();
          await ctx.sync();
        } catch {
          /* stile in uso e non eliminabile: resta la definizione precedente */
        }
      }
    }
    if (replaceAll) {
      ctx.document.body.clear();
      ctx.document.body.insertOoxml(pkg, "Start");
    } else {
      ctx.document.getSelection().insertOoxml(pkg, "Replace");
    }
    await ctx.sync();
  });

  return { diagrams: prep.mermaid.length, failed, missingImages: missing, images: images.size };
}

export async function insertDiagram(code: string, fmt: MermaidFmt): Promise<void> {
  const pkg = buildDiagramPackage(code, await renderMermaid(code, fmt), "Diagramma", fmt);
  await Word.run(async (ctx) => {
    ctx.document.getSelection().insertOoxml(pkg, "Replace");
    await ctx.sync();
  });
}

/** Sorgente Mermaid dell'immagine selezionata, oppure null se la selezione non è un diagramma WordMD. */
export async function selectedDiagramCode(): Promise<string | null> {
  return Word.run(async (ctx) => {
    const pics = ctx.document.getSelection().inlinePictures;
    pics.load("items/altTextTitle,items/altTextDescription");
    await ctx.sync();
    const pic = pics.items.find((p) => p.altTextTitle === MERMAID_TITLE);
    return pic ? pic.altTextDescription : null;
  });
}

export async function replaceSelectedDiagram(code: string, fmt: MermaidFmt): Promise<boolean> {
  const pkg = buildDiagramPackage(code, await renderMermaid(code, fmt), "Diagramma", fmt);
  return Word.run(async (ctx) => {
    const pics = ctx.document.getSelection().inlinePictures;
    pics.load("items/altTextTitle");
    await ctx.sync();
    const pic = pics.items.find((p) => p.altTextTitle === MERMAID_TITLE);
    if (!pic) return false;
    pic.getRange().insertOoxml(pkg, "Replace");
    await ctx.sync();
    return true;
  });
}

// ---------- export ----------

export interface ExportResult {
  files: OutFile[];
  diagrams: number;
  failed: { fileName: string; message: string }[];
}

/** Legge il documento e prepara <nome>.md più un SVG per ogni diagramma Mermaid e i file delle immagini. */
export async function exportMarkdown(baseName: string, fmt: MermaidFmt): Promise<ExportResult> {
  const flat = await Word.run(async (ctx) => {
    const ooxml = ctx.document.body.getOoxml();
    await ctx.sync();
    return ooxml.value;
  });
  const out = ooxmlToMarkdown(flat, baseName);

  const files: OutFile[] = [{ name: `${baseName}.md`, data: out.markdown, type: "text/markdown;charset=utf-8" }];
  const failed: ExportResult["failed"] = [];
  for (const d of out.diagrams) {
    try {
      files.push({ name: d.fileName, data: (await renderMermaid(d.code, fmt)).svg, type: "image/svg+xml" });
    } catch (e) {
      failed.push({ fileName: d.fileName, message: errText(e) });
    }
  }
  for (const img of out.images) {
    files.push({ name: img.fileName, data: base64ToBytes(img.base64), type: "application/octet-stream" });
  }
  return { files, diagrams: out.diagrams.length, failed };
}

/** Il documento intero come file .docx/.mdocx (Word lo produce già con le immagini incorporate). */
export function currentDocumentBytes(): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(res.error.message));
      const file = res.value;
      const chunks: Uint8Array[] = [];
      const next = (i: number): void => {
        file.getSliceAsync(i, (s) => {
          if (s.status !== Office.AsyncResultStatus.Succeeded) {
            file.closeAsync();
            return reject(new Error(s.error.message));
          }
          chunks.push(Uint8Array.from(s.value.data as number[]));
          if (i + 1 < file.sliceCount) return next(i + 1);
          file.closeAsync();
          const all = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
          let off = 0;
          for (const c of chunks) { all.set(c, off); off += c.length; }
          resolve(all);
        });
      };
      next(0);
    });
  });
}

/**
 * .docx: immagini fisse (PNG + SVG) senza sorgente Mermaid, più i file SVG dei diagrammi.
 * .mdocx: lo stesso pacchetto con il sorgente Mermaid nel testo alternativo, quindi i diagrammi restano modificabili.
 */
export async function exportWord(baseName: string, editable: boolean): Promise<ExportResult> {
  const bytes = await currentDocumentBytes();
  const diagrams = listDiagrams(bytes);
  const files: OutFile[] = [{
    name: `${baseName}.${editable ? "mdocx" : "docx"}`,
    data: editable ? bytes : flattenDocx(bytes),
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  }];
  const failed: ExportResult["failed"] = [];
  if (!editable) {
    diagrams.forEach((d, i) => {
      const fileName = `${baseName}-diagramma-${i + 1}.svg`;
      if (d.svg) files.push({ name: fileName, data: d.svg, type: "image/svg+xml" });
      else failed.push({ fileName, message: "SVG non presente nel documento" });
    });
  }
  return { files, diagrams: diagrams.length, failed };
}

/** Apre un file .mdocx (o .docx) in una nuova finestra di Word: i diagrammi restano modificabili dal pannello. */
export async function openWordFile(bytes: Uint8Array): Promise<number> {
  const count = listDiagrams(bytes).length;
  await Word.run(async (ctx) => {
    ctx.application.createDocument(bytesToBase64(bytes)).open();
    await ctx.sync();
  });
  return count;
}
