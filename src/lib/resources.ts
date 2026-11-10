import { mimeFromName, toImageRes, type ImageRes } from "./imageInfo";
import { svgToPng } from "./mermaidRender";

/** File scelti dall'utente accanto al Markdown (immagini): chiave = percorso relativo in minuscolo. */
export type LocalFiles = Map<string, File>;

export function addLocalFiles(into: LocalFiles, files: Iterable<File>): void {
  for (const f of files) {
    const rel = ((f as File & { webkitRelativePath?: string }).webkitRelativePath || f.name).replace(/\\/g, "/").toLowerCase();
    into.set(rel, f);
    // con la cartella scelta, "immagini/a.png" deve valere anche senza il nome della cartella radice
    const noRoot = rel.includes("/") ? rel.slice(rel.indexOf("/") + 1) : rel;
    if (!into.has(noRoot)) into.set(noRoot, f);
    const base = rel.split("/").pop()!;
    if (!into.has(base)) into.set(base, f);
  }
}

function normalizePath(src: string): string[] {
  let p = src;
  try { p = decodeURI(src); } catch { /* percorso non codificato */ }
  p = p.split(/[?#]/)[0].replace(/\\/g, "/").replace(/^\.\//, "").toLowerCase();
  const parts: string[] = [];
  for (const seg of p.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg && seg !== ".") parts.push(seg);
  }
  const full = parts.join("/");
  return [full, parts.slice(-1)[0] ?? ""];
}

function dataUri(src: string): { bytes: Uint8Array; mime: string } | null {
  const m = /^data:([^;,]+)((?:;[^;,]+)*),(.*)$/s.exec(src);
  if (!m) return null;
  const isB64 = m[2].includes(";base64");
  const raw = isB64 ? atob(decodeURIComponent(m[3])) : decodeURIComponent(m[3]);
  const bytes = Uint8Array.from(raw, (c) => c.charCodeAt(0));
  return { bytes: isB64 ? bytes : new TextEncoder().encode(raw), mime: m[1] };
}

export interface ImageResolution {
  images: Map<string, ImageRes>;
  /** Immagini non trovate o non scaricabili, con il motivo. */
  missing: { src: string; reason: string }[];
}

/**
 * Risolve gli indirizzi delle immagini del Markdown: data URI, file scelti dall'utente (percorsi relativi)
 * e indirizzi http(s) (se il server consente il download dal pannello).
 */
export async function resolveImages(srcs: string[], local: LocalFiles): Promise<ImageResolution> {
  const images = new Map<string, ImageRes>();
  const missing: ImageResolution["missing"] = [];
  for (const src of srcs) {
    try {
      let bytes: Uint8Array | undefined;
      let mime: string | undefined;
      const data = dataUri(src);
      if (data) {
        ({ bytes, mime } = data);
      } else if (/^https?:\/\//i.test(src)) {
        const res = await fetch(src);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        bytes = new Uint8Array(await res.arrayBuffer());
        mime = res.headers.get("content-type")?.split(";")[0] || mimeFromName(src);
      } else {
        const [full, base] = normalizePath(src);
        const file = local.get(full) ?? local.get(base);
        if (!file) throw new Error("file non fornito: scegli anche le immagini o la loro cartella");
        bytes = new Uint8Array(await file.arrayBuffer());
        mime = file.type || mimeFromName(file.name);
      }
      if (!mime || !mime.startsWith("image/")) mime = mimeFromName(src) ?? "image/png";
      images.set(src, await toImageRes(bytes, mime, svgToPng));
    } catch (e) {
      missing.push({ src, reason: e instanceof Error ? e.message : String(e) });
    }
  }
  return { images, missing };
}
