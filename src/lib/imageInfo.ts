import { bytesToBase64 } from "./opc";

/** Immagine pronta per essere inserita in Word. Gli SVG portano anche un PNG di ripiego. */
export interface ImageRes {
  base64: string;
  contentType: string;
  ext: string;
  widthPx: number;
  heightPx: number;
  svgText?: string;
  pngBase64?: string;
}

const MIME_BY_EXT: Record<string, string> = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", svg: "image/svg+xml", bmp: "image/bmp",
  webp: "image/webp",
};

export function mimeFromName(name: string): string | undefined {
  return MIME_BY_EXT[name.split(/[?#]/)[0].split(".").pop()?.toLowerCase() ?? ""];
}

export function extFromMime(mime: string): string {
  return mime === "image/jpeg" ? "jpg" : mime === "image/svg+xml" ? "svg" : (mime.split("/")[1] ?? "png");
}

/** Dimensioni dai byte dell'immagine (PNG, JPEG, GIF, BMP); null se il formato non è riconosciuto. */
export function rasterSize(b: Uint8Array): { width: number; height: number } | null {
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50) return { width: dv.getUint32(16), height: dv.getUint32(20) };
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49) return { width: dv.getUint16(6, true), height: dv.getUint16(8, true) };
  if (b.length > 26 && b[0] === 0x42 && b[1] === 0x4d) return { width: Math.abs(dv.getInt32(18, true)), height: Math.abs(dv.getInt32(22, true)) };
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i + 9 < b.length) {
      if (b[i] !== 0xff) { i++; continue; }
      const marker = b[i + 1];
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { height: dv.getUint16(i + 5), width: dv.getUint16(i + 7) };
      }
      i += 2 + dv.getUint16(i + 2);
    }
  }
  return null;
}

/** Dimensioni dichiarate da un SVG (width/height oppure viewBox). */
export function svgSize(svg: string): { width: number; height: number } {
  const tag = /<svg\b[^>]*>/i.exec(svg)?.[0] ?? "";
  const num = (name: string) => {
    const m = new RegExp(`\\b${name}\\s*=\\s*["']([\\d.]+)(px)?["']`, "i").exec(tag);
    return m ? parseFloat(m[1]) : NaN;
  };
  let width = num("width");
  let height = num("height");
  const vb = /viewBox\s*=\s*["']([^"']+)["']/i.exec(tag)?.[1].split(/[\s,]+/).map(Number);
  if ((!(width > 0) || !(height > 0)) && vb && vb.length === 4) {
    width = width > 0 ? width : vb[2];
    height = height > 0 ? height : vb[3];
  }
  return { width: width > 0 ? width : 300, height: height > 0 ? height : 150 };
}

/** Dimensioni tramite il decodificatore del browser (WebP e formati non letti dai byte). */
export function decodeSize(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || 300, height: img.naturalHeight || 150 });
    img.onerror = () => reject(new Error("Immagine non decodificabile"));
    img.src = dataUrl;
  });
}

/** Costruisce un ImageRes dai byte; per gli SVG chiede il PNG di ripiego a `svgToPng`. */
export async function toImageRes(
  bytes: Uint8Array,
  contentType: string,
  svgToPng?: (svg: string, w: number, h: number) => Promise<string>,
): Promise<ImageRes> {
  if (contentType === "image/svg+xml") {
    const svgText = new TextDecoder().decode(bytes);
    const { width, height } = svgSize(svgText);
    const sizedSvg = svgText.replace(/<svg\b([^>]*)>/i, (m, attrs: string) =>
      /\bwidth\s*=/.test(attrs) && /\bheight\s*=/.test(attrs) ? m : `<svg width="${width}" height="${height}"${attrs}>`);
    const pngBase64 = svgToPng ? await svgToPng(sizedSvg, width, height) : undefined;
    return { base64: bytesToBase64(bytes), contentType, ext: "svg", widthPx: width, heightPx: height, svgText, pngBase64 };
  }
  const size = rasterSize(bytes) ?? (await decodeSize(`data:${contentType};base64,${bytesToBase64(bytes)}`));
  return { base64: bytesToBase64(bytes), contentType, ext: extFromMime(contentType), widthPx: size.width, heightPx: size.height };
}
