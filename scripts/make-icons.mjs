// Genera le icone PNG del pannello (16, 32, 80 px) senza dipendenze: quadrato arrotondato blu con una "M" bianca.
import { writeFileSync, mkdirSync } from "node:fs";
import { deflateSync } from "node:zlib";

const BG = [0x1f, 0x4e, 0x79];
const FG = [255, 255, 255];

function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}

// distanza di un punto da un segmento
function seg(px, py, ax, ay, bx, by) {
  const dx = bx - ax, dy = by - ay;
  const t = Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)));
  return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
}

function pixel(x, y, s) {
  const u = x / s, v = y / s; // 0..1
  const r = 0.18;
  const cx = Math.min(Math.max(u, r), 1 - r), cy = Math.min(Math.max(v, r), 1 - r);
  if (Math.hypot(u - cx, v - cy) > r) return [0, 0, 0, 0];
  const w = 0.075;
  const d = Math.min(
    seg(u, v, 0.26, 0.74, 0.26, 0.26), seg(u, v, 0.74, 0.74, 0.74, 0.26),
    seg(u, v, 0.26, 0.26, 0.5, 0.56), seg(u, v, 0.74, 0.26, 0.5, 0.56),
  );
  const a = Math.max(0, Math.min(1, (w - d) / (1 / s) + 0.5));
  return [0, 1, 2].map((i) => Math.round(BG[i] * (1 - a) + FG[i] * a)).concat(255);
}

function png(s) {
  const SS = 4; // supersampling
  const raw = Buffer.alloc((s * 4 + 1) * s);
  for (let y = 0; y < s; y++) {
    raw[y * (s * 4 + 1)] = 0;
    for (let x = 0; x < s; x++) {
      const acc = [0, 0, 0, 0];
      for (let j = 0; j < SS; j++) for (let i = 0; i < SS; i++) {
        const p = pixel(x + (i + 0.5) / SS, y + (j + 0.5) / SS, s);
        const al = p[3] / 255;
        acc[0] += p[0] * al; acc[1] += p[1] * al; acc[2] += p[2] * al; acc[3] += p[3];
      }
      const n = SS * SS, al = acc[3] / n / 255;
      const o = y * (s * 4 + 1) + 1 + x * 4;
      raw[o] = al ? Math.round(acc[0] / n / al) : 0; raw[o + 1] = al ? Math.round(acc[1] / n / al) : 0;
      raw[o + 2] = al ? Math.round(acc[2] / n / al) : 0; raw[o + 3] = Math.round(acc[3] / n);
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(s, 0); ihdr.writeUInt32BE(s, 4); ihdr[8] = 8; ihdr[9] = 6;
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk("IHDR", ihdr), chunk("IDAT", deflateSync(raw)), chunk("IEND", Buffer.alloc(0))]);
}

mkdirSync("assets", { recursive: true });
for (const s of [16, 32, 80]) writeFileSync(`assets/icon-${s}.png`, png(s));
console.log("icone create in assets/");
