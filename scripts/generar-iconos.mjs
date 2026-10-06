// Genera los íconos PNG del complemento en public/assets/ sin dependencias externas.
// Uso: node scripts/generar-iconos.mjs
import { mkdirSync, writeFileSync } from "node:fs";
import { deflateSync } from "node:zlib";

const SIZES = [16, 32, 64, 80, 128];
const SS = 4; // supermuestreo para suavizar bordes

const GREEN = [16, 124, 65];
const WHITE = [255, 255, 255];
const GOLD = [255, 200, 61];

function crcTable() {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
}
const CRC = crcTable();
function crc32(buf) {
  let c = 0xffffffff;
  for (const b of buf) c = CRC[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}
function png(size, rgba) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bits por canal
  ihdr[9] = 6; // RGBA
  const raw = Buffer.alloc((size * 4 + 1) * size);
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    rgba.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

// Devuelve el color del punto (u, v) en coordenadas normalizadas [0, 1], o null si es transparente.
function sample(u, v) {
  const r = 0.2;
  const inRounded = (() => {
    const cx = Math.min(Math.max(u, r), 1 - r);
    const cy = Math.min(Math.max(v, r), 1 - r);
    return (u - cx) ** 2 + (v - cy) ** 2 <= r * r;
  })();
  if (!inRounded) return null;

  // Destello (estrella de 4 puntas) arriba a la derecha.
  const sx = (u - 0.7) / 0.24;
  const sy = (v - 0.3) / 0.24;
  if (Math.abs(sx) ** (2 / 3) + Math.abs(sy) ** (2 / 3) <= 1) return GOLD;

  // Cuadrícula de hoja de cálculo.
  const line = 0.055;
  const inGrid = u > 0.18 && u < 0.82 && v > 0.18 && v < 0.82;
  if (inGrid) {
    for (const g of [0.18, 0.39, 0.6, 0.82]) {
      if (Math.abs(u - g) < line / 2 || Math.abs(v - g) < line / 2) return WHITE;
    }
  }
  return GREEN;
}

mkdirSync("public/assets", { recursive: true });
for (const size of SIZES) {
  const rgba = Buffer.alloc(size * size * 4);
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let r = 0, g = 0, b = 0, a = 0;
      for (let j = 0; j < SS; j++) {
        for (let i = 0; i < SS; i++) {
          const c = sample((x + (i + 0.5) / SS) / size, (y + (j + 0.5) / SS) / size);
          if (c) {
            r += c[0];
            g += c[1];
            b += c[2];
            a += 1;
          }
        }
      }
      const o = (y * size + x) * 4;
      if (a > 0) {
        rgba[o] = Math.round(r / a);
        rgba[o + 1] = Math.round(g / a);
        rgba[o + 2] = Math.round(b / a);
      }
      rgba[o + 3] = Math.round((a / (SS * SS)) * 255);
    }
  }
  writeFileSync(`public/assets/icon-${size}.png`, png(size, rgba));
  console.log(`public/assets/icon-${size}.png`);
}
