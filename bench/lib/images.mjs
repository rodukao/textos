// PDF → imagens e comparação de páginas pixel a pixel.
import { createCanvas, loadImage } from '@napi-rs/canvas';
import { readFile } from 'node:fs/promises';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

/** Resolução de comparação: 96 DPI, a mesma do CSS (1 px = 1/96 pol.). */
export const DPI = 96;

/** @typedef {{ width: number, height: number, data: Uint8ClampedArray }} Raster */

/** @returns {Promise<Raster[]>} */
export async function rasterizePdf(pdfPath, dpi = DPI) {
  const data = new Uint8Array(await readFile(pdfPath));
  const task = getDocument({ data, isEvalSupported: false, verbosity: 0 });
  const pdf = await task.promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: dpi / 72 });
    // Usa a tela do próprio pdf.js: ele traz sua cópia do @napi-rs/canvas, e objetos
    // (Path2D) de uma cópia não funcionam na tela da outra.
    const { canvas, context: ctx } = pdf.canvasFactory.create(Math.round(viewport.width), Math.round(viewport.height));
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvas, canvasContext: ctx, viewport }).promise;
    pages.push({ width: canvas.width, height: canvas.height, data: ctx.getImageData(0, 0, canvas.width, canvas.height).data });
  }
  await task.destroy();
  return pages;
}

/** @returns {Promise<Raster>} */
export async function decodePng(buffer) {
  const img = await loadImage(buffer);
  const canvas = createCanvas(img.width, img.height);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, img.width, img.height);
  ctx.drawImage(img, 0, 0);
  return { width: img.width, height: img.height, data: ctx.getImageData(0, 0, img.width, img.height).data };
}

/** Copia `r` para uma tela branca de `w`×`h` (alinhada no canto superior esquerdo). */
function pad(r, w, h) {
  if (r.width === w && r.height === h) return r.data;
  const out = new Uint8ClampedArray(w * h * 4).fill(255);
  for (let y = 0; y < Math.min(h, r.height); y++) {
    const row = r.data.subarray(y * r.width * 4, y * r.width * 4 + Math.min(w, r.width) * 4);
    out.set(row, y * w * 4);
  }
  return out;
}

export async function encodePng(raster) {
  const canvas = createCanvas(raster.width, raster.height);
  const ctx = canvas.getContext('2d');
  const img = ctx.createImageData(raster.width, raster.height);
  img.data.set(raster.data);
  ctx.putImageData(img, 0, 0);
  return canvas.encode('png');
}

/** Pixel com "tinta": nitidamente diferente do branco do papel. */
function inkMask(data, w, h) {
  const m = new Uint8Array(w * h);
  for (let i = 0, p = 0; p < m.length; i += 4, p++) {
    const r = data[i], g = data[i + 1], b = data[i + 2];
    // luminância baixa (texto, linhas) ou cor saturada (realces, fundos coloridos)
    const lum = 0.299 * r + 0.587 * g + 0.114 * b;
    const sat = Math.max(r, g, b) - Math.min(r, g, b);
    if (lum < 200 || sat > 40) m[p] = 1;
  }
  return m;
}

/** Dilatação quadrada de raio `r` (máximo numa janela), separável: horizontal e depois vertical. */
function dilate(mask, w, h, r) {
  const tmp = new Uint8Array(w * h);
  const out = new Uint8Array(w * h);
  for (let y = 0; y < h; y++) {
    const row = y * w;
    for (let x = 0; x < w; x++) {
      if (!mask[row + x]) continue;
      for (let k = Math.max(0, x - r); k <= Math.min(w - 1, x + r); k++) tmp[row + k] = 1;
    }
  }
  for (let x = 0; x < w; x++) {
    for (let y = 0; y < h; y++) {
      if (!tmp[y * w + x]) continue;
      for (let k = Math.max(0, y - r); k <= Math.min(h - 1, y + r); k++) out[k * w + x] = 1;
    }
  }
  return out;
}

/** Pixel bem preenchido (não é borda suavizada). */
const isCore = (d, i) => {
  const lum = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
  return lum < 130 || Math.max(d[i], d[i + 1], d[i + 2]) - Math.min(d[i], d[i + 1], d[i + 2]) > 90;
};

const colorDist = (d1, i, d2, j) => Math.hypot(d1[i] - d2[j], d1[i + 1] - d2[j + 1], d1[i + 2] - d2[j + 2]);

/**
 * Compara duas páginas tolerando o que não é diferença de verdade: suavização de
 * fonte (antialiasing) e deslocamentos de até `radius` px.
 *
 * `diff` = (tinta só na referência + tinta só no comparado + tinta com cor diferente)
 *          ÷ (tinta total). 0 = idênticas.
 *
 * Imagem de diferenças: cinza = igual, vermelho = só na referência (Word),
 * azul = só no comparado (Textos), laranja = mesma posição com cor diferente.
 * @returns {{ diff: number, diffImage: Raster, sizeMismatch: boolean }}
 */
export function comparePages(a, b, radius = 2) {
  const w = Math.max(a.width, b.width);
  const h = Math.max(a.height, b.height);
  const da = pad(a, w, h);
  const db = pad(b, w, h);
  const ia = inkMask(da, w, h);
  const ib = inkMask(db, w, h);
  const na = dilate(ia, w, h, radius);
  const nb = dilate(ib, w, h, radius);

  const out = new Uint8ClampedArray(w * h * 4).fill(255);
  const paint = (p, r, g, bl) => {
    out[p * 4] = r;
    out[p * 4 + 1] = g;
    out[p * 4 + 2] = bl;
  };

  let onlyA = 0, onlyB = 0, wrongColor = 0, ink = 0;
  for (let p = 0; p < w * h; p++) {
    if (!ia[p] && !ib[p]) continue;
    ink++;
    if (ia[p] && !nb[p]) {
      onlyA++;
      paint(p, 230, 40, 40);
    } else if (ib[p] && !na[p]) {
      onlyB++;
      paint(p, 40, 90, 230);
    } else if (ia[p] && isCore(da, p * 4)) {
      // Há tinta perto nas duas: confere a cor (só no "miolo" do traço; a borda
      // suavizada varia entre renderizadores) com o vizinho mais parecido.
      const x = p % w, y = (p - x) / w;
      let best = Infinity;
      for (let dy = -radius; dy <= radius && best > 0; dy++) {
        for (let dx = -radius; dx <= radius; dx++) {
          const xx = x + dx, yy = y + dy;
          if (xx < 0 || yy < 0 || xx >= w || yy >= h) continue;
          const q = yy * w + xx;
          if (ib[q]) best = Math.min(best, colorDist(da, p * 4, db, q * 4));
        }
      }
      if (best > 90) {
        wrongColor++;
        paint(p, 245, 150, 20);
      } else paint(p, 200, 200, 200);
    } else paint(p, 200, 200, 200);
  }

  return {
    diff: ink ? Math.min(1, (onlyA + onlyB + wrongColor) / ink) : 0,
    diffImage: { width: w, height: h, data: out },
    sizeMismatch: Math.abs(a.width - b.width) > 2 || Math.abs(a.height - b.height) > 2,
  };
}
