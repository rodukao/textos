// Imprime a altura de linha simples do Word para fontes TrueType/OpenType: as métricas do GDI
// (usWinAscent + usWinDescent + entrelinha externa) divididas por unitsPerEm. É a tabela usada pelo
// motor em packages/layout-engine/measuring/dom/src/wordLineMetrics.ts.
// Uso: node bench/tools/metricas-fonte.mjs C:\Windows\Fonts\calibri.ttf C:\Windows\Fonts\cambria.ttc ...
import { readFileSync } from 'node:fs';

const tables = (buf, off) => {
  const t = {};
  for (let i = 0; i < buf.readUInt16BE(off + 4); i++) {
    const r = off + 12 + i * 16;
    t[buf.toString('latin1', r, r + 4)] = buf.readUInt32BE(r + 8);
  }
  return t;
};
const names = (buf, t) => {
  const o = t.name, so = o + buf.readUInt16BE(o + 4), n = {};
  for (let i = 0; i < buf.readUInt16BE(o + 2); i++) {
    const r = o + 6 + i * 12;
    const [pid, lid, nid, len, off] = [0, 4, 6, 8, 10].map((k) => buf.readUInt16BE(r + k));
    if (pid === 3 && lid === 0x409 && (nid === 1 || nid === 2)) n[nid] = Buffer.from(buf.subarray(so + off, so + off + len)).swap16().toString('utf16le');
  }
  return n;
};

for (const file of process.argv.slice(2)) {
  const buf = readFileSync(file);
  const offsets = buf.toString('latin1', 0, 4) === 'ttcf'
    ? Array.from({ length: buf.readUInt32BE(8) }, (_, i) => buf.readUInt32BE(12 + i * 4))
    : [0];
  for (const off of offsets) {
    const t = tables(buf, off), n = names(buf, t);
    const upm = buf.readUInt16BE(t.head + 18);
    const [hA, hD, hG] = [4, 6, 8].map((k) => buf.readInt16BE(t.hhea + k));
    const [wA, wD] = [74, 76].map((k) => buf.readUInt16BE(t['OS/2'] + k));
    const externa = Math.max(0, hG - ((wA + wD) - (hA - hD)));
    const razao = (wA + wD + externa) / upm;
    console.log(`${`${n[1]} ${n[2]}`.padEnd(32)} simples = ${razao.toFixed(5)} × tamanho  (win ${wA}+${wD}, entrelinha externa ${externa}, upm ${upm})`);
  }
}
