// Lista as linhas de texto de uma página de um PDF (por exemplo, o PDF que o Word gera) com a
// posição e a largura de cada uma, em pt. Para comparar onde o Word quebra as linhas.
// Uso: node bench/tools/linhas-pdf.mjs <arquivo.pdf> <página> [trecho]
//   com [trecho], mostra só as linhas que o contêm (o texto de fontes como Cambria pode vir trocado no PDF).
import { readFileSync } from 'node:fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
const [file, pageNo, filter = ''] = process.argv.slice(2);
const pdf = await getDocument({ data: new Uint8Array(readFileSync(file)), verbosity: 0 }).promise;
const page = await pdf.getPage(Number(pageNo));
const { items } = await page.getTextContent();
const lines = new Map();
for (const it of items) { if (!it.str.trim()) continue; const k = it.transform[5].toFixed(1); (lines.get(k) ?? lines.set(k, []).get(k)).push(it); }
for (const [y, its] of [...lines].sort((a, b) => b[0] - a[0])) {
  its.sort((a, b) => a.transform[4] - b.transform[4]);
  const x0 = its[0].transform[4], x1 = its.at(-1).transform[4] + its.at(-1).width;
  const t = its.map((i) => i.str).join(' ');
  if (t.includes(filter)) console.log(y.padStart(6), '| x', x0.toFixed(1), '→', x1.toFixed(1), '| larg.', (x1 - x0).toFixed(2), '|', t.slice(0, 95));
}
