// Mostra a estrutura XML (só nomes de elementos) ao redor de um trecho de texto num .docx.
import { readFileSync } from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';
const [file, needle, part = 'word/document.xml'] = process.argv.slice(2);
const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(file)))[part]);
const doc = new DOMParser().parseFromString(xml, 'application/xml');
// Acha o w:t com o trecho e sobe até o w:p de nível mais alto do corpo.
const ts = [...doc.getElementsByTagName('w:t')].filter((t) => t.textContent.includes(needle));
if (!ts.length) { console.log('trecho não encontrado'); process.exit(); }
let top = ts[0]; const chain = [];
while (top.parentNode && top.parentNode.nodeName !== 'w:body') { top = top.parentNode; chain.push(top.nodeName); }
console.log('caminho até o corpo:', chain.reverse().join(' > '));
const outline = (el, depth = 0, max = 14) => {
  if (depth > max) return;
  const attrs = [...(el.attributes ?? [])].filter((a) => !/rsid|paraId|textId|xmlns/.test(a.name)).map((a) => `${a.name}="${a.value.slice(0, 30)}"`).join(' ');
  const txt = el.nodeName === 'w:t' ? ` «${el.textContent.slice(0, 40)}»` : '';
  console.log('  '.repeat(depth) + el.nodeName + (attrs ? ' ' + attrs : '') + txt);
  for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) outline(c, depth + 1, max);
};
outline(top);
