// Lista os parágrafos do corpo (pPr + resumo dos runs) a partir de um trecho.
import { readFileSync } from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
const [file, from, count = '3'] = process.argv.slice(2);
const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(file)))['word/document.xml']);
const doc = new DOMParser().parseFromString(xml, 'application/xml');
const body = doc.getElementsByTagName('w:body')[0];
const paras = [...body.childNodes].filter((n) => n.nodeType === 1);
const textOf = (p) => [...p.getElementsByTagName('w:t')].filter((t) => { let a = t; while (a && a !== p) { if (a.nodeName === 'w:txbxContent' || a.nodeName === 'mc:Fallback') return false; a = a.parentNode; } return true; }).map((t) => t.textContent).join('');
const i = paras.findIndex((p) => textOf(p).includes(from));
const ser = new XMLSerializer();
for (const p of paras.slice(Math.max(0, i - 1), i + Number(count))) {
  const ppr = [...p.childNodes].find((c) => c.nodeName === 'w:pPr');
  const kids = [...p.childNodes].filter((c) => c.nodeType === 1 && c.nodeName !== 'w:pPr').map((c) => c.nodeName === 'w:r' ? 'r[' + [...c.childNodes].filter((x) => x.nodeType === 1 && x.nodeName !== 'w:rPr').map((x) => x.nodeName).join(',') + ']' : c.nodeName);
  console.log(`\n¶ «${textOf(p).slice(0, 50)}»`);
  console.log('  pPr:', ppr ? ser.serializeToString(ppr).replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ w:rsid\w*="[^"]*"/g, '') : '(nenhum)');
  console.log('  filhos:', kids.join(' '));
}
