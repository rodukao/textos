// Imprime o XML (sem os parágrafos de texto) do wp:anchor que contém um trecho.
import { readFileSync } from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
const [file, needle] = process.argv.slice(2);
const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(file)))['word/document.xml']);
const doc = new DOMParser().parseFromString(xml, 'application/xml');
const t = [...doc.getElementsByTagName('w:t')].find((x) => x.textContent.includes(needle));
let a = t; while (a && a.nodeName !== 'mc:AlternateContent' && a.nodeName !== 'w:drawing') a = a.parentNode;
// Remove o conteúdo de texto para focar nas propriedades da caixa.
for (const c of [...a.getElementsByTagName('w:txbxContent')]) while (c.firstChild) c.removeChild(c.firstChild);
console.log(new XMLSerializer().serializeToString(a).replace(/ xmlns:\w+="[^"]*"/g, '').replace(/></g, '>\n<'));
