// Imprime o XML dos parágrafos dentro da caixa de texto que contém um trecho.
import { readFileSync } from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
const [file, needle] = process.argv.slice(2);
const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(file)))['word/document.xml']);
const doc = new DOMParser().parseFromString(xml, 'application/xml');
const t = [...doc.getElementsByTagName('w:t')].find((x) => x.textContent.includes(needle));
let c = t; while (c && c.nodeName !== 'w:txbxContent') c = c.parentNode;
console.log(new XMLSerializer().serializeToString(c).replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ w:rsid\w*="[^"]*"/g, '').replace(/></g, '>\n<'));
