// Imprime a definição de um estilo (w:style por styleId) e os padrões do documento (w:docDefaults).
// Uso: node bench/tools/estilo-xml.mjs <arquivo.docx> <styleId>
import { readFileSync } from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
const [file, styleId] = process.argv.slice(2);
const xml = strFromU8(unzipSync(new Uint8Array(readFileSync(file)))['word/styles.xml']);
const doc = new DOMParser().parseFromString(xml, 'application/xml');
const ser = (n) => new XMLSerializer().serializeToString(n).replace(/ xmlns:\w+="[^"]*"/g, '').replace(/ w:rsid\w*="[^"]*"/g, '').replace(/></g, '>\n<');
const styles = [...doc.getElementsByTagName('w:style')];
const s = styles.find((x) => x.getAttribute('w:styleId') === styleId);
console.log(s ? ser(s) : `estilo ${styleId} não encontrado`);
const dd = doc.getElementsByTagName('w:docDefaults')[0];
console.log('--- docDefaults\n' + (dd ? ser(dd) : '(nenhum)'));
