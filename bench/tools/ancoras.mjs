// Lista os parágrafos do corpo com os objetos flutuantes ancorados em cada um (tamanho, posição,
// tipo de contorno e o texto de caixas de texto). Ajuda a entender quebras de linha ao lado de
// imagens e caixas flutuantes.
// Uso: node bench/tools/ancoras.mjs <arquivo.docx>
import { readFileSync } from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';

const [file] = process.argv.slice(2);
const zip = unzipSync(new Uint8Array(readFileSync(file)));
const doc = new DOMParser().parseFromString(strFromU8(zip['word/document.xml']), 'application/xml');
const inside = (el, name) => { for (let p = el.parentNode; p; p = p.parentNode) if (p.nodeName === name) return true; return false; };
const pt = (emu) => (Number(emu) / 12700).toFixed(1);
const pos = (el) => {
  const off = el.getElementsByTagName('wp:posOffset')[0];
  const align = el.getElementsByTagName('wp:align')[0];
  return `${el.getAttribute('relativeFrom')} ${off ? pt(off.textContent) + 'pt' : align?.textContent}`;
};

const body = doc.getElementsByTagName('w:body')[0];
let n = 0;
for (let el = body.firstChild; el; el = el.nextSibling) {
  if (el.nodeName !== 'w:p' && el.nodeName !== 'w:tbl') continue;
  n++;
  const text = [...el.getElementsByTagName('w:t')].filter((t) => !inside(t, 'w:txbxContent')).map((t) => t.textContent).join('');
  const anchors = [...el.getElementsByTagName('wp:anchor')].filter((a) => !inside(a, 'mc:Fallback')).map((a) => {
    const ext = a.getElementsByTagName('wp:extent')[0];
    const wrap = [...a.childNodes].find((c) => c.nodeName?.startsWith('wp:wrap'));
    const box = [...a.getElementsByTagName('w:t')].map((t) => t.textContent).join('');
    return `\n     [${pt(ext.getAttribute('cx'))} × ${pt(ext.getAttribute('cy'))} pt, H ${pos(a.getElementsByTagName('wp:positionH')[0])}, ` +
      `V ${pos(a.getElementsByTagName('wp:positionV')[0])}, ${wrap?.nodeName.slice(3)}${wrap?.getAttribute('wrapText') ? ' ' + wrap.getAttribute('wrapText') : ''}` +
      `${box ? `, caixa: "${box.slice(0, 50)}"` : ''}]`;
  });
  console.log(`${String(n).padStart(3)} ${el.nodeName === 'w:tbl' ? 'tabela' : ''}${JSON.stringify(text.slice(0, 70))}${anchors.join('')}`);
}
