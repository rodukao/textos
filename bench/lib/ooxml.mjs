// Compara o conteúdo OOXML de dois .docx e lista o que se perdeu no caminho.
import { DOMParser } from '@xmldom/xmldom';
import { createHash } from 'node:crypto';
import { strFromU8, unzipSync } from 'fflate';

const NS = {
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main': 'w',
  'http://schemas.microsoft.com/office/word/2010/wordml': 'w14',
  'http://schemas.microsoft.com/office/word/2012/wordml': 'w15',
  'http://schemas.microsoft.com/office/word/2018/wordml/cex': 'w16cex',
  'http://schemas.microsoft.com/office/word/2016/wordml/cid': 'w16cid',
  'http://schemas.microsoft.com/office/word/2018/wordml': 'w16',
  'http://schemas.microsoft.com/office/word/2023/wordml/word16du': 'w16du',
  'http://schemas.microsoft.com/office/word/2015/wordml/symex': 'w16se',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing': 'wp',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing': 'wp14',
  'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
  'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape': 'wps',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup': 'wpg',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'urn:schemas-microsoft-com:vml': 'v',
  'urn:schemas-microsoft-com:office:office': 'o',
  'urn:schemas-microsoft-com:office:word': 'w10',
  'http://schemas.openxmlformats.org/officeDocument/2006/math': 'm',
  'http://schemas.openxmlformats.org/package/2006/relationships': 'rel',
  'http://schemas.openxmlformats.org/package/2006/content-types': 'ct',
  'http://www.w3.org/XML/1998/namespace': 'xml',
};

/**
 * Atributos que mudam a cada salvamento sem afetar a formatação
 * (identificadores de sessão de edição, ids internos de parágrafo, datas).
 */
const VOLATILE_ATTR = /^(w:rsid\w*|w14:paraId|w14:textId|w:id|w16du:dateUtc|r:id|r:embed|r:link|w:date)$/;
const VOLATILE_PARTS = /^(docProps\/|word\/webSettings\.xml$|customXml\/)/;

const qname = (node) => {
  const prefix = NS[node.namespaceURI] ?? node.namespaceURI ?? '';
  return prefix ? `${prefix}:${node.localName}` : node.localName;
};

/** Assinaturas de um XML: `pai>filho` e `elemento@atributo=valor`, com contagem. */
function signatures(xml) {
  const doc = new DOMParser({ onError: () => {} }).parseFromString(xml, 'application/xml');
  const counts = new Map();
  const add = (k) => counts.set(k, (counts.get(k) ?? 0) + 1);
  const walk = (el, parentName) => {
    const name = qname(el);
    add(`<${parentName}> ${name}`);
    for (let i = 0; i < el.attributes.length; i++) {
      const at = el.attributes[i];
      if (at.namespaceURI === 'http://www.w3.org/2000/xmlns/' || at.name.startsWith('xmlns')) continue;
      const an = qname(at);
      if (VOLATILE_ATTR.test(an)) continue;
      add(`${name} @${an}=${at.value}`);
    }
    for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) walk(c, name);
  };
  if (doc.documentElement) walk(doc.documentElement, '#');
  return counts;
}

/** Texto visível de uma parte (w:t, w:delText, w:tab, w:br). */
function plainText(xml) {
  const doc = new DOMParser({ onError: () => {} }).parseFromString(xml, 'application/xml');
  let out = '';
  const walk = (el) => {
    const name = qname(el);
    // Separadores de nota (w:type="separator" etc.) são decoração gerada pelo Word, não texto.
    if ((name === 'w:footnote' || name === 'w:endnote') && el.getAttributeNS(Object.keys(NS)[0], 'type')) return;
    // mc:Fallback é uma cópia (VML) do mesmo conteúdo do mc:Choice, só para Word 2007;
    // contá-la duplicaria o texto de caixas de texto. A perda dela aparece na lista de XML.
    if (name === 'mc:Fallback') return;
    if (name === 'w:t' || name === 'w:delText') out += el.textContent;
    else if (name === 'w:tab') out += '\t';
    else if (name === 'w:br' || name === 'w:cr') out += '\n';
    else if (name === 'w:p') {
      for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) walk(c);
      out += '\n';
      return;
    }
    for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) walk(c);
  };
  if (doc.documentElement) walk(doc.documentElement);
  return out;
}

const sha = (u8) => createHash('sha1').update(u8).digest('hex');
const isXml = (name) => /\.(xml|rels)$/i.test(name);

/** Agrupa partes numeradas (header1.xml, header2.xml…) para comparar o conjunto. */
const partGroup = (name) => name.replace(/\d+(?=\.xml(\.rels)?$)/, 'N');

/**
 * @returns {{
 *   textEqual: boolean, textDiffAt: number | null,
 *   missingParts: string[], addedParts: string[], changedMedia: string[],
 *   lost: { part: string, signature: string, count: number }[],
 *   lostTotal: number
 * }}
 */
export function diffDocx(originalBytes, roundtripBytes) {
  const a = unzipSync(new Uint8Array(originalBytes));
  const b = unzipSync(new Uint8Array(roundtripBytes));

  const groupFiles = (zip) => {
    const g = new Map();
    for (const name of Object.keys(zip)) {
      if (name.endsWith('/')) continue;
      const k = partGroup(name);
      if (!g.has(k)) g.set(k, []);
      g.get(k).push(name);
    }
    return g;
  };
  const ga = groupFiles(a);
  const gb = groupFiles(b);

  const missingParts = [...ga.keys()].filter((k) => !gb.has(k) && !VOLATILE_PARTS.test(k));
  const addedParts = [...gb.keys()].filter((k) => !ga.has(k));

  // Mídia: compara pelo conteúdo (hash), independentemente do nome do arquivo.
  const mediaHashes = (zip) => new Set(Object.keys(zip).filter((n) => n.startsWith('word/media/')).map((n) => sha(zip[n])));
  const mb = mediaHashes(b);
  const changedMedia = Object.keys(a).filter((n) => n.startsWith('word/media/') && !mb.has(sha(a[n])));

  const lost = [];
  for (const [group, files] of ga) {
    if (!isXml(group) || VOLATILE_PARTS.test(group) || !gb.has(group)) continue;
    const sum = (zip, names) => {
      const total = new Map();
      for (const n of names) for (const [k, v] of signatures(strFromU8(zip[n]))) total.set(k, (total.get(k) ?? 0) + v);
      return total;
    };
    const sa = sum(a, files);
    const sb = sum(b, gb.get(group));
    for (const [sig, n] of sa) {
      const m = sb.get(sig) ?? 0;
      if (m < n) lost.push({ part: group, signature: sig, count: n - m });
    }
  }
  lost.sort((x, y) => y.count - x.count);

  const textOf = (zip) =>
    ['word/document.xml', ...Object.keys(zip).filter((n) => /^word\/(header|footer|footnotes|endnotes)\d*\.xml$/.test(n)).sort()]
      .filter((n) => zip[n])
      .map((n) => plainText(strFromU8(zip[n])))
      .join('\n');
  // Linhas vazias não contam: o Word cria cabeçalhos/rodapés vazios que outros editores
  // omitem sem mudar nada na página.
  const ta = textOf(a).replace(/\n{2,}/g, '\n').trim();
  const tb = textOf(b).replace(/\n{2,}/g, '\n').trim();
  let textDiffAt = null;
  if (ta !== tb) {
    textDiffAt = 0;
    while (textDiffAt < ta.length && ta[textDiffAt] === tb[textDiffAt]) textDiffAt++;
  }

  return {
    textEqual: ta === tb,
    textDiffAt,
    textContext: textDiffAt === null ? null : { original: ta.slice(Math.max(0, textDiffAt - 30), textDiffAt + 40), roundtrip: tb.slice(Math.max(0, textDiffAt - 30), textDiffAt + 40) },
    missingParts,
    addedParts,
    changedMedia,
    lost,
    lostTotal: lost.reduce((s, l) => s + l.count, 0),
  };
}
