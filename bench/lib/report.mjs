// Gera bench/out/report.html a partir dos resultados.
import { writeFile } from 'node:fs/promises';
import path from 'node:path';
import { SLOW_CPU } from './superdoc.mjs';

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const maxDiff = (pages) => (pages?.length ? Math.max(...pages.map((p) => (p.diff == null ? 100 : p.diff))) : null);
const fmtSec = (ms) => (ms == null ? '—' : `${(ms / 1000).toLocaleString('pt-BR', { maximumFractionDigits: 1 })} s`);
const fmtPct = (x) => (x == null ? '—' : `${x.toLocaleString('pt-BR', { maximumFractionDigits: 2 })}%`);

/** Veredito de "salvar sem perda". */
function saveVerdict(d) {
  if (!d.roundtripPath) return { cls: 'bad', label: 'Não exportou' };
  if (!d.word_roundtrip?.ok) return { cls: 'bad', label: 'Word recusou o arquivo' };
  const xml = d.result.save.xml;
  const m = maxDiff(d.result.save.pages);
  const samePages = d.word_roundtrip.pages === d.word_original?.pages;
  if (!xml.textEqual) return { cls: 'bad', label: 'Texto alterado' };
  if (samePages && m === 0) return { cls: 'good', label: 'Idêntico no Word' };
  if (samePages && m < 1) return { cls: 'warn', label: 'Diferença pequena' };
  return { cls: 'bad', label: 'Formatação perdida' };
}

function displayVerdict(d) {
  const pages = d.result.display.pages;
  if (!pages?.length) return { cls: 'bad', label: 'Não exibiu' };
  const m = maxDiff(pages);
  const samePages = d.superdoc.pages === d.word_original?.pages;
  if (!samePages) return { cls: 'bad', label: 'Paginação diferente' };
  if (m < 5) return { cls: 'good', label: 'Muito próximo' };
  if (m < 25) return { cls: 'warn', label: 'Diferenças pequenas' };
  return { cls: 'bad', label: 'Diferenças grandes' };
}

function pageRows(d, kind) {
  const pages = d.result[kind].pages ?? [];
  const left = d.wordPages ?? [];
  const right = kind === 'display' ? d.sdShots ?? [] : d.roundtripPages ?? [];
  const rightLabel = kind === 'display' ? 'SuperDoc' : 'Word abrindo o arquivo salvo';
  return pages
    .map((p) => {
      const img = (f) => (f ? `<a href="${esc(d.id)}/${esc(f)}" target="_blank"><img loading="lazy" src="${esc(d.id)}/${esc(f)}"></a>` : '<div class="noimg">—</div>');
      return `<div class="pagerow">
        <div class="pagehead">Página ${p.page} · diferença ${fmtPct(p.diff)}${p.sizeMismatch ? ' · tamanho de página diferente' : ''}${p.missing ? ` · falta no ${p.missing}` : ''}</div>
        <div class="triple">
          <figure>${img(left[p.page - 1])}<figcaption>Word (original)</figcaption></figure>
          <figure>${img(right[p.page - 1])}<figcaption>${rightLabel}</figcaption></figure>
          <figure>${img(p.diffImage)}<figcaption>Diferenças (vermelho)</figcaption></figure>
        </div></div>`;
    })
    .join('');
}

function xmlSection(xml) {
  if (!xml) return '';
  const items = xml.lost
    .slice(0, 60)
    .map((l) => `<tr><td>${esc(l.part)}</td><td><code>${esc(l.signature)}</code></td><td class="num">${l.count}</td></tr>`)
    .join('');
  return `
    <p>Texto: ${xml.textEqual ? '<b class="good-t">idêntico</b>' : `<b class="bad-t">diferente</b> — original: <code>${esc(xml.textContext?.original)}</code> · salvo: <code>${esc(xml.textContext?.roundtrip)}</code>`}</p>
    ${xml.missingParts.length ? `<p>Partes do arquivo que sumiram: <code>${xml.missingParts.map(esc).join('</code>, <code>')}</code></p>` : ''}
    ${xml.changedMedia.length ? `<p>Imagens alteradas/perdidas: ${xml.changedMedia.length}</p>` : ''}
    <p>${xml.lostTotal} item(ns) de XML do original não aparecem no arquivo salvo${xml.lost.length > 60 ? ' (mostrando os 60 mais frequentes)' : ''}:</p>
    ${items ? `<table class="xml"><thead><tr><th>Parte</th><th>Item perdido</th><th>Qtd.</th></tr></thead><tbody>${items}</tbody></table>` : ''}`;
}

export async function writeReport(outDir, docs) {
  const rows = docs
    .map((d) => {
      const sv = saveVerdict(d);
      const dv = displayVerdict(d);
      return `<tr>
        <td><a href="#${esc(d.id)}">${esc(path.basename(d.file))}</a></td>
        <td class="num">${d.word_original?.pages ?? '—'}</td>
        <td><span class="badge ${sv.cls}">${sv.label}</span></td>
        <td class="num">${fmtPct(maxDiff(d.result.save.pages))}</td>
        <td class="num">${d.result.save.xml?.lostTotal ?? '—'}</td>
        <td><span class="badge ${dv.cls}">${dv.label}</span></td>
        <td class="num">${d.superdoc.pages ?? '—'}</td>
        <td class="num">${fmtPct(maxDiff(d.result.display.pages))}</td>
        <td class="num">${fmtSec(d.superdoc.perf?.slowOpenMs)}</td>
        <td class="num">${d.superdoc.perf?.heapMB != null ? `${d.superdoc.perf.heapMB} MB` : '—'}</td>
      </tr>`;
    })
    .join('');

  const count = (fn, cls) => docs.filter((d) => fn(d).cls === cls).length;
  const details = docs
    .map((d) => {
      const problems = [
        ...(d.superdoc.errors ?? []),
        d.superdoc.exportError,
        d.word_original?.error && `Word (original): ${d.word_original.error}`,
        d.word_roundtrip?.error && `Word (arquivo salvo): ${d.word_roundtrip.error}`,
        d.superdoc.blockedHosts?.length && `Acessos externos bloqueados: ${d.superdoc.blockedHosts.join(', ')}`,
      ].filter(Boolean);
      return `<section id="${esc(d.id)}">
        <h2>${esc(path.basename(d.file))}</h2>
        <p class="muted">${esc(d.file)} · abriu em ${fmtSec(d.superdoc.openMs)} (nesta máquina) e ${fmtSec(d.superdoc.perf?.slowOpenMs)} com CPU ${SLOW_CPU}× mais lenta · memória ${d.superdoc.perf?.heapMB ?? '—'} MB</p>
        ${problems.length ? `<ul class="problems">${problems.map((p) => `<li>${esc(p)}</li>`).join('')}</ul>` : ''}
        <details open><summary>Salvar sem perda — ${saveVerdict(d).label}</summary>
          ${xmlSection(d.result.save.xml)}
          ${pageRows(d, 'save')}
        </details>
        <details><summary>Exibição na tela — ${displayVerdict(d).label}</summary>${pageRows(d, 'display')}</details>
      </section>`;
    })
    .join('');

  const html = `<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Fidelidade Word — Textos</title>
<style>
  :root { --bg:#f6f7f9; --card:#fff; --text:#1f2328; --muted:#5f6670; --line:#d9dde3; --good:#1a7f37; --warn:#9a6700; --bad:#cf222e; color-scheme: light; }
  @media (prefers-color-scheme: dark) { :root { --bg:#16181c; --card:#1f2228; --text:#e6e8eb; --muted:#9aa1a9; --line:#343840; --good:#4ac26b; --warn:#d4a72c; --bad:#ff6b6b; color-scheme: dark; } }
  body { margin:0; background:var(--bg); color:var(--text); font:14px/1.5 system-ui,'Segoe UI',sans-serif; }
  main { max-width:1200px; margin:0 auto; padding:24px 16px 64px; }
  h1 { font-size:22px; margin:0 0 4px; } h2 { font-size:17px; margin:0 0 4px; }
  .muted { color:var(--muted); margin-top:0; }
  .cards { display:flex; gap:12px; flex-wrap:wrap; margin:16px 0 24px; }
  .card { background:var(--card); border:1px solid var(--line); border-radius:8px; padding:12px 16px; min-width:200px; }
  .card b { font-size:20px; }
  table { border-collapse:collapse; width:100%; background:var(--card); border:1px solid var(--line); border-radius:8px; overflow:hidden; }
  th, td { padding:8px 10px; border-bottom:1px solid var(--line); text-align:left; vertical-align:top; }
  th { font-weight:600; font-size:12px; color:var(--muted); }
  .num { text-align:right; font-variant-numeric:tabular-nums; }
  .badge { display:inline-block; padding:1px 8px; border-radius:999px; font-size:12px; font-weight:600; border:1px solid currentColor; }
  .good { color:var(--good); } .warn { color:var(--warn); } .bad { color:var(--bad); }
  .good-t { color:var(--good); } .bad-t { color:var(--bad); }
  section { background:var(--card); border:1px solid var(--line); border-radius:8px; padding:16px; margin-top:20px; }
  details { margin-top:12px; } summary { cursor:pointer; font-weight:600; }
  .problems { color:var(--bad); }
  .pagerow { margin:12px 0; } .pagehead { font-size:12px; color:var(--muted); margin-bottom:4px; }
  .triple { display:grid; grid-template-columns:repeat(3,1fr); gap:8px; }
  figure { margin:0; } figure img { width:100%; border:1px solid var(--line); background:#fff; display:block; }
  figcaption { font-size:12px; color:var(--muted); text-align:center; }
  .noimg { aspect-ratio:210/297; border:1px dashed var(--line); display:grid; place-items:center; color:var(--muted); }
  table.xml code { font-size:12px; word-break:break-all; }
  .scroll { overflow-x:auto; }
  @media (max-width:700px) { .triple { grid-template-columns:1fr; } }
</style></head>
<body><main>
  <h1>Fidelidade com o Word</h1>
  <p class="muted">Gerado em ${new Date().toLocaleString('pt-BR')} · motor: SuperDoc · referência: Microsoft Word</p>
  <div class="cards">
    <div class="card">Documentos<br><b>${docs.length}</b></div>
    <div class="card">Salvos sem perda<br><b class="good-t">${count(saveVerdict, 'good')}</b> / ${docs.length}</div>
    <div class="card">Exibição muito próxima<br><b class="good-t">${count(displayVerdict, 'good')}</b> / ${docs.length}</div>
    <div class="card">Abrir no pior caso (CPU ${SLOW_CPU}× lenta)<br><b>${fmtSec(Math.max(0, ...docs.map((d) => d.superdoc.perf?.slowOpenMs ?? 0)))}</b></div>
  </div>
  <p class="muted"><b>Salvar sem perda:</b> o Textos abre e salva o arquivo sem editar; o Word abre o resultado e comparamos com o original, página a página.
  <b>Exibição:</b> página desenhada pelo Textos comparada com a do Word. "Diferença" = pixels diferentes ÷ pixels com conteúdo.</p>
  <div class="scroll"><table><thead><tr>
    <th>Documento</th><th class="num">Págs. Word</th><th>Salvar</th><th class="num">Maior dif.</th><th class="num">Itens XML perdidos</th>
    <th>Exibição</th><th class="num">Págs. Textos</th><th class="num">Maior dif.</th>
    <th class="num">Abrir (CPU ${SLOW_CPU}× lenta)</th><th class="num">Memória</th>
  </tr></thead><tbody>${rows}</tbody></table></div>
  ${details}
</main></body></html>`;
  const file = path.join(outDir, 'report.html');
  await writeFile(file, html);
  return file;
}
