#!/usr/bin/env node
/**
 * Bancada de fidelidade do Textos, usando o Microsoft Word como referência.
 *
 * Uso:
 *   node bench/run.mjs [pasta-ou-arquivo ...] [--out bench/out] [--limit N]
 * Sem argumentos, usa bench/corpus.
 *
 * Para cada .docx:
 *  1. SALVAR SEM PERDA — o SuperDoc abre e exporta sem editar nada; o Word abre o
 *     resultado e gera um PDF, comparado página a página com o PDF do original.
 *     Também compara o XML interno e lista o que sumiu.
 *  2. EXIBIÇÃO — as páginas desenhadas pelo SuperDoc são comparadas com o PDF do Word.
 */
import { mkdir, readdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { comparePages, decodePng, encodePng, rasterizePdf } from './lib/images.mjs';
import { diffDocx } from './lib/ooxml.mjs';
import { writeReport } from './lib/report.mjs';
import { startSuperDoc } from './lib/superdoc.mjs';
import { runWord } from './lib/word.mjs';

const benchDir = path.dirname(fileURLToPath(import.meta.url));

function parseArgs(argv) {
  const opts = { inputs: [], out: path.join(benchDir, 'out'), limit: Infinity };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--out') opts.out = path.resolve(argv[++i]);
    else if (argv[i] === '--limit') opts.limit = Number(argv[++i]);
    else opts.inputs.push(path.resolve(argv[i]));
  }
  if (opts.inputs.length === 0) opts.inputs.push(path.join(benchDir, 'corpus'));
  return opts;
}

async function collectDocx(inputs) {
  const files = [];
  const visit = async (p) => {
    const s = await stat(p);
    if (s.isDirectory()) for (const e of await readdir(p)) await visit(path.join(p, e));
    else if (/\.docx$/i.test(p) && !path.basename(p).startsWith('~$')) files.push(p);
  };
  for (const i of inputs) await visit(i);
  return files.sort();
}

const slug = (file, used) => {
  let s = path.basename(file, path.extname(file)).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w.-]+/g, '_').slice(0, 60);
  while (used.has(s)) s += '_';
  used.add(s);
  return s;
};

const pct = (x) => (x == null ? null : Math.round(x * 10000) / 100);
const log = (...a) => console.log(new Date().toLocaleTimeString('pt-BR'), ...a);

async function comparePageSets(ref, other, dir, prefix) {
  const pages = [];
  const n = Math.max(ref.length, other.length);
  for (let i = 0; i < n; i++) {
    const a = ref[i];
    const b = other[i];
    const entry = { page: i + 1, diff: null, missing: !a ? 'referência' : !b ? 'comparado' : null };
    if (a && b) {
      const c = comparePages(a, b);
      entry.diff = pct(c.diff);
      entry.sizeMismatch = c.sizeMismatch;
      if (c.diff > 0) {
        entry.diffImage = `${prefix}-diff-${i + 1}.png`;
        await writeFile(path.join(dir, entry.diffImage), await encodePng(c.diffImage));
      }
    }
    pages.push(entry);
  }
  return pages;
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  const files = (await collectDocx(opts.inputs)).slice(0, opts.limit);
  if (files.length === 0) throw new Error('Nenhum .docx encontrado.');
  await rm(opts.out, { recursive: true, force: true });
  await mkdir(opts.out, { recursive: true });
  log(`${files.length} documento(s). Saída: ${opts.out}`);

  const used = new Set();
  const docs = files.map((file) => ({ file, id: slug(file, used) }));

  // 1) SuperDoc: abrir, fotografar páginas, exportar.
  const sd = await startSuperDoc();
  try {
    for (const [i, d] of docs.entries()) {
      d.dir = path.join(opts.out, d.id);
      await mkdir(d.dir, { recursive: true });
      d.bytes = await readFile(d.file);
      log(`[${i + 1}/${docs.length}] SuperDoc: ${path.basename(d.file)}`);
      const r = await sd.process(d.bytes, path.basename(d.file));
      d.superdoc = { pages: r.pages, errors: r.errors, loadMs: r.loadMs, exportError: r.exportError, consoleErrors: r.consoleErrors.slice(0, 10), blockedHosts: r.blockedHosts };
      d.sdShots = [];
      for (const [k, png] of r.screenshots.entries()) {
        const name = `superdoc-${k + 1}.png`;
        await writeFile(path.join(d.dir, name), png);
        d.sdShots.push(name);
      }
      if (r.exported) {
        d.roundtripPath = path.join(d.dir, 'roundtrip.docx');
        await writeFile(d.roundtripPath, r.exported);
      }
    }
  } finally {
    await sd.stop();
  }

  // 2) Word: PDF do original e do arquivo salvo pelo SuperDoc (num único lote).
  log('Word: gerando PDFs de referência…');
  const jobs = [];
  for (const d of docs) {
    jobs.push({ op: 'pdf', in: d.file, out: path.join(d.dir, 'word-original.pdf'), id: d.id, kind: 'original' });
    if (d.roundtripPath) jobs.push({ op: 'pdf', in: d.roundtripPath, out: path.join(d.dir, 'word-roundtrip.pdf'), id: d.id, kind: 'roundtrip' });
  }
  const results = await runWord(jobs.map(({ op, in: i, out }) => ({ op, in: i, out })));
  results.forEach((res, k) => {
    const d = docs.find((x) => x.id === jobs[k].id);
    d[`word_${jobs[k].kind}`] = { ok: res.ok, pages: res.pages, error: res.error, ms: res.ms };
  });

  // 3) Comparações.
  for (const [i, d] of docs.entries()) {
    log(`[${i + 1}/${docs.length}] comparando: ${path.basename(d.file)}`);
    d.result = { save: {}, display: {} };
    const orig = d.word_original?.ok ? await rasterizePdf(path.join(d.dir, 'word-original.pdf')) : null;
    if (orig) {
      d.wordPages = [];
      for (const [k, r] of orig.entries()) {
        const name = `word-${k + 1}.png`;
        await writeFile(path.join(d.dir, name), await encodePng(r));
        d.wordPages.push(name);
      }
    }

    // Exibição: SuperDoc × Word.
    if (orig && d.sdShots.length) {
      const shots = [];
      for (const s of d.sdShots) shots.push(await decodePng(await readFile(path.join(d.dir, s))));
      d.result.display.pages = await comparePageSets(orig, shots, d.dir, 'display');
    }

    // Salvar: Word(original) × Word(arquivo salvo pelo SuperDoc).
    if (d.roundtripPath) {
      d.result.save.xml = diffDocx(d.bytes, await readFile(d.roundtripPath));
      if (orig && d.word_roundtrip?.ok) {
        const rt = await rasterizePdf(path.join(d.dir, 'word-roundtrip.pdf'));
        d.roundtripPages = [];
        for (const [k, r] of rt.entries()) {
          const name = `roundtrip-${k + 1}.png`;
          await writeFile(path.join(d.dir, name), await encodePng(r));
          d.roundtripPages.push(name);
        }
        d.result.save.pages = await comparePageSets(orig, rt, d.dir, 'save');
      }
    }
    delete d.bytes;
  }

  const reportPath = await writeReport(opts.out, docs);
  await writeFile(path.join(opts.out, 'report.json'), JSON.stringify(docs, null, 2));
  log(`Relatório: ${reportPath}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
