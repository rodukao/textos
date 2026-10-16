// Abre um .docx na página da bancada (mesmo motor e configuração do Textos) e executa um trecho
// de JavaScript na página já diagramada, imprimindo o resultado. Para investigar a diagramação
// pelo DOM (posições de imagens, linhas, fragmentos).
// Uso: node bench/tools/inspecionar.mjs <arquivo.docx> <script.js>
//   o script é o corpo de uma função assíncrona; use `return` para devolver o resultado (JSON).
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';
import { createServer } from 'vite';

const [file, scriptFile] = process.argv.slice(2);
const harnessDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'harness');
const server = await createServer({ root: harnessDir, configFile: path.join(harnessDir, 'vite.config.mjs'), server: { port: 0 } });
await server.listen();
const url = server.resolvedUrls.local[0];
const origin = new URL(url).origin;
const browser = await chromium.launch({ channel: 'msedge', headless: true, args: ['--disable-lcd-text'] });
try {
  const context = await browser.newContext({ viewport: { width: 1000, height: 1200 }, deviceScaleFactor: 1 });
  // Os documentos ficam na máquina: nada sai do servidor local.
  await context.route('**/*', (r) => {
    const u = r.request().url();
    return u.startsWith(origin) || u.startsWith('data:') || u.startsWith('blob:') ? r.continue() : r.abort();
  });
  const page = await context.newPage();
  page.on('console', (m) => m.type() === 'error' && console.error('[console]', m.text().slice(0, 300)));
  await page.goto(url);
  await page.waitForFunction(() => !!window.bench);
  await page.evaluate((b64) => window.bench.load(b64, 'documento.docx'), readFileSync(file).toString('base64'));
  const body = readFileSync(scriptFile, 'utf8');
  const result = await page.evaluate(`(async () => { ${body} })()`);
  console.log(typeof result === 'string' ? result : JSON.stringify(result, null, 1));
} finally {
  await browser.close();
  await server.close();
}
