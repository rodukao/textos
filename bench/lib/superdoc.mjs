// Abre cada .docx no SuperDoc (num Edge sem janela), fotografa as páginas e exporta de volta.
import { chromium } from 'playwright-core';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createServer } from 'vite';

const harnessDir = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'harness');

export async function startSuperDoc() {
  const server = await createServer({
    root: harnessDir,
    configFile: path.join(harnessDir, 'vite.config.mjs'),
    server: { port: 0, strictPort: false },
  });
  await server.listen();
  const url = server.resolvedUrls.local[0];
  const origin = new URL(url).origin;

  // Edge já vem com o Windows: nada de baixar navegador.
  // --disable-lcd-text: suavização em tons de cinza (como num PDF), sem as franjas coloridas do ClearType.
  const browser = await chromium.launch({ channel: 'msedge', headless: true, args: ['--disable-lcd-text'] });

  async function process(bytes, name) {
    const context = await browser.newContext({ viewport: { width: 1000, height: 1200 }, deviceScaleFactor: 1 });
    const page = await context.newPage();
    // Os documentos ficam na máquina: qualquer requisição para fora do servidor local é bloqueada.
    const blocked = new Set();
    await context.route('**/*', (route) => {
      const u = route.request().url();
      if (u.startsWith(origin) || u.startsWith('data:') || u.startsWith('blob:')) return route.continue();
      blocked.add(new URL(u).host);
      return route.abort();
    });
    const consoleErrors = [];
    page.on('console', (m) => m.type() === 'error' && consoleErrors.push(m.text().slice(0, 300)));
    page.on('pageerror', (e) => consoleErrors.push(String(e).slice(0, 300)));

    try {
      await page.goto(url);
      await page.waitForFunction(() => !!window.bench);
      const t0 = Date.now();
      const loaded = await page.evaluate(([b64, n]) => window.bench.load(b64, n), [Buffer.from(bytes).toString('base64'), name]);
      const loadMs = Date.now() - t0;

      const pageEls = await page.$$('.superdoc-page');
      const screenshots = [];
      for (const el of pageEls) {
        await el.scrollIntoViewIfNeeded();
        await page.waitForTimeout(150);
        screenshots.push(await el.screenshot({ animations: 'disabled' }));
      }

      let exported = null;
      let exportError = null;
      try {
        exported = Buffer.from(await page.evaluate(() => window.bench.exportDocx()), 'base64');
      } catch (e) {
        exportError = String(e).slice(0, 500);
      }
      return { pages: loaded.pages ?? pageEls.length, errors: loaded.errors, loadMs, screenshots, exported, exportError, consoleErrors, blockedHosts: [...blocked] };
    } catch (e) {
      return { pages: 0, errors: [String(e).slice(0, 500)], loadMs: null, screenshots: [], exported: null, exportError: 'não carregou', consoleErrors, blockedHosts: [...blocked] };
    } finally {
      await context.close();
    }
  }

  async function stop() {
    await browser.close();
    await server.close();
  }

  return { process, stop };
}
