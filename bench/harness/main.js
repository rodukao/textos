// Página que a bancada abre no navegador (via Playwright) para carregar um .docx
// no SuperDoc, esperar a diagramação e exportar de volta.
import { SuperDoc } from 'superdoc';
import 'superdoc/style.css';
// Mesma configuração do app: a bancada mede exatamente o que o Textos faz.
import { baseSuperDocConfig } from '../../src/engine/superdoc-config.ts';

let superdoc = null;

const toBase64 = async (blob) => {
  const bytes = new Uint8Array(await blob.arrayBuffer());
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
};

const fromBase64 = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));

/** Espera a diagramação estabilizar: nenhuma atualização de paginação por `quietMs`. */
function waitForLayout(state, quietMs = 1200, maxMs = 60000) {
  return new Promise((resolve) => {
    const start = performance.now();
    const tick = () => {
      const now = performance.now();
      if (now - state.lastPagination >= quietMs || now - start > maxMs) resolve();
      else setTimeout(tick, 100);
    };
    tick();
  });
}

window.bench = {
  /** Carrega o .docx (base64) e resolve quando estiver diagramado. */
  async load(b64, name = 'documento.docx') {
    const start = performance.now();
    const state = { lastPagination: start, pages: null, errors: [], paginations: 0 };
    const file = new File([fromBase64(b64)], name, {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await new Promise((resolve, reject) => {
      superdoc = new SuperDoc({
        ...baseSuperDocConfig,
        selector: '#editor',
        document: file,
        documentMode: 'viewing',
        onReady: () => resolve(),
        onPaginationUpdate: (p) => {
          state.lastPagination = performance.now();
          state.paginations += 1;
          state.pages = p.totalPages;
        },
        onException: (e) => {
          state.errors.push(String(e?.error?.message ?? e?.message ?? e?.stage ?? 'erro'));
        },
      });
      setTimeout(() => reject(new Error('SuperDoc não ficou pronto em 60 s')), 60000);
    });
    await document.fonts.ready;
    await waitForLayout(state);
    window.__benchState = state;
    // Tempo até a última diagramação: documento aberto e pronto para ler.
    return {
      pages: state.pages,
      errors: state.errors,
      openMs: Math.round(state.lastPagination - start),
      // Quantas vezes o documento foi diagramado ao abrir (reflows por fontes que chegam tarde etc.).
      paginations: state.paginations,
    };
  },

  /** Exporta o documento atual para .docx (base64), sem nenhuma edição. */
  async exportDocx() {
    const blob = await superdoc.export({ exportType: ['docx'], triggerDownload: false });
    return toBase64(blob);
  },
};
