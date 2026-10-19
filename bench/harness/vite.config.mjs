import path from 'node:path';
import { defineConfig } from 'vite';

// TEXTOS_ENGINE_DIR=<pasta do textos-engine>: a bancada usa o motor compilado direto do fork
// (packages/superdoc/dist), sem instalar nada no node_modules do Textos — dá para medir uma
// mudança no motor enquanto o app continua rodando com a versão publicada.
const engineDir = process.env.TEXTOS_ENGINE_DIR;
const engineDist = engineDir ? path.resolve(engineDir, 'packages/superdoc/dist') : null;

export default defineConfig({
  resolve: engineDist
    ? {
        alias: [
          { find: /^superdoc$/, replacement: path.join(engineDist, 'superdoc.es.js') },
          { find: /^superdoc\/style\.css$/, replacement: path.join(engineDist, 'style.css') },
        ],
      }
    : {},
  // O SuperDoc localiza alguns recursos por caminho relativo ao próprio arquivo;
  // o pré-empacotamento do Vite muda esse caminho.
  optimizeDeps: { exclude: ['superdoc'] },
  define: {
    __VUE_OPTIONS_API__: 'true',
    __VUE_PROD_DEVTOOLS__: 'false',
    __VUE_PROD_HYDRATION_MISMATCH_DETAILS__: 'false',
  },
  server: {
    port: 1430,
    strictPort: true,
    // O motor fora do projeto precisa ser servido também (além do próprio Textos).
    fs: engineDir ? { allow: [path.resolve(import.meta.dirname, '../..'), path.resolve(engineDir)] } : undefined,
  },
  logLevel: 'warn',
});
