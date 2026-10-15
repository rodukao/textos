import { defineConfig } from 'vite';

export default defineConfig({
  // O SuperDoc localiza alguns recursos por caminho relativo ao próprio arquivo;
  // o pré-empacotamento do Vite muda esse caminho.
  optimizeDeps: { exclude: ['superdoc'] },
  define: {
    __VUE_OPTIONS_API__: 'true',
    __VUE_PROD_DEVTOOLS__: 'false',
    __VUE_PROD_HYDRATION_MISMATCH_DETAILS__: 'false',
  },
  server: { port: 1430, strictPort: true },
  logLevel: 'warn',
});
