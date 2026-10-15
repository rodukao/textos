import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Porta fixa: o Tauri (src-tauri/tauri.conf.json → devUrl) espera o servidor em 1420.
export default defineConfig({
  plugins: [react()],
  clearScreen: false,
  server: {
    port: 1420,
    strictPort: true,
    watch: { ignored: ['**/src-tauri/**', '**/bench/**'] },
  },
  // O SuperDoc localiza alguns recursos por caminho relativo ao próprio arquivo;
  // o pré-empacotamento do Vite muda esse caminho.
  optimizeDeps: { exclude: ['superdoc'] },
  // O SuperDoc usa Vue internamente; estas flags evitam avisos e código morto.
  define: {
    __VUE_OPTIONS_API__: 'true',
    __VUE_PROD_DEVTOOLS__: 'false',
    __VUE_PROD_HYDRATION_MISMATCH_DETAILS__: 'false',
  },
  envPrefix: ['VITE_', 'TAURI_ENV_'],
  build: {
    target: 'es2022',
    // O motor de documentos do Word é grande por natureza.
    chunkSizeWarningLimit: 6000,
  },
});
