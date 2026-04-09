import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  // For GitHub Pages deployment under a subpath (e.g. /pptx-svg/react/)
  base: process.env.VITE_BASE_PATH || '/',
  // pptx-svg uses `new URL('./main.wasm', import.meta.url)` internally.
  // Excluding it from pre-bundling ensures the URL resolves correctly in dev mode.
  optimizeDeps: {
    exclude: ['pptx-svg'],
  },
});
