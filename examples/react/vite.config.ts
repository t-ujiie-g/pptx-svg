import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  // pptx-svg uses `new URL('./main.wasm', import.meta.url)` internally.
  // Excluding it from pre-bundling ensures the URL resolves correctly in dev mode.
  optimizeDeps: {
    exclude: ['pptx-svg'],
  },
});
