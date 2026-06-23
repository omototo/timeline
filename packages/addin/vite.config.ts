import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Dev/preview ports live in the 9588x band to avoid colliding with anything on 3000.
// strictPort: fail loudly rather than drift to another port, so manifest.xml URLs always match.
const DEV_PORT = 9588;
const PREVIEW_PORT = 9589;

export default defineConfig({
  plugins: [react()],
  server: {
    port: DEV_PORT,
    strictPort: true,
  },
  preview: {
    port: PREVIEW_PORT,
    strictPort: true,
  },
  build: {
    outDir: 'dist',
  },
});
