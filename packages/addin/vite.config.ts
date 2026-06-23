import { defineConfig, type UserConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { getHttpsServerOptions } from 'office-addin-dev-certs';

// Dev/preview ports live in the 9588x band to avoid colliding with anything on 3000.
// strictPort: fail loudly rather than drift to another port, so manifest.xml URLs always match.
const DEV_PORT = 9588;
const PREVIEW_PORT = 9589;

export default defineConfig(async ({ command }): Promise<UserConfig> => {
  const config: UserConfig = {
    plugins: react(),
    server: {
      host: 'localhost',
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
  };

  if (command === 'serve') {
    // Excel sideload needs HTTPS against the trusted Office localhost cert, but
    // plain-browser dev does not. When the dev certs are missing, fall back to
    // HTTP instead of hard-failing so `bun run dev` still serves the pane.
    try {
      config.server = {
        ...config.server,
        https: await getHttpsServerOptions(),
      };
    } catch {
      console.warn(
        '[vite] Office dev certs not found — serving over HTTP. Run `bun run --filter @timeline/addin certs` for the trusted HTTPS server required by Excel sideload.',
      );
    }
  }

  return config;
});
