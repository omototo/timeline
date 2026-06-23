import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    projects: [
      {
        test: {
          name: 'engine',
          environment: 'node',
          root: './packages/engine',
          include: ['test/**/*.{test,spec}.ts'],
        },
      },
      {
        test: {
          name: 'addin',
          environment: 'jsdom',
          root: './packages/addin',
          setupFiles: ['./test/setup.ts'],
          include: ['test/**/*.{test,spec}.{ts,tsx}'],
        },
      },
    ],
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html'],
      include: ['packages/*/src/**/*.{ts,tsx}'],
      // The DOM bootstrap (createRoot wrapper) has no logic to test; exclude it.
      exclude: ['packages/addin/src/main.tsx'],
      thresholds: {
        statements: 80,
        branches: 80,
        functions: 80,
        lines: 80,
      },
    },
  },
});
