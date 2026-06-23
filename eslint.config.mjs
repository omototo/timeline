import js from '@eslint/js';
import tseslint from 'typescript-eslint';
import reactPlugin from 'eslint-plugin-react';
import reactHooks from 'eslint-plugin-react-hooks';
import jsxA11y from 'eslint-plugin-jsx-a11y';
import vitest from '@vitest/eslint-plugin';
import prettier from 'eslint-config-prettier';

export default tseslint.config(
  {
    ignores: ['**/dist/**', '**/coverage/**', '**/node_modules/**', 'scripts/**'],
  },
  js.configs.recommended,
  ...tseslint.configs.strictTypeChecked,
  ...tseslint.configs.stylisticTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      // Leading-underscore params/vars are an intentional "unused on purpose"
      // marker — e.g. not-yet-implemented interface stubs that must keep their
      // frozen signatures. Honour the convention rather than fight it.
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
        },
      ],
    },
  },
  // Engine-purity wall: the engine must not reach for Office.js, React, or DOM globals.
  {
    files: ['packages/engine/**/*.{ts,tsx}'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          paths: [
            { name: 'office-js', message: 'The engine must stay pure: no Office.js.' },
            { name: '@microsoft/office-js', message: 'The engine must stay pure: no Office.js.' },
            { name: 'react', message: 'The engine must stay pure: no React.' },
            { name: 'react-dom', message: 'The engine must stay pure: no React DOM.' },
          ],
          patterns: [
            {
              group: ['office-js/*', '@microsoft/office-js/*'],
              message: 'The engine must stay pure: no Office.js.',
            },
          ],
        },
      ],
    },
  },
  // Addin override: React + a11y rules for the host package.
  {
    files: ['packages/addin/**/*.{ts,tsx}'],
    ...reactPlugin.configs.flat.recommended,
    languageOptions: {
      ...reactPlugin.configs.flat.recommended.languageOptions,
    },
    settings: {
      react: { version: 'detect' },
    },
    plugins: {
      react: reactPlugin,
      'react-hooks': reactHooks,
      'jsx-a11y': jsxA11y,
    },
    rules: {
      ...reactPlugin.configs.flat.recommended.rules,
      ...reactHooks.configs.recommended.rules,
      ...jsxA11y.flatConfigs.recommended.rules,
      'react/react-in-jsx-scope': 'off',
    },
  },
  // Vitest rules on test files.
  {
    files: ['**/*.{test,spec}.{ts,tsx}', '**/test/**/*.{ts,tsx}'],
    plugins: { vitest },
    rules: {
      ...vitest.configs.recommended.rules,
    },
  },
  // Config files run outside the TS project graph; lint them untyped.
  {
    files: ['*.{js,mjs,cjs}', '**/*.config.{ts,js,mjs}'],
    ...tseslint.configs.disableTypeChecked,
  },
  // eslint-config-prettier last to switch off stylistic rules that conflict with Prettier.
  prettier,
);
