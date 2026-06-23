// Access to the live Office.js host globals.
//
// The adapters are written against injected shims so they stay headless-testable
// (see `office-types.ts`). This module is the one place that reaches for the
// real `Excel`/`Office` globals — guarded so that outside a real Excel host
// (jsdom tests, a browser tab) the caller falls back to the fake data source.

import type { IsSetSupported } from './office-types.ts';

/** The real `Excel.run` signature, sourced from `@types/office-js`. */
export type RealExcelRun = typeof Excel.run;

interface ExcelHostGlobal {
  Excel?: { run: RealExcelRun };
  Office?: {
    context?: {
      requirements?: { isSetSupported(name: string, version?: string): boolean };
      document?: { url?: string };
    };
  };
}

/** The live `Excel.run`, or `null` when not running inside an Excel host. */
export function getExcelRun(): RealExcelRun | null {
  const host = globalThis as ExcelHostGlobal;
  return host.Excel?.run ?? null;
}

/**
 * A feature-detector backed by `Office.context.requirements`, or an
 * assume-modern default when Office is unavailable.
 */
export function getIsSetSupported(): IsSetSupported {
  const requirements = (globalThis as ExcelHostGlobal).Office?.context?.requirements;
  if (!requirements) {
    return () => true;
  }
  return (name, version) => requirements.isSetSupported(name, version);
}

/**
 * A stable per-workbook key (the document url) used to isolate persisted history
 * so two workbooks never share a timeline. Null when not running in Office (one
 * shared store) — see ADR-0007 (history is otherwise origin-scoped).
 */
export function getWorkbookKey(): string | null {
  return (globalThis as ExcelHostGlobal).Office?.context?.document?.url ?? null;
}
