// Access to the live Office.js host globals.
//
// The adapters are written against injected shims so they stay headless-testable
// (see `office-types.ts`). This module is the one place that reaches for the
// real `Excel`/`Office` globals — guarded so that outside a real Excel host
// (jsdom tests, a browser tab) the caller falls back to the fake data source.

import type { WorkbookStamp } from '@timeline/engine';
import type { IsSetSupported } from './office-types.ts';
import { OfficeWorkbookStamp } from './workbook-stamp.ts';

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
 * The document url (file path), or null. A weak workbook identity — empty for
 * unsaved workbooks and unstable across move/rename — used only as a last-resort
 * fallback for {@link resolveWorkbookKey}.
 */
export function getDocumentUrl(): string | null {
  return (globalThis as ExcelHostGlobal).Office?.context?.document?.url ?? null;
}

/** A fresh workbook GUID (crypto-random; falls back to getRandomValues). */
export function newWorkbookGuid(): string {
  const cryptoObj = globalThis.crypto;
  if (typeof cryptoObj.randomUUID === 'function') {
    return cryptoObj.randomUUID();
  }
  const bytes = new Uint8Array(16);
  cryptoObj.getRandomValues(bytes);
  return Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
}

/**
 * A stable, per-workbook key to isolate persisted history. Prefers the GUID
 * stamped into `workbook.settings` (ADR-0006) — it lives inside the .xlsx, so it
 * travels with the file, is unique per workbook, and survives save/rename. If no
 * stamp exists yet (a fresh workbook), one is minted and written back. Returns
 * null only when the stamp is unavailable AND there is no document url — the
 * caller must then avoid the shared store (use a session-unique key) so two
 * unidentifiable workbooks never share a timeline.
 *
 * `makeStamp`/`makeGuid` are injectable for headless tests.
 */
export async function resolveWorkbookKey(
  makeStamp: () => WorkbookStamp = () => new OfficeWorkbookStamp(),
  makeGuid: () => string = newWorkbookGuid,
): Promise<string | null> {
  try {
    const stamp = makeStamp();
    const existing = await stamp.read();
    if (existing !== null && existing.workbookGuid !== '') {
      return existing.workbookGuid;
    }
    const workbookGuid = makeGuid();
    await stamp.write({ workbookGuid, tipHash: existing?.tipHash ?? '' });
    return workbookGuid;
  } catch {
    // Office settings unavailable (e.g. not in a real host): fall back to the url.
    return getDocumentUrl();
  }
}
