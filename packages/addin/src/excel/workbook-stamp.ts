/**
 * `OfficeWorkbookStamp` — the in-file workbook stamp (ADR-0006).
 *
 * Reads/writes a tiny `{ workbookGuid, tipHash }` stamp into
 * `Office.context.document.settings`, which lives inside the `.xlsx` and so
 * travels with the file. The history store is keyed by `workbookGuid`; the
 * `tipHash` lets `attach` decide clean-resume vs. drift on launch.
 *
 * The settings object is injected (default `Office.context.document.settings`)
 * so headless tests can pass a fake — we don't hard-depend on a global `Office`
 * at construction time.
 */
import type { WorkbookStamp, WorkbookStampData } from '@timeline/engine';

/** The single settings key under which the stamp JSON is stored. */
const STAMP_KEY = 'timeline.workbookStamp';

/**
 * The slice of `Office.Settings` this adapter needs. Both `get`/`set` are
 * synchronous in the Office.js API; `saveAsync` persists into the file.
 */
export interface SettingsLike {
  get(key: string): unknown;
  set(key: string, value: unknown): void;
  saveAsync(callback?: (result: { status: string }) => void): void;
}

/** True when the value has the shape of a stamp. */
function isStampData(value: unknown): value is WorkbookStampData {
  if (typeof value !== 'object' || value === null) {
    return false;
  }
  const candidate = value as { workbookGuid?: unknown; tipHash?: unknown };
  return typeof candidate.workbookGuid === 'string' && typeof candidate.tipHash === 'string';
}

/** Resolves the default settings object from the ambient Office global, if present. */
function defaultSettings(): SettingsLike | undefined {
  const office = (globalThis as { Office?: { context?: { document?: { settings?: unknown } } } })
    .Office;
  return office?.context?.document?.settings as SettingsLike | undefined;
}

export class OfficeWorkbookStamp implements WorkbookStamp {
  readonly #settings: SettingsLike;

  /**
   * @param settings injectable Office settings object (default
   *   `Office.context.document.settings`). Throws if neither is available.
   */
  constructor(settings: SettingsLike | undefined = defaultSettings()) {
    if (settings === undefined) {
      throw new Error('OfficeWorkbookStamp: no settings object (Office not initialised).');
    }
    this.#settings = settings;
  }

  read(): Promise<WorkbookStampData | null> {
    const raw = this.#settings.get(STAMP_KEY);
    return Promise.resolve(
      isStampData(raw) ? { workbookGuid: raw.workbookGuid, tipHash: raw.tipHash } : null,
    );
  }

  write(data: WorkbookStampData): Promise<void> {
    this.#settings.set(STAMP_KEY, { workbookGuid: data.workbookGuid, tipHash: data.tipHash });
    return new Promise<void>((resolve, reject) => {
      this.#settings.saveAsync((result) => {
        if (result.status === 'succeeded') {
          resolve();
        } else {
          reject(new Error(`OfficeWorkbookStamp: saveAsync failed (${result.status}).`));
        }
      });
    });
  }
}
