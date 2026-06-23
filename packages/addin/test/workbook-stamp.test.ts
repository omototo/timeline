import { describe, it, expect, vi, afterEach } from 'vitest';
import { OfficeWorkbookStamp, type SettingsLike } from '../src/excel/workbook-stamp.ts';
import type { WorkbookStampData } from '@timeline/engine';

const STAMP_KEY = 'timeline.workbookStamp';
const STAMP: WorkbookStampData = { workbookGuid: 'guid-123', tipHash: 'hash-abc' };

/** An in-memory `Office.Settings` double whose `saveAsync` reports `status`. */
function fakeSettings(saveStatus: 'succeeded' | 'failed' = 'succeeded'): SettingsLike & {
  readonly store: Map<string, unknown>;
} {
  const store = new Map<string, unknown>();
  return {
    store,
    get: (key) => store.get(key),
    set: (key, value) => {
      store.set(key, value);
    },
    saveAsync: (callback) => {
      callback?.({ status: saveStatus });
    },
  };
}

describe('OfficeWorkbookStamp', () => {
  afterEach(() => {
    delete (globalThis as { Office?: unknown }).Office;
  });

  it('returns null when no stamp has been written', async () => {
    const stamp = new OfficeWorkbookStamp(fakeSettings());
    expect(await stamp.read()).toBeNull();
  });

  it('round-trips a written stamp', async () => {
    const settings = fakeSettings();
    const stamp = new OfficeWorkbookStamp(settings);
    await stamp.write(STAMP);
    expect(settings.store.get(STAMP_KEY)).toEqual(STAMP);
    expect(await stamp.read()).toEqual(STAMP);
  });

  it('persists via saveAsync on write', async () => {
    const settings = fakeSettings();
    const spy = vi.spyOn(settings, 'saveAsync');
    const stamp = new OfficeWorkbookStamp(settings);
    await stamp.write(STAMP);
    expect(spy).toHaveBeenCalledOnce();
  });

  it('rejects when saveAsync reports failure', async () => {
    const stamp = new OfficeWorkbookStamp(fakeSettings('failed'));
    await expect(stamp.write(STAMP)).rejects.toThrow('saveAsync failed');
  });

  it('returns null for a malformed stored value', async () => {
    const settings = fakeSettings();
    settings.set(STAMP_KEY, { workbookGuid: 'only-guid' });
    const stamp = new OfficeWorkbookStamp(settings);
    expect(await stamp.read()).toBeNull();
  });

  it('returns null for a non-object stored value', async () => {
    const settings = fakeSettings();
    settings.set(STAMP_KEY, 'not-an-object');
    const stamp = new OfficeWorkbookStamp(settings);
    expect(await stamp.read()).toBeNull();
  });

  it('falls back to the ambient Office settings when none is injected', async () => {
    const settings = fakeSettings();
    (globalThis as { Office?: unknown }).Office = {
      context: { document: { settings } },
    };
    const stamp = new OfficeWorkbookStamp();
    await stamp.write(STAMP);
    expect(await stamp.read()).toEqual(STAMP);
  });

  it('throws at construction when no settings object is available', () => {
    expect(() => new OfficeWorkbookStamp()).toThrow('Office not initialised');
  });
});
