import { afterEach, describe, expect, it } from 'vitest';
import type { WorkbookStamp, WorkbookStampData } from '@timeline/engine';
import { newWorkbookGuid, resolveWorkbookKey } from '../src/excel/excel-host.ts';

class FakeStamp implements WorkbookStamp {
  data: WorkbookStampData | null;
  readonly written: WorkbookStampData[] = [];
  constructor(initial: WorkbookStampData | null = null) {
    this.data = initial;
  }
  read(): Promise<WorkbookStampData | null> {
    return Promise.resolve(this.data);
  }
  write(data: WorkbookStampData): Promise<void> {
    this.written.push(data);
    this.data = data;
    return Promise.resolve();
  }
}

describe('newWorkbookGuid', () => {
  it('returns distinct non-empty ids', () => {
    const a = newWorkbookGuid();
    const b = newWorkbookGuid();
    expect(a).not.toBe('');
    expect(a).not.toBe(b);
  });
});

describe('resolveWorkbookKey', () => {
  afterEach(() => {
    delete (globalThis as { Office?: unknown }).Office;
  });

  it('reuses the workbook GUID already stamped in settings', async () => {
    const stamp = new FakeStamp({ workbookGuid: 'wb-123', tipHash: 'abc' });
    const key = await resolveWorkbookKey(
      () => stamp,
      () => 'should-not-be-used',
    );
    expect(key).toBe('wb-123');
    expect(stamp.written).toHaveLength(0); // existing stamp left untouched
  });

  it('mints and writes a GUID for an unstamped (fresh) workbook', async () => {
    const stamp = new FakeStamp(null);
    const key = await resolveWorkbookKey(
      () => stamp,
      () => 'fresh-guid',
    );
    expect(key).toBe('fresh-guid');
    expect(stamp.written).toEqual([{ workbookGuid: 'fresh-guid', tipHash: '' }]);
  });

  it('preserves an existing tipHash when minting a missing guid', async () => {
    const stamp = new FakeStamp({ workbookGuid: '', tipHash: 'keep-me' });
    await resolveWorkbookKey(
      () => stamp,
      () => 'new-guid',
    );
    expect(stamp.written).toEqual([{ workbookGuid: 'new-guid', tipHash: 'keep-me' }]);
  });

  it('falls back to the document url when the stamp is unavailable', async () => {
    (globalThis as { Office?: unknown }).Office = {
      context: { document: { url: '/path/Book.xlsx' } },
    };
    const key = await resolveWorkbookKey(() => {
      throw new Error('no settings');
    });
    expect(key).toBe('/path/Book.xlsx');
  });

  it('returns null when neither a stamp nor a url is available', async () => {
    const key = await resolveWorkbookKey(() => {
      throw new Error('no settings');
    });
    expect(key).toBeNull();
  });
});
