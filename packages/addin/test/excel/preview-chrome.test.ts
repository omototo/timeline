import { describe, expect, it } from 'vitest';
import { OfficePreviewChrome } from '../../src/excel/preview-chrome.ts';
import { createFakeExcel, FakeWorkbook } from './fake-excel.ts';

describe('OfficePreviewChrome', () => {
  it('hides every real sheet on enter and restores them exactly on exit', async () => {
    const workbook = new FakeWorkbook();
    const sheet1 = workbook.addSheet('Sheet1');
    const ark1 = workbook.addSheet('Ark1');
    ark1.visibility = 'Hidden'; // a sheet the user had already hidden
    const ark2 = workbook.addSheet('Ark2');
    // An engine-owned preview surface must be left alone.
    const preview = workbook.addSheet('__tl_preview_0a0a0a0a');
    preview.visibility = 'VeryHidden';
    const { run } = createFakeExcel(workbook);
    const chrome = new OfficePreviewChrome(run);

    await chrome.enter();
    expect(sheet1.visibility).toBe('Hidden'); // was Visible -> hidden
    expect(ark2.visibility).toBe('Hidden');
    expect(ark1.visibility).toBe('Hidden'); // already Hidden, untouched
    expect(preview.visibility).toBe('VeryHidden'); // internal surface untouched

    await chrome.exit();
    expect(sheet1.visibility).toBe('Visible'); // restored
    expect(ark2.visibility).toBe('Visible'); // restored
    expect(ark1.visibility).toBe('Hidden'); // stays as the user had it
    expect(preview.visibility).toBe('VeryHidden');
  });

  it('exit is a no-op when nothing was hidden', async () => {
    const workbook = new FakeWorkbook();
    const { run } = createFakeExcel(workbook);
    const chrome = new OfficePreviewChrome(run);
    await chrome.exit(); // must not throw
    expect(workbook.syncCount).toBe(0);
  });

  it('recover() deletes orphan preview surfaces and un-hides stranded real sheets', async () => {
    const workbook = new FakeWorkbook();
    const sheet1 = workbook.addSheet('Sheet1');
    sheet1.visibility = 'Hidden'; // stranded hidden by an interrupted preview
    const secret = workbook.addSheet('Secret');
    secret.visibility = 'VeryHidden'; // a sheet the USER hid -> must stay hidden
    const orphan = workbook.addSheet('__tl_preview_0a0a0a0a'); // crash leftover
    orphan.visibility = 'VeryHidden';
    const { run } = createFakeExcel(workbook);

    await new OfficePreviewChrome(run).recover();

    expect(workbook.findSheet('__tl_preview_0a0a0a0a')).toBeUndefined(); // orphan gone
    expect(sheet1.visibility).toBe('Visible'); // stranded sheet recovered
    expect(secret.visibility).toBe('VeryHidden'); // user's hidden sheet untouched
  });

  it('recover() is a no-op when there are no orphan preview surfaces', async () => {
    const workbook = new FakeWorkbook();
    const hidden = workbook.addSheet('Sheet1');
    hidden.visibility = 'Hidden'; // user hid this; no crash signal -> leave it
    const { run } = createFakeExcel(workbook);
    await new OfficePreviewChrome(run).recover();
    expect(hidden.visibility).toBe('Hidden');
  });
});
