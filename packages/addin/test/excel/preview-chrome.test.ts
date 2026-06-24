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
});
