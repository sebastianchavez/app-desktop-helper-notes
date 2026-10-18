import { describe, it, expect } from 'vitest';
import * as XLSX from 'xlsx';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { patchSpreadsheet } from '../electron/spreadsheet-edit';
import { readSpreadsheetAppearance } from '../electron/spreadsheet-appearance';

/** Crea un libro y recursos centinela; retorna Base64 sin usar archivos del usuario. */
function fixture(): string {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([[10, 20], ['original']]), 'Datos');
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Otra hoja']]), 'Resumen');
  const archive = unzipSync(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
  archive['xl/media/image1.png'] = new Uint8Array([1, 2, 3, 4]);
  archive['xl/charts/chart1.xml'] = strToU8('<chart>conservar</chart>');
  return Buffer.from(zipSync(archive)).toString('base64');
}

describe('XLSX edits preserving the original package', () => {
  it('persists merges, formulas, dimensions and editable font/fill/alignment/number formats', () => {
    const original = fixture(),
      book = XLSX.read(Buffer.from(original, 'base64'), { type: 'buffer' });
    book.Sheets['Datos']['!merges'] = [{ s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }];
    book.Sheets['Datos']['C1'] = { t: 'n', v: 30, f: 'SUM(A1:B1)' };
    book.Sheets['Datos']['!ref'] = 'A1:C2';
    const result = patchSpreadsheet(original, book, 0, {
      formats: [
        {
          row: 0,
          col: 0,
          format: {
            bold: true,
            italic: true,
            underline: true,
            fontSize: 20,
            color: '#123456',
            background: '#abcdef',
            align: 'center',
            wrap: true,
            numberFormat: '0.00%',
          },
        },
      ],
      colWidths: [140, 80, 90],
      rowHeights: [40, 20],
    });
    const reopened = XLSX.read(Buffer.from(result, 'base64'), { type: 'buffer', cellStyles: true });
    expect(reopened.Sheets['Datos']['C1'].f).toBe('SUM(A1:B1)');
    expect(reopened.Sheets['Datos']['C1'].v).toBe(30);
    expect(reopened.Sheets['Datos']['!merges']).toHaveLength(1);
    expect(reopened.Sheets['Datos']['A1'].w).toBe('1000.00%');
    const css = readSpreadsheetAppearance(Buffer.from(result, 'base64'))['Datos']['0:0'];
    expect(css['font-weight']).toBe('bold');
    expect(css['color'].toLowerCase()).toBe('#123456');
    expect(css['background-color'].toLowerCase()).toBe('#abcdef');
    expect(css['text-align']).toBe('center');
    const before = unzipSync(Buffer.from(original, 'base64')),
      after = unzipSync(Buffer.from(result, 'base64'));
    for (const name of ['xl/media/image1.png', 'xl/charts/chart1.xml', 'xl/worksheets/sheet2.xml'])
      expect(after[name]).toEqual(before[name]);
  });
  it('keeps formatting when values change and moves style IDs with rows', () => {
    let original = fixture();
    let book = XLSX.read(Buffer.from(original, 'base64'), { type: 'buffer' });
    original = patchSpreadsheet(original, book, 0, {
      formats: [{ row: 0, col: 0, format: { bold: true } }],
    });
    book = XLSX.read(Buffer.from(original, 'base64'), { type: 'buffer' });
    book.Sheets['Datos']['A1'].v = 55;
    const edited = patchSpreadsheet(original, book, 0);
    expect(
      readSpreadsheetAppearance(Buffer.from(edited, 'base64'))['Datos']['0:0']['font-weight'],
    ).toBe('bold');
    book.Sheets['Datos'] = XLSX.utils.aoa_to_sheet([[null], [55]]);
    const moved = patchSpreadsheet(edited, book, 0, { origins: { A2: 'A1' } });
    expect(
      readSpreadsheetAppearance(Buffer.from(moved, 'base64'))['Datos']['1:0']['font-weight'],
    ).toBe('bold');
    expect(
      strFromU8(unzipSync(Buffer.from(moved, 'base64'))['xl/worksheets/sheet1.xml']),
    ).not.toContain('r="A1" s=');
  });
  it('removes merges when splitting cells, retaining covered values', () => {
    let original = fixture();
    const book = XLSX.read(Buffer.from(original, 'base64'), { type: 'buffer' });
    book.Sheets['Datos']['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    original = patchSpreadsheet(original, book, 0);
    book.Sheets['Datos']['!merges'] = [];
    const reopened = XLSX.read(Buffer.from(patchSpreadsheet(original, book, 0), 'base64'), {
      type: 'buffer',
    });
    expect(reopened.Sheets['Datos']['!merges'] ?? []).toHaveLength(0);
    expect(reopened.Sheets['Datos']['B1'].v).toBe(20);
  });
});
