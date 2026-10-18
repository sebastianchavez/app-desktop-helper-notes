import { describe, it, expect } from 'vitest';
import * as XLSX from 'xlsx';
import { unzipSync, strFromU8 } from 'fflate';
import { patchSpreadsheet } from '../electron/spreadsheet-edit';
import { spreadsheetErrorCode, spreadsheetErrorText } from '../electron/spreadsheet-values';

/** No recibe argumentos; retorna libro aislado con época 1904, nombre definido, matriz y errores de Excel. */
function fixture(): string {
  const book = XLSX.utils.book_new(),
    sheet = XLSX.utils.aoa_to_sheet([[1, 2, 3]]);
  sheet.A2 = { t: 'e', v: 7 };
  sheet.B2 = { t: 'n', v: 99, f: '_xlfn.FILTER(A1:C1,A1:C1>1)', F: 'B2:C2' };
  sheet.C2 = { t: 'e', v: 29, f: '_xlfn.UNKNOWN(1)' };
  sheet['!ref'] = 'A1:C2';
  XLSX.utils.book_append_sheet(book, sheet, 'Data');
  book.Workbook = { WBProps: { date1904: true }, Names: [{ Name: 'Rate', Ref: '0.2' }] };
  return Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString('base64');
}
describe('spreadsheet formula file compatibility', () => {
  it('preserves epoch, names and array definitions when rebuilding cell data', () => {
    const original = fixture(),
      book = XLSX.read(Buffer.from(original, 'base64'), { type: 'buffer', xlfn: true });
    book.Sheets.Data.A1.v = 5;
    delete book.Sheets.Data.B2.F;
    const saved = patchSpreadsheet(original, book, 0),
      xml = strFromU8(unzipSync(Buffer.from(saved, 'base64'))['xl/worksheets/sheet1.xml']);
    expect(xml).toContain('t="array"');
    expect(xml).toContain('ref="B2:C2"');
    expect(xml).toContain('_xlfn.FILTER');
    const read = XLSX.read(Buffer.from(saved, 'base64'), { type: 'buffer', xlfn: true });
    expect(read.Workbook?.WBProps?.date1904).toBe(true);
    expect(read.Workbook?.Names?.[0].Name).toBe('Rate');
    expect(read.Sheets.Data.A2.t).toBe('e');
  });
  it('restores native error type when a renderer sends the visible imported error text', () => {
    const original = fixture(),
      book = XLSX.read(Buffer.from(original, 'base64'), { type: 'buffer', xlfn: true });
    book.Sheets.Data.A2 = { t: 's', v: '#DIV/0!' };
    const read = XLSX.read(Buffer.from(patchSpreadsheet(original, book, 0), 'base64'), {
      type: 'buffer',
    });
    expect(read.Sheets.Data.A2.t).toBe('e');
    expect(read.Sheets.Data.A2.v).toBe(7);
    expect(read.Sheets.Data.C2.t).toBe('e');
    expect(read.Sheets.Data.C2.v).toBe(29);
  });
  it('maps known error codes including #NULL! zero and normalizes internal errors without converting regular text', () => {
    expect(spreadsheetErrorCode('#NULL!')).toBe(0);
    expect(spreadsheetErrorText(7)).toBe('#DIV/0!');
    expect(spreadsheetErrorCode('#CYCLE!')).toBe(23);
    expect(spreadsheetErrorCode('#ERROR!')).toBe(15);
    expect(spreadsheetErrorCode('hello')).toBeUndefined();
  });
});
