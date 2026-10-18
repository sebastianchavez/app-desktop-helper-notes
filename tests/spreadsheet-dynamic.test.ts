import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { unzipSync, strFromU8 } from 'fflate';
import { readSpreadsheetStructure } from '../electron/spreadsheet-structure';
import { patchSpreadsheet } from '../electron/spreadsheet-edit';
/** No recibe argumentos; retorna el archivo sintético generado por XlsxWriter como productor independiente. */
function fixture(): string {
  return readFileSync('tests/fixtures/dynamic-source.xlsx').toString('base64');
}
describe('native dynamic arrays and table structures', () => {
  it('preserves an unchanged native table byte for byte when the editor includes a leading formula equals sign', () => {
    const original = fixture(),
      book = XLSX.read(Buffer.from(original, 'base64'), { type: 'buffer', xlfn: true }),
      table = readSpreadsheetStructure(original).Data.tables[0];
    const result = patchSpreadsheet(original, book, 0, {
      tables: [{ ...table, calculated: { Total: '=' + book.Sheets.Data.D2.f } }],
    });
    expect(strFromU8(unzipSync(Buffer.from(result, 'base64'))['xl/tables/table1.xml'])).toBe(
      strFromU8(unzipSync(Buffer.from(original, 'base64'))['xl/tables/table1.xml']),
    );
  });
  it('updates mixed-case names without overwriting a global name shadowed by a local definition', () => {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([[1], [2]]), 'Data');
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([[3]]), 'Other');
    book.Workbook = {
      Names: [
        { Name: 'SalesRange', Ref: 'Data!$A$1:$A$2' },
        { Name: 'SalesRange', Ref: 'Data!$B$1:$B$2', Sheet: 0 },
      ],
    };
    const original = Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString(
      'base64',
    );
    const first = patchSpreadsheet(original, book, 0, {
      definedNames: { SALESRANGE: 'Data!$B$2:$B$3' },
    });
    let names = XLSX.read(Buffer.from(first, 'base64')).Workbook!.Names!;
    expect(names.find((n) => n.Sheet === undefined)!.Ref).toBe('Data!$A$1:$A$2');
    expect(names.find((n) => n.Sheet === 0)!.Ref).toBe('Data!$B$2:$B$3');
    const second = patchSpreadsheet(first, book, 1, {
      definedNames: { SALESRANGE: 'Data!$A$2:$A$3' },
    });
    names = XLSX.read(Buffer.from(second, 'base64')).Workbook!.Names!;
    expect(names.find((n) => n.Sheet === undefined)!.Ref).toBe('Data!$A$2:$A$3');
    expect(names.find((n) => n.Sheet === 0)!.Ref).toBe('Data!$B$2:$B$3');
  });
  it('resizes native tables and filters while preserving IDs, styles and calculated columns', () => {
    const original = fixture(),
      book = XLSX.read(Buffer.from(original, 'base64'), { type: 'buffer', xlfn: true });
    const table = readSpreadsheetStructure(original).Data.tables[0];
    const result = patchSpreadsheet(original, book, 0, {
      tables: [{ ...table, ref: 'A1:E8', columns: [...table.columns, 'Extra'] }],
    });
    expect(readSpreadsheetStructure(result).Data.tables[0]).toEqual({
      ...table,
      ref: 'A1:E8',
      columns: [...table.columns, 'Extra'],
      columnIds: [1, 2, 3, 4, 5],
    });
    const xml = strFromU8(unzipSync(Buffer.from(result, 'base64'))['xl/tables/table1.xml']);
    expect(xml).toContain('autoFilter ref="A1:E8"');
    expect(xml).toContain('tableStyleInfo');
    expect(xml).toContain('calculatedColumnFormula');
    expect(xml).toContain('tableColumn id="4" name="Total"');
    expect(xml).toContain('tableColumn id="5" name="Extra"');
    const shrunk = patchSpreadsheet(result, book, 0, {
      tables: [
        { ...table, ref: 'A1:C3', columns: table.columns.slice(0, 3), columnIds: [1, 2, 3] },
      ],
    });
    expect(readSpreadsheetStructure(shrunk).Data.tables[0].columns).toHaveLength(3);
    expect(strFromU8(unzipSync(Buffer.from(shrunk, 'base64'))['xl/tables/table1.xml'])).toContain(
      'autoFilter ref="A1:C3"',
    );
  });
  it('reads imported table definitions and XLDAPR dynamic anchors', () => {
    const metadata = readSpreadsheetStructure(fixture());
    expect(metadata.Data.tables[0]).toEqual({
      id: 1,
      columnIds: [1, 2, 3, 4],
      name: 'Sales',
      sheet: 'Data',
      ref: 'A1:D4',
      headerRows: 1,
      totalRows: 0,
      columns: ['Region', 'Amount', 'Units', 'Total'],
    });
    expect(metadata.Dashboard.arrays.G1).toEqual({ ref: 'G1:H3', dynamic: true });
  });
  it('updates array bounds, reuses existing cell metadata and preserves unrelated native table XML', () => {
    const original = fixture(),
      book = XLSX.read(Buffer.from(original, 'base64'), { type: 'buffer', xlfn: true });
    book.Sheets.Dashboard.G1 = { t: 'n', v: 10, f: '_xlfn.SEQUENCE(2,2,10,5)' };
    delete book.Sheets.Dashboard.G3;
    delete book.Sheets.Dashboard.H3;
    const result = patchSpreadsheet(original, book, 1, {
      spills: [{ row: 0, col: 6, rows: 2, cols: 2 }],
    });
    expect(readSpreadsheetStructure(result).Dashboard.arrays.G1).toEqual({
      ref: 'G1:H2',
      dynamic: true,
    });
    const before = unzipSync(Buffer.from(original, 'base64')),
      after = unzipSync(Buffer.from(result, 'base64'));
    expect(strFromU8(after['xl/tables/table1.xml'])).toBe(
      strFromU8(before['xl/tables/table1.xml']),
    );
    expect(strFromU8(after['xl/metadata.xml']).match(/<metadataType /g)).toHaveLength(1);
    expect(readSpreadsheetStructure(result).Dashboard.arrays.A1.dynamic).toBe(true);
  });
  it('creates native dynamic metadata on a plain workbook without replacing other relationships', () => {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([[1], [2]]), 'Data');
    const original = Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString(
      'base64',
    );
    book.Sheets.Data.A1.f = '_xlfn.SEQUENCE(2)';
    const result = patchSpreadsheet(original, book, 0, {
      spills: [{ row: 0, col: 0, rows: 2, cols: 1 }],
    });
    expect(readSpreadsheetStructure(result).Data.arrays.A1).toEqual({
      ref: 'A1:A2',
      dynamic: true,
    });
    const archive = unzipSync(Buffer.from(result, 'base64'));
    expect(strFromU8(archive['xl/_rels/workbook.xml.rels'])).toContain('/sheetMetadata');
    expect(strFromU8(archive['[Content_Types].xml'])).toContain('sheetMetadata+xml');
  });
  it('writes modern spill errors as native error values rather than strings or numeric error codes', () => {
    const original = fixture(),
      book = XLSX.read(Buffer.from(original, 'base64'), { type: 'buffer', xlfn: true });
    book.Sheets.Dashboard.G1 = { t: 's', v: '#SPILL!', f: '_xlfn.SEQUENCE(2)' };
    const result = patchSpreadsheet(original, book, 1, {
      spills: [{ row: 0, col: 6, rows: 1, cols: 1 }],
    });
    expect(readSpreadsheetStructure(result).Dashboard.errors.G1).toBe('#SPILL!');
    expect(readSpreadsheetStructure(result).Dashboard.arrays.G1.ref).toBe('G1');
  });
});
