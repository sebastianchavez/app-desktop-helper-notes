import { describe, it, expect } from 'vitest';
import * as XLSX from 'xlsx';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { patchSpreadsheet } from '../electron/spreadsheet-edit';
import { readSpreadsheetStructure, NativeTable } from '../electron/spreadsheet-structure';

/** No recibe argumentos; retorna libro sencillo y paquete original con dos hojas para comprobar aislamiento de partes. */
function fixture() {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    book,
    XLSX.utils.aoa_to_sheet([
      ['Product', 'Qty', 'Total'],
      ['A', 2, 10],
      ['B', 3, 15],
      ['Total', 5, 25],
    ]),
    'Data',
  );
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['keep']]), 'Other');
  return {
    book,
    base64: Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString('base64'),
  };
}
const table: NativeTable = {
  id: 1,
  columnIds: [1, 2, 3],
  name: 'Orders',
  sheet: 'Data',
  ref: 'A1:C4',
  columns: ['Product', 'Qty', 'Total'],
  headerRows: 1,
  totalRows: 1,
  totals: ['', 'sum', 'sum'],
  calculated: { Total: '[@Qty]*5' },
};
describe('table package lifecycle', () => {
  it('updates imported pivot cache sources that refer to a renamed table and requests refresh on open', () => {
    const { book, base64 } = fixture(),
      created = patchSpreadsheet(base64, book, 0, { tables: [table] }),
      parts = unzipSync(Buffer.from(created, 'base64'));
    const part = 'xl/pivotCache/pivotCacheDefinition1.xml';
    parts[part] = strToU8(
      '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource name="Orders"/></cacheSource></pivotCacheDefinition>',
    );
    const result = patchSpreadsheet(Buffer.from(zipSync(parts)).toString('base64'), book, 0, {
      tables: [{ ...table, name: 'Archive' }],
    });
    const xml = strFromU8(unzipSync(Buffer.from(result, 'base64'))[part]);
    expect(xml).toContain('name="Archive"');
    expect(xml).toContain('refreshOnLoad="1"');
  });
  it('creates native relationships, column formulas, styles and total metadata without changing another worksheet', () => {
    const { book, base64 } = fixture(),
      result = patchSpreadsheet(base64, book, 0, { tables: [table] }),
      parts = unzipSync(Buffer.from(result, 'base64'));
    expect(readSpreadsheetStructure(result).Data.tables[0]).toEqual({
      id: 1,
      columnIds: [1, 2, 3],
      name: 'Orders',
      sheet: 'Data',
      ref: 'A1:C4',
      columns: ['Product', 'Qty', 'Total'],
      headerRows: 1,
      totalRows: 1,
      totals: ['', 'sum', 'sum'],
    });
    const xml = strFromU8(parts['xl/tables/table1.xml']);
    expect(xml).toContain('autoFilter ref="A1:C3"');
    expect(xml).toContain('totalsRowFunction="sum"');
    expect(xml).toContain('<calculatedColumnFormula>[@Qty]*5</calculatedColumnFormula>');
    expect(strFromU8(parts['[Content_Types].xml'])).toContain('/xl/tables/table1.xml');
    expect(strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels'])).toContain(
      '../tables/table1.xml',
    );
    expect(strFromU8(parts['xl/worksheets/sheet2.xml'])).toBe(
      strFromU8(unzipSync(Buffer.from(base64, 'base64'))['xl/worksheets/sheet2.xml']),
    );
  });
  it('keeps IDs through repeated table/column renames, column reordering and custom totals', () => {
    const { book, base64 } = fixture();
    let result = patchSpreadsheet(base64, book, 0, { tables: [table] });
    result = patchSpreadsheet(result, book, 0, {
      tables: [
        {
          ...table,
          name: 'Ventas',
          sourceName: 'Orders',
          columns: ['Product', 'Importe', 'Qty'],
          columnIds: [1, 3, 2],
          totals: ['', 'custom', 'sum'],
          totalFormulas: { Importe: 'SUM(B2:B3)*2' },
          calculated: { Importe: '[@Qty]*5' },
        },
      ],
    });
    const xml = strFromU8(unzipSync(Buffer.from(result, 'base64'))['xl/tables/table1.xml']);
    expect(xml).toContain('id="3" name="Importe"');
    expect(xml).toContain('<totalsRowFormula>SUM(B2:B3)*2</totalsRowFormula>');
    result = patchSpreadsheet(result, book, 0, {
      tables: [
        {
          ...table,
          name: 'Archive',
          sourceName: 'Orders',
          columns: ['Product', 'Amount', 'Qty'],
          columnIds: [1, 3, 2],
        },
      ],
    });
    expect(readSpreadsheetStructure(result).Data.tables[0].name).toBe('Archive');
    expect(readSpreadsheetStructure(result).Data.tables[0].columnIds).toEqual([1, 3, 2]);
  });
  it('removes a deleted table and its native relationships while retaining cells', () => {
    const { book, base64 } = fixture(),
      created = patchSpreadsheet(base64, book, 0, { tables: [table] }),
      removed = patchSpreadsheet(created, book, 0, { tables: [] });
    const parts = unzipSync(Buffer.from(removed, 'base64'));
    expect(parts['xl/tables/table1.xml']).toBeUndefined();
    expect(readSpreadsheetStructure(removed).Data.tables).toEqual([]);
    expect(strFromU8(parts['xl/worksheets/sheet1.xml'])).not.toContain('tableParts');
    expect(XLSX.read(Buffer.from(removed, 'base64')).Sheets.Data.C2.v).toBe(10);
  });
  it('round-trips imported hidden rows and their structural relocation', () => {
    const { book, base64 } = fixture();
    const saved = patchSpreadsheet(base64, book, 0, { hiddenRows: [1] });
    expect(readSpreadsheetStructure(saved).Data.hiddenRows).toEqual([1]);
    const moved = patchSpreadsheet(saved, book, 0, { hiddenRows: [2] });
    expect(readSpreadsheetStructure(moved).Data.hiddenRows).toEqual([2]);
  });
});
