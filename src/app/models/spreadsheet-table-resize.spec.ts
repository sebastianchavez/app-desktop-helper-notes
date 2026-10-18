import { resizeSpreadsheetTable, tableBounds } from './spreadsheet-table-resize';
import { SpreadsheetTable } from './spreadsheet-table';

const table: SpreadsheetTable = {
  name: 'Sales',
  sheet: 'Data',
  ref: 'A1:B3',
  columns: ['Region', 'Amount'],
  headerRows: 1,
  totalRows: 0,
};
describe('table range resizing', () => {
  it('extends rows and columns while preserving identity and existing headers', () => {
    expect(
      resizeSpreadsheetTable(table, '$A$1:$D$8', [['Region', 'Amount', 'Units']], [table]),
    ).toEqual({ ...table, ref: 'A1:D8', columns: ['Region', 'Amount', 'Units', 'Column4'] });
  });
  it('shrinks the definition without mutating outside values', () => {
    const data = [
      ['Region', 'Amount'],
      ['North', 4],
      ['South', 7],
    ];
    expect(resizeSpreadsheetTable(table, 'A1:A2', data, [table]).columns).toEqual(['Region']);
    expect(data[2]).toEqual(['South', 7]);
  });
  it('rejects moved origins, duplicate headers, overlapping tables and excessive ranges', () => {
    expect(() => resizeSpreadsheetTable(table, 'B1:C8', [], [table])).toThrow('TABLE_RANGE_ERROR');
    expect(() =>
      resizeSpreadsheetTable(table, 'A1:C8', [['Region', 'Amount', 'AMOUNT']], [table]),
    ).toThrow('TABLE_HEADER_ERROR');
    expect(() =>
      resizeSpreadsheetTable(
        table,
        'A1:D8',
        [],
        [table, { ...table, name: 'Other', ref: 'D6:F9' }],
      ),
    ).toThrow('TABLE_OVERLAP');
    expect(() => resizeSpreadsheetTable(table, 'A1:B90000', [], [table])).toThrow(
      'TABLE_RANGE_ERROR',
    );
    expect(tableBounds('A1:XFE10')).toBeNull();
  });
  it('allows resizing totals rows; the editor separately checks that their destination is empty', () => {
    const totals = { ...table, totalRows: 1 };
    expect(resizeSpreadsheetTable(totals, 'A1:B5', [], [totals]).ref).toBe('A1:B5');
    expect(resizeSpreadsheetTable(totals, 'A1:C3', [], [totals]).ref).toBe('A1:C3');
  });
});
