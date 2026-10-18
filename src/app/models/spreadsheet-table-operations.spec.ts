import {
  remapSpreadsheetTable,
  validateTableName,
  tableTotalFormula,
} from './spreadsheet-table-operations';
import {
  renameSpreadsheetTableReferences,
  renameSpreadsheetColumnReferences,
  remapSpreadsheetFormula,
} from './spreadsheet-formula';
import { SpreadsheetTable } from './spreadsheet-table';

const table: SpreadsheetTable = {
  name: 'Sales',
  sheet: 'Data',
  ref: 'A1:C5',
  columns: ['Item', 'Units', 'Amount'],
  headerRows: 1,
  totalRows: 1,
  totals: ['', 'countNums', 'sum'],
};
describe('editable native tables and structural references', () => {
  it('renames column references simultaneously, including imported current-row dialects and escaped headings', () => {
    expect(
      renameSpreadsheetColumnReferences(
        'SUM(Sales[A])+SUM(Sales[B])+Other[A]+"Sales[A]"',
        'Sales',
        { A: 'B', B: 'A' },
      ),
    ).toBe('SUM(Sales[B])+SUM(Sales[A])+Other[A]+"Sales[A]"');
    expect(
      renameSpreadsheetColumnReferences(
        '[[#This Row],Amount]*[@Amount]',
        'Sales',
        { Amount: 'Precio' },
        true,
      ),
    ).toBe('[[#This Row],Precio]*[@Precio]');
    expect(renameSpreadsheetColumnReferences('[@Amount]', 'Sales', { Amount: 'Precio' })).toBe(
      '[@Amount]',
    );
    expect(
      renameSpreadsheetColumnReferences('Sales[Amount]', 'Sales', { Amount: 'Price [net]' }),
    ).toBe("Sales[Price '[net']]");
  });
  it('validates workbook-wide names and protects cell references and defined names', () => {
    expect(validateTableName(' Ventas_2026 ', [])).toBe('Ventas_2026');
    for (const name of ['Sales', 'A1', 'R1C1', 'R', 'with space'])
      expect(() => validateTableName(name, [table])).toThrow('TABLE_NAME_ERROR');
    expect(() => validateTableName('Rate', [], undefined, ['RATE'])).toThrow('TABLE_NAME_ERROR');
    expect(validateTableName('Sales', [table], 'Sales')).toBe('Sales');
  });
  it('renames explicit and bare table references without changing text, columns or sheet names', () => {
    expect(
      renameSpreadsheetTableReferences(
        'SUM(Sales[Amount])+SUM(Sales)+LEN("Sales[Amount]")+Sales!A1+Other[Sales]',
        'Sales',
        'Ventas',
      ),
    ).toBe('SUM(Ventas[Amount])+SUM(Ventas)+LEN("Sales[Amount]")+Sales!A1+Other[Sales]');
    expect(renameSpreadsheetTableReferences('SUM(Sales[Amount])', 'Sales', null)).toBe(
      'SUM(#REF!)',
    );
  });
  it('inserts data rows and columns while moving totals and retaining column identities', () => {
    const rows = remapSpreadsheetTable(
      table,
      (r, c) => ({ row: r >= 2 ? r + 1 : r, col: c }),
      true,
    )!;
    expect(rows.ref).toBe('A1:C6');
    expect(rows.totals).toEqual(table.totals);
    const cols = remapSpreadsheetTable(
      table,
      (r, c) => ({ row: r, col: c >= 1 ? c + 1 : c }),
      true,
    )!;
    expect(cols.ref).toBe('A1:D5');
    expect(cols.columns).toEqual(['Item', 'Column2', 'Units', 'Amount']);
    expect(cols.totals).toEqual(['', '', 'countNums', 'sum']);
  });
  it('supports data deletion and whole-table moves but rejects detached rows, headers and empty data', () => {
    expect(
      remapSpreadsheetTable(table, (r, c) => (r === 2 ? null : { row: r > 2 ? r - 1 : r, col: c }))!
        .ref,
    ).toBe('A1:C4');
    expect(remapSpreadsheetTable(table, (r, c) => ({ row: r + 2, col: c + 1 }))!.ref).toBe('B3:D7');
    expect(remapSpreadsheetTable(table, () => null)).toBeNull();
    expect(() =>
      remapSpreadsheetTable(table, (r, c) => (r === 0 ? null : { row: r - 1, col: c })),
    ).toThrow('TABLE_STRUCTURE_ERROR');
    expect(() =>
      remapSpreadsheetTable(table, (r, c) => ({ row: r === 2 ? 10 : r, col: c })),
    ).toThrow('TABLE_STRUCTURE_ERROR');
    expect(() =>
      remapSpreadsheetTable(table, (r, c) =>
        r > 0 && r < 4 ? null : { row: r === 4 ? 1 : r, col: c },
      ),
    ).toThrow('TABLE_STRUCTURE_ERROR');
    expect(tableTotalFormula('sum', 'C2:C4')).toBe('SUBTOTAL(109,C2:C4)');
  });
  it('contracts partially deleted ranges, preserving absolute and qualified endpoints', () => {
    const remove = (r: number, c: number) => (r === 1 ? null : { row: r > 1 ? r - 1 : r, col: c });
    expect(remapSpreadsheetFormula('SUM($A$2:$B$5)+"A2:B5"', 'Data', 'Data', remove)).toBe(
      'SUM($A$2:$B$4)+"A2:B5"',
    );
    expect(remapSpreadsheetFormula("SUM('Data'!A1:B2)", 'Other', 'Data', remove)).toBe(
      "SUM('Data'!A1:B1)",
    );
    expect(remapSpreadsheetFormula('SUM(A2:B2)', 'Data', 'Data', remove)).toBe('SUM(#REF!)');
    expect(remapSpreadsheetFormula('SUM(Other!A2:B5)', 'Data', 'Data', remove)).toBe(
      'SUM(Other!A2:B5)',
    );
    expect(
      remapSpreadsheetFormula('SUM(A1:C5)', 'Data', 'Data', (r, c) =>
        c === 0 ? null : { row: r, col: c - 1 },
      ),
    ).toBe('SUM(A1:B5)');
    expect(
      remapSpreadsheetFormula('SUM($A:$C)', 'Data', 'Data', (r, c) =>
        c === 0 ? null : { row: r, col: c - 1 },
      ),
    ).toBe('SUM($A:$B)');
    expect(
      remapSpreadsheetFormula('SUM(A:C)', 'Data', 'Data', (r, c) =>
        r === 0 ? null : { row: r - 1, col: c },
      ),
    ).toBe('SUM(A:C)');
    expect(
      remapSpreadsheetFormula('SUM(1:5)', 'Data', 'Data', (r, c) =>
        c === 0 ? null : { row: r, col: c - 1 },
      ),
    ).toBe('SUM(1:5)');
  });
});
