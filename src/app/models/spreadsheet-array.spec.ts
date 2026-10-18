import { describe, it, expect } from 'vitest';
import {
  evaluateSpreadsheetResult,
  evaluateSpreadsheetFormula,
  ooxmlSpreadsheetFormula,
  translateSpreadsheetFormula,
  canonicalSpreadsheetFormula,
} from './spreadsheet-formula';
import {
  recalculateSpreadsheetWorkbook,
  SpreadsheetCalculationSheet,
} from './spreadsheet-workbook';
import { resolveSpreadsheetTable, SpreadsheetTable } from './spreadsheet-table';

const cells: Record<string, string | number | boolean | null> = {
  A1: 1,
  A2: 2,
  A3: 3,
  B1: 'North',
  B2: 'South',
  B3: 'North',
  C1: 10,
  C2: 20,
  C3: 30,
};
/** Recibe expresión; retorna su matriz/escalar con tres filas ocupadas y lectura de celdas aislada. */
function run(formula: string) {
  return evaluateSpreadsheetResult(formula, (ref) => cells[ref.replace(/\$/g, '')] ?? null, {
    usedRows: 3,
  });
}
describe('dynamic array calculations', () => {
  it('generates rectangular sequences and supports arithmetic broadcasting', () => {
    expect(run('=SEQUENCE(2;3;5;2)')).toEqual([
      [5, 7, 9],
      [11, 13, 15],
    ]);
    expect(run('=A1:A3*2')).toEqual([[2], [4], [6]]);
    expect(run('=A1:A3>1')).toEqual([[false], [true], [true]]);
    expect(run('=TRANSPOSE(A1:A3)')).toEqual([[1, 2, 3]]);
    expect(run('=IF(A1:A3>1;A1:A3;99)')).toEqual([[99], [2], [3]]);
    expect(run('=IFERROR(A1:A3;1/0)')).toEqual([[1], [2], [3]]);
  });
  it('filters rows with combined conditions and evaluates the empty fallback only when needed', () => {
    expect(run('=FILTER(A1:C3;(B1:B3="North")*(C1:C3>10))')).toEqual([[3, 'North', 30]]);
    expect(run('=FILTER(A1:A3;A1:A3>5;"Nada")')).toBe('Nada');
    expect(run('=FILTER(A1:A3;A1:A3>0;1/0)')).toEqual([[1], [2], [3]]);
    expect(run('=FILTER(A1:A3;A1:A3>5)')).toBe('#CALC!');
  });
  it('sorts by a selected column, preserves ties and extracts unique rows', () => {
    expect(run('=SORT(A1:C3;3;-1)')).toEqual([
      [3, 'North', 30],
      [2, 'South', 20],
      [1, 'North', 10],
    ]);
    expect(run('=UNIQUE(B1:B3)')).toEqual([['North'], ['South']]);
    expect(run('=UNIQUE(B1:B3;FALSE;TRUE)')).toEqual([['South']]);
    expect(run('=SORT(UNIQUE(B1:B3))')).toEqual([['North'], ['South']]);
  });
  it('returns multiple lookup columns and INDEX slices', () => {
    expect(run('=XLOOKUP(2;A1:A3;B1:C3)')).toEqual([['South', 20]]);
    expect(run('=INDEX(A1:C3;2;0)')).toEqual([[2, 'South', 20]]);
  });
  it('validates dimensions and output limits without allocating sheet-sized results', () => {
    expect(run('=SEQUENCE(100001)')).toBe('#NUM!');
    expect(run('=FILTER(A1:A3;C1:C2>0)')).toBe('#VALUE!');
    expect(run('=A1:A3+C1:C2')).toBe('#VALUE!');
  });
  it('exports spill references and future functions using native OOXML names', () => {
    expect(ooxmlSpreadsheetFormula("=SUM('Data'!A1#)+FILTER(A1:A3;A1:A3>1)")).toBe(
      "=SUM(_xlfn.ANCHORARRAY('Data'!A1))+_xlfn._xlws.FILTER(A1:A3,A1:A3>1)",
    );
  });
  it('copies complete ranges while keeping structured column names and quoted literals unchanged', () => {
    expect(translateSpreadsheetFormula('=SUM(A:A)+SUM($B:$C)+SUM(1:2)+Sales[A1]', 1, 1)).toBe(
      '=SUM(B:B)+SUM($B:$C)+SUM(2:3)+Sales[A1]',
    );
    expect(canonicalSpreadsheetFormula('=SUMA(Sales[Cost;Total])')).toBe('=SUM(Sales[Cost;Total])');
  });
});
describe('whole row and column ranges', () => {
  it('aggregates used values while keeping the logical blank tail', () => {
    expect(run('=SUM(A:A)')).toBe(6);
    expect(run('=AVERAGE(A:A)')).toBe(2);
    expect(run('=COUNT(A:A)')).toBe(3);
    expect(run('=ROWS(A:A)')).toBe(1048576);
    expect(run('=COUNTBLANK(A:A)')).toBe(1048573);
    expect(run('=COUNTIF(A:A;"")')).toBe(1048573);
    expect(run('=COUNTIFS(A:A;"";B:B;"")')).toBe(1048573);
    expect(run('=SUMIFS(C:C;B:B;"North")')).toBe(40);
    expect(run('=SUMPRODUCT(A:A;C:C)')).toBe(140);
  });
  it('includes row width, last-row lookups and qualified complete ranges', () => {
    expect(run('=SUM(1:1)')).toBe(11);
    expect(run('=COLUMNS(1:1)')).toBe(16384);
    expect(run('=XMATCH("";A:A;0;-1)')).toBe(1048576);
    expect(run('=INDEX(A:A;1048576)')).toBe(0);
  });
});
const table: SpreadsheetTable = {
  name: 'Sales',
  sheet: 'Data',
  ref: 'A1:C5',
  headerRows: 1,
  totalRows: 1,
  columns: ['Region', 'Amount', 'Units'],
};
describe('structured references', () => {
  it('resolves data, headers, totals, column spans and current rows', () => {
    expect(resolveSpreadsheetTable('[[#This Row],Amount]', [table], 'Data', 2, 1)).toBe(
      "'Data'!$B$3",
    );
    expect(resolveSpreadsheetTable('Sales[Amount]', [table], 'Data', 2)).toBe("'Data'!$B$2:$B$4");
    expect(resolveSpreadsheetTable('Sales[[#Headers],[Amount]]', [table], 'Data', 2)).toBe(
      "'Data'!$B$1:$B$1",
    );
    expect(resolveSpreadsheetTable('Sales[[#Totals],[Amount]]', [table], 'Data', 2)).toBe(
      "'Data'!$B$5:$B$5",
    );
    expect(resolveSpreadsheetTable('Sales[[Amount]:[Units]]', [table], 'Data', 2)).toBe(
      "'Data'!$B$2:$C$4",
    );
    expect(resolveSpreadsheetTable('[@Amount]', [table], 'Data', 2)).toBe("'Data'!$B$3");
  });
  it('calculates imported table expressions without rewriting them into plain A1 formulas', () => {
    expect(
      evaluateSpreadsheetFormula('=SUM(Sales[Amount])', (ref) => (/!B\d+$/.test(ref) ? 10 : null), {
        resolveStructured: (ref) => resolveSpreadsheetTable(ref, [table], 'Data', 2),
      }),
    ).toBe(30);
    expect(resolveSpreadsheetTable('Sales[Missing]', [table], 'Data', 2)).toBe('#REF!');
  });
});
describe('workbook spill ownership', () => {
  /** Recibe datos y fórmulas; retorna hojas mutables y el resultado de un cálculo aislado. */
  function calculate(data: (string | number | boolean | null)[][], formulas: [string, string][]) {
    const sheets: SpreadsheetCalculationSheet[] = [
      { name: 'Data', data: structuredClone(data), colCount: 6 },
    ];
    const expressions = new Map(formulas);
    const result = recalculateSpreadsheetWorkbook(sheets, expressions, new Map(), new Map());
    return { sheets, expressions, result };
  }
  it('populates children, resolves A1# dependencies and supports a scalar that precedes its new spill parent', () => {
    const { sheets, result } = calculate(
      [],
      [
        ['0:0:3', '=A2+10'],
        ['0:0:0', '=SEQUENCE(3)'],
        ['0:0:5', '=SUM(A1#)'],
      ],
    );
    expect(sheets[0].data[1][0]).toBe(2);
    expect(sheets[0].data[0][3]).toBe(12);
    expect(sheets[0].data[0][5]).toBe(6);
    expect(result.owners.get('0:1:0')).toBe('0:0:0');
  });
  it('preserves a blocking value and a merge instead of overwriting either', () => {
    const blocked = calculate([[null], [99]], [['0:0:0', '=SEQUENCE(3)']]);
    expect(blocked.sheets[0].data[0][0]).toBe('#SPILL!');
    expect(blocked.sheets[0].data[1][0]).toBe(99);
    const sheets: SpreadsheetCalculationSheet[] = [
      { name: 'Data', data: [], colCount: 3, merges: [{ s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }] },
    ];
    recalculateSpreadsheetWorkbook(
      sheets,
      new Map([['0:0:0', '=SEQUENCE(3)']]),
      new Map(),
      new Map(),
    );
    expect(sheets[0].data[0][0]).toBe('#SPILL!');
  });
  it('clears obsolete children on shrink and after removing the anchor formula', () => {
    const first = calculate([], [['0:0:0', '=SEQUENCE(3)']]);
    first.expressions.set('0:0:0', '=SEQUENCE(2)');
    const second = recalculateSpreadsheetWorkbook(
      first.sheets,
      first.expressions,
      new Map(),
      first.result.spills,
    );
    expect(first.sheets[0].data[2][0]).toBeNull();
    first.expressions.clear();
    first.sheets[0].data[0][0] = 99;
    recalculateSpreadsheetWorkbook(first.sheets, first.expressions, new Map(), second.spills);
    expect(first.sheets[0].data[0][0]).toBe(99);
    expect(first.sheets[0].data[1][0]).toBeNull();
  });
  it('preserves the entire cached output of an imported dynamic function that is still unsupported', () => {
    const sheets: SpreadsheetCalculationSheet[] = [
      { name: 'Data', data: [[111], [222]], colCount: 1 },
    ];
    const result = recalculateSpreadsheetWorkbook(
      sheets,
      new Map([['0:0:0', '=UNKNOWNARRAY()']]),
      new Map([['0:0:0', { formula: '=UNKNOWNARRAY()', value: 111 }]]),
      new Map([['0:0:0', { sheet: 0, row: 0, col: 0, rows: 2, cols: 1 }]]),
    );
    expect(sheets[0].data).toEqual([[111], [222]]);
    expect(result.cachedCount).toBe(1);
    expect(result.owners.get('0:1:0')).toBe('0:0:0');
  });
});
