import { describe, it, expect } from 'vitest';
import * as XLSX from 'xlsx';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  calculatePivot,
  saveSpreadsheetPivot,
  readSpreadsheetPivots,
  refreshSpreadsheetPivots,
  PivotDefinition,
} from '../electron/spreadsheet-pivot';
import { patchSpreadsheet } from '../electron/spreadsheet-edit';

/** Construye un libro aislado; no recibe argumentos y retorna XLSX Base64 con datos y recursos centinela. */
function fixture(): string {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    book,
    XLSX.utils.aoa_to_sheet([
      ['Region', 'Product', 'Seller', 'Amount', 'Units'],
      ['North', 'A', 'Ana', 100, 2],
      ['North', 'B', 'Luis', 50, 3],
      ['South', 'A', 'Ana', 200, 4],
      ['North', 'A', 'Ana', 150, 6],
      ['South', 'B', null, null, 1],
    ]),
    'Sales',
  );
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Keep me']]), 'Other');
  const archive = unzipSync(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
  archive['xl/charts/chart1.xml'] = strToU8('<chart>preserve</chart>');
  archive['xl/media/image1.png'] = new Uint8Array([1, 2, 3]);
  return Buffer.from(zipSync(archive)).toString('base64');
}
/** Retorna una configuración nueva para que las pruebas no compartan arrays mutables. */
function definition(): PivotDefinition {
  return {
    name: 'Summary',
    sourceSheet: 'Sales',
    sourceRange: 'A1:E6',
    rows: [0],
    columns: [1],
    values: [{ field: 3, aggregate: 'sum' }],
    filters: [],
    totals: true,
  };
}

describe('native spreadsheet pivot tables', () => {
  it('does not aggregate numeric BIFF error codes as amounts and preserves the previous result with a warning', () => {
    const saved = saveSpreadsheetPivot(fixture(), definition());
    const book = XLSX.read(Buffer.from(saved.base64, 'base64'), { type: 'buffer' });
    book.Sheets.Sales.D2 = { t: 'e', v: 7, f: '1/0' };
    const result = refreshSpreadsheetPivots(patchSpreadsheet(saved.base64, book, 0));
    expect(result.warnings).toHaveLength(1);
    expect(readSpreadsheetPivots(result.base64)[0].error).toBe('PIVOT_ERRORS');
    expect(
      XLSX.read(Buffer.from(result.base64, 'base64'), { type: 'buffer' }).Sheets.Summary.D4.v,
    ).toBe(500);
  });
  it('aggregates a row/column cross and calculates both grand totals', () => {
    const result = calculatePivot(fixture(), definition());
    expect(result.matrix).toEqual([
      ['Region', 'A', 'B', 'Total general'],
      ['North', 250, 50, 300],
      ['South', 200, 0, 200],
      ['Total general', 450, 50, 500],
    ]);
    expect(result.matched).toBe(5);
  });
  it('calculates weighted grand averages from records rather than averaging group averages', () => {
    const config = definition();
    config.values = [{ field: 3, aggregate: 'average' }];
    const matrix = calculatePivot(fixture(), config).matrix;
    expect(matrix[1]).toEqual(['North', 125, 50, 100]);
    expect(matrix.at(-1)).toEqual(['Total general', 150, 50, 125]);
  });
  it('supports multiple values, row levels, typed filters and min/max/count variants', () => {
    const config = definition();
    config.rows = [0, 1];
    config.columns = [];
    config.filters = [{ field: 2, value: 'Ana' }];
    config.values = [
      { field: 3, aggregate: 'count' },
      { field: 3, aggregate: 'countNums' },
      { field: 3, aggregate: 'min' },
      { field: 3, aggregate: 'max' },
    ];
    const result = calculatePivot(fixture(), config);
    expect(result.matched).toBe(3);
    expect(result.matrix[1]).toEqual(['North', 'A', 2, 2, 100, 150]);
    expect(result.matrix.at(-1)).toEqual(['Total general', null, 3, 3, 100, 200]);
  });
  it('keeps text, booleans and numeric categories distinct and excludes completely blank records', () => {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      book,
      XLSX.utils.aoa_to_sheet([
        ['Key', 'Value'],
        [1, 2],
        ['1', 4],
        [true, 8],
        [null, null],
      ]),
      'Typed',
    );
    const base64 = Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString(
      'base64',
    );
    const result = calculatePivot(base64, {
      ...definition(),
      sourceSheet: 'Typed',
      sourceRange: 'A1:B5',
      columns: [],
      values: [{ field: 1, aggregate: 'sum' }],
    });
    expect(result.rows).toHaveLength(3);
    expect(result.matched).toBe(3);
    expect(result.matrix.at(-1)).toEqual(['Total general', 14]);
  });
  it('validates headers, ranges, conflicting axes and the result sheet name before saving', () => {
    expect(() => calculatePivot(fixture(), { ...definition(), sourceRange: 'A0:E6' })).toThrow(
      'PIVOT_RANGE',
    );
    expect(() =>
      calculatePivot(fixture(), { ...definition(), sourceRange: 'A2:E6' }),
    ).not.toThrow();
    expect(() => calculatePivot(fixture(), { ...definition(), columns: [0] })).toThrow(
      'PIVOT_FIELDS',
    );
    expect(() => saveSpreadsheetPivot(fixture(), { ...definition(), name: 'Other' })).toThrow(
      'PIVOT_DUPLICATE',
    );
    expect(() => calculatePivot(fixture(), { ...definition(), sourceRange: 'A1:F6' })).toThrow(
      'PIVOT_HEADERS',
    );
  });
  it('creates native cache records, field definitions, workbook relationships and a dedicated output sheet', () => {
    const original = fixture(),
      saved = saveSpreadsheetPivot(original, definition());
    const entries = unzipSync(Buffer.from(saved.base64, 'base64')),
      before = unzipSync(Buffer.from(original, 'base64'));
    for (const name of [
      'xl/worksheets/sheet1.xml',
      'xl/worksheets/sheet2.xml',
      'xl/charts/chart1.xml',
      'xl/media/image1.png',
    ])
      expect(entries[name]).toEqual(before[name]);
    expect(strFromU8(entries[saved.pivot.tablePath])).toContain('pivotTableDefinition');
    expect(strFromU8(entries[saved.pivot.cachePath])).toContain(
      'worksheetSource ref="A1:E6" sheet="Sales"',
    );
    expect(strFromU8(entries[saved.pivot.recordsPath])).toContain('count="5"');
    const book = XLSX.read(Buffer.from(saved.base64, 'base64'), { type: 'buffer' });
    expect(book.SheetNames).toContain('Summary');
    expect(book.Sheets.Summary.D4.v).toBe(500);
    expect(readSpreadsheetPivots(saved.base64)[0].rows).toEqual([0]);
  });
  it('edits an existing pivot without adding sheets, cache IDs or duplicate table parts', () => {
    const saved = saveSpreadsheetPivot(fixture(), definition());
    const edited = saveSpreadsheetPivot(saved.base64, {
      ...saved.pivot,
      values: [{ field: 3, aggregate: 'average' }],
    });
    expect(edited.pivot.cacheId).toBe(saved.pivot.cacheId);
    expect(edited.pivot.tablePath).toBe(saved.pivot.tablePath);
    const book = XLSX.read(Buffer.from(edited.base64, 'base64'), { type: 'buffer' });
    expect(book.SheetNames).toHaveLength(3);
    expect(book.Sheets.Summary.D4.v).toBe(125);
    expect(readSpreadsheetPivots(edited.base64)).toHaveLength(1);
  });
  it('refreshes after source edits while preserving pivot parts across regular worksheet saves', () => {
    const saved = saveSpreadsheetPivot(fixture(), definition());
    const book = XLSX.read(Buffer.from(saved.base64, 'base64'), { type: 'buffer' });
    book.Sheets.Sales.D2.v = 300;
    const updated = patchSpreadsheet(saved.base64, book, 0),
      refreshed = refreshSpreadsheetPivots(updated);
    expect(refreshed.warnings).toEqual([]);
    expect(
      XLSX.read(Buffer.from(refreshed.base64, 'base64'), { type: 'buffer' }).Sheets.Summary.D4.v,
    ).toBe(700);
    const entries = unzipSync(Buffer.from(refreshed.base64, 'base64'));
    expect(strFromU8(entries[readSpreadsheetPivots(refreshed.base64)[0].cachePath])).toContain(
      '<n v="300"/>',
    );
  });
  it('retains the last valid result and warns when a source becomes invalid', () => {
    const saved = saveSpreadsheetPivot(fixture(), definition());
    const book = XLSX.read(Buffer.from(saved.base64, 'base64'), { type: 'buffer' });
    book.Sheets.Sales.A1.v = '';
    const refreshed = refreshSpreadsheetPivots(patchSpreadsheet(saved.base64, book, 0));
    expect(refreshed.warnings).toHaveLength(1);
    expect(readSpreadsheetPivots(refreshed.base64)[0].error).toBe('PIVOT_HEADERS');
    expect(
      XLSX.read(Buffer.from(refreshed.base64, 'base64'), { type: 'buffer' }).Sheets.Summary.D4.v,
    ).toBe(500);
  });
  it('recovers its native definition when another editor removes private metadata', () => {
    const config = definition();
    config.filters = [{ field: 2, value: 'Ana' }];
    const saved = saveSpreadsheetPivot(fixture(), config),
      archive = unzipSync(Buffer.from(saved.base64, 'base64'));
    delete archive['helperNotes/pivots.json'];
    const reopened = Buffer.from(zipSync(archive)).toString('base64');
    const recovered = readSpreadsheetPivots(reopened)[0];
    expect(recovered.filters).toEqual(config.filters);
    expect(recovered.sourceRange).toBe('A1:E6');
    expect(refreshSpreadsheetPivots(reopened).warnings).toEqual([]);
  });
  it('preserves unrelated cells added to the output sheet by another editor', () => {
    const saved = saveSpreadsheetPivot(fixture(), definition());
    const book = XLSX.read(Buffer.from(saved.base64, 'base64'), { type: 'buffer' });
    book.Sheets.Summary.Z20 = { t: 's', v: 'External annotation' };
    book.Sheets.Summary['!ref'] = 'A1:Z20';
    const edited = patchSpreadsheet(saved.base64, book, 2),
      refreshed = refreshSpreadsheetPivots(edited);
    const reopened = XLSX.read(Buffer.from(refreshed.base64, 'base64'), { type: 'buffer' });
    expect(reopened.Sheets.Summary.Z20.v).toBe('External annotation');
    expect(reopened.Sheets.Summary.D4.v).toBe(500);
  });
});
