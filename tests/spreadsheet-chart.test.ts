import { describe, it, expect } from 'vitest';
import * as XLSX from 'xlsx';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import {
  buildSpreadsheetChartData,
  saveSpreadsheetChart,
  readSpreadsheetCharts,
  refreshSpreadsheetCharts,
  deleteSpreadsheetChart,
  ChartDefinition,
  ChartType,
} from '../electron/spreadsheet-chart';
import { patchSpreadsheet } from '../electron/spreadsheet-edit';
import {
  saveSpreadsheetPivot,
  readSpreadsheetPivots,
  refreshSpreadsheetPivots,
} from '../electron/spreadsheet-pivot';

/** No recibe argumentos; retorna un libro aislado con encabezados, huecos, negativos y una hoja con apóstrofo. */
function fixture(): string {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    book,
    XLSX.utils.aoa_to_sheet([
      ['Month', 'Revenue', 'Cost', 'X'],
      ['Jan', 100, 30, 1],
      ['Feb', 200, null, 2],
      ['Mar', 50, -10, 3],
    ]),
    "Sales '26",
  );
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Keep']]), 'Dashboard');
  const archive = unzipSync(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
  archive['xl/media/sentinel.png'] = new Uint8Array([1, 2, 3]);
  return Buffer.from(zipSync(archive)).toString('base64');
}
/** Recibe tipo opcional; retorna una definición nueva sin estado compartido entre pruebas. */
function definition(type: ChartType = 'column'): ChartDefinition {
  return {
    hostSheet: 'Dashboard',
    sourceSheet: "Sales '26",
    sourceRange: 'A1:D4',
    title: 'Revenue & costs <2026>',
    type,
    headers: true,
    category: 0,
    series: [1, 2],
    legend: true,
    labels: true,
    stacked: false,
    xTitle: 'Month',
    yTitle: 'Amount',
    colors: ['#2563eb', '#16a34a'],
    anchor: { row: 2, col: 1, offsetX: 12, offsetY: 6, width: 600, height: 360 },
  };
}
/** Recibe ZIP y ruta; retorna XML sin descomprimir otros archivos para las aserciones. */
function part(base64: string, path: string): string {
  return strFromU8(unzipSync(Buffer.from(base64, 'base64'))[path]);
}

describe('native spreadsheet charts', () => {
  it('uses numeric cells, preserves gaps and escapes categories without interpreting markup', () => {
    const data = buildSpreadsheetChartData(fixture(), definition());
    expect(data.categories).toEqual(['Jan', 'Feb', 'Mar']);
    expect(data.series.map((s) => s.values)).toEqual([
      [100, 200, 50],
      [30, null, -10],
    ]);
    expect(data.series.map((s) => s.name)).toEqual(['Revenue', 'Cost']);
  });
  for (const type of ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'] as ChartType[])
    it(`roundtrips native ${type} with linked ranges and geometry`, () => {
      const d = definition(type);
      if (type === 'pie' || type === 'doughnut') d.series = [1];
      if (type === 'scatter') d.category = 3;
      const saved = saveSpreadsheetChart(fixture(), d),
        read = readSpreadsheetCharts(saved.base64);
      expect(read).toHaveLength(1);
      expect(read[0].type).toBe(type);
      expect(read[0].title).toBe(d.title);
      expect(read[0].anchor).toEqual(d.anchor);
      expect(read[0].sourceRange).toBe(d.sourceRange);
      expect(read[0].data).toEqual(saved.chart.data);
      expect(part(saved.base64, saved.chart.chartPath)).toContain("'Sales ''26'!$B$2:$B$4");
      expect(part(saved.base64, saved.chart.chartPath)).toContain(
        'Revenue &amp; costs &lt;2026&gt;',
      );
      expect(unzipSync(Buffer.from(saved.base64, 'base64'))['xl/media/sentinel.png']).toEqual(
        new Uint8Array([1, 2, 3]),
      );
    });
  it('accepts lowercase/absolute ranges and retains unused source columns for later editing', () => {
    const d = definition();
    d.sourceRange = '$a$1:$d$4';
    d.series = [1];
    const saved = saveSpreadsheetChart(fixture(), d);
    expect(readSpreadsheetCharts(saved.base64)[0].sourceRange).toBe('A1:D4');
    expect(part(saved.base64, saved.chart.chartPath)).toContain('$B$2:$B$4');
  });
  it('edits an object without duplicating anchors or chart parts and deletes only that object', () => {
    const first = saveSpreadsheetChart(fixture(), definition());
    const second = saveSpreadsheetChart(first.base64, {
      ...definition('line'),
      anchor: { ...definition().anchor, col: 8 },
    });
    const edited = saveSpreadsheetChart(second.base64, {
      ...first.chart,
      type: 'area',
      stacked: true,
      title: 'Edited',
    });
    expect(readSpreadsheetCharts(edited.base64)).toHaveLength(2);
    expect(
      Object.keys(unzipSync(Buffer.from(edited.base64, 'base64'))).filter((p) =>
        /^xl\/charts\/chart\d+\.xml$/.test(p),
      ),
    ).toHaveLength(2);
    const removed = deleteSpreadsheetChart(edited.base64, first.chart.id);
    expect(readSpreadsheetCharts(removed).map((c) => c.id)).toEqual([second.chart.id]);
    expect(XLSX.read(Buffer.from(removed, 'base64')).Sheets["Sales '26"].B2.v).toBe(100);
  });
  it('refreshes linked caches from edited cells while preserving native pivot parts', () => {
    const pivot = saveSpreadsheetPivot(fixture(), {
      name: 'Summary',
      sourceSheet: "Sales '26",
      sourceRange: 'A1:D4',
      rows: [0],
      columns: [],
      values: [{ field: 1, aggregate: 'sum' }],
      filters: [],
      totals: true,
    });
    const saved = saveSpreadsheetChart(pivot.base64, definition());
    const book = XLSX.read(Buffer.from(saved.base64, 'base64'));
    book.Sheets["Sales '26"].B2.v = 400;
    const patched = patchSpreadsheet(saved.base64, book, 0);
    const refreshed = refreshSpreadsheetCharts(patched);
    expect(refreshed.warnings).toEqual([]);
    expect(readSpreadsheetCharts(refreshed.base64)[0].data.series[0].values[0]).toBe(400);
    expect(part(refreshed.base64, saved.chart.chartPath)).toContain('<c:v>400</c:v>');
    expect(readSpreadsheetPivots(refreshed.base64)).toHaveLength(1);
    expect(part(refreshed.base64, 'xl/pivotTables/pivotTable1.xml')).toBe(
      part(saved.base64, 'xl/pivotTables/pivotTable1.xml'),
    );
  });
  it('reads native objects without private JSON and preserves foreign chart formatting on ordinary saves', () => {
    const saved = saveSpreadsheetChart(fixture(), definition());
    const archive = unzipSync(Buffer.from(saved.base64, 'base64'));
    delete archive['helperNotes/charts.json'];
    archive[saved.chart.chartPath] = strToU8(
      part(saved.base64, saved.chart.chartPath).replace(
        '<c:lang val="es-CL"/>',
        '<c:lang val="en-US"/>',
      ),
    );
    const foreign = Buffer.from(zipSync(archive)).toString('base64'),
      read = readSpreadsheetCharts(foreign);
    expect(read).toHaveLength(1);
    expect(read[0].sourceRange).toBe('A1:C4');
    expect(part(refreshSpreadsheetCharts(foreign).base64, saved.chart.chartPath)).toContain(
      'en-US',
    );
    const edited = saveSpreadsheetChart(foreign, { ...read[0], title: 'Editable' });
    expect(readSpreadsheetCharts(edited.base64)[0].title).toBe('Editable');
  });
  it('keeps other drawing anchors and does not opt foreign charts into automatic rewriting', () => {
    const first = saveSpreadsheetChart(fixture(), definition());
    const archive = unzipSync(Buffer.from(first.base64, 'base64'));
    delete archive['helperNotes/charts.json'];
    const sentinel =
      '<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="100" cy="100"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="99" name="KeepShape"/><xdr:cNvSpPr/></xdr:nvSpPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>';
    archive[first.chart.drawingPath] = strToU8(
      part(first.base64, first.chart.drawingPath).replace('</xdr:wsDr>', sentinel + '</xdr:wsDr>'),
    );
    const foreign = Buffer.from(zipSync(archive)).toString('base64');
    const added = saveSpreadsheetChart(foreign, definition('line'));
    const refreshed = refreshSpreadsheetCharts(added.base64).base64;
    expect(part(refreshed, first.chart.chartPath)).toBe(part(foreign, first.chart.chartPath));
    expect(part(refreshed, first.chart.drawingPath)).toContain(sentinel);
    expect(readSpreadsheetCharts(refreshed).filter((c) => c.managed)).toHaveLength(1);
    expect(
      part(deleteSpreadsheetChart(refreshed, added.chart.id), first.chart.drawingPath),
    ).toContain(sentinel);
  });
  it('rejects invalid origins and keeps the last valid chart if all values become blank', () => {
    expect(() =>
      buildSpreadsheetChartData(fixture(), { ...definition(), sourceSheet: 'Missing' }),
    ).toThrow('CHART_SOURCE');
    const saved = saveSpreadsheetChart(fixture(), { ...definition(), series: [1] });
    const book = XLSX.read(Buffer.from(saved.base64, 'base64'));
    for (const address of ['B2', 'B3', 'B4']) delete book.Sheets["Sales '26"][address];
    const patched = patchSpreadsheet(saved.base64, book, 0),
      refresh = refreshSpreadsheetCharts(patched);
    expect(refresh.warnings).toHaveLength(1);
    expect(part(refresh.base64, saved.chart.chartPath)).toBe(
      part(saved.base64, saved.chart.chartPath),
    );
    expect(readSpreadsheetCharts(refresh.base64)[0].error).toBe('CHART_NUMBERS');
  });
  it('validates range, duplicate/category series, positions, scatter and pie semantics', () => {
    for (const change of [
      { sourceRange: 'A1:A4' },
      { sourceRange: 'A0:D4' },
      { sourceRange: 'A1:D4000' },
    ])
      expect(() => buildSpreadsheetChartData(fixture(), { ...definition(), ...change })).toThrow();
    for (const series of [[0], [1, 1], []])
      expect(() => buildSpreadsheetChartData(fixture(), { ...definition(), series })).toThrow(
        'CHART_FIELDS',
      );
    expect(() =>
      buildSpreadsheetChartData(fixture(), {
        ...definition(),
        anchor: { ...definition().anchor, width: 10 },
      }),
    ).toThrow('CHART_POSITION');
    expect(() => buildSpreadsheetChartData(fixture(), definition('scatter'))).toThrow(
      'CHART_SCATTER',
    );
    expect(() => buildSpreadsheetChartData(fixture(), definition('pie'))).toThrow(
      'CHART_PIE_SERIES',
    );
    expect(() =>
      buildSpreadsheetChartData(fixture(), { ...definition('pie'), series: [2] }),
    ).toThrow('CHART_PIE_VALUES');
  });
  it('reopens circular charts with more than 64 categories and a compact editable palette', () => {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      book,
      XLSX.utils.aoa_to_sheet([
        ['Category', 'Amount'],
        ...Array.from({ length: 100 }, (_, i) => ['C' + i, i + 1]),
      ]),
      'Data',
    );
    const base64 = Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' })).toString(
      'base64',
    );
    const result = saveSpreadsheetChart(base64, {
      ...definition('pie'),
      hostSheet: 'Data',
      sourceSheet: 'Data',
      sourceRange: 'A1:B101',
      series: [1],
    });
    const read = readSpreadsheetCharts(result.base64)[0];
    expect(read.error).toBeUndefined();
    expect(read.data.categories).toHaveLength(100);
    expect(read.colors.length).toBeLessThanOrEqual(64);
  });
  it('updates a chart whose source is a refreshed pivot result', () => {
    const pivot = saveSpreadsheetPivot(fixture(), {
      name: 'Summary',
      sourceSheet: "Sales '26",
      sourceRange: 'A1:D4',
      rows: [0],
      columns: [],
      values: [{ field: 1, aggregate: 'sum' }],
      filters: [],
      totals: true,
    });
    const chart = saveSpreadsheetChart(pivot.base64, {
      ...definition(),
      sourceSheet: 'Summary',
      sourceRange: 'A1:B4',
      series: [1],
    });
    const book = XLSX.read(Buffer.from(chart.base64, 'base64'));
    book.Sheets["Sales '26"].B2.v = 400;
    const updated = refreshSpreadsheetPivots(patchSpreadsheet(chart.base64, book, 0));
    const result = refreshSpreadsheetCharts(updated.base64);
    expect(result.warnings).toEqual([]);
    expect(readSpreadsheetCharts(result.base64)[0].data.series[0].values).toEqual([200, 400, 50]);
    expect(part(result.base64, chart.chart.chartPath)).toContain('<c:v>400</c:v>');
  });
});
