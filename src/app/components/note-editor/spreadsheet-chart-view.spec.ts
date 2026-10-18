import { describe, expect, it } from 'vitest';
import { spreadsheetChartGeometry } from './spreadsheet-chart-view';
import type { ChartDefinition, ChartData } from '../../models/spreadsheet-chart';

/** Recibe tipo; retorna configuración aislada del renderizador para comprobar geometría sin navegador. */
function definition(type: ChartDefinition['type']): ChartDefinition {
  return {
    hostSheet: 'Data',
    sourceSheet: 'Data',
    sourceRange: 'A1:B4',
    title: 'Test',
    type,
    headers: true,
    category: 0,
    series: [1],
    legend: true,
    labels: true,
    stacked: false,
    xTitle: 'X',
    yTitle: 'Y',
    colors: ['#2563eb'],
    anchor: { row: 0, col: 0, offsetX: 0, offsetY: 0, width: 600, height: 360 },
  };
}
const data: ChartData = {
  categories: ['A', 'B', 'C'],
  x: [1, 2, 10],
  series: [{ name: 'Values', color: '#2563eb', values: [10, null, -5] }],
};
describe('spreadsheet chart geometry', () => {
  it('draws negative and positive columns on opposite sides of zero, skipping gaps', () => {
    const geometry = spreadsheetChartGeometry(definition('column'), data);
    const bars = geometry.marks.filter((m) => m.kind === 'rect');
    expect(bars).toHaveLength(2);
    expect(bars[0].y! + bars[0].height!).toBeCloseTo(geometry.zero);
    expect(bars[1].y).toBeCloseTo(geometry.zero);
    expect(bars.every((m) => m.width! > 0 && m.height! > 0)).toBe(true);
  });
  it('separates line segments across null cells', () => {
    const geometry = spreadsheetChartGeometry(definition('line'), data);
    expect(geometry.marks.filter((m) => m.kind === 'path')).toHaveLength(2);
    expect(geometry.marks.filter((m) => m.kind === 'circle')).toHaveLength(2);
  });
  it('renders a complete single-slice doughnut with a hole and finite coordinates', () => {
    const geometry = spreadsheetChartGeometry(definition('doughnut'), {
      categories: ['All'],
      x: [null],
      series: [{ name: 'Values', color: '#2563eb', values: [100] }],
    });
    expect(geometry.marks).toHaveLength(1);
    expect(geometry.marks[0].d?.match(/ A /g)).toHaveLength(4);
    expect(geometry.marks[0].d).not.toMatch(/NaN|Infinity/);
    expect(geometry.labels[0].text).toBe('100%');
  });
  it('positions scatter points by numeric X distances rather than row order', () => {
    const geometry = spreadsheetChartGeometry(definition('scatter'), {
      ...data,
      series: [{ ...data.series[0], values: [10, 10, 10] }],
    });
    const points = geometry.marks.filter((m) => m.kind === 'circle');
    expect((points[2].cx! - points[1].cx!) / (points[1].cx! - points[0].cx!)).toBeCloseTo(8);
  });
  it('stacks series and supports all-zero values without invalid SVG', () => {
    const d = definition('bar');
    d.stacked = true;
    const g = spreadsheetChartGeometry(d, {
      categories: ['A'],
      x: [null],
      series: [
        { name: 'One', color: '#2563eb', values: [10] },
        { name: 'Two', color: '#16a34a', values: [20] },
      ],
    });
    expect(g.marks[0].x! + g.marks[0].width!).toBeCloseTo(g.marks[1].x!);
    const zeros = spreadsheetChartGeometry(definition('area'), {
      ...data,
      series: [{ ...data.series[0], values: [0, 0, 0] }],
    });
    expect(JSON.stringify(zeros)).not.toMatch(/NaN|Infinity/);
  });
});
