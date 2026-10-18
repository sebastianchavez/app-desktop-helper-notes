import { Component, computed, input } from '@angular/core';
import type { ChartData, ChartDefinition } from '../../models/spreadsheet-chart';

export const CHART_COLORS = [
  '#2563eb',
  '#16a34a',
  '#ea580c',
  '#9333ea',
  '#0891b2',
  '#dc2626',
  '#ca8a04',
  '#db2777',
];
type Mark = {
  kind: 'rect' | 'path' | 'circle';
  color: string;
  tooltip: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  d?: string;
  cx?: number;
  cy?: number;
  r?: number;
  opacity?: number;
  outline?: boolean;
};
type Label = { x: number; y: number; text: string; anchor: 'start' | 'middle' | 'end' };
type ChartGeometry = {
  marks: Mark[];
  labels: Label[];
  ticks: { value: string; position: number }[];
  categories: Label[];
  legend: { name: string; color: string }[];
  horizontal: boolean;
  pie: boolean;
  plot: { left: number; top: number; width: number; height: number };
  zero: number;
};

/** Recibe un número; retorna una etiqueta corta de eje sin modificar el valor real de los datos. */
function numberLabel(value: number): string {
  return new Intl.NumberFormat('es', {
    maximumFractionDigits: 2,
    notation: Math.abs(value) >= 1000000 ? 'compact' : 'standard',
  }).format(value);
}
/** Recibe límites numéricos; retorna un dominio redondeado y seis marcas como máximo aproximado. */
function domain(min: number, max: number): { min: number; max: number; step: number } {
  if (!Number.isFinite(min) || !Number.isFinite(max) || min === max) {
    min = Math.min(0, min || 0);
    max = min + 1;
  }
  const raw = (max - min) / 5,
    power = 10 ** Math.floor(Math.log10(raw)),
    unit = raw / power;
  const step = ([1, 2, 5, 10].find((v) => v >= unit) ?? 10) * power;
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
}

/**
 * Construye geometría SVG para las mismas series que se escriben en el XLSX, incluidos negativos y huecos.
 * @param definition Presentación elegida por el usuario.
 * @param data Valores tipados del libro; null representa una celda no numérica o vacía.
 * @returns Marcas, ejes y leyenda en un lienzo lógico de 640×400, sin HTML ni JavaScript de los datos.
 */
export function spreadsheetChartGeometry(
  definition: ChartDefinition,
  data: ChartData,
): ChartGeometry {
  const pie = definition.type === 'pie' || definition.type === 'doughnut',
    horizontal = definition.type === 'bar';
  const plot = {
    left: horizontal ? 115 : 68,
    top: 48,
    width: horizontal ? 495 : 540,
    height: definition.legend ? 245 : 280,
  };
  const marks: Mark[] = [],
    labels: Label[] = [],
    categories: Label[] = [];
  const legend = pie
    ? data.categories.map((name, i) => ({
        name,
        color: definition.colors[i] ?? CHART_COLORS[i % CHART_COLORS.length],
      }))
    : data.series.map((s) => ({ name: s.name, color: s.color }));
  const result: ChartGeometry = {
    marks,
    labels,
    ticks: [],
    categories,
    legend,
    horizontal,
    pie,
    plot,
    zero: 0,
  };
  if (!data.series.length || !data.categories.length) return result;
  if (pie) {
    const values = data.series[0].values,
      total = values.reduce<number>((sum, v) => sum + (v ?? 0), 0),
      cx = 320,
      cy = 185,
      r = 115,
      inner = definition.type === 'doughnut' ? 63 : 0;
    let angle = -Math.PI / 2;
    values.forEach((value, i) => {
      if (value === null || value <= 0 || total <= 0) return;
      const delta = (value / total) * Math.PI * 2,
        end = angle + delta;
      // Dos arcos conservan el círculo completo cuando sólo existe una categoría.
      const slice = (radius: number, start: number, finish: number, reverse = false) => {
        const mid = (start + finish) / 2;
        return `L ${cx + Math.cos(start) * radius} ${cy + Math.sin(start) * radius} A ${radius} ${radius} 0 0 ${reverse ? 0 : 1} ${cx + Math.cos(mid) * radius} ${cy + Math.sin(mid) * radius} A ${radius} ${radius} 0 0 ${reverse ? 0 : 1} ${cx + Math.cos(finish) * radius} ${cy + Math.sin(finish) * radius}`;
      };
      const d = `M ${inner ? cx + Math.cos(angle) * inner : cx} ${inner ? cy + Math.sin(angle) * inner : cy} ${slice(r, angle, end)} ${inner ? slice(inner, end, angle, true) : `L ${cx} ${cy}`} Z`;
      marks.push({
        kind: 'path',
        d,
        color: definition.colors[i] ?? CHART_COLORS[i % CHART_COLORS.length],
        tooltip: `${data.categories[i]}: ${numberLabel(value)} (${numberLabel((value / total) * 100)}%)`,
      });
      if (definition.labels && values.length <= 40) {
        const mid = angle + delta / 2,
          at = inner ? (r + inner) / 2 : r * 0.68;
        labels.push({
          x: cx + Math.cos(mid) * at,
          y: cy + Math.sin(mid) * at,
          text: `${numberLabel((value / total) * 100)}%`,
          anchor: 'middle',
        });
      }
      angle = end;
    });
    return result;
  }
  const count = data.categories.length,
    stacked = definition.stacked && ['bar', 'column', 'area'].includes(definition.type);
  const positive = Array(count).fill(0),
    negative = Array(count).fill(0);
  let min = 0,
    max = 0;
  if (stacked)
    for (let i = 0; i < count; i++) {
      const p = data.series.reduce((s, v) => s + Math.max(0, v.values[i] ?? 0), 0),
        n = data.series.reduce((s, v) => s + Math.min(0, v.values[i] ?? 0), 0);
      max = Math.max(max, p);
      min = Math.min(min, n);
    }
  else
    for (const series of data.series)
      for (const value of series.values)
        if (value !== null) {
          min = Math.min(min, value);
          max = Math.max(max, value);
        }
  const y = domain(min, max),
    numeric = (value: number) =>
      horizontal
        ? plot.left + ((value - y.min) / (y.max - y.min)) * plot.width
        : plot.top + ((y.max - value) / (y.max - y.min)) * plot.height;
  result.zero = numeric(0);
  for (let value = y.min, i = 0; value <= y.max + y.step * 0.01 && i < 12; value += y.step, i++)
    result.ticks.push({ value: numberLabel(value), position: numeric(value) });
  const scatter = definition.type === 'scatter';
  const xValues = data.x.filter((v): v is number => v !== null),
    xd = domain(Math.min(...xValues, 0), Math.max(...xValues, 1));
  const step = (horizontal ? plot.height : plot.width) / Math.max(1, count);
  const position = (i: number) =>
    scatter
      ? plot.left + (((data.x[i] ?? 0) - xd.min) / (xd.max - xd.min)) * plot.width
      : horizontal
        ? plot.top + (i + 0.5) * step
        : plot.left + (i + 0.5) * step;
  const labelStep = Math.max(1, Math.ceil(count / 8));
  if (scatter)
    for (let v = xd.min, i = 0; v <= xd.max + xd.step * 0.01 && i < 12; v += xd.step, i++)
      categories.push({
        x: plot.left + ((v - xd.min) / (xd.max - xd.min)) * plot.width,
        y: plot.top + plot.height + 20,
        text: numberLabel(v),
        anchor: 'middle',
      });
  else
    data.categories.forEach((category, i) => {
      if (i % labelStep === 0)
        categories.push({
          x: horizontal ? plot.left - 8 : position(i),
          y: horizontal ? position(i) + 4 : plot.top + plot.height + 20,
          text: category.length > 17 ? category.slice(0, 16) + '…' : category,
          anchor: horizontal ? 'end' : 'middle',
        });
    });
  data.series.forEach((series, si) => {
    let segment: { x: number; y: number; base: number }[] = [];
    const finish = () => {
      if (!segment.length) return;
      const path = segment.map((p, i) => `${i ? 'L' : 'M'} ${p.x} ${p.y}`).join(' ');
      if (definition.type === 'area') {
        const bottom = [...segment]
          .reverse()
          .map((p) => `L ${p.x} ${p.base}`)
          .join(' ');
        marks.push({
          kind: 'path',
          d: `${path} ${bottom} Z`,
          color: series.color,
          tooltip: series.name,
          opacity: 0.35,
        });
      }
      marks.push({
        kind: 'path',
        d: path,
        color: series.color,
        tooltip: series.name,
        outline: true,
      });
      segment = [];
    };
    series.values.forEach((value, i) => {
      if (value === null || (scatter && data.x[i] === null)) {
        finish();
        return;
      }
      const base = stacked ? (value >= 0 ? positive[i] : negative[i]) : 0,
        top = base + value;
      if (stacked) {
        if (value >= 0) positive[i] = top;
        else negative[i] = top;
      }
      const tooltip = `${data.categories[i]} · ${series.name}: ${numberLabel(value)}`;
      if (definition.type === 'column' || definition.type === 'bar') {
        const width = (step * 0.76) / (stacked ? 1 : data.series.length),
          offset = stacked ? -width / 2 : -step * 0.38 + si * width;
        marks.push(
          horizontal
            ? {
                kind: 'rect',
                x: Math.min(numeric(base), numeric(top)),
                y: position(i) + offset,
                width: Math.abs(numeric(top) - numeric(base)),
                height: Math.max(0.3, width - 1),
                color: series.color,
                tooltip,
              }
            : {
                kind: 'rect',
                x: position(i) + offset,
                y: Math.min(numeric(base), numeric(top)),
                width: Math.max(0.3, width - 1),
                height: Math.abs(numeric(top) - numeric(base)),
                color: series.color,
                tooltip,
              },
        );
      } else if (scatter)
        marks.push({
          kind: 'circle',
          cx: position(i),
          cy: numeric(value),
          r: 3.5,
          color: series.color,
          tooltip,
        });
      else {
        segment.push({ x: position(i), y: numeric(top), base: numeric(base) });
        if (count <= 100)
          marks.push({
            kind: 'circle',
            cx: position(i),
            cy: numeric(top),
            r: 2.5,
            color: series.color,
            tooltip,
          });
      }
      if (definition.labels && count * data.series.length <= 80)
        labels.push(
          horizontal
            ? {
                x: numeric(top) + (value < 0 ? -4 : 4),
                y: position(i) + 4,
                text: numberLabel(value),
                anchor: value < 0 ? 'end' : 'start',
              }
            : { x: position(i), y: numeric(top) - 7, text: numberLabel(value), anchor: 'middle' },
        );
    });
    finish();
  });
  return result;
}

@Component({
  selector: 'app-spreadsheet-chart-view',
  template: `<svg
    viewBox="0 0 640 400"
    role="img"
    [attr.aria-label]="definition().title || 'Gráfico'"
  >
    <title>{{ definition().title }}</title>
    <rect width="640" height="400" fill="white" />
    <text class="chart-title-text" x="320" y="27" text-anchor="middle" font-weight="600">
      {{
        definition().title.length > 65 ? definition().title.slice(0, 64) + '…' : definition().title
      }}
    </text>
    @if (!geometry().pie) {
      @for (tick of geometry().ticks; track $index) {
        @if (geometry().horizontal) {
          <line
            [attr.x1]="tick.position"
            [attr.x2]="tick.position"
            [attr.y1]="geometry().plot.top"
            [attr.y2]="geometry().plot.top + geometry().plot.height"
            stroke="#e2e8f0"
          />
          <text
            [attr.x]="tick.position"
            [attr.y]="geometry().plot.top + geometry().plot.height + 20"
            text-anchor="middle"
          >
            {{ tick.value }}
          </text>
        } @else {
          <line
            [attr.x1]="geometry().plot.left"
            [attr.x2]="geometry().plot.left + geometry().plot.width"
            [attr.y1]="tick.position"
            [attr.y2]="tick.position"
            stroke="#e2e8f0"
          />
          <text [attr.x]="geometry().plot.left - 8" [attr.y]="tick.position + 4" text-anchor="end">
            {{ tick.value }}
          </text>
        }
      }
      @for (label of geometry().categories; track $index) {
        <text [attr.x]="label.x" [attr.y]="label.y" [attr.text-anchor]="label.anchor">
          {{ label.text }}
        </text>
      }
      <text x="330" y="350" text-anchor="middle">
        {{ geometry().horizontal ? definition().yTitle : definition().xTitle }}
      </text>
      <text transform="translate(17 175) rotate(-90)" text-anchor="middle">
        {{ geometry().horizontal ? definition().xTitle : definition().yTitle }}
      </text>
    }
    @for (mark of geometry().marks; track $index) {
      @if (mark.kind === 'rect') {
        <rect
          [attr.x]="mark.x"
          [attr.y]="mark.y"
          [attr.width]="mark.width"
          [attr.height]="mark.height"
          [attr.fill]="mark.color"
        >
          <title>{{ mark.tooltip }}</title>
        </rect>
      } @else if (mark.kind === 'circle') {
        <circle [attr.cx]="mark.cx" [attr.cy]="mark.cy" [attr.r]="mark.r" [attr.fill]="mark.color">
          <title>{{ mark.tooltip }}</title>
        </circle>
      } @else {
        <path
          [attr.d]="mark.d"
          [attr.fill]="mark.outline ? 'none' : mark.color"
          [attr.stroke]="mark.color"
          [attr.stroke-width]="mark.outline ? 2 : 1"
          [attr.opacity]="mark.opacity ?? 1"
        >
          <title>{{ mark.tooltip }}</title>
        </path>
      }
    }
    @for (label of geometry().labels; track $index) {
      <text
        [attr.x]="label.x"
        [attr.y]="label.y"
        [attr.text-anchor]="label.anchor"
        font-weight="600"
      >
        {{ label.text }}
      </text>
    }
    @if (definition().legend) {
      @for (item of geometry().legend.slice(0, 12); track $index; let i = $index) {
        <rect
          [attr.x]="24 + (i % 4) * 155"
          [attr.y]="363 + Math.floor(i / 4) * 13"
          width="9"
          height="9"
          [attr.fill]="item.color"
        />
        <text [attr.x]="38 + (i % 4) * 155" [attr.y]="371 + Math.floor(i / 4) * 13">
          {{ item.name.length > 21 ? item.name.slice(0, 20) + '…' : item.name }}
        </text>
      }
    }
  </svg>`,
  styles: [
    `
      :host {
        display: block;
        width: 100%;
        height: 100%;
        min-height: 0;
      }
      svg {
        display: block;
        width: 100%;
        height: 100%;
        background: white;
      }
      .chart-title-text {
        font-size: 18px;
      }
      text {
        font-family: 'Segoe UI', sans-serif;
        font-size: 11px;
        fill: #334155;
      }
    `,
  ],
})
export class SpreadsheetChartView {
  readonly definition = input.required<ChartDefinition>();
  readonly data = input.required<ChartData>();
  protected readonly Math = Math;
  protected readonly geometry = computed(() =>
    spreadsheetChartGeometry(this.definition(), this.data()),
  );
}
