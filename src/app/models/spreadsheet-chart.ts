/** Configuración y datos de gráficos nativos XLSX. Coordenadas base cero; tamaños y desplazamientos en píxeles CSS. */
export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';
export type ChartDefinition = {
  id?: string;
  hostSheet: string;
  sourceSheet: string;
  sourceRange: string;
  title: string;
  type: ChartType;
  headers: boolean;
  category: number;
  series: number[];
  legend: boolean;
  labels: boolean;
  stacked: boolean;
  xTitle: string;
  yTitle: string;
  colors: string[];
  anchor: {
    row: number;
    col: number;
    offsetX: number;
    offsetY: number;
    width: number;
    height: number;
  };
};
export type ChartData = {
  categories: string[];
  x: (number | null)[];
  series: { name: string; values: (number | null)[]; color: string }[];
};
export type SavedChart = ChartDefinition & {
  id: string;
  chartPath: string;
  drawingPath: string;
  relationshipId: string;
  managed?: boolean;
  data: ChartData;
  error?: string;
};
