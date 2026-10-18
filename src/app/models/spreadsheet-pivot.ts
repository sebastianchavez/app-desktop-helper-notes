/** Funciones de agregación nativas de las tablas dinámicas XLSX. */
export type PivotAggregate = 'sum' | 'count' | 'countNums' | 'average' | 'min' | 'max';
/** Configuración de campos; sus índices son relativos a la primera columna del rango de origen. */
export type PivotDefinition = {
  id?: string;
  name: string;
  sourceSheet: string;
  sourceRange: string;
  rows: number[];
  columns: number[];
  values: { field: number; aggregate: PivotAggregate }[];
  filters: { field: number; value: string | number | boolean | null }[];
  totals: boolean;
};
/** Definición persistida con los identificadores de las partes nativas dentro del XLSX. */
export type SavedPivot = PivotDefinition & {
  id: string;
  targetSheet: string;
  tablePath: string;
  cachePath: string;
  recordsPath: string;
  cacheId: number;
  error?: string;
};
