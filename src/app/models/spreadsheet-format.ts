/** Formato editable de una celda; dimensiones y coordenadas usan píxeles e índices base cero. */
export type SpreadsheetFormat = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align?: 'left' | 'center' | 'right';
  fontSize?: number;
  color?: string;
  background?: string;
  wrap?: boolean;
  numberFormat?: string;
};

export type SpreadsheetEditOptions = {
  tables?: import('./spreadsheet-table').SpreadsheetTable[];
  definedNames?: Record<string, string>;
  spills?: { row: number; col: number; rows: number; cols: number }[];
  formats?: { row: number; col: number; format: SpreadsheetFormat }[];
  colWidths?: number[];
  rowHeights?: number[];
  hiddenRows?: number[];
  origins?: Record<string, string>;
};
