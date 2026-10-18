import type { SheetValue } from './spreadsheet-formula';
export type SpreadsheetValue = SheetValue | SheetValue[][];
const dimensions = new WeakMap<SheetValue[][], { rows: number; cols: number }>();
export const MAX_ARRAY_CELLS = 100000;

/** Recibe valor y dimensiones lógicas opcionales; retorna su matriz, conservando las colas vacías de columnas completas sin reservarlas. */
export function spreadsheetMatrix(value: SpreadsheetValue): SheetValue[][] {
  return Array.isArray(value) ? value : [[value]];
}
/** Recibe matriz y dimensiones; retorna la misma matriz con tamaño lógico de hoja completa. */
export function markSpreadsheetRange(
  matrix: SheetValue[][],
  rows: number,
  cols: number,
): SheetValue[][] {
  dimensions.set(matrix, { rows, cols });
  return matrix;
}
/** Recibe valor; retorna sus dimensiones lógicas, aunque la cola vacía no esté materializada. */
export function spreadsheetShape(value: SpreadsheetValue): { rows: number; cols: number } {
  return Array.isArray(value)
    ? (dimensions.get(value) ?? { rows: value.length, cols: value[0]?.length ?? 0 })
    : { rows: 1, cols: 1 };
}
/** Recibe valor e índices base cero; retorna la celda o null si pertenece a una cola vacía. */
export function spreadsheetAt(value: SpreadsheetValue, row: number, col: number): SheetValue {
  return Array.isArray(value) ? (value[row]?.[col] ?? null) : row === 0 && col === 0 ? value : null;
}
/** Recibe valor; retorna cantidad lógica de celdas vacías omitidas de la representación. */
export function spreadsheetTail(value: SpreadsheetValue): number {
  const s = spreadsheetShape(value);
  return s.rows * s.cols - (Array.isArray(value) ? value.length * (value[0]?.length ?? 0) : 1);
}
/** Recibe valor; retorna matriz completa hasta el límite seguro o rechaza un resultado excesivo. */
export function materializeSpreadsheetArray(value: SpreadsheetValue): SheetValue[][] {
  const s = spreadsheetShape(value);
  if (s.rows * s.cols > MAX_ARRAY_CELLS || !s.rows || !s.cols) throw Error('#NUM!');
  return Array.from({ length: s.rows }, (_, r) =>
    Array.from({ length: s.cols }, (_, c) => spreadsheetAt(value, r, c)),
  );
}
/** Recibe uno o dos valores y operación escalar; retorna escalar o matriz con difusión por filas/columnas y errores por celda. */
export function mapSpreadsheetArray(
  left: SpreadsheetValue,
  right: SpreadsheetValue,
  operation: (a: SheetValue, b: SheetValue) => SheetValue,
): SpreadsheetValue {
  if (!Array.isArray(left) && !Array.isArray(right)) return operation(left, right);
  const a = spreadsheetShape(left),
    b = spreadsheetShape(right),
    rows = Math.max(a.rows, b.rows),
    cols = Math.max(a.cols, b.cols);
  if (
    (a.rows !== b.rows && a.rows !== 1 && b.rows !== 1) ||
    (a.cols !== b.cols && a.cols !== 1 && b.cols !== 1)
  )
    throw Error('#VALUE!');
  if (rows * cols > MAX_ARRAY_CELLS) throw Error('#NUM!');
  return Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, c) => {
      try {
        return operation(
          spreadsheetAt(left, a.rows === 1 ? 0 : r, a.cols === 1 ? 0 : c),
          spreadsheetAt(right, b.rows === 1 ? 0 : r, b.cols === 1 ? 0 : c),
        );
      } catch (error) {
        return error instanceof Error ? error.message : '#VALUE!';
      }
    }),
  );
}

/** Recibe función, argumentos evaluados y comparación; retorna matriz calculada o undefined si corresponde a otra biblioteca. */
export function evaluateSpreadsheetArray(
  name: string,
  values: SpreadsheetValue[],
  compare: (a: SheetValue, b: SheetValue) => number,
): SpreadsheetValue | undefined {
  const scalar = (i: number, fallback: SheetValue = null) =>
    values[i] === undefined ? fallback : spreadsheetAt(values[i], 0, 0);
  const numeric = (i: number, fallback = 0) => {
    const value = Number(scalar(i, fallback));
    if (!Number.isFinite(value)) throw Error('#VALUE!');
    return value;
  };
  const transpose = (array: SheetValue[][]) =>
    array[0].map((_, c) => array.map((row) => row[c] ?? null));
  switch (name) {
    case 'SEQUENCE': {
      const rows = Math.trunc(numeric(0)),
        cols = Math.trunc(numeric(1, 1)),
        start = numeric(2, 1),
        step = numeric(3, 1);
      if (rows < 1 || cols < 1 || rows > 1048576 || cols > 16384 || rows * cols > MAX_ARRAY_CELLS)
        throw Error('#NUM!');
      return Array.from({ length: rows }, (_, r) =>
        Array.from({ length: cols }, (_, c) => start + (r * cols + c) * step),
      );
    }
    case 'TRANSPOSE':
      return transpose(materializeSpreadsheetArray(values[0]));
    case 'SORT':
    case 'UNIQUE': {
      let array = materializeSpreadsheetArray(values[0]);
      const byColumn = Boolean(scalar(name === 'SORT' ? 3 : 1, false));
      if (byColumn) array = transpose(array);
      if (name === 'SORT') {
        const index = Math.trunc(numeric(1, 1)) - 1,
          order = numeric(2, 1);
        if (index < 0 || index >= array[0].length || ![-1, 1].includes(order))
          throw Error('#VALUE!');
        array = array
          .map((row, index) => ({ row, index }))
          .sort((a, b) => compare(a.row[index], b.row[index]) * order || a.index - b.index)
          .map((v) => v.row);
      } else {
        const key = (row: SheetValue[]) =>
          JSON.stringify(
            row.map((v) => (typeof v === 'string' ? v.toLocaleLowerCase() : (v ?? 0))),
          );
        const counts = new Map<string, number>();
        for (const row of array) counts.set(key(row), (counts.get(key(row)) ?? 0) + 1);
        const seen = new Set<string>();
        array = array.filter((row) => {
          const k = key(row);
          if (seen.has(k) || (Boolean(scalar(2, false)) && counts.get(k) !== 1)) return false;
          seen.add(k);
          return true;
        });
        if (!array.length) throw Error('#CALC!');
      }
      return byColumn ? transpose(array) : array;
    }
    case 'FILTER': {
      const array = materializeSpreadsheetArray(values[0]),
        mask = materializeSpreadsheetArray(values[1]);
      const truth = (v: SheetValue) => {
        if (typeof v === 'string') throw Error(v.startsWith('#') ? v : '#VALUE!');
        return Boolean(v);
      };
      let result: SheetValue[][];
      if (mask[0].length === 1 && mask.length === array.length)
        result = array.filter((_, r) => truth(mask[r][0]));
      else if (mask.length === 1 && mask[0].length === array[0].length) {
        const selected = mask[0].map(truth);
        result = array.map((row) => row.filter((_, c) => selected[c]));
        if (!result[0].length) result = [];
      } else throw Error('#VALUE!');
      if (!result.length) {
        if (values[2] !== undefined) return values[2];
        throw Error('#CALC!');
      }
      return result;
    }
    default:
      return undefined;
  }
}
