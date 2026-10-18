import { spreadsheetAddress } from './spreadsheet-formula';
import { SpreadsheetTable, spreadsheetColumnLabel } from './spreadsheet-table';
import { tableBounds } from './spreadsheet-table-resize';

export type CellMapper = (row: number, col: number) => { row: number; col: number } | null;
export const TABLE_TOTAL_FUNCTIONS = [
  '',
  'sum',
  'average',
  'count',
  'countNums',
  'min',
  'max',
  'stdDev',
  'stdDevP',
  'var',
  'varP',
];

/** Recibe nombre, tablas y nombres definidos; valida identidad única y retorna el nombre sin espacios exteriores. */
export function validateTableName(
  name: string,
  tables: SpreadsheetTable[],
  current?: string,
  definedNames: string[] = [],
): string {
  name = name.trim();
  if (
    !/^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(name) ||
    name.length > 255 ||
    /^[RC]$/i.test(name) ||
    /^R\d+C\d+$/i.test(name) ||
    spreadsheetAddress(name) ||
    tables.some((t) => t.name !== current && t.name.toLowerCase() === name.toLowerCase()) ||
    definedNames.some((n) => n.toLowerCase() === name.toLowerCase())
  )
    throw Error('NOTE_EDITOR.TABLE_NAME_ERROR');
  return name;
}

/** Recibe encabezados existentes y posición; retorna un nombre nuevo único sin renombrar columnas previas. */
export function newTableColumn(columns: string[], index: number): string {
  const base = `Column${index + 1}`;
  let name = base;
  for (let n = 2; columns.some((c) => c.toLowerCase() === name.toLowerCase()); n++)
    name = base + '_' + n;
  return name;
}

/**
 * Remapea una tabla ante inserciones, eliminaciones o movimientos completos de filas/columnas.
 * @param table Definición original (máximo 100.000 celdas).
 * @param map Transformación de coordenadas; null identifica celdas eliminadas.
 * @param allowGaps Sólo las inserciones pueden incorporar filas o columnas nuevas entre las existentes.
 * @returns Nueva definición, null si se elimina toda la tabla, o lanza una clave traducible si se parte su estructura.
 */
export function remapSpreadsheetTable(
  table: SpreadsheetTable,
  map: CellMapper,
  allowGaps = false,
): SpreadsheetTable | null {
  const b = tableBounds(table.ref);
  if (!b || (b.b.row - b.a.row + 1) * (b.b.col - b.a.col + 1) > 100000)
    throw Error('NOTE_EDITOR.TABLE_STRUCTURE_ERROR');
  const points: { row: number; col: number; oldRow: number; oldCol: number }[] = [];
  for (let r = b.a.row; r <= b.b.row; r++)
    for (let c = b.a.col; c <= b.b.col; c++) {
      const next = map(r, c);
      if (next) points.push({ ...next, oldRow: r, oldCol: c });
    }
  if (!points.length) return null;
  const rows = [...new Set(points.map((p) => p.row))].sort((a, b) => a - b),
    cols = [...new Set(points.map((p) => p.col))].sort((a, b) => a - b);
  const r1 = rows[0],
    r2 = rows.at(-1)!,
    c1 = cols[0],
    c2 = cols.at(-1)!;
  if (
    r1 < 0 ||
    c1 < 0 ||
    r2 >= 1048576 ||
    c2 >= 16384 ||
    (!allowGaps && (rows.length !== r2 - r1 + 1 || cols.length !== c2 - c1 + 1)) ||
    r2 - r1 + 1 <= table.headerRows + table.totalRows ||
    (r2 - r1 + 1) * (c2 - c1 + 1) > 100000 ||
    (table.headerRows && !points.some((p) => p.oldRow === b.a.row && p.row === r1)) ||
    (table.totalRows && !points.some((p) => p.oldRow === b.b.row && p.row === r2))
  )
    throw Error('NOTE_EDITOR.TABLE_STRUCTURE_ERROR');
  const columns: string[] = [],
    totals: string[] = [],
    columnIds: number[] = [];
  let nextId = Math.max(0, ...(table.columnIds ?? table.columns.map((_, i) => i + 1))) + 1;
  const byColumn = new Map<number, (typeof points)[number]>();
  for (const point of points) if (!byColumn.has(point.col)) byColumn.set(point.col, point);
  for (let c = c1; c <= c2; c++) {
    const old = byColumn.get(c);
    columnIds.push(
      old ? (table.columnIds?.[old.oldCol - b.a.col] ?? old.oldCol - b.a.col + 1) : nextId++,
    );
    columns.push(
      old
        ? table.columns[old.oldCol - b.a.col]
        : newTableColumn([...table.columns, ...columns], c - c1),
    );
    totals.push(old ? (table.totals?.[old.oldCol - b.a.col] ?? '') : '');
  }
  return {
    ...table,
    ref: `${spreadsheetColumnLabel(c1)}${r1 + 1}:${spreadsheetColumnLabel(c2)}${r2 + 1}`,
    columns,
    columnIds,
    ...(table.totalRows ? { totals } : {}),
  };
}

/** Recibe función de total y referencia A1 de datos; retorna SUBTOTAL compatible con Excel o null si es una etiqueta/custom. */
export function tableTotalFormula(kind: string, reference: string): string | null {
  const codes: Record<string, number> = {
    average: 101,
    countNums: 102,
    count: 103,
    max: 104,
    min: 105,
    stdDev: 107,
    stdDevP: 108,
    sum: 109,
    var: 110,
    varP: 111,
  };
  return codes[kind] ? `SUBTOTAL(${codes[kind]},${reference})` : null;
}
