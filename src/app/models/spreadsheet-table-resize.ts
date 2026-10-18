import { spreadsheetAddress } from './spreadsheet-formula';
import { SpreadsheetTable, spreadsheetColumnLabel } from './spreadsheet-table';

/** Recibe un rango A1; retorna extremos ordenados o null para rangos inválidos/fuera de Excel. */
export function tableBounds(ref: string) {
  const parts = ref.trim().toUpperCase().replace(/\$/g, '').split(':');
  if (parts.length > 2 || parts.some((p) => !/^[A-Z]{1,3}[1-9]\d*$/.test(p))) return null;
  const a = spreadsheetAddress(parts[0]),
    b = spreadsheetAddress(parts[1] ?? parts[0]);
  return a && b && a.row <= b.row && a.col <= b.col && b.row < 1048576 && b.col < 16384
    ? { a, b }
    : null;
}

/**
 * Valida el redimensionamiento sin borrar datos fuera del nuevo rango.
 * @param table Tabla original; el encabezado y el origen permanecen fijos.
 * @param ref Rango nuevo, incluyendo encabezados.
 * @param data Valores de la hoja para validar encabezados nuevos.
 * @param tables Otras tablas de la hoja para evitar superposiciones.
 * @returns Definición nueva o lanza una clave traducible antes de mutar datos.
 */
export function resizeSpreadsheetTable(
  table: SpreadsheetTable,
  ref: string,
  data: (string | number | boolean | null)[][],
  tables: SpreadsheetTable[],
): SpreadsheetTable {
  const old = tableBounds(table.ref),
    next = tableBounds(ref);
  if (
    !old ||
    !next ||
    next.a.row !== old.a.row ||
    next.a.col !== old.a.col ||
    next.b.row - next.a.row + 1 <= table.headerRows + table.totalRows ||
    (next.b.row - next.a.row + 1) * (next.b.col - next.a.col + 1) > 100000
  )
    throw Error('NOTE_EDITOR.TABLE_RANGE_ERROR');
  for (const other of tables) {
    if (other.name === table.name) continue;
    const bounds = tableBounds(other.ref);
    if (
      bounds &&
      next.a.row <= bounds.b.row &&
      next.b.row >= bounds.a.row &&
      next.a.col <= bounds.b.col &&
      next.b.col >= bounds.a.col
    )
      throw Error('NOTE_EDITOR.TABLE_OVERLAP');
  }
  const columns = table.columns.slice(0, next.b.col - next.a.col + 1);
  for (let col = old.b.col + 1; col <= next.b.col; col++) {
    const value = table.headerRows ? data[next.a.row]?.[col] : null;
    let name = value == null || value === '' ? `Column${col - next.a.col + 1}` : String(value);
    if (value == null || value === '') {
      const base = name;
      for (let i = 2; columns.some((c) => c.toLowerCase() === name.toLowerCase()); i++)
        name = base + '_' + i;
    }
    if (name.length > 255 || columns.some((c) => c.toLowerCase() === name.toLowerCase()))
      throw Error('NOTE_EDITOR.TABLE_HEADER_ERROR');
    columns.push(name);
  }
  return {
    ...table,
    ref: `${spreadsheetColumnLabel(next.a.col)}${next.a.row + 1}:${spreadsheetColumnLabel(next.b.col)}${next.b.row + 1}`,
    columns,
  };
}
