import { spreadsheetAddress } from './spreadsheet-formula';
/** Definición nativa importada; ref incluye encabezados y totales si existen. */
export type SpreadsheetTable = {
  name: string;
  id?: number;
  columnIds?: number[];
  sheet: string;
  ref: string;
  columns: string[];
  headerRows: number;
  totalRows: number;
  sourceName?: string;
  totals?: string[];
  totalFormulas?: Record<string, string>;
  calculated?: Record<string, string>;
  columnSources?: Record<string, string>;
};

/** Recibe índice de columna base cero; retorna su etiqueta A1 sin modificar referencias existentes. */
export function spreadsheetColumnLabel(col: number): string {
  let label = '';
  for (col++; col; col = Math.floor((col - 1) / 26))
    label = String.fromCharCode(65 + ((col - 1) % 26)) + label;
  return label;
}

/**
 * Resuelve referencias estructuradas de tablas existentes, sin transformar textos de la fórmula.
 * @param reference Table[Column], Table[[#Headers],[Column]], [@Column] o selección contigua de columnas.
 * @param tables Tablas importadas del libro.
 * @param ownSheet Hoja propietaria de la fórmula.
 * @param row Fila base cero de la fórmula, usada por @/#This Row.
 * @returns Referencia A1 calificada, undefined si no hay tabla y #REF! si la selección no es válida.
 */
export function resolveSpreadsheetTable(
  reference: string,
  tables: SpreadsheetTable[],
  ownSheet: string,
  row: number,
  col?: number,
): string | undefined {
  const start = reference.indexOf('[');
  if (start < 0) return undefined;
  const name = reference.slice(0, start),
    body = reference.slice(start);
  const table = name
    ? tables.find((t) => t.name.toLowerCase() === name.toLowerCase())
    : tables.find((t) => {
        const [a, b] = t.ref.split(':').map(spreadsheetAddress);
        return (
          t.sheet.toLowerCase() === ownSheet.toLowerCase() &&
          a &&
          b &&
          row >= a.row &&
          row <= b.row &&
          (col === undefined || (col >= a.col && col <= b.col))
        );
      });
  if (!table) return undefined;
  const [first, last] = table.ref.split(':').map(spreadsheetAddress);
  if (!first || !last) return '#REF!';
  const members: string[] = [];
  let depth = 0,
    text = '';
  for (let i = 0; i < body.length; i++) {
    const c = body[i];
    if (c === "'" && /[\[\]#@']/.test(body[i + 1] ?? '')) {
      text += body[++i];
      continue;
    }
    if (c === '[') {
      depth++;
      if (depth > 1) text = '';
    } else if (c === ']') {
      if (text && text !== ',' && text !== ':') members.push(text.replace(/^,\s*/, ''));
      text = '';
      depth--;
    } else if (depth) text += c;
  }
  const selectors = members.filter((v) => v.startsWith('#')).map((v) => v.toLowerCase());
  const thisRow =
    body.startsWith('[@') || selectors.some((v) => ['#this row', '#esta fila'].includes(v));
  const columns = members.filter((v) => !v.startsWith('#')).map((v) => v.replace(/^@/, ''));
  const indices = columns.length
    ? columns.map((c) => table.columns.findIndex((n) => n.toLowerCase() === c.toLowerCase()))
    : [0, table.columns.length - 1];
  if (indices.some((i) => i < 0) || indices.length > 2) return '#REF!';
  // Las uniones válidas aún no implementadas se reconocen como no soportadas para conservar su caché importada.
  if ((columns.length === 2 && !/\]\s*:\s*\[/.test(body)) || selectors.length > 1) return '#NAME?';
  if (
    selectors.some(
      (v) =>
        ![
          '#all',
          '#todo',
          '#data',
          '#datos',
          '#headers',
          '#encabezados',
          '#totals',
          '#totales',
          '#this row',
          '#esta fila',
        ].includes(v),
    )
  )
    return '#REF!';
  let r1 = first.row + table.headerRows,
    r2 = last.row - table.totalRows;
  if (selectors.some((v) => ['#all', '#todo'].includes(v))) {
    r1 = first.row;
    r2 = last.row;
  } else if (selectors.some((v) => ['#headers', '#encabezados'].includes(v))) {
    if (!table.headerRows) return '#REF!';
    r1 = first.row;
    r2 = first.row + table.headerRows - 1;
  } else if (selectors.some((v) => ['#totals', '#totales'].includes(v))) {
    if (!table.totalRows) return '#REF!';
    r1 = last.row - table.totalRows + 1;
    r2 = last.row;
  }
  if (thisRow) {
    if (row < first.row + table.headerRows || row > last.row - table.totalRows) return '#VALUE!';
    r1 = row;
    r2 = row;
  }
  if (r1 > r2) return '#REF!';
  const prefix = `'${table.sheet.replace(/'/g, "''")}'!`,
    c1 = first.col + Math.min(...indices),
    c2 = first.col + Math.max(...indices);
  if (thisRow && c1 === c2) return `${prefix}$${spreadsheetColumnLabel(c1)}$${r1 + 1}`;
  return `${prefix}$${spreadsheetColumnLabel(c1)}$${r1 + 1}:$${spreadsheetColumnLabel(c2)}$${r2 + 1}`;
}
