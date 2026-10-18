import * as XLSX from 'xlsx';
import { spreadsheetErrorText } from './spreadsheet-values';
import { DOMParser, XMLSerializer, Element as XmlElement } from '@xmldom/xmldom';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import path from 'node:path';
import { randomUUID } from 'node:crypto';

type Cell = string | number | boolean | null;
export type PivotAggregate = 'sum' | 'count' | 'countNums' | 'average' | 'min' | 'max';
export type PivotDefinition = {
  id?: string;
  name: string;
  sourceSheet: string;
  sourceRange: string;
  rows: number[];
  columns: number[];
  values: { field: number; aggregate: PivotAggregate }[];
  filters: { field: number; value: Cell }[];
  totals: boolean;
};
export type SavedPivot = PivotDefinition & {
  id: string;
  targetSheet: string;
  tablePath: string;
  cachePath: string;
  recordsPath: string;
  cacheId: number;
  error?: string;
};
type PivotResult = {
  headers: string[];
  records: Cell[][];
  shared: Cell[][];
  rows: Cell[][];
  columns: Cell[][];
  matrix: Cell[][];
  matched: number;
  headerRows: number;
  definition: PivotDefinition;
};
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const METADATA = 'helperNotes/pivots.json';
const AGGREGATES = new Set<PivotAggregate>(['sum', 'count', 'countNums', 'average', 'min', 'max']);
const CAPTIONS: Record<PivotAggregate, string> = {
  sum: 'Suma',
  count: 'Conteo',
  countNums: 'Conteo numérico',
  average: 'Promedio',
  min: 'Mínimo',
  max: 'Máximo',
};

/** Recibe texto sin confianza; retorna contenido XML escapado y sin caracteres de control inválidos. */
function xml(value: unknown): string {
  return String(value ?? '')
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
/** Recibe nodo y nombre opcional; retorna sólo hijos XML directos. */
function children(element: XmlElement | undefined, name?: string): XmlElement[] {
  return Array.from(element?.childNodes ?? [])
    .filter((n) => n.nodeType === 1)
    .map((n) => n as XmlElement)
    .filter((n) => !name || n.localName === name);
}

/**
 * Sustituye el área calculada conservando celdas y ajustes ajenos que otro editor haya añadido a la hoja.
 * @param archive Paquete actual.
 * @param targetPath Hoja propietaria del resultado.
 * @param fresh Hoja generada con la matriz nueva desde A1.
 * @param previous Tabla existente, si se está actualizando.
 * @returns Documento de hoja con los datos nuevos y los elementos externos intactos.
 */
function mergePivotOutput(
  archive: Record<string, Uint8Array>,
  targetPath: string,
  fresh: ReturnType<typeof parse>,
  previous?: SavedPivot,
) {
  if (!previous || !archive[targetPath]) return fresh;
  const original = parse(archive, targetPath),
    root = original.documentElement!,
    data = children(root, 'sheetData')[0];
  const location =
    children(parse(archive, previous.tablePath).documentElement!, 'location')[0]?.getAttribute(
      'ref',
    ) ?? 'A1';
  const oldRange = XLSX.utils.decode_range(location),
    newRange = XLSX.utils.decode_range(
      children(fresh.documentElement!, 'dimension')[0]?.getAttribute('ref') ?? 'A1',
    );
  oldRange.s = { r: 0, c: 0 };
  const inRange = (r: number, c: number, range: XLSX.Range) =>
    r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c;
  const styles = new Map<string, string>();
  for (const row of children(data, 'row'))
    for (const cell of children(row, 'c')) {
      const ref = cell.getAttribute('r')!,
        address = XLSX.utils.decode_cell(ref);
      if (cell.hasAttribute('s')) styles.set(ref, cell.getAttribute('s')!);
      if (inRange(address.r, address.c, oldRange) || inRange(address.r, address.c, newRange))
        row.removeChild(cell);
    }
  for (const row of children(children(fresh.documentElement!, 'sheetData')[0], 'row')) {
    const number = Number(row.getAttribute('r'));
    let target = children(data, 'row').find((r) => Number(r.getAttribute('r')) === number);
    if (!target) {
      target = original.createElementNS(NS, 'row');
      target.setAttribute('r', String(number));
      data.insertBefore(
        target,
        children(data, 'row').find((r) => Number(r.getAttribute('r')) > number) ?? null,
      );
    }
    for (const cell of children(row, 'c')) {
      const copy = original.importNode(cell, true) as XmlElement,
        ref = copy.getAttribute('r')!,
        col = XLSX.utils.decode_cell(ref).c;
      if (styles.has(ref)) copy.setAttribute('s', styles.get(ref)!);
      target.insertBefore(
        copy,
        children(target, 'c').find((c) => XLSX.utils.decode_cell(c.getAttribute('r')!).c > col) ??
          null,
      );
    }
  }
  const dimension = children(root, 'dimension')[0];
  if (dimension) {
    const addresses = Array.from(data.getElementsByTagNameNS(NS, 'c')).map((c) =>
      XLSX.utils.decode_cell(c.getAttribute('r')!),
    );
    dimension.setAttribute(
      'ref',
      XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: {
          r: Math.max(newRange.e.r, ...addresses.map((a) => a.r)),
          c: Math.max(newRange.e.c, ...addresses.map((a) => a.c)),
        },
      }),
    );
  }
  return original;
}
/** Recibe Base64 XLSX/XLS/CSV; retorna el ZIP XLSX original o su conversión inicial. */
function openArchive(base64: string): Record<string, Uint8Array> {
  let bytes = Buffer.from(base64, 'base64');
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b)
    bytes = Buffer.from(
      XLSX.write(XLSX.read(bytes, { type: 'buffer' }), { type: 'buffer', bookType: 'xlsx' }),
    );
  return unzipSync(bytes);
}
/** Recibe el ZIP y una parte XML; retorna el documento o una raíz nueva para relaciones opcionales. */
function parse(archive: Record<string, Uint8Array>, name: string, fallback?: string) {
  return new DOMParser().parseFromString(
    archive[name] ? strFromU8(archive[name]) : (fallback ?? ''),
    'application/xml',
  );
}
/** Recibe documento XML y archivo destino; escribe sólo esa parte del ZIP en memoria y no retorna valor. */
function store(
  archive: Record<string, Uint8Array>,
  name: string,
  document: ReturnType<typeof parse>,
): void {
  archive[name] = strToU8(new XMLSerializer().serializeToString(document));
}
/** Recibe un destino relativo y parte propietaria; retorna la ruta normalizada dentro del XLSX. */
function resolve(owner: string, target: string): string {
  return target.startsWith('/')
    ? target.slice(1)
    : path.posix.normalize(path.posix.join(path.posix.dirname(owner), target));
}
/** Recibe una parte; retorna la ruta de sus relaciones OPC. */
function relPath(part: string): string {
  return path.posix.join(path.posix.dirname(part), '_rels', path.posix.basename(part) + '.rels');
}

/**
 * Valida el origen y calcula el cruce sin sumar promedios parciales ni convertir texto en números.
 * @param base64 Libro completo cuyos valores calculados ya están guardados.
 * @param input Distribución de campos relativa a la primera columna del rango de origen.
 * @returns Matriz de resultados, registros filtrados y miembros de caché tipados; rechaza configuraciones inválidas.
 */
export function calculatePivot(base64: string, input: PivotDefinition): PivotResult {
  if (!input || typeof input.name !== 'string' || !input.name.trim() || input.name.length > 80)
    throw Error('PIVOT_NAME');
  if (
    typeof input.sourceRange !== 'string' ||
    !/^\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$/i.test(input.sourceRange)
  )
    throw Error('PIVOT_RANGE');
  const range = XLSX.utils.decode_range(input.sourceRange.replace(/\$/g, '').toUpperCase());
  if (
    range.s.r < 0 ||
    range.s.c < 0 ||
    range.e.r <= range.s.r ||
    range.e.c < range.s.c ||
    range.e.r >= 1048576 ||
    range.e.c >= 16384 ||
    (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1) > 1000000
  )
    throw Error('PIVOT_RANGE');
  const workbook = XLSX.read(Buffer.from(base64, 'base64'), { type: 'buffer', cellDates: false });
  const sheet = workbook.Sheets[input.sourceSheet];
  if (!sheet) throw Error('PIVOT_SOURCE');
  const headers = Array.from({ length: range.e.c - range.s.c + 1 }, (_, i) =>
    String(sheet[XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c + i })]?.v ?? '').trim(),
  );
  if (
    headers.some((h) => !h) ||
    new Set(headers.map((h) => h.toLowerCase())).size !== headers.length
  )
    throw Error('PIVOT_HEADERS');
  const field = (i: number) => Number.isInteger(i) && i >= 0 && i < headers.length;
  if (
    !Array.isArray(input.rows) ||
    !input.rows.length ||
    input.rows.length > 5 ||
    !Array.isArray(input.columns) ||
    input.columns.length > 3 ||
    !Array.isArray(input.values) ||
    !input.values.length ||
    input.values.length > 8 ||
    !Array.isArray(input.filters) ||
    input.filters.length > headers.length ||
    [
      ...input.rows,
      ...input.columns,
      ...input.filters.map((f) => f.field),
      ...input.values.map((v) => v.field),
    ].some((i) => !field(i)) ||
    new Set([...input.rows, ...input.columns, ...input.filters.map((f) => f.field)]).size !==
      input.rows.length + input.columns.length + input.filters.length ||
    input.values.some((v) => !AGGREGATES.has(v.aggregate)) ||
    new Set(input.values.map((v) => `${v.field}:${v.aggregate}`)).size !== input.values.length
  )
    throw Error('PIVOT_FIELDS');
  const definition = {
    ...input,
    name: input.name.trim(),
    sourceRange: XLSX.utils.encode_range(range),
  };
  const records: Cell[][] = [];
  for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const record = headers.map((_, i) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: range.s.c + i })];
      if (cell?.t === 'e') {
        // Los códigos BIFF de error son números: nunca sumarlos como si fueran cantidades del origen.
        if (input.values.some((v) => v.field === i && !v.aggregate.startsWith('count')))
          throw Error('PIVOT_ERRORS');
        return spreadsheetErrorText(cell.v);
      }
      const value = cell?.v;
      return (typeof value === 'number' && Number.isFinite(value)) ||
        typeof value === 'boolean' ||
        typeof value === 'string'
        ? value
        : null;
    });
    if (record.some((v) => v !== null && v !== '')) records.push(record);
  }
  if (records.length > 100000) throw Error('PIVOT_LIMIT');
  const shared = headers.map((_, i) =>
    Array.from(new Map(records.map((r) => [JSON.stringify(r[i]), r[i]])).values()),
  );
  const filtered = records.filter((r) =>
    input.filters.every((f) => JSON.stringify(r[f.field]) === JSON.stringify(f.value)),
  );
  const tuples = (fields: number[]) => {
    const unique = Array.from(
      new Map(
        filtered.map((r) => {
          const tuple = fields.map((f) => r[f]);
          return [JSON.stringify(tuple), tuple];
        }),
      ).values(),
    );
    return unique.sort((a, b) => {
      for (let i = 0; i < a.length; i++) {
        const x = a[i],
          y = b[i];
        const comparison =
          typeof x === 'number' && typeof y === 'number'
            ? x - y
            : String(x ?? '').localeCompare(String(y ?? ''), undefined, { numeric: true });
        if (comparison) return comparison;
      }
      return 0;
    });
  };
  const rows = tuples(input.rows),
    columns = input.columns.length ? tuples(input.columns) : [[]];
  if (
    (rows.length + 2) * (Math.max(1, columns.length) * input.values.length + input.rows.length) >
    100000
  )
    throw Error('PIVOT_LIMIT');
  type Stats = { sum: number; numeric: number; count: number; min: number; max: number };
  const stats = new Map<string, Stats[]>();
  const tupleKey = (values: Cell[] | null) => (values === null ? '*' : JSON.stringify(values));
  const statKey = (r: Cell[] | null, c: Cell[] | null) =>
    JSON.stringify([tupleKey(r), tupleKey(c)]);
  for (const record of filtered) {
    const r = input.rows.map((i) => record[i]),
      c = input.columns.map((i) => record[i]);
    for (const [row, col] of [
      [r, c],
      [r, null],
      [null, c],
      [null, null],
    ] as [Cell[] | null, Cell[] | null][]) {
      const key = statKey(row, col);
      let group = stats.get(key);
      if (!group) {
        group = input.values.map(() => ({
          sum: 0,
          numeric: 0,
          count: 0,
          min: Infinity,
          max: -Infinity,
        }));
        stats.set(key, group);
      }
      input.values.forEach((measure, i) => {
        const v = record[measure.field],
          s = group![i];
        if (v !== null && v !== '') s.count++;
        if (typeof v === 'number') {
          s.numeric++;
          s.sum += v;
          s.min = Math.min(s.min, v);
          s.max = Math.max(s.max, v);
        }
      });
    }
  }
  const value = (row: Cell[] | null, col: Cell[] | null, i: number): Cell => {
    const s = stats.get(statKey(row, col))?.[i];
    if (!s)
      return input.values[i].aggregate === 'sum' || input.values[i].aggregate.startsWith('count')
        ? 0
        : null;
    switch (input.values[i].aggregate) {
      case 'sum':
        return s.sum;
      case 'count':
        return s.count;
      case 'countNums':
        return s.numeric;
      case 'average':
        return s.numeric ? s.sum / s.numeric : null;
      case 'min':
        return s.numeric ? s.min : null;
      case 'max':
        return s.numeric ? s.max : null;
    }
  };
  const headerRows = Math.max(1, input.columns.length + (input.values.length > 1 ? 1 : 0));
  const totalCols = input.totals && input.columns.length > 0;
  const matrix: Cell[][] = Array.from({ length: headerRows }, () => input.rows.map(() => null));
  input.rows.forEach((field, i) => (matrix[headerRows - 1][i] = headers[field]));
  for (const column of [...columns, ...(totalCols ? [null] : [])]) {
    input.values.forEach((measure) => {
      for (let level = 0; level < headerRows; level++) {
        const caption = `${CAPTIONS[measure.aggregate]} · ${headers[measure.field]}`;
        matrix[level].push(
          column === null
            ? level === 0
              ? 'Total general'
              : caption
            : input.columns.length === 0
              ? caption
              : level < input.columns.length
                ? (column[level] ?? '(Vacío)')
                : caption,
        );
      }
    });
  }
  for (const row of [...rows, ...(input.totals ? [null] : [])]) {
    const output: Cell[] =
      row === null
        ? input.rows.map((_, i) => (i === 0 ? 'Total general' : null))
        : row.map((v) => v ?? '(Vacío)');
    for (const column of [...columns, ...(totalCols ? [null] : [])])
      input.values.forEach((_, i) => output.push(value(row, column, i)));
    matrix.push(output);
  }
  return {
    headers,
    records,
    shared,
    rows,
    columns,
    matrix,
    matched: filtered.length,
    headerRows,
    definition,
  };
}

/** Recibe un valor de caché; retorna el elemento tipado SpreadsheetML correspondiente. */
function cacheValue(value: Cell): string {
  return value === null
    ? '<m/>'
    : typeof value === 'number'
      ? `<n v="${value}"/>`
      : typeof value === 'boolean'
        ? `<b v="${value ? 1 : 0}"/>`
        : `<s v="${xml(value)}"/>`;
}

/**
 * Lee las tablas gestionadas por Helper Notes sin modificar cachés o partes nativas ajenas.
 * @param base64 Libro XLSX completo.
 * @returns Definiciones persistidas, o lista vacía si el archivo aún no tiene tablas creadas en Helper Notes.
 */
export function readSpreadsheetPivots(base64: string): SavedPivot[] {
  const archive = openArchive(base64);
  if (!archive[METADATA]) return recoverNativePivots(archive);
  const values: unknown = JSON.parse(strFromU8(archive[METADATA]));
  if (!Array.isArray(values) || values.length > 100) throw Error('PIVOT_METADATA');
  return values.filter(
    (p) =>
      p &&
      typeof p.id === 'string' &&
      /^xl\/pivotTables\/[^/]+\.xml$/.test(p.tablePath) &&
      /^xl\/pivotCache\/[^/]+\.xml$/.test(p.cachePath) &&
      /^xl\/pivotCache\/[^/]+\.xml$/.test(p.recordsPath) &&
      archive[p.tablePath] &&
      archive[p.cachePath],
  ) as SavedPivot[];
}

/**
 * Recupera tablas creadas en Helper Notes si otro editor quitó los metadatos privados al guardar.
 * @param archive ZIP del XLSX, leído sin modificarlo.
 * @returns Sólo tablas propias con origen de rango y agregaciones compatibles; las tablas externas se conservan intactas.
 */
function recoverNativePivots(archive: Record<string, Uint8Array>): SavedPivot[] {
  const workbook = parse(archive, 'xl/workbook.xml'),
    relations = parse(archive, 'xl/_rels/workbook.xml.rels');
  const pivots: SavedPivot[] = [];
  for (const sheet of Array.from(workbook.getElementsByTagNameNS(NS, 'sheet'))) {
    const rel = children(relations.documentElement!).find(
      (r) => r.getAttribute('Id') === sheet.getAttribute('r:id'),
    );
    if (!rel) continue;
    const sheetPath = resolve('xl/workbook.xml', rel.getAttribute('Target')!);
    if (!archive[relPath(sheetPath)]) continue;
    const sheetRels = parse(archive, relPath(sheetPath));
    for (const tableRel of children(sheetRels.documentElement!).filter(
      (r) => r.getAttribute('Type') === `${REL}/pivotTable`,
    )) {
      const tablePath = resolve(sheetPath, tableRel.getAttribute('Target')!);
      if (!archive[tablePath]) continue;
      const table = parse(archive, tablePath).documentElement!,
        name = table.getAttribute('name') ?? '';
      if (!name.startsWith('HelperPivot_')) continue;
      if (!archive[relPath(tablePath)]) continue;
      const cacheRel = children(parse(archive, relPath(tablePath)).documentElement!).find(
        (r) => r.getAttribute('Type') === `${REL}/pivotCacheDefinition`,
      );
      if (!cacheRel) continue;
      const cachePath = resolve(tablePath, cacheRel.getAttribute('Target')!);
      if (!archive[cachePath]) continue;
      const cache = parse(archive, cachePath).documentElement!,
        source = cache.getElementsByTagNameNS(NS, 'worksheetSource')[0];
      if (!source?.getAttribute('sheet') || !source.getAttribute('ref')) continue;
      const recordsRel = archive[relPath(cachePath)]
        ? children(parse(archive, relPath(cachePath)).documentElement!).find(
            (r) => r.getAttribute('Type') === `${REL}/pivotCacheRecords`,
          )
        : undefined;
      if (!recordsRel) continue;
      const rows = children(children(table, 'rowFields')[0] ?? table, 'field').map((f) =>
        Number(f.getAttribute('x')),
      );
      const columns = children(children(table, 'colFields')[0] ?? table, 'field')
        .map((f) => Number(f.getAttribute('x')))
        .filter((i) => i >= 0);
      const values = children(children(table, 'dataFields')[0] ?? table, 'dataField').map((f) => ({
        field: Number(f.getAttribute('fld')),
        aggregate: (f.getAttribute('subtotal') || 'sum') as PivotAggregate,
      }));
      if (!rows.length || values.some((v) => !AGGREGATES.has(v.aggregate))) continue;
      const fields = children(children(cache, 'cacheFields')[0] ?? cache, 'cacheField');
      const pivotFields = children(children(table, 'pivotFields')[0] ?? table, 'pivotField');
      const filters: PivotDefinition['filters'] = [];
      for (const page of children(children(table, 'pageFields')[0] ?? table, 'pageField')) {
        const field = Number(page.getAttribute('fld'));
        if (!page.hasAttribute('item')) continue;
        const item = children(
          children(pivotFields[field], 'items')[0] ?? pivotFields[field],
          'item',
        )[Number(page.getAttribute('item'))];
        const shared = children(children(fields[field], 'sharedItems')[0] ?? fields[field])[
          Number(item?.getAttribute('x'))
        ];
        if (!shared) continue;
        const text = shared.getAttribute('v');
        filters.push({
          field,
          value:
            shared.localName === 'n'
              ? Number(text)
              : shared.localName === 'b'
                ? text === '1'
                : shared.localName === 'm'
                  ? null
                  : (text ?? ''),
        });
      }
      const origin = XLSX.utils.decode_range(
        children(table, 'location')[0]?.getAttribute('ref') ?? 'A1',
      ).s;
      if (origin.c !== 0 || origin.r !== (filters.length ? filters.length + 1 : 0)) continue;
      pivots.push({
        id: name.slice('HelperPivot_'.length),
        name: sheet.getAttribute('name')!,
        targetSheet: sheet.getAttribute('name')!,
        sourceSheet: source.getAttribute('sheet')!,
        sourceRange: source.getAttribute('ref')!,
        rows,
        columns,
        values,
        filters,
        totals:
          table.getAttribute('rowGrandTotals') !== '0' ||
          table.getAttribute('colGrandTotals') !== '0',
        tablePath,
        cachePath,
        recordsPath: resolve(cachePath, recordsRel.getAttribute('Target')!),
        cacheId: Number(table.getAttribute('cacheId')),
      });
    }
  }
  return pivots;
}

/**
 * Crea/actualiza una tabla dinámica nativa y su hoja de salida, conservando el resto del paquete original.
 * @param base64 Libro con datos actuales.
 * @param input Definición; id existente reutiliza su hoja y sus partes, sin duplicarlas.
 * @returns Archivo Base64 completo y definición guardada.
 */
export function saveSpreadsheetPivot(
  base64: string,
  input: PivotDefinition,
): { base64: string; pivot: SavedPivot } {
  const result = calculatePivot(base64, input),
    archive = openArchive(base64),
    pivots = readSpreadsheetPivots(base64);
  const previous = input.id ? pivots.find((p) => p.id === input.id) : undefined;
  if (input.id && !previous) throw Error('PIVOT_NOT_FOUND');
  if (pivots.length >= 100 && !previous) throw Error('PIVOT_LIMIT');
  const workbook = parse(archive, 'xl/workbook.xml'),
    relationships = parse(archive, 'xl/_rels/workbook.xml.rels');
  const types = parse(archive, '[Content_Types].xml');
  const sheetNodes = Array.from(workbook.getElementsByTagNameNS(NS, 'sheet'));
  const savedName =
    (previous?.targetSheet ?? input.name)
      .replace(/[\\/?*\[\]:]/g, ' ')
      .slice(0, 31)
      .trim() || 'Tabla dinámica';
  if (
    !previous &&
    sheetNodes.some((s) => s.getAttribute('name')?.toLowerCase() === savedName.toLowerCase())
  )
    throw Error('PIVOT_DUPLICATE');
  if (
    previous?.targetSheet === input.sourceSheet ||
    pivots.some((p) => p.targetSheet === input.sourceSheet)
  )
    throw Error('PIVOT_SOURCE');
  const id = previous?.id ?? randomUUID();
  const occupied = new Set(Object.keys(archive));
  let part = 1;
  while (
    occupied.has(`xl/pivotTables/pivotTable${part}.xml`) ||
    occupied.has(`xl/pivotCache/pivotCacheDefinition${part}.xml`)
  )
    part++;
  const cacheIds = Array.from(workbook.getElementsByTagNameNS(NS, 'pivotCache')).map((c) =>
    Number(c.getAttribute('cacheId')),
  );
  const pivot: SavedPivot = {
    ...result.definition,
    id,
    targetSheet: previous?.targetSheet ?? savedName,
    tablePath: previous?.tablePath ?? `xl/pivotTables/pivotTable${part}.xml`,
    cachePath: previous?.cachePath ?? `xl/pivotCache/pivotCacheDefinition${part}.xml`,
    recordsPath: previous?.recordsPath ?? `xl/pivotCache/pivotCacheRecords${part}.xml`,
    cacheId: previous?.cacheId ?? Math.max(-1, ...cacheIds) + 1,
  };
  delete pivot.error;
  /** Recibe tipo y destino de relación; retorna un ID nuevo y agrega la relación al libro en memoria. */
  const relationship = (type: string, target: string): string => {
    const ids = new Set(children(relationships.documentElement!).map((r) => r.getAttribute('Id')));
    let i = 1;
    while (ids.has(`rId${i}`)) i++;
    const node = relationships.createElementNS(PACKAGE_REL, 'Relationship');
    node.setAttribute('Id', `rId${i}`);
    node.setAttribute('Type', `${REL}/${type}`);
    node.setAttribute('Target', target);
    relationships.documentElement!.appendChild(node);
    return `rId${i}`;
  };
  /** Recibe ruta y tipo MIME; registra la parte si falta y no retorna valor. */
  const contentType = (name: string, mime: string): void => {
    if (
      children(types.documentElement!, 'Override').some(
        (n) => n.getAttribute('PartName') === `/${name}`,
      )
    )
      return;
    const node = types.createElementNS(TYPES, 'Override');
    node.setAttribute('PartName', `/${name}`);
    node.setAttribute('ContentType', mime);
    types.documentElement!.appendChild(node);
  };
  let targetPath: string;
  if (previous) {
    const sheet = sheetNodes.find((s) => s.getAttribute('name') === previous.targetSheet);
    if (!sheet) throw Error('PIVOT_NOT_FOUND');
    const rel = children(relationships.documentElement!).find(
      (r) => r.getAttribute('Id') === sheet.getAttribute('r:id'),
    );
    if (!rel) throw Error('PIVOT_NOT_FOUND');
    targetPath = resolve('xl/workbook.xml', rel.getAttribute('Target')!);
  } else {
    let i = 1;
    while (occupied.has(`xl/worksheets/sheet${i}.xml`)) i++;
    targetPath = `xl/worksheets/sheet${i}.xml`;
    const sheet = workbook.createElementNS(NS, 'sheet');
    sheet.setAttribute('name', pivot.targetSheet);
    sheet.setAttribute(
      'sheetId',
      String(Math.max(0, ...sheetNodes.map((s) => Number(s.getAttribute('sheetId')))) + 1),
    );
    sheet.setAttributeNS(REL, 'r:id', relationship('worksheet', targetPath.slice(3)));
    workbook.getElementsByTagNameNS(NS, 'sheets')[0].appendChild(sheet);
    let caches = workbook.getElementsByTagNameNS(NS, 'pivotCaches')[0];
    if (!caches) {
      caches = workbook.createElementNS(NS, 'pivotCaches');
      const after = children(workbook.documentElement!).find((n) =>
        [
          'smartTagPr',
          'smartTagTypes',
          'webPublishing',
          'fileRecoveryPr',
          'webPublishObjects',
          'extLst',
        ].includes(n.localName!),
      );
      workbook.documentElement!.insertBefore(caches, after ?? null);
    }
    const cache = workbook.createElementNS(NS, 'pivotCache');
    cache.setAttribute('cacheId', String(pivot.cacheId));
    cache.setAttributeNS(
      REL,
      'r:id',
      relationship('pivotCacheDefinition', pivot.cachePath.slice(3)),
    );
    caches.appendChild(cache);
  }
  // La hoja de salida es exclusiva de esta tabla; todas las otras hojas y recursos quedan intactos.
  const output = XLSX.utils.book_new();
  const pageRows: Cell[][] = input.filters.length
    ? [...input.filters.map((f) => [result.headers[f.field], f.value]), []]
    : [];
  const ws = XLSX.utils.aoa_to_sheet([...pageRows, ...result.matrix]);
  const tableRange = XLSX.utils.encode_range({
    s: { r: pageRows.length, c: 0 },
    e: { r: pageRows.length + result.matrix.length - 1, c: result.matrix[0].length - 1 },
  });
  ws['!cols'] = result.matrix[0].map(() => ({ wch: 20 }));
  XLSX.utils.book_append_sheet(output, ws, 'Output');
  const generated = unzipSync(
    XLSX.write(output, { type: 'buffer', bookType: 'xlsx', bookSST: false }),
  );
  const outputDoc = mergePivotOutput(
      archive,
      targetPath,
      parse(generated, 'xl/worksheets/sheet1.xml'),
      previous,
    ),
    outputRoot = outputDoc.documentElement!;
  const outputRels = parse(archive, relPath(targetPath), `<Relationships xmlns="${PACKAGE_REL}"/>`);
  const oldIds = new Set(
    children(outputRels.documentElement!)
      .filter(
        (r) =>
          r.getAttribute('Type') === `${REL}/pivotTable` &&
          resolve(targetPath, r.getAttribute('Target')!) === pivot.tablePath,
      )
      .map((r) => r.getAttribute('Id')),
  );
  children(outputRels.documentElement!)
    .filter((r) => oldIds.has(r.getAttribute('Id')))
    .forEach((r) => outputRels.documentElement!.removeChild(r));
  let relationId = 'rIdHelperPivot',
    suffix = 1;
  while (children(outputRels.documentElement!).some((r) => r.getAttribute('Id') === relationId))
    relationId = `rIdHelperPivot${suffix++}`;
  const outRel = outputRels.createElementNS(PACKAGE_REL, 'Relationship');
  outRel.setAttribute('Id', relationId);
  outRel.setAttribute('Type', `${REL}/pivotTable`);
  outRel.setAttribute(
    'Target',
    path.posix.relative(path.posix.dirname(targetPath), pivot.tablePath),
  );
  outputRels.documentElement!.appendChild(outRel);
  const parts =
    children(outputRoot, 'pivotParts')[0] ?? outputDoc.createElementNS(NS, 'pivotParts');
  children(parts, 'pivotPart')
    .filter((p) => oldIds.has(p.getAttribute('r:id')))
    .forEach((p) => parts.removeChild(p));
  const tablePart = outputDoc.createElementNS(NS, 'pivotPart');
  tablePart.setAttributeNS(REL, 'r:id', relationId);
  parts.appendChild(tablePart);
  parts.setAttribute('count', String(children(parts, 'pivotPart').length));
  if (!parts.parentNode) outputRoot.insertBefore(parts, children(outputRoot, 'extLst')[0] ?? null);
  store(archive, targetPath, outputDoc);
  store(archive, relPath(targetPath), outputRels);
  archive[relPath(pivot.tablePath)] = strToU8(
    `<Relationships xmlns="${PACKAGE_REL}"><Relationship Id="rId1" Type="${REL}/pivotCacheDefinition" Target="${xml(path.posix.relative(path.posix.dirname(pivot.tablePath), pivot.cachePath))}"/></Relationships>`,
  );
  archive[relPath(pivot.cachePath)] = strToU8(
    `<Relationships xmlns="${PACKAGE_REL}"><Relationship Id="rId1" Type="${REL}/pivotCacheRecords" Target="${xml(path.posix.basename(pivot.recordsPath))}"/></Relationships>`,
  );
  archive[pivot.cachePath] = strToU8(
    `<pivotCacheDefinition xmlns="${NS}" xmlns:r="${REL}" r:id="rId1" refreshOnLoad="0" saveData="1" enableRefresh="1" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3" recordCount="${result.records.length}"><cacheSource type="worksheet"><worksheetSource ref="${xml(result.definition.sourceRange)}" sheet="${xml(input.sourceSheet)}"/></cacheSource><cacheFields count="${result.headers.length}">${result.headers
      .map((header, i) => {
        const items = result.shared[i],
          strings = items.some((v) => typeof v === 'string'),
          numbers = items.some((v) => typeof v === 'number'),
          booleans = items.some((v) => typeof v === 'boolean'),
          blank = items.some((v) => v === null);
        return `<cacheField name="${xml(header)}" numFmtId="0"><sharedItems count="${items.length}" containsString="${strings ? 1 : 0}" containsNumber="${numbers ? 1 : 0}" containsBlank="${blank ? 1 : 0}" containsNonDate="1" containsSemiMixedTypes="${blank ? 1 : 0}" containsMixedTypes="${[strings, numbers, booleans].filter(Boolean).length > 1 ? 1 : 0}">${items.map(cacheValue).join('')}</sharedItems></cacheField>`;
      })
      .join('')}</cacheFields></pivotCacheDefinition>`,
  );
  const indices = result.shared.map(
    (items) => new Map(items.map((v, i) => [JSON.stringify(v), i])),
  );
  archive[pivot.recordsPath] = strToU8(
    `<pivotCacheRecords xmlns="${NS}" count="${result.records.length}">${result.records.map((record) => `<r>${record.map((v, i) => `<x v="${indices[i].get(JSON.stringify(v))}"/>`).join('')}</r>`).join('')}</pivotCacheRecords>`,
  );
  const multiple = input.values.length > 1,
    totalCols = input.totals && input.columns.length > 0;
  const item = (tuple: Cell[], fields: number[], dataIndex?: number) =>
    `<i${dataIndex === undefined ? '' : ` i="${dataIndex}"`}>${tuple.map((v, i) => `<x v="${indices[fields[i]].get(JSON.stringify(v))}"/>`).join('')}${multiple && dataIndex !== undefined ? `<x v="${dataIndex}"/>` : ''}</i>`;
  const rowItems =
    result.rows.map((row) => item(row, input.rows)).join('') +
    (input.totals ? '<i t="grand"><x/></i>' : '');
  const colItems =
    result.columns
      .map((col) => input.values.map((_, i) => item(col, input.columns, i)).join(''))
      .join('') +
    (totalCols ? input.values.map((_, i) => `<i t="grand" i="${i}"><x/></i>`).join('') : '');
  archive[pivot.tablePath] = strToU8(
    `<pivotTableDefinition xmlns="${NS}" name="HelperPivot_${xml(id.replace(/[^A-Za-z0-9_]/g, '_'))}" cacheId="${pivot.cacheId}" dataCaption="Valores" rowGrandTotals="${input.totals ? 1 : 0}" colGrandTotals="${input.totals ? 1 : 0}" compact="0" compactData="0" gridDropZones="1" outline="0" outlineData="0" showDrill="1" multipleFieldFilters="0" createdVersion="3" updatedVersion="3" minRefreshableVersion="3"><location ref="${tableRange}" rowPageCount="${input.filters.length}" colPageCount="${input.filters.length ? 1 : 0}" firstHeaderRow="0" firstDataRow="${result.headerRows}" firstDataCol="${input.rows.length}"/><pivotFields count="${result.headers.length}">${result.headers
      .map((_, i) => {
        const filter = input.filters.find((f) => f.field === i),
          axis = input.rows.includes(i)
            ? 'axisRow'
            : input.columns.includes(i)
              ? 'axisCol'
              : filter
                ? 'axisPage'
                : null;
        return `<pivotField${axis ? ` axis="${axis}"` : ''}${input.values.some((v) => v.field === i) ? ' dataField="1"' : ''} showAll="0" defaultSubtotal="0" compact="0" outline="0">${axis ? `<items count="${result.shared[i].length}">${result.shared[i].map((v, n) => `<item x="${n}"${filter && JSON.stringify(filter.value) !== JSON.stringify(v) ? ' h="1"' : ''}/>`).join('')}</items>` : ''}</pivotField>`;
      })
      .join(
        '',
      )}</pivotFields><rowFields count="${input.rows.length}">${input.rows.map((i) => `<field x="${i}"/>`).join('')}</rowFields><rowItems count="${result.rows.length + (input.totals ? 1 : 0)}">${rowItems}</rowItems>${input.columns.length || multiple ? `<colFields count="${input.columns.length + (multiple ? 1 : 0)}">${input.columns.map((i) => `<field x="${i}"/>`).join('')}${multiple ? '<field x="-2"/>' : ''}</colFields>` : ''}<colItems count="${result.columns.length * input.values.length + (totalCols ? input.values.length : 0)}">${colItems}</colItems>${input.filters.length ? `<pageFields count="${input.filters.length}">${input.filters.map((f) => `<pageField fld="${f.field}" hier="-1" item="${indices[f.field].get(JSON.stringify(f.value)) ?? 0}"/>`).join('')}</pageFields>` : ''}<dataFields count="${input.values.length}">${input.values.map((v) => `<dataField name="${xml(CAPTIONS[v.aggregate] + ' de ' + result.headers[v.field])}" fld="${v.field}" subtotal="${v.aggregate}" baseField="0" baseItem="0" numFmtId="2"/>`).join('')}</dataFields><pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" showColHeaders="1" showRowStripes="1" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`,
  );
  const entries = pivots.filter((p) => p.id !== id);
  entries.push(pivot);
  archive[METADATA] = strToU8(JSON.stringify(entries));
  contentType(
    targetPath,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  );
  contentType(
    pivot.tablePath,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  );
  contentType(
    pivot.cachePath,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  );
  contentType(
    pivot.recordsPath,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  );
  contentType(METADATA, 'application/json');
  store(archive, 'xl/workbook.xml', workbook);
  store(archive, 'xl/_rels/workbook.xml.rels', relationships);
  store(archive, '[Content_Types].xml', types);
  return { base64: Buffer.from(zipSync(archive)).toString('base64'), pivot };
}

/**
 * Actualiza todas las tablas creadas en Helper Notes después de guardar cambios de sus orígenes.
 * @param base64 Libro actual con las fórmulas recalculadas.
 * @returns Libro actualizado; cada origen inválido conserva su último resultado y publica un error por tabla.
 */
export function refreshSpreadsheetPivots(base64: string): { base64: string; warnings: string[] } {
  const warnings: string[] = [];
  const failures = new Map<string, string>();
  for (const pivot of readSpreadsheetPivots(base64)) {
    try {
      base64 = saveSpreadsheetPivot(base64, pivot).base64;
    } catch (error) {
      failures.set(pivot.id, error instanceof Error ? error.message : 'PIVOT_SOURCE');
      warnings.push(`${pivot.name}: ${error instanceof Error ? error.message : 'PIVOT_SOURCE'}`);
    }
  }
  if (failures.size) {
    const archive = openArchive(base64);
    const saved = readSpreadsheetPivots(base64).map((p) =>
      failures.has(p.id) ? { ...p, error: failures.get(p.id) } : p,
    );
    archive[METADATA] = strToU8(JSON.stringify(saved));
    base64 = Buffer.from(zipSync(archive)).toString('base64');
  }
  return { base64, warnings };
}
