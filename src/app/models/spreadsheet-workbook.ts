import {
  canonicalSpreadsheetFormula,
  evaluateSpreadsheetResult,
  spreadsheetAddress,
  SheetValue,
} from './spreadsheet-formula';
import {
  spreadsheetColumnLabel,
  SpreadsheetTable,
  resolveSpreadsheetTable,
} from './spreadsheet-table';
import { isSpreadsheetError } from './spreadsheet-function-library';
export type SpreadsheetSpill = {
  sheet: number;
  row: number;
  col: number;
  rows: number;
  cols: number;
};
export type SpreadsheetCalculationSheet = {
  name: string;
  data: SheetValue[][];
  colCount: number;
  hiddenRows?: number[];
  hasActiveFilters?: boolean;
  date1904?: boolean;
  definedNames?: Record<string, string>;
  tables?: SpreadsheetTable[];
  merges?: { s: { r: number; c: number }; e: { r: number; c: number } }[];
};
export type SpreadsheetFormulaCache = { formula: string; value: SheetValue };

/** Recibe descriptor de matriz; retorna sus coordenadas base cero en orden de filas sin incluir el ancla si se solicita. */
export function spreadsheetSpillCells(
  spill: SpreadsheetSpill,
  includeAnchor = false,
): { row: number; col: number }[] {
  return Array.from({ length: spill.rows * spill.cols }, (_, i) => ({
    row: spill.row + Math.floor(i / spill.cols),
    col: spill.col + (i % spill.cols),
  })).filter((c) => includeAnchor || c.row !== spill.row || c.col !== spill.col);
}

/**
 * Recalcula fórmulas y desbordamientos sin sobrescribir celdas ocupadas, combinadas o pertenecientes a tablas.
 * @param sheets Hojas mutables que recibirán resultados escalares y celdas derivadas.
 * @param formulas Expresiones indexadas como hoja:fila:columna.
 * @param caches Resultados importados de expresiones no soportadas, ligados a su expresión original.
 * @param previousSpills Matrices previas cuyos resultados derivados se pueden limpiar y regenerar.
 * @returns Propietarios de celdas derivadas, matrices actuales, anclas nativas (incluidas las bloqueadas) y cantidad de cachés conservadas.
 */
export function recalculateSpreadsheetWorkbook(
  sheets: SpreadsheetCalculationSheet[],
  formulas: Map<string, string>,
  caches: Map<string, SpreadsheetFormulaCache>,
  previousSpills: Map<string, SpreadsheetSpill>,
) {
  let previous = previousSpills,
    spills = new Map<string, SpreadsheetSpill>(),
    arrays = new Map<string, SpreadsheetSpill>(),
    owners = new Map<string, string>(),
    cachedCount = 0;
  const cachedKeys = new Set<string>();
  const now = new Date(),
    tables = sheets.flatMap((s) => s.tables ?? []),
    hidden = sheets.map((s) => new Set(s.hiddenRows ?? [])),
    key = (s: number, r: number, c: number) => `${s}:${r}:${c}`;
  const set = (s: number, r: number, c: number, value: SheetValue) => {
    while (sheets[s].data.length <= r) sheets[s].data.push([]);
    sheets[s].data[r][c] = value;
  };
  const signature = (map: Map<string, SpreadsheetSpill>) =>
    JSON.stringify(
      [...map].map(([k, v]) => [
        k,
        v,
        sheets[v.sheet].data
          .slice(v.row, v.row + v.rows)
          .map((r) => r.slice(v.col, v.col + v.cols)),
      ]),
    );
  for (let pass = 0; pass < 8; pass++) {
    const priorSignature = signature(previous),
      oldOwners = new Map<string, string>();
    const priorValues = new Map(
      [...previous].map(([anchor, spill]) => [
        anchor,
        Array.from({ length: spill.rows }, (_, r) =>
          Array.from(
            { length: spill.cols },
            (_, c) => sheets[spill.sheet].data[spill.row + r]?.[spill.col + c] ?? null,
          ),
        ),
      ]),
    );
    for (const [anchor, spill] of previous)
      for (const cell of spreadsheetSpillCells(spill)) {
        const k = key(spill.sheet, cell.row, cell.col);
        oldOwners.set(k, anchor);
        if (!formulas.has(k)) set(spill.sheet, cell.row, cell.col, null);
      }
    spills = new Map();
    arrays = new Map();
    owners = new Map();
    cachedCount = 0;
    cachedKeys.clear();
    const complete = new Map<string, SheetValue>(),
      visiting = new Set<string>();
    const usedRows = Math.max(1, ...sheets.map((s) => s.data.length));
    /** Recibe coordenadas; retorna resultado de celda o de su ancla, resolviendo referencias entre hojas de forma perezosa. */
    const calculate = (index: number, row: number, col: number): SheetValue => {
      const k = key(index, row, col);
      if (complete.has(k)) return complete.get(k)!;
      if (visiting.has(k)) return '#CYCLE!';
      const sheet = sheets[index];
      if (!sheet) return '#REF!';
      const formula = formulas.get(k);
      if (!formula) {
        const owner = owners.get(k) ?? oldOwners.get(k);
        if (owner) {
          const [s, r, c] = owner.split(':').map(Number);
          calculate(s, r, c);
        }
        return sheet.data[row]?.[col] ?? null;
      }
      visiting.add(k);
      const resolve = (reference: string) => {
        const address = spreadsheetAddress(reference);
        if (!address) return null;
        const target =
          address.sheet === undefined
            ? index
            : sheets.findIndex((s) => s.name.toLowerCase() === address.sheet!.toLowerCase());
        return target < 0 ? null : { ...address, sheet: target };
      };
      const result = evaluateSpreadsheetResult(
        formula,
        (reference) => {
          const a = resolve(reference);
          return a ? calculate(a.sheet, a.row, a.col) : '#REF!';
        },
        {
          date1904: sheet.date1904,
          now,
          row,
          col,
          usedRows,
          omitSubtotalCell: (reference, ignoreHidden) => {
            const a = resolve(reference);
            if (!a) return false;
            // El archivo no distingue ocultación manual de filtrada: los códigos 1–11 conservan su caché si hay filtros activos.
            if (!ignoreHidden && sheets[a.sheet].hasActiveFilters) throw Error('#NAME?');
            return (
              !!(ignoreHidden && hidden[a.sheet].has(a.row)) ||
              /\bSUBTOTAL\s*\(/i.test(formulas.get(key(a.sheet, a.row, a.col)) ?? '')
            );
          },
          resolveName: (name) => sheet.definedNames?.[name.toUpperCase()],
          resolveStructured: (reference) =>
            resolveSpreadsheetTable(reference, tables, sheet.name, row, col),
          readSpill: (reference) => {
            const a = resolve(reference);
            if (!a) return '#REF!';
            const origin = key(a.sheet, a.row, a.col);
            calculate(a.sheet, a.row, a.col);
            const spill = spills.get(origin);
            if (!spill) return '#REF!';
            return Array.from({ length: spill.rows }, (_, r) =>
              Array.from(
                { length: spill.cols },
                (_, c) => sheets[a.sheet].data[a.row + r]?.[a.col + c] ?? null,
              ),
            );
          },
        },
      );
      let value: SheetValue;
      if (Array.isArray(result)) {
        const spill = { sheet: index, row, col, rows: result.length, cols: result[0]?.length ?? 0 };
        arrays.set(k, { ...spill, rows: 1, cols: 1 });
        const cells = spreadsheetSpillCells(spill, true);
        const blocked =
          row + spill.rows > 100000 ||
          col + spill.cols > 16384 ||
          cells.some((cell) => {
            const targetKey = key(index, cell.row, cell.col);
            if (
              sheet.merges?.some(
                (m) =>
                  cell.row >= m.s.r && cell.row <= m.e.r && cell.col >= m.s.c && cell.col <= m.e.c,
              )
            )
              return true;
            if (
              (sheet.tables ?? []).some((t) => {
                const [a, b] = t.ref.split(':').map(spreadsheetAddress);
                return (
                  a &&
                  b &&
                  cell.row >= a.row &&
                  cell.row <= b.row &&
                  cell.col >= a.col &&
                  cell.col <= b.col
                );
              })
            )
              return true;
            if (targetKey === k) return false;
            return (
              formulas.has(targetKey) ||
              owners.has(targetKey) ||
              (sheet.data[cell.row]?.[cell.col] !== null &&
                sheet.data[cell.row]?.[cell.col] !== undefined)
            );
          });
        if (blocked) value = '#SPILL!';
        else {
          spills.set(k, spill);
          arrays.set(k, spill);
          result.forEach((values, r) =>
            values.forEach((v, c) => {
              const output = v ?? 0;
              set(index, row + r, col + c, output);
              if (r || c) owners.set(key(index, row + r, col + c), k);
            }),
          );
          sheet.colCount = Math.max(sheet.colCount, col + spill.cols);
          value = result[0]?.[0] ?? 0;
        }
      } else {
        const imported = caches.get(k),
          keep =
            (result === '#NAME?' || result === '#ERROR!') &&
            imported?.value != null &&
            canonicalSpreadsheetFormula(formula) === imported.formula;
        if (keep) {
          cachedCount++;
          cachedKeys.add(k);
          const prior = previous.get(k),
            values = priorValues.get(k);
          if (prior && values) {
            spills.set(k, prior);
            arrays.set(k, prior);
            values.forEach((values, r) =>
              values.forEach((v, c) => {
                set(index, row + r, col + c, v);
                if (r || c) owners.set(key(index, row + r, col + c), k);
              }),
            );
          }
        }
        value = keep ? imported!.value : result;
        if (
          /\b(?:FILTER|SORT|UNIQUE|SEQUENCE|TRANSPOSE)\s*\(/i.test(
            canonicalSpreadsheetFormula(formula),
          )
        )
          arrays.set(k, { sheet: index, row, col, rows: 1, cols: 1 });
      }
      visiting.delete(k);
      complete.set(k, value);
      set(index, row, col, value);
      return value;
    };
    for (const k of formulas.keys()) {
      const [s, r, c] = k.split(':').map(Number);
      calculate(s, r, c);
    }
    if (signature(spills) === priorSignature || (!spills.size && !previous.size)) break;
    previous = spills;
    if (pass === 7) {
      for (const spill of spills.values())
        for (const cell of spreadsheetSpillCells(spill)) set(spill.sheet, cell.row, cell.col, null);
      for (const [k, spill] of arrays) {
        set(spill.sheet, spill.row, spill.col, '#CYCLE!');
      }
      spills.clear();
      owners.clear();
    }
  }
  return { spills, arrays, owners, cachedCount, cachedKeys };
}
