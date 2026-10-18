import {
  spreadsheetShape,
  spreadsheetAt,
  spreadsheetTail,
  SpreadsheetValue,
} from './spreadsheet-array';
import type { SheetValue } from './spreadsheet-formula';
import { evaluateSpreadsheetDate, SpreadsheetDateOptions } from './spreadsheet-date';
import { evaluateSpreadsheetExtra } from './spreadsheet-extra-functions';
type Value = SpreadsheetValue;
export type FormulaLibraryContext = SpreadsheetDateOptions & {
  matches: (value: SheetValue, criterion: SheetValue) => boolean;
  compare: (left: SheetValue, right: SheetValue) => number;
};

/** Recibe escalar o matriz; retorna valores en orden de filas sin convertir sus tipos. */
function flat(value: Value | undefined): SheetValue[] {
  return Array.isArray(value) ? value.flat() : [value ?? null];
}
/** Recibe escalar o matriz; retorna su primer valor o null. */
function scalar(value: Value | undefined): SheetValue {
  return Array.isArray(value) ? (value[0]?.[0] ?? null) : (value ?? null);
}
/** Recibe un valor; indica si es un error de cálculo reconocido. */
export function isSpreadsheetError(value: unknown): boolean {
  return (
    typeof value === 'string' &&
    /^#(?:NULL!|REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|ERROR!|CYCLE!|SPILL!|CALC!|GETTING_DATA)$/.test(
      value,
    )
  );
}
/** Recibe escalar o matriz; retorna número finito o rechaza texto no convertible y errores. */
function number(value: Value | undefined): number {
  const v = scalar(value);
  if (isSpreadsheetError(v)) throw Error(String(v));
  const n = Number(v ?? 0);
  if (!Number.isFinite(n)) throw Error('#VALUE!');
  return n;
}
/** Recibe valores y rango esperado; rechaza matrices con formas distintas, sin truncar datos. */
function sameShape(values: Value[]): void {
  const shape = (v: Value) => JSON.stringify(spreadsheetShape(v));
  if (values.some((v) => shape(v) !== shape(values[0]))) throw Error('#VALUE!');
}
/** Recibe vector y dirección; retorna sus índices en el orden de búsqueda solicitado. */
function indices(length: number, reverse = false): number[] {
  return Array.from({ length }, (_, i) => (reverse ? length - 1 - i : i));
}

/** Recibe patrón Excel; retorna expresión regular con * y ? y escape ~, sin anclas para permitir HALLAR. */
export function spreadsheetWildcardPattern(pattern: string): string {
  let result = '';
  for (let i = 0; i < pattern.length; i++) {
    let character = pattern[i],
      literal = false;
    if (character === '~' && i + 1 < pattern.length) {
      character = pattern[++i];
      literal = true;
    }
    result +=
      !literal && character === '*'
        ? '[\\s\\S]*'
        : !literal && character === '?'
          ? '[\\s\\S]'
          : character.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return result;
}

/**
 * Calcula funciones adicionales sin ejecutar código procedente del documento.
 * @param name Nombre canónico inglés.
 * @param values Argumentos evaluados; las matrices conservan dimensiones y tipos.
 * @param context Comparación Excel, criterios y sistema de fechas del libro.
 * @returns Escalar calculado, o undefined si la función pertenece al núcleo/no está soportada.
 */
export function evaluateSpreadsheetLibrary(
  name: string,
  values: Value[],
  context: FormulaLibraryContext,
): Value | undefined {
  const n = (i: number, fallback = 0) => (values[i] === undefined ? fallback : number(values[i]));
  const text = (i: number) => String(scalar(values[i]) ?? '');
  const all = values.flatMap(flat),
    nums = all.filter((v): v is number => typeof v === 'number');
  const matrix = (i: number): SheetValue[][] =>
    Array.isArray(values[i]) ? (values[i] as SheetValue[][]) : [[scalar(values[i])]];
  const check = () => {
    const error = all.find(isSpreadsheetError);
    if (error) throw Error(String(error));
  };
  if (
    !/^(IS|SUBTOTAL$|COUNTBLANK$|SUMIFS$|COUNTIFS$|AVERAGEIFS$|MINIFS$|MAXIFS$|AVERAGEIF$|XLOOKUP$|XMATCH$|COUNT$|COUNTA$)/.test(
      name,
    )
  )
    check();
  switch (name) {
    case 'SUBTOTAL': {
      const code = n(0),
        kind = code > 100 ? code - 100 : code;
      if (
        !Number.isInteger(code) ||
        kind < 1 ||
        kind > 11 ||
        code > 111 ||
        (code > 11 && code < 101) ||
        values.length < 2
      )
        throw Error('#VALUE!');
      const cells = values.slice(1).flatMap(flat),
        numbers = cells.filter((v): v is number => typeof v === 'number');
      if (kind === 2) return numbers.length;
      if (kind === 3) return cells.filter((v) => v !== null).length;
      const failure = cells.find(isSpreadsheetError);
      if (failure) throw Error(String(failure));
      const sum = numbers.reduce((a, b) => a + b, 0);
      if (kind === 9) return sum;
      if (kind === 4) return numbers.length ? Math.max(...numbers) : 0;
      if (kind === 5) return numbers.length ? Math.min(...numbers) : 0;
      if (kind === 6) return numbers.length ? numbers.reduce((a, b) => a * b, 1) : 0;
      if (kind === 1) {
        if (!numbers.length) throw Error('#DIV/0!');
        return sum / numbers.length;
      }
      const divisor = numbers.length - ([7, 10].includes(kind) ? 1 : 0);
      if (divisor <= 0) throw Error('#DIV/0!');
      const mean = sum / numbers.length,
        variance = numbers.reduce((a, b) => a + (b - mean) ** 2, 0) / divisor;
      return [7, 8].includes(kind) ? Math.sqrt(variance) : variance;
    }
    case 'ISERROR':
      return isSpreadsheetError(scalar(values[0]));
    case 'ISERR':
      return isSpreadsheetError(scalar(values[0])) && scalar(values[0]) !== '#N/A';
    case 'ISNA':
      return scalar(values[0]) === '#N/A';
    case 'ISNUMBER':
      return typeof scalar(values[0]) === 'number';
    case 'ISTEXT':
      return typeof scalar(values[0]) === 'string' && !isSpreadsheetError(scalar(values[0]));
    case 'ISBLANK':
      return scalar(values[0]) === null;
    case 'ISLOGICAL':
      return typeof scalar(values[0]) === 'boolean';
    case 'N': {
      const v = scalar(values[0]);
      if (isSpreadsheetError(v)) throw Error(String(v));
      return typeof v === 'number' ? v : v === true ? 1 : 0;
    }
    case 'NA':
      return '#N/A';
    case 'TRUE':
      return true;
    case 'FALSE':
      return false;
    case 'COUNTBLANK':
      return (
        flat(values[0]).filter((v) => v === null || v === '').length + spreadsheetTail(values[0])
      );
    case 'SUMIFS':
    case 'COUNTIFS':
    case 'AVERAGEIFS':
    case 'MINIFS':
    case 'MAXIFS': {
      const offset = name === 'COUNTIFS' ? 0 : 1;
      if (values.length < offset + 2 || (values.length - offset) % 2 !== 0) throw Error('#VALUE!');
      const ranges = values.filter((_, i) => i >= offset && (i - offset) % 2 === 0);
      sameShape(offset ? [values[0], ...ranges] : ranges);
      const tests = ranges.map(flat),
        used = Math.max(...tests.map((t) => t.length)),
        matches = indices(used).filter((i) =>
          tests.every((range, j) =>
            context.matches(range[i] ?? null, scalar(values[offset + j * 2 + 1])),
          ),
        );
      if (name === 'COUNTIFS') {
        const shape = spreadsheetShape(ranges[0]),
          tail = shape.rows * shape.cols - used;
        return (
          matches.length +
          (tests.every((_, j) => context.matches(null, scalar(values[offset + j * 2 + 1])))
            ? tail
            : 0)
        );
      }
      const output = flat(values[0]),
        selected = matches.map((i) => output[i]);
      const error = selected.find(isSpreadsheetError);
      if (error) throw Error(String(error));
      const numbers = selected.filter((v): v is number => typeof v === 'number');
      if (name === 'AVERAGEIFS') {
        if (!numbers.length) throw Error('#DIV/0!');
        return numbers.reduce((a, b) => a + b, 0) / numbers.length;
      }
      if (name === 'MINIFS') return numbers.length ? Math.min(...numbers) : 0;
      if (name === 'MAXIFS') return numbers.length ? Math.max(...numbers) : 0;
      return numbers.reduce((a, b) => a + b, 0);
    }
    case 'AVERAGEIF': {
      if (values.length < 2) throw Error('#VALUE!');
      sameShape([values[0], values[2] ?? values[0]]);
      const range = flat(values[0]),
        sum = flat(values[2] ?? values[0]),
        selected = sum.filter((_, i) => context.matches(range[i], scalar(values[1])));
      const error = selected.find(isSpreadsheetError);
      if (error) throw Error(String(error));
      const numbers = selected.filter((v): v is number => typeof v === 'number');
      if (!numbers.length) throw Error('#DIV/0!');
      return numbers.reduce((a, b) => a + b, 0) / numbers.length;
    }
    case 'SUMPRODUCT': {
      if (!values.length) throw Error('#VALUE!');
      check();
      sameShape(values);
      const ranges = values.map(flat);
      return ranges[0].reduce<number>(
        (sum, _, i) =>
          sum +
          ranges.reduce<number>(
            (product, range) => product * (typeof range[i] === 'number' ? (range[i] as number) : 0),
            1,
          ),
        0,
      );
    }
    case 'XLOOKUP':
    case 'XMATCH': {
      const lookup = flat(values[1]);
      if (Array.isArray(values[1]) && matrix(1).length > 1 && matrix(1)[0].length > 1)
        throw Error('#VALUE!');
      const mode = n(name === 'XLOOKUP' ? 4 : 2),
        search = n(name === 'XLOOKUP' ? 5 : 3, 1);
      // Las búsquedas binarias requieren ordenación validada; se conservan para recálculo externo.
      if (![0, -1, 1, 2].includes(mode) || ![1, -1].includes(search)) throw Error('#NUM!');
      const query = scalar(values[0]);
      if (isSpreadsheetError(query)) throw Error(String(query));
      const lookupShape = spreadsheetShape(values[1]),
        logicalLength = lookupShape.rows * lookupShape.cols;
      const order = indices(lookup.length, search === -1);
      if (logicalLength > lookup.length) {
        if (search === -1) order.unshift(logicalLength - 1);
        else order.push(lookup.length);
      }
      const lookupAt = (i: number) => lookup[i] ?? null;
      let found = order.find((i) =>
        mode === 2
          ? context.matches(lookupAt(i), query)
          : context.compare(lookupAt(i), query) === 0,
      );
      if (found === undefined && (mode === -1 || mode === 1)) {
        for (const i of order) {
          const compare = context.compare(lookupAt(i), query);
          if (
            (mode === -1 ? compare < 0 : compare > 0) &&
            (found === undefined ||
              (mode === -1
                ? context.compare(lookupAt(i), lookupAt(found)) > 0
                : context.compare(lookup[i], lookup[found]) < 0))
          )
            found = i;
        }
      }
      if (found === undefined) throw Error('#N/A');
      if (name === 'XMATCH') return found + 1;
      const output = spreadsheetShape(values[2]);
      if (lookupShape.cols === 1 && output.rows === lookupShape.rows) {
        if (output.cols === 1) return spreadsheetAt(values[2], found, 0) ?? 0;
        return [
          Array.from({ length: output.cols }, (_, c) => spreadsheetAt(values[2], found, c) ?? 0),
        ];
      }
      if (lookupShape.rows === 1 && output.cols === lookupShape.cols) {
        if (output.rows === 1) return spreadsheetAt(values[2], 0, found) ?? 0;
        return Array.from({ length: output.rows }, (_, r) => [
          spreadsheetAt(values[2], r, found) ?? 0,
        ]);
      }
      throw Error('#VALUE!');
    }
    case 'ROWS':
      return spreadsheetShape(values[0]).rows;
    case 'COLUMNS':
      return spreadsheetShape(values[0]).cols;
    case 'MEDIAN':
    case 'LARGE':
    case 'SMALL':
    case 'STDEV.S':
    case 'STDEV.P':
    case 'STDEV':
    case 'STDEVP':
    case 'VAR.S':
    case 'VAR.P':
    case 'VAR':
    case 'VARP': {
      check();
      const input = (
        name === 'LARGE' || name === 'SMALL'
          ? flat(values[0]).filter((v): v is number => typeof v === 'number')
          : nums
      ).sort((a, b) => a - b);
      if (name === 'LARGE' || name === 'SMALL') {
        const k = Math.trunc(n(1));
        if (k < 1 || k > input.length) throw Error('#NUM!');
        return input[name === 'SMALL' ? k - 1 : input.length - k];
      }
      if (!input.length) throw Error(name === 'MEDIAN' ? '#NUM!' : '#DIV/0!');
      if (name === 'MEDIAN') {
        const mid = Math.floor(input.length / 2);
        return input.length % 2 ? input[mid] : (input[mid - 1] + input[mid]) / 2;
      }
      const sample = ['STDEV.S', 'STDEV', 'VAR.S', 'VAR'].includes(name),
        denominator = input.length - (sample ? 1 : 0);
      if (denominator <= 0) throw Error('#DIV/0!');
      const mean = input.reduce((a, b) => a + b, 0) / input.length,
        variance = input.reduce((sum, v) => sum + (v - mean) ** 2, 0) / denominator;
      return name.startsWith('VAR') ? variance : Math.sqrt(variance);
    }
    case 'PRODUCT':
      check();
      return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
    case 'SIGN':
      return Math.sign(n(0));
    case 'TRUNC':
      return Math.trunc(n(0) * 10 ** n(1)) / 10 ** n(1);
    case 'QUOTIENT':
      if (!n(1)) throw Error('#DIV/0!');
      return Math.trunc(n(0) / n(1));
    case 'CEILING.MATH':
    case 'FLOOR.MATH': {
      const value = n(0),
        step = Math.abs(n(1, 1));
      if (!step) return 0;
      const round = name === 'CEILING.MATH' ? Math.ceil : Math.floor;
      return value < 0 && n(2) !== 0 ? -round(-value / step) * step : round(value / step) * step;
    }
    case 'REPT': {
      const count = Math.trunc(n(1));
      if (count < 0 || text(0).length * count > 32767) throw Error('#VALUE!');
      return text(0).repeat(count);
    }
    case 'CLEAN':
      return text(0).replace(/[\x00-\x1f]/g, '');
    case 'PROPER':
      return text(0)
        .toLowerCase()
        .replace(/(^|[^\p{L}])\p{L}/gu, (v) => v.toUpperCase());
    case 'EXACT':
      check();
      return text(0) === text(1);
    case 'FIND':
    case 'SEARCH': {
      const start = Math.trunc(n(2, 1));
      if (start < 1 || start > text(1).length + 1) throw Error('#VALUE!');
      if (name === 'SEARCH') {
        const match = new RegExp(spreadsheetWildcardPattern(text(0)), 'i').exec(
          text(1).slice(start - 1),
        );
        if (!match) throw Error('#VALUE!');
        return start + match.index;
      }
      const at = text(1).indexOf(text(0), start - 1);
      if (at < 0) throw Error('#VALUE!');
      return at + 1;
    }
    case 'REPLACE': {
      const start = Math.trunc(n(1)),
        length = Math.trunc(n(2));
      if (start < 1 || length < 0) throw Error('#VALUE!');
      return text(0).slice(0, start - 1) + text(3) + text(0).slice(start - 1 + length);
    }
    case 'SUBSTITUTE': {
      const source = text(0),
        from = text(1),
        to = text(2);
      if (!from) return source;
      if (values[3] === undefined) return source.split(from).join(to);
      const occurrence = Math.trunc(n(3));
      if (occurrence < 1) throw Error('#VALUE!');
      let at = -1;
      for (let i = 0; i < occurrence; i++) {
        at = source.indexOf(from, at < 0 ? 0 : at + from.length);
        if (at < 0) return source;
      }
      return source.slice(0, at) + to + source.slice(at + from.length);
    }
    case 'TEXTJOIN': {
      check();
      if (
        !Boolean(scalar(values[1])) &&
        text(0).length &&
        values.slice(2).reduce<number>((sum, v) => sum + spreadsheetTail(v), 0) * text(0).length >
          32767
      )
        throw Error('#VALUE!');
      const output = values
        .slice(2)
        .flatMap(flat)
        .filter((v) => !Boolean(scalar(values[1])) || (v !== null && v !== ''))
        .map((v) => v ?? '')
        .join(text(0));
      if (output.length > 32767) throw Error('#VALUE!');
      return output;
    }
  }
  if (
    [
      'TODAY',
      'NOW',
      'DATE',
      'YEAR',
      'MONTH',
      'DAY',
      'DAYS',
      'EDATE',
      'EOMONTH',
      'WEEKDAY',
      'TIME',
      'HOUR',
      'MINUTE',
      'SECOND',
      'NETWORKDAYS',
      'WORKDAY',
    ].includes(name)
  ) {
    const holidays =
      name === 'WORKDAY' || name === 'NETWORKDAYS'
        ? flat(values[2]).filter((v): v is number => typeof v === 'number')
        : [];
    check();
    return evaluateSpreadsheetDate(name, values.slice(0, 3).map(number), holidays, context);
  }
  return evaluateSpreadsheetExtra(name, values);
}
