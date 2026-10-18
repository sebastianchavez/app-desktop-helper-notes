import type { SpreadsheetValue } from './spreadsheet-array';
import type { SheetValue } from './spreadsheet-formula';
const EXTRA_ARITY: Record<string, [number, number]> = {
  PMT: [3, 5],
  PV: [3, 5],
  FV: [3, 5],
  NPER: [3, 5],
  IPMT: [4, 6],
  PPMT: [4, 6],
  NPV: [2, 255],
  MROUND: [2, 2],
  LN: [1, 1],
  LOG10: [1, 1],
  LOG: [1, 2],
  EXP: [1, 1],
  FACT: [1, 1],
  GCD: [1, 255],
  LCM: [1, 255],
  T: [1, 1],
  NUMBERVALUE: [1, 3],
  TEXTBEFORE: [2, 6],
  TEXTAFTER: [2, 6],
};

/**
 * Calcula funciones financieras, numéricas y de texto sin dependencias de Office.
 * @param name Nombre canónico inglés.
 * @param args Argumentos evaluados; los rangos conservan texto, números y blancos.
 * @returns Valor calculado o undefined para funciones ajenas; lanza errores Excel ante entradas inválidas.
 */
export function evaluateSpreadsheetExtra(
  name: string,
  args: SpreadsheetValue[],
): SheetValue | undefined {
  const arity = EXTRA_ARITY[name];
  if (arity && (args.length < arity[0] || args.length > arity[1])) throw Error('#VALUE!');
  const scalar = (i: number): SheetValue =>
    Array.isArray(args[i])
      ? ((args[i] as SheetValue[][])[0]?.[0] ?? null)
      : ((args[i] as SheetValue) ?? null);
  const num = (i: number, fallback = 0): number => {
    if (args[i] === undefined) return fallback;
    const v = scalar(i);
    if (typeof v === 'string' && v.startsWith('#')) throw Error(v);
    const n = Number(v);
    if (!Number.isFinite(n)) throw Error('#VALUE!');
    return n;
  };
  const text = (i: number, fallback = ''): string =>
    args[i] === undefined ? fallback : String(scalar(i) ?? '');
  const finite = (value: number): number => {
    if (!Number.isFinite(value)) throw Error('#NUM!');
    return value;
  };
  if (['PMT', 'PV', 'FV', 'NPER', 'IPMT', 'PPMT'].includes(name)) {
    const rate = num(0),
      nper = num(1),
      pmtOrPv = num(2),
      fourth = num(3),
      type = num(name === 'IPMT' || name === 'PPMT' ? 5 : 4);
    if (![0, 1].includes(type) || rate <= -1) throw Error('#NUM!');
    if (name === 'NPER') {
      const pmt = num(1),
        pv = num(2),
        fv = num(3);
      if (rate === 0) {
        if (pmt === 0) throw Error('#DIV/0!');
        return -(pv + fv) / pmt;
      }
      const a = (pmt * (1 + rate * type)) / rate;
      return finite(Math.log((a - fv) / (pv + a)) / Math.log1p(rate));
    }
    if (name === 'IPMT' || name === 'PPMT') {
      const per = num(1),
        periods = num(2),
        pv = num(3),
        fv = num(4),
        when = num(5);
      if (![0, 1].includes(when) || per < 1 || per > periods || periods <= 0) throw Error('#NUM!');
      const payment = evaluateSpreadsheetExtra('PMT', [rate, periods, pv, fv, when]) as number;
      const balance =
        rate === 0
          ? pv + payment * (per - 1)
          : pv * Math.pow(1 + rate, per - 1) +
            (payment * (1 + rate * when) * Math.expm1((per - 1) * Math.log1p(rate))) / rate;
      const interest = when && per === 1 ? 0 : (-balance * rate) / (1 + rate * when);
      return finite(name === 'IPMT' ? interest : payment - interest);
    }
    const factor = Math.pow(1 + rate, nper),
      annuity =
        rate === 0 ? nper : (Math.expm1(nper * Math.log1p(rate)) / rate) * (1 + rate * type);
    if (name === 'FV') return finite(-(pmtOrPv * annuity + fourth * factor));
    if (name === 'PV') return finite(-(pmtOrPv * annuity + fourth) / factor);
    if (annuity === 0) throw Error('#DIV/0!');
    return finite(-(pmtOrPv * factor + fourth) / annuity);
  }
  switch (name) {
    case 'NPV': {
      const rate = num(0);
      if (rate === -1) throw Error('#DIV/0!');
      const values = args
        .slice(1)
        .flatMap((v) => (Array.isArray(v) ? v.flat() : [v]))
        .filter((v): v is number => typeof v === 'number');
      return finite(values.reduce((sum, v, i) => sum + v / Math.pow(1 + rate, i + 1), 0));
    }
    case 'MROUND': {
      const n = num(0),
        multiple = num(1);
      if (n * multiple < 0) throw Error('#NUM!');
      return multiple === 0
        ? 0
        : finite(
            Math.sign(n) *
              Math.floor(Math.abs(n / multiple) + 0.5 + Number.EPSILON) *
              Math.abs(multiple),
          );
    }
    case 'LN':
      if (num(0) <= 0) throw Error('#NUM!');
      return Math.log(num(0));
    case 'LOG10':
      if (num(0) <= 0) throw Error('#NUM!');
      return Math.log10(num(0));
    case 'LOG': {
      const n = num(0),
        base = num(1, 10);
      if (n <= 0 || base <= 0) throw Error('#NUM!');
      if (base === 1) throw Error('#DIV/0!');
      return Math.log(n) / Math.log(base);
    }
    case 'EXP':
      return finite(Math.exp(num(0)));
    case 'FACT': {
      const n = Math.trunc(num(0));
      if (num(0) < 0 || n > 170) throw Error('#NUM!');
      let f = 1;
      for (let i = 2; i <= n; i++) f *= i;
      return f;
    }
    case 'GCD':
    case 'LCM': {
      const values = args
        .flatMap((v) => (Array.isArray(v) ? v.flat() : [v]))
        .filter((v) => typeof v === 'number') as number[];
      if (values.some((v) => v < 0 || v > Number.MAX_SAFE_INTEGER)) throw Error('#NUM!');
      /** Recibe enteros no negativos; retorna su máximo común divisor mediante Euclides. */
      const gcd = (a: number, b: number): number => {
        while (b) {
          const n = a % b;
          a = b;
          b = n;
        }
        return a;
      };
      const result = values
        .map(Math.trunc)
        .reduce(
          (a, b) =>
            name === 'GCD' ? gcd(a, b) : a === 0 || b === 0 ? 0 : finite((a / gcd(a, b)) * b),
          name === 'GCD' ? 0 : 1,
        );
      if (result > Number.MAX_SAFE_INTEGER) throw Error('#NUM!');
      return result;
    }
    case 'T':
      return typeof scalar(0) === 'string' ? scalar(0) : '';
    case 'NUMBERVALUE': {
      const decimal = text(1, '.'),
        group = text(2, ',');
      if (!decimal || !group || decimal[0] === group[0]) throw Error('#VALUE!');
      let value = text(0).replace(/\s/g, ''),
        percent = 0;
      while (value.endsWith('%')) {
        percent++;
        value = value.slice(0, -1);
      }
      const point = value.indexOf(decimal[0]);
      if (point >= 0 && value.slice(point + 1).includes(group[0])) throw Error('#VALUE!');
      value = value.split(group[0]).join('').split(decimal[0]).join('.');
      if (!value) return 0;
      if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value)) throw Error('#VALUE!');
      return finite(Number(value) / Math.pow(100, percent));
    }
    case 'TEXTBEFORE':
    case 'TEXTAFTER': {
      const input = text(0),
        delimiter = text(1),
        instance = Math.trunc(num(2, 1)),
        match = num(3),
        end = num(4);
      if (!instance || ![0, 1].includes(match) || ![0, 1].includes(end)) throw Error('#VALUE!');
      if (!input) return '';
      if (!delimiter)
        return name === 'TEXTBEFORE' ? (instance > 0 ? '' : input) : instance > 0 ? input : '';
      const haystack = match ? input.toLowerCase() : input,
        needle = match ? delimiter.toLowerCase() : delimiter,
        positions: number[] = [];
      for (
        let at = haystack.indexOf(needle);
        at >= 0;
        at = haystack.indexOf(needle, at + needle.length)
      )
        positions.push(at);
      if (end) instance > 0 ? positions.push(input.length) : positions.unshift(-delimiter.length);
      const at = instance > 0 ? positions[instance - 1] : positions[positions.length + instance];
      if (at === undefined) {
        if (args[5] !== undefined) return scalar(5);
        throw Error('#N/A');
      }
      return name === 'TEXTBEFORE'
        ? input.slice(0, Math.max(0, at))
        : input.slice(at + delimiter.length);
    }
  }
  return undefined;
}
