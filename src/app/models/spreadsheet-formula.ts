import {
  evaluateSpreadsheetArray,
  mapSpreadsheetArray,
  spreadsheetShape,
  spreadsheetAt,
  markSpreadsheetRange,
  materializeSpreadsheetArray,
  spreadsheetTail,
  SpreadsheetValue,
} from './spreadsheet-array';
import { SUPPORTED_SPREADSHEET_FUNCTIONS } from './spreadsheet-supported-functions';
import {
  evaluateSpreadsheetLibrary,
  isSpreadsheetError,
  spreadsheetWildcardPattern,
} from './spreadsheet-function-library';
import type { SpreadsheetDateOptions } from './spreadsheet-date';
export type SheetValue = string | number | boolean | null;
/** Contexto del libro: fechas, nombres definidos y posición actual para FILA/COLUMNA. */
export type SpreadsheetFormulaOptions = SpreadsheetDateOptions & {
  usedRows?: number;
  omitSubtotalCell?: (reference: string, ignoreHidden: boolean) => boolean;
  readSpill?: (reference: string) => SheetValue[][] | SheetValue;
  resolveStructured?: (reference: string) => string | undefined;
  resolveName?: (name: string) => string | undefined;
  row?: number;
  col?: number;
  nameStack?: Set<string>;
};
type RangeValue = SheetValue[][];
type Value = SheetValue | RangeValue;
type Node =
  | { kind: 'literal'; value: SheetValue }
  | { kind: 'ref'; value: string }
  | { kind: 'whole'; value: string }
  | { kind: 'spill'; value: string }
  | { kind: 'name'; value: string }
  | { kind: 'range'; start: string; end: string }
  | { kind: 'call'; name: string; args: Node[] }
  | { kind: 'binary'; op: string; left: Node; right: Node }
  | { kind: 'unary'; op: string; node: Node };
const aliases: Record<string, string> = {
  SUBTOTALES: 'SUBTOTAL',
  PAGO: 'PMT',
  VA: 'PV',
  VF: 'FV',
  NPER: 'NPER',
  PAGOINT: 'IPMT',
  PAGOPRIN: 'PPMT',
  VNA: 'NPV',
  'REDOND.MULT': 'MROUND',
  'M.C.D': 'GCD',
  'M.C.M': 'LCM',
  'VALOR.NUMERO': 'NUMBERVALUE',
  TEXTOANTES: 'TEXTBEFORE',
  TEXTODESPUES: 'TEXTAFTER',
  FILTRAR: 'FILTER',
  ORDENAR: 'SORT',
  UNICOS: 'UNIQUE',
  SECUENCIA: 'SEQUENCE',
  TRANSPONER: 'TRANSPOSE',
  'SUMAR.SI.CONJUNTO': 'SUMIFS',
  'CONTAR.SI.CONJUNTO': 'COUNTIFS',
  'PROMEDIO.SI': 'AVERAGEIF',
  'PROMEDIO.SI.CONJUNTO': 'AVERAGEIFS',
  'MIN.SI.CONJUNTO': 'MINIFS',
  'MAX.SI.CONJUNTO': 'MAXIFS',
  SUMAPRODUCTO: 'SUMPRODUCT',
  BUSCARX: 'XLOOKUP',
  COINCIDIRX: 'XMATCH',
  'SI.ND': 'IFNA',
  'SI.CONJUNTO': 'IFS',
  CAMBIAR: 'SWITCH',
  'CONTAR.BLANCO': 'COUNTBLANK',
  ESERROR: 'ISERROR',
  ESERR: 'ISERR',
  ESNOD: 'ISNA',
  ESNUMERO: 'ISNUMBER',
  ESTEXTO: 'ISTEXT',
  ESBLANCO: 'ISBLANK',
  ESLOGICO: 'ISLOGICAL',
  NOD: 'NA',
  VERDADERO: 'TRUE',
  FALSO: 'FALSE',
  MEDIANA: 'MEDIAN',
  'K.ESIMO.MAYOR': 'LARGE',
  'K.ESIMO.MENOR': 'SMALL',
  'DESVEST.M': 'STDEV.S',
  'DESVEST.P': 'STDEV.P',
  DESVEST: 'STDEV',
  DESVESTP: 'STDEVP',
  'VAR.S': 'VAR.S',
  PRODUCTO: 'PRODUCT',
  SIGNO: 'SIGN',
  TRUNCAR: 'TRUNC',
  COCIENTE: 'QUOTIENT',
  'MULTIPLO.SUPERIOR.MAT': 'CEILING.MATH',
  'MULTIPLO.INFERIOR.MAT': 'FLOOR.MATH',
  REPETIR: 'REPT',
  LIMPIAR: 'CLEAN',
  NOMPROPIO: 'PROPER',
  IGUAL: 'EXACT',
  ENCONTRAR: 'FIND',
  HALLAR: 'SEARCH',
  REEMPLAZAR: 'REPLACE',
  SUSTITUIR: 'SUBSTITUTE',
  UNIRCADENAS: 'TEXTJOIN',
  FECHA: 'DATE',
  HOY: 'TODAY',
  AHORA: 'NOW',
  AÑO: 'YEAR',
  MES: 'MONTH',
  DIA: 'DAY',
  DIAS: 'DAYS',
  'FECHA.MES': 'EDATE',
  'FIN.MES': 'EOMONTH',
  DIASEM: 'WEEKDAY',
  NSHORA: 'TIME',
  HORA: 'HOUR',
  MINUTO: 'MINUTE',
  SEGUNDO: 'SECOND',
  'DIAS.LAB': 'NETWORKDAYS',
  'DIA.LAB': 'WORKDAY',
  FILAS: 'ROWS',
  COLUMNAS: 'COLUMNS',
  FILA: 'ROW',
  COLUMNA: 'COLUMN',
  RAIZ: 'SQRT',
  POTENCIA: 'POWER',
  RESIDUO: 'MOD',
  'REDONDEAR.MAS': 'ROUNDUP',
  'REDONDEAR.MENOS': 'ROUNDDOWN',
  ENTERO: 'INT',
  SUMA: 'SUM',
  PROMEDIO: 'AVERAGE',
  CONTAR: 'COUNT',
  CONTARA: 'COUNTA',
  SI: 'IF',
  SIERROR: 'IFERROR',
  'SI.ERROR': 'IFERROR',
  'CONTAR.SI': 'COUNTIF',
  'SUMAR.SI': 'SUMIF',
  REDONDEAR: 'ROUND',
  Y: 'AND',
  O: 'OR',
  NO: 'NOT',
  BUSCARV: 'VLOOKUP',
  BUSCARH: 'HLOOKUP',
  INDICE: 'INDEX',
  COINCIDIR: 'MATCH',
  CONCATENAR: 'CONCATENATE',
  LARGO: 'LEN',
  IZQUIERDA: 'LEFT',
  DERECHA: 'RIGHT',
  EXTRAE: 'MID',
  ESPACIOS: 'TRIM',
  MAYUSC: 'UPPER',
  MINUSC: 'LOWER',
};

/** Recibe expresión canónica; retorna tokens, manteniendo juntas referencias estructuradas anidadas, nombres escapados y rangos completos. */
function tokenizeSpreadsheetFormula(source: string): string[] {
  const result: string[] = [];
  const token =
    /"(?:[^"]|"")*"|(?:(?:'(?:[^']|'')+'|[\p{L}_][\p{L}\d_.]*)!)?\$?(?:[A-Z]{1,3}:\$?[A-Z]{1,3}|\d+:\$?\d+)|(?:'(?:[^']|'')+'|[\p{L}_][\p{L}\d_.]*)!\$?[A-Z]{1,3}\$?\d+|\$?[A-Z]{1,3}\$?\d+|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|#[A-Z0-9/?!]+|[\p{L}_][\p{L}\d_.]*|<>|>=|<=|[^\s]/giuy;
  let index = 0;
  while (index < source.length) {
    if (/\s/.test(source[index])) {
      index++;
      continue;
    }
    const prefix = source.slice(index).match(/^(?:[\p{L}_][\p{L}\d_.]*)?\[/u);
    if (prefix) {
      let depth = 0,
        end = index + prefix[0].length - 1;
      do {
        const character = source[end++];
        if (character === "'" && /[\[\]#@']/.test(source[end] ?? '')) {
          end++;
          continue;
        }
        if (character === '[') depth++;
        if (character === ']') depth--;
      } while (depth > 0 && end < source.length);
      if (depth !== 0) throw Error('#REF!');
      result.push(source.slice(index, end));
      index = end;
      continue;
    }
    token.lastIndex = index;
    const match = token.exec(source);
    if (!match) throw Error('#ERROR!');
    result.push(match[0]);
    index = token.lastIndex;
  }
  return result;
}

/** Recibe dos escalares; retorna orden Excel: números antes que texto, texto sin distinguir mayúsculas y booleanos al final. */
export function compareSpreadsheetValues(left: SheetValue, right: SheetValue): number {
  if (left === null) left = typeof right === 'string' ? '' : typeof right === 'boolean' ? false : 0;
  if (right === null) right = typeof left === 'string' ? '' : typeof left === 'boolean' ? false : 0;
  const type = (value: SheetValue) =>
    typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2;
  if (type(left) !== type(right)) return type(left) - type(right);
  const a = typeof left === 'string' ? left.toLocaleLowerCase() : left!,
    b = typeof right === 'string' ? right.toLocaleLowerCase() : right!;
  return a === b ? 0 : a < b ? -1 : 1;
}

/** Recibe celda y criterio; retorna coincidencia con operadores y comodines Excel, incluyendo ~ para literales. */
export function spreadsheetMatchesCriteria(value: SheetValue, criterion: SheetValue): boolean {
  if (isSpreadsheetError(criterion)) throw Error(String(criterion));
  if (typeof criterion !== 'string') return compareSpreadsheetValues(value, criterion ?? 0) === 0;
  const match = criterion.match(/^(>=|<=|<>|>|<|=)?(.*)$/s)!,
    op = match[1] ?? '=',
    text = match[2];
  if ((op === '=' || op === '<>') && /[*?~]/.test(text)) {
    const result =
      typeof value === 'string' &&
      new RegExp('^' + spreadsheetWildcardPattern(text) + '$', 'i').test(value);
    return op === '=' ? result : !result;
  }
  let right: SheetValue = text;
  if (text.trim() !== '' && Number.isFinite(Number(text))) {
    right = Number(text);
    if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value)))
      value = Number(value);
  }
  const compared = compareSpreadsheetValues(value, right);
  return op === '='
    ? compared === 0
    : op === '<>'
      ? compared !== 0
      : op === '>'
        ? compared > 0
        : op === '<'
          ? compared < 0
          : op === '>='
            ? compared >= 0
            : compared <= 0;
}

/** Recibe fórmula localizada; retorna funciones modernas con sus prefijos OOXML, preservando textos y nombres de hoja. */
export function ooxmlSpreadsheetFormula(formula: string): string {
  const future = new Set([
    'TEXTBEFORE',
    'TEXTAFTER',
    'NUMBERVALUE',
    'FILTER',
    'SORT',
    'UNIQUE',
    'SEQUENCE',
    'ANCHORARRAY',
    'XLOOKUP',
    'XMATCH',
    'IFS',
    'SWITCH',
    'IFNA',
    'CONCAT',
    'TEXTJOIN',
    'MINIFS',
    'MAXIFS',
    'CEILING.MATH',
    'FLOOR.MATH',
    'STDEV.S',
    'STDEV.P',
    'VAR.S',
    'VAR.P',
  ]);
  const tokens = tokenizeSpreadsheetFormula(canonicalSpreadsheetFormula(formula));
  const output: string[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (spreadsheetAddress(token) && tokens[i + 1] === '#') {
      output.push('_xlfn.ANCHORARRAY(' + token + ')');
      i++;
      continue;
    }
    output.push(
      tokens[i + 1] === '(' && future.has(token)
        ? (token === 'FILTER' || token === 'SORT' ? '_xlfn._xlws.' : '_xlfn.') + token
        : token,
    );
  }
  return output.join('');
}

/** Recibe fórmula localizada; retorna nombres OOXML sin modificar textos entre comillas. */
export function canonicalSpreadsheetFormula(formula: string): string {
  try {
    const tokens = tokenizeSpreadsheetFormula(formula);
    return tokens
      .map((token, i) => {
        if (tokens[i + 1] !== '(' || !/^[\p{L}_][\p{L}\d_.]*$/u.test(token))
          return token === ';' ? ',' : token;
        const normalized = token.toUpperCase();
        return (aliases[normalized] ?? normalized)
          .replace(/^_XLFN\./, '_xlfn.')
          .replace(/_XLWS\./g, '_xlws.');
      })
      .join('');
  } catch {
    return formula;
  }
}

/** Recibe referencia A1 opcionalmente calificada; retorna hoja, fila y columna con base cero. */
export function spreadsheetAddress(
  reference: string,
): { sheet?: string; row: number; col: number } | null {
  const match = reference.match(/^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?\$?([A-Z]{1,3})\$?(\d+)$/i);
  if (!match) return null;
  let col = 0;
  for (const c of match[3].toUpperCase()) col = col * 26 + c.charCodeAt(0) - 64;
  const row = Number(match[4]) - 1;
  return row >= 0 && row < 1048576 && col <= 16384
    ? { sheet: match[1]?.replace(/''/g, "'") ?? match[2], row, col: col - 1 }
    : null;
}

/**
 * Evalúa fórmulas sin ejecutar JavaScript: operadores, funciones anidadas, rangos y referencias entre hojas.
 * @param formula Expresión Excel con signo igual y separadores coma o punto y coma.
 * @param read Resuelve una referencia A1; puede devolver errores de dependencias.
 * @returns Valor calculado o error de hoja de cálculo; conserva las fórmulas desconocidas como #NAME?.
 */
export function evaluateSpreadsheetResult(
  formula: string,
  read: (reference: string) => SheetValue,
  options: SpreadsheetFormulaOptions = {},
): SpreadsheetValue {
  try {
    const source = canonicalSpreadsheetFormula(formula).replace(/^=/, '');
    const tokens = tokenizeSpreadsheetFormula(source);
    let position = 0;
    const take = () => tokens[position++];
    const precedence: Record<string, number> = {
      '=': 1,
      '<>': 1,
      '>': 1,
      '<': 1,
      '>=': 1,
      '<=': 1,
      '&': 2,
      '+': 3,
      '-': 3,
      '*': 4,
      '/': 4,
      '^': 5,
    };
    /** Lee tokens con precedencia; retorna un árbol para permitir SI/SI.ERROR perezosos. */
    const parse = (minimum = 0): Node => {
      const token = take();
      if (!token) throw Error('#ERROR!');
      let node: Node;
      if (token === '+' || token === '-') node = { kind: 'unary', op: token, node: parse(6) };
      else if (token === '(') {
        node = parse();
        if (take() !== ')') throw Error('#ERROR!');
      } else if (token.startsWith('"'))
        node = { kind: 'literal', value: token.slice(1, -1).replace(/""/g, '"') };
      else if (
        /^(?:(?:'(?:[^']|'')+'|[\p{L}_][\p{L}\d_.]*)!)?\$?(?:[A-Z]{1,3}:\$?[A-Z]{1,3}|\d+:\$?\d+)$/iu.test(
          token,
        )
      )
        node = { kind: 'whole', value: token };
      else if (/^\d|^\./.test(token)) node = { kind: 'literal', value: Number(token) };
      else if (spreadsheetAddress(token)) {
        if (tokens[position] === '#') {
          take();
          node = { kind: 'spill', value: token };
        } else if (tokens[position] === ':') {
          take();
          const end = take();
          if (!spreadsheetAddress(end)) throw Error('#REF!');
          node = { kind: 'range', start: token, end };
        } else node = { kind: 'ref', value: token };
      } else if (tokens[position] === '(') {
        take();
        const args: Node[] = [];
        while (tokens[position] !== ')') {
          args.push(tokens[position] === ',' ? { kind: 'literal', value: 0 } : parse());
          if (tokens[position] !== ',' && tokens[position] !== ';') break;
          take();
          if (tokens[position] === ')') args.push({ kind: 'literal', value: 0 });
        }
        if (take() !== ')') throw Error('#ERROR!');
        node = {
          kind: 'call',
          name: token.toUpperCase().replace(/^_XLFN\.(?:_XLWS\.)?|^_XLWS\./, ''),
          args,
        };
      } else if (/^(TRUE|FALSE|VERDADERO|FALSO)$/i.test(token))
        node = { kind: 'literal', value: /^(TRUE|VERDADERO)$/i.test(token) };
      else if (token.startsWith('#')) node = { kind: 'literal', value: token };
      else node = { kind: 'name', value: token };
      while (position < tokens.length) {
        if (tokens[position] === '%') {
          take();
          node = { kind: 'unary', op: '%', node };
          continue;
        }
        const op = tokens[position],
          level = precedence[op];
        if (level === undefined || level < minimum) break;
        take();
        node = { kind: 'binary', op, left: node, right: parse(level + 1) };
      }
      return node;
    };
    const scalar = (v: Value): SheetValue => (Array.isArray(v) ? (v[0]?.[0] ?? null) : v);
    const error = isSpreadsheetError;
    const number = (v: Value): number => {
      const s = scalar(v);
      if (error(s)) throw Error(String(s));
      const n = Number(s ?? 0);
      if (!Number.isFinite(n)) throw Error('#VALUE!');
      return n;
    };
    const truth = (value: SheetValue) => {
      if (error(value)) throw Error(String(value));
      return Boolean(value);
    };
    const flat = (v: Value): SheetValue[] => (Array.isArray(v) ? v.flat() : [v]);
    const matches = spreadsheetMatchesCriteria;
    const compare = compareSpreadsheetValues;
    /** Evalúa el árbol; recibe nodo y retorna escalar/rango con propagación de errores. */
    let subtotalMode: number | null = null;
    /** Recibe referencia; retorna su valor o blanco al excluir filas ocultas/subtotales anidados. */
    const readCell = (reference: string): SheetValue =>
      subtotalMode !== null && options.omitSubtotalCell?.(reference, subtotalMode >= 101)
        ? null
        : read(reference);
    const evaluate = (node: Node): Value => {
      if (node.kind === 'literal') return node.value;
      if (node.kind === 'ref') return readCell(node.value);
      if (node.kind === 'spill') return options.readSpill?.(node.value) ?? '#REF!';
      if (node.kind === 'whole') {
        const split = node.value.lastIndexOf('!'),
          prefix = split < 0 ? '' : node.value.slice(0, split + 1),
          [start, end] = node.value
            .slice(split + 1)
            .replace(/\$/g, '')
            .split(':');
        if (/^\d+$/.test(start)) {
          const a = Number(start),
            b = Number(end);
          if (a < 1 || b < 1 || a > 1048576 || b > 1048576) throw Error('#REF!');
          return evaluate({
            kind: 'range',
            start: prefix + 'A' + Math.min(a, b),
            end: 'XFD' + Math.max(a, b),
          });
        }
        const a = spreadsheetAddress(start + '1'),
          b = spreadsheetAddress(end + '1');
        if (!a || !b) throw Error('#REF!');
        const last = Math.max(1, options.usedRows ?? 1);
        const range = evaluate({
          kind: 'range',
          start: prefix + start + '1',
          end: end + last,
        }) as RangeValue;
        return markSpreadsheetRange(range, 1048576, Math.abs(a.col - b.col) + 1);
      }
      if (node.kind === 'name') {
        const name = node.value.toUpperCase(),
          seen = new Set(options.nameStack ?? []);
        if (seen.has(name)) throw Error('#CYCLE!');
        seen.add(name);
        const expression = (
          node.value.includes('[')
            ? options.resolveStructured?.(node.value)
            : (options.resolveName?.(node.value) ??
              options.resolveStructured?.(node.value + '[#Data]'))
        )?.replace(/^=/, '');
        if (expression === undefined) throw Error('#NAME?');
        const refs = expression.match(
          /^((?:(?:'(?:[^']|'')+'|[\p{L}_][\p{L}\d_.]*)!)?\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?$/iu,
        );
        if (refs)
          return evaluate(
            refs[2]
              ? { kind: 'range', start: refs[1], end: refs[2] }
              : { kind: 'ref', value: refs[1] },
          );
        return evaluateSpreadsheetResult('=' + expression, readCell, {
          ...options,
          nameStack: seen,
        });
      }
      if (node.kind === 'range') {
        const a = spreadsheetAddress(node.start)!,
          b = spreadsheetAddress(node.end)!;
        if (b.sheet && b.sheet !== a.sheet) throw Error('#REF!');
        if ((Math.abs(a.row - b.row) + 1) * (Math.abs(a.col - b.col) + 1) > 100000)
          throw Error('#NUM!');
        const rows: RangeValue = [];
        for (let row = Math.min(a.row, b.row); row <= Math.max(a.row, b.row); row++) {
          const cells: SheetValue[] = [];
          for (let col = Math.min(a.col, b.col); col <= Math.max(a.col, b.col); col++) {
            let label = '',
              c = col + 1;
            while (c) {
              label = String.fromCharCode(65 + ((c - 1) % 26)) + label;
              c = Math.floor((c - 1) / 26);
            }
            cells.push(
              readCell(`${a.sheet ? `'${a.sheet.replace(/'/g, "''")}'!` : ''}${label}${row + 1}`),
            );
          }
          rows.push(cells);
        }
        return rows;
      }
      if (node.kind === 'unary') {
        return mapSpreadsheetArray(evaluate(node.node), 0, (value) => {
          const n = number(value);
          return node.op === '-' ? -n : node.op === '%' ? n / 100 : n;
        });
      }
      if (node.kind === 'binary') {
        return mapSpreadsheetArray(evaluate(node.left), evaluate(node.right), (left, right) => {
          if (error(left)) throw Error(String(left));
          if (error(right)) throw Error(String(right));
          if (node.op === '&') return String(left ?? '') + String(right ?? '');
          if (['=', '<>', '>', '<', '>=', '<='].includes(node.op)) {
            const c = compare(left, right);
            return node.op === '='
              ? c === 0
              : node.op === '<>'
                ? c !== 0
                : node.op === '>'
                  ? c > 0
                  : node.op === '<'
                    ? c < 0
                    : node.op === '>='
                      ? c >= 0
                      : c <= 0;
          }
          const a = number(left),
            b = number(right);
          switch (node.op) {
            case '+':
              return a + b;
            case '-':
              return a - b;
            case '*':
              return a * b;
            case '/':
              if (!b) throw Error('#DIV/0!');
              return a / b;
            case '^':
              return a ** b;
          }
          throw Error('#VALUE!');
        });
      }
      if (node.kind !== 'call') throw Error('#ERROR!');
      if (
        !['FILTER', 'SORT', 'UNIQUE', 'SEQUENCE', 'TRANSPOSE', 'ANCHORARRAY'].includes(node.name) &&
        !SUPPORTED_SPREADSHEET_FUNCTIONS.has(node.name)
      )
        throw Error('#NAME?');
      const args = node.args;
      if (node.name === 'ANCHORARRAY') {
        if (args[0]?.kind !== 'ref') throw Error('#REF!');
        return options.readSpill?.(args[0].value) ?? '#REF!';
      }
      if (node.name === 'FILTER') {
        const array = evaluate(args[0]),
          mask = evaluate(args[1]);
        try {
          return evaluateSpreadsheetArray('FILTER', [array, mask], compare)!;
        } catch (error) {
          if (error instanceof Error && error.message === '#CALC!' && args[2])
            return evaluate(args[2]);
          throw error;
        }
      }
      if (node.name === 'IF') {
        const condition = evaluate(args[0]);
        if (!Array.isArray(condition))
          return evaluate(
            truth(condition) ? args[1] : (args[2] ?? { kind: 'literal', value: false }),
          );
        const mask = materializeSpreadsheetArray(condition),
          needsTrue = mask.some((r) => r.some((v) => !error(v) && Boolean(v))),
          needsFalse = mask.some((r) => r.some((v) => !error(v) && !Boolean(v)));
        const branch = (node: Node): Value => {
          try {
            return evaluate(node);
          } catch (error) {
            return error instanceof Error ? error.message : '#VALUE!';
          }
        };
        const yes = needsTrue ? branch(args[1]) : null,
          no = needsFalse ? branch(args[2] ?? { kind: 'literal', value: false }) : null;
        const shapes = [spreadsheetShape(mask), spreadsheetShape(yes), spreadsheetShape(no)],
          rows = Math.max(...shapes.map((s) => s.rows)),
          cols = Math.max(...shapes.map((s) => s.cols));
        if (
          rows * cols > 100000 ||
          shapes.some((s) => (s.rows !== 1 && s.rows !== rows) || (s.cols !== 1 && s.cols !== cols))
        )
          throw Error('#VALUE!');
        const at = (value: Value, index: number, r: number, c: number) =>
          spreadsheetAt(value, shapes[index].rows === 1 ? 0 : r, shapes[index].cols === 1 ? 0 : c);
        return Array.from({ length: rows }, (_, r) =>
          Array.from({ length: cols }, (_, c) => {
            const value = at(mask, 0, r, c);
            return error(value) ? value : truth(value) ? at(yes, 1, r, c) : at(no, 2, r, c);
          }),
        );
      }
      if (node.name === 'IFERROR') {
        try {
          const v = evaluate(args[0]);
          if (Array.isArray(v)) {
            if (!v.some((row) => row.some(error))) return v;
            let fallback: Value;
            try {
              fallback = evaluate(args[1]);
            } catch (e) {
              fallback = e instanceof Error ? e.message : '#VALUE!';
            }
            return mapSpreadsheetArray(v, fallback, (value, fallback) =>
              error(value) ? fallback : value,
            );
          }
          if (error(scalar(v))) throw Error();
          return v;
        } catch {
          return evaluate(args[1]);
        }
      }
      if (node.name === 'IFNA') {
        try {
          const result = evaluate(args[0]);
          return scalar(result) === '#N/A' ? evaluate(args[1]) : result;
        } catch (e) {
          if (e instanceof Error && e.message === '#N/A') return evaluate(args[1]);
          throw e;
        }
      }
      if (node.name === 'IFS') {
        if (args.length % 2) throw Error('#VALUE!');
        for (let i = 0; i < args.length; i += 2) {
          const v = scalar(evaluate(args[i]));
          if (error(v)) throw Error(String(v));
          if (Boolean(v)) return evaluate(args[i + 1]);
        }
        throw Error('#N/A');
      }
      if (node.name === 'SWITCH') {
        const value = scalar(evaluate(args[0]));
        if (error(value)) throw Error(String(value));
        for (let i = 1; i + 1 < args.length; i += 2)
          if (compare(value, scalar(evaluate(args[i]))) === 0) return evaluate(args[i + 1]);
        if (args.length % 2 === 0) return evaluate(args.at(-1)!);
        throw Error('#N/A');
      }
      if (node.name === 'SUBTOTAL') {
        const code = number(evaluate(args[0])),
          previous = subtotalMode;
        subtotalMode = code;
        try {
          return evaluateSpreadsheetLibrary('SUBTOTAL', [code, ...args.slice(1).map(evaluate)], {
            ...options,
            matches,
            compare,
          })!;
        } finally {
          subtotalMode = previous;
        }
      }
      if (node.name === 'ROW' || node.name === 'COLUMN') {
        const first = args[0];
        if (!first) return (node.name === 'ROW' ? (options.row ?? 0) : (options.col ?? 0)) + 1;
        const address =
          first.kind === 'ref'
            ? spreadsheetAddress(first.value)
            : first.kind === 'range'
              ? spreadsheetAddress(first.start)
              : null;
        if (!address) throw Error('#VALUE!');
        return (node.name === 'ROW' ? address.row : address.col) + 1;
      }
      if (node.name === 'TEXTBEFORE' || node.name === 'TEXTAFTER') {
        if (args.length < 2 || args.length > 6) throw Error('#VALUE!');
        try {
          return evaluateSpreadsheetLibrary(node.name, args.slice(0, 5).map(evaluate), {
            ...options,
            matches,
            compare,
          })!;
        } catch (error) {
          if (error instanceof Error && error.message === '#N/A' && args[5])
            return evaluate(args[5]);
          throw error;
        }
      }
      if (node.name === 'XLOOKUP') {
        const values = [
          evaluate(args[0]),
          evaluate(args[1]),
          evaluate(args[2]),
          null,
          args[4] ? evaluate(args[4]) : 0,
          args[5] ? evaluate(args[5]) : 1,
        ];
        try {
          return evaluateSpreadsheetLibrary('XLOOKUP', values, { ...options, matches, compare })!;
        } catch (e) {
          if (e instanceof Error && e.message === '#N/A' && args[3]) return evaluate(args[3]);
          throw e;
        }
      }
      // Las funciones de inspección reciben el error como valor; no dejan que una división lo interrumpa antes.
      const values = args.map((arg) => {
          try {
            return evaluate(arg);
          } catch (e) {
            if (
              [
                'ISERROR',
                'ISERR',
                'ISNA',
                'ISNUMBER',
                'ISTEXT',
                'ISBLANK',
                'ISLOGICAL',
                'COUNT',
              ].includes(node.name) &&
              e instanceof Error &&
              e.message.startsWith('#')
            )
              return e.message;
            throw e;
          }
        }),
        all = values.flatMap(flat),
        nums = all.filter((v) => typeof v === 'number') as number[];
      const dynamic = evaluateSpreadsheetArray(node.name, values, compare);
      if (dynamic !== undefined) return dynamic;
      const extended = evaluateSpreadsheetLibrary(node.name, values, {
        ...options,
        matches,
        compare,
      });
      if (extended !== undefined) return extended;
      const failure = node.name === 'COUNT' || node.name === 'COUNTA' ? undefined : all.find(error);
      if (failure) throw Error(String(failure));
      const n = (i: number, defaultValue = 0) =>
        values[i] === undefined ? defaultValue : number(values[i]);
      const text = (i: number) => String(scalar(values[i]) ?? '');
      switch (node.name) {
        case 'SUM':
          return nums.reduce((a, b) => a + b, 0);
        case 'AVERAGE':
          if (!nums.length) throw Error('#DIV/0!');
          return nums.reduce((a, b) => a + b, 0) / nums.length;
        case 'COUNT':
          return nums.length;
        case 'COUNTA':
          return all.filter((v) => v !== null && v !== '').length;
        case 'MIN':
          return nums.length ? Math.min(...nums) : 0;
        case 'MAX':
          return nums.length ? Math.max(...nums) : 0;
        case 'ABS':
          return Math.abs(n(0));
        case 'INT':
          return Math.floor(n(0));
        case 'SQRT':
          return Math.sqrt(n(0));
        case 'POWER':
          return n(0) ** n(1);
        case 'MOD':
          if (!n(1)) throw Error('#DIV/0!');
          return ((n(0) % n(1)) + n(1)) % n(1);
        case 'ROUND':
          return (
            (Math.sign(n(0)) * Math.round(Math.abs(n(0)) * 10 ** n(1) + Number.EPSILON)) /
            10 ** n(1)
          );
        case 'ROUNDUP':
          return (Math.sign(n(0)) * Math.ceil(Math.abs(n(0)) * 10 ** n(1))) / 10 ** n(1);
        case 'ROUNDDOWN':
          return Math.trunc(n(0) * 10 ** n(1)) / 10 ** n(1);
        case 'AND':
          return all.every(Boolean);
        case 'OR':
          return all.some(Boolean);
        case 'NOT':
          return !Boolean(scalar(values[0]));
        case 'CONCAT':
        case 'CONCATENATE':
          return all.map((v) => v ?? '').join('');
        case 'LEN':
          return text(0).length;
        case 'LEFT':
          if (n(1, 1) < 0) throw Error('#VALUE!');
          return text(0).slice(0, Math.trunc(n(1, 1)));
        case 'RIGHT':
          if (n(1, 1) < 0) throw Error('#VALUE!');
          return n(1, 1) ? text(0).slice(-n(1, 1)) : '';
        case 'MID':
          if (n(1) < 1 || n(2) < 0) throw Error('#VALUE!');
          return text(0).slice(n(1) - 1, n(1) - 1 + n(2));
        case 'TRIM':
          return text(0).trim().replace(/ +/g, ' ');
        case 'UPPER':
          return text(0).toUpperCase();
        case 'LOWER':
          return text(0).toLowerCase();
        case 'COUNTIF':
          return (
            flat(values[0]).filter((v) => matches(v, scalar(values[1]))).length +
            (matches(null, scalar(values[1])) ? spreadsheetTail(values[0]) : 0)
          );
        case 'SUMIF':
          return flat(values[0]).reduce<number>(
            (sum, v, i) =>
              sum +
              (matches(v, scalar(values[1])) ? Number(flat(values[2] ?? values[0])[i]) || 0 : 0),
            0,
          );
        case 'INDEX': {
          const shape = spreadsheetShape(values[0]),
            row = Math.trunc(n(1)),
            col = Math.trunc(n(2, 1));
          if (row < 0 || col < 0 || row > shape.rows || col > shape.cols) throw Error('#REF!');
          if (row && col) return spreadsheetAt(values[0], row - 1, col - 1) ?? 0;
          const table = materializeSpreadsheetArray(values[0]);
          return row ? [table[row - 1]] : col ? table.map((r) => [r[col - 1]]) : table;
        }
        case 'MATCH': {
          const array = flat(values[1]);
          const mode = n(2, 1);
          if (![-1, 0, 1].includes(mode)) throw Error('#N/A');
          const i = array.findIndex((v) =>
            mode === 0 ? matches(v, scalar(values[0])) : compare(v, scalar(values[0])) === 0,
          );
          if (i >= 0) return i + 1;
          let best = -1;
          if (mode !== 0)
            array.forEach((v, index) => {
              if (
                mode === 1 ? compare(v, scalar(values[0])) <= 0 : compare(v, scalar(values[0])) >= 0
              )
                best = index;
            });
          if (best >= 0) return best + 1;
          throw Error('#N/A');
        }
        case 'VLOOKUP':
        case 'HLOOKUP': {
          let table = Array.isArray(values[1]) ? values[1] : [[scalar(values[1])]];
          if (node.name === 'HLOOKUP')
            table = table[0].map((_, i) => table.map((row) => row[i] ?? null));
          const exact = values[3] !== undefined && !Boolean(scalar(values[3]));
          const row =
            table.find((row) =>
              exact ? matches(row[0], scalar(values[0])) : compare(row[0], scalar(values[0])) === 0,
            ) ??
            (!exact
              ? table.filter((row) => compare(row[0], scalar(values[0])) <= 0).at(-1)
              : undefined);
          const column = Math.trunc(n(2));
          if (column < 1) throw Error('#VALUE!');
          if (column > (table[0]?.length ?? 0)) throw Error('#REF!');
          return row ? (row[column - 1] ?? 0) : '#N/A';
        }
        default:
          throw Error('#NAME?');
      }
    };
    const tree = parse();
    if (position !== tokens.length) return '#ERROR!';
    const result = evaluate(tree);
    if (Array.isArray(result))
      return materializeSpreadsheetArray(result).map((row) =>
        row.map((v) => (typeof v === 'number' && !Number.isFinite(v) ? '#NUM!' : v)),
      );
    return typeof result === 'number' && !Number.isFinite(result) ? '#NUM!' : result;
  } catch (error) {
    return error instanceof Error && error.message.startsWith('#') ? error.message : '#ERROR!';
  }
}

/** Recibe expresión, lector y contexto; retorna el primer resultado escalar para consumidores que no administran desbordamientos. */
export function evaluateSpreadsheetFormula(
  formula: string,
  read: (reference: string) => SheetValue,
  options: SpreadsheetFormulaOptions = {},
): SheetValue {
  const result = evaluateSpreadsheetResult(formula, read, options);
  return Array.isArray(result) ? (result[0]?.[0] ?? null) : result;
}

/** Recibe fórmula y desplazamiento de pegado; retorna referencias relativas movidas, respetando $ y cadenas. */
export function translateSpreadsheetFormula(formula: string, rows: number, cols: number): string {
  try {
    return tokenizeSpreadsheetFormula(formula)
      .map((token) => {
        if (token.startsWith('"') || token.includes('[')) return token;
        const bang = token.lastIndexOf('!'),
          prefix = bang < 0 ? '' : token.slice(0, bang + 1),
          local = token.slice(bang + 1);
        const address = spreadsheetAddress(token);
        if (address) {
          const flags = local.match(/^(\$?)[A-Z]+(\$?)\d+$/i)!,
            r = address.row + (flags[2] ? 0 : rows),
            c = address.col + (flags[1] ? 0 : cols);
          if (r < 0 || c < 0 || r >= 1048576 || c >= 16384) return '#REF!';
          return prefix + flags[1] + formulaColumnLabel(c) + flags[2] + (r + 1);
        }
        if (/^\$?(?:[A-Z]{1,3}:\$?[A-Z]{1,3}|\d+:\$?\d+)$/i.test(local)) {
          const mapped = local.split(':').map((part) => {
            const absolute = part.startsWith('$'),
              v = part.replace('$', '');
            if (/^\d+$/.test(v)) {
              const r = Number(v) + (absolute ? 0 : rows);
              return r < 1 || r > 1048576 ? '#REF!' : (absolute ? '$' : '') + r;
            }
            const a = spreadsheetAddress(v + '1')!;
            const c = a.col + (absolute ? 0 : cols);
            return c < 0 || c >= 16384 ? '#REF!' : (absolute ? '$' : '') + formulaColumnLabel(c);
          });
          return mapped.includes('#REF!') ? '#REF!' : prefix + mapped.join(':');
        }
        return token;
      })
      .join('');
  } catch {
    return formula;
  }
}

/**
 * Reescribe referencias tras mover/insertar/eliminar filas o columnas en una hoja.
 * @param formula Expresión original, con literales de texto preservados.
 * @param ownSheet Nombre de la hoja que contiene la fórmula.
 * @param changedSheet Nombre de la hoja modificada.
 * @param map Traduce una coordenada antigua o retorna null si se eliminó.
 * @returns Fórmula con referencias absolutas y calificadas actualizadas sólo donde corresponde.
 */
export function remapSpreadsheetFormula(
  formula: string,
  ownSheet: string,
  changedSheet: string,
  map: (row: number, col: number) => { row: number; col: number } | null,
): string {
  try {
    const tokens = tokenizeSpreadsheetFormula(formula),
      output: string[] = [];
    const move = (reference: string, inherited?: string): string => {
      const address = spreadsheetAddress(reference);
      if (!address) return reference;
      if ((address.sheet ?? inherited ?? ownSheet).toLowerCase() !== changedSheet.toLowerCase())
        return reference;
      const mapped = map(address.row, address.col);
      if (!mapped) return '#REF!';
      const bang = reference.lastIndexOf('!'),
        prefix = bang < 0 ? '' : reference.slice(0, bang + 1),
        local = reference.slice(bang + 1),
        flags = local.match(/^(\$?)[A-Z]+(\$?)\d+$/i)!;
      return prefix + flags[1] + formulaColumnLabel(mapped.col) + flags[2] + (mapped.row + 1);
    };
    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i],
        address = spreadsheetAddress(token);
      if (address) {
        const first = move(token);
        if (tokens[i + 1] === ':' && spreadsheetAddress(tokens[i + 2])) {
          const lastToken = tokens[i + 2],
            end = spreadsheetAddress(lastToken)!;
          let last = move(lastToken, address.sheet),
            start = first;
          // Una eliminación parcial contrae el rango a sus celdas supervivientes, incluso con extremos absolutos.
          if (
            (address.sheet ?? ownSheet).toLowerCase() === changedSheet.toLowerCase() &&
            (first === '#REF!' || last === '#REF!')
          ) {
            const r1 = Math.min(address.row, end.row),
              r2 = Math.max(address.row, end.row),
              c1 = Math.min(address.col, end.col),
              c2 = Math.max(address.col, end.col);
            let row = r1;
            while (row <= r2 && !map(row, c1) && !map(row, c2)) row++;
            if (row <= r2) {
              let left = c1,
                right = c2,
                top = r1,
                bottom = r2;
              while (left <= right && !map(row, left)) left++;
              while (right >= left && !map(row, right)) right--;
              while (top <= bottom && !map(top, left)) top++;
              while (bottom >= top && !map(bottom, left)) bottom--;
              const a = map(top, left),
                b = map(bottom, right);
              /** Recibe token original y coordenada final; retorna referencia preservando calificación y signos $. */
              const format = (token: string, p: { row: number; col: number }): string => {
                const bang = token.lastIndexOf('!'),
                  prefix = bang < 0 ? '' : token.slice(0, bang + 1),
                  flags = token.slice(bang + 1).match(/^(\$?)[A-Z]+(\$?)\d+$/i)!;
                return prefix + flags[1] + formulaColumnLabel(p.col) + flags[2] + (p.row + 1);
              };
              if (a && b) {
                start = format(token, a);
                last = format(lastToken, b);
              }
            }
          }
          output.push(start === '#REF!' || last === '#REF!' ? '#REF!' : start + ':' + last);
          i += 2;
        } else output.push(first);
        continue;
      }
      const bang = token.lastIndexOf('!'),
        prefix = bang < 0 ? '' : token.slice(0, bang + 1),
        local = token.slice(bang + 1);
      if (!token.includes('[') && /^\$?(?:[A-Z]{1,3}:\$?[A-Z]{1,3}|\d+:\$?\d+)$/i.test(local)) {
        const sheet =
          bang < 0
            ? ownSheet
            : prefix.slice(0, -1).startsWith("'")
              ? prefix.slice(1, -2).replace(/''/g, "'")
              : prefix.slice(0, -1);
        if (sheet.toLowerCase() === changedSheet.toLowerCase()) {
          const raw = local.split(':'),
            isRow = /^\$?\d+$/.test(raw[0]);
          const indices = raw.map((v) =>
            isRow
              ? Number(v.replace('$', '')) - 1
              : spreadsheetAddress(v.replace('$', '') + '1')!.col,
          );
          let first = Math.min(...indices),
            last = Math.max(...indices);
          // Coordenadas centinela fuera del eje editado evitan que borrar la fila 1 invalide columnas completas, o viceversa.
          const mapAxis = (index: number) => (isRow ? map(index, 16384) : map(1048576, index));
          while (first <= last && !mapAxis(first)) first++;
          while (last >= first && !mapAxis(last)) last--;
          if (first > last) output.push('#REF!');
          else {
            const a = mapAxis(first)!,
              b = mapAxis(last)!;
            const format = (p: { row: number; col: number }, i: number) =>
              (raw[i].startsWith('$') ? '$' : '') +
              (isRow ? String(p.row + 1) : formulaColumnLabel(p.col));
            output.push(prefix + format(a, 0) + ':' + format(b, 1));
          }
          continue;
        }
      }
      output.push(token);
    }
    return output.join('');
  } catch {
    return formula;
  }
}

/** Recibe columna base cero; retorna etiqueta A1 para traducciones de referencias sin alterar nombres de tablas. */
function formulaColumnLabel(col: number): string {
  let label = '';
  for (col++; col; col = Math.floor((col - 1) / 26))
    label = String.fromCharCode(65 + ((col - 1) % 26)) + label;
  return label;
}

/** Recibe fórmula y nombres de tabla; retorna referencias renombradas sin tocar cadenas, encabezados ni nombres de hoja. null convierte referencias a una tabla eliminada en #REF!. */
export function renameSpreadsheetTableReferences(
  formula: string,
  previous: string,
  next: string | null,
): string {
  try {
    return tokenizeSpreadsheetFormula(formula)
      .map((token) => {
        if (token.toLowerCase() === previous.toLowerCase()) return next ?? '#REF!';
        const bracket = token.indexOf('[');
        return bracket > 0 && token.slice(0, bracket).toLowerCase() === previous.toLowerCase()
          ? next === null
            ? '#REF!'
            : next + token.slice(bracket)
          : token;
      })
      .join('');
  } catch {
    return formula;
  }
}

/** Recibe fórmula, tabla, nombres de columnas anteriores/finales y alcance implícito; retorna referencias actualizadas simultáneamente, conservando literales. */
export function renameSpreadsheetColumnReferences(
  formula: string,
  table: string,
  names: Record<string, string>,
  implicit = false,
): string {
  const lookup = new Map(Object.entries(names).map(([a, b]) => [a.toLowerCase(), b]));
  try {
    return tokenizeSpreadsheetFormula(formula)
      .map((token) => {
        const bracket = token.indexOf('[');
        if (
          bracket < 0 ||
          (bracket === 0
            ? !implicit
            : token.slice(0, bracket).toLowerCase() !== table.toLowerCase())
        )
          return token;
        const shorthand = token.match(/^(.*\[\[#[^\]]+\],)([^\[\]]+)\]$/);
        if (shorthand) {
          const next = lookup.get(shorthand[2].replace(/'([\[\]#@'])/g, '$1').toLowerCase());
          return next === undefined
            ? token
            : shorthand[1] + next.replace(/[\[\]#@']/g, "'$&") + ']';
        }
        return token.replace(/\[((?:'[^]|[^\[\]])*)\]/g, (match, member: string) => {
          if (member.startsWith('#')) return match;
          const row = member.startsWith('@'),
            body = (row ? member.slice(1) : member).replace(/'([\[\]#@'])/g, '$1'),
            next = lookup.get(body.toLowerCase());
          return next === undefined
            ? match
            : '[' + (row ? '@' : '') + next.replace(/[\[\]#@']/g, "'$&") + ']';
        });
      })
      .join('');
  } catch {
    return formula;
  }
}
