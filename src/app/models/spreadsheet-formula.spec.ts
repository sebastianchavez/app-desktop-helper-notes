import {
  canonicalSpreadsheetFormula,
  evaluateSpreadsheetFormula,
  translateSpreadsheetFormula,
  remapSpreadsheetFormula,
  ooxmlSpreadsheetFormula,
} from './spreadsheet-formula';

describe('spreadsheet formula engine', () => {
  const cells: Record<string, string | number | boolean | null> = {
    A1: 10,
    A2: 20,
    B1: 'Ana',
    B2: 'Luis',
    C1: 3,
    C2: 4,
    "'Resumen'!A1": 100,
  };
  const run = (formula: string) => evaluateSpreadsheetFormula(formula, (ref) => cells[ref] ?? null);
  it('supports nested functions inside arithmetic, quoted arguments, semicolons and qualified references', () => {
    expect(run('=SUMA(A1:A2)*2+SI(A1>5;3;0)')).toBe(63);
    expect(run('=SI(A1>5;CONCATENAR("hola; ";B1);"no")')).toBe('hola; Ana');
    expect(run("='Resumen'!A1+SUM(A1:A2)")).toBe(130);
  });
  it('propagates errors and evaluates only the required IF / IFERROR branches', () => {
    expect(run('=IF(A1=10;5;1/0)')).toBe(5);
    expect(run('=IFERROR(10/0;99)')).toBe(99);
    expect(run('=SUM(A1:A2)/0')).toBe('#DIV/0!');
    expect(run('=UNKNOWN(A1)')).toBe('#NAME?');
    expect(run('=UNKNOWN(1/0)')).toBe('#NAME?');
  });
  it('calculates criteria, lookups, text, booleans and percentage arithmetic', () => {
    expect(run('=SUMIF(A1:A2;">10";C1:C2)')).toBe(4);
    expect(run('=COUNTIF(B1:B2;"L*")')).toBe(1);
    expect(run('=VLOOKUP(20;A1:C2;2;FALSE)')).toBe('Luis');
    expect(run('=ROUND(200*10%;2)')).toBe(20);
    expect(run('=AND(A1>5;A2<30)')).toBe(true);
  });
  it('keeps absolute references and string literals during copy/paste', () => {
    expect(translateSpreadsheetFormula('=A1+$B$2+C$3+$D4+"A1"', 2, 1)).toBe(
      '=B3+$B$2+D$3+$D6+"A1"',
    );
    expect(canonicalSpreadsheetFormula('=SI(A1;"SUMA(A1)";SUMA(A1:A2))')).toBe(
      '=IF(A1,"SUMA(A1)",SUM(A1:A2))',
    );
    expect(translateSpreadsheetFormula('=A1', -1, 0)).toBe('=#REF!');
  });
  it('updates only the affected sheet, including absolute references and qualified ranges', () => {
    const move = (row: number, col: number) => ({ row: row + 1, col });
    expect(remapSpreadsheetFormula('=SUM(Datos!A1:A3)+A1+"A1"', 'Resumen', 'Datos', move)).toBe(
      '=SUM(Datos!A2:A4)+A1+"A1"',
    );
    expect(remapSpreadsheetFormula('=$A$1+Otra!A1', 'Datos', 'Datos', move)).toBe('=$A$2+Otra!A1');
  });
  const examples: [string, unknown][] = [
    ['=SUMIFS(A1:A2;B1:B2;"A*";C1:C2;">2")', 10],
    ['=COUNTIFS(A1:A2;">5";C1:C2;">3")', 1],
    ['=AVERAGEIFS(A1:A2;C1:C2;">2")', 15],
    ['=MINIFS(A1:A2;C1:C2;">3")', 20],
    ['=MAXIFS(A1:A2;C1:C2;">3")', 20],
    ['=AVERAGEIF(C1:C2;">2";A1:A2)', 15],
    ['=SUMPRODUCT(A1:A2;C1:C2)', 110],
    ['=XLOOKUP("luis";B1:B2;A1:A2)', 20],
    ['=BUSCARX("Missing";B1:B2;A1:A2;"No encontrado")', 'No encontrado'],
    ['=_xlfn.XLOOKUP(15;A1:A2;C1:C2;"";-1)', 3],
    ['=XLOOKUP(15;A1:A2;C1:C2;"";1)', 4],
    ['=XLOOKUP("L*";B1:B2;A1:A2;"";2)', 20],
    ['=XLOOKUP(10;A1:A2;C1:C2;1/0)', 3],
    ['=XMATCH(20;A1:A2)', 2],
    ['=MATCH(15;A1:A2)', 1],
    ['=MATCH(15;A1:A2;0)', '#N/A'],
    ['=IFNA(MATCH(99;A1:A2;0);123)', 123],
    ['=IFNA(1/0;123)', '#DIV/0!'],
    ['=IFS(FALSE;1/0;A1>5;"yes";TRUE;"no")', 'yes'],
    ['=SWITCH(A1;10;"ten";20;1/0;"other")', 'ten'],
    ['=ISERROR(1/0)', true],
    ['=ISNA(NA())', true],
    ['=ISERR(NA())', false],
    ['=ISNUMBER(A1)', true],
    ['=ISTEXT(B1)', true],
    ['=ISBLANK(D1)', true],
    ['=ISBLANK("")', false],
    ['=COUNTBLANK(A1:D2)', 2],
    ['=COUNT(A1:A2;NA())', 2],
    ['=MEDIAN(A1:A2;100)', 20],
    ['=LARGE(A1:C2;2)', 10],
    ['=SMALL(A1:C2;2)', 4],
    ['=VAR.S(A1:A2)', 50],
    ['=STDEV.P(A1:A2)', 5],
    ['=PRODUCT(A1:A2)', 200],
    ['=SIGN(-4)', -1],
    ['=TRUNC(-1.234;2)', -1.23],
    ['=QUOTIENT(-7;3)', -2],
    ['=CEILING.MATH(-4.3;2)', -4],
    ['=FLOOR.MATH(-4.3;2)', -6],
    ['=TEXTJOIN("-";TRUE;B1:B2;"")', 'Ana-Luis'],
    ['=SUBSTITUTE("a-b-a";"a";"x";2)', 'a-b-x'],
    ['=REPLACE("abcdef";2;3;"X")', 'aXef'],
    ['=FIND("b";"abc")', 2],
    ['=SEARCH("B?";"aBc")', 2],
    ['=REPT("ab";3)', 'ababab'],
    ['=PROPER("hOLA MUNDO")', 'Hola Mundo'],
    ['=EXACT("A";"a")', false],
    ['=ROW(C2)', 2],
    ['=COLUMN(C2)', 3],
    ['=ROWS(A1:C2)', 2],
    ['=COLUMNS(A1:C2)', 3],
    ['=COUNTIF(B1:B2;"a*")', 1],
    ['="Ana"="a*"', false],
    ['="ANA"="ana"', true],
    ['=COUNTIF(B1:B2;"~*")', 0],
    ['=IF(TRUE;;99)', 0],
    ['=IF(FALSE;99;)', 0],
    ['=-2^2', 4],
    ['=2^3^2', 64],
    ['=IF(1/0;1;2)', '#DIV/0!'],
    ['=LEFT("abc";-1)', '#VALUE!'],
    ['=MID("abc";0;1)', '#VALUE!'],
  ];
  for (const [formula, expected] of examples)
    it(`calculates ${formula}`, () => expect(run(formula)).toEqual(expected));
  it('honors modern OOXML prefixes and protects quoted sheet names and copy boundaries', () => {
    expect(
      canonicalSpreadsheetFormula('=\'Sum(a);A1\'!A1+sumar.si.conjunto(A1:A2;B1:B2;"x;")'),
    ).toBe('=\'Sum(a);A1\'!A1+SUMIFS(A1:A2,B1:B2,"x;")');
    expect(ooxmlSpreadsheetFormula('=BUSCARX(A1;B1:B2;C1:C2)+SUMA(A1:A2)')).toBe(
      '=_xlfn.XLOOKUP(A1,B1:B2,C1:C2)+SUM(A1:A2)',
    );
    expect(translateSpreadsheetFormula("='A1'!A1+Sheet1!B2+TaxA1", 1, 1)).toBe(
      "='A1'!B2+Sheet1!C3+TaxA1",
    );
  });
  it('resolves named ranges, constants and aliases and rejects recursive definitions', () => {
    const names: Record<string, string> = { SALES: '$A$1:$A$2', RATE: '0.2', LOOP: 'LOOP' };
    const named = (formula: string) =>
      evaluateSpreadsheetFormula(formula, (ref) => cells[ref] ?? null, {
        resolveName: (name) => names[name.toUpperCase()],
      });
    expect(named('=SUM(Sales)*Rate')).toBe(6);
    expect(named('=Loop')).toBe('#CYCLE!');
    expect(named('=Missing')).toBe('#NAME?');
  });
  it('validates criteria shapes, invalid lookup columns and literal wildcard escapes', () => {
    expect(run('=COUNTIFS(A1:A2;">0";A1:C1;">0")')).toBe('#VALUE!');
    const lookup = (f: string) =>
      evaluateSpreadsheetFormula(
        f,
        (ref) => ({ A1: '*', A2: '?', A3: '~', A4: 'abc' })[ref] ?? null,
      );
    expect(lookup('=COUNTIF(A1:A4;"~*")')).toBe(1);
    expect(lookup('=COUNTIF(A1:A4;"~?")')).toBe(1);
    expect(lookup('=COUNTIF(A1:A4;"~~")')).toBe(1);
  });
});
