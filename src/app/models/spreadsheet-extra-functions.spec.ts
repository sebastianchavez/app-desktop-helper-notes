import {
  evaluateSpreadsheetFormula as evaluate,
  ooxmlSpreadsheetFormula,
} from './spreadsheet-formula';
import {
  recalculateSpreadsheetWorkbook,
  SpreadsheetCalculationSheet,
} from './spreadsheet-workbook';

describe('financial, subtotal and text compatibility', () => {
  const calc = (f: string) => evaluate(f, () => null);
  it('calculates annuities with end/beginning payments and zero interest', () => {
    expect(calc('PAGO(0.01;12;1000)')).toBeCloseTo(-88.8487886783417, 10);
    expect(calc('PMT(0,10,1000)')).toBe(-100);
    expect(calc('FV(0,10,-100,-1000)')).toBe(2000);
    expect(calc('PV(0,10,-100)')).toBe(1000);
    expect(calc('NPER(0,-100,1000)')).toBe(10);
    expect(calc('NPER(0.01,-88.8487886783417,1000)')).toBeCloseTo(12, 9);
    expect(calc('PV(0.01,12,PMT(0.01,12,1000))')).toBeCloseTo(1000, 9);
    expect(calc('PMT(0.01,12,1000,0,1)*(1.01)')).toBeCloseTo(
      calc('PMT(0.01,12,1000)') as number,
      10,
    );
    expect(calc('IPMT(0.01,1,12,1000)')).toBe(-10);
    expect(calc('IPMT(0.01,1,12,1000,0,1)')).toBe(0);
    expect(calc('IPMT(0.01,2,12,1000)+PPMT(0.01,2,12,1000)')).toBeCloseTo(
      calc('PMT(0.01,12,1000)') as number,
      10,
    );
    expect(calc('NPV(0.1,100,200)')).toBeCloseTo(100 / 1.1 + 200 / 1.21, 9);
    expect(calc('PMT(0,0,100)')).toBe('#DIV/0!');
    expect(calc('IPMT(0.01,13,12,1000)')).toBe('#NUM!');
  });
  it('supports numeric functions and explicit locale conversion', () => {
    expect(calc('MROUND(13,5)')).toBe(15);
    expect(calc('MROUND(-13,-5)')).toBe(-15);
    expect(calc('MROUND(-13,5)')).toBe('#NUM!');
    expect(calc('LOG(8,2)')).toBe(3);
    expect(calc('LN(EXP(2))')).toBeCloseTo(2, 12);
    expect(calc('FACT(5.9)')).toBe(120);
    expect(calc('GCD(24,36)')).toBe(12);
    expect(calc('LCM(4,6)')).toBe(12);
    expect(calc('VALOR.NUMERO("1.234,50";",";".")')).toBe(1234.5);
    expect(calc('NUMBERVALUE("12.5%")')).toBe(0.125);
    expect(calc('NUMBERVALUE("1.2.3")')).toBe('#VALUE!');
  });
  it('extracts delimited text with reverse instances, case modes and fallbacks', () => {
    expect(calc('TEXTBEFORE("a/b/c","/",-1)')).toBe('a/b');
    expect(calc('TEXTAFTER("a/b/c","/",2)')).toBe('c');
    expect(calc('TEXTOANTES("aXXb";"xx";1;1)')).toBe('a');
    expect(calc('TEXTAFTER("abc","/",1,0,0,"none")')).toBe('none');
    expect(calc('TEXTBEFORE("abc","/",1,0,1)')).toBe('abc');
    expect(calc('TEXTAFTER("abc","/",1,0,1)')).toBe('');
    expect(calc('TEXTAFTER("abc","/",0)')).toBe('#VALUE!');
    expect(calc('TEXTBEFORE("a/b","/",1,0,0,1/0)')).toBe('a');
    expect(ooxmlSpreadsheetFormula('TEXTOANTES("A/B";"/")')).toBe('_xlfn.TEXTBEFORE("A/B","/")');
  });
  it('recalculates table totals and ignores nested subtotals and hidden rows', () => {
    const sheets: SpreadsheetCalculationSheet[] = [
      { name: 'Data', colCount: 3, data: [[10], [20], [30], [], []], hiddenRows: [1] },
    ];
    const formulas = new Map([
      ['0:3:0', 'SUBTOTAL(9,A1:A3)'],
      ['0:4:0', 'SUBTOTAL(109,A1:A4)'],
      ['0:0:1', 'SUBTOTAL(2,A1:A4)'],
      ['0:1:1', 'SUBTOTAL(3,A1:A4)'],
    ]);
    recalculateSpreadsheetWorkbook(sheets, formulas, new Map(), new Map());
    expect(sheets[0].data[3][0]).toBe(60);
    expect(sheets[0].data[4][0]).toBe(40);
    expect(sheets[0].data[0][1]).toBe(3);
    expect(sheets[0].data[1][1]).toBe(3);
    expect(calc('SUBTOTALES(109;1;2;3)')).toBe(6);
    expect(calc('SUBTOTAL(112,1)')).toBe('#VALUE!');
  });
  it('preserves an imported subtotal cache when active filters make manual/filtered visibility ambiguous', () => {
    const sheets: SpreadsheetCalculationSheet[] = [
      {
        name: 'Data',
        colCount: 2,
        data: [[10], [20], [30], [], []],
        hiddenRows: [1],
        hasActiveFilters: true,
      },
    ];
    const formulas = new Map([
      ['0:3:0', 'SUBTOTAL(9,A1:A3)'],
      ['0:4:0', 'SUBTOTAL(109,A1:A3)'],
    ]);
    const result = recalculateSpreadsheetWorkbook(
      sheets,
      formulas,
      new Map([['0:3:0', { formula: 'SUBTOTAL(9,A1:A3)', value: 77 }]]),
      new Map(),
    );
    expect(sheets[0].data[3][0]).toBe(77);
    expect(sheets[0].data[4][0]).toBe(40);
    expect(result.cachedCount).toBe(1);
  });
});
