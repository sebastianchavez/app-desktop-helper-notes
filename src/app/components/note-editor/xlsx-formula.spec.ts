import { evaluateArithmeticExpression, nextSpreadsheetCell } from './note-editor';

describe('xlsx arithmetic formulas', () => {
  it('evaluates parentheses and operator precedence', () => {
    expect(evaluateArithmeticExpression('(100*10)')).toBe(1000);
    expect(evaluateArithmeticExpression('2+3*4')).toBe(14);
    expect(evaluateArithmeticExpression('(2+3)*4')).toBe(20);
  });

  it('evaluates cell references, powers and percentages', () => {
    const cells: Record<string, number> = { A1: 12, B1: 3 };
    expect(evaluateArithmeticExpression('A1/B1+2^3', ref => cells[ref.toUpperCase()] ?? null)).toBe(12);
    expect(evaluateArithmeticExpression('200*10%')).toBe(20);
  });

  it('returns spreadsheet errors for invalid operations', () => {
    expect(evaluateArithmeticExpression('10/0')).toBe('#DIV/0!');
    expect(evaluateArithmeticExpression('(10*')).toBe('#ERROR!');
  });
});

describe('xlsx keyboard navigation', () => {
  it('moves down with Enter semantics and supports every arrow direction', () => {
    expect(nextSpreadsheetCell({ row: 2, col: 3 }, 'down', 10, 10)).toEqual({ row: 3, col: 3 });
    expect(nextSpreadsheetCell({ row: 2, col: 3 }, 'up', 10, 10)).toEqual({ row: 1, col: 3 });
    expect(nextSpreadsheetCell({ row: 2, col: 3 }, 'left', 10, 10)).toEqual({ row: 2, col: 2 });
    expect(nextSpreadsheetCell({ row: 2, col: 3 }, 'right', 10, 10)).toEqual({ row: 2, col: 4 });
  });

  it('does not move outside the visible grid', () => {
    expect(nextSpreadsheetCell({ row: 0, col: 0 }, 'up', 10, 10)).toEqual({ row: 0, col: 0 });
    expect(nextSpreadsheetCell({ row: 9, col: 9 }, 'down', 10, 10)).toEqual({ row: 9, col: 9 });
  });
});
