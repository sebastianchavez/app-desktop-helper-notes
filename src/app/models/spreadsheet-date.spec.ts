import { describe, it, expect } from 'vitest';
import { evaluateSpreadsheetFormula } from './spreadsheet-formula';
import { spreadsheetDateParts, spreadsheetDateSerial } from './spreadsheet-date';
/** Recibe expresión y época opcional; retorna su cálculo con reloj fijo sin leer celdas del usuario. */
function run(formula: string, date1904 = false) {
  return evaluateSpreadsheetFormula(formula, () => null, {
    date1904,
    now: new Date(2026, 9, 15, 13, 14, 15),
  });
}
describe('Excel date and time compatibility', () => {
  it('supports both epochs with the historical leap-day convention', () => {
    expect(run('=DATE(1900;1;1)')).toBe(1);
    expect(run('=DATE(1904;1;1)', true)).toBe(0);
    expect((run('=DATE(2026;10;15)') as number) - (run('=DATE(2026;10;15)', true) as number)).toBe(
      1462,
    );
    expect(run('=DATE(1900;2;29)')).toBe(60);
    expect(run('=DAY(60)')).toBe(29);
    expect(run('=MONTH(60)')).toBe(2);
    expect(run('=DATE(1900;3;1)')).toBe(61);
    expect(run('=EOMONTH(32;0)')).toBe(60);
  });
  it('normalizes overflow and clamps month offsets rather than overflowing the target month', () => {
    expect(run('=DATE(2026;13;1)')).toBe(run('=DATE(2027;1;1)'));
    expect(run('=EDATE(DATE(2024;1;31);1)')).toBe(run('=DATE(2024;2;29)'));
    expect(run('=EOMONTH(DATE(2025;1;10);1)')).toBe(run('=DATE(2025;2;28)'));
    expect(run('=YEAR(DATE(2026;10;15))')).toBe(2026);
    expect(run('=DAYS(DATE(2026;10;15);DATE(2026;10;1))')).toBe(14);
  });
  it('counts weekdays and holiday exclusions in both directions', () => {
    expect(run('=WEEKDAY(DATE(2026;10;15);2)')).toBe(4);
    expect(run('=NETWORKDAYS(DATE(2026;10;12);DATE(2026;10;16))')).toBe(5);
    expect(run('=NETWORKDAYS(DATE(2026;10;16);DATE(2026;10;12);DATE(2026;10;14))')).toBe(-4);
    expect(run('=WORKDAY(DATE(2026;10;16);1)')).toBe(run('=DATE(2026;10;19)'));
  });
  it('calculates fractional times and uses one deterministic local clock', () => {
    expect(run('=HOUR(TIME(13;14;15))')).toBe(13);
    expect(run('=MINUTE(TIME(13;14;15))')).toBe(14);
    expect(run('=SECOND(TIME(13;14;15))')).toBe(15);
    expect(run('=TODAY()')).toBe(run('=DATE(2026;10;15)'));
    expect(run('=NOW()-TODAY()')).toBeCloseTo((13 * 3600 + 14 * 60 + 15) / 86400);
  });
  it('rejects dates outside supported bounds and keeps conversion independent of timezone', () => {
    expect(run('=DATE(-1;1;1)')).toBe('#NUM!');
    expect(run('=TIME(-1;0;0)')).toBe('#NUM!');
    const date = new Date(Date.UTC(2026, 9, 15));
    expect(spreadsheetDateParts(spreadsheetDateSerial(date))).toEqual({
      year: 2026,
      month: 10,
      day: 15,
    });
  });
});
