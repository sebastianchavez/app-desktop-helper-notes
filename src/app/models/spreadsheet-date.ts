/** Contexto de fechas: época del libro y reloj opcional para cálculos reproducibles. */
export type SpreadsheetDateOptions = { date1904?: boolean; now?: Date };
const DAY = 86400000;

/** Recibe año/mes/día gregorianos; retorna UTC sin la conversión implícita de años 0–99 de JavaScript. */
function utc(year: number, month: number, day: number): Date {
  const date = new Date(0);
  date.setUTCFullYear(year, month, day);
  date.setUTCHours(0, 0, 0, 0);
  return date;
}

/** Recibe fecha UTC y época; retorna el serial Excel, incluyendo el salto del 29/02/1900 ficticio. */
export function spreadsheetDateSerial(date: Date, date1904 = false): number {
  const value =
    (date.getTime() - utc(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 31).getTime()) /
    DAY;
  return value + (!date1904 && date.getTime() >= utc(1900, 2, 1).getTime() ? 1 : 0);
}

/** Recibe serial y época; retorna sus componentes, preservando el día ficticio que Excel identifica como serial 60. */
export function spreadsheetDateParts(
  serial: number,
  date1904 = false,
): { year: number; month: number; day: number } {
  const integer = Math.floor(serial);
  if (!Number.isFinite(serial) || integer < 0 || integer > (date1904 ? 2957003 : 2958465))
    throw Error('#NUM!');
  if (!date1904 && integer === 60) return { year: 1900, month: 2, day: 29 };
  const date = new Date(
    utc(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 31).getTime() +
      (integer - (!date1904 && integer > 60 ? 1 : 0)) * DAY,
  );
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
}

/** Recibe función, argumentos numéricos, feriados y contexto; retorna su valor o undefined si no es una función de fecha soportada. */
export function evaluateSpreadsheetDate(
  name: string,
  args: number[],
  holidays: number[],
  options: SpreadsheetDateOptions,
): number | undefined {
  const epoch = options.date1904 ?? false;
  const serial = (y: number, m: number, d: number) =>
    spreadsheetDateSerial(utc(y, m - 1, d), epoch);
  const checked = (value: number) => {
    spreadsheetDateParts(value, epoch);
    return value;
  };
  const today = () => {
    const now = options.now ?? new Date();
    return serial(now.getFullYear(), now.getMonth() + 1, now.getDate());
  };
  const weekday = (day: number) => (((Math.floor(day) + (epoch ? 5 : 6)) % 7) + 7) % 7; // Domingo = 0, con la convención histórica de Excel.
  switch (name) {
    case 'TODAY':
      return today();
    case 'NOW': {
      const now = options.now ?? new Date();
      return today() + (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) / 86400;
    }
    case 'DATE': {
      let [y, m, d] = args.map(Math.trunc);
      if (y < 0 || y >= 10000) throw Error('#NUM!');
      if (y < 1900) y += 1900;
      // El día se suma al primer día del mes para que DATE(1900,2,29) produzca el serial 60.
      return checked(serial(y, m, 1) + d - 1);
    }
    case 'YEAR':
      return spreadsheetDateParts(args[0], epoch).year;
    case 'MONTH':
      return spreadsheetDateParts(args[0], epoch).month;
    case 'DAY':
      return spreadsheetDateParts(args[0], epoch).day;
    case 'DAYS':
      return Math.floor(args[0]) - Math.floor(args[1]);
    case 'EDATE':
    case 'EOMONTH': {
      const p = spreadsheetDateParts(args[0], epoch),
        month = p.month + Math.trunc(args[1]),
        end = serial(p.year, month + 1, 1) - 1;
      return checked(
        name === 'EOMONTH' ? end : Math.min(serial(p.year, month, 1) + p.day - 1, end),
      );
    }
    case 'WEEKDAY': {
      spreadsheetDateParts(args[0], epoch);
      const mode = args[1] ?? 1,
        w = weekday(args[0]);
      if (mode === 1) return w + 1;
      if (mode === 2) return ((w + 6) % 7) + 1;
      if (mode === 3) return (w + 6) % 7;
      if (mode >= 11 && mode <= 17) return ((w + 7 - (mode === 17 ? 0 : mode - 10)) % 7) + 1;
      throw Error('#NUM!');
    }
    case 'TIME': {
      const [h, m, s] = args.map(Math.trunc);
      if (args.some((v) => v < 0 || v > 32767)) throw Error('#NUM!');
      return ((h * 3600 + m * 60 + s) % 86400) / 86400;
    }
    case 'HOUR':
    case 'MINUTE':
    case 'SECOND': {
      if (args[0] < 0) throw Error('#NUM!');
      const seconds = Math.floor((args[0] % 1) * 86400 + 0.000001);
      return name === 'HOUR'
        ? Math.floor(seconds / 3600)
        : name === 'MINUTE'
          ? Math.floor(seconds / 60) % 60
          : seconds % 60;
    }
    case 'NETWORKDAYS':
    case 'WORKDAY': {
      spreadsheetDateParts(args[0], epoch);
      const excluded = new Set(holidays.map(Math.floor)),
        work = (d: number) => weekday(d) !== 0 && weekday(d) !== 6 && !excluded.has(d);
      let start = Math.floor(args[0]);
      if (name === 'NETWORKDAYS') {
        spreadsheetDateParts(args[1], epoch);
        const end = Math.floor(args[1]),
          sign = end >= start ? 1 : -1;
        if (Math.abs(end - start) > 100000) throw Error('#NUM!');
        let count = 0;
        for (let d = Math.min(start, end); d <= Math.max(start, end); d++) if (work(d)) count++;
        return sign * count;
      }
      let remaining = Math.abs(Math.trunc(args[1])),
        sign = args[1] < 0 ? -1 : 1,
        iterations = 0;
      if (remaining > 70000) throw Error('#NUM!');
      while (remaining) {
        start += sign;
        if (++iterations > 100000) throw Error('#NUM!');
        if (work(start)) remaining--;
      }
      return checked(start);
    }
    default:
      return undefined;
  }
}
