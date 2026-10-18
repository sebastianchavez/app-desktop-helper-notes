/** Códigos de error SpreadsheetML/BIFF que SheetJS escribe como valores de error nativos. */
export const SPREADSHEET_ERROR_CODES: Record<string, number> = {
  '#NULL!': 0,
  '#DIV/0!': 7,
  '#VALUE!': 15,
  '#REF!': 23,
  '#NAME?': 29,
  '#NUM!': 36,
  '#N/A': 42,
  '#GETTING_DATA': 43,
};
/** Recibe código numérico del archivo; retorna el texto de error, o #VALUE! si no se reconoce. */
export function spreadsheetErrorText(code: unknown): string {
  return Object.entries(SPREADSHEET_ERROR_CODES).find(([, v]) => v === code)?.[0] ?? '#VALUE!';
}
/** Recibe resultado del motor; retorna código nativo o undefined si es texto normal, normalizando errores internos. */
export function spreadsheetErrorCode(value: unknown): number | undefined {
  if (typeof value !== 'string') return undefined;
  return SPREADSHEET_ERROR_CODES[
    value === '#CYCLE!' ? '#REF!' : value === '#ERROR!' ? '#VALUE!' : value
  ];
}
