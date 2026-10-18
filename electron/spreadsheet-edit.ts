import {
  ensureDynamicSpreadsheetMetadata,
  patchSpreadsheetTables,
  NativeTable,
} from './spreadsheet-structure';
import { DOMParser, XMLSerializer, Element as XmlElement } from '@xmldom/xmldom';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import * as XLSX from 'xlsx';
import path from 'node:path';
import { spreadsheetErrorText } from './spreadsheet-values';

export type SpreadsheetFormat = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align?: 'left' | 'center' | 'right';
  fontSize?: number;
  color?: string;
  background?: string;
  wrap?: boolean;
  numberFormat?: string;
};
export type SpreadsheetEditOptions = {
  tables?: NativeTable[];
  definedNames?: Record<string, string>;
  spills?: { row: number; col: number; rows: number; cols: number }[];
  formats?: { row: number; col: number; format: SpreadsheetFormat }[];
  colWidths?: number[];
  rowHeights?: number[];
  hiddenRows?: number[];
  origins?: Record<string, string>;
};
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Recibe un elemento y nombre local; retorna sus hijos XML directos. */
function children(element: XmlElement, name?: string): XmlElement[] {
  return Array.from(element.childNodes)
    .filter((n) => n.nodeType === 1)
    .map((n) => n as XmlElement)
    .filter((n) => !name || n.localName === name);
}

/**
 * Actualiza únicamente datos/formatos de una hoja dentro del ZIP original.
 * Conserva gráficos, imágenes, estilos ajenos, relaciones y hojas no editadas.
 * @param original Archivo completo Base64, también admite XLS/CSV normalizados a XLSX.
 * @param generated Libro con los valores/fórmulas nuevos y las combinaciones finales.
 * @param index Hoja modificada con base cero.
 * @param options Cambios de formato, dimensiones y origen de celdas desplazadas.
 * @returns XLSX completo Base64; rechaza una hoja inexistente o coordenadas inválidas.
 */
export function patchSpreadsheet(
  original: string,
  generated: XLSX.WorkBook,
  index: number,
  options: SpreadsheetEditOptions = {},
): string {
  let bytes = Buffer.from(original, 'base64');
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b)
    bytes = Buffer.from(
      XLSX.write(XLSX.read(bytes, { type: 'buffer', cellStyles: true }), {
        type: 'buffer',
        bookType: 'xlsx',
      }),
    );
  const archive = unzipSync(bytes);
  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const parse = (key: string) => parser.parseFromString(strFromU8(archive[key]), 'application/xml');
  const workbook = parse('xl/workbook.xml');
  const sheet = Array.from(workbook.getElementsByTagNameNS(NS, 'sheet'))[index];
  if (!sheet) throw new Error('Hoja inexistente.');
  const relationship = Array.from(
    parse('xl/_rels/workbook.xml.rels').getElementsByTagName('Relationship'),
  ).find((r) => r.getAttribute('Id') === sheet.getAttribute('r:id'));
  if (!relationship) throw new Error('Relación de hoja inexistente.');
  const target = relationship.getAttribute('Target')!;
  const sheetPath = target.startsWith('/') ? target.slice(1) : path.posix.normalize(`xl/${target}`);
  const doc = parse(sheetPath);
  const root = doc.documentElement!;
  const oldData = children(root, 'sheetData')[0];
  const oldCells = new Map(
    Array.from(oldData.getElementsByTagNameNS(NS, 'c')).map((c) => [c.getAttribute('r')!, c]),
  );
  // Generar una sola hoja con cadenas inline: no reemplazar sharedStrings del archivo original.
  const temporary = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(temporary, generated.Sheets[generated.SheetNames[index]], 'Edited');
  const fresh = unzipSync(
    XLSX.write(temporary, { type: 'buffer', bookType: 'xlsx', bookSST: false }),
  );
  const freshDoc = parser.parseFromString(
    strFromU8(fresh['xl/worksheets/sheet1.xml']),
    'application/xml',
  );
  const freshRoot = freshDoc.documentElement!;
  const newData = doc.importNode(children(freshRoot, 'sheetData')[0], true) as XmlElement;
  const rows = new Map(children(newData, 'row').map((row) => [Number(row.getAttribute('r')), row]));
  const cells = new Map(
    Array.from(newData.getElementsByTagNameNS(NS, 'c')).map((c) => [c.getAttribute('r')!, c]),
  );
  /** Recibe referencia A1; retorna la celda, creando fila/celda en orden si estaba vacía. */
  const ensureCell = (reference: string): XmlElement => {
    if (!/^[A-Z]{1,3}[1-9]\d*$/.test(reference)) throw new Error('Celda inválida.');
    let cell = cells.get(reference);
    if (cell) return cell;
    const address = XLSX.utils.decode_cell(reference);
    let row = rows.get(address.r + 1);
    if (!row) {
      row = doc.createElementNS(NS, 'row');
      row.setAttribute('r', String(address.r + 1));
      const after = children(newData, 'row').find(
        (r) => Number(r.getAttribute('r')) > address.r + 1,
      );
      newData.insertBefore(row, after ?? null);
      rows.set(address.r + 1, row);
    }
    cell = doc.createElementNS(NS, 'c');
    cell.setAttribute('r', reference);
    const after = children(row, 'c').find(
      (c) => XLSX.utils.decode_cell(c.getAttribute('r')!).c > address.c,
    );
    row.insertBefore(cell, after ?? null);
    cells.set(reference, cell);
    return cell;
  };
  // Copiar IDs de estilo, incluso para celdas vacías; origins mueve estilos con filas/columnas.
  const origins =
    options.origins ?? Object.fromEntries(Array.from(oldCells.keys()).map((r) => [r, r]));
  for (const [destination, source] of Object.entries(origins)) {
    const style = oldCells.get(source)?.getAttribute('s');
    if (style) ensureCell(destination).setAttribute('s', style);
    const old = oldCells.get(source),
      cell = cells.get(destination);
    if (!old || !cell) continue;
    const oldFormula = children(old, 'f')[0],
      newFormula = children(cell, 'f')[0];
    // Las matrices que el motor no calcula conservan su definición original si la expresión no cambió.
    if (
      oldFormula?.getAttribute('t') === 'array' &&
      newFormula?.textContent === oldFormula.textContent
    ) {
      for (const attr of Array.from(oldFormula.attributes))
        newFormula.setAttribute(attr.name, attr.value);
    }
    if (oldFormula && newFormula?.textContent === oldFormula.textContent) {
      for (const name of ['cm', 'vm'])
        if (old.hasAttribute(name)) cell.setAttribute(name, old.getAttribute(name)!);
    }
    // Un error importado sin fórmula no debe transformarse en texto al reconstruir sheetData.
    const originalError = children(old, 'v')[0]?.textContent ?? '';
    const errorText = originalError.startsWith('#')
      ? originalError
      : spreadsheetErrorText(Number(originalError));
    const text = children(cell, 'is')[0]?.textContent ?? children(cell, 'v')[0]?.textContent;
    if (old.getAttribute('t') === 'e' && !newFormula && text === errorText) {
      for (const child of Array.from(cell.childNodes)) cell.removeChild(child);
      cell.setAttribute('t', 'e');
      const value = doc.createElementNS(NS, 'v');
      value.textContent = errorText;
      cell.appendChild(value);
    }
  }
  if (options.spills?.length) {
    const cm = ensureDynamicSpreadsheetMetadata(archive);
    for (const spill of options.spills) {
      if (
        ![spill.row, spill.col, spill.rows, spill.cols].every(Number.isInteger) ||
        spill.rows < 1 ||
        spill.cols < 1 ||
        spill.row < 0 ||
        spill.col < 0 ||
        spill.row + spill.rows > 1048576 ||
        spill.col + spill.cols > 16384 ||
        spill.rows * spill.cols > 100000
      )
        throw Error('Matriz inválida');
      const address = XLSX.utils.encode_cell({ r: spill.row, c: spill.col }),
        cell = cells.get(address),
        f = cell ? children(cell, 'f')[0] : undefined;
      if (!cell || !f) continue;
      f.setAttribute('t', 'array');
      f.setAttribute(
        'ref',
        XLSX.utils.encode_range({
          s: { r: spill.row, c: spill.col },
          e: { r: spill.row + spill.rows - 1, c: spill.col + spill.cols - 1 },
        }),
      );
      cell.setAttribute('cm', String(cm));
    }
  }
  for (const cell of cells.values())
    if (children(cell, 'f')[0]) {
      const v = children(cell, 'v')[0],
        text = children(cell, 'is')[0]?.textContent ?? v?.textContent;
      if (text === '#SPILL!' || text === '#CALC!') {
        for (const node of [...children(cell, 'v'), ...children(cell, 'is')])
          cell.removeChild(node);
        cell.setAttribute('t', 'e');
        const value = doc.createElementNS(NS, 'v');
        value.textContent = text;
        cell.appendChild(value);
      }
    }
  let styles = archive['xl/styles.xml'] ? parse('xl/styles.xml') : null;
  if (options.formats?.length && styles) {
    const styleRoot = styles.documentElement!;
    const fonts = children(styleRoot, 'fonts')[0],
      fills = children(styleRoot, 'fills')[0];
    const xfs = children(styleRoot, 'cellXfs')[0];
    const cache = new Map<string, string>();
    for (const update of options.formats) {
      const cell = ensureCell(XLSX.utils.encode_cell({ r: update.row, c: update.col }));
      const baseId = Number(cell.getAttribute('s') || 0),
        format = update.format;
      const key = JSON.stringify([baseId, format]);
      if (cache.has(key)) {
        cell.setAttribute('s', cache.get(key)!);
        continue;
      }
      const xf = children(xfs, 'xf')[baseId]?.cloneNode(true) as XmlElement;
      if (!xf) throw new Error('Estilo de celda inválido.');
      if (['bold', 'italic', 'underline', 'fontSize', 'color'].some((k) => k in format)) {
        const font = children(fonts, 'font')[Number(xf.getAttribute('fontId') || 0)].cloneNode(
          true,
        ) as XmlElement;
        for (const [prop, tag] of [
          ['bold', 'b'],
          ['italic', 'i'],
          ['underline', 'u'],
          ['fontSize', 'sz'],
          ['color', 'color'],
        ] as const) {
          if (format[prop] === undefined) continue;
          children(font, tag).forEach((n) => font.removeChild(n));
          const value = format[prop];
          if (value === false) continue;
          const node = styles.createElementNS(NS, tag);
          if (prop === 'fontSize') node.setAttribute('val', String(Number(value) * 0.75));
          if (prop === 'color' && /^#[\da-f]{6}$/i.test(String(value)))
            node.setAttribute('rgb', `FF${String(value).slice(1)}`);
          font.appendChild(node);
        }
        xf.setAttribute('fontId', String(children(fonts, 'font').length));
        fonts.appendChild(font);
        fonts.setAttribute('count', String(children(fonts, 'font').length));
        xf.setAttribute('applyFont', '1');
      }
      if (format.background && /^#[\da-f]{6}$/i.test(format.background)) {
        const fill = styles.createElementNS(NS, 'fill'),
          pattern = styles.createElementNS(NS, 'patternFill');
        pattern.setAttribute('patternType', 'solid');
        const color = styles.createElementNS(NS, 'fgColor');
        color.setAttribute('rgb', `FF${format.background.slice(1)}`);
        pattern.appendChild(color);
        fill.appendChild(pattern);
        xf.setAttribute('fillId', String(children(fills, 'fill').length));
        fills.appendChild(fill);
        fills.setAttribute('count', String(children(fills, 'fill').length));
        xf.setAttribute('applyFill', '1');
      }
      if (format.align !== undefined || format.wrap !== undefined) {
        let alignment = children(xf, 'alignment')[0];
        if (!alignment) {
          alignment = styles.createElementNS(NS, 'alignment');
          xf.appendChild(alignment);
        }
        if (format.align) alignment.setAttribute('horizontal', format.align);
        if (format.wrap !== undefined) alignment.setAttribute('wrapText', format.wrap ? '1' : '0');
        xf.setAttribute('applyAlignment', '1');
      }
      if (format.numberFormat !== undefined) {
        let formats = children(styleRoot, 'numFmts')[0];
        if (!formats) {
          formats = styles.createElementNS(NS, 'numFmts');
          styleRoot.insertBefore(formats, styleRoot.firstChild);
        }
        const id =
          Math.max(163, ...children(formats).map((n) => Number(n.getAttribute('numFmtId')))) + 1;
        const number = styles.createElementNS(NS, 'numFmt');
        number.setAttribute('numFmtId', String(id));
        number.setAttribute('formatCode', format.numberFormat);
        formats.appendChild(number);
        formats.setAttribute('count', String(children(formats).length));
        xf.setAttribute('numFmtId', String(id));
        xf.setAttribute('applyNumberFormat', '1');
      }
      const id = String(children(xfs, 'xf').length);
      xfs.appendChild(xf);
      xfs.setAttribute('count', String(children(xfs, 'xf').length));
      cell.setAttribute('s', id);
      cache.set(key, id);
    }
    archive['xl/styles.xml'] = strToU8(serializer.serializeToString(styles));
  }
  // Altura y atributos originales de las filas sobreviven cuando no hubo cambios estructurales.
  for (const row of children(newData, 'row')) {
    const i = Number(row.getAttribute('r')) - 1;
    if (options.rowHeights?.[i]) {
      row.setAttribute('ht', String(options.rowHeights[i] * 0.75));
      row.setAttribute('customHeight', '1');
    } else {
      const previous = children(oldData, 'row').find(
        (r) => r.getAttribute('r') === row.getAttribute('r'),
      );
      for (const key of ['ht', 'customHeight', 'hidden', 'outlineLevel'])
        if (previous?.hasAttribute(key)) row.setAttribute(key, previous.getAttribute(key)!);
    }
  }
  if (options.hiddenRows) {
    children(newData, 'row').forEach((r) => r.removeAttribute('hidden'));
    for (const row of options.hiddenRows) {
      if (!Number.isInteger(row) || row < 0 || row >= 1048576) throw Error('Fila oculta inválida');
      ensureCell('A' + (row + 1));
      rows.get(row + 1)!.setAttribute('hidden', '1');
    }
  }
  root.replaceChild(newData, oldData);
  const originalCols = children(root, 'cols')[0];
  for (const name of ['dimension', 'mergeCells', 'cols']) {
    if (name === 'cols' && !options.colWidths) continue;
    children(root, name).forEach((node) => root.removeChild(node));
    let replacement = children(freshRoot, name)[0];
    if (name === 'cols' && options.colWidths) {
      replacement = freshDoc.createElementNS(NS, 'cols');
      options.colWidths.forEach((width, i) => {
        const column = freshDoc.createElementNS(NS, 'col');
        const previous = children(originalCols ?? root, 'col').find(
          (c) => Number(c.getAttribute('min')) <= i + 1 && Number(c.getAttribute('max')) >= i + 1,
        );
        for (const attribute of ['hidden', 'outlineLevel', 'collapsed', 'style', 'bestFit'])
          if (previous?.hasAttribute(attribute))
            column.setAttribute(attribute, previous.getAttribute(attribute)!);
        column.setAttribute('min', String(i + 1));
        column.setAttribute('max', String(i + 1));
        column.setAttribute('width', String(Math.max(1, ((Number(width) || 80) - 5) / 7)));
        column.setAttribute('customWidth', '1');
        replacement.appendChild(column);
      });
    }
    if (replacement) {
      const imported = doc.importNode(replacement, true);
      if (name === 'dimension') root.insertBefore(imported, root.firstChild);
      else if (name === 'cols') root.insertBefore(imported, newData);
      else root.insertBefore(imported, newData.nextSibling);
    }
  }
  archive[sheetPath] = strToU8(serializer.serializeToString(doc));
  if (options.tables) patchSpreadsheetTables(archive, sheetPath, options.tables);
  // Las definiciones conservan su ámbito; actualizar únicamente nombres ya presentes en el libro.
  if (options.definedNames) {
    const names = Array.from(workbook.getElementsByTagNameNS(NS, 'definedName'));
    for (const name of names) {
      if (name.hasAttribute('localSheetId') && Number(name.getAttribute('localSheetId')) !== index)
        continue;
      const key = name.getAttribute('name')!.toUpperCase();
      if (
        !name.hasAttribute('localSheetId') &&
        names.some(
          (n) =>
            n.hasAttribute('localSheetId') &&
            Number(n.getAttribute('localSheetId')) === index &&
            n.getAttribute('name')?.toUpperCase() === key,
        )
      )
        continue;
      const value = options.definedNames[key];
      if (value !== undefined) name.textContent = value;
    }
  }
  // Excel/Calc recalculan funciones no cubiertas por el motor local al abrir el archivo.
  let calc = workbook.getElementsByTagNameNS(NS, 'calcPr')[0];
  if (!calc) {
    calc = workbook.createElementNS(NS, 'calcPr');
    workbook.documentElement!.appendChild(calc);
  }
  calc.setAttribute('fullCalcOnLoad', '1');
  calc.setAttribute('forceFullCalc', '1');
  calc.setAttribute('calcMode', 'auto');
  archive['xl/workbook.xml'] = strToU8(serializer.serializeToString(workbook));
  return Buffer.from(zipSync(archive)).toString('base64');
}
