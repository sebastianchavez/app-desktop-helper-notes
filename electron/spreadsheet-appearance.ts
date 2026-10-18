import { DOMParser, Element as XmlElement } from '@xmldom/xmldom';
import { unzipSync } from 'fflate';
import path from 'path';

/** CSS por nombre de hoja y coordenada fila:columna, con base cero. */
export type SpreadsheetAppearance = Record<string, Record<string, Record<string, string>>>;

/**
 * Obtiene hijos XML directos independientemente del prefijo del espacio de nombres.
 * @param element Nodo padre opcional.
 * @param name Nombre local opcional para filtrar los hijos.
 * @returns Elementos hijos en orden documental.
 */
function children(element: XmlElement | null | undefined, name?: string): XmlElement[] {
  return element
    ? Array.from(element.childNodes)
        .filter((node) => node.nodeType === 1)
        .map((node) => node as XmlElement)
        .filter((node) => !name || node.localName === name)
    : [];
}

/**
 * Recupera fuentes, rellenos, colores, alineación y bordes originales de un XLSX.
 * @param buffer Archivo original; XLS y CSV no usan este contenedor OOXML.
 * @returns Estilos CSS por hoja y celda; un objeto vacío para formatos sin OOXML.
 */
export function readSpreadsheetAppearance(buffer: Buffer): SpreadsheetAppearance {
  if (buffer[0] !== 0x50 || buffer[1] !== 0x4b) return {};
  const archive = unzipSync(buffer, {
    filter: (entry) => /\.(xml|rels)$/.test(entry.name) && entry.originalSize <= 50 * 1024 * 1024,
  });
  /** Recibe la ruta interna XML; retorna su raíz o null cuando no existe. */
  const parse = (name: string) =>
    archive[name]
      ? new DOMParser().parseFromString(new TextDecoder().decode(archive[name]), 'application/xml')
          .documentElement
      : null;
  const styles = parse('xl/styles.xml');
  if (!styles) return {};
  const theme = parse('xl/theme/theme1.xml');
  const scheme = theme
    ? Array.from(theme.getElementsByTagName('*')).find((node) => node.localName === 'clrScheme')
    : null;
  const themeColors = [
    'lt1',
    'dk1',
    'lt2',
    'dk2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ].map((name) => {
    const node = children(children(scheme, name)[0])[0];
    return node?.getAttribute('lastClr') || node?.getAttribute('val') || '';
  });

  /** Recibe un color OOXML; retorna RGB CSS válido, incluyendo tonos del tema, o null. */
  const color = (node: XmlElement | undefined): string | null => {
    if (!node) return null;
    const indexed = [
      '000000',
      'FFFFFF',
      'FF0000',
      '00FF00',
      '0000FF',
      'FFFF00',
      'FF00FF',
      '00FFFF',
    ];
    const rgb =
      node.getAttribute('rgb') ||
      (node.hasAttribute('theme') ? themeColors[Number(node.getAttribute('theme'))] : '') ||
      (node.hasAttribute('indexed') ? indexed[Number(node.getAttribute('indexed')) % 8] : '');
    if (!/^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(rgb ?? '')) return null;
    const tint = Math.max(-1, Math.min(1, Number(node.getAttribute('tint')) || 0));
    const channels = (rgb!.slice(-6).match(/../g) ?? []).map((channel) => {
      const value = parseInt(channel, 16);
      return Math.round(tint < 0 ? value * (1 + tint) : value + (255 - value) * tint)
        .toString(16)
        .padStart(2, '0');
    });
    return `#${channels.join('')}`;
  };
  const fonts = children(children(styles, 'fonts')[0]);
  const fills = children(children(styles, 'fills')[0]);
  const borders = children(children(styles, 'borders')[0]);
  const formats = children(children(styles, 'cellXfs')[0]).map((format) => {
    const css: Record<string, string> = {};
    const font = fonts[Number(format.getAttribute('fontId')) || 0];
    /** Recibe una propiedad booleana de fuente; retorna si está habilitada. */
    const enabled = (name: string) => {
      const node = children(font, name)[0];
      return !!node && !['0', 'false'].includes(node.getAttribute('val') ?? '');
    };
    if (enabled('b')) css['font-weight'] = 'bold';
    if (enabled('i')) css['font-style'] = 'italic';
    if (enabled('u')) css['text-decoration'] = 'underline';
    const size = Number(children(font, 'sz')[0]?.getAttribute('val'));
    if (size > 0 && size <= 200) css['font-size'] = `${size}pt`;
    const family = children(font, 'name')[0]?.getAttribute('val');
    if (family) css['font-family'] = family.replace(/[;{}<>]/g, '');
    const foreground = color(children(font, 'color')[0]);
    if (foreground) css['color'] = foreground;
    const fill = children(fills[Number(format.getAttribute('fillId')) || 0], 'patternFill')[0];
    const background = color(children(fill, 'fgColor')[0]);
    if (background && fill?.getAttribute('patternType') === 'solid')
      css['background-color'] = background;
    const alignment = children(format, 'alignment')[0];
    const horizontal = alignment?.getAttribute('horizontal');
    if (['left', 'center', 'right', 'justify'].includes(horizontal ?? ''))
      css['text-align'] = horizontal!;
    const vertical = alignment?.getAttribute('vertical');
    if (['top', 'center', 'bottom'].includes(vertical ?? ''))
      css['vertical-align'] = vertical === 'center' ? 'middle' : vertical!;
    if (alignment?.getAttribute('wrapText') === '1') css['white-space'] = 'pre-wrap';
    const border = borders[Number(format.getAttribute('borderId')) || 0];
    for (const side of ['left', 'right', 'top', 'bottom']) {
      const node = children(border, side)[0];
      const type = node?.getAttribute('style');
      if (!type || type === 'none') continue;
      const width = type === 'thick' ? 3 : type === 'medium' || type === 'double' ? 2 : 1;
      const line =
        type === 'double'
          ? 'double'
          : /dash/i.test(type)
            ? 'dashed'
            : type === 'dotted'
              ? 'dotted'
              : 'solid';
      css[`border-${side}`] =
        `${width}px ${line} ${color(children(node, 'color')[0]) ?? '#000000'}`;
    }
    return css;
  });
  const workbook = parse('xl/workbook.xml');
  const relations = children(parse('xl/_rels/workbook.xml.rels'));
  const result: SpreadsheetAppearance = {};
  for (const sheet of children(children(workbook, 'sheets')[0])) {
    const relation = relations.find(
      (node) => node.getAttribute('Id') === sheet.getAttribute('r:id'),
    );
    const target = relation?.getAttribute('Target');
    if (!target) continue;
    const file = parse(
      target.startsWith('/') ? target.slice(1) : path.posix.normalize(`xl/${target}`),
    );
    const cells: Record<string, Record<string, string>> = {};
    for (const cell of file
      ? Array.from(file.getElementsByTagName('*')).filter((node) => node.localName === 'c')
      : []) {
      const reference = cell.getAttribute('r')?.match(/^([A-Z]+)(\d+)$/);
      if (!reference) continue;
      const column =
        Array.from(reference[1]).reduce(
          (value, character) => value * 26 + character.charCodeAt(0) - 64,
          0,
        ) - 1;
      const css = formats[Number(cell.getAttribute('s')) || 0];
      if (css && Object.keys(css).length) cells[`${Number(reference[2]) - 1}:${column}`] = css;
    }
    result[sheet.getAttribute('name') || ''] = cells;
  }
  return result;
}
