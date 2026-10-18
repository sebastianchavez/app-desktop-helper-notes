import { DOMParser, Element as XmlElement } from '@xmldom/xmldom';
import { unzipSync } from 'fflate';

/**
 * Escapa texto antes de incorporarlo al HTML de búsqueda o edición.
 * @param text Texto plano extraído del documento.
 * @returns Texto con caracteres especiales representados como entidades HTML.
 */
export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Extrae un resumen HTML del DOCX mediante XML real, sin confundir w:t con w:tabs o w:tbl.
 * La vista del documento usa el archivo original; este HTML sirve para buscar y crear una nota editable.
 * @param buffer Contenido binario del archivo DOCX.
 * @returns Párrafos, saltos de línea y formato básico sin exponer etiquetas OOXML.
 * @throws Si el contenedor o el documento principal no son válidos.
 */
export function parseDocxToHtml(buffer: Buffer): string {
  const entries = unzipSync(buffer, {
    filter: (entry) => entry.name === 'word/document.xml' && entry.originalSize <= 50 * 1024 * 1024,
  });
  const bytes = entries['word/document.xml'];
  if (!bytes) throw new Error('El DOCX no contiene word/document.xml.');
  const xml = new DOMParser().parseFromString(new TextDecoder().decode(bytes), 'application/xml');
  const root = xml.documentElement;
  if (
    root?.localName !== 'document' ||
    root.namespaceURI !== 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  )
    throw new Error('El DOCX no contiene un documento Word válido.');
  const paragraphs = Array.from(xml.getElementsByTagName('w:p')).map((paragraph) => {
    const alignment = paragraph.getElementsByTagName('w:jc')[0]?.getAttribute('w:val');
    const align = alignment === 'both' ? 'justify' : alignment;
    const attribute = ['left', 'center', 'right', 'justify'].includes(align ?? '')
      ? ` style="text-align:${align}"`
      : '';
    const style = paragraph.getElementsByTagName('w:pStyle')[0]?.getAttribute('w:val') ?? '';
    const tag = /^heading[1-3]$/i.test(style) ? `h${style.slice(-1)}` : 'p';
    const runs = Array.from(paragraph.getElementsByTagName('w:r'))
      .map((run) => {
        let text = Array.from(run.childNodes)
          .map((node) => {
            if (node.nodeType !== 1) return '';
            const element = node as XmlElement;
            if (element.tagName === 'w:t') return escapeHtml(element.textContent ?? '');
            if (element.tagName === 'w:tab') return '&#9;';
            if (element.tagName === 'w:br' || element.tagName === 'w:cr') return '<br>';
            return '';
          })
          .join('');
        for (const [property, html] of [
          ['w:b', 'b'],
          ['w:i', 'i'],
          ['w:u', 'u'],
          ['w:strike', 's'],
        ]) {
          const setting = run.getElementsByTagName(property)[0];
          if (setting && !['false', '0', 'none'].includes(setting.getAttribute('w:val') ?? ''))
            text = `<${html}>${text}</${html}>`;
        }
        return text;
      })
      .join('');
    return `<${tag}${attribute}>${runs || '<br>'}</${tag}>`;
  });
  return paragraphs.join('') || '<p></p>';
}
