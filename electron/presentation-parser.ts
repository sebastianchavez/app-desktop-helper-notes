import { unzipSync } from 'fflate';

export interface ParsedPresentationSlide {
  id: string;
  title: string;
  body: string;
  notes: string;
}

export interface ParsedPresentationDocument {
  kind: 'helper-notes-presentation';
  version: 1;
  slides: ParsedPresentationSlide[];
}

const decoder = new TextDecoder('utf-8');

/**
 * Inspecciona el directorio central antes de descomprimir para rechazar ZIP
 * cifrados, tamaños anómalos y posibles bombas de compresión.
 */
function assertSafePptx(buffer: Uint8Array): void {
  if (buffer.byteLength > 100 * 1024 * 1024) throw new Error('El archivo PPTX supera el límite de 100 MB.');
  const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
  let eocd = -1;
  for (let offset = buffer.byteLength - 22; offset >= Math.max(0, buffer.byteLength - 65_557); offset--) {
    if (view.getUint32(offset, true) === 0x06054b50) {
      eocd = offset;
      break;
    }
  }
  if (eocd < 0) throw new Error('El archivo PPTX no es un ZIP válido.');

  const entryCount = view.getUint16(eocd + 10, true);
  const centralSize = view.getUint32(eocd + 12, true);
  const centralOffset = view.getUint32(eocd + 16, true);
  if (entryCount > 5_000 || centralOffset + centralSize > buffer.byteLength) {
    throw new Error('El archivo PPTX contiene una estructura no permitida.');
  }

  let cursor = centralOffset;
  let totalUncompressed = 0;
  for (let index = 0; index < entryCount; index++) {
    if (cursor + 46 > buffer.byteLength || view.getUint32(cursor, true) !== 0x02014b50) {
      throw new Error('El directorio del archivo PPTX está dañado.');
    }
    const flags = view.getUint16(cursor + 8, true);
    const uncompressedSize = view.getUint32(cursor + 24, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    if ((flags & 0x0001) !== 0 || uncompressedSize === 0xffffffff || uncompressedSize > 50 * 1024 * 1024) {
      throw new Error('El archivo PPTX contiene una entrada no permitida.');
    }
    totalUncompressed += uncompressedSize;
    if (totalUncompressed > 250 * 1024 * 1024) throw new Error('El contenido PPTX descomprimido es demasiado grande.');
    cursor += 46 + nameLength + extraLength + commentLength;
  }
}

/** Decodifica las entidades XML que pueden aparecer dentro de un run OOXML. */
function decodeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/&#(\d+);/g, (_match, value) => String.fromCodePoint(Number(value)))
    .replace(/&#x([0-9a-f]+);/gi, (_match, value) => String.fromCodePoint(parseInt(value, 16)));
}

/** Extrae y normaliza los fragmentos `<a:t>` de un bloque XML. */
function textRuns(xml: string): string[] {
  return [...xml.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)]
    .map(match => decodeXmlText(match[1]).replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

/** Conserva la separación por párrafos de una forma de PowerPoint. */
function shapeParagraphs(shapeXml: string): string[] {
  const paragraphs = [...shapeXml.matchAll(/<a:p(?:\s[^>]*)?>([\s\S]*?)<\/a:p>/g)];
  const values = paragraphs.map(paragraph => textRuns(paragraph[1]).join('')).filter(Boolean);
  return values.length > 0 ? values : textRuns(shapeXml);
}

/** Obtiene el índice numérico para ordenar slide1, slide2, etc. */
function slideNumber(fileName: string): number {
  return Number(fileName.match(/slide(\d+)\.xml$/)?.[1] ?? Number.MAX_SAFE_INTEGER);
}

/**
 * Convierte un PPTX validado al modelo editable interno. Identifica títulos
 * mediante placeholders OOXML y recupera las notas del orador asociadas.
 */
export function parsePptxBuffer(buffer: Uint8Array): ParsedPresentationDocument {
  assertSafePptx(buffer);
  const files = unzipSync(buffer);
  const slideFiles = Object.keys(files)
    .filter(name => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => slideNumber(a) - slideNumber(b));

  if (slideFiles.length === 0) throw new Error('El archivo PPTX no contiene diapositivas válidas.');

  const slides = slideFiles.map((fileName, index) => {
    const xml = decoder.decode(files[fileName]);
    const shapes = [...xml.matchAll(/<p:sp(?:\s[^>]*)?>([\s\S]*?)<\/p:sp>/g)].map(match => match[1]);
    let title = '';
    const bodyParts: string[] = [];

    for (const shape of shapes) {
      const paragraphs = shapeParagraphs(shape);
      if (paragraphs.length === 0) continue;
      const isTitle = /<p:ph\b[^>]*\btype="(?:title|ctrTitle)"/i.test(shape);
      if (isTitle && !title) title = paragraphs.join(' ');
      else bodyParts.push(...paragraphs);
    }

    if (!title && bodyParts.length > 0) title = bodyParts.shift() ?? '';

    const number = slideNumber(fileName);
    const notesFile = files[`ppt/notesSlides/notesSlide${number}.xml`];
    const notes = notesFile
      ? textRuns(decoder.decode(notesFile)).filter(text => !/^\d+$/.test(text)).join('\n')
      : '';

    return {
      id: `imported-${index + 1}`,
      title: title || `Diapositiva ${index + 1}`,
      body: bodyParts.join('\n'),
      notes,
    };
  });

  return { kind: 'helper-notes-presentation', version: 1, slides };
}
