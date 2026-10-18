import { readFileSync, statSync } from 'fs';
import path from 'path';
import { parseDocxToHtml } from './document-parser';
import { parsePptxBuffer } from './presentation-parser';
import { isSupportedDocument } from './document-types';

export interface ImportedDocument {
  title: string;
  content: string;
  file_data?: string;
  file_type?: string;
}

/**
 * Convierte documentos físicos sin modificar sus originales. Los errores de un archivo
 * no impiden importar los demás; la persistencia de las notas corresponde al renderer.
 * @param paths Rutas absolutas recibidas del diálogo, del arrastre o del sistema operativo.
 * @returns Documentos convertidos y errores por archivo. Las rutas repetidas se leen una sola vez.
 * @throws Si el argumento no es una lista de cadenas.
 */
export function importDocumentPaths(paths: string[]): {
  files: ImportedDocument[];
  errors: string[];
} {
  if (!Array.isArray(paths) || paths.some((filePath) => typeof filePath !== 'string')) {
    throw new Error('Lista de archivos inválida.');
  }
  const files: ImportedDocument[] = [];
  const errors: string[] = [];
  for (const filePath of new Set(paths)) {
    try {
      if (
        !path.isAbsolute(filePath) ||
        !isSupportedDocument(filePath) ||
        !statSync(filePath).isFile()
      ) {
        throw new Error('Formato o ruta no compatible.');
      }
      const extension = path.extname(filePath).toLowerCase();
      const title = path.basename(filePath, path.extname(filePath));
      const buffer = readFileSync(filePath);
      if (extension === '.txt') {
        files.push({
          title,
          content: decodePlainText(buffer),
          file_data: buffer.toString('base64'),
          file_type: 'txt',
        });
      } else if (extension === '.html' || extension === '.htm') {
        const html = buffer.toString('utf8');
        files.push({
          title,
          content: (html.match(/<body[^>]*>([\s\S]*?)<\/body>/i)?.[1] ?? html).trim(),
        });
      } else if (extension === '.docx') {
        files.push({
          title,
          content: parseDocxToHtml(buffer),
          file_data: buffer.toString('base64'),
          file_type: 'docx',
        });
      } else if (extension === '.pptx') {
        files.push({
          title,
          content: JSON.stringify(parsePptxBuffer(buffer)),
          file_data: buffer.toString('base64'),
          file_type: 'pptx',
        });
      } else {
        files.push({
          title,
          content: '',
          file_data: buffer.toString('base64'),
          file_type: extension === '.pdf' ? 'pdf' : 'xlsx',
        });
      }
    } catch (error) {
      errors.push(
        `${path.basename(filePath)}: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
  }
  return { files, errors };
}

/**
 * Decodifica texto Unicode respetando BOM y manteniendo saltos de línea y espacios.
 * @param buffer Bytes originales de un TXT UTF-8, UTF-16LE o UTF-16BE.
 * @returns Texto plano sin convertirlo en HTML.
 */
function decodePlainText(buffer: Buffer): string {
  const encoding =
    buffer[0] === 0xff && buffer[1] === 0xfe
      ? 'utf-16le'
      : buffer[0] === 0xfe && buffer[1] === 0xff
        ? 'utf-16be'
        : 'utf-8';
  return new TextDecoder(encoding).decode(buffer);
}
