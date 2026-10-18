import path from 'path';

/** Extensiones admitidas por el importador, el diálogo y el puente de Electron. */
export const SUPPORTED_DOCUMENT_EXTENSIONS = [
  'txt',
  'html',
  'htm',
  'docx',
  'pdf',
  'xlsx',
  'xls',
  'csv',
  'pptx',
] as const;

/**
 * Comprueba la extensión sin leer ni ejecutar el archivo.
 * @param filePath Nombre o ruta del documento, con cualquier combinación de mayúsculas.
 * @returns `true` cuando el importador admite su extensión.
 */
export function isSupportedDocument(filePath: string): boolean {
  return (SUPPORTED_DOCUMENT_EXTENSIONS as readonly string[]).includes(
    path.extname(filePath).slice(1).toLowerCase(),
  );
}
