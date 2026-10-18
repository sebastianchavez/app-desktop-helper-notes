import { parseDocxToHtml } from './document-parser';
import { getNoteById, updateNote, NoteRow } from './db/note.repository';

const MAX_WORD_BYTES = 100 * 1024 * 1024;

/**
 * Guarda el DOCX editado y su índice de búsqueda en una única escritura SQLite.
 * @param noteId Identificador de un documento Word existente y no eliminado.
 * @param base64 Contenedor DOCX exportado por el editor, codificado como Base64 estándar.
 * @returns Metadatos actualizados sin duplicar el adjunto binario en la respuesta IPC.
 * @throws Si la nota, el Base64 o el documento no son válidos; conserva el original.
 */
export function saveWordDocument(noteId: number, base64: string): Omit<NoteRow, 'file_data'> {
  if (!Number.isSafeInteger(noteId) || noteId <= 0) throw new Error('Documento no válido.');
  if (
    typeof base64 !== 'string' ||
    !base64.length ||
    base64.length > Math.ceil(MAX_WORD_BYTES / 3) * 4 ||
    base64.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(base64)
  )
    throw new Error('El contenido Word no es válido o supera los 100 MB.');
  const note = getNoteById(noteId) as (NoteRow & { deleted_at?: string | null }) | undefined;
  if (!note || note.file_type !== 'docx' || note.deleted_at)
    throw new Error('El documento Word ya no está disponible.');
  const bytes = Buffer.from(base64, 'base64');
  if (bytes.length > MAX_WORD_BYTES || bytes.toString('base64') !== base64)
    throw new Error('El contenido Word no es válido.');
  // Validar y extraer antes de escribir impide reemplazar el original por una exportación fallida.
  const content = parseDocxToHtml(bytes);
  const updated = updateNote(noteId, { file_data: base64, content });
  const { file_data: _binary, ...metadata } = updated;
  return metadata;
}
