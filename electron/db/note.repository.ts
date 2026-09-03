import { getDb } from './db';

/** Registro completo de nota, incluidos adjuntos Base64 y tipo de archivo. */
export interface NoteRow {
  id: number;
  folder_id: number | null;
  title: string;
  content: string;
  file_data: string | null;
  file_type: string | null;
  created_at: string;
  updated_at: string;
}

/**
 * Lista notas visibles; `null` obtiene toda la biblioteca para transferencias,
 * mientras que un identificador limita el resultado a una carpeta.
 */
export function getNotesByFolder(folderId: number | null): NoteRow[] {
  const db = getDb();
  const cols = 'id, folder_id, title, content, file_type, created_at, updated_at';
  if (folderId === null) {
    return db.prepare(`SELECT ${cols} FROM notes WHERE deleted_at IS NULL ORDER BY updated_at DESC`).all() as NoteRow[];
  }
  return db
    .prepare(`SELECT ${cols} FROM notes WHERE folder_id = ? AND deleted_at IS NULL ORDER BY updated_at DESC`)
    .all(folderId) as NoteRow[];
}

/** Recupera una nota con sus datos binarios, incluso si fue borrada lógicamente. */
export function getNoteById(id: number): NoteRow | undefined {
  const db = getDb();
  return db.prepare('SELECT * FROM notes WHERE id = ?').get(id) as NoteRow | undefined;
}

/** Inserta una nota vacía y devuelve la fila generada. */
export function createNote(folderId: number | null, title?: string): NoteRow {
  const db = getDb();
  const stmt = db.prepare('INSERT INTO notes (folder_id, title, content) VALUES (?, ?, ?)');
  const result = stmt.run(folderId, title ?? '', '');
  return db.prepare('SELECT * FROM notes WHERE id = ?').get(result.lastInsertRowid) as NoteRow;
}

/** Construye un UPDATE parcial sólo con los campos recibidos. */
export function updateNote(id: number, data: { title?: string; content?: string; folder_id?: number | null }): NoteRow {
  const db = getDb();
  const fields: string[] = [];
  const values: unknown[] = [];

  if (data.title !== undefined) {
    fields.push('title = ?');
    values.push(data.title);
  }
  if (data.content !== undefined) {
    fields.push('content = ?');
    values.push(data.content);
  }
  if (data.folder_id !== undefined) {
    fields.push('folder_id = ?');
    values.push(data.folder_id);
  }
  if ((data as any).file_data !== undefined) {
    fields.push('file_data = ?');
    values.push((data as any).file_data);
  }
  if ((data as any).file_type !== undefined) {
    fields.push('file_type = ?');
    values.push((data as any).file_type);
  }

  if (fields.length > 0) {
    fields.push("updated_at = datetime('now')");
    values.push(id);
    db.prepare(`UPDATE notes SET ${fields.join(', ')} WHERE id = ?`).run(...values);
  }

  return db.prepare('SELECT * FROM notes WHERE id = ?').get(id) as NoteRow;
}

/** Aplica borrado lógico sin destruir el contenido de la nota. */
export function deleteNote(id: number): void {
  const db = getDb();
  db.prepare("UPDATE notes SET deleted_at = datetime('now') WHERE id = ?").run(id);
}
