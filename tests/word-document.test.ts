import { DatabaseSync } from 'node:sqlite';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { strToU8, zipSync } from 'fflate';

const state = vi.hoisted(() => ({ db: null as unknown }));
vi.mock('../electron/db/db', () => ({ getDb: () => state.db }));
import { saveWordDocument } from '../electron/word-document';

/**
 * Construye un DOCX mínimo con contenido y una imagen para comprobar la conservación binaria.
 * @param text Texto que debe aparecer en el índice de búsqueda después de guardar.
 * @returns Contenedor ZIP codificado como Base64 para el canal IPC.
 */
function documentBase64(text: string): string {
  return Buffer.from(
    zipSync({
      'word/document.xml': strToU8(
        `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
      ),
      'word/media/image1.png': new Uint8Array([1, 2, 3, 4]),
    }),
  ).toString('base64');
}

describe('Word document persistence', () => {
  let sqlite: DatabaseSync;
  beforeEach(() => {
    sqlite = new DatabaseSync(':memory:');
    sqlite.exec(`CREATE TABLE notes (
      id INTEGER PRIMARY KEY, folder_id INTEGER, title TEXT, content TEXT,
      file_data TEXT, file_type TEXT, created_at TEXT, updated_at TEXT, deleted_at TEXT
    ); INSERT INTO notes VALUES (1, 7, 'Título actual', 'original', 'original', 'docx', 'original', 'original', NULL);`);
    state.db = sqlite;
  });
  afterEach(() => sqlite.close());

  it('saves the complete exported binary and search text without overwriting title or folder', () => {
    const base64 = documentBase64('Texto editado');
    const metadata = saveWordDocument(1, base64);
    expect(metadata).toMatchObject({
      id: 1,
      title: 'Título actual',
      folder_id: 7,
      content: '<p>Texto editado</p>',
    });
    expect(metadata).not.toHaveProperty('file_data');
    expect(sqlite.prepare('SELECT file_data FROM notes WHERE id = 1').get()?.['file_data']).toBe(
      base64,
    );
  });

  it('keeps the original intact when the export is corrupt', () => {
    for (const invalid of ['not base64', 'AAAA', Buffer.from('not a zip').toString('base64')])
      expect(() => saveWordDocument(1, invalid)).toThrow();
    expect(sqlite.prepare('SELECT file_data, content FROM notes WHERE id = 1').get()).toMatchObject(
      { file_data: 'original', content: 'original' },
    );
  });

  it('rejects a ZIP whose document entry contains unrelated XML', () => {
    const invalid = Buffer.from(zipSync({ 'word/document.xml': strToU8('<unrelated/>') })).toString(
      'base64',
    );
    expect(() => saveWordDocument(1, invalid)).toThrow();
    expect(sqlite.prepare('SELECT file_data FROM notes WHERE id = 1').get()?.['file_data']).toBe(
      'original',
    );
  });

  it('rejects missing, deleted and non-Word notes', () => {
    const base64 = documentBase64('Texto');
    for (const id of [0, -1, 1.5, 99]) expect(() => saveWordDocument(id, base64)).toThrow();
    sqlite.exec("UPDATE notes SET file_type = 'xlsx' WHERE id = 1");
    expect(() => saveWordDocument(1, base64)).toThrow();
    sqlite.exec("UPDATE notes SET file_type = 'docx', deleted_at = 'deleted' WHERE id = 1");
    expect(() => saveWordDocument(1, base64)).toThrow();
  });
});
