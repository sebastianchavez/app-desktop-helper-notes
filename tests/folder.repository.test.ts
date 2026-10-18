import { DatabaseSync } from 'node:sqlite';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const state = vi.hoisted(() => ({ db: null as any }));
vi.mock('../electron/db/db', () => ({ getDb: () => state.db }));
import { moveFolder } from '../electron/db/folder.repository';

describe('moveFolder persistence', () => {
  let sqlite: DatabaseSync;
  beforeEach(() => {
    sqlite = new DatabaseSync(':memory:');
    sqlite.exec(`
      CREATE TABLE folders (
        id INTEGER PRIMARY KEY, name TEXT, parent_id INTEGER REFERENCES folders(id),
        created_at TEXT DEFAULT 'original', updated_at TEXT DEFAULT 'original', deleted_at TEXT
      );
      CREATE TABLE notes (id INTEGER PRIMARY KEY, folder_id INTEGER REFERENCES folders(id), content TEXT);
      INSERT INTO folders (id, name, parent_id) VALUES
        (1, 'Origin', NULL), (2, 'Child', 1), (3, 'Grandchild', 2), (4, 'Target', NULL);
      INSERT INTO folders (id, name, deleted_at) VALUES (5, 'Deleted', 'today');
      INSERT INTO notes VALUES (21, 2, 'preserved content');
    `);
    state.db = {
      prepare: (sql: string) => sqlite.prepare(sql),
      transaction: (operation: () => unknown) => () => {
        sqlite.exec('BEGIN');
        try {
          const result = operation();
          sqlite.exec('COMMIT');
          return result;
        } catch (error) {
          sqlite.exec('ROLLBACK');
          throw error;
        }
      },
    };
  });
  afterEach(() => sqlite.close());

  it('persists the parent and preserves descendant folders and notes', () => {
    expect(moveFolder(1, 4).parent_id).toBe(4);
    expect(sqlite.prepare('SELECT parent_id FROM folders WHERE id = 2').get()?.['parent_id']).toBe(1);
    expect(sqlite.prepare('SELECT * FROM notes WHERE id = 21').get()).toMatchObject({ folder_id: 2, content: 'preserved content' });
  });

  it('moves a child back to the root', () => {
    expect(moveFolder(2, null).parent_id).toBeNull();
  });

  it('rejects cycles without writing changes', () => {
    for (const destination of [1, 2, 3]) expect(() => moveFolder(1, destination)).toThrow();
    expect(sqlite.prepare('SELECT parent_id FROM folders WHERE id = 1').get()?.['parent_id']).toBeNull();
  });

  it('rejects missing or deleted sources and destinations', () => {
    for (const id of [5, 99]) {
      expect(() => moveFolder(id, 4)).toThrow();
      expect(() => moveFolder(1, id)).toThrow();
    }
  });

  it('rejects malformed IPC arguments', () => {
    for (const id of [0, -1, 1.5, NaN]) expect(() => moveFolder(id, 4)).toThrow();
    expect(() => moveFolder(1, undefined as unknown as null)).toThrow();
  });

  it('leaves an unchanged parent untouched', () => {
    expect(moveFolder(2, 1).updated_at).toBe('original');
  });
});
