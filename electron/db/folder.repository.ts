import { getDb } from './db';

/** Representación SQLite de una carpeta visible. */
export interface FolderRow {
  id: number;
  name: string;
  parent_id: number | null;
  created_at: string;
  updated_at: string;
}

/** Devuelve carpetas no eliminadas ordenadas alfabéticamente. */
export function getAllFolders(): FolderRow[] {
  const db = getDb();
  return db.prepare('SELECT * FROM folders WHERE deleted_at IS NULL ORDER BY name ASC').all() as FolderRow[];
}

/** Crea una carpeta después de comprobar que su padre siga disponible. */
export function createFolder(name: string, parentId: number | null = null): FolderRow {
  const db = getDb();
  if (parentId !== null) {
    const parent = db.prepare('SELECT id FROM folders WHERE id = ? AND deleted_at IS NULL').get(parentId);
    if (!parent) throw new Error('La carpeta padre no existe o fue eliminada.');
  }
  const stmt = db.prepare('INSERT INTO folders (name, parent_id) VALUES (?, ?)');
  const result = stmt.run(name, parentId);
  return db.prepare('SELECT * FROM folders WHERE id = ?').get(result.lastInsertRowid) as FolderRow;
}

/** Renombra la carpeta y actualiza su marca temporal. */
export function updateFolder(id: number, name: string): FolderRow {
  const db = getDb();
  db.prepare('UPDATE folders SET name = ?, updated_at = datetime(\'now\') WHERE id = ?').run(name, id);
  return db.prepare('SELECT * FROM folders WHERE id = ?').get(id) as FolderRow;
}

/**
 * Marca como eliminada una rama completa y todas sus notas en una transacción.
 * El CTE recursivo resuelve descendientes sin depender del árbol del renderer.
 */
export function deleteFolder(id: number): void {
  const db = getDb();
  const deleteTree = db.transaction((folderId: number) => {
    const descendants = db.prepare(`
      WITH RECURSIVE folder_tree(id) AS (
        SELECT id FROM folders WHERE id = ? AND deleted_at IS NULL
        UNION ALL
        SELECT child.id
        FROM folders child
        JOIN folder_tree parent ON child.parent_id = parent.id
        WHERE child.deleted_at IS NULL
      )
      SELECT id FROM folder_tree
    `).all(folderId) as Array<{ id: number }>;

    if (descendants.length === 0) return;
    const ids = descendants.map(folder => folder.id);
    const placeholders = ids.map(() => '?').join(', ');
    db.prepare(`UPDATE notes SET deleted_at = datetime('now') WHERE folder_id IN (${placeholders}) AND deleted_at IS NULL`).run(...ids);
    db.prepare(`UPDATE folders SET deleted_at = datetime('now') WHERE id IN (${placeholders})`).run(...ids);
  });
  deleteTree(id);
}
