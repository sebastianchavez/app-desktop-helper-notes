import { getDb } from './db';

export interface FolderRow {
  id: number;
  name: string;
  created_at: string;
  updated_at: string;
}

export function getAllFolders(): FolderRow[] {
  const db = getDb();
  return db.prepare('SELECT * FROM folders ORDER BY name ASC').all() as FolderRow[];
}

export function createFolder(name: string): FolderRow {
  const db = getDb();
  const stmt = db.prepare('INSERT INTO folders (name) VALUES (?)');
  const result = stmt.run(name);
  return db.prepare('SELECT * FROM folders WHERE id = ?').get(result.lastInsertRowid) as FolderRow;
}

export function updateFolder(id: number, name: string): FolderRow {
  const db = getDb();
  db.prepare('UPDATE folders SET name = ?, updated_at = datetime(\'now\') WHERE id = ?').run(name, id);
  return db.prepare('SELECT * FROM folders WHERE id = ?').get(id) as FolderRow;
}

export function deleteFolder(id: number): void {
  const db = getDb();
  db.prepare('DELETE FROM folders WHERE id = ?').run(id);
}
