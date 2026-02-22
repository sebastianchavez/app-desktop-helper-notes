import Database from 'better-sqlite3';
import path from 'path';
import { app } from 'electron';

let db: Database.Database;

export function getDb() {
  if (db) return db;

  const dbPath = path.join(app.getPath('userData'), 'helpernote.db');
  db = new Database(dbPath);

  db.pragma('journal_mode = WAL');
  return db;
}

export function initDb() {
  const db = getDb();

  db.pragma('foreign_keys = OFF');

  try {
    db.exec(`
      CREATE TABLE IF NOT EXISTS folders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        created_at TEXT DEFAULT (datetime('now')),
        updated_at TEXT DEFAULT (datetime('now'))
      );

      CREATE TABLE IF NOT EXISTS notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        folder_id INTEGER,
        title TEXT NOT NULL DEFAULT '',
        content TEXT NOT NULL DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        updated_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY (folder_id) REFERENCES folders(id) ON DELETE CASCADE
      );

      CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL
      );
    `);
    // Migrations
    const hasFileType = db.prepare("SELECT COUNT(*) as cnt FROM pragma_table_info('notes') WHERE name = 'file_type'").get() as { cnt: number };
    if (hasFileType.cnt === 0) {
      db.exec("ALTER TABLE notes ADD COLUMN file_data TEXT DEFAULT NULL");
      db.exec("ALTER TABLE notes ADD COLUMN file_type TEXT DEFAULT NULL");
    }

    // Soft delete migrations
    const hasNoteDeletedAt = db.prepare("SELECT COUNT(*) as cnt FROM pragma_table_info('notes') WHERE name = 'deleted_at'").get() as { cnt: number };
    if (hasNoteDeletedAt.cnt === 0) {
      db.exec("ALTER TABLE notes ADD COLUMN deleted_at TEXT DEFAULT NULL");
    }
    const hasFolderDeletedAt = db.prepare("SELECT COUNT(*) as cnt FROM pragma_table_info('folders') WHERE name = 'deleted_at'").get() as { cnt: number };
    if (hasFolderDeletedAt.cnt === 0) {
      db.exec("ALTER TABLE folders ADD COLUMN deleted_at TEXT DEFAULT NULL");
    }

    // AI Conversations tables
    db.exec(`
      CREATE TABLE IF NOT EXISTS ai_conversations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL DEFAULT 'Nueva conversación',
        model TEXT NOT NULL DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        updated_at TEXT DEFAULT (datetime('now'))
      );

      CREATE TABLE IF NOT EXISTS ai_messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        conversation_id INTEGER NOT NULL,
        role TEXT NOT NULL,
        content TEXT NOT NULL DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY (conversation_id) REFERENCES ai_conversations(id) ON DELETE CASCADE
      );
    `);
  } catch (error) {
    console.error('Error durante inicialización de BD:', error);
  } finally {
    db.pragma('foreign_keys = ON');
  }
}