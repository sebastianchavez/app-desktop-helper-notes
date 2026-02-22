import { getDb } from './db';

export interface ConversationRow {
  id: number;
  title: string;
  model: string;
  created_at: string;
  updated_at: string;
}

export interface MessageRow {
  id: number;
  conversation_id: number;
  role: string;
  content: string;
  created_at: string;
}

export function getAllConversations(): ConversationRow[] {
  const db = getDb();
  return db
    .prepare('SELECT * FROM ai_conversations ORDER BY updated_at DESC')
    .all() as ConversationRow[];
}

export function getConversationById(id: number): ConversationRow | undefined {
  const db = getDb();
  return db.prepare('SELECT * FROM ai_conversations WHERE id = ?').get(id) as ConversationRow | undefined;
}

export function createConversation(title: string, model: string): ConversationRow {
  const db = getDb();
  const result = db
    .prepare('INSERT INTO ai_conversations (title, model) VALUES (?, ?)')
    .run(title, model);
  return db
    .prepare('SELECT * FROM ai_conversations WHERE id = ?')
    .get(result.lastInsertRowid) as ConversationRow;
}

export function updateConversation(id: number, data: { title?: string; model?: string }): ConversationRow {
  const db = getDb();
  const fields: string[] = [];
  const values: unknown[] = [];

  if (data.title !== undefined) {
    fields.push('title = ?');
    values.push(data.title);
  }
  if (data.model !== undefined) {
    fields.push('model = ?');
    values.push(data.model);
  }

  if (fields.length > 0) {
    fields.push("updated_at = datetime('now')");
    values.push(id);
    db.prepare(`UPDATE ai_conversations SET ${fields.join(', ')} WHERE id = ?`).run(...values);
  }

  return db.prepare('SELECT * FROM ai_conversations WHERE id = ?').get(id) as ConversationRow;
}

export function deleteConversation(id: number): void {
  const db = getDb();
  db.prepare('DELETE FROM ai_messages WHERE conversation_id = ?').run(id);
  db.prepare('DELETE FROM ai_conversations WHERE id = ?').run(id);
}

export function getMessagesByConversation(conversationId: number): MessageRow[] {
  const db = getDb();
  return db
    .prepare('SELECT * FROM ai_messages WHERE conversation_id = ? ORDER BY created_at ASC')
    .all(conversationId) as MessageRow[];
}

export function addMessage(conversationId: number, role: string, content: string): MessageRow {
  const db = getDb();
  const result = db
    .prepare('INSERT INTO ai_messages (conversation_id, role, content) VALUES (?, ?, ?)')
    .run(conversationId, role, content);

  // Touch conversation updated_at
  db.prepare("UPDATE ai_conversations SET updated_at = datetime('now') WHERE id = ?").run(conversationId);

  return db.prepare('SELECT * FROM ai_messages WHERE id = ?').get(result.lastInsertRowid) as MessageRow;
}

export function updateMessage(id: number, content: string): void {
  const db = getDb();
  db.prepare('UPDATE ai_messages SET content = ? WHERE id = ?').run(content, id);
}
