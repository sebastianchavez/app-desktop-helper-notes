import { getDb } from './db';

/** Cabecera persistida de una sesión de chat. */
export interface ConversationRow {
  id: number;
  title: string;
  model: string;
  assistant_id: string;
  created_at: string;
  updated_at: string;
}

/** Mensaje individual asociado a una conversación. */
export interface MessageRow {
  id: number;
  conversation_id: number;
  role: string;
  content: string;
  created_at: string;
}

/** Lista sesiones mostrando primero la modificada más recientemente. */
export function getAllConversations(): ConversationRow[] {
  const db = getDb();
  return db
    .prepare('SELECT * FROM ai_conversations ORDER BY updated_at DESC')
    .all() as ConversationRow[];
}

/** Recupera una conversación por su clave primaria. */
export function getConversationById(id: number): ConversationRow | undefined {
  const db = getDb();
  return db.prepare('SELECT * FROM ai_conversations WHERE id = ?').get(id) as ConversationRow | undefined;
}

/** Crea una sesión vinculada al modelo y perfil que originarán sus respuestas. */
export function createConversation(title: string, model: string, assistantId = 'general'): ConversationRow {
  const db = getDb();
  const result = db
    .prepare('INSERT INTO ai_conversations (title, model, assistant_id) VALUES (?, ?, ?)')
    .run(title, model, assistantId);
  return db
    .prepare('SELECT * FROM ai_conversations WHERE id = ?')
    .get(result.lastInsertRowid) as ConversationRow;
}

/** Actualiza selectivamente metadatos y renueva el orden por actividad. */
export function updateConversation(id: number, data: { title?: string; model?: string; assistant_id?: string }): ConversationRow {
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
  if (data.assistant_id !== undefined) {
    fields.push('assistant_id = ?');
    values.push(data.assistant_id);
  }

  if (fields.length > 0) {
    fields.push("updated_at = datetime('now')");
    values.push(id);
    db.prepare(`UPDATE ai_conversations SET ${fields.join(', ')} WHERE id = ?`).run(...values);
  }

  return db.prepare('SELECT * FROM ai_conversations WHERE id = ?').get(id) as ConversationRow;
}

/** Elimina mensajes y cabecera explícitamente dentro de la misma conexión. */
export function deleteConversation(id: number): void {
  const db = getDb();
  db.prepare('DELETE FROM ai_messages WHERE conversation_id = ?').run(id);
  db.prepare('DELETE FROM ai_conversations WHERE id = ?').run(id);
}

/** Devuelve el historial en orden cronológico. */
export function getMessagesByConversation(conversationId: number): MessageRow[] {
  const db = getDb();
  return db
    .prepare('SELECT * FROM ai_messages WHERE conversation_id = ? ORDER BY created_at ASC')
    .all(conversationId) as MessageRow[];
}

/** Inserta un mensaje y marca su conversación como recientemente activa. */
export function addMessage(conversationId: number, role: string, content: string): MessageRow {
  const db = getDb();
  const result = db
    .prepare('INSERT INTO ai_messages (conversation_id, role, content) VALUES (?, ?, ?)')
    .run(conversationId, role, content);

  // Touch conversation updated_at
  db.prepare("UPDATE ai_conversations SET updated_at = datetime('now') WHERE id = ?").run(conversationId);

  return db.prepare('SELECT * FROM ai_messages WHERE id = ?').get(result.lastInsertRowid) as MessageRow;
}

/** Reemplaza el contenido, usado para consolidar respuestas recibidas por streaming. */
export function updateMessage(id: number, content: string): void {
  const db = getDb();
  db.prepare('UPDATE ai_messages SET content = ? WHERE id = ?').run(content, id);
}
