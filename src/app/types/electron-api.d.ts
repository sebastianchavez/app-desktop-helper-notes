import { Folder } from '../models/folder.model';
import { Note } from '../models/note.model';

export interface ElectronAPI {
  platform: string;

  // Dialog
  dialogConfirm(options: { title: string; message: string }): Promise<boolean>;

  // Folders
  foldersGetAll(): Promise<Folder[]>;
  foldersCreate(name: string): Promise<Folder>;
  foldersUpdate(id: number, name: string): Promise<Folder>;
  foldersDelete(id: number): Promise<{ success: boolean }>;

  // Notes
  notesGetByFolder(folderId: number | null): Promise<Note[]>;
  notesGetById(id: number): Promise<Note | undefined>;
  notesCreate(folderId: number | null, title?: string): Promise<Note>;
  notesUpdate(id: number, data: { title?: string; content?: string; folder_id?: number | null; file_data?: string | null; file_type?: string | null }): Promise<Note>;
  notesDelete(id: number): Promise<{ success: boolean }>;

  // Settings
  settingsGet(key: string): Promise<string | undefined>;
  settingsSet(key: string, value: string): Promise<{ success: boolean }>;
  settingsGetAll(): Promise<{ key: string; value: string }[]>;

  // File Export
  fileSaveAs(options: {
    defaultName: string;
    filters: { name: string; extensions: string[] }[];
    content: string;
    encoding?: string;
  }): Promise<{ success: boolean; filePath?: string }>;
  fileSaveAsBuffer(options: {
    defaultName: string;
    filters: { name: string; extensions: string[] }[];
    buffer: number[];
  }): Promise<{ success: boolean; filePath?: string }>;
  fileImport(): Promise<{ title: string; content: string; file_data?: string; file_type?: string }[]>;
  fileShowInFolder(filePath: string): Promise<void>;
  filePrintToPdf(options: {
    defaultName: string;
    htmlContent: string;
  }): Promise<{ success: boolean; filePath?: string }>;

  // Window
  windowOpenNote(noteId: number): Promise<void>;

  // Ollama AI
  ollamaCheckInstalled(): Promise<{ installed: boolean }>;
  ollamaStatus(): Promise<{ running: boolean }>;
  ollamaStart(): Promise<{ success: boolean; alreadyRunning?: boolean; error?: string }>;
  ollamaStop(): Promise<{ success: boolean }>;
  ollamaOpenDownload(): Promise<{ success: boolean }>;
  ollamaListModels(): Promise<{ success: boolean; models?: { name: string; size: number; modified_at: string }[]; error?: string }>;
  ollamaChat(model: string, messages: { role: string; content: string }[]): Promise<{ success: boolean; error?: string }>;
  ollamaAbort(): Promise<void>;

  // AI Conversations
  conversationsGetAll(): Promise<{ id: number; title: string; model: string; created_at: string; updated_at: string }[]>;
  conversationsGetById(id: number): Promise<{ id: number; title: string; model: string; created_at: string; updated_at: string } | undefined>;
  conversationsCreate(title: string, model: string): Promise<{ id: number; title: string; model: string; created_at: string; updated_at: string }>;
  conversationsUpdate(id: number, data: { title?: string; model?: string }): Promise<{ id: number; title: string; model: string; created_at: string; updated_at: string }>;
  conversationsDelete(id: number): Promise<{ success: boolean }>;
  conversationsGetMessages(conversationId: number): Promise<{ id: number; conversation_id: number; role: string; content: string; created_at: string }[]>;
  conversationsAddMessage(conversationId: number, role: string, content: string): Promise<{ id: number; conversation_id: number; role: string; content: string; created_at: string }>;
  conversationsUpdateMessage(id: number, content: string): Promise<{ success: boolean }>;

  onOllamaToken(callback: (token: string) => void): () => void;
  onOllamaDone(callback: () => void): () => void;
  onOllamaError(callback: (error: string) => void): () => void;

  // Events
  onNoteChanged(callback: (noteId: number) => void): () => void;
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}
