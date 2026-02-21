import { Folder } from '../models/folder.model';
import { Note } from '../models/note.model';

export interface ElectronAPI {
  platform: string;

  // Folders
  foldersGetAll(): Promise<Folder[]>;
  foldersCreate(name: string): Promise<Folder>;
  foldersUpdate(id: number, name: string): Promise<Folder>;
  foldersDelete(id: number): Promise<{ success: boolean }>;

  // Notes
  notesGetByFolder(folderId: number | null): Promise<Note[]>;
  notesGetById(id: number): Promise<Note | undefined>;
  notesCreate(folderId: number | null, title?: string): Promise<Note>;
  notesUpdate(id: number, data: { title?: string; content?: string; folder_id?: number | null }): Promise<Note>;
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
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}
