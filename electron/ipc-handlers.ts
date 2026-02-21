import { ipcMain, dialog, BrowserWindow } from 'electron';
import { writeFileSync } from 'fs';
import { getAllFolders, createFolder, updateFolder, deleteFolder } from './db/folder.repository';
import { getNotesByFolder, getNoteById, createNote, updateNote, deleteNote } from './db/note.repository';
import { getSetting, setSetting, getAllSettings } from './db/settings.repository';

export function registerIpcHandlers() {
  // ── Folders ──
  ipcMain.handle('folders:getAll', () => {
    return getAllFolders();
  });

  ipcMain.handle('folders:create', (_event, name: string) => {
    return createFolder(name);
  });

  ipcMain.handle('folders:update', (_event, id: number, name: string) => {
    return updateFolder(id, name);
  });

  ipcMain.handle('folders:delete', (_event, id: number) => {
    deleteFolder(id);
    return { success: true };
  });

  // ── Notes ──
  ipcMain.handle('notes:getByFolder', (_event, folderId: number | null) => {
    return getNotesByFolder(folderId);
  });

  ipcMain.handle('notes:getById', (_event, id: number) => {
    return getNoteById(id);
  });

  ipcMain.handle('notes:create', (_event, folderId: number | null, title?: string) => {
    return createNote(folderId, title);
  });

  ipcMain.handle('notes:update', (_event, id: number, data: { title?: string; content?: string; folder_id?: number | null }) => {
    return updateNote(id, data);
  });

  ipcMain.handle('notes:delete', (_event, id: number) => {
    deleteNote(id);
    return { success: true };
  });

  // ── Settings ──
  ipcMain.handle('settings:get', (_event, key: string) => {
    return getSetting(key);
  });

  ipcMain.handle('settings:set', (_event, key: string, value: string) => {
    setSetting(key, value);
    return { success: true };
  });

  ipcMain.handle('settings:getAll', () => {
    return getAllSettings();
  });

  // ── File Export ──
  ipcMain.handle('file:saveAs', async (_event, options: { defaultName: string; filters: { name: string; extensions: string[] }[]; content: string; encoding?: string }) => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showSaveDialog(win!, {
      defaultPath: options.defaultName,
      filters: options.filters,
    });
    if (result.canceled || !result.filePath) return { success: false };
    writeFileSync(result.filePath, options.content, options.encoding as BufferEncoding ?? 'utf-8');
    return { success: true, filePath: result.filePath };
  });

  ipcMain.handle('file:saveAsBuffer', async (_event, options: { defaultName: string; filters: { name: string; extensions: string[] }[]; buffer: number[] }) => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showSaveDialog(win!, {
      defaultPath: options.defaultName,
      filters: options.filters,
    });
    if (result.canceled || !result.filePath) return { success: false };
    writeFileSync(result.filePath, Buffer.from(options.buffer));
    return { success: true, filePath: result.filePath };
  });
}
