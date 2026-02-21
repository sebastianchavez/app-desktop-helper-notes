import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('electronAPI', {
  platform: process.platform,

  // ── Folders ──
  foldersGetAll: () => ipcRenderer.invoke('folders:getAll'),
  foldersCreate: (name: string) => ipcRenderer.invoke('folders:create', name),
  foldersUpdate: (id: number, name: string) => ipcRenderer.invoke('folders:update', id, name),
  foldersDelete: (id: number) => ipcRenderer.invoke('folders:delete', id),

  // ── Notes ──
  notesGetByFolder: (folderId: number | null) => ipcRenderer.invoke('notes:getByFolder', folderId),
  notesGetById: (id: number) => ipcRenderer.invoke('notes:getById', id),
  notesCreate: (folderId: number | null, title?: string) => ipcRenderer.invoke('notes:create', folderId, title),
  notesUpdate: (id: number, data: { title?: string; content?: string; folder_id?: number | null }) =>
    ipcRenderer.invoke('notes:update', id, data),
  notesDelete: (id: number) => ipcRenderer.invoke('notes:delete', id),

  // ── Settings ──
  settingsGet: (key: string) => ipcRenderer.invoke('settings:get', key),
  settingsSet: (key: string, value: string) => ipcRenderer.invoke('settings:set', key, value),
  settingsGetAll: () => ipcRenderer.invoke('settings:getAll'),

  // ── File Export ──
  fileSaveAs: (options: { defaultName: string; filters: { name: string; extensions: string[] }[]; content: string; encoding?: string }) =>
    ipcRenderer.invoke('file:saveAs', options),
  fileSaveAsBuffer: (options: { defaultName: string; filters: { name: string; extensions: string[] }[]; buffer: number[] }) =>
    ipcRenderer.invoke('file:saveAsBuffer', options),
});
