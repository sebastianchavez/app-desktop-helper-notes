import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('electronAPI', {
  platform: process.platform,

  // ── Dialog ──
  dialogConfirm: (options: { title: string; message: string }) =>
    ipcRenderer.invoke('dialog:confirm', options),

  // ── Folders ──
  foldersGetAll: () => ipcRenderer.invoke('folders:getAll'),
  foldersCreate: (name: string) => ipcRenderer.invoke('folders:create', name),
  foldersUpdate: (id: number, name: string) => ipcRenderer.invoke('folders:update', id, name),
  foldersDelete: (id: number) => ipcRenderer.invoke('folders:delete', id),

  // ── Notes ──
  notesGetByFolder: (folderId: number | null) => ipcRenderer.invoke('notes:getByFolder', folderId),
  notesGetById: (id: number) => ipcRenderer.invoke('notes:getById', id),
  notesCreate: (folderId: number | null, title?: string) => ipcRenderer.invoke('notes:create', folderId, title),
  notesUpdate: (id: number, data: { title?: string; content?: string; folder_id?: number | null; file_data?: string | null; file_type?: string | null }) =>
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
  fileImport: () => ipcRenderer.invoke('file:import'),
  fileShowInFolder: (filePath: string) => ipcRenderer.invoke('file:showInFolder', filePath),
  filePrintToPdf: (options: { defaultName: string; htmlContent: string }) =>
    ipcRenderer.invoke('file:printToPdf', options),

  // ── Window ──
  windowOpenNote: (noteId: number) => ipcRenderer.invoke('window:openNote', noteId),

  // ── Ollama AI ──
  ollamaCheckInstalled: () => ipcRenderer.invoke('ollama:checkInstalled'),
  ollamaStatus: () => ipcRenderer.invoke('ollama:status'),
  ollamaStart: () => ipcRenderer.invoke('ollama:start'),
  ollamaStop: () => ipcRenderer.invoke('ollama:stop'),
  ollamaOpenDownload: () => ipcRenderer.invoke('ollama:openDownload'),
  ollamaListModels: () => ipcRenderer.invoke('ollama:listModels'),
  ollamaChat: (model: string, messages: { role: string; content: string }[]) =>
    ipcRenderer.invoke('ollama:chat', model, messages),
  ollamaAbort: () => ipcRenderer.invoke('ollama:abort'),

  // ── AI Conversations ──
  conversationsGetAll: () => ipcRenderer.invoke('conversations:getAll'),
  conversationsGetById: (id: number) => ipcRenderer.invoke('conversations:getById', id),
  conversationsCreate: (title: string, model: string) => ipcRenderer.invoke('conversations:create', title, model),
  conversationsUpdate: (id: number, data: { title?: string; model?: string }) => ipcRenderer.invoke('conversations:update', id, data),
  conversationsDelete: (id: number) => ipcRenderer.invoke('conversations:delete', id),
  conversationsGetMessages: (conversationId: number) => ipcRenderer.invoke('conversations:getMessages', conversationId),
  conversationsAddMessage: (conversationId: number, role: string, content: string) => ipcRenderer.invoke('conversations:addMessage', conversationId, role, content),
  conversationsUpdateMessage: (id: number, content: string) => ipcRenderer.invoke('conversations:updateMessage', id, content),

  onOllamaToken: (callback: (token: string) => void) => {
    const handler = (_e: Electron.IpcRendererEvent, token: string) => callback(token);
    ipcRenderer.on('ollama:token', handler);
    return () => ipcRenderer.removeListener('ollama:token', handler);
  },
  onOllamaDone: (callback: () => void) => {
    const handler = () => callback();
    ipcRenderer.on('ollama:done', handler);
    return () => ipcRenderer.removeListener('ollama:done', handler);
  },
  onOllamaError: (callback: (error: string) => void) => {
    const handler = (_e: Electron.IpcRendererEvent, error: string) => callback(error);
    ipcRenderer.on('ollama:error', handler);
    return () => ipcRenderer.removeListener('ollama:error', handler);
  },

  // ── Events ──
  onNoteChanged: (callback: (noteId: number) => void) => {
    const handler = (_event: Electron.IpcRendererEvent, noteId: number) => callback(noteId);
    ipcRenderer.on('note:changed', handler);
    return () => ipcRenderer.removeListener('note:changed', handler);
  },
});
