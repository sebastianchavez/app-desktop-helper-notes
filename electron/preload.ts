import { contextBridge, ipcRenderer } from 'electron';

/**
 * Superficie segura entre el renderer y el proceso principal.
 *
 * Cada método traduce una operación de dominio a un canal IPC explícito. Los
 * listeners devuelven una función de limpieza para que Angular pueda evitar
 * fugas al destruir componentes. Este objeto debe conservar paridad con la
 * interfaz `ElectronAPI` declarada en el renderer.
 */
contextBridge.exposeInMainWorld('electronAPI', {
  platform: process.platform,

  // ── Updates ──
  updatesGetState: () => ipcRenderer.invoke('updates:getState'),
  updatesCheck: () => ipcRenderer.invoke('updates:check'),
  updatesDownload: () => ipcRenderer.invoke('updates:download'),
  updatesInstall: () => ipcRenderer.invoke('updates:install'),
  onUpdateState: (callback: (state: unknown) => void) => {
    const handler = (_event: Electron.IpcRendererEvent, state: unknown) => callback(state);
    ipcRenderer.on('updates:state', handler);
    return () => ipcRenderer.removeListener('updates:state', handler);
  },

  // ── Dialog ──
  dialogConfirm: (options: { title: string; message: string }) =>
    ipcRenderer.invoke('dialog:confirm', options),

  // ── Folders ──
  foldersGetAll: () => ipcRenderer.invoke('folders:getAll'),
  foldersCreate: (name: string, parentId: number | null = null) =>
    ipcRenderer.invoke('folders:create', name, parentId),
  foldersUpdate: (id: number, name: string) => ipcRenderer.invoke('folders:update', id, name),
  foldersDelete: (id: number) => ipcRenderer.invoke('folders:delete', id),

  // ── Notes ──
  notesGetByFolder: (folderId: number | null) => ipcRenderer.invoke('notes:getByFolder', folderId),
  notesGetById: (id: number) => ipcRenderer.invoke('notes:getById', id),
  notesCreate: (folderId: number | null, title?: string) =>
    ipcRenderer.invoke('notes:create', folderId, title),
  notesUpdate: (
    id: number,
    data: {
      title?: string;
      content?: string;
      folder_id?: number | null;
      file_data?: string | null;
      file_type?: string | null;
    },
  ) => ipcRenderer.invoke('notes:update', id, data),
  notesDelete: (id: number) => ipcRenderer.invoke('notes:delete', id),

  // ── Settings ──
  settingsGet: (key: string) => ipcRenderer.invoke('settings:get', key),
  settingsSet: (key: string, value: string) => ipcRenderer.invoke('settings:set', key, value),
  settingsGetAll: () => ipcRenderer.invoke('settings:getAll'),

  // ── File Export ──
  fileSaveAs: (options: {
    defaultName: string;
    filters: { name: string; extensions: string[] }[];
    content: string;
    encoding?: string;
  }) => ipcRenderer.invoke('file:saveAs', options),
  fileSaveAsBuffer: (options: {
    defaultName: string;
    filters: { name: string; extensions: string[] }[];
    buffer: number[];
  }) => ipcRenderer.invoke('file:saveAsBuffer', options),
  backupExport: (options: { folderIds: number[]; noteIds: number[]; format: 'json' | 'zip' }) =>
    ipcRenderer.invoke('backup:export', options),
  backupImportFiles: () => ipcRenderer.invoke('backup:importFiles'),
  backupImportFolder: () => ipcRenderer.invoke('backup:importFolder'),
  fileImport: () => ipcRenderer.invoke('file:import'),
  fileImportJson: () => ipcRenderer.invoke('file:importJson'),
  fileExportXlsx: (options: { defaultName: string; htmlContent: string; title: string }) =>
    ipcRenderer.invoke('file:exportXlsx', options),
  fileExportXlsxRaw: (options: {
    defaultName: string;
    base64: string;
    format?: 'xlsx' | 'xls' | 'csv';
    sheetIndex?: number;
  }) => ipcRenderer.invoke('file:exportXlsxRaw', options),
  xlsxParseData: (base64: string) => ipcRenderer.invoke('xlsx:parseData', base64),
  xlsxUpdateCell: (
    base64: string,
    sheetIndex: number,
    row: number,
    col: number,
    value: string | number | boolean | null,
    formula?: string | null,
  ) => ipcRenderer.invoke('xlsx:updateCell', base64, sheetIndex, row, col, value, formula),
  xlsxCreateBlank: () => ipcRenderer.invoke('xlsx:createBlank'),
  xlsxBatchUpdateCells: (
    base64: string,
    sheetIndex: number,
    cells: {
      row: number;
      col: number;
      value: string | number | boolean | null;
      formula?: string | null;
    }[],
  ) => ipcRenderer.invoke('xlsx:batchUpdateCells', base64, sheetIndex, cells),
  xlsxReplaceSheetData: (
    base64: string,
    sheetIndex: number,
    data: (string | number | boolean | null)[][],
    formulas: {
      row: number;
      col: number;
      value: string | number | boolean | null;
      formula: string;
    }[],
    merges: { s: { r: number; c: number }; e: { r: number; c: number } }[],
  ) => ipcRenderer.invoke('xlsx:replaceSheetData', base64, sheetIndex, data, formulas, merges),
  xlsxUpdateDimensions: (
    base64: string,
    sheetIndex: number,
    colWidths: number[],
    rowHeights: number[],
  ) => ipcRenderer.invoke('xlsx:updateDimensions', base64, sheetIndex, colWidths, rowHeights),
  fileShowInFolder: (filePath: string) => ipcRenderer.invoke('file:showInFolder', filePath),
  filePrintToPdf: (options: { defaultName: string; htmlContent: string }) =>
    ipcRenderer.invoke('file:printToPdf', options),

  // ── Window ──
  windowOpenNote: (noteId: number) => ipcRenderer.invoke('window:openNote', noteId),
  windowSetFullScreen: (enabled: boolean) => ipcRenderer.invoke('window:setFullScreen', enabled),

  // ── Ollama AI ──
  llamaManagerStatus: (endpoint: string) => ipcRenderer.invoke('llama:managerStatus', endpoint),
  llamaInstall: () => ipcRenderer.invoke('llama:install'),
  llamaCatalog: (query: string) => ipcRenderer.invoke('llama:catalog', query),
  llamaModelFiles: (repository: string) => ipcRenderer.invoke('llama:modelFiles', repository),
  llamaDownloadModel: (repository: string, fileName: string) =>
    ipcRenderer.invoke('llama:downloadModel', repository, fileName),
  llamaCancelDownload: () => ipcRenderer.invoke('llama:cancelDownload'),
  llamaLocalModels: () => ipcRenderer.invoke('llama:localModels'),
  llamaStart: (modelPath: string, endpoint: string) =>
    ipcRenderer.invoke('llama:start', modelPath, endpoint),
  llamaStop: () => ipcRenderer.invoke('llama:stop'),
  aiStatus: (config: { provider: 'ollama' | 'llamacpp'; endpoint: string; temperature: number }) =>
    ipcRenderer.invoke('ai:status', config),
  aiListModels: (config: {
    provider: 'ollama' | 'llamacpp';
    endpoint: string;
    temperature: number;
  }) => ipcRenderer.invoke('ai:listModels', config),
  aiChat: (
    config: { provider: 'ollama' | 'llamacpp'; endpoint: string; temperature: number },
    model: string,
    messages: { role: string; content: string }[],
  ) => ipcRenderer.invoke('ai:chat', config, model, messages),
  aiAbort: () => ipcRenderer.invoke('ai:abort'),
  ollamaCheckInstalled: () => ipcRenderer.invoke('ollama:checkInstalled'),
  ollamaStatus: () => ipcRenderer.invoke('ollama:status'),
  ollamaStart: () => ipcRenderer.invoke('ollama:start'),
  ollamaStop: () => ipcRenderer.invoke('ollama:stop'),
  ollamaOpenDownload: () => ipcRenderer.invoke('ollama:openDownload'),
  ollamaListModels: () => ipcRenderer.invoke('ollama:listModels'),
  ollamaPull: (model: string) => ipcRenderer.invoke('ollama:pull', model),
  ollamaSignin: () => ipcRenderer.invoke('ollama:signin'),
  ollamaChat: (model: string, messages: { role: string; content: string }[]) =>
    ipcRenderer.invoke('ollama:chat', model, messages),
  ollamaAbort: () => ipcRenderer.invoke('ollama:abort'),

  // ── AI Conversations ──
  conversationsGetAll: () => ipcRenderer.invoke('conversations:getAll'),
  conversationsGetById: (id: number) => ipcRenderer.invoke('conversations:getById', id),
  conversationsCreate: (title: string, model: string, assistantId?: string) =>
    ipcRenderer.invoke('conversations:create', title, model, assistantId),
  conversationsUpdate: (
    id: number,
    data: { title?: string; model?: string; assistant_id?: string },
  ) => ipcRenderer.invoke('conversations:update', id, data),
  conversationsDelete: (id: number) => ipcRenderer.invoke('conversations:delete', id),
  conversationsGetMessages: (conversationId: number) =>
    ipcRenderer.invoke('conversations:getMessages', conversationId),
  conversationsAddMessage: (conversationId: number, role: string, content: string) =>
    ipcRenderer.invoke('conversations:addMessage', conversationId, role, content),
  conversationsUpdateMessage: (id: number, content: string) =>
    ipcRenderer.invoke('conversations:updateMessage', id, content),

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
  onOllamaPullProgress: (
    callback: (progress: {
      model: string;
      status: string;
      completed: number;
      total: number;
      percent: number;
    }) => void,
  ) => {
    const handler = (
      _e: Electron.IpcRendererEvent,
      progress: {
        model: string;
        status: string;
        completed: number;
        total: number;
        percent: number;
      },
    ) => callback(progress);
    ipcRenderer.on('ollama:pullProgress', handler);
    return () => ipcRenderer.removeListener('ollama:pullProgress', handler);
  },
  onAiToken: (callback: (token: string) => void) => {
    const handler = (_e: Electron.IpcRendererEvent, token: string) => callback(token);
    ipcRenderer.on('ai:token', handler);
    return () => ipcRenderer.removeListener('ai:token', handler);
  },
  onAiDone: (callback: () => void) => {
    const handler = () => callback();
    ipcRenderer.on('ai:done', handler);
    return () => ipcRenderer.removeListener('ai:done', handler);
  },
  onAiError: (callback: (error: string) => void) => {
    const handler = (_e: Electron.IpcRendererEvent, error: string) => callback(error);
    ipcRenderer.on('ai:error', handler);
    return () => ipcRenderer.removeListener('ai:error', handler);
  },
  onLlamaProgress: (
    callback: (progress: {
      operation: 'runtime' | 'model';
      id: string;
      received: number;
      total: number;
      percent: number;
      message: string;
    }) => void,
  ) => {
    const handler = (
      _e: Electron.IpcRendererEvent,
      progress: {
        operation: 'runtime' | 'model';
        id: string;
        received: number;
        total: number;
        percent: number;
        message: string;
      },
    ) => callback(progress);
    ipcRenderer.on('llama:progress', handler);
    return () => ipcRenderer.removeListener('llama:progress', handler);
  },

  // ── Events ──
  onNoteChanged: (callback: (noteId: number) => void) => {
    const handler = (_event: Electron.IpcRendererEvent, noteId: number) => callback(noteId);
    ipcRenderer.on('note:changed', handler);
    return () => ipcRenderer.removeListener('note:changed', handler);
  },
});
