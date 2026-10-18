import { Folder } from '../models/folder.model';
import { Note } from '../models/note.model';
import { AssistantRuntimeConfig } from '../models/assistant.model';

/**
 * Contrato del puente expuesto por `electron/preload.ts` en `window`.
 *
 * El renderer sólo conoce estas operaciones serializables y nunca recibe acceso
 * directo a Node, SQLite ni al sistema de archivos. Cada firma debe mantenerse
 * sincronizada con el canal IPC correspondiente del preload y del proceso main.
 */
export interface ElectronAPI {
  platform: string;
  wordClipboardExecute(command: 'copy' | 'cut' | 'paste' | 'pastePlain'): Promise<void>;
  wordClipboardWriteImage(pngBase64: string): Promise<void>;
  onWordClipboardCommand(
    callback: (command: 'copy' | 'cut' | 'paste' | 'pastePlain') => void,
  ): () => void;

  updatesGetState(): Promise<UpdateState>;
  updatesCheck(): Promise<UpdateState>;
  updatesDownload(): Promise<UpdateState>;
  updatesInstall(): Promise<{ success: boolean; error?: string }>;
  onUpdateState(callback: (state: UpdateState) => void): () => void;

  // Dialog
  dialogConfirm(options: { title: string; message: string }): Promise<boolean>;

  // Folders
  foldersGetAll(): Promise<Folder[]>;
  foldersCreate(name: string, parentId?: number | null): Promise<Folder>;
  foldersUpdate(id: number, name: string): Promise<Folder>;
  foldersMove(id: number, parentId: number | null): Promise<Folder>;
  foldersDelete(id: number): Promise<{ success: boolean }>;

  // Notes
  notesGetByFolder(folderId: number | null): Promise<Note[]>;
  notesGetById(id: number): Promise<Note | undefined>;
  notesCreate(folderId: number | null, title?: string): Promise<Note>;
  notesUpdate(
    id: number,
    data: {
      title?: string;
      content?: string;
      folder_id?: number | null;
      file_data?: string | null;
      file_type?: string | null;
    },
  ): Promise<Note>;
  notesDelete(id: number): Promise<{ success: boolean }>;
  /** Persiste el DOCX completo y devuelve metadatos sin el adjunto Base64. */
  wordSave(noteId: number, base64: string): Promise<Omit<Note, 'file_data'>>;

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
  backupExport(options: {
    folderIds: number[];
    noteIds: number[];
    format: 'json' | 'zip';
  }): Promise<{
    success: boolean;
    canceled?: boolean;
    filePath?: string;
    folders?: number;
    notes?: number;
    error?: string;
  }>;
  backupImportFiles(): Promise<{
    success: boolean;
    partial?: boolean;
    canceled?: boolean;
    folders?: number;
    notes?: number;
    sources?: number;
    errors?: string[];
    error?: string;
  }>;
  backupImportFolder(): ReturnType<ElectronAPI['backupImportFiles']>;
  fileImport(): Promise<
    { title: string; content: string; file_data?: string; file_type?: string }[]
  >;
  fileSupportedExtensions: string[];
  fileGetPath(file: File): string;
  fileImportPaths(paths: string[]): Promise<{
    files: { title: string; content: string; file_data?: string; file_type?: string }[];
    errors: string[];
  }>;
  fileTakePendingOpen(): Promise<string[]>;
  onFileOpenRequested(callback: () => void): () => void;
  fileImportJson(): Promise<
    { title: string; content: string; file_data?: string; file_type?: string }[]
  >;
  fileExportXlsx(options: {
    defaultName: string;
    htmlContent: string;
    title: string;
  }): Promise<{ success: boolean; filePath?: string }>;
  fileExportXlsxRaw(options: {
    defaultName: string;
    base64: string;
    format?: 'xlsx' | 'xls' | 'csv';
    sheetIndex?: number;
  }): Promise<{ success: boolean; filePath?: string }>;
  xlsxParseData(base64: string): Promise<{
    success: boolean;
    pivots?: import('../models/spreadsheet-pivot').SavedPivot[];
    charts?: import('../models/spreadsheet-chart').SavedChart[];
    sheets?: {
      name: string;
      data: (string | number | boolean | null)[][];
      colCount: number;
      date1904?: boolean;
      definedNames?: Record<string, string>;
      formulas: {
        row: number;
        col: number;
        formula: string;
        arrayRange?: string;
        dynamic?: boolean;
      }[];
      tables?: import('../models/spreadsheet-table').SpreadsheetTable[];
      merges?: { s: { r: number; c: number }; e: { r: number; c: number } }[];
      colWidths?: number[];
      rowHeights?: number[];
      cellStyles?: Record<string, Record<string, string>>;
      numberFormats?: Record<string, string>;
      formattedCells?: Record<string, { value: string | number | boolean | null; text: string }>;
    }[];
    error?: string;
  }>;
  xlsxChart(
    base64: string,
    action: 'preview' | 'save' | 'delete' | 'refresh',
    definition?: import('../models/spreadsheet-chart').ChartDefinition,
    id?: string,
  ): Promise<{
    success: boolean;
    base64?: string;
    data?: import('../models/spreadsheet-chart').ChartData;
    chart?: import('../models/spreadsheet-chart').SavedChart;
    warnings?: string[];
    error?: string;
  }>;
  xlsxPivot(
    base64: string,
    action: 'preview' | 'save' | 'refresh',
    definition?: import('../models/spreadsheet-pivot').PivotDefinition,
  ): Promise<{
    success: boolean;
    base64?: string;
    pivot?: import('../models/spreadsheet-pivot').SavedPivot;
    warnings?: string[];
    preview?: {
      matrix: (string | number | boolean | null)[][];
      matched: number;
      rows: number;
      columns: number;
    };
    error?: string;
  }>;
  xlsxUpdateCell(
    base64: string,
    sheetIndex: number,
    row: number,
    col: number,
    value: string | number | boolean | null,
    formula?: string | null,
  ): Promise<{
    success: boolean;
    base64?: string;
    error?: string;
  }>;
  xlsxCreateBlank(): Promise<{
    success: boolean;
    base64?: string;
    error?: string;
  }>;
  xlsxBatchUpdateCells(
    base64: string,
    sheetIndex: number,
    cells: {
      row: number;
      col: number;
      value: string | number | boolean | null;
      formula?: string | null;
    }[],
  ): Promise<{
    success: boolean;
    base64?: string;
    error?: string;
  }>;
  xlsxReplaceSheetData(
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
    options?: import('../models/spreadsheet-format').SpreadsheetEditOptions,
  ): Promise<{
    success: boolean;
    base64?: string;
    error?: string;
  }>;
  xlsxUpdateDimensions(
    base64: string,
    sheetIndex: number,
    colWidths: number[],
    rowHeights: number[],
  ): Promise<{
    success: boolean;
    base64?: string;
    error?: string;
  }>;
  fileShowInFolder(filePath: string): Promise<void>;
  filePrintToPdf(options: {
    defaultName: string;
    htmlContent: string;
  }): Promise<{ success: boolean; filePath?: string }>;
  filePrint(options: {
    title: string;
    htmlContent?: string;
    pdfBase64?: string;
    landscape?: boolean;
  }): Promise<{ success: boolean; canceled?: boolean; error?: string }>;

  // Window
  windowOpenNote(noteId: number): Promise<void>;
  windowSetFullScreen(enabled: boolean): Promise<{ success: boolean }>;

  // Ollama AI
  llamaManagerStatus(endpoint: string): Promise<{
    installed: boolean;
    running: boolean;
    executable?: string;
    models: { name: string; path: string; size: number }[];
  }>;
  llamaInstall(): Promise<{ success: boolean; path?: string; version?: string; error?: string }>;
  llamaCatalog(query: string): Promise<{
    success: boolean;
    models: { id: string; downloads: number; likes: number; updatedAt: string }[];
    error?: string;
  }>;
  llamaModelFiles(repository: string): Promise<{
    success: boolean;
    files: { name: string; size: number; quantization: string }[];
    error?: string;
  }>;
  llamaDownloadModel(
    repository: string,
    fileName: string,
  ): Promise<{ success: boolean; path?: string; error?: string }>;
  llamaCancelDownload(): Promise<{ success: boolean }>;
  llamaLocalModels(): Promise<{ name: string; path: string; size: number }[]>;
  llamaStart(modelPath: string, endpoint: string): Promise<{ success: boolean; error?: string }>;
  llamaStop(): Promise<{ success: boolean; error?: string }>;
  aiStatus(config: AssistantRuntimeConfig): Promise<{ available: boolean; error?: string }>;
  aiListModels(config: AssistantRuntimeConfig): Promise<{
    success: boolean;
    models: { name: string; size: number; modified_at: string }[];
    error?: string;
  }>;
  aiChat(
    config: AssistantRuntimeConfig,
    model: string,
    messages: { role: string; content: string }[],
  ): Promise<{ success: boolean; error?: string }>;
  aiAbort(): Promise<{ success: boolean }>;
  ollamaCheckInstalled(): Promise<{ installed: boolean }>;
  ollamaStatus(): Promise<{ running: boolean }>;
  ollamaStart(): Promise<{ success: boolean; alreadyRunning?: boolean; error?: string }>;
  ollamaStop(): Promise<{ success: boolean; error?: string }>;
  ollamaOpenDownload(): Promise<{ success: boolean }>;
  ollamaListModels(): Promise<{
    success: boolean;
    models?: { name: string; size: number; modified_at: string }[];
    error?: string;
  }>;
  ollamaPull(model: string): Promise<{ success: boolean; error?: string }>;
  ollamaSignin(): Promise<{ success: boolean; error?: string }>;
  ollamaChat(
    model: string,
    messages: { role: string; content: string }[],
  ): Promise<{ success: boolean; error?: string }>;
  ollamaAbort(): Promise<void>;

  // AI Conversations
  conversationsGetAll(): Promise<
    {
      id: number;
      title: string;
      model: string;
      assistant_id: string;
      created_at: string;
      updated_at: string;
    }[]
  >;
  conversationsGetById(id: number): Promise<
    | {
        id: number;
        title: string;
        model: string;
        assistant_id: string;
        created_at: string;
        updated_at: string;
      }
    | undefined
  >;
  conversationsCreate(
    title: string,
    model: string,
    assistantId?: string,
  ): Promise<{
    id: number;
    title: string;
    model: string;
    assistant_id: string;
    created_at: string;
    updated_at: string;
  }>;
  conversationsUpdate(
    id: number,
    data: { title?: string; model?: string; assistant_id?: string },
  ): Promise<{
    id: number;
    title: string;
    model: string;
    assistant_id: string;
    created_at: string;
    updated_at: string;
  }>;
  conversationsDelete(id: number): Promise<{ success: boolean }>;
  conversationsGetMessages(
    conversationId: number,
  ): Promise<
    { id: number; conversation_id: number; role: string; content: string; created_at: string }[]
  >;
  conversationsAddMessage(
    conversationId: number,
    role: string,
    content: string,
  ): Promise<{
    id: number;
    conversation_id: number;
    role: string;
    content: string;
    created_at: string;
  }>;
  conversationsUpdateMessage(id: number, content: string): Promise<{ success: boolean }>;

  onOllamaToken(callback: (token: string) => void): () => void;
  onOllamaDone(callback: () => void): () => void;
  onOllamaError(callback: (error: string) => void): () => void;
  onOllamaPullProgress(
    callback: (progress: {
      model: string;
      status: string;
      completed: number;
      total: number;
      percent: number;
    }) => void,
  ): () => void;
  onAiToken(callback: (token: string) => void): () => void;
  onAiDone(callback: () => void): () => void;
  onAiError(callback: (error: string) => void): () => void;
  onLlamaProgress(
    callback: (progress: {
      operation: 'runtime' | 'model';
      id: string;
      received: number;
      total: number;
      percent: number;
      message: string;
    }) => void,
  ): () => void;

  // Events
  onNoteChanged(callback: (noteId: number) => void): () => void;
}

export interface UpdateState {
  status: 'idle' | 'checking' | 'available' | 'downloading' | 'downloaded' | 'up-to-date' | 'error';
  currentVersion: string;
  update?: {
    version: string;
    releaseNotes: string;
    downloadUrl: string;
    checksumSha256: string;
    fileName: string;
    fileSize: number;
    mandatory: boolean;
  };
  progress?: number;
  filePath?: string;
  error?: string;
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}
