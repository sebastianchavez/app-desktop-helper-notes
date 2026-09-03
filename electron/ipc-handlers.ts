import { ipcMain, dialog, BrowserWindow, shell } from 'electron';
import { readFileSync, writeFileSync, existsSync } from 'fs';
import { execSync, spawn, ChildProcess } from 'child_process';
import http from 'http';
import path from 'path';
import { inflateRawSync } from 'zlib';
import * as XLSX from 'xlsx';
import { getAllFolders, createFolder, updateFolder, deleteFolder } from './db/folder.repository';
import { getNotesByFolder, getNoteById, createNote, updateNote, deleteNote } from './db/note.repository';
import { getSetting, setSetting, getAllSettings } from './db/settings.repository';
import {
  getAllConversations,
  getConversationById,
  createConversation,
  updateConversation,
  deleteConversation,
  getMessagesByConversation,
  addMessage,
  updateMessage,
} from './db/conversation.repository';
import { abortAiChat, checkAiProvider, listAiModels, streamAiChat, AiRuntimeConfig } from './ai-provider';
import {
  cancelLlamaDownload,
  downloadLlamaModel,
  installLlamaCpp,
  listLlamaModelFiles,
  listLocalLlamaModels,
  llamaManagerStatus,
  searchLlamaModels,
  startManagedLlamaServer,
  stopManagedLlamaServer,
} from './llama-manager';
import {
  backupToJson,
  backupToZip,
  createBackup,
  findBackupFiles,
  importBackupPaths,
} from './backup-manager';
import { parsePptxBuffer } from './presentation-parser';

// Families that are NOT conversational (embedding, vision-only, code-only, etc.)
const NON_CHAT_PATTERNS = /\b(embed|embedding|all-minilm|bge-|nomic-embed|mxbae|snowflake-arctic-embed)\b/i;

type XlsxCellValue = string | number | boolean | null;
type XlsxCellUpdate = { row: number; col: number; value: XlsxCellValue; formula?: string | null };

/**
 * Aplica valor o fórmula a una celda y elimina metadatos incompatibles del
 * estado anterior. La función se comparte entre cambios individuales y lotes.
 */
export function setXlsxCell(sheet: XLSX.WorkSheet, update: XlsxCellUpdate): void {
  const cellRef = XLSX.utils.encode_cell({ r: update.row, c: update.col });
  const formula = update.formula?.trim().replace(/^=/, '').slice(0, 8192);
  if ((update.value === null || update.value === '') && !formula) {
    delete sheet[cellRef];
    return;
  }

  const value = update.value ?? 0;
  const cell: XLSX.CellObject = typeof value === 'number'
    ? { t: 'n', v: value }
    : typeof value === 'boolean'
      ? { t: 'b', v: value }
      : { t: 's', v: String(value) };
  if (formula) cell.f = formula;
  sheet[cellRef] = cell;

  const range = XLSX.utils.decode_range(sheet['!ref'] || cellRef);
  range.s.r = Math.min(range.s.r, update.row);
  range.s.c = Math.min(range.s.c, update.col);
  range.e.r = Math.max(range.e.r, update.row);
  range.e.c = Math.max(range.e.c, update.col);
  sheet['!ref'] = XLSX.utils.encode_range(range);
}

/** Serializa un libro XLSX completo como Base64 para almacenarlo en SQLite. */
export function workbookToBase64(workbook: XLSX.WorkBook): string {
  (workbook as XLSX.WorkBook & { CalcPr?: Record<string, string> }).CalcPr = {
    calcMode: 'auto',
    fullCalcOnLoad: '1',
    forceFullCalc: '1',
  };
  const newBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  return Buffer.from(newBuffer).toString('base64');
}

/**
 * Registra la capa de aplicación disponible para el preload.
 *
 * Este es el límite privilegiado del programa: valida entradas del renderer,
 * coordina repositorios, diálogos, archivos, ventanas y procesos de IA, y
 * devuelve únicamente datos serializables. Debe ejecutarse una sola vez.
 */
export function registerIpcHandlers() {
  // ── Dialog ──
  ipcMain.handle('dialog:confirm', async (_event, options: { title: string; message: string }) => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showMessageBox(win!, {
      type: 'warning',
      buttons: ['Cancelar', 'Eliminar'],
      defaultId: 0,
      cancelId: 0,
      title: options.title,
      message: options.message,
    });
    return result.response === 1;
  });

  // ── Folders ──
  ipcMain.handle('folders:getAll', () => {
    return getAllFolders();
  });

  ipcMain.handle('folders:create', (_event, name: string, parentId: number | null = null) => {
    return createFolder(name, parentId);
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

  ipcMain.handle('notes:update', (event, id: number, data: { title?: string; content?: string; folder_id?: number | null; file_data?: string | null; file_type?: string | null }) => {
    const result = updateNote(id, data);
    // Broadcast to all OTHER windows so they refresh
    const senderWebContentsId = event.sender.id;
    for (const win of BrowserWindow.getAllWindows()) {
      if (win.webContents.id !== senderWebContentsId && !win.isDestroyed()) {
        win.webContents.send('note:changed', id);
      }
    }
    return result;
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

  // ── Structured backup import/export ──
  ipcMain.handle('backup:export', async (_event, options: { folderIds: number[]; noteIds: number[]; format: 'json' | 'zip' }) => {
    try {
      const backup = createBackup({ folderIds: options.folderIds ?? [], noteIds: options.noteIds ?? [] });
      const extension = options.format === 'zip' ? 'zip' : 'json';
      const selectedCount = backup.folders.length + backup.notes.length;
      const baseName = selectedCount === 1 && backup.folders.length === 1
        ? backup.folders[0].name.replace(/[<>:"/\\|?*]/g, '_').trim() || 'carpeta'
        : 'helper-notes-export';
      const win = BrowserWindow.getFocusedWindow();
      const result = await dialog.showSaveDialog(win!, {
        defaultPath: `${baseName}.${extension}`,
        filters: [{ name: options.format === 'zip' ? 'Archivo ZIP' : 'Respaldo JSON', extensions: [extension] }],
      });
      if (result.canceled || !result.filePath) return { success: false, canceled: true };
      const filePath = result.filePath.toLowerCase().endsWith(`.${extension}`) ? result.filePath : `${result.filePath}.${extension}`;
      if (options.format === 'zip') writeFileSync(filePath, Buffer.from(await backupToZip(backup)));
      else writeFileSync(filePath, backupToJson(backup), 'utf8');
      return {
        success: true,
        filePath,
        folders: backup.folders.length,
        notes: backup.notes.length + backup.folders.reduce((total, folder) => total + folder.notes.length, 0),
      };
    } catch (error) {
      return { success: false, error: error instanceof Error ? error.message : String(error) };
    }
  });

  ipcMain.handle('backup:importFiles', async () => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showOpenDialog(win!, {
      title: 'Importar respaldo de Helper Notes',
      filters: [
        { name: 'Respaldos de Helper Notes', extensions: ['json', 'zip'] },
        { name: 'JSON', extensions: ['json'] },
        { name: 'ZIP', extensions: ['zip'] },
      ],
      properties: ['openFile', 'multiSelections'],
    });
    if (result.canceled || result.filePaths.length === 0) return { success: false, canceled: true };
    const summary = await importBackupPaths(result.filePaths);
    return {
      success: summary.sources > 0 && summary.errors.length === 0,
      partial: summary.folders + summary.notes > 0 && summary.errors.length > 0,
      ...summary,
    };
  });

  ipcMain.handle('backup:importFolder', async () => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showOpenDialog(win!, {
      title: 'Seleccionar carpeta con respaldos',
      properties: ['openDirectory'],
    });
    if (result.canceled || result.filePaths.length === 0) return { success: false, canceled: true };
    try {
      const files = findBackupFiles(result.filePaths[0]);
      if (files.length === 0) return { success: false, error: 'La carpeta no contiene archivos JSON o ZIP compatibles.' };
      const summary = await importBackupPaths(files);
      return {
        success: summary.sources > 0 && summary.errors.length === 0,
        partial: summary.folders + summary.notes > 0 && summary.errors.length > 0,
        ...summary,
      };
    } catch (error) {
      return { success: false, error: error instanceof Error ? error.message : String(error) };
    }
  });

  // ── File Import ──
  ipcMain.handle('file:import', async () => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showOpenDialog(win!, {
      title: 'Importar archivo',
      filters: [
        { name: 'Documentos', extensions: ['txt', 'docx', 'pdf', 'html', 'htm', 'xlsx', 'xls', 'csv', 'pptx'] },
        { name: 'Texto', extensions: ['txt'] },
        { name: 'HTML', extensions: ['html', 'htm'] },
        { name: 'Word', extensions: ['docx'] },
        { name: 'PDF', extensions: ['pdf'] },
        { name: 'Excel', extensions: ['xlsx', 'xls', 'csv'] },
        { name: 'PowerPoint', extensions: ['pptx'] },
      ],
      properties: ['openFile', 'multiSelections'],
    });
    if (result.canceled || result.filePaths.length === 0) return [];

    const imported: { title: string; content: string; file_data?: string; file_type?: string }[] = [];

    for (const filePath of result.filePaths) {
      const ext = path.extname(filePath).toLowerCase();
      const baseName = path.basename(filePath, ext);

      if (ext === '.txt') {
        const text = readFileSync(filePath, 'utf-8');
        // Convert plain text lines to HTML paragraphs
        const html = text
          .split(/\r?\n/)
          .map((line) => `<p>${escapeHtml(line) || '<br>'}</p>`)
          .join('');
        imported.push({ title: baseName, content: html });
      } else if (ext === '.html' || ext === '.htm') {
        try {
          const rawHtml = readFileSync(filePath, 'utf-8');
          // Extract <body> content if present, otherwise use the full HTML
          const bodyMatch = rawHtml.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
          const content = bodyMatch ? bodyMatch[1].trim() : rawHtml.trim();
          imported.push({ title: baseName, content });
        } catch {
          imported.push({ title: baseName, content: '<p>Error al importar el archivo .html</p>' });
        }
      } else if (ext === '.docx') {
        try {
          const html = parseDocxToHtml(readFileSync(filePath));
          imported.push({ title: baseName, content: html });
        } catch {
          // If parsing fails, import as empty note with the filename
          imported.push({ title: baseName, content: '<p>Error al importar el archivo .docx</p>' });
        }
      } else if (ext === '.pdf') {
        try {
          const pdfBuffer = readFileSync(filePath);
          const base64 = pdfBuffer.toString('base64');
          imported.push({ title: baseName, content: '', file_data: base64, file_type: 'pdf' });
        } catch {
          imported.push({ title: baseName, content: '<p>Error al importar el archivo .pdf</p>' });
        }
      } else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
        try {
          const fileBuffer = readFileSync(filePath);
          const base64 = fileBuffer.toString('base64');
          imported.push({ title: baseName, content: '', file_data: base64, file_type: 'xlsx' });
        } catch {
          imported.push({ title: baseName, content: '<p>Error al importar el archivo Excel</p>' });
        }
      } else if (ext === '.pptx') {
        try {
          const fileBuffer = readFileSync(filePath);
          const presentation = parsePptxBuffer(fileBuffer);
          imported.push({
            title: baseName,
            content: JSON.stringify(presentation),
            file_data: fileBuffer.toString('base64'),
            file_type: 'pptx',
          });
        } catch {
          imported.push({ title: baseName, content: '<p>Error al importar el archivo PowerPoint</p>' });
        }
      }
    }

    return imported;
  });

  // ── File Import JSON ──
  ipcMain.handle('file:importJson', async () => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showOpenDialog(win!, {
      title: 'Importar notas desde JSON',
      filters: [
        { name: 'JSON', extensions: ['json'] },
      ],
      properties: ['openFile', 'multiSelections'],
    });
    if (result.canceled || result.filePaths.length === 0) return [];

    const imported: { title: string; content: string; file_data?: string; file_type?: string }[] = [];

    for (const filePath of result.filePaths) {
      try {
        const raw = readFileSync(filePath, 'utf-8');
        const json = JSON.parse(raw);

        // Support both single note and array of notes
        const notes = Array.isArray(json) ? json : [json];
        for (const note of notes) {
          if (note && typeof note.title === 'string' && typeof note.content === 'string') {
            const entry: { title: string; content: string; file_data?: string; file_type?: string } = {
              title: note.title,
              content: note.content,
            };
            if (note.file_data) entry.file_data = note.file_data;
            if (note.file_type) entry.file_type = note.file_type;
            imported.push(entry);
          }
        }
      } catch {
        // If JSON parsing fails, skip this file
        const baseName = path.basename(filePath, '.json');
        imported.push({ title: baseName, content: '<p>Error al importar el archivo .json</p>' });
      }
    }

    return imported;
  });

  // ── File Export Excel ──
  ipcMain.handle('file:exportXlsx', async (_event, options: { defaultName: string; htmlContent: string; title: string }) => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showSaveDialog(win!, {
      defaultPath: options.defaultName,
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    });
    if (result.canceled || !result.filePath) return { success: false };

    try {
      // Parse HTML tables into a workbook
      const div = `<html><body>${options.htmlContent}</body></html>`;
      const workbook = XLSX.read(div, { type: 'string' });

      // If no sheets were parsed (no tables in HTML), create one from plain text
      if (workbook.SheetNames.length === 0) {
        const tmpEl = options.htmlContent.replace(/<[^>]+>/g, '\n');
        const lines = tmpEl.split('\n').filter((l: string) => l.trim());
        const data = lines.map((line: string) => [line.trim()]);
        if (data.length === 0) data.push([options.title]);
        const ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, ws, 'Nota');
      }

      XLSX.writeFile(workbook, result.filePath);
      return { success: true, filePath: result.filePath };
    } catch {
      return { success: false };
    }
  });

  // ── File Export spreadsheet from base64 (for workbook notes) ──
  ipcMain.handle('file:exportXlsxRaw', async (_event, options: {
    defaultName: string;
    base64: string;
    format?: 'xlsx' | 'xls' | 'csv';
    sheetIndex?: number;
  }) => {
    const format = options.format ?? 'xlsx';
    const extension = format;
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showSaveDialog(win!, {
      defaultPath: options.defaultName,
      filters: [{
        name: format === 'csv' ? 'CSV' : format === 'xls' ? 'Excel 97-2003' : 'Excel',
        extensions: [extension],
      }],
    });
    if (result.canceled || !result.filePath) return { success: false };
    try {
      const source = XLSX.read(Buffer.from(options.base64, 'base64'), { type: 'buffer' });
      let workbook = source;
      if (format === 'csv') {
        const selectedName = source.SheetNames[options.sheetIndex ?? 0] ?? source.SheetNames[0];
        if (!selectedName) return { success: false };
        workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, source.Sheets[selectedName], selectedName);
      }
      const filePath = result.filePath.toLowerCase().endsWith(`.${extension}`)
        ? result.filePath
        : `${result.filePath}.${extension}`;
      XLSX.writeFile(workbook, filePath, {
        bookType: format === 'xls' ? 'biff8' : format,
      });
      return { success: true, filePath };
    } catch {
      return { success: false };
    }
  });

  // ── XLSX Parse: read base64 and return structured sheet data ──
  ipcMain.handle('xlsx:parseData', (_event, base64: string) => {
    try {
      const buffer = Buffer.from(base64, 'base64');
      const workbook = XLSX.read(buffer, { type: 'buffer', cellStyles: true });
      const sheets: { name: string; data: XlsxCellValue[][]; colCount: number; formulas: { row: number; col: number; formula: string }[]; merges?: { s: { r: number; c: number }; e: { r: number; c: number } }[]; colWidths?: number[]; rowHeights?: number[] }[] = [];

      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(sheet, {
          header: 1,
          defval: null,
          blankrows: true,
        });

        // Determine actual column count from range
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
        const colCount = range.e.c + 1;

        // Get merges
        const merges = sheet['!merges']?.map((m: XLSX.Range) => ({
          s: { r: m.s.r, c: m.s.c },
          e: { r: m.e.r, c: m.e.c },
        })) || [];

        const formulas = Object.entries(sheet).flatMap(([cellRef, cell]) => {
          if (cellRef.startsWith('!') || !cell || typeof cell !== 'object' || !('f' in cell) || !cell.f) return [];
          const decoded = XLSX.utils.decode_cell(cellRef);
          return [{ row: decoded.r, col: decoded.c, formula: `=${cell.f}` }];
        });

        const colWidths = (sheet['!cols'] ?? []).map(column => {
          const width = column?.wpx ?? (column?.wch ? column.wch * 7 : 80);
          return Math.max(36, Math.min(600, Number(width) || 80));
        });
        const rowHeights = (sheet['!rows'] ?? []).map(row => {
          const height = row?.hpx ?? (row?.hpt ? row.hpt * (96 / 72) : 20);
          return Math.max(18, Math.min(300, Number(height) || 20));
        });

        sheets.push({ name: sheetName, data: jsonData as XlsxCellValue[][], colCount, formulas, merges, colWidths, rowHeights });
      }

      return { success: true, sheets };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // ── XLSX Update Cell: update a cell value and return new base64 ──
  ipcMain.handle('xlsx:updateCell', (_event, base64: string, sheetIndex: number, row: number, col: number, value: XlsxCellValue, formula?: string | null) => {
    try {
      const buffer = Buffer.from(base64, 'base64');
      const workbook = XLSX.read(buffer, { type: 'buffer', cellStyles: true });
      const sheetName = workbook.SheetNames[sheetIndex];
      if (!sheetName) return { success: false, error: 'Sheet not found' };

      const sheet = workbook.Sheets[sheetName];
      setXlsxCell(sheet, { row, col, value, formula });

      // Update range if needed
      const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
      if (row > range.e.r) range.e.r = row;
      if (col > range.e.c) range.e.c = col;
      sheet['!ref'] = XLSX.utils.encode_range(range);

      return { success: true, base64: workbookToBase64(workbook) };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // ── XLSX Create Blank: create an empty workbook and return base64 ──
  ipcMain.handle('xlsx:createBlank', () => {
    try {
      const workbook = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([[]]);
      XLSX.utils.book_append_sheet(workbook, ws, 'Sheet1');
      const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
      const base64 = Buffer.from(buffer).toString('base64');
      return { success: true, base64 };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // ── XLSX Batch Update Cells: update multiple cells at once and return new base64 ──
  ipcMain.handle('xlsx:batchUpdateCells', (_event, base64: string, sheetIndex: number, cells: XlsxCellUpdate[]) => {
    try {
      const buffer = Buffer.from(base64, 'base64');
      const workbook = XLSX.read(buffer, { type: 'buffer', cellStyles: true });
      const sheetName = workbook.SheetNames[sheetIndex];
      if (!sheetName) return { success: false, error: 'Sheet not found' };

      const sheet = workbook.Sheets[sheetName];
      const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');

      for (const cell of cells) {
        setXlsxCell(sheet, cell);
        if (cell.row > range.e.r) range.e.r = cell.row;
        if (cell.col > range.e.c) range.e.c = cell.col;
      }

      sheet['!ref'] = XLSX.utils.encode_range(range);
      return { success: true, base64: workbookToBase64(workbook) };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  ipcMain.handle('xlsx:replaceSheetData', (_event, base64: string, sheetIndex: number, data: XlsxCellValue[][], formulas: XlsxCellUpdate[], merges: XLSX.Range[] = []) => {
    try {
      const workbook = XLSX.read(Buffer.from(base64, 'base64'), { type: 'buffer', cellStyles: true });
      const sheetName = workbook.SheetNames[sheetIndex];
      if (!sheetName) return { success: false, error: 'Sheet not found' };
      if (!Array.isArray(data) || data.length > 100_000) return { success: false, error: 'Invalid sheet data' };

      const previous = workbook.Sheets[sheetName];
      const replacement = XLSX.utils.aoa_to_sheet(data);
      if (previous['!cols']) replacement['!cols'] = previous['!cols'];
      if (previous['!rows']) replacement['!rows'] = previous['!rows'];
      if (Array.isArray(merges) && merges.length > 0) replacement['!merges'] = merges.slice(0, 10_000);
      for (const formula of formulas.slice(0, 100_000)) setXlsxCell(replacement, formula);
      workbook.Sheets[sheetName] = replacement;
      return { success: true, base64: workbookToBase64(workbook) };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  ipcMain.handle('xlsx:updateDimensions', (_event, base64: string, sheetIndex: number, colWidths: number[], rowHeights: number[]) => {
    try {
      const workbook = XLSX.read(Buffer.from(base64, 'base64'), { type: 'buffer', cellStyles: true });
      const sheetName = workbook.SheetNames[sheetIndex];
      if (!sheetName) return { success: false, error: 'Sheet not found' };
      if (!Array.isArray(colWidths) || !Array.isArray(rowHeights) || colWidths.length > 16_384 || rowHeights.length > 100_000) {
        return { success: false, error: 'Invalid dimensions' };
      }
      const sheet = workbook.Sheets[sheetName];
      sheet['!cols'] = colWidths.map(width => ({ wpx: Math.max(36, Math.min(600, Number(width) || 80)) }));
      sheet['!rows'] = rowHeights.map(height => ({ hpx: Math.max(18, Math.min(300, Number(height) || 20)) }));
      return { success: true, base64: workbookToBase64(workbook) };
    } catch (err) {
      return { success: false, error: String(err) };
    }
  });

  // ── Window ──
  ipcMain.handle('window:openNote', (_event, noteId: number) => {
    const parentWin = BrowserWindow.getFocusedWindow();
    const devUrl = process.env['ELECTRON_DEV_URL'];

    const noteWin = new BrowserWindow({
      width: 900,
      height: 700,
      minWidth: 600,
      minHeight: 400,
      title: 'Helper Notes',
      icon: path.join(__dirname, '../assets/favicon.ico'),
      resizable: true,
      parent: parentWin ?? undefined,
      webPreferences: {
        preload: path.join(__dirname, 'preload.js'),
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: false,
      },
    });

    noteWin.setMenu(null);

    if (devUrl) {
      noteWin.loadURL(`${devUrl}?noteId=${noteId}`);
    } else {
      noteWin.loadFile(
        path.join(__dirname, '../dist/app-desktop-helper-notes/browser/index.html'),
        { search: `noteId=${noteId}` }
      );
    }
  });

  // ── File Utilities ──
  ipcMain.handle('file:showInFolder', (_event, filePath: string) => {
    shell.showItemInFolder(filePath);
  });

  ipcMain.handle('file:printToPdf', async (_event, options: { defaultName: string; htmlContent: string }) => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showSaveDialog(win!, {
      defaultPath: options.defaultName,
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
    });
    if (result.canceled || !result.filePath) return { success: false };

    // Create a hidden window to render HTML → PDF
    const pdfWin = new BrowserWindow({
      show: false,
      width: 800,
      height: 600,
      webPreferences: { offscreen: true, javascript: false, sandbox: true },
    });

    const styledHtml = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; padding: 40px 60px; color: #1e293b; line-height: 1.6; font-size: 14px; }
  h1 { font-size: 28px; margin-bottom: 8px; }
  h2 { font-size: 22px; margin-top: 20px; }
  h3 { font-size: 18px; margin-top: 16px; }
  p { margin: 8px 0; }
  ul, ol { margin: 8px 0; padding-left: 24px; }
  b, strong { font-weight: 700; }
  i, em { font-style: italic; }
  u { text-decoration: underline; }
  s, del, strike { text-decoration: line-through; }
  img { max-width: 100%; height: auto; }
  table { width: 100%; border-collapse: collapse; margin: 12px 0; }
  th, td { border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; }
</style></head><body>${options.htmlContent}</body></html>`;
    try {
      await pdfWin.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(styledHtml));
      const pdfData = await pdfWin.webContents.printToPDF({
        printBackground: true,
        pageSize: 'A4',
        margins: { top: 0.4, bottom: 0.4, left: 0.4, right: 0.4 },
      });
      writeFileSync(result.filePath, pdfData);
      return { success: true, filePath: result.filePath };
    } catch (error) {
      return { success: false, error: error instanceof Error ? error.message : String(error) };
    } finally {
      if (!pdfWin.isDestroyed()) pdfWin.destroy();
    }
  });

  ipcMain.handle('window:setFullScreen', (_event, enabled: boolean) => {
    const win = BrowserWindow.getFocusedWindow();
    if (!win) return { success: false };
    win.setFullScreen(Boolean(enabled));
    return { success: true };
  });

  // ── Ollama AI ──

  ipcMain.handle('llama:managerStatus', (_event, endpoint: string) => llamaManagerStatus(endpoint));

  ipcMain.handle('llama:install', event => installLlamaCpp(progress => {
    if (!event.sender.isDestroyed()) event.sender.send('llama:progress', progress);
  }));

  ipcMain.handle('llama:catalog', (_event, query: string) => searchLlamaModels(query));

  ipcMain.handle('llama:modelFiles', (_event, repository: string) => listLlamaModelFiles(repository));

  ipcMain.handle('llama:downloadModel', (event, repository: string, fileName: string) =>
    downloadLlamaModel(repository, fileName, progress => {
      if (!event.sender.isDestroyed()) event.sender.send('llama:progress', progress);
    })
  );

  ipcMain.handle('llama:cancelDownload', () => {
    cancelLlamaDownload();
    return { success: true };
  });

  ipcMain.handle('llama:localModels', () => listLocalLlamaModels());

  ipcMain.handle('llama:start', (_event, modelPath: string, endpoint: string) => startManagedLlamaServer(modelPath, endpoint));

  ipcMain.handle('llama:stop', () => stopManagedLlamaServer());

  ipcMain.handle('ai:status', (_event, config: AiRuntimeConfig) => checkAiProvider(config));

  ipcMain.handle('ai:listModels', (_event, config: AiRuntimeConfig) => listAiModels(config));

  ipcMain.handle('ai:chat', (event, config: AiRuntimeConfig, model: string, messages: { role: string; content: string }[]) => {
    const senderId = event.sender.id;
    streamAiChat(senderId, config, model, messages, token => {
      const win = BrowserWindow.getAllWindows().find(item => item.webContents.id === senderId && !item.isDestroyed());
      win?.webContents.send('ai:token', token);
    }).then(() => {
      const win = BrowserWindow.getAllWindows().find(item => item.webContents.id === senderId && !item.isDestroyed());
      win?.webContents.send('ai:done');
    }).catch(error => {
      const win = BrowserWindow.getAllWindows().find(item => item.webContents.id === senderId && !item.isDestroyed());
      win?.webContents.send('ai:error', error instanceof Error ? error.message : String(error));
    });
    return { success: true };
  });

  ipcMain.handle('ai:abort', event => {
    abortAiChat(event.sender.id);
    return { success: true };
  });

  ipcMain.handle('ollama:checkInstalled', () => {
    return { installed: isOllamaInstalled() };
  });

  ipcMain.handle('ollama:status', async () => {
    try {
      await ollamaRequest('GET', '/api/tags');
      return { running: true };
    } catch {
      return { running: false };
    }
  });

  ipcMain.handle('ollama:start', async () => {
    try {
      // Check if already running
      try {
        await ollamaRequest('GET', '/api/tags');
        return { success: true, alreadyRunning: true };
      } catch { /* not running, proceed to start */ }

      const ollamaPath = findOllamaPath();
      if (!ollamaPath) {
        return { success: false, error: 'No se encontró el ejecutable de Ollama' };
      }

      ollamaProcess = spawn(ollamaPath, ['serve'], {
        detached: true,
        stdio: 'ignore',
        windowsHide: true,
      });
      ollamaProcess.unref();

      // Wait for it to become available (up to 10s)
      for (let i = 0; i < 20; i++) {
        await new Promise(r => setTimeout(r, 500));
        try {
          await ollamaRequest('GET', '/api/tags');
          return { success: true };
        } catch { /* keep waiting */ }
      }
      return { success: false, error: 'Ollama no respondió a tiempo' };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('ollama:stop', async () => {
    try {
      if (process.platform === 'win32') {
        execSync('taskkill /F /IM ollama.exe /T', { stdio: 'ignore' });
      } else {
        execSync('pkill -f ollama', { stdio: 'ignore' });
      }
      ollamaProcess = null;
      return { success: true };
    } catch {
      return { success: true }; // Already stopped
    }
  });

  ipcMain.handle('ollama:openDownload', () => {
    shell.openExternal('https://ollama.com/download');
    return { success: true };
  });

  ipcMain.handle('ollama:listModels', async () => {
    try {
      const data = await ollamaRequest('GET', '/api/tags');
      const allModels: any[] = (data as any).models ?? [];
      // Filter out non-conversational models (embedding, etc.)
      const chatModels = allModels.filter((m: any) => !NON_CHAT_PATTERNS.test(m.name));
      return { success: true, models: chatModels };
    } catch (err: any) {
      return { success: false, models: [], error: err.message };
    }
  });

  ipcMain.handle('ollama:chat', (event, model: string, messages: { role: string; content: string }[]) => {
    const senderId = event.sender.id;
    ollamaStream(model, messages, (chunk: string) => {
      // Send each token to the renderer
      const win = BrowserWindow.getAllWindows().find(w => w.webContents.id === senderId && !w.isDestroyed());
      if (win) {
        win.webContents.send('ollama:token', chunk);
      }
    }).then(() => {
      const win = BrowserWindow.getAllWindows().find(w => w.webContents.id === senderId && !w.isDestroyed());
      if (win) win.webContents.send('ollama:done');
    }).catch((err: any) => {
      const win = BrowserWindow.getAllWindows().find(w => w.webContents.id === senderId && !w.isDestroyed());
      if (win) win.webContents.send('ollama:error', err.message);
    });
    // Return immediately; tokens will be streamed via events
    return { success: true };
  });

  ipcMain.handle('ollama:abort', () => {
    if (currentOllamaRequest) {
      currentOllamaRequest.destroy();
      currentOllamaRequest = null;
    }
    return { success: true };
  });

  // ── AI Conversations ──

  ipcMain.handle('conversations:getAll', () => {
    return getAllConversations();
  });

  ipcMain.handle('conversations:getById', (_event, id: number) => {
    return getConversationById(id);
  });

  ipcMain.handle('conversations:create', (_event, title: string, model: string, assistantId?: string) => {
    return createConversation(title, model, assistantId);
  });

  ipcMain.handle('ollama:pull', async (event, model: string) => {
    const senderId = event.sender.id;
    return pullOllamaModel(model, progress => {
      const win = BrowserWindow.getAllWindows().find(item => item.webContents.id === senderId && !item.isDestroyed());
      win?.webContents.send('ollama:pullProgress', progress);
    });
  });

  ipcMain.handle('ollama:signin', () => {
    try {
      const ollamaPath = findOllamaPath();
      if (!ollamaPath) return { success: false, error: 'No se encontró el ejecutable de Ollama' };
      const signin = spawn(ollamaPath, ['signin'], {
        detached: true,
        stdio: 'ignore',
        windowsHide: true,
      });
      signin.unref();
      return { success: true };
    } catch (error) {
      return { success: false, error: error instanceof Error ? error.message : String(error) };
    }
  });

  ipcMain.handle('conversations:update', (_event, id: number, data: { title?: string; model?: string; assistant_id?: string }) => {
    return updateConversation(id, data);
  });

  ipcMain.handle('conversations:delete', (_event, id: number) => {
    deleteConversation(id);
    return { success: true };
  });

  ipcMain.handle('conversations:getMessages', (_event, conversationId: number) => {
    return getMessagesByConversation(conversationId);
  });

  ipcMain.handle('conversations:addMessage', (_event, conversationId: number, role: string, content: string) => {
    return addMessage(conversationId, role, content);
  });

  ipcMain.handle('conversations:updateMessage', (_event, id: number, content: string) => {
    updateMessage(id, content);
    return { success: true };
  });
}

// ── Helpers ──

/** Escapa texto no confiable antes de incorporarlo a documentos HTML. */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Minimal .docx (OOXML) parser — extracts word/document.xml from the ZIP,
 * then converts <w:p> / <w:r> elements to HTML.
 */
/** Extrae párrafos y formato básico de un DOCX sin ejecutar contenido embebido. */
function parseDocxToHtml(buffer: Buffer): string {
  // Read ZIP entries
  const entries = readZipEntries(buffer);
  const docEntry = entries.find((e) => e.name === 'word/document.xml');
  if (!docEntry) return '<p></p>';

  const xml = docEntry.data.toString('utf-8');

  // Extract paragraphs
  const paragraphs: string[] = [];
  const pRegex = /<w:p[\s>]([\s\S]*?)<\/w:p>/g;
  let pMatch: RegExpExecArray | null;

  while ((pMatch = pRegex.exec(xml)) !== null) {
    const pContent = pMatch[1];

    // Check paragraph properties for alignment and heading style
    let align = '';
    let tag = 'p';
    const pPrMatch = pContent.match(/<w:pPr>([\s\S]*?)<\/w:pPr>/);
    if (pPrMatch) {
      const pPr = pPrMatch[1];
      const jcMatch = pPr.match(/<w:jc\s+w:val="([^"]+)"/);
      if (jcMatch) {
        const jcVal = jcMatch[1];
        if (jcVal === 'center') align = ' style="text-align:center"';
        else if (jcVal === 'right') align = ' style="text-align:right"';
        else if (jcVal === 'both') align = ' style="text-align:justify"';
      }
      // Heading styles
      const styleMatch = pPr.match(/<w:pStyle\s+w:val="([^"]+)"/);
      if (styleMatch) {
        const s = styleMatch[1].toLowerCase();
        if (s.includes('heading1') || s === 'titulo' || s === 'title') tag = 'h1';
        else if (s.includes('heading2')) tag = 'h2';
        else if (s.includes('heading3')) tag = 'h3';
      }
    }

    // Extract runs
    const runs: string[] = [];
    const rRegex = /<w:r[\s>]([\s\S]*?)<\/w:r>/g;
    let rMatch: RegExpExecArray | null;

    while ((rMatch = rRegex.exec(pContent)) !== null) {
      const rContent = rMatch[1];

      // Text
      const tMatch = rContent.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/);
      if (!tMatch) continue;
      let text = escapeHtml(tMatch[1]);

      // Run properties
      const rPrMatch = rContent.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
      if (rPrMatch) {
        const rPr = rPrMatch[1];
        if (/<w:b[\s/>]/.test(rPr) && !/<w:b\s+w:val="false"/.test(rPr)) text = `<b>${text}</b>`;
        if (/<w:i[\s/>]/.test(rPr) && !/<w:i\s+w:val="false"/.test(rPr)) text = `<i>${text}</i>`;
        if (/<w:u[\s/>]/.test(rPr)) text = `<u>${text}</u>`;
        if (/<w:strike[\s/>]/.test(rPr)) text = `<s>${text}</s>`;
      }

      runs.push(text);
    }

    const inner = runs.join('') || '<br>';
    paragraphs.push(`<${tag}${align}>${inner}</${tag}>`);
  }

  return paragraphs.join('') || '<p></p>';
}

/**
 * Minimal ZIP reader — extracts file entries from a ZIP buffer.
 * Supports stored (method 0) and deflated (method 8) entries.
 */
/** Lee las entradas ZIP mínimas necesarias para importar DOCX. */
function readZipEntries(buf: Buffer): { name: string; data: Buffer }[] {
  const entries: { name: string; data: Buffer }[] = [];
  let offset = 0;

  while (offset + 4 <= buf.length) {
    const sig = buf.readUInt32LE(offset);
    if (sig !== 0x04034b50) break; // Not a local file header

    const method = buf.readUInt16LE(offset + 8);
    const compressedSize = buf.readUInt32LE(offset + 18);
    const uncompressedSize = buf.readUInt32LE(offset + 22);
    const nameLen = buf.readUInt16LE(offset + 26);
    const extraLen = buf.readUInt16LE(offset + 28);
    const name = buf.toString('utf-8', offset + 30, offset + 30 + nameLen);
    const dataStart = offset + 30 + nameLen + extraLen;
    const rawData = buf.subarray(dataStart, dataStart + compressedSize);

    let data: Buffer;
    if (method === 0) {
      data = rawData;
    } else if (method === 8) {
      data = inflateRawSync(rawData);
    } else {
      data = Buffer.alloc(0);
    }

    entries.push({ name, data });
    offset = dataStart + compressedSize;
  }

  return entries;
}

// ── Ollama helpers ──

const OLLAMA_HOST = 'http://127.0.0.1';
const OLLAMA_PORT = 11434;

let currentOllamaRequest: http.ClientRequest | null = null;
let ollamaProcess: ChildProcess | null = null;

/** Detecta Ollama en ubicaciones conocidas o a través del PATH del sistema. */
function isOllamaInstalled(): boolean {
  try {
    if (process.platform === 'win32') {
      // Check common install locations
      const localAppData = process.env.LOCALAPPDATA || '';
      const programFiles = process.env.ProgramFiles || 'C:\\Program Files';
      const paths = [
        path.join(localAppData, 'Programs', 'Ollama', 'ollama.exe'),
        path.join(programFiles, 'Ollama', 'ollama.exe'),
        path.join(localAppData, 'Ollama', 'ollama.exe'),
      ];
      if (paths.some(p => existsSync(p))) return true;
      // Also try to find via where command
      try {
        execSync('where ollama', { stdio: 'ignore' });
        return true;
      } catch { return false; }
    } else {
      execSync('which ollama', { stdio: 'ignore' });
      return true;
    }
  } catch {
    return false;
  }
}

/** Resuelve la ruta ejecutable que se usará para iniciar y administrar Ollama. */
function findOllamaPath(): string | null {
  try {
    if (process.platform === 'win32') {
      const localAppData = process.env.LOCALAPPDATA || '';
      const programFiles = process.env.ProgramFiles || 'C:\\Program Files';
      const paths = [
        path.join(localAppData, 'Programs', 'Ollama', 'ollama.exe'),
        path.join(programFiles, 'Ollama', 'ollama.exe'),
        path.join(localAppData, 'Ollama', 'ollama.exe'),
      ];
      const found = paths.find(p => existsSync(p));
      if (found) return found;
      try {
        return execSync('where ollama', { encoding: 'utf-8' }).trim().split('\n')[0];
      } catch { return null; }
    } else {
      return execSync('which ollama', { encoding: 'utf-8' }).trim();
    }
  } catch {
    return null;
  }
}

/** Cliente JSON de bajo nivel para operaciones no transmitidas de Ollama. */
function ollamaRequest(method: string, path: string, body?: any): Promise<any> {
  return new Promise((resolve, reject) => {
    const postData = body ? JSON.stringify(body) : undefined;
    const req = http.request({
      hostname: '127.0.0.1',
      port: OLLAMA_PORT,
      path,
      method,
      headers: postData ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(postData) } : {},
      timeout: 5000,
    }, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch { resolve(data); }
      });
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('Ollama no está disponible')); });
    if (postData) req.write(postData);
    req.end();
  });
}

/** Descarga un modelo interpretando progreso NDJSON y reenviándolo al renderer. */
function pullOllamaModel(
  model: string,
  onProgress: (progress: { model: string; status: string; completed: number; total: number; percent: number }) => void,
): Promise<{ success: boolean; error?: string }> {
  if (!/^[a-zA-Z0-9][a-zA-Z0-9._/:+-]{0,199}$/.test(model)) {
    return Promise.resolve({ success: false, error: 'El identificador del modelo no es válido' });
  }
  return new Promise(resolve => {
    const payload = JSON.stringify({ model, stream: true });
    const request = http.request({
      hostname: '127.0.0.1',
      port: OLLAMA_PORT,
      path: '/api/pull',
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload) },
    }, response => {
      let buffer = '';
      let settled = false;
      const finish = (result: { success: boolean; error?: string }): void => {
        if (settled) return;
        settled = true;
        resolve(result);
      };
      const processLine = (line: string): void => {
        if (!line.trim()) return;
        try {
          const data = JSON.parse(line) as Record<string, unknown>;
          if (data['error']) {
            finish({ success: false, error: String(data['error']) });
            return;
          }
          const completed = Number(data['completed'] ?? 0);
          const total = Number(data['total'] ?? 0);
          onProgress({
            model,
            status: String(data['status'] ?? 'Preparando modelo'),
            completed,
            total,
            percent: total > 0 ? Math.min(100, Math.round((completed / total) * 100)) : 0,
          });
        } catch {
          // Wait for the rest of a partial NDJSON line.
        }
      };
      response.setEncoding('utf8');
      response.on('data', chunk => {
        buffer += chunk;
        const lines = buffer.split('\n');
        buffer = lines.pop() ?? '';
        lines.forEach(processLine);
      });
      response.on('end', () => {
        processLine(buffer);
        if ((response.statusCode ?? 500) >= 400) finish({ success: false, error: `Ollama respondió con HTTP ${response.statusCode}` });
        else finish({ success: true });
      });
      response.on('error', error => finish({ success: false, error: error.message }));
    });
    request.on('error', error => resolve({ success: false, error: error.message }));
    request.write(payload);
    request.end();
  });
}

/** Consume `/api/chat` como NDJSON y entrega incrementalmente cada token. */
function ollamaStream(model: string, messages: { role: string; content: string }[], onToken: (chunk: string) => void): Promise<void> {
  return new Promise((resolve, reject) => {
    const postData = JSON.stringify({ model, messages, stream: true });
    const req = http.request({
      hostname: '127.0.0.1',
      port: OLLAMA_PORT,
      path: '/api/chat',
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(postData) },
    }, (res) => {
      res.on('data', (chunk) => {
        const lines = chunk.toString().split('\n').filter((l: string) => l.trim());
        for (const line of lines) {
          try {
            const json = JSON.parse(line);
            if (json.message?.content) {
              onToken(json.message.content);
            }
            if (json.done) {
              currentOllamaRequest = null;
              resolve();
            }
          } catch { /* ignore parse errors on partial data */ }
        }
      });
      res.on('end', () => { currentOllamaRequest = null; resolve(); });
      res.on('error', (err) => { currentOllamaRequest = null; reject(err); });
    });
    req.on('error', (err) => { currentOllamaRequest = null; reject(err); });
    currentOllamaRequest = req;
    req.write(postData);
    req.end();
  });
}
