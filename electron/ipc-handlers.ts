import { ipcMain, dialog, BrowserWindow, shell } from 'electron';
import { readFileSync, writeFileSync, existsSync } from 'fs';
import { execSync, spawn, ChildProcess } from 'child_process';
import http from 'http';
import path from 'path';
import { inflateRawSync } from 'zlib';
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

// Families that are NOT conversational (embedding, vision-only, code-only, etc.)
const NON_CHAT_PATTERNS = /\b(embed|embedding|all-minilm|bge-|nomic-embed|mxbae|snowflake-arctic-embed)\b/i;

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

  // ── File Import ──
  ipcMain.handle('file:import', async () => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showOpenDialog(win!, {
      title: 'Importar archivo',
      filters: [
        { name: 'Documentos', extensions: ['txt', 'docx', 'pdf', 'html', 'htm'] },
        { name: 'Texto', extensions: ['txt'] },
        { name: 'HTML', extensions: ['html', 'htm'] },
        { name: 'Word', extensions: ['docx'] },
        { name: 'PDF', extensions: ['pdf'] },
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
      }
    }

    return imported;
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
      webPreferences: { offscreen: true },
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
</style></head><body>${options.htmlContent}</body></html>`;

    await pdfWin.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(styledHtml));
    // Wait for content to render
    await new Promise(resolve => setTimeout(resolve, 500));

    const pdfData = await pdfWin.webContents.printToPDF({
      printBackground: true,
      pageSize: 'Letter',
      margins: { top: 0.4, bottom: 0.4, left: 0.4, right: 0.4 },
    });

    pdfWin.destroy();
    writeFileSync(result.filePath, pdfData);
    return { success: true, filePath: result.filePath };
  });

  // ── Ollama AI ──

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

  ipcMain.handle('conversations:create', (_event, title: string, model: string) => {
    return createConversation(title, model);
  });

  ipcMain.handle('conversations:update', (_event, id: number, data: { title?: string; model?: string }) => {
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
