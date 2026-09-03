import { existsSync, readFileSync, readdirSync, statSync } from 'fs';
import path from 'path';
import yauzl from 'yauzl';
import { strToU8, zipSync } from 'fflate';
import PptxGenJS from 'pptxgenjs';
import { getDb } from './db/db';
import { getAllFolders } from './db/folder.repository';
import { getNoteById, getNotesByFolder, NoteRow } from './db/note.repository';
import { parsePptxBuffer } from './presentation-parser';

const SCHEMA = 'helper-notes-export';
const VERSION = 1;
const MAX_BACKUP_BYTES = 512 * 1024 * 1024;
const MAX_BACKUP_FILES = 100;
const MAX_FOLDERS = 500;
const MAX_NOTES = 10_000;

/** Nota independiente del esquema interno de SQLite usada en respaldos. */
export interface PortableNote {
  title: string;
  content: string;
  fileData: string | null;
  fileType: string | null;
  createdAt?: string;
  updatedAt?: string;
}

export interface PortableFolder {
  name: string;
  /** Ancestor names, relative to the root selected for export. */
  parentPath?: string[];
  createdAt?: string;
  updatedAt?: string;
  notes: PortableNote[];
}

export interface HelperNotesBackup {
  schema: typeof SCHEMA;
  version: typeof VERSION;
  exportedAt: string;
  folders: PortableFolder[];
  notes: PortableNote[];
}

export interface BackupSelection {
  folderIds: number[];
  noteIds: number[];
}

export interface ImportSummary {
  folders: number;
  notes: number;
  sources: number;
  errors: string[];
}

function uniqueIds(values: number[]): number[] {
  return [...new Set(values.filter(value => Number.isSafeInteger(value) && value > 0))];
}

function portableNote(note: NoteRow): PortableNote {
  return {
    title: note.title,
    content: note.content,
    fileData: note.file_data ?? null,
    fileType: note.file_type ?? null,
    createdAt: note.created_at,
    updatedAt: note.updated_at,
  };
}

/**
 * Construye un respaldo autocontenido a partir de una selección. Incluye los
 * ancestros necesarios para conservar jerarquía y evita duplicar notas ya
 * cubiertas por carpetas seleccionadas.
 */
export function createBackup(selection: BackupSelection): HelperNotesBackup {
  const requestedFolderIds = uniqueIds(selection.folderIds);
  const noteIds = uniqueIds(selection.noteIds);
  const allFolders = getAllFolders();
  const foldersById = new Map(allFolders.map(folder => [folder.id, folder]));
  const requestedSet = new Set(requestedFolderIds);
  const hasSelectedAncestor = (folderId: number): boolean => {
    let current = foldersById.get(folderId)?.parent_id ?? null;
    const visited = new Set<number>();
    while (current !== null && !visited.has(current)) {
      if (requestedSet.has(current)) return true;
      visited.add(current);
      current = foldersById.get(current)?.parent_id ?? null;
    }
    return false;
  };
  const roots = requestedFolderIds.filter(id => foldersById.has(id) && !hasSelectedAncestor(id));
  const folderSet = new Set<number>();
  const folderPaths = new Map<number, string[]>();
  const collect = (folderId: number, parentPath: string[]): void => {
    const folder = foldersById.get(folderId);
    if (!folder || folderSet.has(folderId)) return;
    folderSet.add(folderId);
    folderPaths.set(folderId, parentPath);
    const childPath = [...parentPath, folder.name];
    allFolders.filter(child => child.parent_id === folderId).forEach(child => collect(child.id, childPath));
  };
  roots.forEach(id => collect(id, []));

  const folders: PortableFolder[] = [...folderSet].flatMap(folderId => {
    const folder = foldersById.get(folderId);
    if (!folder) return [];
    const folderNotes = getNotesByFolder(folderId).map(note => getNoteById(note.id)).filter((note): note is NoteRow => !!note);
    return [{
      name: folder.name,
      parentPath: folderPaths.get(folderId) ?? [],
      createdAt: folder.created_at,
      updatedAt: folder.updated_at,
      notes: folderNotes.map(portableNote),
    }];
  });

  const notes = noteIds.flatMap(noteId => {
    const note = getNoteById(noteId);
    if (!note || folderSet.has(note.folder_id ?? -1) || (note as NoteRow & { deleted_at?: string }).deleted_at) return [];
    return [portableNote(note)];
  });

  if (folders.length === 0 && notes.length === 0) throw new Error('No hay carpetas o notas válidas seleccionadas.');
  return { schema: SCHEMA, version: VERSION, exportedAt: new Date().toISOString(), folders, notes };
}

/** Serializa el respaldo con formato legible y estable para intercambio. */
export function backupToJson(backup: HelperNotesBackup): string {
  return JSON.stringify(backup, null, 2);
}

/** Limpia nombres para impedir rutas inválidas o traversal dentro del ZIP. */
function safeArchiveName(value: string, fallback: string): string {
  const normalized = value
    .normalize('NFKC')
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_')
    .replace(/[. ]+$/g, '')
    .trim()
    .slice(0, 180);
  const reserved = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i;
  if (!normalized) return fallback;
  return reserved.test(normalized) ? `_${normalized}` : normalized;
}

/** Reserva un nombre único agregando sufijos cuando existen colisiones. */
function uniqueArchiveName(directory: string, title: string, extension: string, usedPaths: Set<string>): string {
  const safeExtension = extension.toLowerCase().replace(/[^a-z0-9]/g, '') || 'bin';
  const suffix = `.${safeExtension}`;
  const titleWithoutExtension = title.toLowerCase().endsWith(suffix)
    ? title.slice(0, -suffix.length)
    : title;
  const base = safeArchiveName(titleWithoutExtension, 'nota');
  let candidate = `${directory}${base}${suffix}`;
  let copy = 2;
  while (usedPaths.has(candidate.toLocaleLowerCase())) {
    candidate = `${directory}${base} (${copy++})${suffix}`;
  }
  usedPaths.add(candidate.toLocaleLowerCase());
  return candidate;
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function noteAsHtml(note: PortableNote): Uint8Array {
  if (/<html[\s>]/i.test(note.content)) return strToU8(note.content);
  const title = escapeHtml(note.title || 'Nota');
  return strToU8(`<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${title}</title>
</head>
<body>
  <h1>${title}</h1>
  ${note.content}
</body>
</html>`);
}

function decodeFileData(fileData: string): Uint8Array {
  return new Uint8Array(Buffer.from(fileData, 'base64'));
}

async function presentationAsPptx(note: PortableNote): Promise<Uint8Array> {
  let slides: { title: string; body: string; notes?: string }[] = [];
  try {
    const document = JSON.parse(note.content) as Record<string, unknown>;
    if (document['kind'] === 'helper-notes-presentation' && Array.isArray(document['slides'])) {
      slides = document['slides'].slice(0, 500).flatMap(value => {
        if (!value || typeof value !== 'object') return [];
        const slide = value as Record<string, unknown>;
        return [{
          title: String(slide['title'] ?? '').slice(0, 500),
          body: String(slide['body'] ?? '').slice(0, 50_000),
          notes: String(slide['notes'] ?? '').slice(0, 50_000),
        }];
      });
    }
  } catch {
    // Imported presentations can still use their original physical PPTX below.
  }

  if (slides.length === 0 && note.fileData) return decodeFileData(note.fileData);
  if (slides.length === 0) slides = [{ title: note.title || 'Presentación', body: '' }];

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Helper Notes';
  pptx.company = 'Helper Notes';
  pptx.subject = 'Presentación exportada desde Helper Notes';
  pptx.title = note.title || 'Presentación';
  pptx.theme = { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' };
  slides.forEach((source, index) => {
    const slide = pptx.addSlide();
    slide.background = { color: index === 0 ? 'FFF7ED' : 'FFFFFF' };
    slide.addText(source.title.trim() || `Diapositiva ${index + 1}`, {
      x: 0.75, y: 0.55, w: 11.85, h: 0.8,
      fontFace: 'Aptos Display', fontSize: 28, bold: true, color: '9A3412',
      valign: 'middle', margin: 0.05, fit: 'shrink',
    });
    if (source.body.trim()) {
      slide.addText(source.body, {
        x: 0.9, y: 1.65, w: 11.45, h: 4.95,
        fontFace: 'Aptos', fontSize: 21, color: '334155',
        valign: 'top', margin: 0.08, fit: 'shrink',
      });
    }
    if (source.notes?.trim()) slide.addNotes(source.notes.trim());
  });
  const output = await pptx.write({ outputType: 'nodebuffer', compression: true });
  return new Uint8Array(output as Buffer);
}

/** Convierte cada tipo de nota a su representación física más interoperable. */
async function physicalNoteFile(note: PortableNote): Promise<{ extension: string; data: Uint8Array }> {
  const type = (note.fileType ?? '').toLowerCase().replace(/^\./, '');
  if (type === 'pptx') return { extension: 'pptx', data: await presentationAsPptx(note) };
  if (note.fileData) {
    const extension = ['pdf', 'xlsx', 'xls', 'csv', 'docx'].includes(type) ? type : 'bin';
    return { extension, data: decodeFileData(note.fileData) };
  }
  return { extension: 'html', data: noteAsHtml(note) };
}

/**
 * Empaqueta la estructura como directorios y archivos reales, e incluye un
 * manifiesto JSON para poder restaurar metadatos sin pérdida.
 */
export async function backupToZip(backup: HelperNotesBackup): Promise<Uint8Array> {
  const entries: Record<string, Uint8Array> = {};
  const usedPaths = new Set<string>();
  const usedDirectories = new Set<string>();

  for (const folder of backup.folders) {
    const rawDirectory = [...(folder.parentPath ?? []), folder.name]
      .map(segment => safeArchiveName(segment, 'carpeta'))
      .join('/');
    let directoryName = rawDirectory;
    let copy = 2;
    while (usedDirectories.has(directoryName.toLocaleLowerCase())) directoryName = `${rawDirectory} (${copy++})`;
    usedDirectories.add(directoryName.toLocaleLowerCase());
    const directory = `${directoryName}/`;
    entries[directory] = new Uint8Array();
    for (const note of folder.notes) {
      const file = await physicalNoteFile(note);
      entries[uniqueArchiveName(directory, note.title, file.extension, usedPaths)] = file.data;
    }
  }
  for (const note of backup.notes) {
    const file = await physicalNoteFile(note);
    entries[uniqueArchiveName('', note.title, file.extension, usedPaths)] = file.data;
  }

  return zipSync(entries, { level: 6 });
}

function asString(value: unknown, maxLength: number): string {
  return typeof value === 'string' ? value.slice(0, maxLength) : '';
}

/** Valida una nota no confiable y limita longitudes antes de importarla. */
function normalizeNote(value: unknown): PortableNote | null {
  if (!value || typeof value !== 'object') return null;
  const note = value as Record<string, unknown>;
  if (typeof note['title'] !== 'string' || typeof note['content'] !== 'string') return null;
  if (note['content'].length > 50_000_000) throw new Error('Una nota supera el tamaño máximo de contenido permitido.');
  const rawFileData = note['fileData'] ?? note['file_data'];
  const rawFileType = note['fileType'] ?? note['file_type'];
  if (typeof rawFileData === 'string' && rawFileData.length > MAX_BACKUP_BYTES) {
    throw new Error('Un archivo adjunto supera el tamaño máximo permitido.');
  }
  return {
    title: asString(note['title'], 1000),
    content: note['content'],
    fileData: typeof rawFileData === 'string' ? rawFileData : null,
    fileType: typeof rawFileType === 'string' ? rawFileType.slice(0, 50) : null,
    createdAt: asString(note['createdAt'] ?? note['created_at'], 40) || undefined,
    updatedAt: asString(note['updatedAt'] ?? note['updated_at'], 40) || undefined,
  };
}

/**
 * Analiza un respaldo JSON como entrada no confiable, valida esquema, límites,
 * identificadores y relaciones antes de devolverlo al importador.
 */
export function parseBackupJson(json: string): HelperNotesBackup {
  if (Buffer.byteLength(json, 'utf8') > MAX_BACKUP_BYTES) throw new Error('El respaldo supera el tamaño máximo permitido.');
  const parsed = JSON.parse(json) as unknown;

  if (Array.isArray(parsed)) {
    const notes = parsed.map(normalizeNote).filter((note): note is PortableNote => !!note);
    if (notes.length === 0) throw new Error('El JSON no contiene notas válidas.');
    return { schema: SCHEMA, version: VERSION, exportedAt: new Date().toISOString(), folders: [], notes };
  }
  if (!parsed || typeof parsed !== 'object') throw new Error('El archivo JSON no tiene una estructura válida.');
  const data = parsed as Record<string, unknown>;

  const legacyNote = normalizeNote(data);
  if (legacyNote) {
    return { schema: SCHEMA, version: VERSION, exportedAt: new Date().toISOString(), folders: [], notes: [legacyNote] };
  }
  if (data['schema'] !== SCHEMA || Number(data['version']) !== VERSION) {
    throw new Error('El archivo no es un respaldo compatible de Helper Notes.');
  }

  const rawFolders = Array.isArray(data['folders']) ? data['folders'] : [];
  const rawNotes = Array.isArray(data['notes']) ? data['notes'] : [];
  if (rawFolders.length > MAX_FOLDERS) throw new Error(`El respaldo contiene más de ${MAX_FOLDERS} carpetas.`);
  const rawNoteCount = rawNotes.length + rawFolders.reduce((total, value) => {
    if (!value || typeof value !== 'object') return total;
    const folderNotes = (value as Record<string, unknown>)['notes'];
    return total + (Array.isArray(folderNotes) ? folderNotes.length : 0);
  }, 0);
  if (rawNoteCount > MAX_NOTES) throw new Error(`El respaldo contiene más de ${MAX_NOTES} notas.`);
  const folders = rawFolders.flatMap(value => {
    if (!value || typeof value !== 'object') return [];
    const folder = value as Record<string, unknown>;
    const name = asString(folder['name'], 500).trim();
    if (!name) return [];
    const notes = (Array.isArray(folder['notes']) ? folder['notes'] : [])
      .map(normalizeNote)
      .filter((note): note is PortableNote => !!note);
    return [{
      name,
      parentPath: Array.isArray(folder['parentPath'])
        ? folder['parentPath'].map(value => asString(value, 500).trim()).filter(Boolean).slice(0, 50)
        : [],
      createdAt: asString(folder['createdAt'], 40) || undefined,
      updatedAt: asString(folder['updatedAt'], 40) || undefined,
      notes,
    }];
  });
  const notes = rawNotes
    .map(normalizeNote)
    .filter((note): note is PortableNote => !!note);
  const noteCount = notes.length + folders.reduce((total, folder) => total + folder.notes.length, 0);
  if (folders.length === 0 && noteCount === 0) throw new Error('El respaldo está vacío.');
  return { schema: SCHEMA, version: VERSION, exportedAt: asString(data['exportedAt'], 40) || new Date().toISOString(), folders, notes };
}

interface PhysicalZipEntry {
  name: string;
  directory: boolean;
  data: Buffer;
}

/** Lee un ZIP por streaming y rechaza entradas peligrosas o límites excedidos. */
function readZipEntries(filePath: string): Promise<PhysicalZipEntry[]> {
  return new Promise((resolve, reject) => {
    yauzl.open(filePath, { lazyEntries: true, validateEntrySizes: true }, (openError, zipFile) => {
      if (openError || !zipFile) return reject(openError ?? new Error('No se pudo abrir el ZIP.'));
      let settled = false;
      let entryCount = 0;
      let totalSize = 0;
      const entries: PhysicalZipEntry[] = [];
      const fail = (error: Error): void => {
        if (settled) return;
        settled = true;
        zipFile.close();
        reject(error);
      };
      zipFile.on('error', fail);
      zipFile.on('end', () => {
        if (settled) return;
        settled = true;
        zipFile.close();
        resolve(entries);
      });
      zipFile.on('entry', entry => {
        entryCount += 1;
        if (entryCount > MAX_NOTES + MAX_FOLDERS) return fail(new Error('El ZIP contiene demasiadas entradas.'));
        const normalized: string = entry.fileName.replace(/\\/g, '/');
        const unsafe = normalized.startsWith('/') || normalized.split('/').some(part => part === '..');
        if (unsafe) return fail(new Error(`El ZIP contiene una ruta insegura: ${entry.fileName}`));
        if ((entry.generalPurposeBitFlag & 1) !== 0) return fail(new Error('Los ZIP cifrados no son compatibles.'));
        const isDirectory = normalized.endsWith('/');
        if (isDirectory) {
          entries.push({ name: normalized, directory: true, data: Buffer.alloc(0) });
          zipFile.readEntry();
          return;
        }
        totalSize += entry.uncompressedSize;
        if (entry.uncompressedSize > MAX_BACKUP_BYTES || totalSize > MAX_BACKUP_BYTES) {
          return fail(new Error('El contenido descomprimido del ZIP es demasiado grande.'));
        }
        zipFile.openReadStream(entry, (streamError, stream) => {
          if (streamError || !stream) return fail(streamError ?? new Error('No se pudo leer un archivo del ZIP.'));
          const chunks: Buffer[] = [];
          let size = 0;
          stream.on('data', chunk => {
            size += chunk.length;
            if (size > entry.uncompressedSize || size > MAX_BACKUP_BYTES) stream.destroy(new Error('Un archivo del ZIP es demasiado grande.'));
            else chunks.push(Buffer.from(chunk));
          });
          stream.on('error', fail);
          stream.on('end', () => {
            if (settled) return;
            entries.push({ name: normalized, directory: false, data: Buffer.concat(chunks) });
            zipFile.readEntry();
          });
        });
      });
      zipFile.readEntry();
    });
  });
}

function textAsHtml(text: string): string {
  return text.split(/\r?\n/).map(line => `<p>${escapeHtml(line) || '<br>'}</p>`).join('');
}

/** Convierte un archivo conocido del ZIP en una nota portable. */
function physicalEntryAsNote(entry: PhysicalZipEntry): PortableNote | null {
  const extension = path.posix.extname(entry.name).toLowerCase();
  const title = path.posix.basename(entry.name, extension).slice(0, 1000) || 'Nota importada';
  const base = { title, createdAt: undefined, updatedAt: undefined };
  if (extension === '.html' || extension === '.htm') {
    const rawHtml = entry.data.toString('utf8');
    const body = rawHtml.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    return { ...base, content: body ? body[1].trim() : rawHtml.trim(), fileData: null, fileType: null };
  }
  if (extension === '.txt') {
    return { ...base, content: textAsHtml(entry.data.toString('utf8')), fileData: null, fileType: null };
  }
  if (extension === '.pdf') {
    return { ...base, content: '', fileData: entry.data.toString('base64'), fileType: 'pdf' };
  }
  if (['.xlsx', '.xls', '.csv'].includes(extension)) {
    return { ...base, content: '', fileData: entry.data.toString('base64'), fileType: 'xlsx' };
  }
  if (extension === '.pptx') {
    const presentation = parsePptxBuffer(entry.data);
    return {
      ...base,
      content: JSON.stringify(presentation),
      fileData: entry.data.toString('base64'),
      fileType: 'pptx',
    };
  }
  return null;
}

/** Reconstruye carpetas y notas cuando el ZIP no posee un manifiesto propio. */
function physicalZipToBackup(entries: PhysicalZipEntry[]): HelperNotesBackup {
  const folders = new Map<string, PortableFolder>();
  const notes: PortableNote[] = [];
  const ensureFolder = (rawName: string): PortableFolder => {
    const parts = rawName.split('/').filter(Boolean).slice(0, 50);
    const name = parts.at(-1)?.slice(0, 500) || 'Carpeta importada';
    const key = parts.join('/').toLocaleLowerCase();
    let folder = folders.get(key);
    if (!folder) {
      if (parts.length > 1) ensureFolder(parts.slice(0, -1).join('/'));
      folder = { name, parentPath: parts.slice(0, -1), notes: [] };
      folders.set(key, folder);
    }
    return folder;
  };

  for (const entry of entries) {
    const parts = entry.name.split('/').filter(Boolean);
    if (parts.length === 0 || parts[0] === '__MACOSX') continue;
    if (entry.directory) {
      ensureFolder(parts.join('/'));
      continue;
    }
    if (['.ds_store', 'thumbs.db'].includes(parts.at(-1)?.toLowerCase() ?? '')) continue;
    let note: PortableNote | null = null;
    try { note = physicalEntryAsNote(entry); }
    catch { note = null; }
    if (!note) continue;
    if (parts.length > 1) ensureFolder(parts.slice(0, -1).join('/')).notes.push(note);
    else notes.push(note);
  }

  const folderList = [...folders.values()];
  if (folderList.length === 0 && notes.length === 0) {
    throw new Error('El ZIP no contiene archivos compatibles de Helper Notes.');
  }
  return { schema: SCHEMA, version: VERSION, exportedAt: new Date().toISOString(), folders: folderList, notes };
}

/** Detecta JSON o ZIP, valida su tamaño y devuelve un único modelo portable. */
export async function readBackupFile(filePath: string): Promise<HelperNotesBackup> {
  if (!existsSync(filePath) || !statSync(filePath).isFile()) throw new Error('El archivo de respaldo no existe.');
  if (statSync(filePath).size > MAX_BACKUP_BYTES) throw new Error('El archivo de respaldo es demasiado grande.');
  const extension = path.extname(filePath).toLowerCase();
  if (extension === '.json') return parseBackupJson(readFileSync(filePath, 'utf8'));
  if (extension === '.zip') {
    const entries = await readZipEntries(filePath);
    const legacyJson = entries.find(entry => !entry.directory && path.posix.basename(entry.name).toLowerCase() === 'helper-notes-export.json');
    return legacyJson ? parseBackupJson(legacyJson.data.toString('utf8')) : physicalZipToBackup(entries);
  }
  throw new Error('Solo se pueden importar respaldos JSON o ZIP.');
}

/** Descubre respaldos compatibles dentro de un directorio seleccionado. */
export function findBackupFiles(directory: string): string[] {
  const root = path.resolve(directory);
  if (!existsSync(root) || !statSync(root).isDirectory()) throw new Error('La carpeta seleccionada no existe.');
  const files: string[] = [];
  const walk = (current: string, depth: number): void => {
    if (depth > 8 || files.length >= MAX_BACKUP_FILES) return;
    for (const entry of readdirSync(current, { withFileTypes: true })) {
      if (files.length >= MAX_BACKUP_FILES || entry.isSymbolicLink()) break;
      const fullPath = path.join(current, entry.name);
      if (entry.isDirectory()) walk(fullPath, depth + 1);
      else if (entry.isFile() && ['.json', '.zip'].includes(path.extname(entry.name).toLowerCase())) files.push(fullPath);
    }
  };
  walk(root, 0);
  return files;
}

function validDate(value?: string): string | null {
  if (!value) return null;
  const sqliteUtc = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(value);
  const date = new Date(sqliteUtc ? `${value.replace(' ', 'T')}Z` : value);
  if (Number.isNaN(date.getTime())) return null;
  return date.toISOString().slice(0, 19).replace('T', ' ');
}

/**
 * Importa varios respaldos en una transacción, reasigna identificadores para
 * evitar colisiones y devuelve un resumen con errores recuperables.
 */
export function importBackups(backups: HelperNotesBackup[]): ImportSummary {
  const db = getDb();
  const errors: string[] = [];
  let folderCount = 0;
  let noteCount = 0;
  const existingNames = new Set(getAllFolders().map(folder => `${folder.parent_id ?? 'root'}:${folder.name.toLocaleLowerCase()}`));

  const uniqueFolderName = (requested: string, parentId: number | null): string => {
    const base = requested.trim().slice(0, 500) || 'Carpeta importada';
    let candidate = base;
    let suffix = 2;
    while (existingNames.has(`${parentId ?? 'root'}:${candidate.toLocaleLowerCase()}`)) candidate = `${base} (importada ${suffix++})`;
    existingNames.add(`${parentId ?? 'root'}:${candidate.toLocaleLowerCase()}`);
    return candidate;
  };

  const insertNote = db.prepare(`
    INSERT INTO notes (folder_id, title, content, file_data, file_type, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, COALESCE(?, datetime('now')), COALESCE(?, datetime('now')))
  `);
  const insertFolder = db.prepare(`
    INSERT INTO folders (name, parent_id, created_at, updated_at)
    VALUES (?, ?, COALESCE(?, datetime('now')), COALESCE(?, datetime('now')))
  `);

  const run = db.transaction(() => {
    for (const backup of backups) {
      const importedPaths = new Map<string, number>();
      const orderedFolders = [...backup.folders].sort((a, b) => (a.parentPath?.length ?? 0) - (b.parentPath?.length ?? 0));
      for (const folder of orderedFolders) {
        const parentPath = folder.parentPath ?? [];
        const parentId = parentPath.length > 0
          ? importedPaths.get(parentPath.join('/').toLocaleLowerCase()) ?? null
          : null;
        const importedName = uniqueFolderName(folder.name, parentId);
        const result = insertFolder.run(importedName, parentId, validDate(folder.createdAt), validDate(folder.updatedAt));
        const folderId = Number(result.lastInsertRowid);
        importedPaths.set([...parentPath, folder.name].join('/').toLocaleLowerCase(), folderId);
        folderCount += 1;
        for (const note of folder.notes) {
          insertNote.run(folderId, note.title, note.content, note.fileData, note.fileType, validDate(note.createdAt), validDate(note.updatedAt));
          noteCount += 1;
        }
      }
      for (const note of backup.notes) {
        insertNote.run(null, note.title, note.content, note.fileData, note.fileType, validDate(note.createdAt), validDate(note.updatedAt));
        noteCount += 1;
      }
    }
  });

  try { run(); }
  catch (error) {
    errors.push(error instanceof Error ? error.message : String(error));
    folderCount = 0;
    noteCount = 0;
  }
  return { folders: folderCount, notes: noteCount, sources: backups.length, errors };
}

/** Lee e importa rutas independientes acumulando fallos sin abortar todo el lote. */
export async function importBackupPaths(paths: string[]): Promise<ImportSummary> {
  const backups: HelperNotesBackup[] = [];
  const errors: string[] = [];
  for (const filePath of paths.slice(0, MAX_BACKUP_FILES)) {
    try { backups.push(await readBackupFile(filePath)); }
    catch (error) { errors.push(`${path.basename(filePath)}: ${error instanceof Error ? error.message : String(error)}`); }
  }
  const result = backups.length > 0 ? importBackups(backups) : { folders: 0, notes: 0, sources: 0, errors: [] };
  return { ...result, errors: [...errors, ...result.errors] };
}
