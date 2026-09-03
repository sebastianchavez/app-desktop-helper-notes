import { app } from 'electron';
import { ChildProcess, execFile, spawn } from 'child_process';
import { createReadStream, createWriteStream, existsSync, mkdirSync, readdirSync, statSync, unlinkSync } from 'fs';
import { createHash } from 'crypto';
import http from 'http';
import https from 'https';
import path from 'path';
import yauzl from 'yauzl';

export interface LlamaProgress {
  operation: 'runtime' | 'model';
  id: string;
  received: number;
  total: number;
  percent: number;
  message: string;
}

export interface LlamaCatalogModel {
  id: string;
  downloads: number;
  likes: number;
  updatedAt: string;
}

export interface LlamaModelFile {
  name: string;
  size: number;
  quantization: string;
}

let llamaProcess: ChildProcess | null = null;
let managedModelPath: string | null = null;
let managedEndpoint: string | null = null;
let activeDownload: http.ClientRequest | null = null;

/** Directorio privado donde la aplicación administra runtime y modelos. */
function rootDir(): string {
  return path.join(app.getPath('userData'), 'llama.cpp');
}

function runtimeDir(): string {
  return path.join(rootDir(), 'runtime');
}

function modelsDir(): string {
  return path.join(rootDir(), 'models');
}

/** Crea de forma idempotente la estructura necesaria bajo `userData`. */
function ensureDirectories(): void {
  mkdirSync(runtimeDir(), { recursive: true });
  mkdirSync(modelsDir(), { recursive: true });
}

/** Busca un archivo por nombre dentro de un árbol ya controlado por la aplicación. */
function findFile(directory: string, fileName: string): string | null {
  if (!existsSync(directory)) return null;
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const fullPath = path.join(directory, entry.name);
    if (entry.isFile() && entry.name.toLowerCase() === fileName.toLowerCase()) return fullPath;
    if (entry.isDirectory()) {
      const nested = findFile(fullPath, fileName);
      if (nested) return nested;
    }
  }
  return null;
}

/** Extrae un ZIP rechazando rutas absolutas o que escapen del directorio destino. */
function extractZipSafely(archivePath: string, destination: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const destinationRoot = path.resolve(destination);
    let settled = false;
    const fail = (error: Error, zipFile?: yauzl.ZipFile): void => {
      if (settled) return;
      settled = true;
      zipFile?.close();
      reject(error);
    };

    yauzl.open(archivePath, { lazyEntries: true, validateEntrySizes: true }, (openError, zipFile) => {
      if (openError || !zipFile) {
        fail(openError ?? new Error('No se pudo abrir el archivo ZIP'));
        return;
      }
      zipFile.on('error', error => fail(error, zipFile));
      zipFile.on('end', () => {
        if (!settled) {
          settled = true;
          resolve();
        }
      });
      zipFile.on('entry', entry => {
        const normalized = entry.fileName.replace(/\\/g, '/');
        const segments = normalized.split('/').filter(Boolean);
        const unixMode = (entry.externalFileAttributes >>> 16) & 0xffff;
        const isSymlink = (unixMode & 0o170000) === 0o120000;
        if (
          !normalized || normalized.startsWith('/') || /^[a-zA-Z]:/.test(normalized) ||
          segments.includes('..') || isSymlink
        ) {
          fail(new Error(`Entrada ZIP insegura: ${entry.fileName}`), zipFile);
          return;
        }

        const targetPath = path.resolve(destinationRoot, ...segments);
        if (targetPath !== destinationRoot && !targetPath.startsWith(destinationRoot + path.sep)) {
          fail(new Error(`La entrada ZIP sale del directorio permitido: ${entry.fileName}`), zipFile);
          return;
        }

        const isDirectory = normalized.endsWith('/') || (unixMode & 0o170000) === 0o040000;
        if (isDirectory) {
          mkdirSync(targetPath, { recursive: true });
          zipFile.readEntry();
          return;
        }

        mkdirSync(path.dirname(targetPath), { recursive: true });
        zipFile.openReadStream(entry, (streamError, readStream) => {
          if (streamError || !readStream) {
            fail(streamError ?? new Error(`No se pudo extraer ${entry.fileName}`), zipFile);
            return;
          }
          const output = createWriteStream(targetPath, { flags: 'w', mode: 0o755 });
          readStream.on('error', error => fail(error, zipFile));
          output.on('error', error => fail(error, zipFile));
          output.on('finish', () => zipFile.readEntry());
          readStream.pipe(output);
        });
      });
      zipFile.readEntry();
    });
  });
}

/** Localiza el binario llama-server instalado para la plataforma actual. */
export function llamaExecutablePath(): string | null {
  return findFile(runtimeDir(), process.platform === 'win32' ? 'llama-server.exe' : 'llama-server');
}

/** Realiza solicitudes JSON acotadas usadas por los catálogos remotos. */
function requestJson(url: string, headers: Record<string, string> = {}): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const requestUrl = new URL(url);
    const transport = requestUrl.protocol === 'https:' ? https : http;
    const request = transport.get(requestUrl, {
      headers: { 'User-Agent': 'Helper-Notes', Accept: 'application/json', ...headers },
      timeout: 15000,
    }, response => {
      if ((response.statusCode === 301 || response.statusCode === 302 || response.statusCode === 307 || response.statusCode === 308) && response.headers.location) {
        response.resume();
        requestJson(new URL(response.headers.location, requestUrl).toString(), headers).then(resolve, reject);
        return;
      }
      let data = '';
      response.setEncoding('utf8');
      response.on('data', chunk => data += chunk);
      response.on('end', () => {
        if ((response.statusCode ?? 500) >= 400) {
          reject(new Error(`HTTP ${response.statusCode}`));
          return;
        }
        try { resolve(JSON.parse(data)); }
        catch { reject(new Error('El servidor devolvió una respuesta no válida')); }
      });
    });
    request.on('timeout', () => request.destroy(new Error('La solicitud superó el tiempo de espera')));
    request.on('error', reject);
  });
}

/** Descarga a disco con progreso, cancelación y limpieza del archivo parcial. */
function downloadFile(
  url: string,
  targetPath: string,
  operation: LlamaProgress['operation'],
  id: string,
  onProgress: (progress: LlamaProgress) => void,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const start = (downloadUrl: URL): void => {
      const transport = downloadUrl.protocol === 'https:' ? https : http;
      const request = transport.get(downloadUrl, {
        headers: { 'User-Agent': 'Helper-Notes', Accept: 'application/octet-stream' },
      }, response => {
        if ((response.statusCode === 301 || response.statusCode === 302 || response.statusCode === 307 || response.statusCode === 308) && response.headers.location) {
          response.resume();
          start(new URL(response.headers.location, downloadUrl));
          return;
        }
        if ((response.statusCode ?? 500) >= 400) {
          response.resume();
          reject(new Error(`La descarga respondió con HTTP ${response.statusCode}`));
          return;
        }
        const total = Number(response.headers['content-length'] ?? 0);
        let received = 0;
        const stream = createWriteStream(targetPath);
        response.on('data', chunk => {
          received += chunk.length;
          onProgress({
            operation,
            id,
            received,
            total,
            percent: total > 0 ? Math.round((received / total) * 100) : 0,
            message: 'Descargando',
          });
        });
        response.pipe(stream);
        stream.on('finish', () => {
          stream.close();
          activeDownload = null;
          resolve();
        });
        stream.on('error', error => {
          activeDownload = null;
          if (existsSync(targetPath)) unlinkSync(targetPath);
          reject(error);
        });
        response.on('aborted', () => {
          stream.destroy();
          activeDownload = null;
          if (existsSync(targetPath)) unlinkSync(targetPath);
          reject(new Error('Descarga cancelada'));
        });
      });
      activeDownload = request;
      request.on('error', error => {
        activeDownload = null;
        if (existsSync(targetPath)) unlinkSync(targetPath);
        reject(error);
      });
    };
    start(new URL(url));
  });
}

/** Calcula SHA-256 por streaming para validar binarios descargados. */
function sha256File(filePath: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const hash = createHash('sha256');
    const stream = createReadStream(filePath);
    stream.on('data', chunk => hash.update(chunk));
    stream.on('error', reject);
    stream.on('end', () => resolve(hash.digest('hex')));
  });
}

/** Descarga, verifica y extrae la versión soportada de llama.cpp. */
export async function installLlamaCpp(onProgress: (progress: LlamaProgress) => void): Promise<{ success: boolean; path?: string; version?: string; error?: string }> {
  if (process.platform !== 'win32' || process.arch !== 'x64') {
    return { success: false, error: 'La instalación automática está disponible actualmente solo para Windows x64.' };
  }
  try {
    ensureDirectories();
    const recentReleases = await requestJson('https://api.github.com/repos/ggml-org/llama.cpp/releases?per_page=30') as Array<Record<string, unknown>>;
    const compatibleRelease = recentReleases.find(item => {
      const releaseAssets = Array.isArray(item['assets']) ? item['assets'] as Array<Record<string, unknown>> : [];
      return releaseAssets.some(asset => /^llama-.+-bin-win-cpu-x64\.zip$/i.test(String(asset['name'] ?? '')));
    });
    if (!compatibleRelease) throw new Error('No se encontró un paquete oficial reciente de llama.cpp para Windows x64 CPU.');
    const assets = compatibleRelease['assets'] as Array<Record<string, unknown>>;
    const asset = assets.find(item => /^llama-.+-bin-win-cpu-x64\.zip$/i.test(String(item['name'] ?? '')))!;
    const downloadUrl = String(asset['browser_download_url'] ?? '');
    const archivePath = path.join(rootDir(), 'llama-runtime.zip');
    const version = String(compatibleRelease['tag_name'] ?? 'latest');
    await downloadFile(downloadUrl, archivePath, 'runtime', version, onProgress);
    const publishedDigest = String(asset['digest'] ?? '');
    if (publishedDigest.startsWith('sha256:')) {
      const actualDigest = await sha256File(archivePath);
      if (actualDigest.toLowerCase() !== publishedDigest.slice(7).toLowerCase()) {
        unlinkSync(archivePath);
        throw new Error('La verificación SHA-256 del paquete de llama.cpp falló.');
      }
    }
    onProgress({ operation: 'runtime', id: version, received: 1, total: 1, percent: 100, message: 'Instalando archivos' });
    await extractZipSafely(archivePath, runtimeDir());
    if (existsSync(archivePath)) unlinkSync(archivePath);
    const executable = llamaExecutablePath();
    if (!executable) throw new Error('La instalación terminó, pero no se encontró llama-server.');
    return { success: true, path: executable, version };
  } catch (error) {
    return { success: false, error: error instanceof Error ? error.message : String(error) };
  }
}

/** Cancela la descarga activa mediante su AbortController compartido. */
export function cancelLlamaDownload(): void {
  activeDownload?.destroy();
  activeDownload = null;
}

/** Busca repositorios compatibles en el catálogo y normaliza sus métricas. */
export async function searchLlamaModels(query: string): Promise<{ success: boolean; models: LlamaCatalogModel[]; error?: string }> {
  try {
    const search = encodeURIComponent(query.trim() || 'GGUF');
    const url = `https://huggingface.co/api/models?search=${search}&filter=gguf&pipeline_tag=text-generation&sort=downloads&direction=-1&limit=30`;
    const data = await requestJson(url) as Array<Record<string, unknown>>;
    const models = data.map(item => ({
      id: String(item['id'] ?? item['modelId'] ?? ''),
      downloads: Number(item['downloads'] ?? 0),
      likes: Number(item['likes'] ?? 0),
      updatedAt: String(item['lastModified'] ?? ''),
    })).filter(item => item.id);
    return { success: true, models };
  } catch (error) {
    return { success: false, models: [], error: error instanceof Error ? error.message : String(error) };
  }
}

function quantizationFromName(fileName: string): string {
  const match = fileName.match(/(?:^|[-_.])(IQ\d(?:_[A-Z0-9]+)?|Q\d(?:_[A-Z0-9]+)?|F16|BF16)(?:[-_.]|$)/i);
  return match?.[1]?.toUpperCase() ?? 'GGUF';
}

/** Lista únicamente archivos GGUF descargables de un repositorio. */
export async function listLlamaModelFiles(repository: string): Promise<{ success: boolean; files: LlamaModelFile[]; error?: string }> {
  if (!/^[\w.-]+\/[\w.-]+$/.test(repository)) return { success: false, files: [], error: 'Repositorio no válido' };
  try {
    const data = await requestJson(`https://huggingface.co/api/models/${repository}?blobs=true`) as Record<string, unknown>;
    const siblings = Array.isArray(data['siblings']) ? data['siblings'] as Array<Record<string, unknown>> : [];
    const files = siblings.map(item => {
      const lfs = item['lfs'] as Record<string, unknown> | undefined;
      return {
        name: String(item['rfilename'] ?? ''),
        size: Number(item['size'] ?? lfs?.['size'] ?? 0),
        quantization: quantizationFromName(String(item['rfilename'] ?? '')),
      };
    }).filter(item => item.name.toLowerCase().endsWith('.gguf') && !item.name.includes('/'))
      .sort((a, b) => a.size - b.size);
    return { success: true, files };
  } catch (error) {
    return { success: false, files: [], error: error instanceof Error ? error.message : String(error) };
  }
}

/** Descarga un GGUF a la biblioteca local conservando información de progreso. */
export async function downloadLlamaModel(
  repository: string,
  fileName: string,
  onProgress: (progress: LlamaProgress) => void,
): Promise<{ success: boolean; path?: string; error?: string }> {
  if (!/^[\w.-]+\/[\w.-]+$/.test(repository) || !/^[\w .()+-]+\.gguf$/i.test(fileName)) {
    return { success: false, error: 'Modelo o archivo no válido' };
  }
  try {
    ensureDirectories();
    const targetPath = path.join(modelsDir(), fileName);
    const encodedFile = encodeURIComponent(fileName).replace(/%2F/gi, '/');
    const url = `https://huggingface.co/${repository}/resolve/main/${encodedFile}?download=true`;
    await downloadFile(url, targetPath, 'model', `${repository}/${fileName}`, onProgress);
    return { success: true, path: targetPath };
  } catch (error) {
    return { success: false, error: error instanceof Error ? error.message : String(error) };
  }
}

/** Enumera modelos locales sin exponer directorios arbitrarios del usuario. */
export function listLocalLlamaModels(): { name: string; path: string; size: number }[] {
  ensureDirectories();
  return readdirSync(modelsDir(), { withFileTypes: true })
    .filter(item => item.isFile() && item.name.toLowerCase().endsWith('.gguf'))
    .map(item => {
      const fullPath = path.join(modelsDir(), item.name);
      return { name: item.name, path: fullPath, size: statSync(fullPath).size };
    });
}

/** Comprueba con timeout si llama-server responde en el endpoint indicado. */
function healthCheck(endpoint: string): Promise<boolean> {
  return new Promise(resolve => {
    try {
      const target = new URL('/health', endpoint);
      const transport = target.protocol === 'https:' ? https : http;
      const request = transport.get(target, { timeout: 1000 }, response => {
        response.resume();
        resolve((response.statusCode ?? 500) < 500);
      });
      request.on('timeout', () => { request.destroy(); resolve(false); });
      request.on('error', () => resolve(false));
    } catch { resolve(false); }
  });
}

/** Resume instalación, proceso y modelos para la pantalla de configuración. */
export async function llamaManagerStatus(endpoint: string): Promise<{ installed: boolean; running: boolean; executable?: string; models: ReturnType<typeof listLocalLlamaModels> }> {
  const executable = llamaExecutablePath();
  return {
    installed: !!executable,
    running: await healthCheck(endpoint),
    ...(executable ? { executable } : {}),
    models: listLocalLlamaModels(),
  };
}

/**
 * Arranca un único llama-server con una ruta de modelo validada dentro de la
 * biblioteca administrada y espera hasta que el endpoint quede saludable.
 */
export async function startManagedLlamaServer(modelPath: string, endpoint: string): Promise<{ success: boolean; error?: string }> {
  try {
    const executable = llamaExecutablePath();
    if (!executable) throw new Error('llama.cpp no está instalado');
    const resolvedModel = path.resolve(modelPath);
    const allowedModelsDir = path.resolve(modelsDir()) + path.sep;
    if (!resolvedModel.startsWith(allowedModelsDir) || !existsSync(resolvedModel)) throw new Error('El modelo no pertenece al catálogo local');
    const target = new URL(endpoint);
    const port = Number(target.port || 8080);
    if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error('Puerto no válido');
    if (llamaProcess && !llamaProcess.killed) {
      if (managedModelPath === resolvedModel && managedEndpoint === endpoint) return { success: true };
      const stopped = await stopManagedLlamaServer();
      if (!stopped.success) throw new Error(stopped.error ?? 'No se pudo cambiar el modelo activo');
    }
    const processInstance = spawn(executable, [
      '-m', resolvedModel,
      '--alias', path.basename(resolvedModel),
      '--host', '127.0.0.1',
      '--port', String(port),
      '--ctx-size', '4096',
    ], {
      cwd: path.dirname(executable),
      windowsHide: true,
      stdio: 'ignore',
    });
    llamaProcess = processInstance;
    managedModelPath = resolvedModel;
    managedEndpoint = endpoint;
    processInstance.on('exit', () => {
      if (llamaProcess === processInstance) {
        llamaProcess = null;
        managedModelPath = null;
        managedEndpoint = null;
      }
    });
    for (let attempt = 0; attempt < 120; attempt++) {
      await new Promise(resolve => setTimeout(resolve, 500));
      if (await healthCheck(endpoint)) return { success: true };
      if (!llamaProcess) break;
    }
    throw new Error('llama-server no respondió dentro del tiempo esperado');
  } catch (error) {
    return { success: false, error: error instanceof Error ? error.message : String(error) };
  }
}

/** Finaliza el proceso administrado y limpia su referencia global. */
export async function stopManagedLlamaServer(): Promise<{ success: boolean; error?: string }> {
  if (!llamaProcess?.pid) return { success: true };
  const processInstance = llamaProcess;
  const pid = processInstance.pid!;
  return new Promise(resolve => {
    if (process.platform === 'win32') {
      execFile('taskkill', ['/PID', String(pid), '/T', '/F'], { windowsHide: true }, error => {
        if (llamaProcess === processInstance) llamaProcess = null;
        managedModelPath = null;
        managedEndpoint = null;
        resolve(error ? { success: false, error: error.message } : { success: true });
      });
    } else {
      const stopped = processInstance.kill('SIGTERM');
      if (llamaProcess === processInstance) llamaProcess = null;
      managedModelPath = null;
      managedEndpoint = null;
      resolve(stopped ? { success: true } : { success: false, error: 'No se pudo detener llama-server' });
    }
  });
}
