import { app, BrowserWindow, dialog, ipcMain, shell } from 'electron';
import { createHash } from 'crypto';
import { createReadStream, createWriteStream, promises as fs } from 'fs';
import { Readable } from 'stream';
import { pipeline } from 'stream/promises';
import { spawn } from 'child_process';
import path from 'path';

export interface UpdateInfo {
  version: string;
  releaseNotes: string;
  downloadUrl: string;
  checksumSha256: string;
  fileName: string;
  fileSize: number;
  mandatory: boolean;
}
export interface UpdateState {
  status: 'idle' | 'checking' | 'available' | 'downloading' | 'downloaded' | 'up-to-date' | 'error';
  currentVersion: string;
  update?: UpdateInfo;
  progress?: number;
  filePath?: string;
  error?: string;
}

/**
 * Coordina la consulta, descarga, validación e instalación de actualizaciones.
 * Vive en el proceso principal para mantener el acceso a archivos y procesos
 * fuera del renderer de Angular.
 */
export class UpdateManager {
  private state: UpdateState = { status: 'idle', currentVersion: app.getVersion() };
  private timer?: NodeJS.Timeout;
  private checking?: Promise<UpdateState>;
  private readonly apiUrl = (
    process.env['HELPER_NOTES_UPDATE_API_URL'] || 'http://127.0.0.1:3000'
  ).replace(/\/$/, '');
  private readonly channel = process.env['HELPER_NOTES_UPDATE_CHANNEL'] || 'stable';
  private readonly enabled = process.env['HELPER_NOTES_AUTO_UPDATE_ENABLED'] !== 'false';
  private readonly autoDownload = process.env['HELPER_NOTES_AUTO_DOWNLOAD_UPDATES'] !== 'false';

  constructor(private readonly window: () => BrowserWindow | null) {
    // El preload expone estos handlers al renderer sin habilitar nodeIntegration.
    ipcMain.handle('updates:getState', () => this.state);
    ipcMain.handle('updates:check', () => this.check(true));
    ipcMain.handle('updates:download', () => this.download());
    ipcMain.handle('updates:install', () => this.install());
  }

  /** Programa una comprobación inicial y verificaciones periódicas en segundo plano. */
  start(): void {
    if (!this.enabled) return;
    const initial = Number(process.env['HELPER_NOTES_UPDATE_INITIAL_DELAY_MS'] || 5000);
    const interval = Math.max(
      300000,
      Number(process.env['HELPER_NOTES_UPDATE_INTERVAL_MS'] || 21600000),
    );
    setTimeout(() => void this.check(false), initial).unref();
    this.timer = setInterval(() => void this.check(false), interval);
    this.timer.unref();
  }

  /**
   * Reutiliza la comprobación activa para impedir solicitudes duplicadas cuando
   * coinciden el temporizador y una acción manual del usuario.
   */
  async check(manual = false): Promise<UpdateState> {
    if (this.checking) return this.checking;
    this.checking = this.performCheck(manual).finally(() => {
      this.checking = undefined;
    });
    return this.checking;
  }

  /** Consulta la versión correspondiente a plataforma, arquitectura y canal actuales. */
  private async performCheck(manual: boolean): Promise<UpdateState> {
    this.setState({ status: 'checking', currentVersion: app.getVersion() });
    try {
      const url = new URL(`${this.apiUrl}/updates/latest`);
      url.searchParams.set('platform', process.platform);
      url.searchParams.set('arch', process.arch);
      url.searchParams.set('channel', this.channel);
      url.searchParams.set('currentVersion', app.getVersion());
      const response = await fetch(url, {
        headers: { accept: 'application/json' },
        signal: AbortSignal.timeout(15000),
      });
      if (response.status === 204) {
        this.setState({ status: 'up-to-date', currentVersion: app.getVersion() });
        if (manual) await this.inform('Helper Notes está actualizado.');
        return this.state;
      }
      if (!response.ok)
        throw new Error(`El servidor de actualizaciones respondió ${response.status}.`);
      const update = (await response.json()) as UpdateInfo;
      if (!update.downloadUrl)
        throw new Error('La versión publicada no contiene una URL de descarga.');
      this.setState({ status: 'available', currentVersion: app.getVersion(), update });
      if (this.autoDownload && app.isPackaged) await this.download();
      else if (manual) await this.inform(`La versión ${update.version} está disponible.`);
      return this.state;
    } catch (error) {
      this.setState({
        status: 'error',
        currentVersion: app.getVersion(),
        error: error instanceof Error ? error.message : String(error),
      });
      if (manual)
        await this.inform(this.state.error || 'No fue posible buscar actualizaciones.', 'error');
      return this.state;
    }
  }

  /** Descarga el instalador y verifica su SHA-256 antes de ofrecer la instalación. */
  async download(): Promise<UpdateState> {
    const update = this.state.update;
    if (!update) throw new Error('No hay una actualización disponible.');
    this.setState({ ...this.state, status: 'downloading', progress: 0 });
    try {
      const response = await fetch(update.downloadUrl, { redirect: 'follow' });
      if (!response.ok || !response.body)
        throw new Error(`No fue posible descargar el instalador (${response.status}).`);
      const directory = path.join(app.getPath('temp'), 'helper-notes-updates');
      await fs.mkdir(directory, { recursive: true });
      const fileName =
        path.basename(update.fileName || new URL(update.downloadUrl).pathname) ||
        `helper-notes-${update.version}.exe`;
      const filePath = path.join(directory, fileName);

      // El stream evita cargar instaladores grandes completos en memoria.
      await pipeline(Readable.fromWeb(response.body as never), createWriteStream(filePath));
      if (update.checksumSha256) {
        const actual = await this.sha256(filePath);
        if (actual.toLowerCase() !== update.checksumSha256.toLowerCase()) {
          await fs.unlink(filePath).catch(() => undefined);
          throw new Error('La firma SHA-256 del instalador no coincide.');
        }
      }
      this.setState({ ...this.state, status: 'downloaded', progress: 100, filePath });
      await this.promptInstall(update);
      return this.state;
    } catch (error) {
      this.setState({
        ...this.state,
        status: 'error',
        error: error instanceof Error ? error.message : String(error),
      });
      return this.state;
    }
  }

  /**
   * En Windows ejecuta NSIS en modo silencioso. En otros sistemas delega la
   * apertura al sistema operativo para respetar el formato de paquete nativo.
   */
  async install(): Promise<{ success: boolean; error?: string }> {
    const filePath = this.state.filePath;
    if (!filePath) return { success: false, error: 'La actualización aún no se ha descargado.' };
    try {
      if (process.platform === 'win32') {
        spawn(filePath, ['/S'], { detached: true, stdio: 'ignore' }).unref();
        app.quit();
      } else {
        const error = await shell.openPath(filePath);
        if (error) throw new Error(error);
      }
      return { success: true };
    } catch (error) {
      return { success: false, error: error instanceof Error ? error.message : String(error) };
    }
  }

  /** Presenta la decisión de instalación y respeta si la actualización es obligatoria. */
  private async promptInstall(update: UpdateInfo): Promise<void> {
    const owner = this.window();
    const options = {
      type: 'info' as const,
      title: 'Actualización lista',
      message: `Helper Notes ${update.version} está listo para instalar.`,
      detail: update.releaseNotes || 'La aplicación se cerrará para completar la instalación.',
      buttons: ['Instalar ahora', update.mandatory ? 'Cerrar' : 'Más tarde'],
      defaultId: 0,
      cancelId: 1,
      noLink: true,
    };
    const result = owner
      ? await dialog.showMessageBox(owner, options)
      : await dialog.showMessageBox(options);
    if (result.response === 0) await this.install();
  }

  /** Muestra un diálogo asociado a la ventana principal cuando está disponible. */
  private async inform(message: string, type: 'info' | 'error' = 'info'): Promise<void> {
    const owner = this.window();
    const options = { type, title: 'Actualizaciones', message, buttons: ['Aceptar'] };
    if (owner) await dialog.showMessageBox(owner, options);
    else await dialog.showMessageBox(options);
  }

  /** Calcula el hash desde disco después de completar la descarga. */
  private async sha256(filePath: string): Promise<string> {
    const hash = createHash('sha256');
    await pipeline(createReadStream(filePath), hash);
    return hash.digest('hex');
  }

  /** Publica cada transición para que una futura vista pueda mostrar el progreso. */
  private setState(state: UpdateState): void {
    this.state = state;
    this.window()?.webContents.send('updates:state', state);
  }
}
