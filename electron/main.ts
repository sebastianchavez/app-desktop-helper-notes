import 'dotenv/config';
import { app, BrowserWindow, Menu, Tray, nativeImage, shell } from 'electron';
import path from 'path';
import { initDb } from './db/db';
import { registerIpcHandlers } from './ipc-handlers';
import { UpdateManager } from './update-manager';
import { applicationIconPath } from './platform-assets';

let win: BrowserWindow | null = null;
let tray: Tray | null = null;
let isQuitting = false;
let updateManager: UpdateManager | null = null;

/** Crea el icono de bandeja y sus acciones para restaurar o cerrar la aplicación. */
function createTray() {
  const iconPath = applicationIconPath();
  tray = new Tray(nativeImage.createFromPath(iconPath));

  const contextMenu = Menu.buildFromTemplate([
    {
      label: 'Abrir Helper Notes',
      click: () => {
        if (win) {
          win.show();
          win.focus();
        }
      },
    },
    { type: 'separator' },
    {
      label: 'Salir',
      click: () => {
        isQuitting = true;
        app.quit();
      },
    },
  ]);

  tray.setToolTip('Helper Notes');
  tray.setContextMenu(contextMenu);

  tray.on('double-click', () => {
    if (win) {
      win.show();
      win.focus();
    }
  });
}

/**
 * Crea la ventana principal con aislamiento de contexto y sin Node en el
 * renderer. En desarrollo carga el servidor Angular; empaquetada usa `dist`.
 */
function createWindow() {
  win = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    title: 'Helper Notes',
    icon: applicationIconPath(),
    resizable: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  const devUrl = process.env['ELECTRON_DEV_URL'];
  if (devUrl) {
    win.loadURL(devUrl);
    win.webContents.openDevTools({ mode: 'detach' });
  } else {
    // PROD: carga dist angular
    win.loadFile(path.join(__dirname, '../dist/app-desktop-helper-notes/browser/index.html'));
  }
  win.setMenu(null);

  // Menú contextual nativo (clic derecho: copiar, cortar, pegar, etc.)
  win.webContents.on('context-menu', (_event, params) => {
    const menuTemplate: Electron.MenuItemConstructorOptions[] = [];

    if (params.isEditable) {
      menuTemplate.push(
        { label: 'Deshacer', role: 'undo' },
        { label: 'Rehacer', role: 'redo' },
        { type: 'separator' },
        { label: 'Cortar', role: 'cut' },
        { label: 'Copiar', role: 'copy' },
        { label: 'Pegar', role: 'paste' },
        { label: 'Seleccionar todo', role: 'selectAll' },
      );
    } else {
      if (params.selectionText) {
        menuTemplate.push({ label: 'Copiar', role: 'copy' });
      }
      if (params.linkURL) {
        menuTemplate.push(
          {
            label: 'Copiar enlace',
            click: () => {
              require('electron').clipboard.writeText(params.linkURL);
            },
          },
          {
            label: 'Abrir enlace',
            click: () => {
              shell.openExternal(params.linkURL);
            },
          },
        );
      }
      if (!params.selectionText && !params.linkURL) {
        menuTemplate.push({ label: 'Seleccionar todo', role: 'selectAll' });
      }
    }

    if (menuTemplate.length > 0) {
      const contextMenu = Menu.buildFromTemplate(menuTemplate);
      contextMenu.popup();
    }
  });

  // Minimizar al tray en vez de cerrar
  win.on('close', (event) => {
    if (!isQuitting) {
      event.preventDefault();
      win?.hide();
    }
  });
}

// Una sola instancia evita dos conexiones SQLite y dos iconos de bandeja.
const gotTheLock = app.requestSingleInstanceLock();

if (!gotTheLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (win) {
      if (win.isMinimized()) win.restore();
      win.show();
      win.focus();
    }
  });

  app.whenReady().then(async () => {
    // El orden garantiza esquema e IPC disponibles antes de renderizar Angular.
    initDb();
    registerIpcHandlers();
    createTray();
    createWindow();
    updateManager = new UpdateManager(() => win);
    updateManager.start();
  });

  app.on('before-quit', () => {
    isQuitting = true;
  });

  app.on('window-all-closed', () => {
    // No hacer nada, la app vive en el tray
  });

  app.on('activate', () => {
    if (win) {
      win.show();
    } else {
      createWindow();
    }
  });
}
