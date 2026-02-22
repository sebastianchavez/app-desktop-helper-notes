import { app, BrowserWindow, Menu, Tray, nativeImage } from 'electron';
import path from 'path';
import { initDb } from './db/db';
import { registerIpcHandlers } from './ipc-handlers';

let win: BrowserWindow | null = null;
let tray: Tray | null = null;
let isQuitting = false;

function createTray() {
    const iconPath = path.join(__dirname, '../assets/favicon.ico');
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

function createWindow() {
    win = new BrowserWindow({
        width: 1280,
        height: 800,
        minWidth: 800,
        minHeight: 600,
        title: "Helper Notes",
        icon: path.join(__dirname, '../assets/favicon.ico'),
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
        win.loadFile(
            path.join(__dirname, '../dist/app-desktop-helper-notes/browser/index.html')
        );
    }
    win.setMenu(null);

    // Minimizar al tray en vez de cerrar
    win.on('close', (event) => {
        if (!isQuitting) {
            event.preventDefault();
            win?.hide();
        }
    });
}

app.whenReady().then(async () => {
    initDb();
    registerIpcHandlers();
    createTray();
    createWindow();
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
