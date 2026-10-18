import { BrowserWindow, Menu } from 'electron';

/**
 * Muestra acciones nativas que delegan la selección al editor Word, también en ventanas independientes.
 * @param window Ventana propietaria del clic derecho.
 * @param x Coordenada horizontal CSS del clic.
 * @param y Coordenada vertical CSS del clic.
 * @returns true si el clic pertenece a Word; false para continuar con el menú habitual.
 */
export async function showWordClipboardMenu(
  window: BrowserWindow,
  x: number,
  y: number,
): Promise<boolean> {
  if (window.isDestroyed() || !Number.isFinite(x) || !Number.isFinite(y)) return false;
  const inWord = await window.webContents
    .executeJavaScript(
      `(() => {
    const element = document.elementFromPoint(${x}, ${y});
    if (element?.closest('.docx-contextmenu-host')) return { word: true, custom: true };
    const host = element?.closest('[data-word-editor-surface]');
    return { word: !!host && !!element?.closest('.docx-paginated-surface'), custom: host?.dataset.wordOwnsContextMenu === 'true' };
  })()`,
    )
    .catch(() => false);
  if (!inWord || (inWord !== true && !inWord.word) || window.isDestroyed()) return false;
  // El menú propio de Word mantiene acciones de tablas, imágenes y formato; evitar un segundo menú nativo.
  if (inWord.custom) return true;
  const command = (action: 'copy' | 'cut' | 'paste' | 'pastePlain'): void => {
    if (!window.isDestroyed()) window.webContents.send('word:clipboard-command', action);
  };
  Menu.buildFromTemplate([
    { label: 'Cortar', click: () => command('cut') },
    { label: 'Copiar', click: () => command('copy') },
    { label: 'Pegar', click: () => command('paste') },
    { label: 'Pegar sin formato', click: () => command('pastePlain') },
  ]).popup({ window });
  return true;
}
