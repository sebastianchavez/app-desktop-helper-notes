import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { BrowserWindow, MenuItemConstructorOptions } from 'electron';

const menu = vi.hoisted(() => ({ popup: vi.fn(), template: [] as MenuItemConstructorOptions[] }));
vi.mock('electron', () => ({
  Menu: {
    buildFromTemplate: (template: MenuItemConstructorOptions[]) => {
      menu.template = template;
      return { popup: menu.popup };
    },
  },
}));
import { showWordClipboardMenu } from '../electron/word-context-menu';

describe('native Word clipboard menu', () => {
  beforeEach(() => {
    menu.popup.mockClear();
    menu.template = [];
  });

  it('routes every action to the owning Word renderer instead of browser DOM editing roles', async () => {
    const send = vi.fn();
    const window = {
      isDestroyed: () => false,
      webContents: { executeJavaScript: async () => true, send },
    } as unknown as BrowserWindow;
    expect(await showWordClipboardMenu(window, 10, 20)).toBe(true);
    expect(menu.popup).toHaveBeenCalledWith({ window });
    const commands = ['cut', 'copy', 'paste', 'pastePlain'];
    menu.template.forEach((item, index) => {
      expect(item.role).toBeUndefined();
      (item.click as () => void)();
      expect(send).toHaveBeenLastCalledWith('word:clipboard-command', commands[index]);
    });
  });

  it('leaves other editors and closed windows to their existing behavior', async () => {
    const window = {
      isDestroyed: () => false,
      webContents: { executeJavaScript: async () => false },
    } as unknown as BrowserWindow;
    expect(await showWordClipboardMenu(window, 10, 20)).toBe(false);
    expect(menu.popup).not.toHaveBeenCalled();
    expect(await showWordClipboardMenu({ isDestroyed: () => true } as BrowserWindow, 10, 20)).toBe(
      false,
    );
    expect(await showWordClipboardMenu(window, NaN, 20)).toBe(false);
  });

  it('does not cover the Word menu with a second native menu', async () => {
    const window = {
      isDestroyed: () => false,
      webContents: { executeJavaScript: async () => ({ word: true, custom: true }) },
    } as unknown as BrowserWindow;
    expect(await showWordClipboardMenu(window, 10, 20)).toBe(true);
    expect(menu.popup).not.toHaveBeenCalled();
  });
});
