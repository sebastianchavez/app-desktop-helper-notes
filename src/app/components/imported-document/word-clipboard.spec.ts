import type { DocxEditorInstance, EditorScope } from '@docx-editor.dev/core';
import { bindWordClipboard, WordClipboardCommand } from './word-clipboard';

describe('Word clipboard selection', () => {
  let host: HTMLDivElement;
  let surface: HTMLDivElement;
  let scope: EditorScope;
  let selection: {
    anchor: { paragraphId: string; offset: number };
    head: { paragraphId: string; offset: number };
  };
  let command: (command: WordClipboardCommand) => void;
  let dispose: () => void;
  const exec = vi.fn(() => ({ ok: true, changed: true }));
  const focus = vi.fn();
  const releaseSelection = vi.fn();
  const native = vi.fn(async (_command: string) => undefined);
  const report = vi.fn();
  const pasteRich = vi.fn(() => true);
  const armForcePlainPaste = vi.fn();

  beforeEach(() => {
    vi.clearAllMocks();
    scope = { kind: 'body' };
    selection = {
      anchor: { paragraphId: 'paragraph', offset: 4 },
      head: { paragraphId: 'paragraph', offset: 8 },
    };
    host = document.createElement('div');
    surface = document.createElement('div');
    surface.className = 'docx-paginated-surface';
    host.append(surface);
    document.body.append(host);
    Object.defineProperty(window, 'electronAPI', {
      configurable: true,
      value: {
        onWordClipboardCommand: (listener: typeof command) => {
          command = listener;
          return () => undefined;
        },
        wordClipboardExecute: native,
      },
    });
    const editor = {
      exec,
      focus,
      releaseSelection,
      getSelectedImage: () => null,
      retainSelection: () => ({ pin: true }),
      snapshot: () => ({ scope, editable: true }),
      surface: {
        state: () => ({ selection }),
        copyFlavours: () => ({ text: 'Selected text', html: '<b>Selected text</b>' }),
        deleteSelection: vi.fn(),
        revealPosition: vi.fn(),
        pasteRich,
        armForcePlainPaste,
      },
    } as unknown as DocxEditorInstance;
    dispose = bindWordClipboard(host, editor, report);
  });
  afterEach(() => {
    dispose();
    host.remove();
  });

  /** Entrega un portapapeles sintético al elemento indicado sin acceder al portapapeles del equipo. */
  function clipboard(type: string, text: string, html = '', target: Element = surface) {
    const data = {
      getData: (format: string) => (format === 'text/plain' ? text : html),
      setData: vi.fn(),
      files: [],
    };
    const event = new Event(type, { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: data });
    target.dispatchEvent(event);
    return { event, data };
  }

  it('pastes rich content through the Word model at the active scope', () => {
    const { event } = clipboard('paste', 'Visible text', '<p><b>Visible text</b></p>');
    expect(event.defaultPrevented).toBe(true);
    expect(pasteRich).toHaveBeenCalledWith('Visible text', '<p><b>Visible text</b></p>');
  });

  it('retains a text box as the paste destination', () => {
    scope = { kind: 'frame', id: 'story', drawingNodeId: 'box', hostParagraphId: 'anchor' };
    clipboard('paste', 'Inside the box');
    expect(pasteRich).toHaveBeenCalledWith('Inside the box', null);
    expect(exec).not.toHaveBeenCalled();
  });

  it('restores the precise context-menu range and releases its pin before the native action', () => {
    const frame: EditorScope = {
      kind: 'frame',
      id: 'story',
      drawingNodeId: 'box',
      hostParagraphId: 'anchor',
    };
    scope = frame;
    surface.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true }));
    scope = { kind: 'body' };
    selection = {
      anchor: { paragraphId: 'other', offset: 0 },
      head: { paragraphId: 'other', offset: 0 },
    };
    command('paste');
    expect(focus).toHaveBeenCalledWith(frame);
    expect(exec).not.toHaveBeenCalled();
    expect(releaseSelection).toHaveBeenCalledTimes(1);
    expect(native).toHaveBeenCalledWith('paste');
  });

  it('copies both plain text and rich HTML from the model selection', () => {
    const { event, data } = clipboard('copy', '');
    expect(event.defaultPrevented).toBe(true);
    expect(data.setData).toHaveBeenCalledWith('text/plain', 'Selected text');
    expect(data.setData).toHaveBeenCalledWith('text/html', '<b>Selected text</b>');
  });

  it('does not intercept paste in toolbar inputs', () => {
    const input = document.createElement('input');
    host.append(input);
    const { event } = clipboard('paste', 'Arial', '', input);
    expect(event.defaultPrevented).toBe(false);
    expect(exec).not.toHaveBeenCalled();
  });

  it('supports paste without formatting with Ctrl+Shift+V', () => {
    surface.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'v', ctrlKey: true, shiftKey: true, bubbles: true }),
    );
    clipboard('paste', 'Plain text', '<b>Plain text</b>');
    expect(armForcePlainPaste).toHaveBeenCalledTimes(1);
    expect(pasteRich).toHaveBeenCalledWith('Plain text', null);
  });
});
