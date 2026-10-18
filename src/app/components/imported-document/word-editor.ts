import { createApp, h, Fragment } from 'vue';
import { DocxEditor } from '@docx-editor.dev/vue';
import type { Editor } from '@docx-editor.dev/core';
import type { WordEditorMount, WordEditorOptions } from './word-editor.factory';
import { prepareWordForEditing } from './word-legacy-drawings';
import { bindWordClipboard } from './word-clipboard';
import type { WordClipboardCommand } from './word-clipboard';

/**
 * Monta el editor Apache 2.0 con controles de texto, imágenes, tablas y configuración de página.
 * @param host Contenedor exclusivo de esta sesión; sus estilos se limitan a .docx-editor.
 * @param bytes DOCX original completo, procesado localmente sin telemetría ni paquetes Pro.
 * @param options Título, idioma y callbacks de archivo de Helper Notes.
 * @returns Handle de inicialización y limpieza que no persiste por cuenta propia.
 */
export async function mountWordEditor(
  host: HTMLElement,
  bytes: Uint8Array,
  options: WordEditorOptions,
): Promise<WordEditorMount> {
  const language = options.language.split('-')[0];
  const translations = language === 'en' ? undefined : await importLocale(language);
  const document = prepareWordForEditing(bytes);
  const clipboardCommand = (command: WordClipboardCommand): void => {
    host.dispatchEvent(new CustomEvent('word-clipboard-command', { detail: command }));
  };
  // Capturar antes del motor evita activar su conversor PDF opcional o una descarga del navegador.
  const shortcut = (event: KeyboardEvent): void => {
    if (!(event.ctrlKey || event.metaKey) || !['s', 'p'].includes(event.key.toLowerCase())) return;
    event.preventDefault();
    event.stopPropagation();
    if (event.key.toLowerCase() === 's') options.save();
    else options.print();
  };
  host.addEventListener('keydown', shortcut, true);
  let resolveReady!: (editor: Editor) => void;
  let rejectReady!: (error: unknown) => void;
  let disposeClipboard: (() => void) | undefined;
  const ready = new Promise<Editor>((resolve, reject) => {
    resolveReady = resolve;
    rejectReady = reject;
  });
  const app = createApp({
    render: () =>
      h(DocxEditor, {
        document,
        mode: 'edit',
        colorMode: 'light',
        popups: {
          contextMenu: () =>
            h(
              DocxEditor.ContextMenu,
              {},
              {
                default: () => [
                  // Conservar las acciones de tablas y formato, reemplazando sólo las rutas de portapapeles.
                  h(DocxEditor.ContextMenu.Cut, { hidden: true }),
                  h(DocxEditor.ContextMenu.Copy, { hidden: true }),
                  h(DocxEditor.ContextMenu.Paste, { hidden: true }),
                  h(DocxEditor.ContextMenu.PasteWithoutFormatting, { hidden: true }),
                  h(DocxEditor.ContextMenu.Item, {
                    label: translations?.contextMenu?.cut ?? 'Cut',
                    shortcut: 'Ctrl+X',
                    onSelect: () => clipboardCommand('cut'),
                  }),
                  h(DocxEditor.ContextMenu.Item, {
                    label: translations?.contextMenu?.copy ?? 'Copy',
                    shortcut: 'Ctrl+C',
                    onSelect: () => clipboardCommand('copy'),
                  }),
                  h(DocxEditor.ContextMenu.Item, {
                    label: translations?.contextMenu?.paste ?? 'Paste',
                    shortcut: 'Ctrl+V',
                    onSelect: () => clipboardCommand('paste'),
                  }),
                  h(DocxEditor.ContextMenu.Item, {
                    label:
                      translations?.contextMenu?.pasteWithoutFormatting ??
                      'Paste without formatting',
                    shortcut: 'Ctrl+Shift+V',
                    onSelect: () => clipboardCommand('pastePlain'),
                  }),
                ],
              },
            ),
        },
        locale: options.language,
        i18n: translations,
        title: options.title,
        menu: {
          onSave: options.save,
          onOpenFile: options.open,
          reportIssue: false,
          // La impresión de Helper Notes usa su botón nativo; estos conversores requieren módulos Pro.
          children: h(DocxEditor.Menu.File, {
            children: h(Fragment, null, [
              h(DocxEditor.Menu.Item, { slotId: 'file.print', hidden: true }),
              h(DocxEditor.Menu.Item, { slotId: 'file.exportPdf', hidden: true }),
              h(DocxEditor.Menu.Item, { slotId: 'file.exportMarkdown', hidden: true }),
            ]),
          }),
        },
        onReady: (editor: Editor) => {
          disposeClipboard = bindWordClipboard(host, editor, (key) =>
            options.clipboardError?.(key),
          );
          resolveReady(editor);
        },
        onError: rejectReady,
      }),
  });
  app.config.errorHandler = (error) => rejectReady(error);
  try {
    app.mount(host);
  } catch (error) {
    rejectReady(error);
  }
  return {
    ready,
    destroy: () => {
      host.removeEventListener('keydown', shortcut, true);
      disposeClipboard?.();
      app.unmount();
    },
  };
}

/** Recibe el idioma de Helper Notes y retorna únicamente su catálogo libre, o inglés por defecto. */
async function importLocale(language: string) {
  switch (language) {
    case 'es':
      return (await import('@docx-editor.dev/i18n/es')).default;
    case 'de':
      return (await import('@docx-editor.dev/i18n/de')).default;
    case 'fr':
      return (await import('@docx-editor.dev/i18n/fr')).default;
    case 'ja':
      return (await import('@docx-editor.dev/i18n/ja')).default;
    case 'pt':
      return (await import('@docx-editor.dev/i18n/pt-BR')).default;
    case 'zh':
      return (await import('@docx-editor.dev/i18n/zh-CN')).default;
    default:
      return undefined;
  }
}
