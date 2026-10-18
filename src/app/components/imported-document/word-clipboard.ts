import type {
  DocxEditorInstance,
  Editor,
  EditorScope,
  SupportedImageMime,
} from '@docx-editor.dev/core';

export type WordClipboardCommand = 'copy' | 'cut' | 'paste' | 'pastePlain';
const IMAGE_TYPES = new Set<SupportedImageMime>([
  'image/png',
  'image/jpeg',
  'image/gif',
  'image/bmp',
  'image/webp',
]);

/**
 * Conecta el portapapeles nativo con la selección del modelo Word, sin editar el DOM pintado.
 * @param host Contenedor exclusivo del editor.
 * @param editor Motor local ya inicializado.
 * @param report Publica errores visibles; recibe null al iniciar otra operación.
 * @returns Limpieza de eventos, selección retenida y suscripción IPC.
 */
export function bindWordClipboard(
  host: HTMLElement,
  editor: Editor,
  report: (key: string | null) => void,
): () => void {
  const instance = editor as DocxEditorInstance;
  let disposed = false;
  let plainNextPaste = false;
  let releaseContext: (() => void) | undefined;
  let contextScope: EditorScope | undefined;
  let contextSelection:
    ReturnType<NonNullable<DocxEditorInstance['surface']>['state']>['selection'] | undefined;

  const isSurface = (target: EventTarget | null): boolean =>
    target instanceof Element && !!target.closest('.docx-paginated-surface');
  const reveal = (): void => {
    const surface = instance.surface;
    if (surface) surface.revealPosition(surface.state().selection.head, { block: 'nearest' });
  };
  const copy = (event: ClipboardEvent): void => {
    if (!isSurface(event.target) || !event.clipboardData || !instance.surface) return;
    const selectedImage = editor.getSelectedImage();
    if (selectedImage) {
      event.preventDefault();
      event.stopImmediatePropagation();
      report(null);
      try {
        const drawing = Array.from(
          host.querySelectorAll<HTMLElement>('[data-drawing-node-id]'),
        ).find((element) => element.dataset['drawingNodeId'] === selectedImage.id);
        const image = drawing instanceof HTMLImageElement ? drawing : drawing?.querySelector('img');
        if (!image?.naturalWidth) throw new Error('La imagen todavía no está lista.');
        const canvas = document.createElement('canvas');
        canvas.width = image.naturalWidth;
        canvas.height = image.naturalHeight;
        if (canvas.width * canvas.height > 25000000) throw new Error('Imagen demasiado grande.');
        const context = canvas.getContext('2d');
        if (!context) throw new Error('No se pudo copiar la imagen.');
        context.drawImage(image, 0, 0);
        const dataUrl = canvas.toDataURL('image/png');
        event.clipboardData.setData('text/html', `<img src="${dataUrl}">`);
        const write = window.electronAPI?.wordClipboardWriteImage;
        if (!write) throw new Error('El portapapeles nativo no está disponible.');
        void write(dataUrl.split(',')[1])
          .then(() => {
            if (!disposed && event.type === 'cut') {
              const result = editor.exec({ type: 'deleteImage', drawingNodeId: selectedImage.id });
              if (!result.ok) report('DOCUMENT.CLIPBOARD_COPY_ERROR');
            }
          })
          .catch(() => report('DOCUMENT.CLIPBOARD_COPY_ERROR'));
      } catch {
        report('DOCUMENT.CLIPBOARD_COPY_ERROR');
      }
      return;
    }
    const flavours = instance.surface.copyFlavours();
    if (flavours.reason) {
      event.preventDefault();
      event.stopImmediatePropagation();
      report('DOCUMENT.CLIPBOARD_COPY_ERROR');
      return;
    }
    if (!flavours.text) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    report(null);
    event.clipboardData.setData('text/plain', flavours.text);
    if (flavours.html) event.clipboardData.setData('text/html', flavours.html);
    if (event.type === 'cut') instance.surface.deleteSelection();
  };
  const paste = (event: ClipboardEvent): void => {
    if (!isSurface(event.target) || !event.clipboardData) return;
    const data = event.clipboardData;
    const image = Array.from(data.files).find((file) => file.type.startsWith('image/'));
    const text = data.getData('text/plain');
    const html = plainNextPaste ? '' : data.getData('text/html');
    const plain = plainNextPaste;
    plainNextPaste = false;
    if (!image && !text && !html) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    report(null);
    if (image && !plain) {
      void insertClipboardImage(editor, image, () => disposed)
        .then(reveal)
        .catch(() => report('DOCUMENT.CLIPBOARD_IMAGE_ERROR'));
      return;
    }
    // La superficie resuelve cuadros/celdas; el comando genérico paste no admite historias frame.
    if (plain) instance.surface?.armForcePlainPaste();
    const changed = editor.snapshot().editable && instance.surface?.pasteRich(text, html || null);
    if (!changed) report('DOCUMENT.CLIPBOARD_PASTE_ERROR');
    else requestAnimationFrame(reveal);
  };
  const keydown = (event: KeyboardEvent): void => {
    if (
      isSurface(event.target) &&
      (event.ctrlKey || event.metaKey) &&
      event.key.toLowerCase() === 'v'
    )
      plainNextPaste = event.shiftKey;
  };
  const context = (event: MouseEvent): void => {
    if (!isSurface(event.target)) return;
    // El menú Electron conserva un pin real del motor, incluidos cuadros y celdas.
    contextScope = editor.snapshot().scope;
    contextSelection = instance.surface?.state().selection;
    releaseContext?.();
    const pin = editor.retainSelection();
    releaseContext = pin ? () => editor.releaseSelection(pin) : undefined;
  };
  const runCommand = (command: WordClipboardCommand): void => {
    if (disposed || !contextScope || !contextSelection) return;
    editor.focus(contextScope);
    // El pin conserva el rango propio de la historia; no sustituirlo por el rango del cuerpo.
    releaseContext?.();
    releaseContext = undefined;
    plainNextPaste = command === 'pastePlain';
    void window.electronAPI
      .wordClipboardExecute(command)
      .catch(() => report('DOCUMENT.CLIPBOARD_PASTE_ERROR'));
  };
  const menuCommand = (event: Event): void => {
    const command = (event as CustomEvent<WordClipboardCommand>).detail;
    if (['copy', 'cut', 'paste', 'pastePlain'].includes(command)) runCommand(command);
  };
  const unsubscribe = window.electronAPI?.onWordClipboardCommand?.(runCommand);
  host.dataset['wordEditorSurface'] = 'true';
  host.dataset['wordOwnsContextMenu'] = 'true';
  host.addEventListener('copy', copy, true);
  host.addEventListener('cut', copy, true);
  host.addEventListener('paste', paste, true);
  host.addEventListener('keydown', keydown, true);
  host.addEventListener('contextmenu', context, true);
  host.addEventListener('word-clipboard-command', menuCommand);
  return () => {
    disposed = true;
    unsubscribe?.();
    releaseContext?.();
    delete host.dataset['wordEditorSurface'];
    delete host.dataset['wordOwnsContextMenu'];
    host.removeEventListener('copy', copy, true);
    host.removeEventListener('cut', copy, true);
    host.removeEventListener('paste', paste, true);
    host.removeEventListener('keydown', keydown, true);
    host.removeEventListener('contextmenu', context, true);
    host.removeEventListener('word-clipboard-command', menuCommand);
  };
}

/**
 * Inserta una imagen en la selección que recibió el pegado, conservando su proporción.
 * @param editor Motor propietario del documento.
 * @param file Imagen entregada por el evento nativo de portapapeles.
 * @param isDisposed Indica si se cerró la sesión durante la lectura asíncrona.
 * @returns Promesa resuelta al insertar o rechazada sin pegar en otro lugar.
 */
async function insertClipboardImage(
  editor: Editor,
  file: File,
  isDisposed: () => boolean,
): Promise<void> {
  if (!IMAGE_TYPES.has(file.type as SupportedImageMime) || file.size > 25 * 1024 * 1024)
    throw new Error('Imagen no compatible.');
  const instance = editor as DocxEditorInstance;
  const selection = JSON.stringify({
    scope: editor.snapshot().scope,
    selection: instance.surface?.state().selection,
  });
  const bytes = new Uint8Array(await file.arrayBuffer());
  const bitmap = await createImageBitmap(file);
  const width = Math.min(432, bitmap.width * 0.75);
  const height = (width * bitmap.height) / bitmap.width;
  bitmap.close();
  if (
    isDisposed() ||
    selection !==
      JSON.stringify({
        scope: editor.snapshot().scope,
        selection: instance.surface?.state().selection,
      })
  )
    throw new Error('La selección cambió antes de pegar.');
  const result = await editor.executeImageCommand({
    type: 'insertImage',
    data: bytes,
    mime: file.type as SupportedImageMime,
    widthPoints: width,
    heightPoints: height,
  });
  if (!result.ok || !result.changed) throw new Error('El destino no admite esta imagen.');
}
