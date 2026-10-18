import {
  Component,
  ElementRef,
  computed,
  effect,
  inject,
  input,
  output,
  signal,
  untracked,
  viewChild,
} from '@angular/core';
import { TranslateModule, TranslateService } from '@ngx-translate/core';
import { Note } from '../../models/note.model';
import { NoteService } from '../../services/note.service';
import { DownloadService } from '../../services/download.service';
import { WordSession, WordSessionService } from '../../services/word-session.service';
import { WORD_EDITOR_FACTORY, WordEditorMount } from './word-editor.factory';
import { restoreWordTextBoxes } from './word-text-boxes';
import { prepareWordForPrint } from './word-legacy-drawings';

/** Edita el DOCX real con su motor local y el TXT como texto plano, conservando cada formato. */
@Component({
  selector: 'app-imported-document',
  imports: [TranslateModule],
  templateUrl: './imported-document.html',
  styleUrl: './imported-document.scss',
})
export class ImportedDocument {
  readonly note = input.required<Note>();
  readonly title = input('');
  readonly text = input('');
  readonly titleChange = output<string>();
  readonly textChange = output<string>();
  readonly printRequested = output<void>();
  protected loading = signal(false);
  protected failed = signal(false);
  protected clipboardError = signal<string | null>(null);
  protected readonly sessions = inject(WordSessionService);
  private readonly host = viewChild<ElementRef<HTMLDivElement>>('documentHost');
  private readonly wordId = computed(() =>
    this.note().file_type === 'docx' ? this.note().id : null,
  );
  private readonly notes = inject(NoteService);
  private readonly downloads = inject(DownloadService);
  private readonly translate = inject(TranslateService);
  private readonly mountEditor = inject(WORD_EDITOR_FACTORY);
  private session: WordSession | null = null;

  constructor() {
    effect((onCleanup) => {
      const noteId = this.wordId();
      const host = this.host()?.nativeElement;
      if (noteId === null || !host) return;
      let current = true;
      let mount: WordEditorMount | undefined;
      let session: WordSession | undefined;
      onCleanup(() => {
        current = false;
        this.session = null;
        // Cada montaje tiene un nodo propio: cerrar el anterior nunca elimina el nuevo editor.
        if (session)
          void session.close().catch((error) => console.error('Word pendiente de guardar:', error));
        else mount?.destroy();
      });
      untracked(() => {
        this.loading.set(true);
        this.failed.set(false);
        this.clipboardError.set(null);
        const container = document.createElement('div');
        container.className = 'word-editor-mount';
        host.replaceChildren(container);
        void (async () => {
          try {
            const fullNote = await window.electronAPI.notesGetById(noteId);
            if (!current) return;
            if (!fullNote?.file_data) throw new Error('El documento Word no contiene datos.');
            const bytes = Uint8Array.from(atob(fullNote.file_data), (character) =>
              character.charCodeAt(0),
            );
            mount = await this.mountEditor(container, bytes, {
              title: fullNote.title,
              language: this.translate.getCurrentLang() || 'es',
              save: () => void this.saveDocument(),
              print: () => this.printRequested.emit(),
              clipboardError: (key) => this.clipboardError.set(key),
              open: (file) => void this.openDocument(file),
            });
            if (!current) {
              mount.destroy();
              return;
            }
            const editor = await mount.ready;
            if (!current) {
              mount.destroy();
              return;
            }
            const mounted = mount;
            session = this.sessions.register(noteId, editor, () => mounted.destroy());
            this.session = session;
          } catch (error) {
            if (current) {
              mount?.destroy();
              this.failed.set(true);
              console.error('Error al abrir el editor Word:', error);
            }
          } finally {
            if (current) this.loading.set(false);
          }
        })();
      });
    });
  }

  /** Guarda las sesiones pendientes; retorna al confirmar o conserva el aviso de error para reintentar. */
  protected async saveDocument(): Promise<void> {
    try {
      await this.sessions.flushAll();
    } catch (error) {
      console.error('No se pudo guardar Word:', error);
    }
  }

  /** Recibe un archivo del menú Word y lo importa como documento independiente; no reemplaza la nota actual. */
  private async openDocument(file: File): Promise<void> {
    try {
      const path = window.electronAPI.fileGetPath(file);
      if (!path) throw new Error('No se pudo obtener la ruta del documento.');
      await this.notes.importPaths([path]);
    } catch (error) {
      console.error('No se pudo importar el documento:', error);
    }
  }

  /**
   * Exporta el DOCX editado o el texto actual del TXT mediante el diálogo nativo.
   * @returns Promesa resuelta al guardar o cancelar; los fallos se muestran en la notificación.
   */
  protected async downloadDocument(): Promise<void> {
    const note = this.note();
    const extension = note.file_type === 'docx' ? 'docx' : 'txt';
    const fileName = `${this.title() || note.title || 'documento'}.${extension}`;
    this.downloads.showSaving(fileName);
    try {
      const filters = [{ name: extension.toUpperCase(), extensions: [extension] }];
      let result: { success: boolean; filePath?: string };
      if (extension === 'txt') {
        result = await window.electronAPI.fileSaveAs({
          defaultName: fileName,
          filters,
          content: this.text(),
          encoding: 'utf8',
        });
      } else {
        const session = this.session;
        if (!session) throw new Error('El editor Word todavía no está listo.');
        result = await window.electronAPI.fileSaveAsBuffer({
          defaultName: fileName,
          filters,
          buffer: Array.from(new Uint8Array(await session.export())),
        });
      }
      if (result.success && result.filePath) this.downloads.showDone(fileName, result.filePath);
      else this.downloads.dismiss(fileName);
    } catch {
      this.downloads.showError(fileName);
    }
  }

  /**
   * Construye una instantánea de impresión desde los últimos bytes editados, sin controles del editor.
   * @returns HTML paginado con estilos y cuadros de texto o null si el motor no está listo.
   */
  async getPrintableContent(): Promise<string | null> {
    const session = this.session;
    if (!session || this.loading() || this.failed()) return null;
    const bytes = await session.export();
    const { renderAsync } = await import('docx-preview');
    const snapshot = document.createElement('div');
    await renderAsync(prepareWordForPrint(new Uint8Array(bytes)), snapshot, snapshot, {
      className: 'word-document',
      ignoreLastRenderedPageBreak: false,
      useBase64URL: true,
      renderHeaders: true,
      renderFooters: true,
      renderAltChunks: false,
      hideWrapperOnPrint: true,
    });
    restoreWordTextBoxes(snapshot);
    snapshot
      .querySelectorAll('script, iframe, object, embed')
      .forEach((element) => element.remove());
    snapshot.querySelectorAll('*').forEach((element) => {
      for (const attribute of Array.from(element.attributes))
        if (attribute.name.startsWith('on')) element.removeAttribute(attribute.name);
    });
    return snapshot.innerHTML;
  }
}
