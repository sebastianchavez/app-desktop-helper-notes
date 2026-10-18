import { Component, DestroyRef, inject, OnInit, signal } from '@angular/core';
import { TranslateModule } from '@ngx-translate/core';
import { Sidebar } from './components/sidebar/sidebar';
import { NoteList } from './components/note-list/note-list';
import { NoteEditor } from './components/note-editor/note-editor';
import { Settings } from './components/settings/settings';
import { ResizeHandle } from './components/resize-handle/resize-handle';
import { DownloadToast } from './components/download-toast/download-toast';
import { AiChat } from './components/ai-chat/ai-chat';
import { AssistantSettings } from './components/assistant-settings/assistant-settings';
import { DataTransfer } from './components/data-transfer/data-transfer';
import { NavigationService } from './services/navigation.service';
import { ThemeService } from './services/theme.service';
import { NoteService } from './services/note.service';
import { I18nService } from './services/i18n.service';
import { WordSessionService } from './services/word-session.service';

@Component({
  selector: 'app-root',
  imports: [
    TranslateModule,
    Sidebar,
    NoteList,
    NoteEditor,
    Settings,
    AssistantSettings,
    DataTransfer,
    ResizeHandle,
    DownloadToast,
    AiChat,
  ],
  templateUrl: './app.html',
  styleUrl: './app.scss',
})
/**
 * Componente raíz de la interfaz.
 *
 * Compone las vistas principales, conserva el tamaño y estado de los paneles y
 * detecta si Electron abrió una ventana dedicada a una sola nota mediante el
 * parámetro `noteId` de la URL.
 */
export class App implements OnInit {
  protected readonly navService = inject(NavigationService);
  protected readonly wordSessions = inject(WordSessionService);

  /** Reintenta guardar los Word pendientes y conserva el aviso si la escritura vuelve a fallar. */
  protected async retryWordSave(): Promise<void> {
    try {
      await this.wordSessions.flushAll();
    } catch (error) {
      console.error('Word pendiente de guardar:', error);
    }
  }
  private readonly themeService = inject(ThemeService);
  private readonly noteService = inject(NoteService);
  private readonly i18nService = inject(I18nService);
  private readonly destroyRef = inject(DestroyRef);
  protected fileDragActive = signal(false);
  protected fileImportBusy = signal(false);
  protected fileImportMessage = signal<string | null>(null);
  private importQueue: Promise<void> = Promise.resolve();

  protected sidebarWidth = signal(256);
  protected noteListWidth = signal(320);
  protected sidebarCollapsed = signal(false);
  protected noteListCollapsed = signal(false);
  protected aiChatOpen = signal(false);
  protected aiChatWidth = signal(380);

  /** When set, app shows only the note editor (standalone window mode) */
  protected standaloneNoteId = signal<number | null>(null);

  private readonly SIDEBAR_MIN = 200;
  private readonly SIDEBAR_MAX = 400;
  private readonly SIDEBAR_COLLAPSED_WIDTH = 64;
  private readonly NOTELIST_MIN = 240;
  private readonly NOTELIST_MAX = 500;
  private readonly AICHAT_MIN = 300;
  private readonly AICHAT_MAX = 600;

  constructor() {
    // Capture noteId query param IMMEDIATELY before Angular router can strip it
    const params = new URLSearchParams(window.location.search);
    const noteIdParam = params.get('noteId');
    if (noteIdParam) {
      const noteId = parseInt(noteIdParam, 10);
      if (!isNaN(noteId)) {
        this.standaloneNoteId.set(noteId);
      }
    }
    // Capture avoids native navigation and nested drop handlers consuming external files.
    const over = (event: DragEvent) => this.onExternalDragOver(event);
    const drop = (event: DragEvent) => this.onExternalDrop(event);
    const leave = (event: DragEvent) => {
      if (!event.relatedTarget) this.fileDragActive.set(false);
    };
    const end = () => this.fileDragActive.set(false);
    window.addEventListener('dragover', over, true);
    window.addEventListener('drop', drop, true);
    window.addEventListener('dragleave', leave, true);
    window.addEventListener('dragend', end, true);
    this.destroyRef.onDestroy(() => {
      window.removeEventListener('dragover', over, true);
      window.removeEventListener('drop', drop, true);
      window.removeEventListener('dragleave', leave, true);
      window.removeEventListener('dragend', end, true);
    });
  }

  /**
   * Inicializa preferencias y carga la nota independiente o las aperturas del sistema.
   * @returns Promesa resuelta cuando el contexto inicial está listo; no recibe parámetros.
   */
  async ngOnInit(): Promise<void> {
    await this.themeService.init();
    await this.i18nService.init();

    // If standalone note window, load the note directly
    const noteId = this.standaloneNoteId();
    if (noteId) {
      const note = await window.electronAPI.notesGetById(noteId);
      if (note) {
        // Put the note into the notes array so selectedNote() computed works
        this.noteService.notes.set([note]);
        this.noteService.selectedNoteId.set(noteId);
      }
    } else {
      const unsubscribe = window.electronAPI.onFileOpenRequested(() => {
        void this.consumePendingFiles();
      });
      this.destroyRef.onDestroy(unsubscribe);
      await this.consumePendingFiles();
    }
  }

  /**
   * Distingue archivos externos del arrastre interno de notas y carpetas.
   * @param event Evento de arrastre de Electron/Chromium.
   * @returns Si la transferencia contiene archivos del sistema.
   */
  private isExternalFileDrag(event: DragEvent): boolean {
    const types = event.dataTransfer?.types ?? [];
    return (
      types.includes('Files') &&
      !types.includes('application/note-id') &&
      !types.includes('application/folder-id')
    );
  }

  /**
   * Activa el destino de importación y evita la navegación nativa al archivo.
   * @param event Evento capturado antes de los manejadores de carpetas y del editor.
   * @returns Nada; actualiza el indicador visual sólo para archivos externos.
   */
  onExternalDragOver(event: DragEvent): void {
    if (!this.isExternalFileDrag(event)) return;
    event.preventDefault();
    event.stopPropagation();
    event.dataTransfer!.dropEffect = 'copy';
    this.fileDragActive.set(true);
  }

  /**
   * Obtiene las rutas físicas mediante el preload y encola el lote para abrirlo.
   * @param event Evento que contiene los archivos soltados.
   * @returns Nada; informa errores de rutas y deja los arrastres internos intactos.
   */
  onExternalDrop(event: DragEvent): void {
    if (!this.isExternalFileDrag(event)) return;
    event.preventDefault();
    event.stopPropagation();
    this.fileDragActive.set(false);
    try {
      const paths = Array.from(event.dataTransfer!.files)
        .map((file) => window.electronAPI.fileGetPath(file))
        .filter(Boolean);
      if (!paths.length) {
        this.fileImportMessage.set('APP.FILE_IMPORT_ERROR');
        return;
      }
      void this.enqueueImport(paths);
    } catch {
      this.fileImportMessage.set('APP.FILE_IMPORT_ERROR');
    }
  }

  /**
   * Recoge archivos enviados por Windows, incluso antes de que Angular estuviera listo.
   * @returns Promesa resuelta al finalizar la importación de las rutas pendientes.
   */
  private async consumePendingFiles(): Promise<void> {
    const paths = await window.electronAPI.fileTakePendingOpen();
    if (Array.isArray(paths) && paths.length) await this.enqueueImport(paths);
  }

  /**
   * Serializa los lotes para que la selección final corresponda al último archivo abierto.
   * @param paths Rutas físicas del lote a importar en la carpeta activa.
   * @returns Promesa del lote, resuelta también ante fallos para no bloquear lotes posteriores.
   */
  private enqueueImport(paths: string[]): Promise<void> {
    this.importQueue = this.importQueue.then(async () => {
      this.fileImportBusy.set(true);
      this.fileImportMessage.set(null);
      try {
        const result = await this.noteService.importPaths(paths);
        if (result.count) {
          this.standaloneNoteId.set(null);
          this.navService.goToNotes();
          this.noteListCollapsed.set(false);
        }
        if (result.errors.length || !result.count)
          this.fileImportMessage.set('APP.FILE_IMPORT_ERROR');
      } catch {
        this.fileImportMessage.set('APP.FILE_IMPORT_ERROR');
      } finally {
        this.fileImportBusy.set(false);
      }
    });
    return this.importQueue;
  }

  /** Expande o contrae la barra lateral de carpetas. */
  toggleSidebar(): void {
    this.sidebarCollapsed.update((v) => !v);
  }

  /** Expande o contrae la lista central de notas. */
  toggleNoteList(): void {
    this.noteListCollapsed.update((v) => !v);
  }

  /** Ajusta el ancho de la barra lateral respetando sus límites visuales. */
  onSidebarResize(delta: number): void {
    if (this.sidebarCollapsed()) return;
    const newWidth = Math.min(
      this.SIDEBAR_MAX,
      Math.max(this.SIDEBAR_MIN, this.sidebarWidth() + delta),
    );
    this.sidebarWidth.set(newWidth);
  }

  /** Ajusta el ancho de la lista de notas respetando sus límites visuales. */
  onNoteListResize(delta: number): void {
    const newWidth = Math.min(
      this.NOTELIST_MAX,
      Math.max(this.NOTELIST_MIN, this.noteListWidth() + delta),
    );
    this.noteListWidth.set(newWidth);
  }

  /** Ajusta el panel de IA; el signo se invierte porque su tirador está a la izquierda. */
  onAiChatResize(delta: number): void {
    // Negative delta = dragging left = panel gets wider
    const newWidth = Math.min(
      this.AICHAT_MAX,
      Math.max(this.AICHAT_MIN, this.aiChatWidth() - delta),
    );
    this.aiChatWidth.set(newWidth);
  }

  /** Muestra u oculta el panel de conversación con IA. */
  toggleAiChat(): void {
    this.aiChatOpen.update((v) => !v);
  }
}
