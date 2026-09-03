import { Component, inject, OnInit, signal } from '@angular/core';
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

@Component({
  selector: 'app-root',
  imports: [Sidebar, NoteList, NoteEditor, Settings, AssistantSettings, DataTransfer, ResizeHandle, DownloadToast, AiChat],
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
  private readonly themeService = inject(ThemeService);
  private readonly noteService = inject(NoteService);
  private readonly i18nService = inject(I18nService);

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
  }

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
    }
  }

  /** Expande o contrae la barra lateral de carpetas. */
  toggleSidebar(): void {
    this.sidebarCollapsed.update(v => !v);
  }

  /** Expande o contrae la lista central de notas. */
  toggleNoteList(): void {
    this.noteListCollapsed.update(v => !v);
  }

  /** Ajusta el ancho de la barra lateral respetando sus límites visuales. */
  onSidebarResize(delta: number): void {
    if (this.sidebarCollapsed()) return;
    const newWidth = Math.min(this.SIDEBAR_MAX, Math.max(this.SIDEBAR_MIN, this.sidebarWidth() + delta));
    this.sidebarWidth.set(newWidth);
  }

  /** Ajusta el ancho de la lista de notas respetando sus límites visuales. */
  onNoteListResize(delta: number): void {
    const newWidth = Math.min(this.NOTELIST_MAX, Math.max(this.NOTELIST_MIN, this.noteListWidth() + delta));
    this.noteListWidth.set(newWidth);
  }

  /** Ajusta el panel de IA; el signo se invierte porque su tirador está a la izquierda. */
  onAiChatResize(delta: number): void {
    // Negative delta = dragging left = panel gets wider
    const newWidth = Math.min(this.AICHAT_MAX, Math.max(this.AICHAT_MIN, this.aiChatWidth() - delta));
    this.aiChatWidth.set(newWidth);
  }

  /** Muestra u oculta el panel de conversación con IA. */
  toggleAiChat(): void {
    this.aiChatOpen.update(v => !v);
  }
}

