import { Component, inject, OnInit, signal, effect, input, output, HostListener, ElementRef } from '@angular/core';
import { TranslateModule, TranslateService } from '@ngx-translate/core';
import { NoteService } from '../../services/note.service';
import { AiChatService } from '../../services/ai-chat.service';

@Component({
  selector: 'app-note-list',
  templateUrl: './note-list.html',
  styleUrl: './note-list.scss',
  imports: [TranslateModule],
})
/**
 * Lista filtrable de notas de la carpeta activa.
 * Gestiona creación por tipo, importación, borrado, apertura en otra ventana
 * y el origen del drag-and-drop hacia el árbol de carpetas.
 */
export class NoteList implements OnInit {
  protected readonly noteService = inject(NoteService);
  private readonly translate = inject(TranslateService);
  private readonly aiChatService = inject(AiChatService);
  private readonly elRef = inject(ElementRef);
  protected searchQuery = signal('');
  protected showCreateMenu = signal(false);

  readonly collapsed = input(false);
  readonly collapseToggle = output();

  constructor() {
    // Reload notes when folder changes
    effect(() => {
      this.noteService.selectedFolderId();
      this.noteService.loadNotes();
    });
  }

  async ngOnInit(): Promise<void> {
    await this.noteService.loadNotes();
  }

  /** Filtra por título o contenido manteniendo el orden entregado por la base. */
  filteredNotes() {
    const query = this.searchQuery().toLowerCase();
    if (!query) return this.noteService.notes();
    return this.noteService.notes().filter(
      (n) => n.title.toLowerCase().includes(query) || n.content.toLowerCase().includes(query)
    );
  }

  async selectNote(id: number): Promise<void> {
    if (this.aiChatService.isStreaming()) return;
    await this.noteService.selectNote(id);
  }

  async createNote(): Promise<void> {
    this.showCreateMenu.set(false);
    await this.noteService.createNote();
  }

  async createExcelNote(): Promise<void> {
    this.showCreateMenu.set(false);
    await this.noteService.createExcelNote();
  }

  async createPresentationNote(): Promise<void> {
    this.showCreateMenu.set(false);
    await this.noteService.createPresentationNote();
  }

  toggleCreateMenu(): void {
    this.showCreateMenu.update(v => !v);
  }

  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    if (this.showCreateMenu() && !this.elRef.nativeElement.querySelector('.create-menu-container')?.contains(event.target as Node)) {
      this.showCreateMenu.set(false);
    }
  }

  /** Solicita confirmación nativa antes de eliminar una nota. */
  async deleteNote(id: number, event: Event): Promise<void> {
    event.stopPropagation();
    const confirmed = await window.electronAPI.dialogConfirm({
      title: this.translate.instant('NOTE_LIST.DELETE_TITLE'),
      message: this.translate.instant('NOTE_LIST.DELETE_MESSAGE'),
    });
    if (!confirmed) return;
    await this.noteService.deleteNote(id);
  }

  async openInNewWindow(id: number): Promise<void> {
    await window.electronAPI.windowOpenNote(id);
  }

  async importNotes(): Promise<void> {
    await this.noteService.importNotes();
  }

  async importJsonNotes(): Promise<void> {
    await this.noteService.importJsonNotes();
  }

  /** Publica el identificador de nota en un MIME interno para moverla entre carpetas. */
  onDragStart(event: DragEvent, noteId: number): void {
    event.dataTransfer?.setData('application/note-id', noteId.toString());
    event.dataTransfer!.effectAllowed = 'move';
  }

  /** Convierte HTML en una vista previa de texto plano limitada a cien caracteres. */
  getPreview(content: string): string {
    if (!content) return this.translate.instant('NOTE_LIST.NO_CONTENT');
    // Strip HTML tags to show plain text preview
    const tmp = document.createElement('div');
    tmp.innerHTML = content;
    const text = tmp.textContent || tmp.innerText || '';
    if (!text.trim()) return this.translate.instant('NOTE_LIST.NO_CONTENT');
    return text.length > 100 ? text.substring(0, 100) + '...' : text;
  }

  /** Presenta fechas recientes de forma compacta y localiza las más antiguas. */
  getRelativeTime(dateStr: string): string {
    const date = new Date(dateStr + 'Z');
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffMin = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMs / 3600000);
    const diffDays = Math.floor(diffMs / 86400000);

    if (diffMin < 1) return this.translate.instant('NOTE_LIST.TIME_NOW');
    if (diffMin < 60) return `${diffMin}m`;
    if (diffHours < 24) return `${diffHours}h`;
    if (diffDays < 7) return `${diffDays}d`;
    return date.toLocaleDateString(this.translate.currentLang || 'es', { month: 'short', day: 'numeric' });
  }
}
