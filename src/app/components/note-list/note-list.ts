import { Component, inject, OnInit, signal, effect, input, output } from '@angular/core';
import { TranslateModule, TranslateService } from '@ngx-translate/core';
import { NoteService } from '../../services/note.service';
import { AiChatService } from '../../services/ai-chat.service';

@Component({
  selector: 'app-note-list',
  templateUrl: './note-list.html',
  styleUrl: './note-list.scss',
  imports: [TranslateModule],
})
export class NoteList implements OnInit {
  protected readonly noteService = inject(NoteService);
  private readonly translate = inject(TranslateService);
  private readonly aiChatService = inject(AiChatService);
  protected searchQuery = signal('');

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
    await this.noteService.createNote();
  }

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

  onDragStart(event: DragEvent, noteId: number): void {
    event.dataTransfer?.setData('application/note-id', noteId.toString());
    event.dataTransfer!.effectAllowed = 'move';
  }

  getPreview(content: string): string {
    if (!content) return this.translate.instant('NOTE_LIST.NO_CONTENT');
    // Strip HTML tags to show plain text preview
    const tmp = document.createElement('div');
    tmp.innerHTML = content;
    const text = tmp.textContent || tmp.innerText || '';
    if (!text.trim()) return this.translate.instant('NOTE_LIST.NO_CONTENT');
    return text.length > 100 ? text.substring(0, 100) + '...' : text;
  }

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
