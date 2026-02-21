import { Component, inject, OnInit, signal, effect } from '@angular/core';
import { NoteService } from '../../services/note.service';

@Component({
  selector: 'app-note-list',
  templateUrl: './note-list.html',
  styleUrl: './note-list.scss',
})
export class NoteList implements OnInit {
  protected readonly noteService = inject(NoteService);
  protected searchQuery = signal('');

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
    await this.noteService.selectNote(id);
  }

  async createNote(): Promise<void> {
    await this.noteService.createNote();
  }

  async deleteNote(id: number, event: Event): Promise<void> {
    event.stopPropagation();
    await this.noteService.deleteNote(id);
  }

  getPreview(content: string): string {
    if (!content) return 'Sin contenido...';
    // Strip HTML tags to show plain text preview
    const tmp = document.createElement('div');
    tmp.innerHTML = content;
    const text = tmp.textContent || tmp.innerText || '';
    if (!text.trim()) return 'Sin contenido...';
    return text.length > 100 ? text.substring(0, 100) + '...' : text;
  }

  getRelativeTime(dateStr: string): string {
    const date = new Date(dateStr + 'Z');
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffMin = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMs / 3600000);
    const diffDays = Math.floor(diffMs / 86400000);

    if (diffMin < 1) return 'ahora';
    if (diffMin < 60) return `${diffMin}m`;
    if (diffHours < 24) return `${diffHours}h`;
    if (diffDays < 7) return `${diffDays}d`;
    return date.toLocaleDateString('es', { month: 'short', day: 'numeric' });
  }
}
