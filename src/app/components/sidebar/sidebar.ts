import { Component, inject, input, OnInit, output, signal } from '@angular/core';
import { TranslateModule, TranslateService } from '@ngx-translate/core';
import { NoteService } from '../../services/note.service';
import { NavigationService } from '../../services/navigation.service';
import { AiChatService } from '../../services/ai-chat.service';

@Component({
  selector: 'app-sidebar',
  imports: [TranslateModule],
  templateUrl: './sidebar.html',
  styleUrl: './sidebar.scss',
})
export class Sidebar implements OnInit {
  protected readonly noteService = inject(NoteService);
  protected readonly navService = inject(NavigationService);
  private readonly translate = inject(TranslateService);
  private readonly aiChatService = inject(AiChatService);
  protected editingFolderId = signal<number | null>(null);
  protected editingFolderName = signal('');
  protected showNewFolderInput = signal(false);
  protected newFolderName = signal('');
  protected dragOverFolderId = signal<number | null | 'all'>(null);

  /** Whether the sidebar is in collapsed state */
  collapsed = input(false);
  /** Emits when user clicks the header to toggle collapse */
  collapseToggle = output<void>();

  async ngOnInit(): Promise<void> {
    await this.noteService.loadFolders();
  }

  async selectFolder(id: number | null): Promise<void> {
    if (this.aiChatService.isStreaming()) return;
    this.navService.goToNotes();
    await this.noteService.selectFolder(id);
  }

  goToSettings(): void {
    if (this.aiChatService.isStreaming()) return;
    this.navService.goToSettings();
  }

  showNewFolder(): void {
    this.showNewFolderInput.set(true);
    this.newFolderName.set('');
  }

  async confirmNewFolder(): Promise<void> {
    const name = this.newFolderName().trim();
    if (name) {
      await this.noteService.createFolder(name);
    }
    this.showNewFolderInput.set(false);
    this.newFolderName.set('');
  }

  cancelNewFolder(): void {
    this.showNewFolderInput.set(false);
    this.newFolderName.set('');
  }

  startEditing(id: number, currentName: string): void {
    this.editingFolderId.set(id);
    this.editingFolderName.set(currentName);
  }

  async confirmEdit(id: number): Promise<void> {
    const name = this.editingFolderName().trim();
    if (name) {
      await this.noteService.updateFolder(id, name);
    }
    this.editingFolderId.set(null);
  }

  cancelEdit(): void {
    this.editingFolderId.set(null);
  }

  startEditingFromButton(id: number, name: string, event: Event): void {
    event.stopPropagation();
    this.startEditing(id, name);
  }

  async deleteFolder(id: number, event: Event): Promise<void> {
    event.stopPropagation();
    const confirmed = await window.electronAPI.dialogConfirm({
      title: this.translate.instant('FOLDER.DELETE_TITLE'),
      message: this.translate.instant('FOLDER.DELETE_MESSAGE'),
    });
    if (!confirmed) return;
    await this.noteService.deleteFolder(id);
  }

  // ── Drag & drop ──
  onDragOver(event: DragEvent, folderId: number | null): void {
    if (event.dataTransfer?.types.includes('application/note-id')) {
      event.preventDefault();
      event.dataTransfer.dropEffect = 'move';
      this.dragOverFolderId.set(folderId ?? 'all');
    }
  }

  onDragLeave(): void {
    this.dragOverFolderId.set(null);
  }

  async onDrop(event: DragEvent, folderId: number | null): Promise<void> {
    event.preventDefault();
    this.dragOverFolderId.set(null);
    const noteIdStr = event.dataTransfer?.getData('application/note-id');
    if (!noteIdStr) return;
    const noteId = parseInt(noteIdStr, 10);
    if (isNaN(noteId)) return;
    await this.noteService.moveNoteToFolder(noteId, folderId);
  }
}
