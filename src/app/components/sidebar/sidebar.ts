import { Component, inject, OnInit, signal } from '@angular/core';
import { NoteService } from '../../services/note.service';
import { NavigationService } from '../../services/navigation.service';

@Component({
  selector: 'app-sidebar',
  templateUrl: './sidebar.html',
  styleUrl: './sidebar.scss',
})
export class Sidebar implements OnInit {
  protected readonly noteService = inject(NoteService);
  protected readonly navService = inject(NavigationService);
  protected editingFolderId = signal<number | null>(null);
  protected editingFolderName = signal('');
  protected showNewFolderInput = signal(false);
  protected newFolderName = signal('');

  async ngOnInit(): Promise<void> {
    await this.noteService.loadFolders();
  }

  async selectFolder(id: number | null): Promise<void> {
    this.navService.goToNotes();
    await this.noteService.selectFolder(id);
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

  async deleteFolder(id: number, event: Event): Promise<void> {
    event.stopPropagation();
    await this.noteService.deleteFolder(id);
  }
}
