import { Injectable, signal, computed, NgZone, inject } from '@angular/core';
import { Folder } from '../models/folder.model';
import { Note } from '../models/note.model';

@Injectable({ providedIn: 'root' })
export class NoteService {
  private api = window.electronAPI;
  private zone = inject(NgZone);

  // ── State ──
  readonly folders = signal<Folder[]>([]);
  readonly notes = signal<Note[]>([]);
  readonly selectedFolderId = signal<number | null>(null);
  readonly selectedNoteId = signal<number | null>(null);

  readonly selectedNote = computed(() => {
    const id = this.selectedNoteId();
    return this.notes().find((n) => n.id === id) ?? null;
  });

  readonly selectedFolder = computed(() => {
    const id = this.selectedFolderId();
    return this.folders().find((f) => f.id === id) ?? null;
  });

  constructor() {
    // Listen for note changes from other windows
    this.api.onNoteChanged((noteId: number) => {
      this.zone.run(async () => {
        await this.loadNotes();
        // If the changed note is the one currently selected, re-select to refresh editor
        if (this.selectedNoteId() === noteId) {
          // Force re-emit by toggling
          this.selectedNoteId.set(null);
          this.selectedNoteId.set(noteId);
        }
      });
    });
  }

  // ── Folders ──
  async loadFolders(): Promise<void> {
    const folders = await this.api.foldersGetAll();
    this.folders.set(folders);
  }

  async createFolder(name: string): Promise<Folder> {
    const folder = await this.api.foldersCreate(name);
    await this.loadFolders();
    return folder;
  }

  async updateFolder(id: number, name: string): Promise<void> {
    await this.api.foldersUpdate(id, name);
    await this.loadFolders();
  }

  async deleteFolder(id: number): Promise<void> {
    await this.api.foldersDelete(id);
    if (this.selectedFolderId() === id) {
      this.selectedFolderId.set(null);
      this.selectedNoteId.set(null);
    }
    await this.loadFolders();
    await this.loadNotes();
  }

  async selectFolder(id: number | null): Promise<void> {
    this.selectedFolderId.set(id);
    this.selectedNoteId.set(null);
    await this.loadNotes();
  }

  // ── Notes ──
  async loadNotes(): Promise<void> {
    const notes = await this.api.notesGetByFolder(this.selectedFolderId());
    this.notes.set(notes);
  }

  async createNote(): Promise<Note> {
    const note = await this.api.notesCreate(this.selectedFolderId());
    await this.loadNotes();
    this.selectedNoteId.set(note.id);
    return note;
  }

  async selectNote(id: number): Promise<void> {
    this.selectedNoteId.set(id);
  }

  async updateNote(id: number, data: { title?: string; content?: string }): Promise<void> {
    await this.api.notesUpdate(id, data);
    await this.loadNotes();
  }

  async deleteNote(id: number): Promise<void> {
    await this.api.notesDelete(id);
    if (this.selectedNoteId() === id) {
      this.selectedNoteId.set(null);
    }
    await this.loadNotes();
  }

  async moveNoteToFolder(noteId: number, folderId: number | null): Promise<void> {
    await this.api.notesUpdate(noteId, { folder_id: folderId });
    // Reload notes so the moved note disappears from the current view
    await this.loadNotes();
    // Deselect if the moved note was selected (it's now in a different folder)
    if (this.selectedNoteId() === noteId) {
      this.selectedNoteId.set(null);
    }
  }

  async importNotes(): Promise<number> {
    const files = await this.api.fileImport();
    if (!files || files.length === 0) return 0;

    let lastNoteId: number | null = null;
    for (const file of files) {
      const note = await this.api.notesCreate(this.selectedFolderId(), file.title);
      const updateData: { content: string; file_data?: string; file_type?: string } = { content: file.content };
      if (file.file_data) updateData.file_data = file.file_data;
      if (file.file_type) updateData.file_type = file.file_type;
      await this.api.notesUpdate(note.id, updateData);
      lastNoteId = note.id;
    }

    await this.loadNotes();
    if (lastNoteId !== null) {
      this.selectedNoteId.set(lastNoteId);
    }
    return files.length;
  }
}
