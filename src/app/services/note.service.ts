import { Injectable, signal, computed, NgZone, inject } from '@angular/core';
import { Folder } from '../models/folder.model';
import { Note } from '../models/note.model';
import { createPresentationDocument } from '../models/presentation.model';

@Injectable({ providedIn: 'root' })
/**
 * Fachada reactiva de carpetas y notas para el renderer de Angular.
 *
 * Delega la persistencia al proceso principal mediante `electronAPI`, refresca
 * las señales después de cada escritura y sincroniza cambios originados en
 * otras ventanas de la aplicación.
 */
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
  /** Carga desde SQLite el árbol completo de carpetas. */
  async loadFolders(): Promise<void> {
    const folders = await this.api.foldersGetAll();
    this.folders.set(folders);
  }

  /** Crea una carpeta, opcionalmente anidada, y refresca el árbol. */
  async createFolder(name: string, parentId: number | null = null): Promise<Folder> {
    const folder = await this.api.foldersCreate(name, parentId);
    await this.loadFolders();
    return folder;
  }

  /** Renombra una carpeta y vuelve a cargar la fuente reactiva. */
  async updateFolder(id: number, name: string): Promise<void> {
    await this.api.foldersUpdate(id, name);
    await this.loadFolders();
  }

  /**
   * Elimina una carpeta y limpia la selección si estaba dentro de la rama
   * eliminada. El conjunto `deletedIds` reproduce en memoria el borrado en
   * cascada ejecutado por SQLite.
   */
  async deleteFolder(id: number): Promise<void> {
    const deletedIds = new Set<number>([id]);
    let changed = true;
    while (changed) {
      changed = false;
      for (const folder of this.folders()) {
        if (folder.parent_id !== null && deletedIds.has(folder.parent_id) && !deletedIds.has(folder.id)) {
          deletedIds.add(folder.id);
          changed = true;
        }
      }
    }
    await this.api.foldersDelete(id);
    const selectedId = this.selectedFolderId();
    if (selectedId !== null && deletedIds.has(selectedId)) {
      this.selectedFolderId.set(null);
      this.selectedNoteId.set(null);
    }
    await this.loadFolders();
    await this.loadNotes();
  }

  /** Selecciona una carpeta, deselecciona la nota actual y carga su contenido. */
  async selectFolder(id: number | null): Promise<void> {
    this.selectedFolderId.set(id);
    this.selectedNoteId.set(null);
    await this.loadNotes();
  }

  // ── Notes ──
  /** Carga las notas pertenecientes a la carpeta seleccionada. */
  async loadNotes(): Promise<void> {
    const notes = await this.api.notesGetByFolder(this.selectedFolderId());
    this.notes.set(notes);
  }

  /**
   * Reload notes from DB and force-refresh the currently selected note
   * so that components watching selectedNote re-run their effects.
   */
  async refreshCurrentNote(): Promise<void> {
    const id = this.selectedNoteId();
    if (id === null) return;
    await this.loadNotes();
    // Must separate the two sets into different microtasks so Angular
    // detects a real change (null → id) instead of batching id → null → id.
    this.selectedNoteId.set(null);
    await new Promise<void>(r => setTimeout(r));
    this.selectedNoteId.set(id);
  }

  /** Crea una nota de texto vacía y la deja seleccionada. */
  async createNote(): Promise<Note> {
    const note = await this.api.notesCreate(this.selectedFolderId());
    await this.loadNotes();
    this.selectedNoteId.set(note.id);
    return note;
  }

  /** Crea un libro XLSX vacío, lo adjunta a una nota y selecciona el resultado. */
  async createExcelNote(): Promise<Note> {
    const result = await this.api.xlsxCreateBlank();
    if (!result.success || !result.base64) throw new Error('Failed to create blank xlsx');
    const note = await this.api.notesCreate(this.selectedFolderId());
    await this.api.notesUpdate(note.id, {
      file_data: result.base64,
      file_type: 'xlsx',
    });
    await this.loadNotes();
    this.selectedNoteId.set(note.id);
    return note;
  }

  /** Crea una nota con el documento de presentación inicial serializado como JSON. */
  async createPresentationNote(): Promise<Note> {
    const note = await this.api.notesCreate(this.selectedFolderId());
    await this.api.notesUpdate(note.id, {
      content: JSON.stringify(createPresentationDocument()),
      file_type: 'pptx',
    });
    await this.loadNotes();
    this.selectedNoteId.set(note.id);
    return note;
  }

  /** Cambia la nota activa sin realizar acceso adicional a la base. */
  async selectNote(id: number): Promise<void> {
    this.selectedNoteId.set(id);
  }

  /** Persiste los campos editables de una nota y refresca el listado. */
  async updateNote(id: number, data: { title?: string; content?: string }): Promise<void> {
    await this.api.notesUpdate(id, data);
    await this.loadNotes();
  }

  /** Elimina una nota y limpia la selección cuando corresponde. */
  async deleteNote(id: number): Promise<void> {
    await this.api.notesDelete(id);
    if (this.selectedNoteId() === id) {
      this.selectedNoteId.set(null);
    }
    await this.loadNotes();
  }

  /** Mueve una nota a otra carpeta y la retira de la vista actual. */
  async moveNoteToFolder(noteId: number, folderId: number | null): Promise<void> {
    await this.api.notesUpdate(noteId, { folder_id: folderId });
    // Reload notes so the moved note disappears from the current view
    await this.loadNotes();
    // Deselect if the moved note was selected (it's now in a different folder)
    if (this.selectedNoteId() === noteId) {
      this.selectedNoteId.set(null);
    }
  }

  /** Importa archivos físicos admitidos y devuelve cuántas notas fueron creadas. */
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

  /** Importa uno o más respaldos JSON legados como notas de la carpeta actual. */
  async importJsonNotes(): Promise<number> {
    const files = await this.api.fileImportJson();
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
