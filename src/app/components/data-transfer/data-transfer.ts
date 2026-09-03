import { Component, computed, inject, OnInit, signal } from '@angular/core';
import { Note } from '../../models/note.model';
import { NavigationService } from '../../services/navigation.service';
import { NoteService } from '../../services/note.service';

@Component({
  selector: 'app-data-transfer',
  templateUrl: './data-transfer.html',
  styleUrl: './data-transfer.scss',
})
/**
 * Pantalla de respaldo y restauración de la biblioteca.
 * Mantiene selecciones independientes de carpetas y notas y delega el acceso
 * al sistema de archivos al proceso principal de Electron.
 */
export class DataTransfer implements OnInit {
  protected readonly noteService = inject(NoteService);
  private readonly navigation = inject(NavigationService);

  protected allNotes = signal<Note[]>([]);
  protected selectedFolderIds = signal<Set<number>>(new Set());
  protected selectedNoteIds = signal<Set<number>>(new Set());
  protected busy = signal(false);
  protected status = signal('');
  protected errors = signal<string[]>([]);
  protected lastExportPath = signal<string | null>(null);

  protected selectedItemCount = computed(() => {
    const folders = this.selectedFolderIds();
    const individualNotes = this.allNotes().filter(note =>
      this.selectedNoteIds().has(note.id) && !folders.has(note.folder_id ?? -1),
    ).length;
    return folders.size + individualNotes;
  });
  protected selectedNoteCount = computed(() => {
    const folders = this.selectedFolderIds();
    return this.allNotes().filter(note => folders.has(note.folder_id ?? -1) || this.selectedNoteIds().has(note.id)).length;
  });

  async ngOnInit(): Promise<void> {
    await this.refreshLibrary();
  }

  protected goBack(): void {
    this.navigation.goToNotes();
  }

  protected folderNoteCount(folderId: number): number {
    return this.allNotes().filter(note => note.folder_id === folderId).length;
  }

  protected folderName(folderId: number | null): string {
    if (folderId === null) return 'Sin carpeta';
    return this.noteService.folders().find(folder => folder.id === folderId)?.name ?? 'Carpeta desconocida';
  }

  protected isFolderSelected(folderId: number): boolean {
    return this.selectedFolderIds().has(folderId);
  }

  protected isNoteSelected(note: Note): boolean {
    return this.selectedFolderIds().has(note.folder_id ?? -1) || this.selectedNoteIds().has(note.id);
  }

  protected noteCoveredByFolder(note: Note): boolean {
    return note.folder_id !== null && this.selectedFolderIds().has(note.folder_id);
  }

  protected toggleFolder(folderId: number, checked: boolean): void {
    const next = new Set(this.selectedFolderIds());
    checked ? next.add(folderId) : next.delete(folderId);
    this.selectedFolderIds.set(next);
  }

  protected toggleNote(noteId: number, checked: boolean): void {
    const next = new Set(this.selectedNoteIds());
    checked ? next.add(noteId) : next.delete(noteId);
    this.selectedNoteIds.set(next);
  }

  protected selectEverything(): void {
    this.selectedFolderIds.set(new Set(this.noteService.folders().map(folder => folder.id)));
    this.selectedNoteIds.set(new Set(this.allNotes().filter(note => note.folder_id === null).map(note => note.id)));
  }

  protected clearSelection(): void {
    this.selectedFolderIds.set(new Set());
    this.selectedNoteIds.set(new Set());
  }

  /** Exporta la selección como respaldo estructurado JSON o ZIP físico. */
  protected async exportSelection(format: 'json' | 'zip'): Promise<void> {
    if (this.selectedItemCount() === 0 || this.busy()) return;
    this.busy.set(true);
    this.status.set('Preparando la exportación…');
    this.errors.set([]);
    this.lastExportPath.set(null);
    try {
      const result = await window.electronAPI.backupExport({
        folderIds: [...this.selectedFolderIds()],
        noteIds: [...this.selectedNoteIds()],
        format,
      });
      if (result.success && result.filePath) {
        this.lastExportPath.set(result.filePath);
        this.status.set(`Exportación completada: ${result.folders ?? 0} carpeta(s) y ${result.notes ?? 0} nota(s).`);
      } else if (result.canceled) {
        this.status.set('Exportación cancelada.');
      } else {
        this.status.set('No se pudo completar la exportación.');
        if (result.error) this.errors.set([result.error]);
      }
    } finally {
      this.busy.set(false);
    }
  }

  protected async importFiles(): Promise<void> {
    await this.runImport(() => window.electronAPI.backupImportFiles());
  }

  protected async importFolder(): Promise<void> {
    await this.runImport(() => window.electronAPI.backupImportFolder());
  }

  protected showExportedFile(): void {
    const filePath = this.lastExportPath();
    if (filePath) void window.electronAPI.fileShowInFolder(filePath);
  }

  /** Ejecuta una importación común, presenta resultados parciales y refresca la biblioteca. */
  private async runImport(operation: () => ReturnType<typeof window.electronAPI.backupImportFiles>): Promise<void> {
    if (this.busy()) return;
    this.busy.set(true);
    this.status.set('Importando estructura…');
    this.errors.set([]);
    try {
      const result = await operation();
      if (result.canceled) {
        this.status.set('Importación cancelada.');
        return;
      }
      this.errors.set(result.errors ?? (result.error ? [result.error] : []));
      if (result.success || result.partial) {
        this.status.set(`Importación completada: ${result.folders ?? 0} carpeta(s) y ${result.notes ?? 0} nota(s).`);
        await this.refreshLibrary();
      } else {
        this.status.set('No se encontró una estructura compatible para importar.');
      }
    } finally {
      this.busy.set(false);
    }
  }

  /** Sincroniza carpetas, notas globales y la vista seleccionada después de importar. */
  private async refreshLibrary(): Promise<void> {
    await this.noteService.loadFolders();
    this.allNotes.set(await window.electronAPI.notesGetByFolder(null));
    await this.noteService.loadNotes();
  }
}
