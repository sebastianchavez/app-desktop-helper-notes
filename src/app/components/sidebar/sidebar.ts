import { NgTemplateOutlet } from '@angular/common';
import { Component, HostListener, inject, input, OnInit, output, signal } from '@angular/core';
import { TranslateModule, TranslateService } from '@ngx-translate/core';
import { NoteService } from '../../services/note.service';
import { NavigationService } from '../../services/navigation.service';
import { AiChatService } from '../../services/ai-chat.service';
import { Folder } from '../../models/folder.model';

@Component({
  selector: 'app-sidebar',
  imports: [TranslateModule, NgTemplateOutlet],
  templateUrl: './sidebar.html',
  styleUrl: './sidebar.scss',
})
/**
 * Árbol de navegación de carpetas y accesos a vistas auxiliares.
 * Implementa edición inline, menú contextual, expansión jerárquica y destino
 * de drag-and-drop para mover notas, archivos y ramas de carpetas.
 */
export class Sidebar implements OnInit {
  protected readonly noteService = inject(NoteService);
  protected readonly navService = inject(NavigationService);
  private readonly translate = inject(TranslateService);
  private readonly aiChatService = inject(AiChatService);
  protected editingFolderId = signal<number | null>(null);
  protected editingFolderName = signal('');
  protected showNewFolderInput = signal(false);
  protected newFolderName = signal('');
  protected newFolderParentId = signal<number | null>(null);
  protected expandedFolderIds = signal<Set<number>>(new Set());
  protected contextMenu = signal<{ x: number; y: number; folderId: number | null } | null>(null);
  protected dragOverFolderId = signal<number | null | 'all'>(null);
  protected draggedFolderId = signal<number | null>(null);
  protected moveError = signal(false);

  /** Whether the sidebar is in collapsed state */
  collapsed = input(false);
  /** Emits when user clicks the header to toggle collapse */
  collapseToggle = output<void>();

  /** Carga el árbol y lo deja expandido durante el primer render. */
  async ngOnInit(): Promise<void> {
    await this.noteService.loadFolders();
    this.expandedFolderIds.set(new Set(this.noteService.folders().map(folder => folder.id)));
  }

  protected rootFolders(): Folder[] {
    return this.noteService.folders().filter(folder => folder.parent_id === null);
  }

  protected folderChildren(parentId: number): Folder[] {
    return this.noteService.folders().filter(folder => folder.parent_id === parentId);
  }

  protected hasChildren(folderId: number): boolean {
    return this.noteService.folders().some(folder => folder.parent_id === folderId);
  }

  protected isExpanded(folderId: number): boolean {
    return this.expandedFolderIds().has(folderId);
  }

  protected toggleExpanded(folderId: number, event: Event): void {
    event.stopPropagation();
    const next = new Set(this.expandedFolderIds());
    next.has(folderId) ? next.delete(folderId) : next.add(folderId);
    this.expandedFolderIds.set(next);
  }

  /** Selecciona una carpeta salvo que una generación de IA esté modificando contexto. */
  async selectFolder(id: number | null): Promise<void> {
    if (this.aiChatService.isStreaming()) return;
    this.navService.goToNotes();
    await this.noteService.selectFolder(id);
  }

  goToSettings(): void {
    if (this.aiChatService.isStreaming()) return;
    this.navService.goToSettings();
  }

  goToAssistantSettings(): void {
    if (this.aiChatService.isStreaming()) return;
    this.navService.goToAssistantSettings();
  }

  goToDataTransfer(): void {
    if (this.aiChatService.isStreaming()) return;
    this.navService.goToDataTransfer();
  }

  showNewFolder(parentId: number | null = null): void {
    this.closeContextMenu();
    if (parentId !== null) {
      const next = new Set(this.expandedFolderIds());
      next.add(parentId);
      this.expandedFolderIds.set(next);
    }
    this.newFolderParentId.set(parentId);
    this.showNewFolderInput.set(true);
    this.newFolderName.set('');
  }

  async confirmNewFolder(): Promise<void> {
    const name = this.newFolderName().trim();
    if (name) {
      await this.noteService.createFolder(name, this.newFolderParentId());
    }
    this.showNewFolderInput.set(false);
    this.newFolderName.set('');
    this.newFolderParentId.set(null);
  }

  cancelNewFolder(): void {
    this.showNewFolderInput.set(false);
    this.newFolderName.set('');
    this.newFolderParentId.set(null);
  }

  /** Abre un menú dentro del viewport para evitar que quede recortado. */
  protected openContextMenu(event: MouseEvent, folderId: number | null): void {
    event.preventDefault();
    event.stopPropagation();
    const width = 230;
    const height = 210;
    this.contextMenu.set({
      x: Math.max(8, Math.min(event.clientX, window.innerWidth - width - 8)),
      y: Math.max(8, Math.min(event.clientY, window.innerHeight - height - 8)),
      folderId,
    });
  }

  protected closeContextMenu(): void {
    this.contextMenu.set(null);
  }

  @HostListener('document:click')
  @HostListener('document:keydown.escape')
  protected closeTransientMenus(): void {
    this.closeContextMenu();
  }

  /** Crea una nota del tipo elegido dentro de la carpeta indicada. */
  protected async createFileInFolder(type: 'note' | 'spreadsheet' | 'presentation', folderId: number | null): Promise<void> {
    this.closeContextMenu();
    if (this.aiChatService.isStreaming()) return;
    await this.selectFolder(folderId);
    if (type === 'spreadsheet') {
      await this.noteService.createExcelNote();
    } else if (type === 'presentation') {
      await this.noteService.createPresentationNote();
    } else {
      await this.noteService.createNote();
    }
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
  onFolderDragStart(event: DragEvent, folderId: number): void {
    const button = (event.target as Element | null)?.closest('button');
    if (!event.dataTransfer || this.aiChatService.isStreaming() ||
        this.editingFolderId() !== null || (button && button !== event.currentTarget)) {
      event.preventDefault();
      return;
    }
    this.closeContextMenu();
    this.moveError.set(false);
    this.draggedFolderId.set(folderId);
    event.dataTransfer.setData('application/folder-id', String(folderId));
    event.dataTransfer.effectAllowed = 'move';
    event.stopPropagation();
  }

  @HostListener('document:dragend')
  @HostListener('document:drop')
  onDragEnd(): void {
    this.draggedFolderId.set(null);
    this.dragOverFolderId.set(null);
  }

  /** Indica sólo destinos válidos para los MIME internos de notas y carpetas. */
  onDragOver(event: DragEvent, folderId: number | null): void {
    const transfer = event.dataTransfer;
    if (!transfer) return;
    const isFolder = transfer.types.includes('application/folder-id');
    const isNote = transfer.types.includes('application/note-id');
    if (!isFolder && !isNote) return;
    event.stopPropagation();
    const sourceId = this.draggedFolderId();
    if (this.aiChatService.isStreaming() || (isFolder &&
        (sourceId === null || !this.noteService.canMoveFolder(sourceId, folderId)))) {
      transfer.dropEffect = 'none';
      this.dragOverFolderId.set(null);
      return;
    }
    event.preventDefault();
    transfer.dropEffect = 'move';
    this.dragOverFolderId.set(folderId ?? 'all');
  }

  onDragLeave(event: DragEvent): void {
    if (event.relatedTarget instanceof Node &&
        (event.currentTarget as HTMLElement).contains(event.relatedTarget)) return;
    this.dragOverFolderId.set(null);
  }

  /** Mueve el elemento y abre el destino para mostrar la nueva ubicación de la rama. */
  async onDrop(event: DragEvent, folderId: number | null): Promise<void> {
    event.preventDefault();
    event.stopPropagation();
    this.onDragEnd();
    if (this.aiChatService.isStreaming()) return;
    const transfer = event.dataTransfer;
    const isFolder = transfer?.types.includes('application/folder-id');
    const rawId = transfer?.getData(isFolder ? 'application/folder-id' : 'application/note-id');
    if (!rawId || !/^\d+$/.test(rawId)) return;
    const id = Number(rawId);
    if (!Number.isSafeInteger(id) || id <= 0) return;
    this.moveError.set(false);
    try {
      if (isFolder) {
        if (!this.noteService.canMoveFolder(id, folderId)) return;
        await this.noteService.moveFolder(id, folderId);
        if (folderId !== null) {
          this.expandedFolderIds.update(ids => new Set([...ids, folderId]));
        }
      } else {
        await this.noteService.moveNoteToFolder(id, folderId);
      }
    } catch {
      this.moveError.set(true);
    }
  }
}
