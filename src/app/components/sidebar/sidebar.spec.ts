import { TestBed } from '@angular/core/testing';
import { signal } from '@angular/core';
import { provideTranslateService } from '@ngx-translate/core';
import { Sidebar } from './sidebar';
import { NoteService } from '../../services/note.service';
import { AiChatService } from '../../services/ai-chat.service';
import { Folder } from '../../models/folder.model';

describe('Sidebar drag and drop', () => {
  let folders: Folder[];
  let foldersMove: ReturnType<typeof vi.fn>;
  let notesUpdate: ReturnType<typeof vi.fn>;
  const isStreaming = signal(false);

  function dragEvent(type: string, mime: string, id: string): DragEvent {
    const data = new Map([[mime, id]]);
    const event = new Event(type, { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'dataTransfer', {
      value: {
        types: [mime],
        getData: (key: string) => data.get(key) ?? '',
        setData: (key: string, value: string) => data.set(key, value),
        dropEffect: 'none',
        effectAllowed: 'none',
      },
    });
    return event as DragEvent;
  }

  beforeEach(async () => {
    isStreaming.set(false);
    folders = [
      { id: 1, name: 'Origen', parent_id: null },
      { id: 2, name: 'Hija', parent_id: 1 },
      { id: 3, name: 'Nieta', parent_id: 2 },
      { id: 4, name: 'Destino', parent_id: null },
    ].map(folder => ({ ...folder, created_at: '', updated_at: '' }));
    foldersMove = vi.fn(async (id: number, parentId: number | null) => {
      folders = folders.map(folder => folder.id === id ? { ...folder, parent_id: parentId } : folder);
      return folders.find(folder => folder.id === id);
    });
    notesUpdate = vi.fn(async () => undefined);
    Object.defineProperty(window, 'electronAPI', {
      configurable: true,
      value: {
        onNoteChanged: () => () => undefined,
        foldersGetAll: async () => folders,
        foldersMove,
        notesGetByFolder: async () => [],
        notesUpdate,
      },
    });
    await TestBed.configureTestingModule({
      imports: [Sidebar],
      providers: [
        provideTranslateService({ fallbackLang: 'es' }),
        { provide: AiChatService, useValue: { isStreaming } },
      ],
    }).compileComponents();
  });

  async function setup(collapsed = false) {
    const fixture = TestBed.createComponent(Sidebar);
    fixture.componentRef.setInput('collapsed', collapsed);
    fixture.detectChanges();
    await fixture.whenStable();
    await vi.waitFor(() => {
      fixture.detectChanges();
      expect(fixture.nativeElement.querySelectorAll('[draggable="true"]')).toHaveLength(4);
    });
    const rows = Array.from(fixture.nativeElement.querySelectorAll('[draggable="true"]')) as HTMLElement[];
    const source = rows.find(row => (collapsed ? row.title : row.textContent)?.includes('Origen'))!;
    const destination = rows.find(row => (collapsed ? row.title : row.textContent)?.includes('Destino'))!;
    return { fixture, rows, source, destination, sidebar: fixture.componentInstance };
  }

  for (const collapsed of [false, true]) {
    it(`moves a folder from the ${collapsed ? 'collapsed' : 'expanded'} sidebar and preserves its subtree and selection`, async () => {
      const { fixture, source, destination } = await setup(collapsed);
      const service = TestBed.inject(NoteService);
      service.selectedFolderId.set(2);
      service.selectedNoteId.set(21);
      const start = dragEvent('dragstart', 'application/folder-id', '');
      source.dispatchEvent(start);
      expect(start.dataTransfer!.getData('application/folder-id')).toBe('1');
      const over = dragEvent('dragover', 'application/folder-id', '1');
      destination.dispatchEvent(over);
      expect(over.defaultPrevented).toBe(true);
      destination.dispatchEvent(dragEvent('drop', 'application/folder-id', '1'));
      await fixture.whenStable();
      expect(foldersMove).toHaveBeenCalledWith(1, 4);
      expect(service.folders().find(folder => folder.id === 2)?.parent_id).toBe(1);
      expect(service.selectedFolderId()).toBe(2);
      expect(service.selectedNoteId()).toBe(21);
    });
  }

  it('rejects self, descendants, unchanged parents and missing destinations', async () => {
    const { sidebar, source, rows } = await setup();
    source.dispatchEvent(dragEvent('dragstart', 'application/folder-id', '1'));
    const descendant = rows.find(row => row.textContent?.includes('Nieta'))!;
    const over = dragEvent('dragover', 'application/folder-id', '1');
    descendant.dispatchEvent(over);
    expect(over.defaultPrevented).toBe(false);
    expect(over.dataTransfer!.dropEffect).toBe('none');
    for (const destination of [1, 2, 3, null, 99]) {
      await sidebar.onDrop(dragEvent('drop', 'application/folder-id', '1'), destination);
    }
    expect(foldersMove).not.toHaveBeenCalled();
  });

  it('moves a subfolder to the root through All notes', async () => {
    const { sidebar } = await setup();
    await sidebar.onDrop(dragEvent('drop', 'application/folder-id', '2'), null);
    expect(foldersMove).toHaveBeenCalledWith(2, null);
  });

  it('continues moving notes and attached files with the existing MIME', async () => {
    const { sidebar } = await setup();
    await sidebar.onDrop(dragEvent('drop', 'application/note-id', '21'), 4);
    expect(notesUpdate).toHaveBeenCalledWith(21, { folder_id: 4 });
    expect(foldersMove).not.toHaveBeenCalled();
  });

  it('rejects malformed identifiers and external files', async () => {
    const { sidebar } = await setup();
    for (const id of ['1bad', '-1', '0', '9007199254740992']) {
      await sidebar.onDrop(dragEvent('drop', 'application/folder-id', id), 4);
    }
    await sidebar.onDrop(dragEvent('drop', 'Files', '1'), 4);
    expect(foldersMove).not.toHaveBeenCalled();
    expect(notesUpdate).not.toHaveBeenCalled();
  });

  it('clears highlighting when a drag is cancelled', async () => {
    const { fixture, source, destination } = await setup();
    source.dispatchEvent(dragEvent('dragstart', 'application/folder-id', '1'));
    destination.dispatchEvent(dragEvent('dragover', 'application/folder-id', '1'));
    fixture.detectChanges();
    expect(destination.classList.contains('ring-2')).toBe(true);
    source.dispatchEvent(dragEvent('dragend', 'application/folder-id', '1'));
    fixture.detectChanges();
    expect(destination.classList.contains('ring-2')).toBe(false);
  });

  it('reports a persistence failure without changing the tree', async () => {
    const { fixture, sidebar } = await setup();
    foldersMove.mockRejectedValueOnce(new Error('Database unavailable'));
    await sidebar.onDrop(dragEvent('drop', 'application/folder-id', '1'), 4);
    fixture.detectChanges();
    expect(fixture.nativeElement.querySelector('[role="alert"]')).toBeTruthy();
    expect(TestBed.inject(NoteService).folders().find(folder => folder.id === 1)?.parent_id).toBeNull();
  });

  it('blocks moves while the assistant is streaming', async () => {
    const { sidebar, source } = await setup();
    isStreaming.set(true);
    const start = dragEvent('dragstart', 'application/folder-id', '1');
    source.dispatchEvent(start);
    expect(start.defaultPrevented).toBe(true);
    await sidebar.onDrop(dragEvent('drop', 'application/folder-id', '1'), 4);
    expect(foldersMove).not.toHaveBeenCalled();
  });
});
