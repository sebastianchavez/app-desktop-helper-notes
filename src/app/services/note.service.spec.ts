import { TestBed } from '@angular/core/testing';
import { Note } from '../models/note.model';
import { Folder } from '../models/folder.model';
import { NoteService } from './note.service';

describe('NoteService', () => {
  const blankNote: Note = {
    id: 21,
    folder_id: 4,
    title: '',
    content: '',
    file_data: null,
    file_type: null,
    created_at: '2026-08-30 12:00:00',
    updated_at: '2026-08-30 12:00:00',
  };
  let createdInFolder: number | null | undefined;
  let lastUpdate: Record<string, unknown> | null;
  let createdFolderParentId: number | null | undefined;

  beforeEach(() => {
    createdInFolder = undefined;
    lastUpdate = null;
    createdFolderParentId = undefined;
    const childFolder: Folder = {
      id: 9,
      name: 'Borradores',
      parent_id: 4,
      created_at: '2026-08-30 12:00:00',
      updated_at: '2026-08-30 12:00:00',
    };
    Object.defineProperty(window, 'electronAPI', {
      configurable: true,
      value: {
        onNoteChanged: () => () => undefined,
        foldersGetAll: async () => [childFolder],
        foldersCreate: async (_name: string, parentId: number | null) => {
          createdFolderParentId = parentId;
          return childFolder;
        },
        notesCreate: async (folderId: number | null) => {
          createdInFolder = folderId;
          return blankNote;
        },
        notesGetByFolder: async () => [blankNote],
        notesUpdate: async (_id: number, data: Record<string, unknown>) => {
          lastUpdate = data;
          return { ...blankNote, ...data };
        },
      },
    });
    TestBed.configureTestingModule({});
  });

  it('creates a subfolder under the requested parent folder', async () => {
    const service = TestBed.inject(NoteService);

    const folder = await service.createFolder('Borradores', 4);

    expect(createdFolderParentId).toBe(4);
    expect(folder.parent_id).toBe(4);
    expect(service.folders()).toEqual([folder]);
  });

  it('creates and selects a blank rich-text note in the active folder', async () => {
    const service = TestBed.inject(NoteService);
    service.selectedFolderId.set(4);

    const note = await service.createNote();

    expect(createdInFolder).toBe(4);
    expect(note.title).toBe('');
    expect(note.content).toBe('');
    expect(note.file_type).toBeNull();
    expect(service.selectedNoteId()).toBe(note.id);
    expect(service.selectedNote()).toEqual(blankNote);
  });

  it('creates a presentation note with one editable slide', async () => {
    const service = TestBed.inject(NoteService);
    const note = await service.createPresentationNote();

    expect(note.id).toBe(blankNote.id);
    expect(lastUpdate?.['file_type']).toBe('pptx');
    const document = JSON.parse(String(lastUpdate?.['content']));
    expect(document.kind).toBe('helper-notes-presentation');
    expect(document.slides).toHaveLength(1);
  });
});
