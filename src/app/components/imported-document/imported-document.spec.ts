import { TestBed } from '@angular/core/testing';
import { provideTranslateService } from '@ngx-translate/core';
import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate';
import { ImportedDocument } from './imported-document';
import { Note } from '../../models/note.model';
import { WORD_EDITOR_FACTORY } from './word-editor.factory';
import type { Editor } from '@docx-editor.dev/core';

/**
 * Genera un DOCX con páginas, una tabla y formato real para verificar el visor completo.
 * @returns Bytes del archivo de prueba; no usa documentos ni datos del usuario.
 */
function wordFixture(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    '_rels/.rels': strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    'word/document.xml': strToU8(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/><w:color w:val="123456"/></w:rPr><w:t>Original Word heading</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:br w:type="page"/><w:t>Second page</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:bottom="720" w:left="720" w:right="720"/></w:sectPr></w:body></w:document>',
    ),
  });
}

describe('ImportedDocument', () => {
  const bytes = wordFixture();
  const base64 = btoa(Array.from(bytes, (byte) => String.fromCharCode(byte)).join(''));
  const word: Note = {
    id: 77,
    folder_id: null,
    title: 'Word document',
    content: '<p>Searchable text</p>',
    file_type: 'docx',
    file_data: base64,
    created_at: '',
    updated_at: '',
  };
  let saveBuffer: ReturnType<typeof vi.fn>;
  let mountEditor: ReturnType<typeof vi.fn>;
  let editor: Editor;
  beforeEach(async () => {
    saveBuffer = vi.fn(async () => ({ success: false }));
    editor = {
      save: vi.fn(async () => bytes.slice().buffer),
      on: vi.fn(() => () => undefined),
    } as unknown as Editor;
    mountEditor = vi.fn(async () => ({ ready: Promise.resolve(editor), destroy: vi.fn() }));
    Object.defineProperty(window, 'electronAPI', {
      configurable: true,
      value: {
        notesGetById: async () => word,
        onNoteChanged: () => () => undefined,
        fileSaveAsBuffer: saveBuffer,
      },
    });
    await TestBed.configureTestingModule({
      imports: [ImportedDocument],
      providers: [
        provideTranslateService({ fallbackLang: 'es' }),
        { provide: WORD_EDITOR_FACTORY, useValue: mountEditor },
      ],
    }).compileComponents();
  });

  it('opens the complete DOCX for editing and retains the engine when metadata changes', async () => {
    const fixture = TestBed.createComponent(ImportedDocument);
    fixture.componentRef.setInput('note', word);
    fixture.componentRef.setInput('title', word.title);
    fixture.detectChanges();
    await fixture.whenStable();
    await vi.waitFor(() => {
      fixture.detectChanges();
      expect(fixture.nativeElement.querySelector('.word-host').getAttribute('aria-busy')).toBe(
        'false',
      );
    });
    expect(mountEditor).toHaveBeenCalledTimes(1);
    expect(Array.from(mountEditor.mock.calls[0][1])).toEqual(Array.from(bytes));
    fixture.componentRef.setInput('note', {
      ...word,
      content: '<p>Updated search text</p>',
      updated_at: 'new',
    });
    fixture.detectChanges();
    await fixture.whenStable();
    expect(mountEditor).toHaveBeenCalledTimes(1);
    expect(await fixture.componentInstance.getPrintableContent()).toContain('Table cell');
  });

  it('exports the live edited DOCX instead of downloading the original attachment', async () => {
    const fixture = TestBed.createComponent(ImportedDocument);
    fixture.componentRef.setInput('note', word);
    fixture.componentRef.setInput('title', word.title);
    fixture.detectChanges();
    await fixture.whenStable();
    await vi.waitFor(() => {
      fixture.detectChanges();
      expect(fixture.nativeElement.querySelector('.word-host').getAttribute('aria-busy')).toBe(
        'false',
      );
    });
    const editedEntries = unzipSync(wordFixture());
    editedEntries['word/document.xml'] = strToU8(
      strFromU8(editedEntries['word/document.xml']).replace(
        'Original Word heading',
        'Edited Word heading',
      ),
    );
    const edited = zipSync(editedEntries);
    editor.save = vi.fn(async () => edited.slice().buffer);
    const buttons = fixture.nativeElement.querySelectorAll(
      'button',
    ) as NodeListOf<HTMLButtonElement>;
    buttons[1].click();
    await vi.waitFor(() => expect(saveBuffer).toHaveBeenCalled());
    expect(saveBuffer.mock.calls[0][0].buffer).toEqual(Array.from(edited));
    expect(saveBuffer.mock.calls[0][0].defaultName).toBe('Word document.docx');
    expect(editor.save).toHaveBeenCalled();
  });

  it('shows TXT markup literally and emits plain text edits', () => {
    const fixture = TestBed.createComponent(ImportedDocument);
    const text = '<b>literal</b>\n\t spaced text';
    fixture.componentRef.setInput('note', { ...word, file_type: 'txt', content: text });
    fixture.componentRef.setInput('text', text);
    const changed = vi.fn();
    fixture.componentInstance.textChange.subscribe(changed);
    fixture.detectChanges();
    const textarea = fixture.nativeElement.querySelector('textarea') as HTMLTextAreaElement;
    expect(textarea.value).toBe(text);
    expect(fixture.nativeElement.querySelector('[contenteditable]')).toBeNull();
    textarea.value = '  edited\n\ttext';
    textarea.dispatchEvent(new Event('input'));
    expect(changed).toHaveBeenCalledWith('  edited\n\ttext');
  });
});
