import { mkdtempSync, writeFileSync, rmSync, readFileSync } from 'fs';
import os from 'os';
import path from 'path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { importDocumentPaths } from '../electron/file-import';
import { SUPPORTED_DOCUMENT_EXTENSIONS, isSupportedDocument } from '../electron/document-types';
import { queueOpenFiles, takePendingOpenFiles } from '../electron/file-open';

describe('external document import', () => {
  let directory: string;
  beforeEach(() => {
    directory = mkdtempSync(path.join(os.tmpdir(), 'helper-notes-import-'));
    takePendingOpenFiles();
  });
  afterEach(() => {
    rmSync(directory, { recursive: true, force: true });
    takePendingOpenFiles();
  });

  /**
   * Escribe un documento en el directorio temporal aislado de este caso.
   * @param name Nombre del archivo con extensión.
   * @param content Contenido textual o binario del documento.
   * @returns Ruta absoluta del archivo creado.
   */
  function fixture(name: string, content: string | Uint8Array): string {
    const filePath = path.join(directory, name);
    writeFileSync(filePath, content);
    return filePath;
  }

  it('imports text and HTML, preserving originals and ignoring duplicate paths', () => {
    const text = fixture('Notas.TXT', '<texto>\nsegunda línea');
    const html = fixture('page.htm', '<html><body><h1>Title</h1></body></html>');
    const result = importDocumentPaths([text, html, text]);
    expect(result.errors).toEqual([]);
    expect(result.files).toEqual([
      {
        title: 'Notas',
        content: '<texto>\nsegunda línea',
        file_type: 'txt',
        file_data: Buffer.from('<texto>\nsegunda línea').toString('base64'),
      },
      { title: 'page', content: '<h1>Title</h1>' },
    ]);
    expect(readFileSync(text, 'utf8')).toBe('<texto>\nsegunda línea');
  });

  it('preserves PDF and spreadsheet attachments as Base64', () => {
    for (const extension of ['pdf', 'xlsx', 'xls', 'csv']) {
      const filePath = fixture(`document.${extension}`, 'original content');
      expect(importDocumentPaths([filePath]).files[0]).toMatchObject({
        file_type: extension === 'pdf' ? 'pdf' : 'xlsx',
        file_data: Buffer.from('original content').toString('base64'),
      });
    }
  });

  it('converts DOCX and PPTX to editable documents', () => {
    const docx = fixture(
      'word.docx',
      zipSync({
        'word/document.xml': strToU8(
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:document>',
        ),
      }),
    );
    expect(importDocumentPaths([docx]).files[0].content).toBe('<p>Hello</p>');
    expect(importDocumentPaths([docx]).files[0]).toMatchObject({
      file_type: 'docx',
      file_data: readFileSync(docx).toString('base64'),
    });
    const pptx = fixture(
      'slides.pptx',
      zipSync({ 'ppt/slides/slide1.xml': strToU8('<p:sld><p:sp><a:t>Hello</a:t></p:sp></p:sld>') }),
    );
    const result = importDocumentPaths([pptx]);
    expect(result.errors).toEqual([]);
    expect(result.files[0].file_type).toBe('pptx');
    expect(JSON.parse(result.files[0].content).kind).toBe('helper-notes-presentation');
  });

  it('never leaks OOXML tab or table markup into text and decodes text entities once', () => {
    const docx = fixture(
      'tabs.docx',
      zipSync({
        'word/document.xml': strToU8(
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:tabs><w:tab w:pos="720"/></w:tabs></w:pPr><w:r><w:tab/><w:t>Tom &amp; Jerry &lt;word&gt;</w:t><w:br/><w:t>next</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Table</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>',
        ),
      }),
    );
    const result = importDocumentPaths([docx]);
    expect(result.files[0].content).toBe(
      '<p>&#9;Tom &amp; Jerry &lt;word&gt;<br>next</p><p>Table</p>',
    );
    expect(result.files[0].content).not.toContain('w:');
    expect(result.files[0].content).not.toContain('&amp;amp;');
  });

  it('preserves whitespace and Unicode in UTF-16 text documents', () => {
    const text = '  Texto\tcon espacios\r\nEspaña';
    const file = fixture(
      'unicode.txt',
      Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')]),
    );
    expect(importDocumentPaths([file]).files[0].content).toBe(text);
  });

  it('reports invalid files individually and still imports compatible files', () => {
    const good = fixture('good.txt', 'valid');
    const executable = fixture('unsupported.exe', 'invalid');
    const corrupt = fixture('corrupt.pptx', 'invalid zip');
    const result = importDocumentPaths([
      good,
      executable,
      corrupt,
      directory,
      path.join(directory, 'missing.txt'),
    ]);
    expect(result.files).toHaveLength(1);
    expect(result.errors).toHaveLength(4);
    expect(() => importDocumentPaths(null as unknown as string[])).toThrow();
  });

  it('queues startup and second-instance paths once until the renderer consumes them', () => {
    expect(queueOpenFiles(['--inspect', 'relative.TXT', 'ignore.exe'], directory)).toBe(true);
    queueOpenFiles([path.join(directory, 'relative.TXT')]);
    expect(takePendingOpenFiles()).toEqual([path.join(directory, 'relative.TXT')]);
    expect(takePendingOpenFiles()).toEqual([]);
    expect(queueOpenFiles(['--flag', 'ignore.exe'], directory)).toBe(false);
  });

  it('registers only supported Explorer extensions and removes its own keys on uninstall', () => {
    const script = readFileSync(path.resolve('build/file-associations.nsh'), 'utf8');
    const registered = [
      ...script.matchAll(/!insertmacro HelperNotesRegisterExtension "([^"]+)"/g),
    ].map((match) => match[1]);
    const unregistered = [
      ...script.matchAll(/!insertmacro HelperNotesUnregisterExtension "([^"]+)"/g),
    ].map((match) => match[1]);
    expect(registered).toEqual([...SUPPORTED_DOCUMENT_EXTENSIONS]);
    expect(unregistered).toEqual(registered);
    expect(registered.every((extension) => isSupportedDocument(`file.${extension}`))).toBe(true);
    expect(script).toContain('Abrir con Helper Notes');
    expect(script).toContain('OpenWithProgids');
    expect(script).not.toContain('WriteRegStr SHELL_CONTEXT "Software\\Classes\\.${EXT}"');
  });
});
