import {
  Component,
  inject,
  signal,
  effect,
  computed,
  ElementRef,
  viewChild,
  HostListener,
  untracked,
} from '@angular/core';
import { NoteService } from '../../services/note.service';

@Component({
  selector: 'app-note-editor',
  templateUrl: './note-editor.html',
  styleUrl: './note-editor.scss',
})
export class NoteEditor {
  protected readonly noteService = inject(NoteService);

  protected title = signal('');
  protected content = signal('');

  // Toolbar state signals
  protected isBold = signal(false);
  protected isItalic = signal(false);
  protected isUnderline = signal(false);
  protected isStrike = signal(false);
  protected isUnorderedList = signal(false);
  protected isOrderedList = signal(false);
  protected activeAlign = signal<'left' | 'center' | 'right' | 'justify'>('left');
  protected activeBlockFormat = signal('p');
  protected activeFontSize = signal('3');
  protected activeTextColor = signal('#ffffff');
  protected showDownloadMenu = signal(false);

  private editorRef = viewChild<ElementRef<HTMLDivElement>>('editorEl');
  private downloadDropdownRef = viewChild<ElementRef<HTMLDivElement>>('downloadDropdown');
  private saveTimeout: ReturnType<typeof setTimeout> | null = null;
  private skipNextSync = false;

  constructor() {
    // Sync local state when selected note changes
    effect(() => {
      const note = this.noteService.selectedNote();
      const skip = untracked(() => this.skipNextSync);
      if (skip) {
        this.skipNextSync = false;
        return;
      }
      if (note) {
        this.title.set(note.title);
        this.content.set(note.content);
        // Directly update the contenteditable element
        const el = untracked(() => this.editorRef()?.nativeElement);
        if (el) {
          el.innerHTML = note.content || '';
        }
      }
    });
  }

  protected readonly wordCount = computed(() => {
    const text = this.stripHtml(this.content());
    if (!text.trim()) return 0;
    return text.trim().split(/\s+/).length;
  });

  protected readonly charCount = computed(() => this.stripHtml(this.content()).length);

  protected readonly lastEdited = computed(() => {
    const note = this.noteService.selectedNote();
    if (!note) return '';
    const date = new Date(note.updated_at + 'Z');
    return date.toLocaleDateString('es', {
      day: 'numeric',
      month: 'long',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  });

  protected readonly folderName = computed(() => {
    const note = this.noteService.selectedNote();
    if (!note || !note.folder_id) return null;
    const folder = this.noteService.folders().find((f) => f.id === note.folder_id);
    return folder?.name ?? null;
  });

  onTitleChange(value: string): void {
    this.title.set(value);
    this.scheduleSave();
  }

  onEditorInput(): void {
    const el = this.editorRef()?.nativeElement;
    if (el) {
      this.content.set(el.innerHTML);
      this.scheduleSave();
    }
  }

  /** Execute a document.execCommand and refocus editor */
  execCommand(command: string, value?: string): void {
    document.execCommand(command, false, value ?? '');
    this.refocusEditor();
    this.refreshToolbarState();
    this.onEditorInput();
  }

  onFormatBlock(tag: string): void {
    document.execCommand('formatBlock', false, `<${tag}>`);
    this.refocusEditor();
    this.refreshToolbarState();
    this.onEditorInput();
  }

  onFontSize(size: string): void {
    document.execCommand('fontSize', false, size);
    this.refocusEditor();
    this.refreshToolbarState();
    this.onEditorInput();
  }

  onTextColor(color: string): void {
    document.execCommand('foreColor', false, color);
    this.activeTextColor.set(color);
    this.refocusEditor();
    this.onEditorInput();
  }

  onHighlightColor(color: string): void {
    document.execCommand('hiliteColor', false, color);
    this.refocusEditor();
    this.onEditorInput();
  }

  clearFormatting(): void {
    document.execCommand('removeFormat', false);
    this.refocusEditor();
    this.refreshToolbarState();
    this.onEditorInput();
  }

  /** Refresh toolbar active states based on current selection */
  refreshToolbarState(): void {
    this.isBold.set(document.queryCommandState('bold'));
    this.isItalic.set(document.queryCommandState('italic'));
    this.isUnderline.set(document.queryCommandState('underline'));
    this.isStrike.set(document.queryCommandState('strikeThrough'));
    this.isUnorderedList.set(document.queryCommandState('insertUnorderedList'));
    this.isOrderedList.set(document.queryCommandState('insertOrderedList'));

    // Alignment
    if (document.queryCommandState('justifyCenter')) {
      this.activeAlign.set('center');
    } else if (document.queryCommandState('justifyRight')) {
      this.activeAlign.set('right');
    } else if (document.queryCommandState('justifyFull')) {
      this.activeAlign.set('justify');
    } else {
      this.activeAlign.set('left');
    }

    // Block format
    const blockVal = document.queryCommandValue('formatBlock')?.toLowerCase() ?? '';
    if (blockVal.includes('h1')) this.activeBlockFormat.set('h1');
    else if (blockVal.includes('h2')) this.activeBlockFormat.set('h2');
    else if (blockVal.includes('h3')) this.activeBlockFormat.set('h3');
    else this.activeBlockFormat.set('p');

    // Font size
    const sizeVal = document.queryCommandValue('fontSize');
    if (sizeVal) this.activeFontSize.set(sizeVal);

    // Text color
    const colorVal = document.queryCommandValue('foreColor');
    if (colorVal) {
      this.activeTextColor.set(this.rgbToHex(colorVal));
    }
  }

  private refocusEditor(): void {
    const el = this.editorRef()?.nativeElement;
    if (el) {
      el.focus();
    }
  }

  private scheduleSave(): void {
    if (this.saveTimeout) clearTimeout(this.saveTimeout);
    this.saveTimeout = setTimeout(() => {
      const noteId = this.noteService.selectedNoteId();
      if (noteId !== null) {
        this.skipNextSync = true;
        this.noteService.updateNote(noteId, {
          title: this.title(),
          content: this.content(),
        });
      }
    }, 500);
  }

  private stripHtml(html: string): string {
    const tmp = document.createElement('div');
    tmp.innerHTML = html;
    return tmp.textContent || tmp.innerText || '';
  }

  private rgbToHex(rgb: string): string {
    // If already hex, return as-is
    if (rgb.startsWith('#')) return rgb;
    const match = rgb.match(/\d+/g);
    if (!match || match.length < 3) return '#ffffff';
    const r = parseInt(match[0]).toString(16).padStart(2, '0');
    const g = parseInt(match[1]).toString(16).padStart(2, '0');
    const b = parseInt(match[2]).toString(16).padStart(2, '0');
    return `#${r}${g}${b}`;
  }

  // ── Close dropdown on outside click ──
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent): void {
    if (this.showDownloadMenu()) {
      const dropdownEl = this.downloadDropdownRef()?.nativeElement;
      if (dropdownEl && !dropdownEl.contains(event.target as Node)) {
        this.showDownloadMenu.set(false);
      }
    }
  }

  // ── Download as .txt ──
  async downloadAsTxt(): Promise<void> {
    this.showDownloadMenu.set(false);
    const noteTitle = this.title() || 'nota';
    const plainText = this.stripHtml(this.content());
    const textContent = `${noteTitle}\n${'='.repeat(noteTitle.length)}\n\n${plainText}`;

    await window.electronAPI.fileSaveAs({
      defaultName: `${this.sanitizeFilename(noteTitle)}.txt`,
      filters: [{ name: 'Archivo de texto', extensions: ['txt'] }],
      content: textContent,
    });
  }

  // ── Download as .docx ──
  async downloadAsDocx(): Promise<void> {
    this.showDownloadMenu.set(false);
    const noteTitle = this.title() || 'nota';
    const htmlBody = this.content() || '';

    const docxBlob = this.generateDocxBytes(noteTitle, htmlBody);
    const buffer = Array.from(docxBlob);

    await window.electronAPI.fileSaveAsBuffer({
      defaultName: `${this.sanitizeFilename(noteTitle)}.docx`,
      filters: [{ name: 'Documento Word', extensions: ['docx'] }],
      buffer,
    });
  }

  private sanitizeFilename(name: string): string {
    return name.replace(/[<>:"/\\|?*]/g, '_').trim() || 'nota';
  }

  /**
   * Generate a minimal valid .docx file (ZIP archive) entirely in-memory.
   * A .docx file is a ZIP containing XML parts:
   *   [Content_Types].xml, _rels/.rels, word/document.xml, word/_rels/document.xml.rels
   * We embed the note HTML as Word-compatible OOXML using altChunk (HTML import),
   * but for maximum compatibility we convert HTML to simplified WordprocessingML.
   */
  private generateDocxBytes(title: string, htmlContent: string): Uint8Array {
    // Convert HTML to simplified OOXML paragraphs
    const bodyXml = this.htmlToOoxml(htmlContent);

    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr><w:t>${this.escapeXml(title)}</w:t></w:r></w:p>
    <w:p/>
${bodyXml}
  </w:body>
</w:document>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:pPr><w:spacing w:after="200"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="200" w:after="100"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:pPr><w:ind w:left="720"/></w:pPr>
  </w:style>
</w:styles>`;

    const wordRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

    // Build ZIP manually
    const files: { path: string; content: Uint8Array }[] = [
      { path: '[Content_Types].xml', content: this.encodeUTF8(contentTypes) },
      { path: '_rels/.rels', content: this.encodeUTF8(rels) },
      { path: 'word/document.xml', content: this.encodeUTF8(documentXml) },
      { path: 'word/styles.xml', content: this.encodeUTF8(stylesXml) },
      { path: 'word/_rels/document.xml.rels', content: this.encodeUTF8(wordRels) },
    ];

    return this.createZip(files);
  }

  /** Convert simple HTML to OOXML paragraphs */
  private htmlToOoxml(html: string): string {
    const div = document.createElement('div');
    div.innerHTML = html;
    const paragraphs: string[] = [];

    const processNode = (node: Node): void => {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent ?? '';
        if (text.trim()) {
          paragraphs.push(`    <w:p><w:r><w:t xml:space="preserve">${this.escapeXml(text)}</w:t></w:r></w:p>`);
        }
        return;
      }

      if (node.nodeType !== Node.ELEMENT_NODE) return;
      const el = node as HTMLElement;
      const tag = el.tagName.toLowerCase();

      if (['h1', 'h2', 'h3'].includes(tag)) {
        const level = tag === 'h1' ? 'Heading1' : tag === 'h2' ? 'Heading2' : 'Heading3';
        const runs = this.elementToRuns(el);
        paragraphs.push(`    <w:p><w:pPr><w:pStyle w:val="${level}"/>${this.getAlignXml(el)}</w:pPr>${runs}</w:p>`);
      } else if (tag === 'p' || tag === 'div') {
        const runs = this.elementToRuns(el);
        paragraphs.push(`    <w:p><w:pPr>${this.getAlignXml(el)}</w:pPr>${runs}</w:p>`);
      } else if (tag === 'ul' || tag === 'ol') {
        el.querySelectorAll('li').forEach((li) => {
          const runs = this.elementToRuns(li);
          paragraphs.push(`    <w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>${runs}</w:p>`);
        });
      } else if (tag === 'br') {
        paragraphs.push(`    <w:p/>`);
      } else {
        // Process children for unknown tags
        el.childNodes.forEach((child) => processNode(child));
      }
    };

    div.childNodes.forEach((child) => processNode(child));

    if (paragraphs.length === 0) {
      const text = div.textContent ?? '';
      if (text.trim()) {
        text.split('\n').filter((l) => l.trim()).forEach((line) => {
          paragraphs.push(`    <w:p><w:r><w:t xml:space="preserve">${this.escapeXml(line)}</w:t></w:r></w:p>`);
        });
      }
    }

    return paragraphs.join('\n');
  }

  /** Convert an element's inline content to OOXML runs */
  private elementToRuns(el: HTMLElement): string {
    const runs: string[] = [];

    const walk = (node: Node, inheritBold: boolean, inheritItalic: boolean, inheritUnderline: boolean, inheritStrike: boolean, inheritColor: string | null, inheritSize: string | null): void => {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent ?? '';
        if (!text) return;
        let rPr = '';
        if (inheritBold) rPr += '<w:b/>';
        if (inheritItalic) rPr += '<w:i/>';
        if (inheritUnderline) rPr += '<w:u w:val="single"/>';
        if (inheritStrike) rPr += '<w:strike/>';
        if (inheritColor) {
          const hex = inheritColor.replace('#', '');
          rPr += `<w:color w:val="${hex}"/>`;
        }
        if (inheritSize) {
          rPr += `<w:sz w:val="${inheritSize}"/><w:szCs w:val="${inheritSize}"/>`;
        }
        const rPrXml = rPr ? `<w:rPr>${rPr}</w:rPr>` : '';
        runs.push(`<w:r>${rPrXml}<w:t xml:space="preserve">${this.escapeXml(text)}</w:t></w:r>`);
        return;
      }

      if (node.nodeType !== Node.ELEMENT_NODE) return;
      const child = node as HTMLElement;
      const childTag = child.tagName.toLowerCase();

      let bold = inheritBold || ['b', 'strong'].includes(childTag);
      let italic = inheritItalic || ['i', 'em'].includes(childTag);
      let underline = inheritUnderline || childTag === 'u';
      let strike = inheritStrike || ['s', 'strike', 'del'].includes(childTag);
      let color = inheritColor;
      let size = inheritSize;

      // Check inline styles
      const style = child.style;
      if (style.fontWeight === 'bold' || parseInt(style.fontWeight) >= 700) bold = true;
      if (style.fontStyle === 'italic') italic = true;
      if (style.textDecoration?.includes('underline')) underline = true;
      if (style.textDecoration?.includes('line-through')) strike = true;
      if (style.color) color = this.rgbToHex(style.color);
      if (child.tagName === 'FONT') {
        const fontColor = child.getAttribute('color');
        if (fontColor) color = fontColor;
        const fontSize = child.getAttribute('size');
        if (fontSize) {
          const sizeMap: Record<string, string> = { '1': '16', '2': '22', '3': '24', '4': '28', '5': '36', '6': '48', '7': '72' };
          size = sizeMap[fontSize] ?? size;
        }
      }

      if (childTag === 'br') {
        runs.push('<w:r><w:br/></w:r>');
        return;
      }

      child.childNodes.forEach((n) => walk(n, bold, italic, underline, strike, color, size));
    };

    el.childNodes.forEach((n) => walk(n, false, false, false, false, null, null));
    return runs.join('');
  }

  private getAlignXml(el: HTMLElement): string {
    const align = el.style.textAlign || el.getAttribute('align') || '';
    if (align === 'center') return '<w:jc w:val="center"/>';
    if (align === 'right') return '<w:jc w:val="right"/>';
    if (align === 'justify') return '<w:jc w:val="both"/>';
    return '';
  }

  private escapeXml(str: string): string {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  private encodeUTF8(str: string): Uint8Array {
    return new TextEncoder().encode(str);
  }

  /**
   * Create a minimal ZIP (PKZIP) file from an array of file entries.
   * Implements the bare-minimum ZIP spec: local file headers + central directory + EOCD.
   */
  private createZip(files: { path: string; content: Uint8Array }[]): Uint8Array {
    const localHeaders: Uint8Array[] = [];
    const centralEntries: Uint8Array[] = [];
    let offset = 0;

    for (const file of files) {
      const pathBytes = this.encodeUTF8(file.path);
      const data = file.content;
      const crc = this.crc32(data);

      // Local file header (30 + pathLen + dataLen)
      const localHeader = new Uint8Array(30 + pathBytes.length + data.length);
      const lv = new DataView(localHeader.buffer);
      lv.setUint32(0, 0x04034b50, true);   // signature
      lv.setUint16(4, 20, true);            // version needed
      lv.setUint16(6, 0, true);             // flags
      lv.setUint16(8, 0, true);             // compression: store
      lv.setUint16(10, 0, true);            // mod time
      lv.setUint16(12, 0, true);            // mod date
      lv.setUint32(14, crc, true);          // crc32
      lv.setUint32(18, data.length, true);  // compressed size
      lv.setUint32(22, data.length, true);  // uncompressed size
      lv.setUint16(26, pathBytes.length, true); // filename length
      lv.setUint16(28, 0, true);            // extra field length
      localHeader.set(pathBytes, 30);
      localHeader.set(data, 30 + pathBytes.length);
      localHeaders.push(localHeader);

      // Central directory entry (46 + pathLen)
      const centralEntry = new Uint8Array(46 + pathBytes.length);
      const cv = new DataView(centralEntry.buffer);
      cv.setUint32(0, 0x02014b50, true);    // signature
      cv.setUint16(4, 20, true);            // version made by
      cv.setUint16(6, 20, true);            // version needed
      cv.setUint16(8, 0, true);             // flags
      cv.setUint16(10, 0, true);            // compression
      cv.setUint16(12, 0, true);            // mod time
      cv.setUint16(14, 0, true);            // mod date
      cv.setUint32(16, crc, true);          // crc32
      cv.setUint32(20, data.length, true);  // compressed size
      cv.setUint32(24, data.length, true);  // uncompressed size
      cv.setUint16(28, pathBytes.length, true); // filename length
      cv.setUint16(30, 0, true);            // extra field length
      cv.setUint16(32, 0, true);            // comment length
      cv.setUint16(34, 0, true);            // disk number
      cv.setUint16(36, 0, true);            // internal attrs
      cv.setUint32(38, 0, true);            // external attrs
      cv.setUint32(42, offset, true);       // local header offset
      centralEntry.set(pathBytes, 46);
      centralEntries.push(centralEntry);

      offset += localHeader.length;
    }

    const centralDirOffset = offset;
    let centralDirSize = 0;
    centralEntries.forEach((e) => (centralDirSize += e.length));

    // End of central directory record (22 bytes)
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);             // signature
    ev.setUint16(4, 0, true);                      // disk number
    ev.setUint16(6, 0, true);                      // disk with CD
    ev.setUint16(8, files.length, true);            // entries on disk
    ev.setUint16(10, files.length, true);           // total entries
    ev.setUint32(12, centralDirSize, true);         // CD size
    ev.setUint32(16, centralDirOffset, true);       // CD offset
    ev.setUint16(20, 0, true);                      // comment length

    // Concatenate all parts
    const totalSize = offset + centralDirSize + 22;
    const result = new Uint8Array(totalSize);
    let pos = 0;
    for (const lh of localHeaders) {
      result.set(lh, pos);
      pos += lh.length;
    }
    for (const ce of centralEntries) {
      result.set(ce, pos);
      pos += ce.length;
    }
    result.set(eocd, pos);

    return result;
  }

  /** Compute CRC-32 for a Uint8Array */
  private crc32(data: Uint8Array): number {
    let crc = 0xffffffff;
    for (let i = 0; i < data.length; i++) {
      crc ^= data[i];
      for (let j = 0; j < 8; j++) {
        crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
      }
    }
    return (crc ^ 0xffffffff) >>> 0;
  }
}
