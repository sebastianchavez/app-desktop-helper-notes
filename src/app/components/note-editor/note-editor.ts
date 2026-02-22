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
  OnDestroy,
  output,
} from '@angular/core';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { TranslateModule, TranslateService } from '@ngx-translate/core';
import { NoteService } from '../../services/note.service';
import { DownloadService } from '../../services/download.service';

@Component({
  selector: 'app-note-editor',
  templateUrl: './note-editor.html',
  styleUrl: './note-editor.scss',
  imports: [TranslateModule],
})
export class NoteEditor implements OnDestroy {
  protected readonly noteService = inject(NoteService);
  private readonly sanitizer = inject(DomSanitizer);
  private readonly downloadService = inject(DownloadService);
  private readonly translate = inject(TranslateService);

  readonly aiChatToggle = output<void>();

  protected title = signal('');
  protected content = signal('');
  protected pdfUrl = signal<SafeResourceUrl | null>(null);
  private pdfBlobUrl: string | null = null;

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
  protected wordWrap = signal(false);

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

      // Clean up previous PDF blob URL
      untracked(() => this.clearPdfUrl());

      if (note) {
        this.title.set(note.title);
        this.content.set(note.content);

        if (note.file_type === 'pdf') {
          // Load PDF data asynchronously
          untracked(() => this.loadPdfData(note.id));
        } else {
          // Defer DOM update so Angular renders the @if block first
          const noteContent = note.content || '';
          setTimeout(() => {
            const el = this.editorRef()?.nativeElement;
            if (el) {
              el.innerHTML = noteContent;
            }
          });
        }
      }
    });
  }

  ngOnDestroy(): void {
    this.clearPdfUrl();
  }

  private clearPdfUrl(): void {
    if (this.pdfBlobUrl) {
      URL.revokeObjectURL(this.pdfBlobUrl);
      this.pdfBlobUrl = null;
    }
    this.pdfUrl.set(null);
  }

  private async loadPdfData(noteId: number): Promise<void> {
    try {
      const fullNote = await window.electronAPI.notesGetById(noteId);
      if (!fullNote?.file_data) return;
      // Check the note is still selected (could have changed during async load)
      if (this.noteService.selectedNoteId() !== noteId) return;

      const binaryStr = atob(fullNote.file_data);
      const bytes = new Uint8Array(binaryStr.length);
      for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
      const blob = new Blob([bytes], { type: 'application/pdf' });
      this.pdfBlobUrl = URL.createObjectURL(blob);
      this.pdfUrl.set(this.sanitizer.bypassSecurityTrustResourceUrl(this.pdfBlobUrl));
    } catch (err) {
      console.error('Error loading PDF data:', err);
    }
  }

  protected async downloadPdf(): Promise<void> {
    const note = this.noteService.selectedNote();
    if (!note) return;
    const fileName = (note.title || 'documento') + '.pdf';
    this.downloadService.showSaving(fileName);
    try {
      const fullNote = await window.electronAPI.notesGetById(note.id);
      if (!fullNote?.file_data) { this.downloadService.showError(fileName); return; }
      const binaryStr = atob(fullNote.file_data);
      const buffer = Array.from(binaryStr, (c) => c.charCodeAt(0));
      const result = await window.electronAPI.fileSaveAsBuffer({
        defaultName: fileName,
        filters: [{ name: 'PDF', extensions: ['pdf'] }],
        buffer,
      });
      if (result.success && result.filePath) {
        this.downloadService.showDone(fileName, result.filePath);
      } else {
        this.downloadService.dismiss(fileName);
      }
    } catch { this.downloadService.showError(fileName); }
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
    return date.toLocaleDateString(this.translate.currentLang || 'es', {
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

  // ── Image support ──

  /** Handle paste: intercept images from clipboard */
  onEditorPaste(event: ClipboardEvent): void {
    const items = event.clipboardData?.items;
    if (!items) return;

    for (let i = 0; i < items.length; i++) {
      if (items[i].type.startsWith('image/')) {
        event.preventDefault();
        const file = items[i].getAsFile();
        if (file) this.insertImageFile(file);
        return;
      }
    }
  }

  /** Handle drop: intercept image files dropped into the editor */
  onEditorDrop(event: DragEvent): void {
    const files = event.dataTransfer?.files;
    if (!files) return;

    for (let i = 0; i < files.length; i++) {
      if (files[i].type.startsWith('image/')) {
        event.preventDefault();
        this.insertImageFile(files[i]);
        return;
      }
    }
  }

  /** Open file picker to insert an image */
  insertImageFromFile(): void {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.onchange = () => {
      const file = input.files?.[0];
      if (file) this.insertImageFile(file);
    };
    input.click();
  }

  /** Read an image file as base64 data URL and insert it at cursor */
  private insertImageFile(file: File): void {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result as string;
      this.refocusEditor();
      document.execCommand('insertImage', false, dataUrl);
      // After insertion, style the image
      setTimeout(() => {
        const el = this.editorRef()?.nativeElement;
        if (el) {
          el.querySelectorAll('img:not([data-styled])').forEach((img: Element) => {
            const imgEl = img as HTMLImageElement;
            imgEl.style.maxWidth = '100%';
            imgEl.style.height = 'auto';
            imgEl.style.borderRadius = '8px';
            imgEl.style.margin = '8px 0';
            imgEl.setAttribute('data-styled', '1');
          });
          this.content.set(el.innerHTML);
          this.scheduleSave();
        }
      });
    };
    reader.readAsDataURL(file);
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
    const fileName = `${this.sanitizeFilename(noteTitle)}.txt`;
    this.downloadService.showSaving(fileName);
    try {
      const result = await window.electronAPI.fileSaveAs({
        defaultName: fileName,
        filters: [{ name: this.translate.instant('NOTE_EDITOR.FILTER_TEXT'), extensions: ['txt'] }],
        content: textContent,
      });
      if (result.success && result.filePath) {
        this.downloadService.showDone(fileName, result.filePath);
      } else {
        this.downloadService.dismiss(fileName);
      }
    } catch { this.downloadService.showError(fileName); }
  }

  // ── Download as .html ──
  async downloadAsHtml(): Promise<void> {
    this.showDownloadMenu.set(false);
    const noteTitle = this.title() || 'nota';
    const htmlBody = this.content() || '';

    // If content already is a full HTML document, export it as-is
    const isFullHtml = /<html[\s>]/i.test(htmlBody);
    const fullHtml = isFullHtml
      ? htmlBody
      : `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${this.escapeXml(noteTitle)}</title>
</head>
<body>
  <h1>${this.escapeXml(noteTitle)}</h1>
  ${htmlBody}
</body>
</html>`;
    const fileName = `${this.sanitizeFilename(noteTitle)}.html`;
    this.downloadService.showSaving(fileName);
    try {
      const result = await window.electronAPI.fileSaveAs({
        defaultName: fileName,
        filters: [{ name: this.translate.instant('NOTE_EDITOR.FILTER_HTML'), extensions: ['html', 'htm'] }],
        content: fullHtml,
        encoding: 'utf-8',
      });
      if (result.success && result.filePath) {
        this.downloadService.showDone(fileName, result.filePath);
      } else {
        this.downloadService.dismiss(fileName);
      }
    } catch { this.downloadService.showError(fileName); }
  }

  // ── Download as .docx ──
  async downloadAsDocx(): Promise<void> {
    this.showDownloadMenu.set(false);
    const noteTitle = this.title() || 'nota';
    const htmlBody = this.content() || '';
    const fileName = `${this.sanitizeFilename(noteTitle)}.docx`;
    this.downloadService.showSaving(fileName);
    try {
      const docxBlob = this.generateDocxBytes(noteTitle, htmlBody);
      const buffer = Array.from(docxBlob);
      const result = await window.electronAPI.fileSaveAsBuffer({
        defaultName: fileName,
        filters: [{ name: this.translate.instant('NOTE_EDITOR.FILTER_WORD'), extensions: ['docx'] }],
        buffer,
      });
      if (result.success && result.filePath) {
        this.downloadService.showDone(fileName, result.filePath);
      } else {
        this.downloadService.dismiss(fileName);
      }
    } catch { this.downloadService.showError(fileName); }
  }

  // ── Download as .pdf (note content rendered to PDF) ──
  async downloadAsPdf(): Promise<void> {
    this.showDownloadMenu.set(false);
    const noteTitle = this.title() || 'nota';
    const fileName = `${this.sanitizeFilename(noteTitle)}.pdf`;
    const htmlContent = `<h1>${this.escapeXml(noteTitle)}</h1>${this.content() || ''}`;
    this.downloadService.showSaving(fileName);
    try {
      const result = await window.electronAPI.filePrintToPdf({
        defaultName: fileName,
        htmlContent,
      });
      if (result.success && result.filePath) {
        this.downloadService.showDone(fileName, result.filePath);
      } else {
        this.downloadService.dismiss(fileName);
      }
    } catch { this.downloadService.showError(fileName); }
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
