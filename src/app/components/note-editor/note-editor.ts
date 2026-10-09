import {
  Component,
  inject,
  signal,
  effect,
  computed,
  ElementRef,
  viewChild,
  viewChildren,
  HostListener,
  untracked,
  OnDestroy,
  AfterViewChecked,
  output,
  QueryList,
} from '@angular/core';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { TranslateModule, TranslateService } from '@ngx-translate/core';
import { DecimalPipe, NgClass, NgStyle } from '@angular/common';
import { NoteService } from '../../services/note.service';
import { DownloadService } from '../../services/download.service';
import { PresentationExportService } from '../../services/presentation-export.service';
import {
  createPresentationSlide,
  parsePresentationDocument,
  PresentationDocument,
  PresentationAnimationType,
  PresentationElementAnimation,
  PresentationLayout,
  PresentationSlide,
  PresentationSlideStyle,
  PresentationTransitionType,
} from '../../models/presentation.model';

type XlsxCellValue = string | number | boolean | null;
type XlsxSheetView = {
  name: string;
  data: XlsxCellValue[][];
  colCount: number;
  formulas: { row: number; col: number; formula: string }[];
  merges?: { s: { r: number; c: number }; e: { r: number; c: number } }[];
  colWidths?: number[];
  rowHeights?: number[];
};

type FormulaSuggestion = { name: string; syntax: string; description: string };

/**
 * Evalúa el subconjunto aritmético seguro de las fórmulas sin usar `eval` ni
 * `Function`. El parser descendente respeta paréntesis, signos, potencias,
 * porcentajes y precedencia de operadores.
 */
export function evaluateArithmeticExpression(
  expression: string,
  resolveCell: (reference: string) => number | null = () => null,
): number | string {
  let position = 0;
  const skipWhitespace = (): void => {
    while (/\s/.test(expression[position] ?? '')) position++;
  };
  const fail = (message = '#ERROR!'): never => { throw new Error(message); };
  const parsePrimary = (): number => {
    skipWhitespace();
    if (expression[position] === '(') {
      position++;
      const value = parseExpression();
      skipWhitespace();
      if (expression[position] !== ')') fail();
      position++;
      return value;
    }
    const rest = expression.slice(position);
    const cell = rest.match(/^[A-Za-z]+\d+/)?.[0];
    if (cell) {
      position += cell.length;
      return resolveCell(cell) ?? 0;
    }
    const numberMatch = rest.match(/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+\-]?\d+)?/);
    if (!numberMatch) return fail();
    const numericLiteral = numberMatch[0];
    position += numericLiteral.length;
    return Number(numericLiteral);
  };
  const parseUnary = (): number => {
    skipWhitespace();
    if (expression[position] === '+') { position++; return parseUnary(); }
    if (expression[position] === '-') { position++; return -parseUnary(); }
    let value = parsePrimary();
    skipWhitespace();
    if (expression[position] === '%') { position++; value /= 100; }
    return value;
  };
  const parsePower = (): number => {
    let value = parseUnary();
    skipWhitespace();
    if (expression[position] === '^') { position++; value = Math.pow(value, parsePower()); }
    return value;
  };
  const parseTerm = (): number => {
    let value = parsePower();
    while (true) {
      skipWhitespace();
      const operator = expression[position];
      if (operator !== '*' && operator !== '/') break;
      position++;
      const right = parsePower();
      if (operator === '/' && right === 0) fail('#DIV/0!');
      value = operator === '*' ? value * right : value / right;
    }
    return value;
  };
  const parseExpression = (): number => {
    let value = parseTerm();
    while (true) {
      skipWhitespace();
      const operator = expression[position];
      if (operator !== '+' && operator !== '-') break;
      position++;
      const right = parseTerm();
      value = operator === '+' ? value + right : value - right;
    }
    return value;
  };
  try {
    const value = parseExpression();
    skipWhitespace();
    return position === expression.length && Number.isFinite(value) ? value : '#ERROR!';
  } catch (error) {
    return error instanceof Error && error.message === '#DIV/0!' ? '#DIV/0!' : '#ERROR!';
  }
}

/**
 * Calcula la celda vecina para navegación con teclado y limita el resultado a
 * los bordes actuales de la hoja.
 */
export function nextSpreadsheetCell(
  cell: { row: number; col: number },
  direction: 'up' | 'down' | 'left' | 'right',
  rowCount: number,
  colCount: number,
): { row: number; col: number } {
  const lastRow = Math.max(0, rowCount - 1);
  const lastCol = Math.max(0, colCount - 1);
  switch (direction) {
    case 'up': return { row: Math.max(0, cell.row - 1), col: cell.col };
    case 'down': return { row: Math.min(lastRow, cell.row + 1), col: cell.col };
    case 'left': return { row: cell.row, col: Math.max(0, cell.col - 1) };
    case 'right': return { row: cell.row, col: Math.min(lastCol, cell.col + 1) };
  }
}

@Component({
  selector: 'app-note-editor',
  templateUrl: './note-editor.html',
  styleUrl: './note-editor.scss',
  imports: [TranslateModule, NgClass, NgStyle, DecimalPipe],
})
/**
 * Editor multimodal de Helper Notes.
 *
 * Selecciona el modo según `file_type`: texto enriquecido, PDF, hoja XLSX o
 * presentación. Centraliza autosave, exportaciones y estado de herramientas;
 * las operaciones privilegiadas sobre archivos se delegan a Electron.
 *
 * La sección XLSX conserva una representación visual en señales y entrega cada
 * mutación al proceso principal para reconstruir el libro real. Sus pilas de
 * undo/redo guardan snapshots Base64, mientras que fórmulas y formatos mantienen
 * metadatos auxiliares por coordenada de celda.
 */
export class NoteEditor implements OnDestroy, AfterViewChecked {
  protected readonly noteService = inject(NoteService);
  private readonly sanitizer = inject(DomSanitizer);
  private readonly downloadService = inject(DownloadService);
  private readonly presentationExport = inject(PresentationExportService);
  private readonly translate = inject(TranslateService);
  protected readonly Math = Math;

  readonly aiChatToggle = output<void>();

  protected title = signal('');
  protected content = signal('');
  protected pdfUrl = signal<SafeResourceUrl | null>(null);
  private pdfBlobUrl: string | null = null;

  // Presentation editor state
  protected presentationSlides = signal<PresentationSlide[]>([]);
  protected activePresentationSlideIndex = signal(0);
  protected presentationMode = signal(false);
  protected presentedSlideIndex = signal(0);
  protected presentationExporting = signal(false);
  protected printing = signal(false);
  protected activePresentationRibbonTab = signal<'home' | 'insert' | 'design' | 'transitions' | 'animations' | 'slideshow'>('home');
  protected activePresentationTextTarget = signal<'title' | 'body'>('title');
  protected activePresentationAnimationTarget = signal<'title' | 'body' | 'both'>('title');
  protected presentationEffectsPreviewing = signal(false);
  private presentationPreviewTimer: ReturnType<typeof setTimeout> | null = null;
  protected readonly presentationThemes = [
    { name: 'Naranja', backgroundColor: '#ffffff', accentColor: '#9a3412', textColor: '#334155' },
    { name: 'Azul', backgroundColor: '#eff6ff', accentColor: '#1d4ed8', textColor: '#1e293b' },
    { name: 'Verde', backgroundColor: '#f0fdf4', accentColor: '#15803d', textColor: '#1f2937' },
    { name: 'Oscuro', backgroundColor: '#0f172a', accentColor: '#fb923c', textColor: '#e2e8f0' },
    { name: 'Minimalista', backgroundColor: '#fafafa', accentColor: '#111827', textColor: '#374151' },
  ];
  protected readonly activePresentationSlide = computed(() =>
    this.presentationSlides()[this.activePresentationSlideIndex()] ?? null,
  );
  protected readonly presentedSlide = computed(() =>
    this.presentationSlides()[this.presentedSlideIndex()] ?? null,
  );

  // Excel viewer state signals
  protected xlsxSheets = signal<XlsxSheetView[]>([]);
  protected activeSheetIndex = signal(0);
  protected editingCell = signal<{ row: number; col: number } | null>(null);
  protected formulaBarValue = signal('');
  protected formulaSuggestions = signal<FormulaSuggestion[]>([]);
  protected activeFormulaSuggestion = signal(0);
  protected selectedCell = signal<{ row: number; col: number } | null>(null);
  private xlsxBase64 = '';
  protected xlsxLoading = signal(false);
  private pendingFocusCellInput = false;
  private xlsxResize: { type: 'column' | 'row'; index: number; start: number; initialSize: number } | null = null;

  private readonly formulaCatalog: FormulaSuggestion[] = [
    { name: 'SUMA', syntax: 'SUMA(número1; número2; …)', description: 'Suma números o rangos.' },
    { name: 'PROMEDIO', syntax: 'PROMEDIO(número1; número2; …)', description: 'Calcula el promedio.' },
    { name: 'CONTAR', syntax: 'CONTAR(valor1; valor2; …)', description: 'Cuenta valores numéricos.' },
    { name: 'MIN', syntax: 'MIN(número1; número2; …)', description: 'Devuelve el valor mínimo.' },
    { name: 'MAX', syntax: 'MAX(número1; número2; …)', description: 'Devuelve el valor máximo.' },
    { name: 'SI', syntax: 'SI(condición; valor_si_verdadero; valor_si_falso)', description: 'Evalúa una condición.' },
    { name: 'CONTAR.SI', syntax: 'CONTAR.SI(rango; criterio)', description: 'Cuenta celdas que cumplen un criterio.' },
    { name: 'SUMAR.SI', syntax: 'SUMAR.SI(rango; criterio; rango_suma)', description: 'Suma celdas que cumplen un criterio.' },
    { name: 'REDONDEAR', syntax: 'REDONDEAR(número; decimales)', description: 'Redondea un número.' },
    { name: 'ABS', syntax: 'ABS(número)', description: 'Devuelve el valor absoluto.' },
  ];

  // Range selection state
  protected xlsxSelectionStart = signal<{ row: number; col: number } | null>(null);
  protected xlsxSelectionEnd = signal<{ row: number; col: number } | null>(null);
  protected xlsxSelectingFullRows = signal(false);
  protected xlsxSelectingFullCols = signal(false);
  private xlsxIsDragging = false;

  // Clipboard
  private xlsxClipboard: { data: (string | number | boolean | null)[][]; startRow: number; startCol: number; isCut: boolean } | null = null;
  protected xlsxHasClipboard = signal(false);

  // Context menu
  protected xlsxContextMenu = signal<{ x: number; y: number } | null>(null);

  // Undo / Redo stacks (store base64 snapshots)
  private xlsxUndoStack: string[] = [];
  private xlsxRedoStack: string[] = [];
  private readonly MAX_UNDO = 50;

  // Excel formatting toolbar signals
  protected xlsxFontSize = signal(12);
  protected xlsxBold = signal(false);
  protected xlsxItalic = signal(false);
  protected xlsxUnderline = signal(false);
  protected xlsxAlign = signal<'left' | 'center' | 'right'>('left');

  // Cell formatting storage: key = "sheetIdx:row:col" => format object
  private xlsxCellFormats = new Map<string, { bold?: boolean; italic?: boolean; underline?: boolean; align?: 'left' | 'center' | 'right'; fontSize?: number }>();

  // Formula storage: key = "sheetIdx:row:col" => formula string (e.g. "=SUM(A1:A10)")
  private xlsxCellFormulas = new Map<string, string>();

  // Formula range-selection mode: when typing a formula and user clicks cells to select ranges
  private xlsxFormulaRangeMode = false;
  private xlsxFormulaRangeDragging = false;
  private xlsxFormulaRangeAnchor: { row: number; col: number } | null = null;
  private xlsxFormulaRangeCurrent: { row: number; col: number } | null = null;
  // Position in formula text where the range reference should be inserted/replaced
  private xlsxFormulaInsertPos = -1;
  private xlsxFormulaInsertEnd = -1;

  // Formula range highlights: colored rectangles showing referenced ranges during editing
  private readonly FORMULA_HIGHLIGHT_COLORS = [
    '#4285f4', '#ea4335', '#9c27b0', '#ff9800', '#0097a7', '#43a047'
  ];
  protected xlsxFormulaHighlights = signal<{ startRow: number; startCol: number; endRow: number; endCol: number; color: string }[]>([]);

  // Minimum visible rows/cols to fill viewport like Excel
  private readonly MIN_VISIBLE_ROWS = 50;
  private readonly MIN_VISIBLE_COLS = 26; // A-Z

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
  private xlsxGridRef = viewChild<ElementRef<HTMLDivElement>>('xlsxGrid');
  private cellInputRefs = viewChildren<ElementRef<HTMLInputElement>>('cellInput');
  private pendingSave: {
    timer: ReturnType<typeof setTimeout>;
    noteId: number;
    title: string;
    content: string;
  } | null = null;
  private skipNextSync = false;

  constructor() {
    // Sync local state when selected note changes
    effect(() => {
      const note = this.noteService.selectedNote();
      const pending = untracked(() => this.pendingSave);
      if (pending && pending.noteId !== note?.id) {
        untracked(() => void this.flushPendingSave());
      }
      const skip = untracked(() => this.skipNextSync);
      if (skip) {
        this.skipNextSync = false;
        return;
      }

      // Clean up previous PDF blob URL
      untracked(() => this.clearPdfUrl());
      if (untracked(() => this.presentationMode())) {
        void window.electronAPI.windowSetFullScreen(false);
      }
      this.presentationMode.set(false);

      if (note) {
        this.title.set(note.title);
        this.content.set(note.content);

        if (note.file_type === 'pdf') {
          // Load PDF data asynchronously
          untracked(() => this.loadPdfData(note.id));
        } else if (note.file_type === 'xlsx') {
          // Load Excel data asynchronously
          untracked(() => this.loadXlsxData(note.id));
        } else if (note.file_type === 'pptx') {
          this.loadPresentationContent(note.content);
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
    void this.flushPendingSave();
    this.clearPdfUrl();
    if (this.presentationMode()) void window.electronAPI.windowSetFullScreen(false);
    if (this.presentationPreviewTimer) clearTimeout(this.presentationPreviewTimer);
  }

  ngAfterViewChecked(): void {
    if (this.pendingFocusCellInput) {
      const inputs = this.cellInputRefs();
      if (inputs.length > 0) {
        const input = inputs[0].nativeElement;
        input.focus();
        // Place cursor at end
        const len = input.value.length;
        input.setSelectionRange(len, len);
        this.pendingFocusCellInput = false;
      }
    }
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

  // ── Presentation editor ──

  private loadPresentationContent(content: string): void {
    const document = parsePresentationDocument(content);
    this.presentationSlides.set(document.slides);
    this.activePresentationSlideIndex.set(0);
    this.presentedSlideIndex.set(0);
  }

  private savePresentationContent(): void {
    const document: PresentationDocument = {
      kind: 'helper-notes-presentation',
      version: 1,
      slides: this.presentationSlides(),
    };
    this.content.set(JSON.stringify(document));
    this.scheduleSave();
  }

  protected selectPresentationSlide(index: number): void {
    if (index < 0 || index >= this.presentationSlides().length) return;
    this.activePresentationSlideIndex.set(index);
  }

  protected updatePresentationSlide(field: 'title' | 'body' | 'notes', value: string): void {
    const index = this.activePresentationSlideIndex();
    const slides = this.presentationSlides().map((slide, slideIndex) =>
      slideIndex === index ? { ...slide, [field]: value } : slide,
    );
    this.presentationSlides.set(slides);
    this.savePresentationContent();
  }

  protected updatePresentationSlideStyle(changes: Partial<PresentationSlideStyle>): void {
    const index = this.activePresentationSlideIndex();
    const slides = this.presentationSlides().map((slide, slideIndex) =>
      slideIndex === index ? { ...slide, style: { ...slide.style, ...changes } } : slide,
    );
    this.presentationSlides.set(slides);
    this.savePresentationContent();
  }

  protected setPresentationLayout(layout: PresentationLayout): void {
    const index = this.activePresentationSlideIndex();
    this.presentationSlides.set(this.presentationSlides().map((slide, slideIndex) =>
      slideIndex === index ? { ...slide, layout } : slide,
    ));
    this.savePresentationContent();
  }

  protected applyPresentationTheme(theme: { backgroundColor: string; accentColor: string; textColor: string }): void {
    this.presentationSlides.set(this.presentationSlides().map(slide => ({
      ...slide,
      style: { ...slide.style, ...theme },
    })));
    this.savePresentationContent();
  }

  protected setPresentationTransition(type: PresentationTransitionType): void {
    const index = this.activePresentationSlideIndex();
    this.presentationSlides.set(this.presentationSlides().map((slide, slideIndex) =>
      slideIndex === index ? { ...slide, transition: { ...slide.transition, type } } : slide,
    ));
    this.savePresentationContent();
    this.previewPresentationEffects();
  }

  protected setPresentationTransitionDuration(duration: number): void {
    const index = this.activePresentationSlideIndex();
    const safeDuration = Math.max(0.1, Math.min(5, duration || 0.45));
    this.presentationSlides.set(this.presentationSlides().map((slide, slideIndex) =>
      slideIndex === index ? { ...slide, transition: { ...slide.transition, duration: safeDuration } } : slide,
    ));
    this.savePresentationContent();
  }

  protected applyPresentationTransitionToAll(): void {
    const transition = this.activePresentationSlide()?.transition;
    if (!transition) return;
    this.presentationSlides.set(this.presentationSlides().map(slide => ({
      ...slide,
      transition: { ...transition },
    })));
    this.savePresentationContent();
  }

  protected setPresentationAnimation(type: PresentationAnimationType): void {
    const index = this.activePresentationSlideIndex();
    const target = this.activePresentationAnimationTarget();
    this.presentationSlides.set(this.presentationSlides().map((slide, slideIndex) => {
      if (slideIndex !== index) return slide;
      return {
        ...slide,
        animations: {
          title: target === 'title' || target === 'both' ? { ...slide.animations.title, type } : slide.animations.title,
          body: target === 'body' || target === 'both' ? { ...slide.animations.body, type } : slide.animations.body,
        },
      };
    }));
    this.savePresentationContent();
    this.previewPresentationEffects();
  }

  protected isPresentationAnimationSelected(slide: PresentationSlide, type: PresentationAnimationType): boolean {
    const target = this.activePresentationAnimationTarget();
    if (target === 'both') return slide.animations.title.type === type && slide.animations.body.type === type;
    return slide.animations[target].type === type;
  }

  protected updatePresentationAnimationTiming(field: 'duration' | 'delay', value: number): void {
    const index = this.activePresentationSlideIndex();
    const target = this.activePresentationAnimationTarget();
    const min = field === 'duration' ? 0.1 : 0;
    const max = field === 'duration' ? 5 : 10;
    const fallback = field === 'duration' ? 0.6 : 0;
    const safeValue = Math.max(min, Math.min(max, Number(value) || fallback));
    const update = (animation: PresentationElementAnimation): PresentationElementAnimation => ({ ...animation, [field]: safeValue });
    this.presentationSlides.set(this.presentationSlides().map((slide, slideIndex) => {
      if (slideIndex !== index) return slide;
      return {
        ...slide,
        animations: {
          title: target === 'title' || target === 'both' ? update(slide.animations.title) : slide.animations.title,
          body: target === 'body' || target === 'both' ? update(slide.animations.body) : slide.animations.body,
        },
      };
    }));
    this.savePresentationContent();
  }

  protected previewPresentationEffects(): void {
    if (this.presentationPreviewTimer) clearTimeout(this.presentationPreviewTimer);
    this.presentationEffectsPreviewing.set(false);
    setTimeout(() => this.presentationEffectsPreviewing.set(true));
    const slide = this.activePresentationSlide();
    const totalSeconds = Math.max(
      slide?.transition.duration ?? 0,
      (slide?.animations.title.delay ?? 0) + (slide?.animations.title.duration ?? 0),
      (slide?.animations.body.delay ?? 0) + (slide?.animations.body.duration ?? 0),
    );
    this.presentationPreviewTimer = setTimeout(() => this.presentationEffectsPreviewing.set(false), (totalSeconds + 0.25) * 1000);
  }

  protected presentationTransitionClass(slide: PresentationSlide, preview = false): string {
    return preview && !this.presentationEffectsPreviewing() ? '' : `ppt-transition-${slide.transition.type}`;
  }

  protected presentationAnimationClass(animation: PresentationElementAnimation, preview = false): string {
    return preview && !this.presentationEffectsPreviewing() ? '' : `ppt-animation-${animation.type}`;
  }

  protected presentationEffectTiming(duration: number, delay = 0): Record<string, string> {
    return { 'animation-duration': `${duration}s`, 'animation-delay': `${delay}s` };
  }

  protected togglePresentationTextStyle(property: 'Bold' | 'Italic' | 'Underline'): void {
    const slide = this.activePresentationSlide();
    if (!slide) return;
    const key = `${this.activePresentationTextTarget()}${property}` as keyof PresentationSlideStyle;
    this.updatePresentationSlideStyle({ [key]: !slide.style[key] } as Partial<PresentationSlideStyle>);
  }

  protected setPresentationTextAlign(align: 'left' | 'center' | 'right'): void {
    const key = `${this.activePresentationTextTarget()}Align` as keyof PresentationSlideStyle;
    this.updatePresentationSlideStyle({ [key]: align } as Partial<PresentationSlideStyle>);
  }

  protected setPresentationFontSize(size: number): void {
    const key = `${this.activePresentationTextTarget()}FontSize` as keyof PresentationSlideStyle;
    this.updatePresentationSlideStyle({ [key]: size } as Partial<PresentationSlideStyle>);
  }

  protected insertPresentationBullet(): void {
    const slide = this.activePresentationSlide();
    if (!slide) return;
    const body = slide.body.trim()
      ? slide.body.split('\n').map(line => line.startsWith('• ') ? line : `• ${line}`).join('\n')
      : '• ';
    this.activePresentationTextTarget.set('body');
    this.updatePresentationSlide('body', body);
  }

  protected getPresentationStageStyle(slide: PresentationSlide): Record<string, string> {
    return { 'background-color': slide.style.backgroundColor };
  }

  protected getPresentationTitleStyle(slide: PresentationSlide): Record<string, string> {
    return {
      color: slide.style.accentColor,
      'font-size': `${slide.style.titleFontSize}px`,
      'font-weight': slide.style.titleBold ? '700' : '400',
      'font-style': slide.style.titleItalic ? 'italic' : 'normal',
      'text-decoration': slide.style.titleUnderline ? 'underline' : 'none',
      'text-align': slide.style.titleAlign,
    };
  }

  protected getPresentationBodyStyle(slide: PresentationSlide): Record<string, string> {
    return {
      color: slide.style.textColor,
      'font-size': `${slide.style.bodyFontSize}px`,
      'font-weight': slide.style.bodyBold ? '700' : '400',
      'font-style': slide.style.bodyItalic ? 'italic' : 'normal',
      'text-decoration': slide.style.bodyUnderline ? 'underline' : 'none',
      'text-align': slide.style.bodyAlign,
    };
  }

  protected addPresentationSlide(): void {
    const slides = [...this.presentationSlides()];
    const insertAt = this.activePresentationSlideIndex() + 1;
    slides.splice(insertAt, 0, createPresentationSlide(insertAt));
    this.presentationSlides.set(slides);
    this.activePresentationSlideIndex.set(insertAt);
    this.savePresentationContent();
  }

  protected duplicatePresentationSlide(): void {
    const current = this.activePresentationSlide();
    if (!current) return;
    const insertAt = this.activePresentationSlideIndex() + 1;
    const slides = [...this.presentationSlides()];
    slides.splice(insertAt, 0, { ...current, id: createPresentationSlide(insertAt).id });
    this.presentationSlides.set(slides);
    this.activePresentationSlideIndex.set(insertAt);
    this.savePresentationContent();
  }

  protected deletePresentationSlide(): void {
    const slides = [...this.presentationSlides()];
    if (slides.length <= 1) {
      slides[0] = createPresentationSlide();
      this.activePresentationSlideIndex.set(0);
    } else {
      slides.splice(this.activePresentationSlideIndex(), 1);
      this.activePresentationSlideIndex.set(Math.min(this.activePresentationSlideIndex(), slides.length - 1));
    }
    this.presentationSlides.set(slides);
    this.savePresentationContent();
  }

  protected movePresentationSlide(direction: -1 | 1): void {
    const index = this.activePresentationSlideIndex();
    const target = index + direction;
    const slides = [...this.presentationSlides()];
    if (target < 0 || target >= slides.length) return;
    [slides[index], slides[target]] = [slides[target], slides[index]];
    this.presentationSlides.set(slides);
    this.activePresentationSlideIndex.set(target);
    this.savePresentationContent();
  }

  protected startPresentation(fromCurrent = false): void {
    if (this.presentationSlides().length === 0) return;
    this.presentedSlideIndex.set(fromCurrent ? this.activePresentationSlideIndex() : 0);
    this.presentationMode.set(true);
    void window.electronAPI.windowSetFullScreen(true);
  }

  protected closePresentation(): void {
    this.presentationMode.set(false);
    void window.electronAPI.windowSetFullScreen(false);
  }

  protected changePresentedSlide(direction: -1 | 1): void {
    const next = Math.max(0, Math.min(this.presentationSlides().length - 1, this.presentedSlideIndex() + direction));
    this.presentedSlideIndex.set(next);
  }

  @HostListener('document:keydown', ['$event'])
  onPresentationKeydown(event: KeyboardEvent): void {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'p') {
      event.preventDefault();
      void this.printCurrentDocument();
      return;
    }
    if (!this.presentationMode()) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      this.closePresentation();
    } else if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'].includes(event.key)) {
      event.preventDefault();
      this.changePresentedSlide(1);
    } else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(event.key)) {
      event.preventDefault();
      this.changePresentedSlide(-1);
    } else if (event.key === 'Home') {
      event.preventDefault();
      this.presentedSlideIndex.set(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      this.presentedSlideIndex.set(this.presentationSlides().length - 1);
    }
  }

  protected async downloadPresentationFile(): Promise<void> {
    if (this.presentationExporting()) return;
    const fileName = `${this.sanitizeFilename(this.title() || 'presentacion')}.pptx`;
    this.presentationExporting.set(true);
    this.downloadService.showSaving(fileName);
    try {
      const bytes = await this.presentationExport.generateSlides(this.title(), this.presentationSlides());
      const result = await window.electronAPI.fileSaveAsBuffer({
        defaultName: fileName,
        filters: [{ name: this.translate.instant('NOTE_EDITOR.FILTER_POWERPOINT'), extensions: ['pptx'] }],
        buffer: Array.from(bytes),
      });
      if (result.success && result.filePath) this.downloadService.showDone(fileName, result.filePath);
      else this.downloadService.dismiss(fileName);
    } catch {
      this.downloadService.showError(fileName);
    } finally {
      this.presentationExporting.set(false);
    }
  }

  // ── Excel (xlsx) viewer methods ──

  private async loadXlsxData(noteId: number): Promise<void> {
    this.xlsxLoading.set(true);
    this.xlsxSheets.set([]);
    this.activeSheetIndex.set(0);
    this.editingCell.set(null);
    this.selectedCell.set(null);
    this.formulaBarValue.set('');
    this.xlsxSelectionStart.set(null);
    this.xlsxSelectionEnd.set(null);
    this.xlsxSelectingFullRows.set(false);
    this.xlsxSelectingFullCols.set(false);
    this.xlsxContextMenu.set(null);
    this.xlsxCellFormats.clear();
    this.xlsxCellFormulas.clear();
    this.xlsxUndoStack = [];
    this.xlsxRedoStack = [];
    try {
      const fullNote = await window.electronAPI.notesGetById(noteId);
      if (!fullNote?.file_data) return;
      if (this.noteService.selectedNoteId() !== noteId) return;
      this.xlsxBase64 = fullNote.file_data;
      const result = await window.electronAPI.xlsxParseData(this.xlsxBase64);
      if (this.noteService.selectedNoteId() !== noteId) return;
      if (result.success && result.sheets) {
        this.loadXlsxFormulaMetadata(result.sheets);
        this.xlsxSheets.set(result.sheets);
        this.recalculateXlsxFormulas();
      }
    } catch (err) {
      console.error('Error loading xlsx data:', err);
    } finally {
      this.xlsxLoading.set(false);
    }
  }

  protected selectSheet(index: number): void {
    this.activeSheetIndex.set(index);
    this.editingCell.set(null);
    this.selectedCell.set(null);
    this.formulaBarValue.set('');
    this.xlsxSelectionStart.set(null);
    this.xlsxSelectionEnd.set(null);
    this.xlsxSelectingFullRows.set(false);
    this.xlsxSelectingFullCols.set(false);
    this.xlsxContextMenu.set(null);
    setTimeout(() => this.recalculateXlsxFormulas());
  }

  protected get activeSheet() {
    const sheets = this.xlsxSheets();
    const idx = this.activeSheetIndex();
    return sheets[idx] ?? null;
  }

  /** Total number of columns to display (data cols + extra to fill viewport) */
  protected get displayColCount(): number {
    const sheet = this.activeSheet;
    const dataCols = sheet ? sheet.colCount : 0;
    return Math.max(dataCols, this.MIN_VISIBLE_COLS);
  }

  /** Total number of rows to display */
  protected get displayRowCount(): number {
    const sheet = this.activeSheet;
    const dataRows = sheet ? sheet.data.length : 0;
    return Math.max(dataRows, this.MIN_VISIBLE_ROWS);
  }

  /** Get cell value for display (handles rows/cols beyond data range) */
  protected getCellValue(row: number, col: number): string | number | boolean | null {
    const sheet = this.activeSheet;
    if (!sheet) return null;
    return sheet.data[row]?.[col] ?? null;
  }

  /** Check if a cell value is numeric for right-alignment */
  protected isCellNumeric(row: number, col: number): boolean {
    const val = this.getCellValue(row, col);
    return typeof val === 'number';
  }

  /** Column indices array for template iteration */
  protected get columnIndices(): number[] {
    return Array.from({ length: this.displayColCount }, (_, i) => i);
  }

  /** Row indices array for template iteration */
  protected get rowIndices(): number[] {
    return Array.from({ length: this.displayRowCount }, (_, i) => i);
  }

  protected getColumnWidth(index: number): number {
    return this.activeSheet?.colWidths?.[index] ?? 80;
  }

  protected getRowHeight(index: number): number {
    return this.activeSheet?.rowHeights?.[index] ?? 20;
  }

  protected startColumnResize(event: MouseEvent, index: number): void {
    event.preventDefault();
    event.stopPropagation();
    this.pushUndoSnapshot();
    this.xlsxResize = { type: 'column', index, start: event.clientX, initialSize: this.getColumnWidth(index) };
  }

  protected startRowResize(event: MouseEvent, index: number): void {
    event.preventDefault();
    event.stopPropagation();
    this.pushUndoSnapshot();
    this.xlsxResize = { type: 'row', index, start: event.clientY, initialSize: this.getRowHeight(index) };
  }

  @HostListener('document:mousemove', ['$event'])
  protected onSpreadsheetResizeMove(event: MouseEvent): void {
    const resize = this.xlsxResize;
    const sheet = this.activeSheet;
    if (!resize || !sheet) return;
    event.preventDefault();
    if (resize.type === 'column') {
      const widths = [...(sheet.colWidths ?? [])];
      widths[resize.index] = Math.max(36, Math.min(600, resize.initialSize + event.clientX - resize.start));
      sheet.colWidths = widths;
    } else {
      const heights = [...(sheet.rowHeights ?? [])];
      heights[resize.index] = Math.max(18, Math.min(300, resize.initialSize + event.clientY - resize.start));
      sheet.rowHeights = heights;
    }
    this.xlsxSheets.set([...this.xlsxSheets()]);
  }

  private async persistXlsxDimensions(): Promise<void> {
    const sheet = this.activeSheet;
    if (!sheet) return;
    try {
      const result = await window.electronAPI.xlsxUpdateDimensions(
        this.xlsxBase64,
        this.activeSheetIndex(),
        sheet.colWidths ?? [],
        sheet.rowHeights ?? [],
      );
      if (!result.success || !result.base64) throw new Error(result.error ?? 'No se pudieron guardar las dimensiones');
      this.xlsxBase64 = result.base64;
      const noteId = this.noteService.selectedNoteId();
      if (noteId !== null) {
        this.skipNextSync = true;
        await window.electronAPI.notesUpdate(noteId, { file_data: result.base64 });
      }
    } catch (error) {
      console.error('Error persisting xlsx dimensions:', error);
    }
  }

  protected selectCell(row: number, col: number): void {
    this.selectedCell.set({ row, col });
    this.xlsxSelectionStart.set({ row, col });
    this.xlsxSelectionEnd.set({ row, col });
    this.xlsxSelectingFullRows.set(false);
    this.xlsxSelectingFullCols.set(false);
    this.xlsxContextMenu.set(null);
    // Show formula if cell has one, otherwise show value
    const formula = this.xlsxCellFormulas.get(this.getCellFormatKey(row, col));
    if (formula) {
      this.formulaBarValue.set(formula);
    } else {
      const value = this.getCellValue(row, col);
      this.formulaBarValue.set(value != null ? String(value) : '');
    }
    this.updateXlsxToolbarFromCell(row, col);
  }

  private focusXlsxGridAfterNavigation(): void {
    setTimeout(() => this.xlsxGridRef()?.nativeElement.focus({ preventScroll: true }));
  }

  protected startEditCell(row: number, col: number, initialValue?: string): void {
    this.editingCell.set({ row, col });
    this.selectedCell.set({ row, col });
    // Show formula for editing if the cell has one
    const formula = this.xlsxCellFormulas.get(this.getCellFormatKey(row, col));
    if (initialValue !== undefined) {
      this.formulaBarValue.set(initialValue);
      this.updateFormulaHighlights(initialValue);
      this.updateFormulaSuggestions(initialValue);
    } else if (formula) {
      this.formulaBarValue.set(formula);
      this.updateFormulaHighlights(formula);
    } else {
      const value = this.getCellValue(row, col);
      this.formulaBarValue.set(value != null ? String(value) : '');
      this.xlsxFormulaHighlights.set([]);
      this.formulaSuggestions.set([]);
    }
    this.pendingFocusCellInput = true;
  }

  /**
   * Parse a string as a number, respecting locale-aware thousands separators.
   * E.g. "50.000" → 50000 (dot as thousands sep), "3.558" → 3.558 (dot as decimal),
   * "1.234.567" → 1234567, "18.730" → 18730.
   * Pattern: if the string matches digits separated by dots in groups of 3, treat dots
   * as thousands separators. Otherwise, fall back to standard JS Number parsing.
   */
  private parseSmartNumber(v: string): string | number | null {
    const trimmed = v.trim();
    if (trimmed === '') return null;

    // Check for locale-formatted numbers with dot as thousands separator:
    // e.g. "50.000", "1.234.567", "252.463"
    // Pattern: 1-3 digits followed by one or more groups of .XXX (exactly 3 digits), no other chars
    const thousandsDotRegex = /^\d{1,3}(\.\d{3})+$/;
    if (thousandsDotRegex.test(trimmed)) {
      // Dots are thousands separators — remove them
      const cleaned = trimmed.replace(/\./g, '');
      return Number(cleaned);
    }

    // Check for locale-formatted with dot thousands + comma decimal:
    // e.g. "1.234,56" → 1234.56
    const thousandsDotCommaDecRegex = /^\d{1,3}(\.\d{3})+(,\d+)$/;
    if (thousandsDotCommaDecRegex.test(trimmed)) {
      const cleaned = trimmed.replace(/\./g, '').replace(',', '.');
      const n = Number(cleaned);
      if (!isNaN(n)) return n;
    }

    // Standard JS number parsing (dot is decimal separator)
    const n = Number(trimmed);
    if (!isNaN(n)) return n;

    // Not a number — return as string
    return trimmed;
  }

  protected async commitCellEdit(fromBlur = false): Promise<void> {
    // A cell click while composing a formula temporarily moves focus so a range can be inserted.
    if (fromBlur && this.xlsxFormulaRangeDragging) return;
    // Once range capture finishes, a real focus change must still commit the edit.
    if (fromBlur && this.xlsxFormulaRangeMode) {
      setTimeout(() => {
        if (!this.xlsxFormulaRangeDragging && this.editingCell()) this.commitCellEdit();
      });
      return;
    }
    const cell = this.editingCell();
    if (!cell) return;
    this.pushUndoSnapshot();
    const rawValue = this.formulaBarValue();
    this.editingCell.set(null);
    this.formulaSuggestions.set([]);
    this.clearFormulaRangeMode();
    this.xlsxFormulaHighlights.set([]);

    const formulaKey = this.getCellFormatKey(cell.row, cell.col);
    let numVal: XlsxCellValue;

    // Check if this is a formula
    if (typeof rawValue === 'string' && rawValue.startsWith('=')) {
      this.xlsxCellFormulas.set(formulaKey, this.canonicalizeFormula(rawValue));
      const computed = this.evaluateFormula(rawValue);
      numVal = computed;
    } else {
      // Remove any previous formula
      this.xlsxCellFormulas.delete(formulaKey);
      const parsed = this.parseSmartNumber(rawValue);
      numVal = parsed !== null ? parsed : rawValue || null;
    }

    // Update local data immediately for responsiveness
    const sheets = this.xlsxSheets();
    const sheetIdx = this.activeSheetIndex();
    const sheet = sheets[sheetIdx];
    if (sheet) {
      while (sheet.data.length <= cell.row) sheet.data.push([]);
      while (sheet.data[cell.row].length <= cell.col) sheet.data[cell.row].push(null);
      sheet.data[cell.row][cell.col] = numVal;
      this.xlsxSheets.set([...sheets]);
    }

    this.recalculateXlsxFormulas();

    // Update formula bar to show the formula, not the computed value
    const formula = this.xlsxCellFormulas.get(formulaKey);
    if (formula) {
      this.formulaBarValue.set(formula);
    } else {
      this.formulaBarValue.set(numVal != null ? String(numVal) : '');
    }

    // Persist to backend
    try {
      const updates = new Map<string, { row: number; col: number; value: XlsxCellValue; formula?: string | null }>();
      updates.set(`${cell.row}:${cell.col}`, {
        row: cell.row,
        col: cell.col,
        value: this.getCellValue(cell.row, cell.col),
        formula: this.xlsxCellFormulas.get(formulaKey) ?? null,
      });
      for (const formulaCell of this.activeFormulaCells()) {
        updates.set(`${formulaCell.row}:${formulaCell.col}`, formulaCell);
      }
      const result = await window.electronAPI.xlsxBatchUpdateCells(this.xlsxBase64, sheetIdx, [...updates.values()]);
      if (result.success && result.base64) {
        this.xlsxBase64 = result.base64;
        // Save updated file_data to note
        const noteId = this.noteService.selectedNoteId();
        if (noteId !== null) {
          this.skipNextSync = true;
          await window.electronAPI.notesUpdate(noteId, { file_data: result.base64 });
        }
      }
    } catch (err) {
      console.error('Error updating cell:', err);
    }
  }

  protected onCellKeydown(event: KeyboardEvent): void {
    // The input lives inside the grid. Once an edit is committed, editingCell becomes
    // null synchronously; without stopping propagation the grid handles the same key
    // again and Enter/arrow navigation advances a second time.
    event.stopPropagation();
    if (this.handleFormulaSuggestionKey(event)) return;
    const cell = this.editingCell() ?? this.selectedCell();
    if (!cell) return;
    if (event.key === 'Enter') {
      event.preventDefault();
      void this.commitCellEdit();
      const next = nextSpreadsheetCell(cell, event.shiftKey ? 'up' : 'down', this.displayRowCount, this.displayColCount);
      this.selectCell(next.row, next.col);
      this.focusXlsxGridAfterNavigation();
    } else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) {
      event.preventDefault();
      void this.commitCellEdit();
      const direction = ({
        ArrowUp: 'up',
        ArrowDown: 'down',
        ArrowLeft: 'left',
        ArrowRight: 'right',
      } as const)[event.key as 'ArrowUp' | 'ArrowDown' | 'ArrowLeft' | 'ArrowRight'];
      const next = nextSpreadsheetCell(cell, direction, this.displayRowCount, this.displayColCount);
      this.selectCell(next.row, next.col);
      this.focusXlsxGridAfterNavigation();
    } else if (event.key === 'Escape') {
      this.editingCell.set(null);
      this.clearFormulaRangeMode();
      this.xlsxFormulaHighlights.set([]);
    } else if (event.key === 'Tab') {
      event.preventDefault();
      void this.commitCellEdit();
      const next = nextSpreadsheetCell(cell, event.shiftKey ? 'left' : 'right', this.displayRowCount, this.displayColCount);
      this.selectCell(next.row, next.col);
      this.focusXlsxGridAfterNavigation();
    }
  }

  /** Called on every input change in cell or formula bar — updates highlights and manages range mode */
  protected onFormulaInput(value: string): void {
    this.formulaBarValue.set(value);
    // Reset range insertion tracking on manual typing
    this.xlsxFormulaInsertPos = -1;
    this.xlsxFormulaInsertEnd = -1;
    if (value.startsWith('=')) {
      this.updateFormulaHighlights(value);
    } else {
      this.xlsxFormulaHighlights.set([]);
    }
    this.updateFormulaSuggestions(value);
  }

  private updateFormulaSuggestions(value: string): void {
    if (!value.startsWith('=')) {
      this.formulaSuggestions.set([]);
      return;
    }
    const token = value.match(/(?:^=|[+\-*/^,(;])\s*([A-Za-zÁÉÍÓÚÑ.]*)$/i)?.[1] ?? null;
    if (token === null) {
      this.formulaSuggestions.set([]);
      return;
    }
    const normalized = token.toLocaleUpperCase();
    const matches = this.formulaCatalog
      .filter(item => item.name.startsWith(normalized))
      .slice(0, 7);
    this.formulaSuggestions.set(matches);
    this.activeFormulaSuggestion.set(0);
  }

  protected acceptFormulaSuggestion(suggestion: FormulaSuggestion): void {
    const value = this.formulaBarValue();
    const match = value.match(/(?:^=|[+\-*/^,(;])\s*([A-Za-zÁÉÍÓÚÑ.]*)$/i);
    if (!match || match.index === undefined) return;
    const tokenStart = value.length - match[1].length;
    const next = `${value.slice(0, tokenStart)}${suggestion.name}(`;
    this.formulaBarValue.set(next);
    this.formulaSuggestions.set([]);
    this.updateFormulaHighlights(next);
    if (!this.editingCell() && this.selectedCell()) this.editingCell.set(this.selectedCell());
    this.pendingFocusCellInput = true;
  }

  private handleFormulaSuggestionKey(event: KeyboardEvent): boolean {
    const suggestions = this.formulaSuggestions();
    if (suggestions.length === 0) return false;
    if (event.altKey && (event.key === 'ArrowDown' || event.key === 'ArrowUp')) {
      event.preventDefault();
      const direction = event.key === 'ArrowDown' ? 1 : -1;
      this.activeFormulaSuggestion.set((this.activeFormulaSuggestion() + direction + suggestions.length) % suggestions.length);
      return true;
    }
    if (event.key === 'Tab' || event.key === 'Enter') {
      event.preventDefault();
      this.acceptFormulaSuggestion(suggestions[this.activeFormulaSuggestion()] ?? suggestions[0]);
      return true;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      this.formulaSuggestions.set([]);
      return true;
    }
    return false;
  }

  protected onFormulaBarKeydown(event: KeyboardEvent): void {
    if (this.handleFormulaSuggestionKey(event)) return;
    if (event.key === 'Enter') {
      event.preventDefault();
      const cell = this.selectedCell();
      if (!cell) return;
      if (!this.editingCell()) this.editingCell.set(cell);
      void this.commitCellEdit();
      const next = nextSpreadsheetCell(cell, event.shiftKey ? 'up' : 'down', this.displayRowCount, this.displayColCount);
      this.selectCell(next.row, next.col);
      this.focusXlsxGridAfterNavigation();
    } else if (event.key === 'Escape') {
      this.editingCell.set(null);
      this.clearFormulaRangeMode();
      this.xlsxFormulaHighlights.set([]);
    }
  }

  protected onGridKeydown(event: KeyboardEvent): void {
    const cell = this.selectedCell();
    if (!cell || this.editingCell()) return;
    const sheet = this.activeSheet;
    if (!sheet) return;

    // Close context menu on any keypress
    this.xlsxContextMenu.set(null);

    // Ctrl+C / Ctrl+X / Ctrl+V / Ctrl+A / Ctrl+Z / Ctrl+Y
    if (event.ctrlKey || event.metaKey) {
      if (event.key === 'c' || event.key === 'C') {
        event.preventDefault();
        this.xlsxCopy();
        return;
      } else if (event.key === 'x' || event.key === 'X') {
        event.preventDefault();
        this.xlsxCut();
        return;
      } else if (event.key === 'v' || event.key === 'V') {
        event.preventDefault();
        this.xlsxPaste();
        return;
      } else if (event.key === 'a' || event.key === 'A') {
        event.preventDefault();
        this.onCornerClick(); // Select all
        return;
      } else if (event.key === 'z' || event.key === 'Z') {
        event.preventDefault();
        if (event.shiftKey) {
          this.xlsxRedo();
        } else {
          this.xlsxUndo();
        }
        return;
      } else if (event.key === 'y' || event.key === 'Y') {
        event.preventDefault();
        this.xlsxRedo();
        return;
      }
    }

    if (event.key === 'ArrowDown') {
      event.preventDefault();
      const newRow = Math.min(cell.row + 1, this.displayRowCount - 1);
      if (event.shiftKey) {
        const anchor = this.xlsxSelectionStart() ?? cell;
        this.xlsxSelectionStart.set(anchor);
        const end = this.xlsxSelectionEnd() ?? cell;
        this.xlsxSelectionEnd.set({ row: Math.min(end.row + 1, this.displayRowCount - 1), col: end.col });
      } else {
        this.selectCell(newRow, cell.col);
      }
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      const newRow = Math.max(cell.row - 1, 0);
      if (event.shiftKey) {
        const anchor = this.xlsxSelectionStart() ?? cell;
        this.xlsxSelectionStart.set(anchor);
        const end = this.xlsxSelectionEnd() ?? cell;
        this.xlsxSelectionEnd.set({ row: Math.max(end.row - 1, 0), col: end.col });
      } else {
        this.selectCell(newRow, cell.col);
      }
    } else if (event.key === 'ArrowRight') {
      event.preventDefault();
      const newCol = Math.min(cell.col + 1, this.displayColCount - 1);
      if (event.shiftKey) {
        const anchor = this.xlsxSelectionStart() ?? cell;
        this.xlsxSelectionStart.set(anchor);
        const end = this.xlsxSelectionEnd() ?? cell;
        this.xlsxSelectionEnd.set({ row: end.row, col: Math.min(end.col + 1, this.displayColCount - 1) });
      } else {
        this.selectCell(cell.row, newCol);
      }
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault();
      const newCol = Math.max(cell.col - 1, 0);
      if (event.shiftKey) {
        const anchor = this.xlsxSelectionStart() ?? cell;
        this.xlsxSelectionStart.set(anchor);
        const end = this.xlsxSelectionEnd() ?? cell;
        this.xlsxSelectionEnd.set({ row: end.row, col: Math.max(end.col - 1, 0) });
      } else {
        this.selectCell(cell.row, newCol);
      }
    } else if (event.key === 'Enter') {
      event.preventDefault();
      const next = nextSpreadsheetCell(cell, event.shiftKey ? 'up' : 'down', this.displayRowCount, this.displayColCount);
      this.selectCell(next.row, next.col);
    } else if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      // Delete selected range
      const r = this.xlsxRange;
      if (r && (r.startRow !== r.endRow || r.startCol !== r.endCol)) {
        this.xlsxDeleteSelection();
      } else {
        this.formulaBarValue.set('');
        this.editingCell.set(cell);
        this.commitCellEdit();
      }
    } else if (event.key === 'Escape') {
      this.xlsxSelectionStart.set(null);
      this.xlsxSelectionEnd.set(null);
      this.xlsxSelectingFullRows.set(false);
      this.xlsxSelectingFullCols.set(false);
    } else if (event.key.length === 1 && !event.ctrlKey && !event.metaKey) {
      // Start typing in cell
      event.preventDefault();
      this.startEditCell(cell.row, cell.col, event.key);
    }
  }

  protected getColumnLabel(index: number): string {
    let label = '';
    let n = index;
    while (true) {
      label = String.fromCharCode(65 + (n % 26)) + label;
      n = Math.floor(n / 26) - 1;
      if (n < 0) break;
    }
    return label;
  }

  protected getCellRef(row: number, col: number): string {
    return `${this.getColumnLabel(col)}${row + 1}`;
  }

  // ── Range selection helpers ──

  /** Normalized selection range */
  protected get xlsxRange(): { startRow: number; startCol: number; endRow: number; endCol: number } | null {
    const s = this.xlsxSelectionStart();
    const e = this.xlsxSelectionEnd();
    if (!s || !e) {
      // Single cell selection
      const sc = this.selectedCell();
      if (!sc) return null;
      return { startRow: sc.row, startCol: sc.col, endRow: sc.row, endCol: sc.col };
    }
    return {
      startRow: Math.min(s.row, e.row),
      startCol: Math.min(s.col, e.col),
      endRow: Math.max(s.row, e.row),
      endCol: Math.max(s.col, e.col),
    };
  }

  /** Check if cell is within selection range */
  protected isCellInRange(row: number, col: number): boolean {
    const r = this.xlsxRange;
    if (!r) return false;
    // Don't highlight the single active cell with range style
    if (r.startRow === r.endRow && r.startCol === r.endCol) return false;
    return row >= r.startRow && row <= r.endRow && col >= r.startCol && col <= r.endCol;
  }

  /** Check if row header should highlight */
  protected isRowInRange(row: number): boolean {
    const r = this.xlsxRange;
    if (!r) return false;
    return row >= r.startRow && row <= r.endRow;
  }

  /** Check if col header should highlight */
  protected isColInRange(col: number): boolean {
    const r = this.xlsxRange;
    if (!r) return false;
    return col >= r.startCol && col <= r.endCol;
  }

  /** Get selection range label for status bar */
  protected get xlsxRangeLabel(): string {
    const r = this.xlsxRange;
    if (!r) return '';
    if (r.startRow === r.endRow && r.startCol === r.endCol) {
      return this.getCellRef(r.startRow, r.startCol);
    }
    return `${this.getCellRef(r.startRow, r.startCol)}:${this.getCellRef(r.endRow, r.endCol)}`;
  }

  protected onCellMousedown(event: MouseEvent, row: number, col: number): void {
    this.xlsxContextMenu.set(null);
    if (event.button === 2) return; // right-click handled separately

    // Formula range-selection mode: if we're editing a formula cell, clicking another cell inserts a ref
    const editing = this.editingCell();
    if (editing && this.isInFormulaRangeMode()) {
      // Don't commit edit — instead start range capture
      event.preventDefault();
      event.stopPropagation();
      this.xlsxFormulaRangeMode = true;
      this.xlsxFormulaRangeDragging = true;
      this.xlsxFormulaRangeAnchor = { row, col };
      this.xlsxFormulaRangeCurrent = { row, col };
      this.insertOrUpdateFormulaRange(row, col, row, col);
      return;
    }

    this.xlsxGridRef()?.nativeElement.focus({ preventScroll: true });

    if (event.shiftKey && this.selectedCell()) {
      // Extend selection with Shift+click
      const anchor = this.xlsxSelectionStart() ?? this.selectedCell()!;
      this.xlsxSelectionStart.set(anchor);
      this.xlsxSelectionEnd.set({ row, col });
      this.xlsxSelectingFullRows.set(false);
      this.xlsxSelectingFullCols.set(false);
      event.preventDefault();
      return;
    }

    // Start new selection
    this.selectedCell.set({ row, col });
    this.xlsxSelectionStart.set({ row, col });
    this.xlsxSelectionEnd.set({ row, col });
    this.xlsxSelectingFullRows.set(false);
    this.xlsxSelectingFullCols.set(false);
    this.xlsxIsDragging = true;

    // Commit any in-progress cell edit before clearing it
    if (this.editingCell()) {
      this.commitCellEdit();
    } else {
      this.editingCell.set(null);
    }
    this.xlsxFormulaHighlights.set([]);

    // Show formula if cell has one, otherwise show value
    const formula = this.xlsxCellFormulas.get(this.getCellFormatKey(row, col));
    if (formula) {
      this.formulaBarValue.set(formula);
    } else {
      const value = this.getCellValue(row, col);
      this.formulaBarValue.set(value != null ? String(value) : '');
    }
    this.updateXlsxToolbarFromCell(row, col);
  }

  protected onCellMouseenter(row: number, col: number): void {
    // Formula range drag
    if (this.xlsxFormulaRangeDragging && this.xlsxFormulaRangeAnchor) {
      this.xlsxFormulaRangeCurrent = { row, col };
      const a = this.xlsxFormulaRangeAnchor;
      this.insertOrUpdateFormulaRange(
        Math.min(a.row, row), Math.min(a.col, col),
        Math.max(a.row, row), Math.max(a.col, col)
      );
      return;
    }
    if (!this.xlsxIsDragging) return;
    this.xlsxSelectionEnd.set({ row, col });
  }

  @HostListener('document:mouseup')
  onDocumentMouseup(): void {
    if (this.xlsxResize) {
      this.xlsxResize = null;
      void this.persistXlsxDimensions();
    }
    this.xlsxIsDragging = false;
    if (this.xlsxFormulaRangeDragging) {
      this.xlsxFormulaRangeDragging = false;
      this.xlsxFormulaRangeAnchor = null;
      this.xlsxFormulaRangeCurrent = null;
      // Re-focus the cell input after range selection
      this.pendingFocusCellInput = true;
    }
  }

  protected onRowHeaderClick(event: MouseEvent, row: number): void {
    this.xlsxContextMenu.set(null);
    if (this.editingCell()) this.commitCellEdit();
    if (event.shiftKey && this.xlsxSelectionStart()) {
      const startRow = this.xlsxSelectionStart()!.row;
      this.xlsxSelectionEnd.set({ row, col: this.displayColCount - 1 });
      this.xlsxSelectionStart.set({ row: startRow, col: 0 });
    } else {
      this.selectedCell.set({ row, col: 0 });
      this.xlsxSelectionStart.set({ row, col: 0 });
      this.xlsxSelectionEnd.set({ row, col: this.displayColCount - 1 });
    }
    this.xlsxSelectingFullRows.set(true);
    this.xlsxSelectingFullCols.set(false);
    this.formulaBarValue.set('');
  }

  protected onRowHeaderContextMenu(event: MouseEvent, row: number): void {
    event.preventDefault();
    event.stopPropagation();
    // Select the full row first
    if (this.editingCell()) this.commitCellEdit();
    this.selectedCell.set({ row, col: 0 });
    this.xlsxSelectionStart.set({ row, col: 0 });
    this.xlsxSelectionEnd.set({ row, col: this.displayColCount - 1 });
    this.xlsxSelectingFullRows.set(true);
    this.xlsxSelectingFullCols.set(false);
    this.formulaBarValue.set('');
    // Open context menu
    this.xlsxContextMenu.set({ x: event.clientX, y: event.clientY });
  }

  protected onColHeaderClick(event: MouseEvent, col: number): void {
    this.xlsxContextMenu.set(null);
    if (this.editingCell()) this.commitCellEdit();
    if (event.shiftKey && this.xlsxSelectionStart()) {
      const startCol = this.xlsxSelectionStart()!.col;
      this.xlsxSelectionEnd.set({ row: this.displayRowCount - 1, col });
      this.xlsxSelectionStart.set({ row: 0, col: startCol });
    } else {
      this.selectedCell.set({ row: 0, col });
      this.xlsxSelectionStart.set({ row: 0, col });
      this.xlsxSelectionEnd.set({ row: this.displayRowCount - 1, col });
    }
    this.xlsxSelectingFullRows.set(false);
    this.xlsxSelectingFullCols.set(true);
    this.formulaBarValue.set('');
  }

  protected onColHeaderContextMenu(event: MouseEvent, col: number): void {
    event.preventDefault();
    event.stopPropagation();
    // Select the full column first
    if (this.editingCell()) this.commitCellEdit();
    this.selectedCell.set({ row: 0, col });
    this.xlsxSelectionStart.set({ row: 0, col });
    this.xlsxSelectionEnd.set({ row: this.displayRowCount - 1, col });
    this.xlsxSelectingFullRows.set(false);
    this.xlsxSelectingFullCols.set(true);
    this.formulaBarValue.set('');
    // Open context menu
    this.xlsxContextMenu.set({ x: event.clientX, y: event.clientY });
  }

  protected onCornerClick(): void {
    if (this.editingCell()) this.commitCellEdit();
    // Select all cells
    this.selectedCell.set({ row: 0, col: 0 });
    this.xlsxSelectionStart.set({ row: 0, col: 0 });
    this.xlsxSelectionEnd.set({ row: this.displayRowCount - 1, col: this.displayColCount - 1 });
    this.xlsxSelectingFullRows.set(true);
    this.xlsxSelectingFullCols.set(true);
  }

  // ── Clipboard operations ──

  protected xlsxCopy(): void {
    const r = this.xlsxRange;
    if (!r) return;
    const data: (string | number | boolean | null)[][] = [];
    for (let row = r.startRow; row <= r.endRow; row++) {
      const rowData: (string | number | boolean | null)[] = [];
      for (let col = r.startCol; col <= r.endCol; col++) {
        rowData.push(this.getCellValue(row, col));
      }
      data.push(rowData);
    }
    this.xlsxClipboard = { data, startRow: r.startRow, startCol: r.startCol, isCut: false };
    this.xlsxHasClipboard.set(true);

    // Also copy to system clipboard as TSV
    const tsv = data.map(row => row.map(v => v != null ? String(v) : '').join('\t')).join('\n');
    navigator.clipboard.writeText(tsv).catch(() => {});
  }

  protected xlsxCut(): void {
    const r = this.xlsxRange;
    if (!r) return;
    const data: (string | number | boolean | null)[][] = [];
    for (let row = r.startRow; row <= r.endRow; row++) {
      const rowData: (string | number | boolean | null)[] = [];
      for (let col = r.startCol; col <= r.endCol; col++) {
        rowData.push(this.getCellValue(row, col));
      }
      data.push(rowData);
    }
    this.xlsxClipboard = { data, startRow: r.startRow, startCol: r.startCol, isCut: true };
    this.xlsxHasClipboard.set(true);

    // Also copy to system clipboard and clear source cells
    const tsv = data.map(row => row.map(v => v != null ? String(v) : '').join('\t')).join('\n');
    navigator.clipboard.writeText(tsv).catch(() => {});

    // Clear source cells
    this.clearRangeCells(r.startRow, r.startCol, r.endRow, r.endCol);
  }

  protected async xlsxPaste(): Promise<void> {
    const sc = this.selectedCell();
    if (!sc) return;
    this.pushUndoSnapshot();

    let data: (string | number | boolean | null)[][] | null = null;

    if (this.xlsxClipboard) {
      data = this.xlsxClipboard.data;
      if (this.xlsxClipboard.isCut) {
        this.xlsxClipboard = null;
        this.xlsxHasClipboard.set(false);
      }
    } else {
      // Try system clipboard
      try {
        const text = await navigator.clipboard.readText();
        if (text) {
          data = text.split('\n').map(row => row.split('\t').map(v => {
            return this.parseSmartNumber(v);
          }));
        }
      } catch { /* clipboard access denied */ }
    }

    if (!data || data.length === 0) return;

    const sheets = this.xlsxSheets();
    const sheetIdx = this.activeSheetIndex();
    const sheet = sheets[sheetIdx];
    if (!sheet) return;

    // Paste data starting at selected cell
    for (let dr = 0; dr < data.length; dr++) {
      const targetRow = sc.row + dr;
      while (sheet.data.length <= targetRow) sheet.data.push([]);
      for (let dc = 0; dc < data[dr].length; dc++) {
        const targetCol = sc.col + dc;
        while (sheet.data[targetRow].length <= targetCol) sheet.data[targetRow].push(null);
        sheet.data[targetRow][targetCol] = data[dr][dc];
      }
    }
    this.xlsxSheets.set([...sheets]);

    // Persist to backend — rebuild the whole xlsx
    await this.persistXlsxData();

    // Select pasted range
    this.xlsxSelectionStart.set(sc);
    this.xlsxSelectionEnd.set({ row: sc.row + data.length - 1, col: sc.col + (data[0]?.length ?? 1) - 1 });
  }

  protected async xlsxDeleteSelection(): Promise<void> {
    const r = this.xlsxRange;
    if (!r) return;
    this.pushUndoSnapshot();
    await this.clearRangeCells(r.startRow, r.startCol, r.endRow, r.endCol);
  }

  private async clearRangeCells(startRow: number, startCol: number, endRow: number, endCol: number): Promise<void> {
    const sheets = this.xlsxSheets();
    const sheetIdx = this.activeSheetIndex();
    const sheet = sheets[sheetIdx];
    if (!sheet) return;

    for (let r = startRow; r <= endRow; r++) {
      if (!sheet.data[r]) continue;
      for (let c = startCol; c <= endCol; c++) {
        if (c < sheet.data[r].length) {
          sheet.data[r][c] = null;
        }
      }
    }
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) this.xlsxCellFormulas.delete(this.getCellFormatKey(r, c));
    }
    this.xlsxSheets.set([...sheets]);
    await this.persistXlsxData();
  }

  /** Push current state onto the undo stack before a mutation */
  private pushUndoSnapshot(): void {
    if (!this.xlsxBase64) return;
    this.xlsxUndoStack.push(this.xlsxBase64);
    if (this.xlsxUndoStack.length > this.MAX_UNDO) {
      this.xlsxUndoStack.shift();
    }
    // Any new mutation clears the redo stack
    this.xlsxRedoStack = [];
  }

  /** Restore a base64 snapshot: re-parse sheets and save to DB */
  private async restoreSnapshot(base64: string): Promise<void> {
    this.xlsxBase64 = base64;
    const result = await window.electronAPI.xlsxParseData(base64);
    if (result.success && result.sheets) {
      this.loadXlsxFormulaMetadata(result.sheets);
      this.xlsxSheets.set(result.sheets);
      this.recalculateXlsxFormulas();
    }
    // Persist to DB
    const noteId = this.noteService.selectedNoteId();
    if (noteId !== null) {
      this.skipNextSync = true;
      await window.electronAPI.notesUpdate(noteId, { file_data: base64 });
    }
  }

  private async xlsxUndo(): Promise<void> {
    if (this.xlsxUndoStack.length === 0) return;
    // Save current state to redo
    if (this.xlsxBase64) {
      this.xlsxRedoStack.push(this.xlsxBase64);
    }
    const snapshot = this.xlsxUndoStack.pop()!;
    await this.restoreSnapshot(snapshot);
  }

  private async xlsxRedo(): Promise<void> {
    if (this.xlsxRedoStack.length === 0) return;
    // Save current state to undo
    if (this.xlsxBase64) {
      this.xlsxUndoStack.push(this.xlsxBase64);
    }
    const snapshot = this.xlsxRedoStack.pop()!;
    await this.restoreSnapshot(snapshot);
  }

  /** Persist all current sheet data to the xlsx backend */
  private async persistXlsxData(): Promise<void> {
    const sheets = this.xlsxSheets();
    const sheetIdx = this.activeSheetIndex();
    const sheet = sheets[sheetIdx];
    if (!sheet) return;

    try {
      this.recalculateXlsxFormulas();
      const result = await window.electronAPI.xlsxReplaceSheetData(
        this.xlsxBase64,
        sheetIdx,
        sheet.data,
        this.activeFormulaCells(),
        sheet.merges ?? [],
      );
      if (!result.success || !result.base64) throw new Error(result.error ?? 'No se pudo guardar la hoja');
      this.xlsxBase64 = result.base64;
      // Save to note
      const noteId = this.noteService.selectedNoteId();
      if (noteId !== null) {
        this.skipNextSync = true;
        await window.electronAPI.notesUpdate(noteId, { file_data: result.base64 });
      }
    } catch (err) {
      console.error('Error persisting xlsx data:', err);
    }
  }

  private loadXlsxFormulaMetadata(sheets: XlsxSheetView[]): void {
    this.xlsxCellFormulas.clear();
    sheets.forEach((sheet, sheetIndex) => {
      for (const formula of sheet.formulas ?? []) {
        this.xlsxCellFormulas.set(`${sheetIndex}:${formula.row}:${formula.col}`, `=${formula.formula.replace(/^=/, '')}`);
      }
    });
  }

  /** XLSX stores function names in English even when the UI accepts localized aliases. */
  private canonicalizeFormula(formula: string): string {
    const aliases: Record<string, string> = {
      'SUMA': 'SUM',
      'PROMEDIO': 'AVERAGE',
      'CONTAR': 'COUNT',
      'SI': 'IF',
      'CONTAR.SI': 'COUNTIF',
      'SUMAR.SI': 'SUMIF',
      'REDONDEAR': 'ROUND',
    };
    return formula.replace(/\b(SUMAR\.SI|CONTAR\.SI|PROMEDIO|REDONDEAR|CONTAR|SUMA|SI)\s*(?=\()/gi, name => aliases[name.toUpperCase()] ?? name);
  }

  private activeFormulaCells(): { row: number; col: number; value: XlsxCellValue; formula: string }[] {
    const sheetIndex = this.activeSheetIndex();
    const prefix = `${sheetIndex}:`;
    const cells: { row: number; col: number; value: XlsxCellValue; formula: string }[] = [];
    for (const [key, formula] of this.xlsxCellFormulas) {
      if (!key.startsWith(prefix)) continue;
      const [, rowText, colText] = key.split(':');
      const row = Number(rowText);
      const col = Number(colText);
      cells.push({ row, col, value: this.getCellValue(row, col), formula: this.canonicalizeFormula(formula) });
    }
    return cells;
  }

  private recalculateXlsxFormulas(): void {
    const sheet = this.activeSheet;
    if (!sheet) return;
    const formulas = this.activeFormulaCells();
    // A few passes resolve simple chains such as C1 -> B1 -> A1.
    for (let pass = 0; pass < Math.min(8, Math.max(1, formulas.length)); pass++) {
      let changed = false;
      for (const formulaCell of formulas) {
        const formula = this.xlsxCellFormulas.get(this.getCellFormatKey(formulaCell.row, formulaCell.col));
        if (!formula) continue;
        const value = this.evaluateFormula(formula);
        while (sheet.data.length <= formulaCell.row) sheet.data.push([]);
        while (sheet.data[formulaCell.row].length <= formulaCell.col) sheet.data[formulaCell.row].push(null);
        if (sheet.data[formulaCell.row][formulaCell.col] !== value) {
          sheet.data[formulaCell.row][formulaCell.col] = value;
          changed = true;
        }
      }
      if (!changed) break;
    }
    this.xlsxSheets.set([...this.xlsxSheets()]);
  }

  private rewriteFormulaReferences(
    formula: string,
    mapper: (row: number, col: number) => { row: number; col: number } | null,
  ): string {
    return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/gi, (_match, colAbsolute: string, colLabel: string, rowAbsolute: string, rowText: string) => {
      const parsed = this.parseCellRef(`${colLabel}${rowText}`);
      if (!parsed) return _match;
      const mapped = mapper(parsed.row, parsed.col);
      if (!mapped) return '#REF!';
      return `${colAbsolute}${this.getColumnLabel(mapped.col)}${rowAbsolute}${mapped.row + 1}`;
    });
  }

  private remapXlsxCellMetadata(mapper: (row: number, col: number) => { row: number; col: number } | null): void {
    const sheetIndex = this.activeSheetIndex();
    const remappedFormulas = new Map<string, string>();
    for (const [key, formula] of this.xlsxCellFormulas) {
      const [sheetText, rowText, colText] = key.split(':');
      if (Number(sheetText) !== sheetIndex) {
        remappedFormulas.set(key, formula);
        continue;
      }
      const mapped = mapper(Number(rowText), Number(colText));
      if (mapped) remappedFormulas.set(`${sheetIndex}:${mapped.row}:${mapped.col}`, this.rewriteFormulaReferences(formula, mapper));
    }
    this.xlsxCellFormulas = remappedFormulas;

    const remappedFormats = new Map<string, { bold?: boolean; italic?: boolean; underline?: boolean; align?: 'left' | 'center' | 'right'; fontSize?: number }>();
    for (const [key, format] of this.xlsxCellFormats) {
      const [sheetText, rowText, colText] = key.split(':');
      if (Number(sheetText) !== sheetIndex) {
        remappedFormats.set(key, format);
        continue;
      }
      const mapped = mapper(Number(rowText), Number(colText));
      if (mapped) remappedFormats.set(`${sheetIndex}:${mapped.row}:${mapped.col}`, format);
    }
    this.xlsxCellFormats = remappedFormats;

    const sheet = this.activeSheet;
    if (sheet?.merges) {
      sheet.merges = sheet.merges.flatMap(merge => {
        const start = mapper(merge.s.r, merge.s.c);
        const end = mapper(merge.e.r, merge.e.c);
        if (!start || !end) return [];
        return [{
          s: { r: Math.min(start.row, end.row), c: Math.min(start.col, end.col) },
          e: { r: Math.max(start.row, end.row), c: Math.max(start.col, end.col) },
        }];
      });
    }
  }

  // ── Context menu ──

  protected onXlsxContextMenu(event: MouseEvent, row: number, col: number): void {
    event.preventDefault();
    event.stopPropagation();

    // If right-clicking outside current range, select the clicked cell
    const r = this.xlsxRange;
    if (!r || row < r.startRow || row > r.endRow || col < r.startCol || col > r.endCol) {
      this.selectedCell.set({ row, col });
      this.xlsxSelectionStart.set({ row, col });
      this.xlsxSelectionEnd.set({ row, col });
    }

    this.xlsxContextMenu.set({ x: event.clientX, y: event.clientY });
  }

  protected closeXlsxContextMenu(): void {
    this.xlsxContextMenu.set(null);
  }

  protected async onContextMenuAction(action: string): Promise<void> {
    this.xlsxContextMenu.set(null);
    switch (action) {
      case 'copy': this.xlsxCopy(); break;
      case 'cut': this.xlsxCut(); break;
      case 'paste': await this.xlsxPaste(); break;
      case 'delete': await this.xlsxDeleteSelection(); break;
      case 'insertRowAbove': await this.xlsxInsertRow(false); break;
      case 'insertRowBelow': await this.xlsxInsertRow(true); break;
      case 'insertColLeft': await this.xlsxInsertCol(false); break;
      case 'insertColRight': await this.xlsxInsertCol(true); break;
      case 'deleteRow': await this.xlsxDeleteRow(); break;
      case 'deleteCol': await this.xlsxDeleteCol(); break;
      case 'moveRowUp': await this.xlsxMoveRows(-1); break;
      case 'moveRowDown': await this.xlsxMoveRows(1); break;
      case 'moveColLeft': await this.xlsxMoveCols(-1); break;
      case 'moveColRight': await this.xlsxMoveCols(1); break;
    }
  }

  protected async xlsxInsertRow(below: boolean): Promise<void> {
    const r = this.xlsxRange;
    if (!r) return;
    this.pushUndoSnapshot();
    const sheets = this.xlsxSheets();
    const sheet = sheets[this.activeSheetIndex()];
    if (!sheet) return;
    const insertAt = below ? r.endRow + 1 : r.startRow;
    this.remapXlsxCellMetadata((row, col) => ({ row: row >= insertAt ? row + 1 : row, col }));
    const emptyRow: (string | number | boolean | null)[] = new Array(sheet.colCount).fill(null);
    sheet.data.splice(insertAt, 0, emptyRow);
    this.xlsxSheets.set([...sheets]);
    await this.persistXlsxData();
  }

  protected async xlsxInsertCol(right: boolean): Promise<void> {
    const r = this.xlsxRange;
    if (!r) return;
    this.pushUndoSnapshot();
    const sheets = this.xlsxSheets();
    const sheet = sheets[this.activeSheetIndex()];
    if (!sheet) return;
    const insertAt = right ? r.endCol + 1 : r.startCol;
    this.remapXlsxCellMetadata((row, col) => ({ row, col: col >= insertAt ? col + 1 : col }));
    for (const row of sheet.data) {
      row.splice(insertAt, 0, null);
    }
    sheet.colCount++;
    this.xlsxSheets.set([...sheets]);
    await this.persistXlsxData();
  }

  protected async xlsxDeleteRow(): Promise<void> {
    const r = this.xlsxRange;
    if (!r) return;
    this.pushUndoSnapshot();
    const sheets = this.xlsxSheets();
    const sheet = sheets[this.activeSheetIndex()];
    if (!sheet) return;
    const count = r.endRow - r.startRow + 1;
    this.remapXlsxCellMetadata((row, col) => {
      if (row >= r.startRow && row <= r.endRow) return null;
      return { row: row > r.endRow ? row - count : row, col };
    });
    sheet.data.splice(r.startRow, count);
    this.xlsxSheets.set([...sheets]);
    this.selectedCell.set(null);
    this.xlsxSelectionStart.set(null);
    this.xlsxSelectionEnd.set(null);
    await this.persistXlsxData();
  }

  protected async xlsxDeleteCol(): Promise<void> {
    const r = this.xlsxRange;
    if (!r) return;
    this.pushUndoSnapshot();
    const sheets = this.xlsxSheets();
    const sheet = sheets[this.activeSheetIndex()];
    if (!sheet) return;
    const count = r.endCol - r.startCol + 1;
    this.remapXlsxCellMetadata((row, col) => {
      if (col >= r.startCol && col <= r.endCol) return null;
      return { row, col: col > r.endCol ? col - count : col };
    });
    for (const row of sheet.data) {
      row.splice(r.startCol, count);
    }
    sheet.colCount = Math.max(0, sheet.colCount - count);
    this.xlsxSheets.set([...sheets]);
    this.selectedCell.set(null);
    this.xlsxSelectionStart.set(null);
    this.xlsxSelectionEnd.set(null);
    await this.persistXlsxData();
  }

  protected async xlsxMoveRows(direction: -1 | 1): Promise<void> {
    if (!this.xlsxSelectingFullRows()) return;
    const range = this.xlsxRange;
    const sheet = this.activeSheet;
    if (!range || !sheet) return;
    const count = range.endRow - range.startRow + 1;
    if ((direction < 0 && range.startRow === 0) || (direction > 0 && range.endRow >= this.displayRowCount - 1)) return;
    this.pushUndoSnapshot();

    while (sheet.data.length <= range.endRow + (direction > 0 ? 1 : 0)) {
      sheet.data.push(new Array(sheet.colCount).fill(null));
    }

    const adjacentRow = direction < 0 ? range.startRow - 1 : range.endRow + 1;
    this.remapXlsxCellMetadata((row, col) => {
      if (row >= range.startRow && row <= range.endRow) return { row: row + direction, col };
      if (row === adjacentRow) return { row: direction < 0 ? range.endRow : range.startRow, col };
      return { row, col };
    });

    const moved = sheet.data.splice(range.startRow, count);
    sheet.data.splice(range.startRow + direction, 0, ...moved);
    const nextStart = range.startRow + direction;
    const nextEnd = range.endRow + direction;
    this.selectedCell.set({ row: nextStart, col: this.selectedCell()?.col ?? 0 });
    this.xlsxSelectionStart.set({ row: nextStart, col: 0 });
    this.xlsxSelectionEnd.set({ row: nextEnd, col: Math.max(0, this.displayColCount - 1) });
    this.xlsxSheets.set([...this.xlsxSheets()]);
    await this.persistXlsxData();
  }

  protected async xlsxMoveCols(direction: -1 | 1): Promise<void> {
    if (!this.xlsxSelectingFullCols()) return;
    const range = this.xlsxRange;
    const sheet = this.activeSheet;
    if (!range || !sheet) return;
    const count = range.endCol - range.startCol + 1;
    if ((direction < 0 && range.startCol === 0) || (direction > 0 && range.endCol >= this.displayColCount - 1)) return;
    this.pushUndoSnapshot();

    sheet.colCount = Math.max(sheet.colCount, range.endCol + (direction > 0 ? 2 : 1));

    const adjacentCol = direction < 0 ? range.startCol - 1 : range.endCol + 1;
    this.remapXlsxCellMetadata((row, col) => {
      if (col >= range.startCol && col <= range.endCol) return { row, col: col + direction };
      if (col === adjacentCol) return { row, col: direction < 0 ? range.endCol : range.startCol };
      return { row, col };
    });

    for (const row of sheet.data) {
      while (row.length < sheet.colCount) row.push(null);
      const moved = row.splice(range.startCol, count);
      row.splice(range.startCol + direction, 0, ...moved);
    }
    const nextStart = range.startCol + direction;
    const nextEnd = range.endCol + direction;
    this.selectedCell.set({ row: this.selectedCell()?.row ?? 0, col: nextStart });
    this.xlsxSelectionStart.set({ row: 0, col: nextStart });
    this.xlsxSelectionEnd.set({ row: Math.max(0, this.displayRowCount - 1), col: nextEnd });
    this.xlsxSheets.set([...this.xlsxSheets()]);
    await this.persistXlsxData();
  }

  // ── Cell formatting (visual only) ──

  private getCellFormatKey(row: number, col: number): string {
    return `${this.activeSheetIndex()}:${row}:${col}`;
  }

  protected getCellFormat(row: number, col: number): { bold?: boolean; italic?: boolean; underline?: boolean; align?: 'left' | 'center' | 'right'; fontSize?: number } {
    return this.xlsxCellFormats.get(this.getCellFormatKey(row, col)) ?? {};
  }

  protected getCellStyle(row: number, col: number): Record<string, string> {
    const fmt = this.getCellFormat(row, col);
    const style: Record<string, string> = {};
    if (fmt.bold) style['font-weight'] = 'bold';
    if (fmt.italic) style['font-style'] = 'italic';
    if (fmt.underline) style['text-decoration'] = 'underline';
    if (fmt.align) style['text-align'] = fmt.align;
    if (fmt.fontSize) style['font-size'] = fmt.fontSize + 'px';

    // Formula range highlight
    for (const h of this.xlsxFormulaHighlights()) {
      if (row >= h.startRow && row <= h.endRow && col >= h.startCol && col <= h.endCol) {
        style['background-color'] = h.color + '1a'; // ~10% opacity
        const borders: string[] = [];
        if (row === h.startRow) borders.push(`solid 2px ${h.color}`);
        if (row === h.endRow) borders.push(`solid 2px ${h.color}`);
        else borders.push('none');
        if (col === h.startCol) borders.push(`solid 2px ${h.color}`);
        else borders.push('none');
        if (col === h.endCol) borders.push(`solid 2px ${h.color}`);
        else borders.push('none');
        // border: top right bottom left
        if (row === h.startRow) style['border-top'] = `2px solid ${h.color}`;
        if (row === h.endRow) style['border-bottom'] = `2px solid ${h.color}`;
        if (col === h.startCol) style['border-left'] = `2px solid ${h.color}`;
        if (col === h.endCol) style['border-right'] = `2px solid ${h.color}`;
        break;
      }
    }

    return style;
  }

  private applyFormatToRange(prop: string, value: unknown): void {
    const r = this.xlsxRange;
    if (!r) return;
    for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startCol; col <= r.endCol; col++) {
        const key = this.getCellFormatKey(row, col);
        const existing = this.xlsxCellFormats.get(key) ?? {};
        (existing as Record<string, unknown>)[prop] = value;
        this.xlsxCellFormats.set(key, existing);
      }
    }
    // Force change detection on sheets
    this.xlsxSheets.set([...this.xlsxSheets()]);
  }

  protected toggleXlsxBold(): void {
    const newVal = !this.xlsxBold();
    this.xlsxBold.set(newVal);
    this.applyFormatToRange('bold', newVal);
  }

  protected toggleXlsxItalic(): void {
    const newVal = !this.xlsxItalic();
    this.xlsxItalic.set(newVal);
    this.applyFormatToRange('italic', newVal);
  }

  protected toggleXlsxUnderline(): void {
    const newVal = !this.xlsxUnderline();
    this.xlsxUnderline.set(newVal);
    this.applyFormatToRange('underline', newVal);
  }

  protected setXlsxAlign(align: 'left' | 'center' | 'right'): void {
    this.xlsxAlign.set(align);
    this.applyFormatToRange('align', align);
  }

  protected setXlsxFontSize(size: number): void {
    this.xlsxFontSize.set(size);
    this.applyFormatToRange('fontSize', size);
  }

  private updateXlsxToolbarFromCell(row: number, col: number): void {
    const fmt = this.getCellFormat(row, col);
    this.xlsxBold.set(!!fmt.bold);
    this.xlsxItalic.set(!!fmt.italic);
    this.xlsxUnderline.set(!!fmt.underline);
    this.xlsxAlign.set(fmt.align ?? 'left');
    this.xlsxFontSize.set(fmt.fontSize ?? 12);
  }

  // ── Formula engine ──

  /** Parse a cell reference like "A1" into { row, col } (0-indexed) */
  private parseCellRef(ref: string): { row: number; col: number } | null {
    const match = ref.match(/^([A-Za-z]+)(\d+)$/);
    if (!match) return null;
    const colStr = match[1].toUpperCase();
    const rowNum = parseInt(match[2], 10) - 1;
    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
      col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
    col -= 1; // 0-indexed
    return { row: rowNum, col };
  }

  /** Get numeric values from a range of cells */
  private getRangeValues(startRef: string, endRef: string): number[] {
    const s = this.parseCellRef(startRef);
    const e = this.parseCellRef(endRef);
    if (!s || !e) return [];
    const sheet = this.activeSheet;
    if (!sheet) return [];

    const startRow = Math.min(s.row, e.row);
    const endRow = Math.max(s.row, e.row);
    const startCol = Math.min(s.col, e.col);
    const endCol = Math.max(s.col, e.col);
    const values: number[] = [];

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const val = sheet.data[r]?.[c];
        if (typeof val === 'number') {
          values.push(val);
        } else if (typeof val === 'string' && val !== '' && !isNaN(Number(val))) {
          values.push(Number(val));
        }
      }
    }
    return values;
  }

  /** Get numeric value from a single cell reference */
  private getCellRefValue(ref: string): number | null {
    const parsed = this.parseCellRef(ref);
    if (!parsed) return null;
    const sheet = this.activeSheet;
    if (!sheet) return null;
    const val = sheet.data[parsed.row]?.[parsed.col];
    if (typeof val === 'number') return val;
    if (typeof val === 'string' && val !== '' && !isNaN(Number(val))) return Number(val);
    return null;
  }

  /** Parse formula arguments — can be ranges (A1:B5), cell refs (A1), or numbers */
  private parseFormulaArgs(argsStr: string): number[] {
    const values: number[] = [];
    // Split by comma or semicolon (Excel uses ; in some locales)
    const parts = argsStr.split(/[,;]/).map(s => s.trim());
    for (const part of parts) {
      if (part.includes(':')) {
        // Range reference
        const [start, end] = part.split(':').map(s => s.trim());
        values.push(...this.getRangeValues(start, end));
      } else if (/^[A-Za-z]+\d+$/.test(part)) {
        // Single cell reference
        const v = this.getCellRefValue(part);
        if (v !== null) values.push(v);
      } else {
        // Literal number
        const n = Number(part);
        if (!isNaN(n)) values.push(n);
      }
    }
    return values;
  }

  /** Evaluate a formula string and return the computed result */
  protected evaluateFormula(formula: string): number | string | null {
    if (!formula.startsWith('=')) return null;
    const expr = formula.substring(1).trim();

    // Match function calls: FUNC(args)
    const funcMatch = expr.match(/^([A-Za-zÁÉÍÓÚÑ.]+)\((.*)\)$/i);
    if (!funcMatch) {
      return evaluateArithmeticExpression(expr, reference => this.getCellRefValue(reference));
    }

    const funcName = funcMatch[1].toUpperCase();
    const argsStr = funcMatch[2];
    const values = this.parseFormulaArgs(argsStr);

    switch (funcName) {
      case 'SUM':
      case 'SUMA':
        return values.reduce((a, b) => a + b, 0);

      case 'AVERAGE':
      case 'PROMEDIO':
        return values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0;

      case 'COUNT':
      case 'CONTAR':
        return values.length;

      case 'MIN':
        return values.length > 0 ? Math.min(...values) : 0;

      case 'MAX':
        return values.length > 0 ? Math.max(...values) : 0;

      case 'ABS':
        return values.length > 0 ? Math.abs(values[0]) : 0;

      case 'ROUND':
      case 'REDONDEAR': {
        // ROUND(value, decimals)
        const parts = argsStr.split(/[,;]/).map(s => s.trim());
        if (parts.length >= 2) {
          const vals0 = this.parseFormulaArgs(parts[0]);
          const vals1 = this.parseFormulaArgs(parts[1]);
          const v = vals0[0] ?? 0;
          const d = vals1[0] ?? 0;
          const factor = Math.pow(10, d);
          return Math.round(v * factor) / factor;
        }
        return values.length > 0 ? Math.round(values[0]) : 0;
      }

      case 'IF':
      case 'SI': {
        // =IF(condition, valueTrue, valueFalse) — basic: compares with >, <, =, >=, <=, <>
        const parts = this.splitFormulaTopLevel(argsStr);
        if (parts.length >= 3) {
          const condResult = this.evaluateCondition(parts[0].trim());
          const trueVal = this.evaluateSimpleValue(parts[1].trim());
          const falseVal = this.evaluateSimpleValue(parts[2].trim());
          return condResult ? trueVal : falseVal;
        }
        return '#ERROR!';
      }

      case 'COUNTIF':
      case 'CONTAR.SI': {
        // =COUNTIF(range, criteria)
        const parts = this.splitFormulaTopLevel(argsStr);
        if (parts.length >= 2) {
          const rangePart = parts[0].trim();
          const criteria = parts[1].trim().replace(/^"(.*)"$/, '$1');
          if (rangePart.includes(':')) {
            const [start, end] = rangePart.split(':').map(s => s.trim());
            return this.countIf(start, end, criteria);
          }
        }
        return 0;
      }

      case 'SUMIF':
      case 'SUMAR.SI': {
        // =SUMIF(range, criteria, [sum_range])
        const parts = this.splitFormulaTopLevel(argsStr);
        if (parts.length >= 2) {
          const rangePart = parts[0].trim();
          const criteria = parts[1].trim().replace(/^"(.*)"$/, '$1');
          const sumRangePart = parts[2]?.trim();
          if (rangePart.includes(':')) {
            const [start, end] = rangePart.split(':').map(s => s.trim());
            let sumStart = start, sumEnd = end;
            if (sumRangePart && sumRangePart.includes(':')) {
              [sumStart, sumEnd] = sumRangePart.split(':').map(s => s.trim());
            }
            return this.sumIf(start, end, criteria, sumStart, sumEnd);
          }
        }
        return 0;
      }

      default:
        return '#NAME?';
    }
  }

  /** Split formula args at top level (respecting nested parentheses) */
  private splitFormulaTopLevel(argsStr: string): string[] {
    const parts: string[] = [];
    let depth = 0;
    let current = '';
    for (const ch of argsStr) {
      if (ch === '(') { depth++; current += ch; }
      else if (ch === ')') { depth--; current += ch; }
      else if ((ch === ',' || ch === ';') && depth === 0) {
        parts.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
    if (current) parts.push(current);
    return parts;
  }

  /** Evaluate a simple condition like "A1>10", "B2>=5", "C3<>0" */
  private evaluateCondition(cond: string): boolean {
    const ops = ['>=', '<=', '<>', '!=', '>', '<', '='];
    for (const op of ops) {
      const idx = cond.indexOf(op);
      if (idx >= 0) {
        const left = this.evaluateSimpleValue(cond.substring(0, idx).trim());
        const right = this.evaluateSimpleValue(cond.substring(idx + op.length).trim());
        const l = typeof left === 'number' ? left : Number(left);
        const r = typeof right === 'number' ? right : Number(right);
        if (isNaN(l) || isNaN(r)) return false;
        switch (op) {
          case '>=': return l >= r;
          case '<=': return l <= r;
          case '<>': case '!=': return l !== r;
          case '>': return l > r;
          case '<': return l < r;
          case '=': return l === r;
        }
      }
    }
    return false;
  }

  /** Evaluate a simple value: cell ref, number, or quoted string */
  private evaluateSimpleValue(val: string): number | string {
    // Quoted string
    if (val.startsWith('"') && val.endsWith('"')) {
      return val.slice(1, -1);
    }
    // Cell reference
    if (/^[A-Za-z]+\d+$/.test(val)) {
      return this.getCellRefValue(val) ?? 0;
    }
    // Nested formula
    if (val.startsWith('=')) {
      return this.evaluateFormula(val) ?? 0;
    }
    // Function call without =
    if (/^[A-Za-z]+\(/.test(val)) {
      return this.evaluateFormula('=' + val) ?? 0;
    }
    // Number
    const n = Number(val);
    return isNaN(n) ? val : n;
  }

  /** COUNTIF implementation */
  private countIf(startRef: string, endRef: string, criteria: string): number {
    const s = this.parseCellRef(startRef);
    const e = this.parseCellRef(endRef);
    if (!s || !e) return 0;
    const sheet = this.activeSheet;
    if (!sheet) return 0;

    const { op, numVal: critNum, textVal: critText } = this.parseCriteria(criteria);
    let count = 0;
    for (let r = Math.min(s.row, e.row); r <= Math.max(s.row, e.row); r++) {
      for (let c = Math.min(s.col, e.col); c <= Math.max(s.col, e.col); c++) {
        const cellVal = sheet.data[r]?.[c];
        if (this.matchesCriteria(cellVal, op, critNum, critText)) count++;
      }
    }
    return count;
  }

  /** SUMIF implementation */
  private sumIf(condStartRef: string, condEndRef: string, criteria: string, sumStartRef: string, sumEndRef: string): number {
    const cs = this.parseCellRef(condStartRef);
    const ce = this.parseCellRef(condEndRef);
    const ss = this.parseCellRef(sumStartRef);
    const se = this.parseCellRef(sumEndRef);
    if (!cs || !ce || !ss || !se) return 0;
    const sheet = this.activeSheet;
    if (!sheet) return 0;

    const { op, numVal: critNum, textVal: critText } = this.parseCriteria(criteria);
    let sum = 0;
    const condRows = Math.max(cs.row, ce.row) - Math.min(cs.row, ce.row) + 1;
    const condCols = Math.max(cs.col, ce.col) - Math.min(cs.col, ce.col) + 1;

    for (let dr = 0; dr < condRows; dr++) {
      for (let dc = 0; dc < condCols; dc++) {
        const condRow = Math.min(cs.row, ce.row) + dr;
        const condCol = Math.min(cs.col, ce.col) + dc;
        const sumRow = Math.min(ss.row, se.row) + dr;
        const sumCol = Math.min(ss.col, se.col) + dc;

        const condVal = sheet.data[condRow]?.[condCol];
        if (this.matchesCriteria(condVal, op, critNum, critText)) {
          const sumVal = sheet.data[sumRow]?.[sumCol];
          if (typeof sumVal === 'number') sum += sumVal;
          else if (typeof sumVal === 'string' && !isNaN(Number(sumVal))) sum += Number(sumVal);
        }
      }
    }
    return sum;
  }

  private parseCriteria(criteria: string): { op: string; numVal: number | null; textVal: string } {
    let op = '=';
    let valStr = criteria;
    if (criteria.startsWith('>=')) { op = '>='; valStr = criteria.substring(2); }
    else if (criteria.startsWith('<=')) { op = '<='; valStr = criteria.substring(2); }
    else if (criteria.startsWith('<>') || criteria.startsWith('!=')) { op = '<>'; valStr = criteria.substring(2); }
    else if (criteria.startsWith('>')) { op = '>'; valStr = criteria.substring(1); }
    else if (criteria.startsWith('<')) { op = '<'; valStr = criteria.substring(1); }
    else if (criteria.startsWith('=')) { op = '='; valStr = criteria.substring(1); }

    const n = Number(valStr);
    return { op, numVal: isNaN(n) ? null : n, textVal: valStr };
  }

  private matchesCriteria(cellVal: string | number | boolean | null | undefined, op: string, critNum: number | null, critText: string): boolean {
    if (cellVal == null) return false;
    const num = typeof cellVal === 'number' ? cellVal : (typeof cellVal === 'string' && !isNaN(Number(cellVal)) ? Number(cellVal) : null);

    if (critNum !== null && num !== null) {
      switch (op) {
        case '=': return num === critNum;
        case '>': return num > critNum;
        case '<': return num < critNum;
        case '>=': return num >= critNum;
        case '<=': return num <= critNum;
        case '<>': return num !== critNum;
      }
    }
    // Text comparison
    const cellStr = String(cellVal).toLowerCase();
    const critLower = critText.toLowerCase();
    switch (op) {
      case '=': return cellStr === critLower;
      case '<>': return cellStr !== critLower;
      default: return false;
    }
  }

  /** Check if a cell has a formula */
  protected cellHasFormula(row: number, col: number): boolean {
    return this.xlsxCellFormulas.has(this.getCellFormatKey(row, col));
  }

  // ── Status bar computations ──

  /** Get sum of numeric values in the selected range */
  protected get xlsxStatusSum(): number | null {
    const r = this.xlsxRange;
    if (!r) return null;
    const sheet = this.activeSheet;
    if (!sheet) return null;

    let sum = 0;
    let hasNum = false;
    for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startCol; col <= r.endCol; col++) {
        const val = sheet.data[row]?.[col];
        if (typeof val === 'number') { sum += val; hasNum = true; }
        else if (typeof val === 'string' && val !== '' && !isNaN(Number(val))) { sum += Number(val); hasNum = true; }
      }
    }
    return hasNum ? sum : null;
  }

  /** Get count of non-empty cells in the selected range */
  protected get xlsxStatusCount(): number {
    const r = this.xlsxRange;
    if (!r) return 0;
    const sheet = this.activeSheet;
    if (!sheet) return 0;

    let count = 0;
    for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startCol; col <= r.endCol; col++) {
        const val = sheet.data[row]?.[col];
        if (val != null && val !== '') count++;
      }
    }
    return count;
  }

  /** Get average of numeric values in the selected range */
  protected get xlsxStatusAverage(): number | null {
    const r = this.xlsxRange;
    if (!r) return null;
    const sheet = this.activeSheet;
    if (!sheet) return null;

    let sum = 0;
    let count = 0;
    for (let row = r.startRow; row <= r.endRow; row++) {
      for (let col = r.startCol; col <= r.endCol; col++) {
        const val = sheet.data[row]?.[col];
        if (typeof val === 'number') { sum += val; count++; }
        else if (typeof val === 'string' && val !== '' && !isNaN(Number(val))) { sum += Number(val); count++; }
      }
    }
    return count > 0 ? sum / count : null;
  }

  // ── Formula range selection helpers ──

  /** Check if the current formula text has an open position for a range reference */
  private isInFormulaRangeMode(): boolean {
    const editing = this.editingCell();
    if (!editing) return false;
    const val = this.formulaBarValue();
    if (!val.startsWith('=')) return false;
    // Check if the last meaningful character is ( , : or an operator — means we expect a cell ref
    const trimmed = val.trimEnd();
    const lastChar = trimmed[trimmed.length - 1];
    return '(,;:+-*/'.includes(lastChar);
  }

  /** Insert or update the range reference in the formula text */
  private insertOrUpdateFormulaRange(startRow: number, startCol: number, endRow: number, endCol: number): void {
    const formula = this.formulaBarValue();
    const ref = startRow === endRow && startCol === endCol
      ? this.getCellRef(startRow, startCol)
      : `${this.getCellRef(startRow, startCol)}:${this.getCellRef(endRow, endCol)}`;

    let newFormula: string;
    if (this.xlsxFormulaInsertPos >= 0 && this.xlsxFormulaInsertEnd >= 0) {
      // Replace existing range ref we previously inserted
      newFormula = formula.substring(0, this.xlsxFormulaInsertPos) + ref + formula.substring(this.xlsxFormulaInsertEnd);
    } else {
      // First insertion — find the position
      this.xlsxFormulaInsertPos = formula.length;
      newFormula = formula + ref;
    }
    this.xlsxFormulaInsertEnd = this.xlsxFormulaInsertPos + ref.length;
    this.formulaBarValue.set(newFormula);

    // Update highlighted ranges to show the one being selected
    this.updateFormulaHighlights(newFormula);
  }

  /** Parse formula to extract all range references and build colored highlight rects */
  private updateFormulaHighlights(formula: string): void {
    if (!formula.startsWith('=')) {
      this.xlsxFormulaHighlights.set([]);
      return;
    }

    const highlights: { startRow: number; startCol: number; endRow: number; endCol: number; color: string }[] = [];
    // Match range references like A1:B5 or single cell refs like A1
    const rangeRegex = /([A-Za-z]+\d+):([A-Za-z]+\d+)/g;
    const singleRegex = /(?<![A-Za-z])([A-Za-z]+)(\d+)(?![\w:])/g;

    let matchIdx = 0;

    // First match ranges (A1:B5)
    let m: RegExpExecArray | null;
    while ((m = rangeRegex.exec(formula)) !== null) {
      const start = this.parseCellRef(m[1]);
      const end = this.parseCellRef(m[2]);
      if (start && end) {
        highlights.push({
          startRow: Math.min(start.row, end.row),
          startCol: Math.min(start.col, end.col),
          endRow: Math.max(start.row, end.row),
          endCol: Math.max(start.col, end.col),
          color: this.FORMULA_HIGHLIGHT_COLORS[matchIdx % this.FORMULA_HIGHLIGHT_COLORS.length]
        });
        matchIdx++;
      }
    }

    // Then match single cell refs that aren't part of a range (not preceded/followed by :)
    // We need to skip refs already captured in ranges
    const rangePositions = new Set<number>();
    rangeRegex.lastIndex = 0;
    while ((m = rangeRegex.exec(formula)) !== null) {
      for (let i = m.index; i < m.index + m[0].length; i++) rangePositions.add(i);
    }

    // Match function names to exclude them: look for refs NOT preceded by function context
    const funcNames = new Set(['SUM', 'SUMA', 'AVERAGE', 'PROMEDIO', 'COUNT', 'CONTAR', 'MIN', 'MAX', 'ABS',
      'ROUND', 'REDONDEAR', 'IF', 'SI', 'COUNTIF', 'SUMIF', 'SUMAR']);
    const refRegex = /\b([A-Z]+)(\d+)\b/gi;
    while ((m = refRegex.exec(formula)) !== null) {
      if (rangePositions.has(m.index)) continue;
      const letters = m[1].toUpperCase();
      if (funcNames.has(letters)) continue;
      const parsed = this.parseCellRef(m[0]);
      if (parsed) {
        highlights.push({
          startRow: parsed.row,
          startCol: parsed.col,
          endRow: parsed.row,
          endCol: parsed.col,
          color: this.FORMULA_HIGHLIGHT_COLORS[matchIdx % this.FORMULA_HIGHLIGHT_COLORS.length]
        });
        matchIdx++;
      }
    }

    this.xlsxFormulaHighlights.set(highlights);
  }

  /** Clear formula-range selection state */
  private clearFormulaRangeMode(): void {
    this.xlsxFormulaRangeMode = false;
    this.xlsxFormulaRangeDragging = false;
    this.xlsxFormulaRangeAnchor = null;
    this.xlsxFormulaRangeCurrent = null;
    this.xlsxFormulaInsertPos = -1;
    this.xlsxFormulaInsertEnd = -1;
  }

  /** Get the highlight color for a cell if it's in any formula highlight range */
  protected getCellHighlightColor(row: number, col: number): string | null {
    for (const h of this.xlsxFormulaHighlights()) {
      if (row >= h.startRow && row <= h.endRow && col >= h.startCol && col <= h.endCol) {
        return h.color;
      }
    }
    return null;
  }

  /** Get the highlight border style for a cell (outline edges of highlighted ranges) */
  protected getCellHighlightBorder(row: number, col: number): string | null {
    for (const h of this.xlsxFormulaHighlights()) {
      if (row >= h.startRow && row <= h.endRow && col >= h.startCol && col <= h.endCol) {
        const borders: string[] = [];
        if (row === h.startRow) borders.push(`border-top: 2px solid ${h.color}`);
        if (row === h.endRow) borders.push(`border-bottom: 2px solid ${h.color}`);
        if (col === h.startCol) borders.push(`border-left: 2px solid ${h.color}`);
        if (col === h.endCol) borders.push(`border-right: 2px solid ${h.color}`);
        return borders.join('; ');
      }
    }
    return null;
  }

  protected async onWorkbookExport(event: Event): Promise<void> {
    const select = event.target as HTMLSelectElement;
    const format = select.value as 'xlsx' | 'xls' | 'csv' | '';
    select.value = '';
    if (format) await this.downloadXlsxFile(format);
  }

  protected async downloadXlsxFile(format: 'xlsx' | 'xls' | 'csv' = 'xlsx'): Promise<void> {
    const note = this.noteService.selectedNote();
    if (!note || !this.xlsxBase64) return;
    const fileName = `${note.title || 'documento'}.${format}`;
    this.downloadService.showSaving(fileName);
    try {
      const result = await window.electronAPI.fileExportXlsxRaw({
        defaultName: fileName,
        base64: this.xlsxBase64,
        format,
        sheetIndex: this.activeSheetIndex(),
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
    const noteId = this.noteService.selectedNoteId();
    if (noteId === null) return;

    if (this.pendingSave) {
      clearTimeout(this.pendingSave.timer);
      if (this.pendingSave.noteId !== noteId) void this.flushPendingSave();
    }

    const snapshot = {
      noteId,
      title: this.title(),
      content: this.content(),
    };
    const timer = setTimeout(() => void this.flushPendingSave(), 500);
    this.pendingSave = { timer, ...snapshot };
  }

  private async flushPendingSave(): Promise<void> {
    const pending = this.pendingSave;
    if (!pending) return;
    clearTimeout(pending.timer);
    this.pendingSave = null;
    if (this.noteService.selectedNoteId() === pending.noteId) this.skipNextSync = true;
    await this.noteService.updateNote(pending.noteId, {
      title: pending.title,
      content: pending.content,
    });
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

  protected async printCurrentDocument(): Promise<void> {
    const note = this.noteService.selectedNote();
    if (!note || this.printing()) return;
    this.printing.set(true);

    try {
      let result: { success: boolean; canceled?: boolean; error?: string };
      if (note.file_type === 'pdf') {
        const fullNote = await window.electronAPI.notesGetById(note.id);
        if (!fullNote?.file_data) throw new Error('El PDF no contiene datos para imprimir.');
        result = await window.electronAPI.filePrint({
          title: this.title() || note.title || 'Documento PDF',
          pdfBase64: fullNote.file_data,
        });
      } else if (note.file_type === 'xlsx') {
        result = await window.electronAPI.filePrint({
          title: this.title() || note.title || 'Hoja de cálculo',
          htmlContent: this.buildPrintableSpreadsheet(),
          landscape: true,
        });
      } else if (note.file_type === 'pptx') {
        result = await window.electronAPI.filePrint({
          title: this.title() || note.title || 'Presentación',
          htmlContent: this.buildPrintablePresentation(),
          landscape: true,
        });
      } else {
        const title = this.title() || note.title || 'Nota';
        result = await window.electronAPI.filePrint({
          title,
          htmlContent: `<article class="print-note"><h1>${this.escapeXml(title)}</h1>${this.content() || ''}</article>`,
        });
      }

      if (!result.success && !result.canceled) {
        console.error('Error al imprimir:', result.error);
        window.alert(this.translate.instant('NOTE_EDITOR.PRINT_ERROR'));
      }
    } catch (error) {
      console.error('Error al preparar la impresión:', error);
      window.alert(this.translate.instant('NOTE_EDITOR.PRINT_ERROR'));
    } finally {
      this.printing.set(false);
    }
  }

  private buildPrintableSpreadsheet(): string {
    const sheet = this.activeSheet;
    const documentTitle = this.escapeXml(this.title() || 'Hoja de cálculo');
    if (!sheet) return `<h1>${documentTitle}</h1><p>${this.escapeXml(this.translate.instant('NOTE_EDITOR.NO_PRINTABLE_CONTENT'))}</p>`;

    let lastRow = -1;
    let lastColumn = -1;
    sheet.data.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
      if (value !== null && value !== '') {
        lastRow = Math.max(lastRow, rowIndex);
        lastColumn = Math.max(lastColumn, columnIndex);
      }
    }));
    const rowCount = Math.max(1, lastRow + 1);
    const columnCount = Math.max(1, lastColumn + 1);
    const headers = Array.from({ length: columnCount }, (_, index) =>
      `<th>${this.getColumnLabel(index)}</th>`,
    ).join('');
    const rows = Array.from({ length: rowCount }, (_, rowIndex) => {
      const cells = Array.from({ length: columnCount }, (_, columnIndex) => {
        const value = sheet.data[rowIndex]?.[columnIndex] ?? '';
        const format = this.xlsxCellFormats.get(`${this.activeSheetIndex()}:${rowIndex}:${columnIndex}`);
        const styles = [
          format?.bold ? 'font-weight:700' : '',
          format?.italic ? 'font-style:italic' : '',
          format?.underline ? 'text-decoration:underline' : '',
          format?.align ? `text-align:${format.align}` : typeof value === 'number' ? 'text-align:right' : '',
          format?.fontSize ? `font-size:${Math.max(8, Math.min(28, format.fontSize))}pt` : '',
        ].filter(Boolean).join(';');
        return `<td${styles ? ` style="${styles}"` : ''}>${this.escapeXml(String(value))}</td>`;
      }).join('');
      return `<tr><th>${rowIndex + 1}</th>${cells}</tr>`;
    }).join('');

    return `<h1 class="print-sheet-title">${documentTitle} — ${this.escapeXml(sheet.name)}</h1><table><thead><tr><th></th>${headers}</tr></thead><tbody>${rows}</tbody></table>`;
  }

  private buildPrintablePresentation(): string {
    const slides = this.presentationSlides();
    if (slides.length === 0) {
      return `<p>${this.escapeXml(this.translate.instant('NOTE_EDITOR.NO_PRINTABLE_CONTENT'))}</p>`;
    }
    return slides.map(slide => {
      const titleStyle = [
        `font-size:${slide.style.titleFontSize}pt`,
        `text-align:${slide.style.titleAlign}`,
        slide.style.titleBold ? 'font-weight:700' : '',
        slide.style.titleItalic ? 'font-style:italic' : '',
        slide.style.titleUnderline ? 'text-decoration:underline' : '',
      ].filter(Boolean).join(';');
      const bodyStyle = [
        `font-size:${slide.style.bodyFontSize}pt`,
        `text-align:${slide.style.bodyAlign}`,
        slide.style.bodyBold ? 'font-weight:700' : '',
        slide.style.bodyItalic ? 'font-style:italic' : '',
        slide.style.bodyUnderline ? 'text-decoration:underline' : '',
      ].filter(Boolean).join(';');
      const body = this.escapeXml(slide.body).replace(/\n/g, '<br>');
      return `<section class="print-slide" style="background:${slide.style.backgroundColor};color:${slide.style.textColor};border-top:8pt solid ${slide.style.accentColor}"><h1 style="${titleStyle}">${this.escapeXml(slide.title)}</h1>${slide.layout === 'title-only' ? '' : `<div class="print-slide-body" style="${bodyStyle}">${body}</div>`}</section>`;
    }).join('');
  }

  // ── Download as .pptx ──
  async downloadAsPptx(): Promise<void> {
    this.showDownloadMenu.set(false);
    const noteTitle = this.title() || 'nota';
    const fileName = `${this.sanitizeFilename(noteTitle)}.pptx`;
    this.downloadService.showSaving(fileName);
    try {
      const bytes = await this.presentationExport.generate(
        noteTitle,
        this.presentationBlocks(this.content()),
      );
      const result = await window.electronAPI.fileSaveAsBuffer({
        defaultName: fileName,
        filters: [{ name: this.translate.instant('NOTE_EDITOR.FILTER_POWERPOINT'), extensions: ['pptx'] }],
        buffer: Array.from(bytes),
      });
      if (result.success && result.filePath) {
        this.downloadService.showDone(fileName, result.filePath);
      } else {
        this.downloadService.dismiss(fileName);
      }
    } catch { this.downloadService.showError(fileName); }
  }

  private presentationBlocks(html: string): {
    kind: 'heading' | 'paragraph' | 'bullet' | 'number' | 'quote';
    text: string;
    level?: number;
  }[] {
    const root = document.createElement('div');
    root.innerHTML = html;
    const blocks: {
      kind: 'heading' | 'paragraph' | 'bullet' | 'number' | 'quote';
      text: string;
      level?: number;
    }[] = [];

    const add = (kind: 'heading' | 'paragraph' | 'bullet' | 'number' | 'quote', text: string, level?: number): void => {
      const normalized = text.replace(/\s+/g, ' ').trim();
      if (normalized) blocks.push({ kind, text: normalized, ...(level ? { level } : {}) });
    };

    const walk = (node: Node): void => {
      if (node.nodeType === Node.TEXT_NODE) {
        add('paragraph', node.textContent ?? '');
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;

      const element = node as HTMLElement;
      const tag = element.tagName.toLowerCase();
      if (/^h[1-3]$/.test(tag)) {
        add('heading', element.textContent ?? '', Number(tag[1]));
        return;
      }
      if (tag === 'ul' || tag === 'ol') {
        Array.from(element.children).filter(child => child.tagName.toLowerCase() === 'li').forEach((item, index) => {
          const text = item.textContent ?? '';
          add(tag === 'ol' ? 'number' : 'bullet', tag === 'ol' ? `${index + 1}. ${text}` : text);
        });
        return;
      }
      if (tag === 'blockquote') {
        add('quote', element.textContent ?? '');
        return;
      }
      if (tag === 'p' || tag === 'div') {
        const hasBlockChildren = Array.from(element.children).some(child =>
          /^(h[1-3]|p|div|ul|ol|blockquote|table)$/.test(child.tagName.toLowerCase()),
        );
        if (hasBlockChildren) element.childNodes.forEach(walk);
        else add('paragraph', element.textContent ?? '');
        return;
      }
      if (tag === 'table') {
        element.querySelectorAll('tr').forEach(row => {
          add('paragraph', Array.from(row.querySelectorAll('th, td')).map(cell => cell.textContent?.trim() ?? '').join('  |  '));
        });
        return;
      }
      element.childNodes.forEach(walk);
    };

    root.childNodes.forEach(walk);
    return blocks;
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

  // ── Download as .xlsx ──
  async downloadAsXlsx(): Promise<void> {
    this.showDownloadMenu.set(false);
    const noteTitle = this.title() || 'nota';
    const htmlContent = this.content() || '';
    const fileName = `${this.sanitizeFilename(noteTitle)}.xlsx`;
    this.downloadService.showSaving(fileName);
    try {
      const result = await window.electronAPI.fileExportXlsx({
        defaultName: fileName,
        htmlContent,
        title: noteTitle,
      });
      if (result.success && result.filePath) {
        this.downloadService.showDone(fileName, result.filePath);
      } else {
        this.downloadService.dismiss(fileName);
      }
    } catch { this.downloadService.showError(fileName); }
  }

  // ── Download as .json ──
  async downloadAsJson(): Promise<void> {
    this.showDownloadMenu.set(false);
    const note = this.noteService.selectedNote();
    if (!note) return;
    const noteTitle = this.title() || 'nota';
    const fileName = `${this.sanitizeFilename(noteTitle)}.json`;

    const exportData: Record<string, unknown> = {
      title: note.title,
      content: note.content,
      created_at: note.created_at,
      updated_at: note.updated_at,
    };

    // If the note has attached file data (e.g. PDF), include it
    if (note.file_type) {
      const fullNote = await window.electronAPI.notesGetById(note.id);
      if (fullNote?.file_data) {
        exportData['file_data'] = fullNote.file_data;
        exportData['file_type'] = fullNote.file_type;
      }
    }

    const jsonContent = JSON.stringify(exportData, null, 2);
    this.downloadService.showSaving(fileName);
    try {
      const result = await window.electronAPI.fileSaveAs({
        defaultName: fileName,
        filters: [{ name: this.translate.instant('NOTE_EDITOR.FILTER_JSON'), extensions: ['json'] }],
        content: jsonContent,
        encoding: 'utf-8',
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
