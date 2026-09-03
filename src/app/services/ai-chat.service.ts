import { Injectable, signal, computed, inject } from '@angular/core';
import { NoteService } from './note.service';
import { I18nService } from './i18n.service';
import { AssistantConfigService } from './assistant-config.service';
import { AssistantModelInfo, AssistantNoteType, AssistantProfile } from '../models/assistant.model';

export interface ChatMessage {
  id?: number;
  role: 'user' | 'assistant';
  content: string;
  timestamp: Date;
  isStreaming?: boolean;
  actions?: ChatAction[];
  displayContent?: string;
}

export interface ChatAction {
  id: string;
  type: 'edit_note' | 'replace_text' | 'create_note' | 'create_folder' | 'edit_excel';
  title?: string;
  content?: string;
  folderName?: string;
  cells?: { row: number; col: number; value: string | number | null }[];
  sheetIndex?: number;
  find?: string;
  replacement?: string;
  status: 'pending' | 'applied' | 'rejected';
  changePreview?: string;
}

export interface Conversation {
  id: number;
  title: string;
  model: string;
  assistant_id: string;
  created_at: string;
  updated_at: string;
}

// Action block regex: [ACTION:type]...[/ACTION] (also matches unclosed blocks at end of message)
const ACTION_REGEX = /\[ACTION:(edit_note|replace_text|create_note|create_folder|edit_excel)\]([\s\S]*?)(?:\[\/ACTION\]|$)/g;

@Injectable({ providedIn: 'root' })
/**
 * Orquesta el chat local entre Angular y los proveedores ejecutados por Electron.
 *
 * El servicio administra modelos, conversaciones persistidas, streaming,
 * mensajes en espera y el motor de acciones que permite a una respuesta de IA
 * proponer cambios verificables sobre notas, carpetas y hojas de cálculo.
 */
export class AiChatService {
  private readonly noteService = inject(NoteService);
  private readonly i18nService = inject(I18nService);
  readonly configService = inject(AssistantConfigService);

  readonly messages = signal<ChatMessage[]>([]);
  readonly models = signal<AssistantModelInfo[]>([]);
  readonly selectedModel = signal<string>('');
  readonly isStreaming = signal(false);
  readonly isThinking = signal(false);
  readonly providerAvailable = signal<boolean | null>(null);
  readonly providerError = signal('');
  readonly ollamaInstalled = signal<boolean | null>(null);
  readonly ollamaRunning = signal<boolean | null>(null);
  readonly ollamaStarting = signal(false);
  readonly ollamaStopping = signal(false);
  readonly queuedMessages = signal<{ id: string; content: string }[]>([]);
  private dispatching = false;

  // Conversations
  readonly conversations = signal<Conversation[]>([]);
  readonly activeConversationId = signal<number | null>(null);
  readonly showSessions = signal(false);

  private cleanupToken: (() => void) | null = null;
  private cleanupDone: (() => void) | null = null;
  private cleanupError: (() => void) | null = null;

  // Excel context: cached CSV representation of the active spreadsheet
  private xlsxContextData: string | null = null;

  readonly hasModels = computed(() => this.models().length > 0 || !!this.selectedModel());
  readonly ollamaAvailable = computed(() => this.ollamaInstalled() === true && this.ollamaRunning() === true);
  readonly activeConversation = computed(() =>
    this.conversations().find(c => c.id === this.activeConversationId()) ?? null
  );

  /** Inicializa perfiles, proveedor activo, modelos disponibles e historial. */
  async initialize(): Promise<void> {
    await this.configService.init();
    const assistant = this.configService.activeAssistant();
    this.selectedModel.set(assistant?.model ?? '');
    if (assistant?.provider === 'ollama') await this.checkOllamaStatus();
    else {
      await this.ensureManagedLlamaServer();
      await this.checkProviderStatus();
    }
    if (this.providerAvailable()) await this.loadModels();
    await this.loadConversations();
  }

  /** Selecciona un modelo y lo guarda dentro del perfil activo. */
  setSelectedModel(name: string): void {
    this.selectedModel.set(name);
    const assistant = this.configService.activeAssistant();
    if (assistant) void this.configService.updateAssistant(assistant.id, { model: name });
  }

  /** Cambia de perfil y vuelve a comprobar el runtime y sus modelos. */
  async selectAssistant(id: string): Promise<void> {
    await this.configService.setActiveAssistant(id);
    this.selectedModel.set(this.configService.activeAssistant()?.model ?? '');
    this.models.set([]);
    if (this.configService.activeAssistant().provider === 'ollama') await this.checkOllamaStatus();
    else {
      await this.ensureManagedLlamaServer();
      await this.checkProviderStatus();
    }
    if (this.providerAvailable()) await this.loadModels();
  }

  /** Proyecta el perfil activo al contrato mínimo entendido por el proceso principal. */
  private runtimeConfig() {
    const assistant = this.configService.activeAssistant();
    return {
      provider: assistant?.provider ?? 'ollama' as const,
      endpoint: assistant?.endpoint || 'http://127.0.0.1:11434',
      temperature: assistant?.temperature ?? 0.7,
    };
  }

  /** Arranca llama.cpp administrado cuando el perfil lo requiere y tiene un modelo local. */
  private async ensureManagedLlamaServer(): Promise<void> {
    const assistant = this.configService.activeAssistant();
    if (assistant.provider !== 'llamacpp' || !assistant.localModelPath) return;
    try {
      const status = await window.electronAPI.llamaManagerStatus(assistant.endpoint);
      if (status.installed && !status.running) {
        await window.electronAPI.llamaStart(assistant.localModelPath, assistant.endpoint);
      }
    } catch {
      // The regular provider status exposes a useful error to the UI.
    }
  }

  /** Comprueba que el endpoint configurado acepte solicitudes. */
  async checkProviderStatus(): Promise<void> {
    try {
      const result = await window.electronAPI.aiStatus(this.runtimeConfig());
      this.providerAvailable.set(result.available);
      this.providerError.set(result.error ?? '');
    } catch (error) {
      this.providerAvailable.set(false);
      this.providerError.set(error instanceof Error ? error.message : String(error));
    }
  }

  /** Comprueba instalación y proceso de Ollama, y sincroniza el estado genérico del proveedor. */
  async checkOllamaStatus(): Promise<void> {
    try {
      const { installed } = await window.electronAPI.ollamaCheckInstalled();
      this.ollamaInstalled.set(installed);
      if (installed) {
        const { running } = await window.electronAPI.ollamaStatus();
        this.ollamaRunning.set(running);
      } else {
        this.ollamaRunning.set(false);
      }
    } catch {
      this.ollamaInstalled.set(false);
      this.ollamaRunning.set(false);
    }
    await this.checkProviderStatus();
  }

  /** Solicita a Electron iniciar Ollama y recarga modelos al quedar disponible. */
  async startOllama(): Promise<void> {
    this.ollamaStarting.set(true);
    try {
      const result = await window.electronAPI.ollamaStart();
      if (result.success) {
        this.ollamaRunning.set(true);
        await this.loadModels();
      }
    } catch { /* ignore */ }
    this.ollamaStarting.set(false);
  }

  /** Detiene Ollama y limpia los modelos que ya no pueden utilizarse. */
  async stopOllama(): Promise<void> {
    this.ollamaStopping.set(true);
    try {
      await window.electronAPI.ollamaStop();
      this.ollamaRunning.set(false);
      this.providerAvailable.set(false);
      this.providerError.set('Ollama está detenido');
      this.models.set([]);
      this.setSelectedModel('');
    } catch { /* ignore */ }
    this.ollamaStopping.set(false);
  }

  /** Abre en el navegador la página oficial de descarga de Ollama. */
  openDownload(): void {
    window.electronAPI.ollamaOpenDownload();
  }

  /** Carga modelos compatibles con chat y conserva la selección persistida cuando existe. */
  async loadModels(): Promise<void> {
    try {
      const result = await window.electronAPI.aiListModels(this.runtimeConfig());
      if (result.success && result.models) {
        this.models.set(result.models);
        this.providerAvailable.set(true);
        const saved = this.selectedModel();
        const modelExists = saved && result.models.some(m => m.name === saved);
        if (!modelExists && !saved && result.models.length > 0) {
          this.setSelectedModel(result.models[0].name);
        }
      } else {
        this.providerError.set(result.error ?? 'No se pudieron cargar los modelos');
        if (this.selectedModel()) {
          this.models.set([{ name: this.selectedModel(), size: 0, modified_at: '' }]);
          this.providerAvailable.set(true);
        } else {
          this.providerAvailable.set(false);
        }
      }
    } catch {
      this.providerAvailable.set(false);
    }
  }

  // ── Conversations ──

  /** Recupera las conversaciones ordenadas por actividad desde SQLite. */
  async loadConversations(): Promise<void> {
    try {
      const list = await window.electronAPI.conversationsGetAll();
      this.conversations.set(list);
    } catch { /* ignore */ }
  }

  /** Crea y activa una conversación vacía asociada al perfil actual. */
  async newConversation(): Promise<void> {
    const model = this.selectedModel() || '';
    const conv = await window.electronAPI.conversationsCreate('Nueva conversación', model, this.configService.activeAssistant().id);
    this.conversations.update(list => [conv, ...list]);
    this.activeConversationId.set(conv.id);
    this.messages.set([]);
    this.showSessions.set(false);
  }

  /** Restaura mensajes, acciones, asistente y modelo de una conversación guardada. */
  async selectConversation(id: number): Promise<void> {
    this.activeConversationId.set(id);
    const dbMessages = await window.electronAPI.conversationsGetMessages(id);
    this.messages.set(
      dbMessages.map(m => {
        const msg: ChatMessage = {
          id: m.id,
          role: m.role as 'user' | 'assistant',
          content: m.content,
          timestamp: new Date(m.created_at),
        };
        // Re-parse actions from stored content
        if (msg.role === 'assistant') {
          this.parseActionsOnMessage(msg);
        }
        return msg;
      })
    );
    const conv = this.conversations().find(c => c.id === id);
    if (conv?.assistant_id && this.configService.assistants().some(item => item.id === conv.assistant_id)) {
      await this.selectAssistant(conv.assistant_id);
    }
    if (conv?.model) {
      this.setSelectedModel(conv.model);
    }
    this.showSessions.set(false);
  }

  /** Elimina una conversación y limpia la vista si era la sesión activa. */
  async deleteConversation(id: number): Promise<void> {
    await window.electronAPI.conversationsDelete(id);
    this.conversations.update(list => list.filter(c => c.id !== id));
    if (this.activeConversationId() === id) {
      this.activeConversationId.set(null);
      this.messages.set([]);
    }
  }

  // ── Actions (RPA Engine) ──

  /**
   * Serializa la hoja activa como TSV para entregar contexto compacto al modelo
   * sin confundir sus columnas con el formato CSV usado por las acciones.
   */
  private async loadXlsxContext(): Promise<void> {
    const note = this.noteService.selectedNote();
    if (!note || note.file_type !== 'xlsx') {
      this.xlsxContextData = null;
      return;
    }
    try {
      const noteData = await window.electronAPI.notesGetById(note.id);
      if (!noteData?.file_data) { this.xlsxContextData = null; return; }
      const result = await window.electronAPI.xlsxParseData(noteData.file_data);
      if (!result.success || !result.sheets) { this.xlsxContextData = null; return; }

      // Use tab-separated format (TSV) to avoid confusing the LLM
      // since the cell edit format (ROW,COL,VALUE) also uses commas.
      let tsv = '';
      for (const sheet of result.sheets) {
        tsv += `--- Sheet: ${sheet.name} ---\n`;
        // Column header reference line
        if (sheet.data.length > 0) {
          const maxCols = Math.max(...sheet.data.map(r => r.length));
          const colHeaders = Array.from({ length: maxCols }, (_, i) => String.fromCharCode(65 + i));
          tsv += `[Columns: ${colHeaders.join('\t')}]\n`;
        }
        for (let r = 0; r < sheet.data.length; r++) {
          const row = sheet.data[r];
          tsv += `Row${r}: ` + row.map(cell => {
            if (cell === null || cell === undefined) return '';
            return String(cell);
          }).join('\t') + '\n';
        }
        tsv += '\n';
      }
      this.xlsxContextData = tsv;
    } catch {
      this.xlsxContextData = null;
    }
  }

  /**
   * Construye el mensaje de sistema según idioma, perfil, tipo de nota y
   * permisos. Incluye sólo los formatos de acción habilitados para reducir
   * propuestas que la aplicación tendría que rechazar.
   */
  private buildSystemPrompt(): string {
    const note = this.noteService.selectedNote();
    const languageName = this.i18nService.getLanguageName();
    const assistant = this.configService.activeAssistant();
    const noteType: AssistantNoteType = note?.file_type === 'xlsx'
      ? 'spreadsheet'
      : note?.file_type === 'pptx' ? 'presentation'
      : note?.file_type === 'pdf' ? 'pdf' : 'text';
    const availableActions = assistant?.actions.filter(action => action.enabled && action.noteTypes.includes(noteType)) ?? [];
    const enabledModes = new Set(availableActions.map(action => action.mode));
    let prompt = `You are "${assistant?.name || 'Helper'}", a configurable assistant in the Helper Notes desktop application.\n`;
    prompt += `Your professional role is: ${assistant?.role || 'general assistant'}.\n`;
    if (assistant?.description) prompt += `Role description: ${assistant.description}\n`;
    if (assistant?.systemPrompt) prompt += `Custom instructions: ${assistant.systemPrompt}\n`;
    const noteInstructions = assistant?.noteTypeInstructions[noteType];
    if (noteInstructions) prompt += `Instructions for this file type: ${noteInstructions}\n`;
    prompt += `You MUST respond in **${languageName}**. Always use ${languageName} for your responses regardless of the language the user writes in.\n\n`;

    if (availableActions.length > 0) {
      prompt += `The following capabilities are enabled for this assistant and this file type:\n`;
      for (const action of availableActions) {
        prompt += `- ${action.name} (${action.mode}): ${action.description}. ${action.instructions}\n`;
      }
      prompt += `Do not use actions that are not in this list. Decide which enabled action fits the user's intent and the current file type.\n\n`;
    }

    if (note) {
      if (note.file_type === 'xlsx') {
        // Excel spreadsheet context
        prompt += `The user currently has a SPREADSHEET (Excel) open:\n`;
        prompt += `- Title: "${note.title || '(sin título)'}"\n`;
        prompt += `- Type: Excel spreadsheet (.xlsx)\n`;
        if (this.xlsxContextData) {
          prompt += `- Spreadsheet data (tab-separated, each row labeled with its 0-based index):\n`;
          prompt += `\`\`\`\n${this.xlsxContextData.substring(0, 4000)}\n\`\`\`\n\n`;
        } else {
          prompt += `- (Spreadsheet data not yet loaded)\n\n`;
        }
      } else if (note.file_type === 'pptx') {
        let presentationText = note.content;
        try {
          const document = JSON.parse(note.content) as { slides?: { title?: string; body?: string; notes?: string }[] };
          presentationText = (document.slides ?? []).map((slide, index) =>
            `Slide ${index + 1}: ${slide.title ?? ''}\n${slide.body ?? ''}\nSpeaker notes: ${slide.notes ?? ''}`,
          ).join('\n\n');
        } catch { /* Keep the raw content as a fallback. */ }
        prompt += `The user currently has a POWERPOINT PRESENTATION open:\n`;
        prompt += `- Title: "${note.title || '(sin título)'}"\n`;
        prompt += `- Slides and speaker notes:\n\`\`\`\n${presentationText.substring(0, 6000)}\n\`\`\`\n\n`;
      } else {
        const plainText = this.stripHtml(note.content);
        const htmlContent = note.content || '';
        prompt += `The user currently has this note open:\n`;
        prompt += `- Title: "${note.title || '(sin título)'}"\n`;
        prompt += `- Plain text preview:\n\`\`\`\n${plainText.substring(0, 2000)}\n\`\`\`\n`;
        prompt += `- Original HTML content (USE THIS as the base when editing):\n\`\`\`html\n${htmlContent.substring(0, 3000)}\n\`\`\`\n\n`;
      }
    } else {
      prompt += `The user does not have any note selected.\n\n`;
    }

    prompt += `You can propose enabled actions using the special blocks below. ONLY use them when the user explicitly asks to modify, create, or edit something.\n\n`;
    prompt += `ACTION FORMAT (you MUST follow this exact format):\n\n`;
    if (enabledModes.has('edit_note')) {
      prompt += `To edit the current note:\n[ACTION:edit_note]\ntitle: Title Here\ncontent:\n<p>Full HTML content here</p>\n[/ACTION]\n\n`;
    }
    if (enabledModes.has('replace_text')) {
      prompt += `To replace an exact text fragment in the current note:\n[ACTION:replace_text]\nfind:\nExact original text\nreplacement:\nReplacement text\n[/ACTION]\n\n`;
    }
    if (enabledModes.has('create_note')) {
      prompt += `To create a new note:\n[ACTION:create_note]\ntitle: Title Here\ncontent:\n<p>Full HTML content here</p>\n[/ACTION]\n\n`;
    }
    if (enabledModes.has('create_folder')) {
      prompt += `To create a folder:\n[ACTION:create_folder]\nname: Folder Name\n[/ACTION]\n\n`;
    }

    // Add Excel-specific action format
    if (note?.file_type === 'xlsx' && enabledModes.has('edit_excel')) {
      prompt += `IMPORTANT: The user has a SPREADSHEET open. Use [ACTION:edit_excel] to modify it. NEVER use [ACTION:edit_note] for spreadsheets.\n\n`;
      prompt += `To edit cells in the current spreadsheet:\n[ACTION:edit_excel]\nsheet: 0\ncells:\nROW,COL,VALUE\nROW,COL,VALUE\n[/ACTION]\n\n`;
      prompt += `EXCEL FORMAT RULES:\n`;
      prompt += `- Each line under "cells:" represents ONE SINGLE CELL: ROW,COL,VALUE\n`;
      prompt += `- ROW and COL are 0-based indexes. A1 = 0,0 | B1 = 0,1 | A2 = 1,0 | C3 = 2,2\n`;
      prompt += `- Column mapping: A=0, B=1, C=2, D=3, E=4, F=5, G=6...\n`;
      prompt += `- VALUE is the content of ONE cell only — a single text or a single number.\n`;
      prompt += `- sheet: is the sheet index (0-based), default is 0\n`;
      prompt += `- **CRITICAL**: NEVER put multiple data values on a single line. If you want to write "Alquiler" in A2 and 800 in B2, write TWO separate lines:\n`;
      prompt += `  1,0,Alquiler\n`;
      prompt += `  1,1,800\n`;
      prompt += `- WRONG (do NOT do this): 1,0,Alquiler,800 ← This puts the text "Alquiler,800" in cell A2\n`;
      prompt += `- RIGHT (do this): 1,0,Alquiler then on next line 1,1,800 ← This puts "Alquiler" in A2 and 800 in B2\n\n`;
      prompt += `EXCEL BEHAVIOR RULES (CRITICAL):\n`;
      prompt += `1. **ALWAYS generate COMPLETE data**. When asked to "add a column", create the HEADER in row 0 AND fill several example rows with realistic data below it.\n`;
      prompt += `2. When asked to "create a table" or "add data", generate a HEADER ROW (row 0) with column names, and AT LEAST 5-10 rows of realistic example data.\n`;
      prompt += `3. When the spreadsheet already has data, PRESERVE existing data. Only add/modify the cells the user requested. Look at the CSV data above to see what already exists.\n`;
      prompt += `4. When adding a new column to existing data, place the header in the SAME ROW as existing headers, and fill data rows to match existing row count.\n`;
      prompt += `5. Be generous with data. Users expect spreadsheets to have meaningful content, not just a single cell.\n`;
      prompt += `6. For numeric columns (prices, quantities, amounts), use realistic numbers.\n`;
      prompt += `7. If the user asks to "change the title/name", use [ACTION:edit_note] with ONLY a title field (no content needed).\n\n`;
      prompt += `EXAMPLE 1 - Creating an expense table from scratch:\n`;
      prompt += `[ACTION:edit_excel]\nsheet: 0\ncells:\n0,0,Concepto\n0,1,Monto\n0,2,Fecha\n0,3,Categoría\n1,0,Alquiler\n1,1,800\n1,2,01/03/2026\n1,3,Vivienda\n2,0,Supermercado\n2,1,150\n2,2,02/03/2026\n2,3,Alimentación\n3,0,Gasolina\n3,1,60\n3,2,03/03/2026\n3,3,Transporte\n4,0,Internet\n4,1,45\n4,2,05/03/2026\n4,3,Servicios\n5,0,Electricidad\n5,1,90\n5,2,05/03/2026\n5,3,Servicios\n[/ACTION]\n\n`;
      prompt += `EXAMPLE 2 - Adding a column to existing data (existing data has 5 rows with headers in row 0):\n`;
      prompt += `[ACTION:edit_excel]\nsheet: 0\ncells:\n0,3,Total\n1,3,500\n2,3,300\n3,3,150\n4,3,200\n5,3,450\n[/ACTION]\n\n`;
    }

    if (enabledModes.has('create_note')) {
      prompt += `EXAMPLE of creating a note with content:\n[ACTION:create_note]\ntitle: Shopping List\ncontent:\n<h2>Groceries</h2>\n<ul>\n<li>Milk</li>\n<li>Bread</li>\n<li>Eggs</li>\n</ul>\n[/ACTION]\n\n`;
    }
    if (enabledModes.has('edit_note')) {
      prompt += `EXAMPLE of editing the current note to add content:\n[ACTION:edit_note]\ntitle: My Note\ncontent:\n<p>Original content here...</p>\n<p>New content added here...</p>\n[/ACTION]\n\n`;
    }
    prompt += `CRITICAL RULES:\n`;
    prompt += `1. For create_note and edit_note, content is mandatory and must contain the complete HTML content.\n`;
    prompt += `2. Prefer edit_note for broad edits and replace_text for a precise substitution when those actions are enabled.\n`;
    prompt += `3. Only use create_note when the user explicitly requests a separate note.\n`;
    prompt += `4. Never emit an action disabled by the active assistant or unavailable for the current file type.\n`;
    prompt += `5. **PRESERVE EXISTING CONTENT**: When editing, retain everything the user did not request to change.\n`;
    prompt += `6. If the user asks to "add something at the end", copy ALL original HTML first, then append new content.\n`;
    prompt += `7. Use action blocks ONLY when the user asks for a change/creation. Otherwise respond normally.\n`;
    prompt += `8. Always generate content in HTML format using <p>, <b>, <i>, <u>, <ul>, <li>, <h1>, <h2>, <h3>, <br> tags.\n`;
    prompt += `9. **ALWAYS close action blocks with [/ACTION]**. Every [ACTION:...] MUST have a matching [/ACTION].\n`;
    prompt += `10. Put ALL content ONLY inside the action block. Do NOT repeat content outside it.\n`;
    prompt += `11. Keep the original note title unless the user asks to change it.\n`;

    return prompt;
  }

  /**
   * Extrae bloques `[ACTION:...]` de una respuesta, valida su estructura y
   * separa el texto visible de las operaciones pendientes de aprobación.
   */
  private parseActionsOnMessage(msg: ChatMessage): void {
    const actions: ChatAction[] = [];
    let cleanContent = msg.content;
    let actionIndex = 0;

    cleanContent = cleanContent.replace(ACTION_REGEX, (_match, type: string, body: string) => {
      const action: ChatAction = {
        id: `${msg.id ?? 'tmp'}-${actionIndex++}`,
        type: type as ChatAction['type'],
        status: 'pending',
      };

      const lines = body.trim().split('\n');
      if (type === 'create_folder') {
        for (const line of lines) {
          const nameMatch = line.match(/^name:\s*(.+)/i);
          if (nameMatch) action.folderName = nameMatch[1].trim();
        }
      } else if (type === 'replace_text') {
        const findMatch = body.match(/^find:\s*\n([\s\S]*?)(?=^replacement:)/mi);
        const replacementMatch = body.match(/^replacement:\s*\n([\s\S]*)$/mi);
        action.find = findMatch?.[1]?.trim() ?? '';
        action.replacement = replacementMatch?.[1]?.trim() ?? '';
        action.changePreview = action.find
          ? `− ${action.find}\n+ ${action.replacement ?? ''}`
          : undefined;
      } else if (type === 'edit_excel') {
        // Parse sheet index and cell data
        action.sheetIndex = 0;
        action.cells = [];
        let inCells = false;
        for (const line of lines) {
          if (inCells) {
            // Parse ROW,COL,VALUE format
            const cellMatch = line.match(/^(\d+)\s*,\s*(\d+)\s*,\s*(.*)/);
            if (cellMatch) {
              const row = parseInt(cellMatch[1], 10);
              const col = parseInt(cellMatch[2], 10);
              const rawVal = cellMatch[3].trim();
              let value: string | number | null = rawVal;
              if (rawVal === '') {
                value = null;
              } else if (!isNaN(Number(rawVal)) && rawVal !== '') {
                value = Number(rawVal);
              }
              action.cells.push({ row, col, value });
            }
            continue;
          }
          const sheetMatch = line.match(/^sheet:\s*(\d+)/i);
          if (sheetMatch) {
            action.sheetIndex = parseInt(sheetMatch[1], 10);
            continue;
          }
          if (/^cells:\s*$/i.test(line)) {
            inCells = true;
            continue;
          }
        }
        // Build a preview of the cell changes
        if (action.cells.length > 0) {
          const preview = action.cells.slice(0, 10).map(c => {
            const colLetter = String.fromCharCode(65 + c.col);
            return `${colLetter}${c.row + 1} = ${c.value ?? ''}`;
          }).join(', ');
          action.changePreview = action.cells.length > 10
            ? `${preview}... (+${action.cells.length - 10} more)`
            : preview;
        }
      } else {
        let titleFound = false;
        let inContent = false;
        const contentLines: string[] = [];
        for (const line of lines) {
          if (inContent) {
            contentLines.push(line);
            continue;
          }
          const titleMatch = line.match(/^title:\s*(.+)/i);
          if (titleMatch) {
            action.title = titleMatch[1].trim();
            titleFound = true;
            continue;
          }
          if (/^content:\s*$/i.test(line) || /^content:\s*(.+)/i.test(line)) {
            inContent = true;
            const inlineContent = line.match(/^content:\s*(.+)/i);
            if (inlineContent) contentLines.push(inlineContent[1]);
            continue;
          }
          // If title was already found and this line isn't title: or content:,
          // treat it as content (model may have omitted 'content:' prefix)
          if (titleFound && line.trim()) {
            inContent = true;
            contentLines.push(line);
          }
        }
        action.content = contentLines.join('\n').trim();
      }

      actions.push(action);

      // Compute change preview for edit_note actions
      if (action.type === 'edit_note' && action.content) {
        this.computeChangePreview(action);
      }

      return ''; // Remove the action block from display content
    });

    if (actions.length > 0) {
      msg.actions = actions;
      msg.displayContent = cleanContent.trim();
    }
  }

  /** Calcula una vista previa simple de las líneas añadidas por una edición. */
  private computeChangePreview(action: ChatAction): void {
    const note = this.noteService.selectedNote();
    if (!note?.content || !action.content) return;

    const originalText = this.stripHtml(note.content).trim();
    const newText = this.stripHtml(action.content).trim();

    if (!originalText) {
      action.changePreview = newText.substring(0, 500);
      return;
    }

    // Line-based diff: find lines that are new or changed
    const originalLines = new Set(
      originalText.split('\n').map(l => l.trim()).filter(l => l)
    );
    const newLines = newText.split('\n').map(l => l.trim()).filter(l => l);
    const addedLines = newLines.filter(l => !originalLines.has(l));

    if (addedLines.length > 0) {
      action.changePreview = addedLines.join('\n');
    }
  }

  /**
   * Ejecuta una acción aprobada sólo si el perfil y el tipo de nota aún la
   * permiten; después refresca la entidad modificada y actualiza su estado UI.
   */
  async applyAction(messageId: number | undefined, actionId: string): Promise<void> {
    const msgs = this.messages();
    const msgIndex = msgs.findIndex(m => m.id === messageId);
    if (msgIndex === -1) return;

    const msg = msgs[msgIndex];
    const action = msg.actions?.find(a => a.id === actionId);
    if (!action || action.status !== 'pending') return;

    const note = this.noteService.selectedNote();
    const noteType: AssistantNoteType = note?.file_type === 'xlsx'
      ? 'spreadsheet'
      : note?.file_type === 'pptx' ? 'presentation'
      : note?.file_type === 'pdf' ? 'pdf' : 'text';
    const actionMode = action.type === 'replace_text' ? 'replace_text' : action.type;
    const allowed = this.configService.activeAssistant().actions.some(definition =>
      definition.enabled && definition.mode === actionMode && definition.noteTypes.includes(noteType)
    );
    if (!allowed) {
      action.status = 'rejected';
      this.messages.set([...msgs]);
      return;
    }

    try {
      switch (action.type) {
        case 'edit_note': {
          const note = this.noteService.selectedNote();
          if (!note) throw new Error('No note selected');
          const updateData: { title?: string; content?: string } = {};
          if (action.title) updateData.title = action.title;
          if (action.content) updateData.content = this.ensureHtml(action.content);
          await this.noteService.updateNote(note.id, updateData);
          // Force refresh so the editor reflects the new content
          await this.noteService.refreshCurrentNote();
          break;
        }
        case 'replace_text': {
          const note = this.noteService.selectedNote();
          if (!note || note.file_type) throw new Error('No text note selected');
          if (!action.find) throw new Error('No source text provided');
          const updatedContent = this.replaceTextInHtml(note.content, action.find, action.replacement ?? '');
          if (updatedContent === null) throw new Error('Source text was not found');
          await this.noteService.updateNote(note.id, { content: updatedContent });
          await this.noteService.refreshCurrentNote();
          break;
        }
        case 'create_note': {
          const newNote = await this.noteService.createNote();
          const updateData: { title?: string; content?: string } = {};
          if (action.title) updateData.title = action.title;
          if (action.content) updateData.content = this.ensureHtml(action.content);
          await this.noteService.updateNote(newNote.id, updateData);
          // Force refresh so the editor reflects the new content
          await this.noteService.refreshCurrentNote();
          break;
        }
        case 'create_folder': {
          if (action.folderName) {
            await this.noteService.createFolder(action.folderName);
          }
          break;
        }
        case 'edit_excel': {
          const note = this.noteService.selectedNote();
          if (!note || note.file_type !== 'xlsx') throw new Error('No spreadsheet selected');
          if (!action.cells || action.cells.length === 0) throw new Error('No cells to update');

          // Get current file_data
          const noteData = await window.electronAPI.notesGetById(note.id);
          if (!noteData?.file_data) throw new Error('No spreadsheet data');

          const result = await window.electronAPI.xlsxBatchUpdateCells(
            noteData.file_data,
            action.sheetIndex ?? 0,
            action.cells
          );
          if (!result.success || !result.base64) throw new Error(result.error || 'Failed to update cells');

          // Save the updated file data
          await window.electronAPI.notesUpdate(note.id, { file_data: result.base64 });

          // Force refresh so the spreadsheet view reloads with updated data
          await this.noteService.refreshCurrentNote();
          break;
        }
      }
      action.status = 'applied';
    } catch {
      action.status = 'rejected';
    }

    // Update the message to trigger UI refresh
    this.messages.update(msgs => {
      const updated = [...msgs];
      updated[msgIndex] = { ...msg };
      return updated;
    });
  }

  /** Marca una propuesta como rechazada sin alterar datos persistidos. */
  rejectAction(messageId: number | undefined, actionId: string): void {
    const msgs = this.messages();
    const msgIndex = msgs.findIndex(m => m.id === messageId);
    if (msgIndex === -1) return;

    const msg = msgs[msgIndex];
    const action = msg.actions?.find(a => a.id === actionId);
    if (!action || action.status !== 'pending') return;

    action.status = 'rejected';
    this.messages.update(msgs => {
      const updated = [...msgs];
      updated[msgIndex] = { ...msg };
      return updated;
    });
  }

  private stripHtml(html: string): string {
    if (!html) return '';
    const div = document.createElement('div');
    div.innerHTML = html;
    return div.textContent || div.innerText || '';
  }

  private ensureHtml(content: string): string {
    // If content already has HTML tags, return as-is
    if (/<[a-z][\s\S]*>/i.test(content)) return content;
    // Otherwise wrap lines in <p> tags
    return content
      .split('\n')
      .filter(line => line.trim() !== '')
      .map(line => `<p>${line}</p>`)
      .join('');
  }

  /** Reemplaza texto únicamente en nodos de texto para preservar la estructura HTML. */
  private replaceTextInHtml(html: string, find: string, replacement: string): string | null {
    const container = document.createElement('div');
    container.innerHTML = html;
    const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
    let changed = false;
    let node = walker.nextNode();
    while (node) {
      if (node.nodeValue?.includes(find)) {
        node.nodeValue = node.nodeValue.split(find).join(replacement);
        changed = true;
      }
      node = walker.nextNode();
    }
    return changed ? container.innerHTML : null;
  }

  // ── Send message ──

  /**
   * Persiste el mensaje, prepara el contexto y conecta listeners de streaming.
   * Si ya existe una generación activa, encola el texto para procesarlo luego.
   */
  async sendMessage(content: string): Promise<void> {
    const trimmed = content.trim();
    if (!trimmed || !this.selectedModel()) return;
    if (this.isStreaming() || this.dispatching) {
      this.queuedMessages.update(queue => [
        ...queue,
        { id: `queued-${Date.now()}-${queue.length}`, content: trimmed },
      ]);
      return;
    }
    this.dispatching = true;
    try {

    // Load Excel context if the active note is a spreadsheet
    await this.loadXlsxContext();

    if (!this.activeConversationId()) {
      const title = trimmed.substring(0, 50) || 'Nueva conversación';
      const conv = await window.electronAPI.conversationsCreate(title, this.selectedModel(), this.configService.activeAssistant().id);
      this.conversations.update(list => [conv, ...list]);
      this.activeConversationId.set(conv.id);
    }

    const convId = this.activeConversationId()!;

    const userMsg = await window.electronAPI.conversationsAddMessage(convId, 'user', trimmed);
    this.messages.update(msgs => [
      ...msgs,
      { id: userMsg.id, role: 'user', content: trimmed, timestamp: new Date() },
    ]);

    const assistantMsg = await window.electronAPI.conversationsAddMessage(convId, 'assistant', '');
    this.messages.update(msgs => [
      ...msgs,
      { id: assistantMsg.id, role: 'assistant', content: '', timestamp: new Date(), isStreaming: true },
    ]);

    this.isStreaming.set(true);
    this.isThinking.set(true);

    const conv = this.conversations().find(c => c.id === convId);
    if (conv && conv.title === 'Nueva conversación') {
      const newTitle = trimmed.substring(0, 50);
      await window.electronAPI.conversationsUpdate(convId, { title: newTitle });
      this.conversations.update(list =>
        list.map(c => c.id === convId ? { ...c, title: newTitle } : c)
      );
    }

    await window.electronAPI.conversationsUpdate(convId, {
      model: this.selectedModel(),
      assistant_id: this.configService.activeAssistant().id,
    });

    // Build history with system prompt
    const systemPrompt = this.buildSystemPrompt();
    const history: { role: string; content: string }[] = [
      { role: 'system', content: systemPrompt },
    ];
    // Add only user/assistant messages (exclude system)
    for (const m of this.messages()) {
      if (m.role === 'user' || m.role === 'assistant') {
        history.push({ role: m.role, content: m.content });
      }
    }

    // Setup streaming listeners
    this.cleanupToken = window.electronAPI.onAiToken((token: string) => {
      this.isThinking.set(false);
      this.messages.update(msgs => {
        const updated = [...msgs];
        const last = updated[updated.length - 1];
        if (last?.role === 'assistant') {
          updated[updated.length - 1] = { ...last, content: last.content + token };
        }
        return updated;
      });
    });

    this.cleanupDone = window.electronAPI.onAiDone(() => {
      this.isThinking.set(false);
      const msgs = this.messages();
      const last = msgs[msgs.length - 1];
      if (last?.role === 'assistant' && last.id) {
        window.electronAPI.conversationsUpdateMessage(last.id, last.content);
      }
      this.messages.update(msgs => {
        const updated = [...msgs];
        const lastMsg = updated[updated.length - 1];
        if (lastMsg?.role === 'assistant') {
          const parsed = { ...lastMsg, isStreaming: false };
          this.parseActionsOnMessage(parsed);
          updated[updated.length - 1] = parsed;
        }
        return updated;
      });
      this.isStreaming.set(false);
      this.removeListeners();
      void this.processNextQueuedMessage();
    });

    this.cleanupError = window.electronAPI.onAiError((error: string) => {
      this.isThinking.set(false);
      const errorContent = `⚠️ Error: ${error}`;
      const msgs = this.messages();
      const last = msgs[msgs.length - 1];
      if (last?.role === 'assistant' && last.id) {
        window.electronAPI.conversationsUpdateMessage(last.id, errorContent);
      }
      this.messages.update(msgs => {
        const updated = [...msgs];
        const lastMsg = updated[updated.length - 1];
        if (lastMsg?.role === 'assistant') {
          updated[updated.length - 1] = {
            ...lastMsg,
            content: errorContent,
            isStreaming: false,
          };
        }
        return updated;
      });
      this.isStreaming.set(false);
      this.removeListeners();
      void this.processNextQueuedMessage();
    });

    await window.electronAPI.aiChat(this.runtimeConfig(), this.selectedModel(), history);
    } finally {
      this.dispatching = false;
      if (!this.isStreaming()) void this.processNextQueuedMessage();
    }
  }

  /** Cancela la generación, conserva el contenido parcial y continúa con la cola. */
  async abort(): Promise<void> {
    await window.electronAPI.aiAbort();
    this.isStreaming.set(false);
    this.isThinking.set(false);
    this.messages.update(msgs => {
      const updated = [...msgs];
      const last = updated[updated.length - 1];
      if (last?.role === 'assistant' && last.isStreaming) {
        const cancelledContent = last.content + ' [cancelado]';
        if (last.id) {
          window.electronAPI.conversationsUpdateMessage(last.id, cancelledContent);
        }
        updated[updated.length - 1] = { ...last, content: cancelledContent, isStreaming: false };
      }
      return updated;
    });
    this.removeListeners();
    await this.processNextQueuedMessage();
  }

  /** Retira un mensaje específico antes de que sea enviado. */
  removeQueuedMessage(id: string): void {
    this.queuedMessages.update(queue => queue.filter(item => item.id !== id));
  }

  /** Descarta todos los mensajes pendientes. */
  clearQueue(): void {
    this.queuedMessages.set([]);
  }

  /** Envía el siguiente elemento de la cola cuando el proveedor vuelve a estar libre. */
  private async processNextQueuedMessage(): Promise<void> {
    if (this.isStreaming() || this.dispatching) return;
    const next = this.queuedMessages()[0];
    if (!next) return;
    this.queuedMessages.update(queue => queue.slice(1));
    await this.sendMessage(next.content);
  }

  /** Elimina la conversación activa y reinicia el historial visible. */
  clearHistory(): void {
    if (this.activeConversationId()) {
      this.deleteConversation(this.activeConversationId()!);
    }
    this.messages.set([]);
    this.activeConversationId.set(null);
  }

  /** Desregistra listeners IPC para evitar duplicados y fugas entre generaciones. */
  private removeListeners(): void {
    this.cleanupToken?.();
    this.cleanupDone?.();
    this.cleanupError?.();
    this.cleanupToken = null;
    this.cleanupDone = null;
    this.cleanupError = null;
  }
}
