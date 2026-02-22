import { Injectable, signal, computed, inject } from '@angular/core';
import { NoteService } from './note.service';
import { I18nService } from './i18n.service';

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
  type: 'edit_note' | 'create_note' | 'create_folder';
  title?: string;
  content?: string;
  folderName?: string;
  status: 'pending' | 'applied' | 'rejected';
  changePreview?: string;
}

export interface OllamaModel {
  name: string;
  size: number;
  modified_at: string;
}

export interface Conversation {
  id: number;
  title: string;
  model: string;
  created_at: string;
  updated_at: string;
}

// Action block regex: [ACTION:type]...[/ACTION] (also matches unclosed blocks at end of message)
const ACTION_REGEX = /\[ACTION:(edit_note|create_note|create_folder)\]([\s\S]*?)(?:\[\/ACTION\]|$)/g;

@Injectable({ providedIn: 'root' })
export class AiChatService {
  private readonly noteService = inject(NoteService);
  private readonly i18nService = inject(I18nService);

  readonly messages = signal<ChatMessage[]>([]);
  readonly models = signal<OllamaModel[]>([]);
  readonly selectedModel = signal<string>(localStorage.getItem('ai-selected-model') || '');
  readonly isStreaming = signal(false);
  readonly isThinking = signal(false);
  readonly ollamaInstalled = signal<boolean | null>(null);
  readonly ollamaRunning = signal<boolean | null>(null);
  readonly ollamaStarting = signal(false);
  readonly ollamaStopping = signal(false);

  // Conversations
  readonly conversations = signal<Conversation[]>([]);
  readonly activeConversationId = signal<number | null>(null);
  readonly showSessions = signal(false);

  private cleanupToken: (() => void) | null = null;
  private cleanupDone: (() => void) | null = null;
  private cleanupError: (() => void) | null = null;

  readonly hasModels = computed(() => this.models().length > 0);
  readonly ollamaAvailable = computed(() => this.ollamaInstalled() === true && this.ollamaRunning() === true);
  readonly activeConversation = computed(() =>
    this.conversations().find(c => c.id === this.activeConversationId()) ?? null
  );

  setSelectedModel(name: string): void {
    this.selectedModel.set(name);
    if (name) {
      localStorage.setItem('ai-selected-model', name);
    } else {
      localStorage.removeItem('ai-selected-model');
    }
  }

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
  }

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

  async stopOllama(): Promise<void> {
    this.ollamaStopping.set(true);
    try {
      await window.electronAPI.ollamaStop();
      this.ollamaRunning.set(false);
      this.models.set([]);
      this.setSelectedModel('');
    } catch { /* ignore */ }
    this.ollamaStopping.set(false);
  }

  openDownload(): void {
    window.electronAPI.ollamaOpenDownload();
  }

  async loadModels(): Promise<void> {
    try {
      const result = await window.electronAPI.ollamaListModels();
      if (result.success && result.models) {
        this.models.set(result.models);
        this.ollamaRunning.set(true);
        const saved = this.selectedModel();
        const modelExists = saved && result.models.some(m => m.name === saved);
        if (!modelExists && result.models.length > 0) {
          this.setSelectedModel(result.models[0].name);
        }
      } else {
        this.ollamaRunning.set(false);
      }
    } catch {
      this.ollamaRunning.set(false);
    }
  }

  // ── Conversations ──

  async loadConversations(): Promise<void> {
    try {
      const list = await window.electronAPI.conversationsGetAll();
      this.conversations.set(list);
    } catch { /* ignore */ }
  }

  async newConversation(): Promise<void> {
    const model = this.selectedModel() || '';
    const conv = await window.electronAPI.conversationsCreate('Nueva conversación', model);
    this.conversations.update(list => [conv, ...list]);
    this.activeConversationId.set(conv.id);
    this.messages.set([]);
    this.showSessions.set(false);
  }

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
    if (conv?.model) {
      this.setSelectedModel(conv.model);
    }
    this.showSessions.set(false);
  }

  async deleteConversation(id: number): Promise<void> {
    await window.electronAPI.conversationsDelete(id);
    this.conversations.update(list => list.filter(c => c.id !== id));
    if (this.activeConversationId() === id) {
      this.activeConversationId.set(null);
      this.messages.set([]);
    }
  }

  // ── Actions (RPA Engine) ──

  private buildSystemPrompt(): string {
    const note = this.noteService.selectedNote();
    const languageName = this.i18nService.getLanguageName();
    let prompt = `You are a helpful AI assistant for a notes application called "Helper Notes". You MUST respond in **${languageName}**. Always use ${languageName} for your responses regardless of the language the user writes in.\n\n`;

    if (note) {
      const plainText = this.stripHtml(note.content);
      const htmlContent = note.content || '';
      prompt += `The user currently has this note open:\n`;
      prompt += `- Title: "${note.title || '(sin título)'}"\n`;
      prompt += `- Plain text preview:\n\`\`\`\n${plainText.substring(0, 2000)}\n\`\`\`\n`;
      prompt += `- Original HTML content (USE THIS as the base when editing):\n\`\`\`html\n${htmlContent.substring(0, 3000)}\n\`\`\`\n\n`;
    } else {
      prompt += `The user does not have any note selected.\n\n`;
    }

    prompt += `You can propose actions on notes using these special blocks. ONLY use them when the user explicitly asks to modify, create, or edit something.\n\n`;
    prompt += `ACTION FORMAT (you MUST follow this exact format):\n\n`;
    prompt += `To edit the current note:\n[ACTION:edit_note]\ntitle: Title Here\ncontent:\n<p>Full HTML content here</p>\n[/ACTION]\n\n`;
    prompt += `To create a new note:\n[ACTION:create_note]\ntitle: Title Here\ncontent:\n<p>Full HTML content here</p>\n[/ACTION]\n\n`;
    prompt += `To create a folder:\n[ACTION:create_folder]\nname: Folder Name\n[/ACTION]\n\n`;
    prompt += `EXAMPLE of creating a note with content:\n[ACTION:create_note]\ntitle: Shopping List\ncontent:\n<h2>Groceries</h2>\n<ul>\n<li>Milk</li>\n<li>Bread</li>\n<li>Eggs</li>\n</ul>\n[/ACTION]\n\n`;
    prompt += `EXAMPLE of editing the current note to add content:\n[ACTION:edit_note]\ntitle: My Note\ncontent:\n<p>Original content here...</p>\n<p>New content added here...</p>\n[/ACTION]\n\n`;
    prompt += `CRITICAL RULES:\n`;
    prompt += `1. **CONTENT IS MANDATORY**: Every [ACTION:create_note] and [ACTION:edit_note] MUST include "content:" followed by the full HTML content on the next lines. NEVER create an action with only a title and no content.\n`;
    prompt += `2. **ALWAYS prefer [ACTION:edit_note]** when the user has a note open and asks to edit, modify, improve, rewrite, add, fix, update, or change it. edit_note replaces the content of the currently open note.\n`;
    prompt += `3. Only use [ACTION:create_note] when the user EXPLICITLY asks to create a NEW, SEPARATE note (e.g. "create a new note about...", "make a new note", "crea una nota").\n`;
    prompt += `4. If the user says things like "edit it", "add more detail", "rewrite this", "improve the note", "add steps", you MUST use [ACTION:edit_note], NOT [ACTION:create_note].\n`;
    prompt += `5. **PRESERVE EXISTING CONTENT**: When using [ACTION:edit_note], COPY the ENTIRE original HTML content and apply the requested changes. Do NOT rewrite from scratch. Copy the original HTML as-is and only modify/append what the user asked for.\n`;
    prompt += `6. If the user asks to "add something at the end", copy ALL original HTML first, then append new content.\n`;
    prompt += `7. Use action blocks ONLY when the user asks for a change/creation. Otherwise respond normally.\n`;
    prompt += `8. Always generate content in HTML format using <p>, <b>, <i>, <u>, <ul>, <li>, <h1>, <h2>, <h3>, <br> tags.\n`;
    prompt += `9. **ALWAYS close action blocks with [/ACTION]**. Every [ACTION:...] MUST have a matching [/ACTION].\n`;
    prompt += `10. Put ALL content ONLY inside the action block. Do NOT repeat content outside it.\n`;
    prompt += `11. Keep the original note title unless the user asks to change it.\n`;

    return prompt;
  }

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

  async applyAction(messageId: number | undefined, actionId: string): Promise<void> {
    const msgs = this.messages();
    const msgIndex = msgs.findIndex(m => m.id === messageId);
    if (msgIndex === -1) return;

    const msg = msgs[msgIndex];
    const action = msg.actions?.find(a => a.id === actionId);
    if (!action || action.status !== 'pending') return;

    try {
      switch (action.type) {
        case 'edit_note': {
          const note = this.noteService.selectedNote();
          if (!note) throw new Error('No note selected');
          const updateData: { title?: string; content?: string } = {};
          if (action.title) updateData.title = action.title;
          if (action.content) updateData.content = this.ensureHtml(action.content);
          await this.noteService.updateNote(note.id, updateData);
          // Force re-select to refresh editor (toggle to null then back)
          this.noteService.selectedNoteId.set(null);
          this.noteService.selectedNoteId.set(note.id);
          break;
        }
        case 'create_note': {
          const newNote = await this.noteService.createNote();
          const updateData: { title?: string; content?: string } = {};
          if (action.title) updateData.title = action.title;
          if (action.content) updateData.content = this.ensureHtml(action.content);
          await this.noteService.updateNote(newNote.id, updateData);
          // Force re-select to refresh editor (toggle to null then back)
          this.noteService.selectedNoteId.set(null);
          this.noteService.selectedNoteId.set(newNote.id);
          break;
        }
        case 'create_folder': {
          if (action.folderName) {
            await this.noteService.createFolder(action.folderName);
          }
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

  // ── Send message ──

  async sendMessage(content: string): Promise<void> {
    if (!content.trim() || this.isStreaming() || !this.selectedModel()) return;

    if (!this.activeConversationId()) {
      const title = content.trim().substring(0, 50) || 'Nueva conversación';
      const conv = await window.electronAPI.conversationsCreate(title, this.selectedModel());
      this.conversations.update(list => [conv, ...list]);
      this.activeConversationId.set(conv.id);
    }

    const convId = this.activeConversationId()!;

    const userMsg = await window.electronAPI.conversationsAddMessage(convId, 'user', content.trim());
    this.messages.update(msgs => [
      ...msgs,
      { id: userMsg.id, role: 'user', content: content.trim(), timestamp: new Date() },
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
      const newTitle = content.trim().substring(0, 50);
      await window.electronAPI.conversationsUpdate(convId, { title: newTitle });
      this.conversations.update(list =>
        list.map(c => c.id === convId ? { ...c, title: newTitle } : c)
      );
    }

    await window.electronAPI.conversationsUpdate(convId, { model: this.selectedModel() });

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
    this.cleanupToken = window.electronAPI.onOllamaToken((token: string) => {
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

    this.cleanupDone = window.electronAPI.onOllamaDone(() => {
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
    });

    this.cleanupError = window.electronAPI.onOllamaError((error: string) => {
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
    });

    await window.electronAPI.ollamaChat(this.selectedModel(), history);
  }

  async abort(): Promise<void> {
    await window.electronAPI.ollamaAbort();
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
  }

  clearHistory(): void {
    if (this.activeConversationId()) {
      this.deleteConversation(this.activeConversationId()!);
    }
    this.messages.set([]);
    this.activeConversationId.set(null);
  }

  private removeListeners(): void {
    this.cleanupToken?.();
    this.cleanupDone?.();
    this.cleanupError?.();
    this.cleanupToken = null;
    this.cleanupDone = null;
    this.cleanupError = null;
  }
}
