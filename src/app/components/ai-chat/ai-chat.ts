import {
  Component,
  inject,
  signal,
  viewChild,
  ElementRef,
  OnInit,
  OnDestroy,
  AfterViewChecked,
} from '@angular/core';
import { AiChatService, ChatAction, ChatMessage } from '../../services/ai-chat.service';
import { NoteService } from '../../services/note.service';
import { TranslateModule, TranslateService } from '@ngx-translate/core';

@Component({
  selector: 'app-ai-chat',
  templateUrl: './ai-chat.html',
  styleUrl: './ai-chat.scss',
  imports: [TranslateModule],
})
export class AiChat implements OnInit, OnDestroy, AfterViewChecked {
  protected readonly chatService = inject(AiChatService);
  protected readonly noteService = inject(NoteService);
  private readonly translate = inject(TranslateService);

  protected inputText = signal('');
  protected showModelSelector = signal(false);
  protected confirmDeleteId = signal<number | null>(null);

  private messagesContainer = viewChild<ElementRef<HTMLDivElement>>('messagesContainer');
  private inputRef = viewChild<ElementRef<HTMLTextAreaElement>>('chatInput');
  private shouldScrollToBottom = false;
  private prevMessageCount = 0;

  async ngOnInit(): Promise<void> {
    await this.chatService.checkOllamaStatus();
    if (this.chatService.ollamaAvailable()) {
      await this.chatService.loadModels();
      await this.chatService.loadConversations();
    }
  }

  ngOnDestroy(): void {
    if (this.chatService.isStreaming()) {
      this.chatService.abort();
    }
  }

  ngAfterViewChecked(): void {
    if (this.shouldScrollToBottom) {
      this.scrollToBottom();
      this.shouldScrollToBottom = false;
    }
    const currentCount = this.chatService.messages().length;
    const lastMsg = this.chatService.messages()[currentCount - 1];
    if (lastMsg?.isStreaming) {
      this.scrollToBottom();
    }
    if (currentCount !== this.prevMessageCount) {
      this.prevMessageCount = currentCount;
      this.shouldScrollToBottom = true;
    }
  }

  async onSend(): Promise<void> {
    const text = this.inputText().trim();
    if (!text) return;
    this.inputText.set('');
    this.shouldScrollToBottom = true;
    this.resetTextareaHeight();
    await this.chatService.sendMessage(text);
  }

  onKeydown(event: KeyboardEvent): void {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      this.onSend();
    }
  }

  onInput(event: Event): void {
    const textarea = event.target as HTMLTextAreaElement;
    this.inputText.set(textarea.value);
    textarea.style.height = 'auto';
    textarea.style.height = Math.min(textarea.scrollHeight, 150) + 'px';
  }

  selectModel(name: string): void {
    this.chatService.setSelectedModel(name);
    this.showModelSelector.set(false);
  }

  onClearChat(): void {
    this.chatService.clearHistory();
  }

  async onNewConversation(): Promise<void> {
    await this.chatService.newConversation();
  }

  async onSelectConversation(id: number): Promise<void> {
    this.shouldScrollToBottom = true;
    await this.chatService.selectConversation(id);
  }

  async onDeleteConversation(id: number, event: Event): Promise<void> {
    event.stopPropagation();
    if (this.confirmDeleteId() === id) {
      await this.chatService.deleteConversation(id);
      this.confirmDeleteId.set(null);
    } else {
      this.confirmDeleteId.set(id);
      // Auto-reset after 3s
      setTimeout(() => {
        if (this.confirmDeleteId() === id) this.confirmDeleteId.set(null);
      }, 3000);
    }
  }

  toggleSessions(): void {
    this.chatService.showSessions.update(v => !v);
  }

  async onRefreshModels(): Promise<void> {
    await this.chatService.checkOllamaStatus();
    if (this.chatService.ollamaAvailable()) {
      await this.chatService.loadModels();
      await this.chatService.loadConversations();
    }
  }

  async onStartOllama(): Promise<void> {
    await this.chatService.startOllama();
    if (this.chatService.ollamaAvailable()) {
      await this.chatService.loadConversations();
    }
  }

  async onStopOllama(): Promise<void> {
    await this.chatService.stopOllama();
  }

  onOpenDownload(): void {
    this.chatService.openDownload();
  }

  formatModelName(name: string): string {
    const base = name.split(':')[0];
    return base.charAt(0).toUpperCase() + base.slice(1);
  }

  formatModelSize(bytes: number): string {
    const gb = bytes / (1024 * 1024 * 1024);
    if (gb >= 1) return gb.toFixed(1) + ' GB';
    const mb = bytes / (1024 * 1024);
    return mb.toFixed(0) + ' MB';
  }

  formatDate(dateStr: string): string {
    const date = new Date(dateStr + 'Z');
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffMins = Math.floor(diffMs / 60000);
    if (diffMins < 1) return this.translate.instant('AI_CHAT.TIME_NOW');
    if (diffMins < 60) return this.translate.instant('AI_CHAT.TIME_MINS_AGO', { count: diffMins });
    const diffHours = Math.floor(diffMins / 60);
    if (diffHours < 24) return this.translate.instant('AI_CHAT.TIME_HOURS_AGO', { count: diffHours });
    const diffDays = Math.floor(diffHours / 24);
    if (diffDays < 7) return this.translate.instant('AI_CHAT.TIME_DAYS_AGO', { count: diffDays });
    return date.toLocaleDateString(this.translate.currentLang || 'es', { day: 'numeric', month: 'short' });
  }

  // ── Actions ──

  getMessageDisplay(msg: ChatMessage): string {
    return msg.displayContent ?? msg.content;
  }

  hasActions(msg: ChatMessage): boolean {
    return (msg.actions?.length ?? 0) > 0;
  }

  getActionLabel(action: ChatAction): string {
    switch (action.type) {
      case 'edit_note': return this.translate.instant('AI_CHAT.ACTION_EDIT_NOTE');
      case 'create_note': return this.translate.instant('AI_CHAT.ACTION_CREATE_NOTE');
      case 'create_folder': return this.translate.instant('AI_CHAT.ACTION_CREATE_FOLDER');
      default: return this.translate.instant('AI_CHAT.ACTION_DEFAULT');
    }
  }

  getActionIcon(action: ChatAction): string {
    switch (action.type) {
      case 'edit_note': return 'edit_note';
      case 'create_note': return 'note_add';
      case 'create_folder': return 'create_new_folder';
      default: return 'auto_fix_high';
    }
  }

  getActionDescription(action: ChatAction): string {
    switch (action.type) {
      case 'edit_note':
        return action.title ? `${this.translate.instant('AI_CHAT.ACTION_TITLE_PREFIX')}: "${action.title}"` : this.translate.instant('AI_CHAT.ACTION_MODIFY_CONTENT');
      case 'create_note':
        return action.title ? `"${action.title}"` : this.translate.instant('AI_CHAT.ACTION_NEW_NOTE');
      case 'create_folder':
        return action.folderName ? `"${action.folderName}"` : this.translate.instant('AI_CHAT.ACTION_NEW_FOLDER');
      default:
        return '';
    }
  }

  getContentPreview(action: ChatAction): string {
    if (action.changePreview) return action.changePreview;
    if (action.content) {
      const div = document.createElement('div');
      div.innerHTML = action.content;
      return (div.textContent || div.innerText || '').trim();
    }
    return '';
  }

  hasChangePreview(action: ChatAction): boolean {
    return action.type === 'edit_note' && !!action.changePreview;
  }

  async onApplyAction(msg: ChatMessage, action: ChatAction): Promise<void> {
    await this.chatService.applyAction(msg.id, action.id);
  }

  onRejectAction(msg: ChatMessage, action: ChatAction): void {
    this.chatService.rejectAction(msg.id, action.id);
  }

  private scrollToBottom(): void {
    const el = this.messagesContainer()?.nativeElement;
    if (el) {
      el.scrollTop = el.scrollHeight;
    }
  }

  private resetTextareaHeight(): void {
    const el = this.inputRef()?.nativeElement;
    if (el) {
      el.style.height = 'auto';
    }
  }
}
