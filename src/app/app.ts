import { Component, inject, OnInit, signal } from '@angular/core';
import { Sidebar } from './components/sidebar/sidebar';
import { NoteList } from './components/note-list/note-list';
import { NoteEditor } from './components/note-editor/note-editor';
import { Settings } from './components/settings/settings';
import { ResizeHandle } from './components/resize-handle/resize-handle';
import { NavigationService } from './services/navigation.service';
import { ThemeService } from './services/theme.service';

@Component({
  selector: 'app-root',
  imports: [Sidebar, NoteList, NoteEditor, Settings, ResizeHandle],
  templateUrl: './app.html',
  styleUrl: './app.scss',
})
export class App implements OnInit {
  protected readonly navService = inject(NavigationService);
  private readonly themeService = inject(ThemeService);

  protected sidebarWidth = signal(256);
  protected noteListWidth = signal(320);

  private readonly SIDEBAR_MIN = 200;
  private readonly SIDEBAR_MAX = 400;
  private readonly NOTELIST_MIN = 240;
  private readonly NOTELIST_MAX = 500;

  async ngOnInit(): Promise<void> {
    await this.themeService.init();
  }

  onSidebarResize(delta: number): void {
    const newWidth = Math.min(this.SIDEBAR_MAX, Math.max(this.SIDEBAR_MIN, this.sidebarWidth() + delta));
    this.sidebarWidth.set(newWidth);
  }

  onNoteListResize(delta: number): void {
    const newWidth = Math.min(this.NOTELIST_MAX, Math.max(this.NOTELIST_MIN, this.noteListWidth() + delta));
    this.noteListWidth.set(newWidth);
  }
}

