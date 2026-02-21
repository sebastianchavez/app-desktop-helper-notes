import { Injectable, signal } from '@angular/core';

export type AppView = 'notes' | 'settings';

@Injectable({ providedIn: 'root' })
export class NavigationService {
  readonly currentView = signal<AppView>('notes');

  goToNotes(): void {
    this.currentView.set('notes');
  }

  goToSettings(): void {
    this.currentView.set('settings');
  }
}
