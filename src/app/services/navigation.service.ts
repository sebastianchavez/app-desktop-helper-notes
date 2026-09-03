import { Injectable, signal } from '@angular/core';

export type AppView = 'notes' | 'settings' | 'assistant-settings' | 'data-transfer';

@Injectable({ providedIn: 'root' })
/** Mantiene la vista activa de la aplicación sin depender del router de Angular. */
export class NavigationService {
  readonly currentView = signal<AppView>('notes');

  /** Regresa al espacio principal de notas. */
  goToNotes(): void {
    this.currentView.set('notes');
  }

  /** Abre las preferencias visuales y de idioma. */
  goToSettings(): void {
    this.currentView.set('settings');
  }

  /** Abre la administración de asistentes y proveedores locales. */
  goToAssistantSettings(): void {
    this.currentView.set('assistant-settings');
  }

  /** Abre la vista de importación y exportación de datos. */
  goToDataTransfer(): void {
    this.currentView.set('data-transfer');
  }
}
