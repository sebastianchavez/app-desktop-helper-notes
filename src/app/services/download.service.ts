import { Injectable, signal } from '@angular/core';

export interface DownloadInfo {
  fileName: string;
  filePath: string;
  status: 'saving' | 'done' | 'error';
}

@Injectable({ providedIn: 'root' })
/**
 * Estado compartido de las notificaciones de descarga/exportación.
 * Cada operación se identifica por nombre de archivo para poder actualizar el
 * mismo aviso desde `saving` hasta `done` o `error`.
 */
export class DownloadService {
  readonly downloads = signal<DownloadInfo[]>([]);

  /** Registra una exportación en curso y reemplaza un aviso anterior del mismo archivo. */
  showSaving(fileName: string): void {
    this.downloads.update((list) => [
      ...list.filter((d) => d.fileName !== fileName),
      { fileName, filePath: '', status: 'saving' as const },
    ]);
  }

  /** Marca la exportación como terminada y la retira automáticamente tras ocho segundos. */
  showDone(fileName: string, filePath: string): void {
    this.downloads.update((list) =>
      list.map((d) => (d.fileName === fileName ? { ...d, filePath, status: 'done' as const } : d))
    );
    setTimeout(() => this.dismiss(fileName), 8000);
  }

  /** Marca la exportación como fallida y conserva el aviso durante cinco segundos. */
  showError(fileName: string): void {
    this.downloads.update((list) =>
      list.map((d) => (d.fileName === fileName ? { ...d, status: 'error' as const } : d))
    );
    setTimeout(() => this.dismiss(fileName), 5000);
  }

  /** Elimina manualmente una notificación de la colección reactiva. */
  dismiss(fileName: string): void {
    this.downloads.update((list) => list.filter((d) => d.fileName !== fileName));
  }
}
