import { Injectable, signal } from '@angular/core';

export interface DownloadInfo {
  fileName: string;
  filePath: string;
  status: 'saving' | 'done' | 'error';
}

@Injectable({ providedIn: 'root' })
export class DownloadService {
  readonly downloads = signal<DownloadInfo[]>([]);

  showSaving(fileName: string): void {
    this.downloads.update((list) => [
      ...list.filter((d) => d.fileName !== fileName),
      { fileName, filePath: '', status: 'saving' as const },
    ]);
  }

  showDone(fileName: string, filePath: string): void {
    this.downloads.update((list) =>
      list.map((d) => (d.fileName === fileName ? { ...d, filePath, status: 'done' as const } : d))
    );
    setTimeout(() => this.dismiss(fileName), 8000);
  }

  showError(fileName: string): void {
    this.downloads.update((list) =>
      list.map((d) => (d.fileName === fileName ? { ...d, status: 'error' as const } : d))
    );
    setTimeout(() => this.dismiss(fileName), 5000);
  }

  dismiss(fileName: string): void {
    this.downloads.update((list) => list.filter((d) => d.fileName !== fileName));
  }
}
