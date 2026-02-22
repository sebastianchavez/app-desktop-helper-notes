import { Component, inject } from '@angular/core';
import { TranslateModule } from '@ngx-translate/core';
import { DownloadService } from '../../services/download.service';

@Component({
  selector: 'app-download-toast',
  templateUrl: './download-toast.html',
  styleUrl: './download-toast.scss',
  imports: [TranslateModule],
})
export class DownloadToast {
  protected readonly downloadService = inject(DownloadService);

  get downloads() { return this.downloadService.downloads; }

  dismiss(fileName: string): void {
    this.downloadService.dismiss(fileName);
  }

  openInFolder(filePath: string): void {
    window.electronAPI.fileShowInFolder(filePath);
  }

  getShortPath(filePath: string): string {
    const parts = filePath.replace(/\\/g, '/').split('/');
    return parts.length > 2 ? '.../' + parts.slice(-2).join('/') : filePath;
  }
}
