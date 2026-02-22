import { Injectable, inject, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';

export type AppLanguage = 'es' | 'en' | 'zh' | 'de' | 'fr' | 'ja' | 'pt';

export interface LanguageOption {
  value: AppLanguage;
  label: string;       // native name
  flag: string;        // emoji flag
}

export const LANGUAGE_OPTIONS: LanguageOption[] = [
  { value: 'es', label: 'Español', flag: '🇪🇸' },
  { value: 'en', label: 'English', flag: '🇺🇸' },
  { value: 'zh', label: '中文', flag: '🇨🇳' },
  { value: 'de', label: 'Deutsch', flag: '🇩🇪' },
  { value: 'fr', label: 'Français', flag: '🇫🇷' },
  { value: 'ja', label: '日本語', flag: '🇯🇵' },
  { value: 'pt', label: 'Português', flag: '🇧🇷' },
];

/** Map language codes to full language names (for AI prompt) */
const LANGUAGE_NAMES: Record<AppLanguage, string> = {
  es: 'Spanish',
  en: 'English',
  zh: 'Chinese',
  de: 'German',
  fr: 'French',
  ja: 'Japanese',
  pt: 'Portuguese',
};

@Injectable({ providedIn: 'root' })
export class I18nService {
  private readonly translate = inject(TranslateService);
  private api = window.electronAPI;

  readonly currentLang = signal<AppLanguage>('es');

  async init(): Promise<void> {
    this.translate.addLangs(['es', 'en', 'zh', 'de', 'fr', 'ja', 'pt']);
    this.translate.setDefaultLang('es');

    const saved = await this.api.settingsGet('language');
    const lang = (saved as AppLanguage) || 'es';
    this.currentLang.set(lang);
    this.translate.use(lang);
  }

  async setLanguage(lang: AppLanguage): Promise<void> {
    this.currentLang.set(lang);
    this.translate.use(lang);
    await this.api.settingsSet('language', lang);
  }

  /** Get full language name in English (for AI system prompts) */
  getLanguageName(): string {
    return LANGUAGE_NAMES[this.currentLang()] || 'Spanish';
  }

  /** Synchronous translate (for use in TS code) */
  instant(key: string, params?: Record<string, unknown>): string {
    return this.translate.instant(key, params);
  }
}
