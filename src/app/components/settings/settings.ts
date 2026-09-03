import { Component, inject } from '@angular/core';
import { TranslateModule } from '@ngx-translate/core';
import { ThemeService, Theme, ColorTheme, COLOR_THEME_OPTIONS } from '../../services/theme.service';
import { I18nService, LANGUAGE_OPTIONS, AppLanguage } from '../../services/i18n.service';

@Component({
  selector: 'app-settings',
  imports: [TranslateModule],
  templateUrl: './settings.html',
  styleUrl: './settings.scss',
})
/** Presenta las preferencias de tema, color e idioma y delega su persistencia. */
export class Settings {
  protected readonly themeService = inject(ThemeService);
  protected readonly i18nService = inject(I18nService);

  protected readonly themeOptions: { value: Theme; icon: string; labelKey: string; descKey: string }[] = [
    { value: 'light', icon: 'light_mode', labelKey: 'SETTINGS.THEME_LIGHT', descKey: 'SETTINGS.THEME_LIGHT_DESC' },
    { value: 'dark', icon: 'dark_mode', labelKey: 'SETTINGS.THEME_DARK', descKey: 'SETTINGS.THEME_DARK_DESC' },
    { value: 'system', icon: 'desktop_windows', labelKey: 'SETTINGS.THEME_SYSTEM', descKey: 'SETTINGS.THEME_SYSTEM_DESC' },
  ];

  protected readonly colorThemeOptions = COLOR_THEME_OPTIONS;
  protected readonly languageOptions = LANGUAGE_OPTIONS;

  async selectTheme(theme: Theme): Promise<void> {
    await this.themeService.setTheme(theme);
  }

  async selectColorTheme(colorTheme: ColorTheme): Promise<void> {
    await this.themeService.setColorTheme(colorTheme);
  }

  async onCustomColorChange(hex: string): Promise<void> {
    await this.themeService.setCustomColor(hex);
  }

  async selectLanguage(lang: AppLanguage): Promise<void> {
    await this.i18nService.setLanguage(lang);
  }

  getColorLabelKey(value: string): string {
    return `SETTINGS.COLOR_${value.toUpperCase()}`;
  }

  getColorDescKey(value: string): string {
    return `SETTINGS.COLOR_${value.toUpperCase()}_DESC`;
  }
}
