import { Component, inject } from '@angular/core';
import { ThemeService, Theme, ColorTheme, COLOR_THEME_OPTIONS } from '../../services/theme.service';

@Component({
  selector: 'app-settings',
  templateUrl: './settings.html',
  styleUrl: './settings.scss',
})
export class Settings {
  protected readonly themeService = inject(ThemeService);

  protected readonly themeOptions: { value: Theme; label: string; icon: string; description: string }[] = [
    { value: 'light', label: 'Claro', icon: 'light_mode', description: 'Interfaz con fondo claro y texto oscuro' },
    { value: 'dark', label: 'Oscuro', icon: 'dark_mode', description: 'Interfaz con fondo oscuro, ideal para uso nocturno' },
    { value: 'system', label: 'Sistema', icon: 'desktop_windows', description: 'Se adapta automáticamente al tema de tu sistema operativo' },
  ];

  protected readonly colorThemeOptions = COLOR_THEME_OPTIONS;

  async selectTheme(theme: Theme): Promise<void> {
    await this.themeService.setTheme(theme);
  }

  async selectColorTheme(colorTheme: ColorTheme): Promise<void> {
    await this.themeService.setColorTheme(colorTheme);
  }

  async onCustomColorChange(hex: string): Promise<void> {
    await this.themeService.setCustomColor(hex);
  }
}
