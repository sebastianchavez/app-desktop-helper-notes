import { Injectable, signal } from '@angular/core';

export type Theme = 'dark' | 'light' | 'system';
export type ColorTheme = 'blue' | 'green' | 'rose' | 'amber' | 'violet' | 'custom';

export interface ColorThemeOption {
  value: ColorTheme;
  label: string;
  icon: string;
  color: string; // hex for preview
  description: string;
}

export const COLOR_THEME_OPTIONS: ColorThemeOption[] = [
  { value: 'blue', label: 'Azul', icon: 'water_drop', color: '#135bec', description: 'Azul eléctrico, el tema clásico' },
  { value: 'green', label: 'Verde', icon: 'eco', color: '#10b981', description: 'Verde esmeralda, fresco y natural' },
  { value: 'rose', label: 'Rosa', icon: 'favorite', color: '#f43f5e', description: 'Rosa intenso, vibrante y moderno' },
  { value: 'amber', label: 'Ámbar', icon: 'light_mode', color: '#f59e0b', description: 'Ámbar cálido, acogedor y dorado' },
  { value: 'violet', label: 'Violeta', icon: 'auto_awesome', color: '#8b5cf6', description: 'Violeta elegante, creativo y único' },
  { value: 'custom', label: 'Personalizado', icon: 'palette', color: '#135bec', description: 'Elige tu propio color de acento' },
];

@Injectable({ providedIn: 'root' })
export class ThemeService {
  private api = window.electronAPI;

  readonly currentTheme = signal<Theme>('dark');
  readonly currentColorTheme = signal<ColorTheme>('blue');
  readonly customColor = signal('#135bec');

  async init(): Promise<void> {
    const [savedTheme, savedColor, savedCustom] = await Promise.all([
      this.api.settingsGet('theme'),
      this.api.settingsGet('colorTheme'),
      this.api.settingsGet('customColor'),
    ]);
    const theme = (savedTheme as Theme) || 'dark';
    const colorTheme = (savedColor as ColorTheme) || 'blue';
    const customColorVal = savedCustom || '#135bec';

    this.currentTheme.set(theme);
    this.currentColorTheme.set(colorTheme);
    this.customColor.set(customColorVal);

    this.applyTheme(theme);
    this.applyColorTheme(colorTheme, customColorVal);
  }

  async setTheme(theme: Theme): Promise<void> {
    this.currentTheme.set(theme);
    this.applyTheme(theme);
    await this.api.settingsSet('theme', theme);
  }

  async setColorTheme(colorTheme: ColorTheme): Promise<void> {
    this.currentColorTheme.set(colorTheme);
    this.applyColorTheme(colorTheme, this.customColor());
    await this.api.settingsSet('colorTheme', colorTheme);
  }

  async setCustomColor(hex: string): Promise<void> {
    this.customColor.set(hex);
    if (this.currentColorTheme() === 'custom') {
      this.applyColorTheme('custom', hex);
    }
    await this.api.settingsSet('customColor', hex);
  }

  private applyTheme(theme: Theme): void {
    const html = document.documentElement;

    if (theme === 'system') {
      const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
      html.classList.toggle('dark', prefersDark);
    } else {
      html.classList.toggle('dark', theme === 'dark');
    }
  }

  private applyColorTheme(colorTheme: ColorTheme, customHex: string): void {
    const html = document.documentElement;

    if (colorTheme === 'custom') {
      // Parse hex to RGB and set CSS variables directly
      const rgb = this.hexToRgbValues(customHex);
      html.removeAttribute('data-color-theme');
      html.style.setProperty('--color-primary', rgb);
      // Generate dark bg variants from the custom color
      const darkBg = this.darkenForBg(customHex);
      const sidebarBg = this.darkenForSidebar(customHex);
      html.style.setProperty('--color-bg-dark', darkBg);
      html.style.setProperty('--color-sidebar-dark', sidebarBg);
    } else {
      // Use data attribute - CSS handles the variables
      html.setAttribute('data-color-theme', colorTheme);
      html.style.removeProperty('--color-primary');
      html.style.removeProperty('--color-bg-dark');
      html.style.removeProperty('--color-sidebar-dark');
    }
  }

  private hexToRgbValues(hex: string): string {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return `${r} ${g} ${b}`;
  }

  private darkenForBg(hex: string): string {
    const r = Math.round(parseInt(hex.slice(1, 3), 16) * 0.08);
    const g = Math.round(parseInt(hex.slice(3, 5), 16) * 0.08);
    const b = Math.round(parseInt(hex.slice(5, 7), 16) * 0.08);
    return `${Math.max(r, 12)} ${Math.max(g, 12)} ${Math.max(b, 18)}`;
  }

  private darkenForSidebar(hex: string): string {
    const r = Math.round(parseInt(hex.slice(1, 3), 16) * 0.1);
    const g = Math.round(parseInt(hex.slice(3, 5), 16) * 0.1);
    const b = Math.round(parseInt(hex.slice(5, 7), 16) * 0.1);
    return `${Math.max(r, 10)} ${Math.max(g, 10)} ${Math.max(b, 20)}`;
  }
}
