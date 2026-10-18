import { InjectionToken } from '@angular/core';
import type { Editor } from '@docx-editor.dev/core';

export interface WordEditorMount {
  readonly ready: Promise<Editor>;
  destroy(): void;
}

export interface WordEditorOptions {
  title: string;
  language: string;
  save(): void;
  print(): void;
  clipboardError?(key: string | null): void;
  open(file: File): void;
}

export type WordEditorFactory = (
  host: HTMLElement,
  bytes: Uint8Array,
  options: WordEditorOptions,
) => Promise<WordEditorMount>;

/** Carga el motor y Vue solamente al abrir un Word; permite probar el ciclo de vida sin navegador. */
export const WORD_EDITOR_FACTORY = new InjectionToken<WordEditorFactory>('Word editor factory', {
  providedIn: 'root',
  factory: () => async (host, bytes, options) =>
    (await import('./word-editor')).mountWordEditor(host, bytes, options),
});
