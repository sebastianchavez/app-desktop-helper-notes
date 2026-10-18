import { DestroyRef, Injectable, inject, signal } from '@angular/core';
import type { Editor } from '@docx-editor.dev/core';
import { NoteService } from './note.service';

export type WordSaveState = 'saved' | 'pending' | 'saving' | 'error';

/**
 * Serializa bytes en bloques acotados sin desbordar la pila con documentos grandes.
 * @param buffer DOCX binario completo exportado por el motor.
 * @returns Base64 estándar para el canal de persistencia de Electron.
 */
export function wordBytesToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  const blocks: string[] = [];
  for (let offset = 0; offset < bytes.length; offset += 32768)
    blocks.push(String.fromCharCode(...bytes.subarray(offset, offset + 32768)));
  return btoa(blocks.join(''));
}

/** Conserva un editor y sus cambios hasta confirmar la escritura del DOCX completo. */
export class WordSession {
  state: WordSaveState = 'saved';
  private revision = 0;
  private savedRevision = 0;
  private timer?: ReturnType<typeof setTimeout>;
  private writing: Promise<void> | null = null;
  private closing = false;
  private disposed = false;
  private readonly unsubscribe: () => void;

  /** Recibe el editor, la persistencia y los callbacks de estado/limpieza; no exporta al abrir. */
  constructor(
    readonly editor: Pick<Editor, 'save' | 'on'>,
    private readonly persist: (base64: string) => Promise<void>,
    private readonly changed: () => void,
    private readonly dispose: () => void,
  ) {
    this.unsubscribe = editor.on('change', (change) => {
      if (change.source || this.disposed) return;
      this.revision++;
      this.state = 'pending';
      this.changed();
      clearTimeout(this.timer);
      this.timer = setTimeout(() => void this.flush().catch(() => undefined), 700);
    });
  }

  /** Devuelve si hay cambios sin confirmar o una escritura en curso; no recibe argumentos. */
  get pending(): boolean {
    return this.revision !== this.savedRevision || this.writing !== null;
  }

  /**
   * Guarda en orden todas las revisiones, incluidas las escritas durante una exportación.
   * @returns Promesa resuelta al confirmar la última revisión o rechazada sin perder el editor.
   */
  flush(): Promise<void> {
    clearTimeout(this.timer);
    if (this.writing) return this.writing;
    if (this.disposed) return Promise.resolve();
    // Diferir el trabajo permite asignar writing antes de ejecutar callbacks síncronos del motor.
    this.writing = Promise.resolve().then(async () => {
      try {
        while (this.revision !== this.savedRevision) {
          const revision = this.revision;
          this.state = 'saving';
          this.changed();
          const bytes = await this.editor.save();
          await this.persist(wordBytesToBase64(bytes));
          this.savedRevision = revision;
        }
        this.state = 'saved';
        if (this.closing) this.finish();
      } catch (error) {
        this.state = 'error';
        throw error;
      } finally {
        this.writing = null;
        this.changed();
      }
    });
    return this.writing;
  }

  /** Exporta el estado actual después de confirmar el guardado; retorna bytes DOCX completos. */
  async export(): Promise<ArrayBuffer> {
    await this.flush();
    return this.editor.save();
  }

  /**
   * Solicita guardar y liberar la sesión al cambiar de nota.
   * @returns Promesa de cierre; un fallo conserva el modelo en memoria para reintentar.
   */
  close(): Promise<void> {
    this.closing = true;
    return this.flush();
  }

  /** Libera listeners y UI después de persistir; no recibe argumentos ni retorna valor. */
  private finish(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.unsubscribe();
    this.dispose();
  }
}

/** Coordina autosave, reintentos y cierre de documentos incluso al cambiar a otra vista. */
@Injectable({ providedIn: 'root' })
export class WordSessionService {
  readonly state = signal<WordSaveState>('saved');
  private readonly notes = inject(NoteService);
  private readonly sessions = new Set<WordSession>();

  constructor() {
    const beforeUnload = (event: BeforeUnloadEvent): void => {
      if (![...this.sessions].some((session) => session.pending)) return;
      event.preventDefault();
      event.returnValue = '';
      void this.flushAll().catch(() => undefined);
    };
    window.addEventListener('beforeunload', beforeUnload);
    inject(DestroyRef).onDestroy(() => window.removeEventListener('beforeunload', beforeUnload));
  }

  /**
   * Registra el motor de una nota sin cambiar su selección.
   * @param noteId Nota propietaria del DOCX.
   * @param editor Contrato del motor local de edición.
   * @param dispose Limpieza de su UI después del último guardado.
   * @returns Sesión que controla guardado, exportación y cierre.
   */
  register(noteId: number, editor: Editor, dispose: () => void): WordSession {
    const session = new WordSession(
      editor,
      (base64) => this.notes.saveWord(noteId, base64),
      () => this.publish(),
      () => {
        dispose();
        this.sessions.delete(session);
      },
    );
    this.sessions.add(session);
    return session;
  }

  /** Reintenta todas las sesiones pendientes; retorna cuando todas confirman su escritura. */
  async flushAll(): Promise<void> {
    const results = await Promise.allSettled([...this.sessions].map((session) => session.flush()));
    const failed = results.find((result) => result.status === 'rejected');
    if (failed?.status === 'rejected') throw failed.reason;
  }

  /** Publica el estado agregado; los errores anteriores siguen visibles al cambiar de nota. */
  private publish(): void {
    const states = [...this.sessions].map((session) => session.state);
    this.state.set(
      states.includes('error')
        ? 'error'
        : states.includes('saving')
          ? 'saving'
          : states.includes('pending')
            ? 'pending'
            : 'saved',
    );
  }
}
