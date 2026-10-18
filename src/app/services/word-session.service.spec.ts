import type { Editor, DocumentChange } from '@docx-editor.dev/core';
import { WordSession } from './word-session.service';

describe('WordSession', () => {
  let change: (change: DocumentChange) => void;
  let session: WordSession;
  let save = vi.fn(async () => new Uint8Array([1, 2, 3]).buffer);
  let persist = vi.fn(async (_base64: string): Promise<void> => undefined);
  let dispose = vi.fn(() => undefined);

  beforeEach(() => {
    vi.useFakeTimers();
    save = vi.fn(async () => new Uint8Array([1, 2, 3]).buffer);
    persist = vi.fn(async (_base64: string): Promise<void> => undefined);
    dispose = vi.fn(() => undefined);
    const editor = {
      save,
      on: (_event: string, callback: typeof change) => {
        change = callback;
        return () => undefined;
      },
    } as unknown as Editor;
    session = new WordSession(
      editor,
      (base64) => persist(base64),
      () => undefined,
      () => dispose(),
    );
  });
  afterEach(() => vi.useRealTimers());

  it('does not overwrite a document merely by opening it', async () => {
    change({ revision: 1, source: 'load' });
    await session.flush();
    expect(save).not.toHaveBeenCalled();
    expect(persist).not.toHaveBeenCalled();
    await session.close();
    expect(dispose).toHaveBeenCalledTimes(1);
  });

  it('debounces edits and confirms the serialized DOCX only after persistence succeeds', async () => {
    change({ revision: 1 });
    change({ revision: 2 });
    expect(session.state).toBe('pending');
    await vi.advanceTimersByTimeAsync(700);
    expect(save).toHaveBeenCalledTimes(1);
    expect(persist).toHaveBeenCalledWith('AQID');
    expect(session.state).toBe('saved');
    expect(session.pending).toBe(false);
  });

  it('serializes revisions made while the previous write is still pending', async () => {
    let release!: () => void;
    persist.mockImplementationOnce(
      () =>
        new Promise<void>((resolve) => {
          release = resolve;
        }),
    );
    change({ revision: 1 });
    const writing = session.flush();
    await vi.advanceTimersByTimeAsync(0);
    change({ revision: 2 });
    release();
    await writing;
    expect(persist).toHaveBeenCalledTimes(2);
    expect(session.pending).toBe(false);
    expect(session.state).toBe('saved');
  });

  it('retains a failed closing session and releases its engine only after retry succeeds', async () => {
    persist.mockRejectedValueOnce(new Error('disk full'));
    change({ revision: 1 });
    await expect(session.close()).rejects.toThrow('disk full');
    expect(session.state).toBe('error');
    expect(dispose).not.toHaveBeenCalled();
    expect(session.pending).toBe(true);
    await session.flush();
    expect(persist).toHaveBeenCalledTimes(2);
    expect(dispose).toHaveBeenCalledTimes(1);
    expect(session.pending).toBe(false);
  });
});
