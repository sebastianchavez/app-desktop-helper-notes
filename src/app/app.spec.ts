import { TestBed } from '@angular/core/testing';
import { provideTranslateService } from '@ngx-translate/core';
import { App } from './app';
import { ElectronAPI } from './types/electron-api';
import { NoteService } from './services/note.service';
import { NavigationService } from './services/navigation.service';

/**
 * Crea un puente Electron de prueba sin acceder a procesos o archivos reales.
 * @param overrides Métodos que el caso de prueba necesita personalizar.
 * @returns Contrato Electron con respuestas vacías para el resto de operaciones.
 */
function createElectronApiMock(overrides: Partial<ElectronAPI> = {}): ElectronAPI {
  const emptyListener = () => () => undefined;
  return new Proxy({ platform: 'win32', ...overrides } as Partial<ElectronAPI>, {
    get(target, property: keyof ElectronAPI) {
      if (property in target) return target[property];
      if (String(property).startsWith('on')) return emptyListener;
      if (property === 'settingsGet') return async () => undefined;
      if (
        property === 'fileTakePendingOpen' ||
        property === 'foldersGetAll' ||
        property === 'notesGetByFolder' ||
        property === 'conversationsGetAll'
      ) {
        return async () => [];
      }
      return async () => ({ success: true });
    },
  }) as ElectronAPI;
}

describe('App', () => {
  beforeEach(async () => {
    Object.defineProperty(window, 'electronAPI', {
      configurable: true,
      value: createElectronApiMock(),
    });

    await TestBed.configureTestingModule({
      imports: [App],
      providers: [provideTranslateService({ fallbackLang: 'es' })],
    }).compileComponents();
  });

  it('should create the app', () => {
    const fixture = TestBed.createComponent(App);
    fixture.detectChanges();
    expect(fixture.componentInstance).toBeTruthy();
  });

  it('should render the application title', async () => {
    const fixture = TestBed.createComponent(App);
    fixture.detectChanges();
    await fixture.whenStable();
    const compiled = fixture.nativeElement as HTMLElement;
    expect(compiled.querySelector('h1')?.textContent).toContain('Helper Notes');
  });

  it('imports dropped physical files even over a nested folder drop handler', async () => {
    const path = 'C:\\documents\\report.pdf';
    const fileGetPath = vi.fn(() => path);
    Object.defineProperty(window, 'electronAPI', {
      configurable: true,
      value: createElectronApiMock({ fileGetPath }),
    });
    const fixture = TestBed.createComponent(App);
    const importer = vi
      .spyOn(TestBed.inject(NoteService), 'importPaths')
      .mockResolvedValue({ count: 1, errors: [] });
    fixture.detectChanges();
    await fixture.whenStable();
    TestBed.inject(NavigationService).goToSettings();
    const target = document.createElement('div');
    const nestedDrop = vi.fn((event: Event) => event.stopPropagation());
    target.addEventListener('drop', nestedDrop);
    document.body.appendChild(target);
    const event = new Event('drop', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'dataTransfer', {
      value: { types: ['Files'], files: [new File(['pdf'], 'report.pdf')] },
    });
    target.dispatchEvent(event);
    await vi.waitFor(() => expect(importer).toHaveBeenCalledWith([path]));
    await fixture.whenStable();
    expect(event.defaultPrevented).toBe(true);
    expect(nestedDrop).not.toHaveBeenCalled();
    expect(TestBed.inject(NavigationService).currentView()).toBe('notes');
    target.remove();
  });

  it('consumes files queued before the renderer was ready', async () => {
    const path = 'C:\\documents\\startup.txt';
    Object.defineProperty(window, 'electronAPI', {
      configurable: true,
      value: createElectronApiMock({ fileTakePendingOpen: async () => [path] }),
    });
    const fixture = TestBed.createComponent(App);
    const importer = vi
      .spyOn(TestBed.inject(NoteService), 'importPaths')
      .mockResolvedValue({ count: 1, errors: [] });
    fixture.detectChanges();
    await vi.waitFor(() => expect(importer).toHaveBeenCalledWith([path]));
  });

  it('serializes dropped batches and continues after an import failure', async () => {
    Object.defineProperty(window, 'electronAPI', {
      configurable: true,
      value: createElectronApiMock({ fileGetPath: (file) => `C:\\documents\\${file.name}` }),
    });
    const fixture = TestBed.createComponent(App);
    let rejectFirst!: (error: Error) => void;
    const importer = vi
      .spyOn(TestBed.inject(NoteService), 'importPaths')
      .mockImplementationOnce(
        () =>
          new Promise((_resolve, reject) => {
            rejectFirst = reject;
          }),
      )
      .mockResolvedValue({ count: 1, errors: [] });
    fixture.detectChanges();
    for (const name of ['first.txt', 'second.txt']) {
      const event = new Event('drop', { bubbles: true, cancelable: true });
      Object.defineProperty(event, 'dataTransfer', {
        value: { types: ['Files'], files: [new File(['text'], name)] },
      });
      window.dispatchEvent(event);
    }
    await vi.waitFor(() => expect(importer).toHaveBeenCalledTimes(1));
    rejectFirst(new Error('Unavailable'));
    await vi.waitFor(() => expect(importer).toHaveBeenCalledTimes(2));
    expect(importer).toHaveBeenLastCalledWith(['C:\\documents\\second.txt']);
  });

  it('leaves internal drag and drop events for the folder tree', () => {
    const fixture = TestBed.createComponent(App);
    const importer = vi.spyOn(TestBed.inject(NoteService), 'importPaths');
    fixture.detectChanges();
    const event = new Event('drop', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'dataTransfer', {
      value: { types: ['application/folder-id'], files: [] },
    });
    window.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
    expect(importer).not.toHaveBeenCalled();
  });
});
