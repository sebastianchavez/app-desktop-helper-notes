import { TestBed } from '@angular/core/testing';
import { provideTranslateService } from '@ngx-translate/core';
import { App } from './app';
import { ElectronAPI } from './types/electron-api';

function createElectronApiMock(): ElectronAPI {
  const emptyListener = () => () => undefined;
  return new Proxy({ platform: 'win32' } as Partial<ElectronAPI>, {
    get(target, property: keyof ElectronAPI) {
      if (property in target) return target[property];
      if (String(property).startsWith('on')) return emptyListener;
      if (property === 'settingsGet') return async () => undefined;
      if (property === 'foldersGetAll' || property === 'notesGetByFolder' || property === 'conversationsGetAll') {
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
});
