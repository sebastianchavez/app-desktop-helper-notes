import { TestBed } from '@angular/core/testing';
import { AssistantConfigService } from './assistant-config.service';

describe('AssistantConfigService', () => {
  let savedValue: string | undefined;

  beforeEach(() => {
    savedValue = undefined;
    Object.defineProperty(window, 'electronAPI', {
      configurable: true,
      value: {
        settingsGet: async () => savedValue,
        settingsSet: async (_key: string, value: string) => {
          savedValue = value;
          return { success: true };
        },
      },
    });
    TestBed.configureTestingModule({});
  });

  it('creates a usable default assistant', async () => {
    const service = TestBed.inject(AssistantConfigService);
    await service.init();

    expect(service.activeAssistant().provider).toBe('ollama');
    expect(service.activeAssistant().actions.some(action => action.mode === 'replace_text')).toBe(true);
  });

  it('persists new assistants and their actions', async () => {
    const service = TestBed.inject(AssistantConfigService);
    await service.init();
    const assistant = await service.addAssistant();
    await service.updateAssistant(assistant.id, { role: 'Profesor' });
    await service.addAction(assistant.id);

    const persisted = JSON.parse(savedValue ?? '{}') as { assistants?: { id: string; role: string; actions: unknown[] }[] };
    const savedAssistant = persisted.assistants?.find(item => item.id === assistant.id);
    expect(savedAssistant?.role).toBe('Profesor');
    expect(savedAssistant?.actions.length).toBeGreaterThan(6);
  });
});
