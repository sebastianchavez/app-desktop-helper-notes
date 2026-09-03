import { Injectable, computed, signal } from '@angular/core';
import {
  AssistantActionDefinition,
  AssistantActionMode,
  AssistantConfiguration,
  AssistantNoteType,
  AssistantProfile,
  AssistantProvider,
} from '../models/assistant.model';

const SETTINGS_KEY = 'assistant-config-v1';

const DEFAULT_ACTIONS: AssistantActionDefinition[] = [
  {
    id: 'analyze',
    name: 'Analizar',
    description: 'Analiza el archivo abierto y entrega observaciones sin modificarlo.',
    instructions: 'Cuando el usuario pida analizar, explica tus conclusiones y no propongas cambios salvo que los solicite.',
    mode: 'analyze',
    enabled: true,
    noteTypes: ['text', 'spreadsheet', 'presentation', 'pdf'],
  },
  {
    id: 'edit-note',
    name: 'Editar nota',
    description: 'Reescribe el contenido completo de una nota de texto.',
    instructions: 'Conserva todo el contenido que el usuario no haya pedido cambiar.',
    mode: 'edit_note',
    enabled: true,
    noteTypes: ['text'],
  },
  {
    id: 'replace-text',
    name: 'Reemplazar texto',
    description: 'Busca un fragmento exacto y lo reemplaza después de mostrar una vista previa.',
    instructions: 'Usa esta acción cuando el usuario identifique claramente el texto original y su reemplazo.',
    mode: 'replace_text',
    enabled: true,
    noteTypes: ['text'],
  },
  {
    id: 'create-note',
    name: 'Crear nota',
    description: 'Crea una nota nueva en la carpeta seleccionada.',
    instructions: 'Úsala solo si el usuario solicita explícitamente una nota nueva.',
    mode: 'create_note',
    enabled: true,
    noteTypes: ['text', 'spreadsheet', 'presentation', 'pdf'],
  },
  {
    id: 'create-folder',
    name: 'Crear carpeta',
    description: 'Crea una carpeta para organizar notas.',
    instructions: 'Solicita o deduce un nombre breve y descriptivo.',
    mode: 'create_folder',
    enabled: true,
    noteTypes: ['text', 'spreadsheet', 'presentation', 'pdf'],
  },
  {
    id: 'edit-excel',
    name: 'Editar hoja de cálculo',
    description: 'Actualiza celdas de la hoja de cálculo abierta.',
    instructions: 'Preserva las celdas no mencionadas y usa índices de fila y columna basados en cero.',
    mode: 'edit_excel',
    enabled: true,
    noteTypes: ['spreadsheet'],
  },
];

function cloneActions(): AssistantActionDefinition[] {
  return DEFAULT_ACTIONS.map(action => ({ ...action, noteTypes: [...action.noteTypes] }));
}

/** Construye el asistente general que garantiza una configuración utilizable. */
function createDefaultAssistant(): AssistantProfile {
  return {
    id: 'general',
    name: 'Asistente general',
    role: 'Asistente',
    description: 'Ayuda general para redactar, analizar y organizar notas.',
    systemPrompt: 'Actúa con precisión, explica tus decisiones y confirma mediante una acción antes de modificar un archivo.',
    provider: 'ollama',
    endpoint: 'http://127.0.0.1:11434',
    model: '',
    localModelPath: '',
    temperature: 0.7,
    actions: cloneActions(),
    noteTypeInstructions: {
      text: 'Prioriza la claridad, conserva el formato HTML y respeta el estilo original.',
      spreadsheet: 'Razona sobre encabezados, filas, columnas y tipos de datos antes de proponer cambios.',
      presentation: 'Analiza la secuencia, claridad y contenido de las diapositivas y las notas del presentador.',
      pdf: 'Los PDF importados son de solo lectura; analízalos sin proponer ediciones directas.',
    },
  };
}

@Injectable({ providedIn: 'root' })
/**
 * Repositorio reactivo de perfiles de asistentes.
 *
 * Mantiene un esquema versionado dentro de `settings`, normaliza información
 * antigua o incompleta y persiste cada mutación a través del proceso Electron.
 */
export class AssistantConfigService {
  readonly configuration = signal<AssistantConfiguration>({
    version: 1,
    activeAssistantId: 'general',
    assistants: [createDefaultAssistant()],
  });
  readonly initialized = signal(false);

  readonly assistants = computed(() => this.configuration().assistants);
  readonly activeAssistant = computed(() => {
    const config = this.configuration();
    return config.assistants.find(item => item.id === config.activeAssistantId) ?? config.assistants[0];
  });

  /** Recupera la configuración persistida o crea la predeterminada si es inválida. */
  async init(): Promise<void> {
    if (this.initialized()) return;
    try {
      const raw = await window.electronAPI.settingsGet(SETTINGS_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as Partial<AssistantConfiguration>;
        const assistants = Array.isArray(parsed.assistants)
          ? parsed.assistants.map(item => this.normalizeAssistant(item))
          : [];
        if (assistants.length > 0) {
          const activeId = assistants.some(item => item.id === parsed.activeAssistantId)
            ? parsed.activeAssistantId as string
            : assistants[0].id;
          this.configuration.set({ version: 1, activeAssistantId: activeId, assistants });
        }
      }
    } catch {
      this.configuration.set({ version: 1, activeAssistantId: 'general', assistants: [createDefaultAssistant()] });
    }
    this.initialized.set(true);
  }

  /** Cambia el asistente activo únicamente si el identificador existe. */
  async setActiveAssistant(id: string): Promise<void> {
    if (!this.assistants().some(item => item.id === id)) return;
    this.configuration.update(config => ({ ...config, activeAssistantId: id }));
    await this.persist();
  }

  /** Añade un perfil nuevo con identificador único y lo convierte en activo. */
  async addAssistant(): Promise<AssistantProfile> {
    const id = `assistant-${Date.now()}`;
    const assistant: AssistantProfile = {
      ...createDefaultAssistant(),
      id,
      name: 'Nuevo asistente',
      role: 'Asistente personalizado',
      description: '',
    };
    this.configuration.update(config => ({
      ...config,
      activeAssistantId: id,
      assistants: [...config.assistants, assistant],
    }));
    await this.persist();
    return assistant;
  }

  /** Duplica un perfil y sus acciones sin compartir referencias mutables. */
  async duplicateAssistant(id: string): Promise<AssistantProfile | null> {
    const source = this.assistants().find(item => item.id === id);
    if (!source) return null;
    const copy = this.normalizeAssistant({
      ...source,
      id: `assistant-${Date.now()}`,
      name: `${source.name} (copia)`,
    });
    this.configuration.update(config => ({
      ...config,
      activeAssistantId: copy.id,
      assistants: [...config.assistants, copy],
    }));
    await this.persist();
    return copy;
  }

  /** Fusiona cambios parciales, vuelve a normalizar el perfil y lo persiste. */
  async updateAssistant(id: string, changes: Partial<AssistantProfile>): Promise<void> {
    this.configuration.update(config => ({
      ...config,
      assistants: config.assistants.map(item => item.id === id
        ? this.normalizeAssistant({ ...item, ...changes, id })
        : item),
    }));
    await this.persist();
  }

  /** Elimina un perfil sin permitir que la colección quede vacía. */
  async deleteAssistant(id: string): Promise<boolean> {
    const config = this.configuration();
    if (config.assistants.length <= 1) return false;
    const assistants = config.assistants.filter(item => item.id !== id);
    this.configuration.set({
      ...config,
      assistants,
      activeAssistantId: config.activeAssistantId === id ? assistants[0].id : config.activeAssistantId,
    });
    await this.persist();
    return true;
  }

  /** Agrega una acción personalizada inicial al asistente indicado. */
  async addAction(assistantId: string): Promise<void> {
    const action: AssistantActionDefinition = {
      id: `custom-${Date.now()}`,
      name: 'Nueva acción',
      description: '',
      instructions: '',
      mode: 'analyze',
      enabled: true,
      noteTypes: ['text'],
    };
    const assistant = this.assistants().find(item => item.id === assistantId);
    if (!assistant) return;
    await this.updateAssistant(assistantId, { actions: [...assistant.actions, action] });
  }

  /** Actualiza una acción preservando su identificador estable. */
  async updateAction(assistantId: string, actionId: string, changes: Partial<AssistantActionDefinition>): Promise<void> {
    const assistant = this.assistants().find(item => item.id === assistantId);
    if (!assistant) return;
    const actions = assistant.actions.map(action => action.id === actionId
      ? { ...action, ...changes, id: action.id }
      : action);
    await this.updateAssistant(assistantId, { actions });
  }

  /** Elimina una acción de un perfil y guarda la configuración resultante. */
  async deleteAction(assistantId: string, actionId: string): Promise<void> {
    const assistant = this.assistants().find(item => item.id === assistantId);
    if (!assistant) return;
    await this.updateAssistant(assistantId, { actions: assistant.actions.filter(action => action.id !== actionId) });
  }

  /**
   * Sanea un perfil procedente de disco: valida proveedor, rangos numéricos,
   * modos de acción y tipos de nota, completando campos ausentes con defaults.
   */
  private normalizeAssistant(value: Partial<AssistantProfile>): AssistantProfile {
    const fallback = createDefaultAssistant();
    const validProvider: AssistantProvider = value.provider === 'llamacpp' ? 'llamacpp' : 'ollama';
    const validModes: AssistantActionMode[] = ['analyze', 'edit_note', 'replace_text', 'create_note', 'create_folder', 'edit_excel'];
    const validTypes: AssistantNoteType[] = ['text', 'spreadsheet', 'presentation', 'pdf'];
    const actions: AssistantActionDefinition[] = Array.isArray(value.actions) ? value.actions.map((action, index) => ({
      id: String(action.id || `action-${index}`),
      name: String(action.name || 'Acción'),
      description: String(action.description || ''),
      instructions: String(action.instructions || ''),
      mode: validModes.includes(action.mode) ? action.mode : 'analyze',
      enabled: action.enabled !== false,
      noteTypes: Array.isArray(action.noteTypes)
        ? action.noteTypes.filter((type): type is AssistantNoteType => validTypes.includes(type))
        : ['text' as AssistantNoteType],
    })) : cloneActions();
    return {
      ...fallback,
      ...value,
      id: String(value.id || `assistant-${Date.now()}`),
      name: String(value.name || fallback.name),
      role: String(value.role || fallback.role),
      description: String(value.description || ''),
      systemPrompt: String(value.systemPrompt || ''),
      provider: validProvider,
      endpoint: String(value.endpoint || (validProvider === 'llamacpp' ? 'http://127.0.0.1:8080' : fallback.endpoint)),
      model: String(value.model || ''),
      localModelPath: String(value.localModelPath || ''),
      temperature: Math.max(0, Math.min(2, Number(value.temperature ?? fallback.temperature))),
      actions,
      noteTypeInstructions: { ...fallback.noteTypeInstructions, ...(value.noteTypeInstructions ?? {}) },
    };
  }

  /** Serializa el estado completo en el almacén de ajustes de Electron. */
  private async persist(): Promise<void> {
    await window.electronAPI.settingsSet(SETTINGS_KEY, JSON.stringify(this.configuration()));
  }
}
