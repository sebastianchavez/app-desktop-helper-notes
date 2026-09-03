/** Proveedor local capaz de ejecutar el modelo seleccionado. */
export type AssistantProvider = 'ollama' | 'llamacpp';
/** Tipo lógico de documento usado para limitar acciones de un asistente. */
export type AssistantNoteType = 'text' | 'spreadsheet' | 'presentation' | 'pdf';
/** Operación que el motor de acciones sabe validar y ejecutar. */
export type AssistantActionMode =
  | 'analyze'
  | 'edit_note'
  | 'replace_text'
  | 'create_note'
  | 'create_folder'
  | 'edit_excel';

/** Regla configurable que autoriza una capacidad para determinados documentos. */
export interface AssistantActionDefinition {
  id: string;
  name: string;
  description: string;
  instructions: string;
  mode: AssistantActionMode;
  enabled: boolean;
  noteTypes: AssistantNoteType[];
}

/** Perfil completo que define personalidad, runtime, modelo y permisos. */
export interface AssistantProfile {
  id: string;
  name: string;
  role: string;
  description: string;
  systemPrompt: string;
  provider: AssistantProvider;
  endpoint: string;
  model: string;
  localModelPath?: string;
  temperature: number;
  actions: AssistantActionDefinition[];
  noteTypeInstructions: Partial<Record<AssistantNoteType, string>>;
}

/** Documento versionado persistido en el almacén local de ajustes. */
export interface AssistantConfiguration {
  version: 1;
  activeAssistantId: string;
  assistants: AssistantProfile[];
}

/** Configuración mínima enviada por IPC al cliente HTTP del proceso principal. */
export interface AssistantRuntimeConfig {
  provider: AssistantProvider;
  endpoint: string;
  temperature: number;
}

/** Metadatos comunes de un modelo descubierto en el proveedor. */
export interface AssistantModelInfo {
  name: string;
  size: number;
  modified_at: string;
}
