import { Component, OnDestroy, OnInit, inject, signal } from '@angular/core';
import { AssistantConfigService } from '../../services/assistant-config.service';
import {
  AssistantActionDefinition,
  AssistantActionMode,
  AssistantNoteType,
  AssistantProfile,
  AssistantProvider,
} from '../../models/assistant.model';

interface OllamaSuggestion {
  id: string;
  name: string;
  detail: string;
  cloud: boolean;
}

@Component({
  selector: 'app-assistant-settings',
  templateUrl: './assistant-settings.html',
  styleUrl: './assistant-settings.scss',
})
/**
 * Editor de perfiles de asistentes y centro de administración de runtimes IA.
 *
 * Permite configurar acciones y tipos de nota, administrar Ollama y descargar
 * tanto llama.cpp como modelos GGUF sin exponer APIs de Node al renderer.
 */
export class AssistantSettings implements OnInit, OnDestroy {
  protected readonly configService = inject(AssistantConfigService);
  protected selectedId = signal('');
  protected statusMessage = signal('');
  protected llamaStatus = signal<{ installed: boolean; running: boolean; executable?: string; models: { name: string; path: string; size: number }[] } | null>(null);
  protected llamaBusy = signal(false);
  protected llamaError = signal('');
  protected ollamaInstalled = signal(false);
  protected ollamaRunning = signal(false);
  protected ollamaModels = signal<{ name: string; size: number; modified_at: string }[]>([]);
  protected ollamaBusy = signal(false);
  protected ollamaModelBusy = signal('');
  protected ollamaPullProgress = signal<{ model: string; status: string; percent: number } | null>(null);
  protected downloadProgress = signal<{ percent: number; message: string; id: string } | null>(null);
  protected modelSearch = signal('');
  protected catalogModels = signal<{ id: string; downloads: number; likes: number; updatedAt: string }[]>([]);
  protected catalogLoading = signal(false);
  protected selectedRepository = signal('');
  protected repositoryFiles = signal<{ name: string; size: number; quantization: string }[]>([]);
  private removeProgressListener: (() => void) | null = null;
  private removeOllamaProgressListener: (() => void) | null = null;

  protected readonly ollamaLocalSuggestions: OllamaSuggestion[] = [
    { id: 'llama3.2:3b', name: 'Llama 3.2 · 3B', detail: 'Ligero y adecuado para notas y tareas generales.', cloud: false },
    { id: 'nemotron-3-nano:4b', name: 'Nemotron 3 Nano · 4B', detail: 'Razonamiento y herramientas con un tamaño moderado.', cloud: false },
    { id: 'qwen3.5:4b', name: 'Qwen 3.5 · 4B', detail: 'Modelo compacto, multilingüe y orientado a herramientas.', cloud: false },
    { id: 'gemma4:e2b', name: 'Gemma 4 · E2B', detail: 'Modelo multimodal eficiente para equipos personales.', cloud: false },
  ];
  protected readonly ollamaCloudSuggestions: OllamaSuggestion[] = [
    { id: 'gemma4:31b-cloud', name: 'Gemma 4 · 31B', detail: 'Cloud gratuito · uso bajo · contexto 256K.', cloud: true },
    { id: 'gpt-oss:20b-cloud', name: 'GPT-OSS · 20B', detail: 'Cloud gratuito · uso bajo · contexto 128K.', cloud: true },
    { id: 'gpt-oss:120b-cloud', name: 'GPT-OSS · 120B', detail: 'Cloud gratuito · uso medio · contexto 128K.', cloud: true },
    { id: 'nemotron-3-nano:30b-cloud', name: 'Nemotron 3 Nano · 30B', detail: 'Cloud gratuito · uso bajo · contexto 1M.', cloud: true },
    { id: 'nemotron-3-super:cloud', name: 'Nemotron 3 Super', detail: 'Cloud gratuito · uso medio · contexto 256K.', cloud: true },
    { id: 'nemotron-3-ultra:cloud', name: 'Nemotron 3 Ultra', detail: 'Cloud gratuito · uso alto · contexto 256K.', cloud: true },
  ];

  protected readonly providers: { value: AssistantProvider; label: string; endpoint: string }[] = [
    { value: 'ollama', label: 'Ollama', endpoint: 'http://127.0.0.1:11434' },
    { value: 'llamacpp', label: 'llama.cpp (llama-server)', endpoint: 'http://127.0.0.1:8080' },
  ];
  protected readonly actionModes: { value: AssistantActionMode; label: string }[] = [
    { value: 'analyze', label: 'Analizar / responder' },
    { value: 'edit_note', label: 'Editar nota completa' },
    { value: 'replace_text', label: 'Reemplazar texto' },
    { value: 'create_note', label: 'Crear nota' },
    { value: 'create_folder', label: 'Crear carpeta' },
    { value: 'edit_excel', label: 'Editar celdas' },
  ];
  protected readonly noteTypes: { value: AssistantNoteType; label: string }[] = [
    { value: 'text', label: 'Nota de texto' },
    { value: 'spreadsheet', label: 'Hoja de cálculo' },
    { value: 'presentation', label: 'Presentación' },
    { value: 'pdf', label: 'PDF' },
  ];

  /** Carga perfiles, registra eventos de progreso y consulta el proveedor activo. */
  async ngOnInit(): Promise<void> {
    await this.configService.init();
    this.selectedId.set(this.configService.activeAssistant()?.id ?? '');
    this.removeProgressListener = window.electronAPI.onLlamaProgress(progress => {
      this.downloadProgress.set({ percent: progress.percent, message: progress.message, id: progress.id });
    });
    this.removeOllamaProgressListener = window.electronAPI.onOllamaPullProgress(progress => {
      this.ollamaPullProgress.set({ model: progress.model, status: progress.status, percent: progress.percent });
    });
    if (this.selectedAssistant()?.provider === 'llamacpp') {
      await this.refreshLlamaStatus();
      await this.searchModels();
    } else {
      await this.refreshOllamaStatus();
    }
  }

  /** Retira listeners IPC de descargas para impedir duplicados al volver a la vista. */
  ngOnDestroy(): void {
    this.removeProgressListener?.();
    this.removeOllamaProgressListener?.();
  }

  protected selectedAssistant(): AssistantProfile | undefined {
    return this.configService.assistants().find(item => item.id === this.selectedId());
  }

  async selectAssistant(id: string): Promise<void> {
    this.selectedId.set(id);
    await this.configService.setActiveAssistant(id);
    if (this.selectedAssistant()?.provider === 'llamacpp') await this.refreshLlamaStatus();
    else await this.refreshOllamaStatus();
  }

  async addAssistant(): Promise<void> {
    const assistant = await this.configService.addAssistant();
    this.selectedId.set(assistant.id);
  }

  async duplicateAssistant(): Promise<void> {
    const copy = await this.configService.duplicateAssistant(this.selectedId());
    if (copy) this.selectedId.set(copy.id);
  }

  async deleteAssistant(): Promise<void> {
    const assistant = this.selectedAssistant();
    if (!assistant) return;
    const confirmed = await window.electronAPI.dialogConfirm({
      title: 'Eliminar asistente',
      message: `¿Eliminar el asistente “${assistant.name}”?`,
    });
    if (!confirmed) return;
    const deleted = await this.configService.deleteAssistant(assistant.id);
    if (!deleted) {
      this.showStatus('Debe existir al menos un asistente.');
      return;
    }
    this.selectedId.set(this.configService.activeAssistant()?.id ?? '');
  }

  /** Persiste un campo genérico del perfil actualmente seleccionado. */
  async updateField<K extends keyof AssistantProfile>(field: K, value: AssistantProfile[K]): Promise<void> {
    const id = this.selectedId();
    if (!id) return;
    await this.configService.updateAssistant(id, { [field]: value } as Pick<AssistantProfile, K>);
    this.showStatus('Configuración guardada');
  }

  /** Cambia de runtime, restablece el modelo y refresca sus capacidades. */
  async changeProvider(provider: AssistantProvider): Promise<void> {
    const option = this.providers.find(item => item.value === provider);
    await this.configService.updateAssistant(this.selectedId(), {
      provider,
      endpoint: option?.endpoint ?? '',
      model: '',
    });
    this.showStatus('Proveedor actualizado');
    if (provider === 'llamacpp') {
      await this.refreshLlamaStatus();
      if (this.catalogModels().length === 0) await this.searchModels();
    } else {
      await this.refreshOllamaStatus();
    }
  }

  protected modelOptions(): { name: string; label: string }[] {
    const assistant = this.selectedAssistant();
    if (!assistant) return [];
    const available = assistant.provider === 'ollama'
      ? this.ollamaModels().map(model => ({ name: model.name, label: model.name }))
      : (this.llamaStatus()?.models ?? []).map(model => ({ name: model.name, label: `${model.name} · ${this.formatBytes(model.size)}` }));
    if (assistant.model && !available.some(model => model.name === assistant.model)) {
      return [{ name: assistant.model, label: `${assistant.model} (no disponible)` }, ...available];
    }
    return available;
  }

  async selectAvailableModel(modelName: string): Promise<void> {
    const assistant = this.selectedAssistant();
    if (!assistant || !modelName) return;
    if (assistant.provider === 'llamacpp') {
      const localModel = this.llamaStatus()?.models.find(model => model.name === modelName);
      if (localModel) {
        await this.selectLocalModel(localModel);
        return;
      }
    }
    await this.configService.updateAssistant(assistant.id, { model: modelName, localModelPath: '' });
    this.showStatus('Modelo seleccionado');
  }

  /** Consulta instalación, proceso y catálogo local de Ollama. */
  async refreshOllamaStatus(): Promise<void> {
    this.ollamaBusy.set(true);
    this.llamaError.set('');
    try {
      const installation = await window.electronAPI.ollamaCheckInstalled();
      this.ollamaInstalled.set(installation.installed);
      const status = await window.electronAPI.ollamaStatus();
      this.ollamaRunning.set(status.running);
      if (status.running) {
        const result = await window.electronAPI.ollamaListModels();
        if (!result.success) throw new Error(result.error || 'No se pudieron cargar los modelos de Ollama');
        this.ollamaModels.set(result.models ?? []);
      } else {
        this.ollamaModels.set([]);
      }
    } catch (error) {
      this.ollamaRunning.set(false);
      this.ollamaModels.set([]);
      this.llamaError.set(error instanceof Error ? error.message : String(error));
    } finally {
      this.ollamaBusy.set(false);
    }
  }

  async startOllama(): Promise<void> {
    this.ollamaBusy.set(true);
    this.llamaError.set('');
    try {
      const result = await window.electronAPI.ollamaStart();
      if (!result.success) throw new Error(result.error || 'No se pudo iniciar Ollama');
      await this.refreshOllamaStatus();
      this.showStatus('Ollama iniciado');
    } catch (error) {
      this.llamaError.set(error instanceof Error ? error.message : String(error));
    } finally {
      this.ollamaBusy.set(false);
    }
  }

  async stopOllama(): Promise<void> {
    this.ollamaBusy.set(true);
    this.llamaError.set('');
    try {
      const result = await window.electronAPI.ollamaStop();
      if (!result.success) throw new Error(result.error || 'No se pudo detener Ollama');
      this.ollamaRunning.set(false);
      this.ollamaModels.set([]);
      this.showStatus('Ollama detenido');
    } catch (error) {
      this.llamaError.set(error instanceof Error ? error.message : String(error));
    } finally {
      this.ollamaBusy.set(false);
    }
  }

  async openOllamaDownload(): Promise<void> {
    await window.electronAPI.ollamaOpenDownload();
  }

  protected isOllamaModelInstalled(model: string): boolean {
    return this.ollamaModels().some(item => item.name === model);
  }

  /** Descarga un modelo Ollama y refleja el progreso emitido por Electron. */
  async installOllamaModel(suggestion: OllamaSuggestion): Promise<void> {
    if (!this.ollamaRunning()) {
      this.llamaError.set('Inicia Ollama antes de agregar un modelo.');
      return;
    }
    this.ollamaModelBusy.set(suggestion.id);
    this.ollamaPullProgress.set({ model: suggestion.id, status: suggestion.cloud ? 'Agregando modelo cloud' : 'Preparando descarga', percent: 0 });
    this.llamaError.set('');
    try {
      const result = await window.electronAPI.ollamaPull(suggestion.id);
      if (!result.success) {
        const signinHelp = suggestion.cloud ? ' Inicia sesión en Ollama para usar modelos cloud.' : '';
        throw new Error((result.error || 'No se pudo agregar el modelo.') + signinHelp);
      }
      await this.refreshOllamaStatus();
      await this.configService.updateAssistant(this.selectedId(), { model: suggestion.id, localModelPath: '' });
      this.showStatus(suggestion.cloud ? 'Modelo cloud agregado y seleccionado' : 'Modelo instalado y seleccionado');
    } catch (error) {
      this.llamaError.set(error instanceof Error ? error.message : String(error));
    } finally {
      this.ollamaModelBusy.set('');
      this.ollamaPullProgress.set(null);
    }
  }

  async signinOllama(): Promise<void> {
    const result = await window.electronAPI.ollamaSignin();
    if (!result.success) {
      this.llamaError.set(result.error || 'No se pudo iniciar el acceso a Ollama.');
      return;
    }
    this.showStatus('Completa el inicio de sesión de Ollama en el navegador');
  }

  /** Consulta instalación, proceso y modelos GGUF administrados por la aplicación. */
  async refreshLlamaStatus(): Promise<void> {
    const assistant = this.selectedAssistant();
    if (!assistant) return;
    try {
      this.llamaStatus.set(await window.electronAPI.llamaManagerStatus(assistant.endpoint));
    } catch (error) {
      this.llamaError.set(error instanceof Error ? error.message : String(error));
    }
  }

  /** Descarga e instala el binario llama.cpp apropiado para la plataforma. */
  async installLlamaCpp(): Promise<void> {
    this.llamaBusy.set(true);
    this.llamaError.set('');
    this.downloadProgress.set({ percent: 0, message: 'Buscando la última versión', id: 'runtime' });
    try {
      const result = await window.electronAPI.llamaInstall();
      if (!result.success) throw new Error(result.error || 'No se pudo instalar llama.cpp');
      this.showStatus(`llama.cpp ${result.version || ''} instalado`);
      await this.refreshLlamaStatus();
    } catch (error) {
      this.llamaError.set(error instanceof Error ? error.message : String(error));
    } finally {
      this.llamaBusy.set(false);
      this.downloadProgress.set(null);
    }
  }

  /** Busca repositorios GGUF en el catálogo remoto mediante el proceso principal. */
  async searchModels(): Promise<void> {
    this.catalogLoading.set(true);
    this.llamaError.set('');
    try {
      const result = await window.electronAPI.llamaCatalog(this.modelSearch());
      if (!result.success) throw new Error(result.error || 'No se pudo consultar Hugging Face');
      this.catalogModels.set(result.models);
    } catch (error) {
      this.llamaError.set(error instanceof Error ? error.message : String(error));
    } finally {
      this.catalogLoading.set(false);
    }
  }

  async openRepository(repository: string): Promise<void> {
    this.selectedRepository.set(repository);
    this.repositoryFiles.set([]);
    this.catalogLoading.set(true);
    try {
      const result = await window.electronAPI.llamaModelFiles(repository);
      if (!result.success) throw new Error(result.error || 'No se pudieron cargar las cuantizaciones');
      this.repositoryFiles.set(result.files);
    } catch (error) {
      this.llamaError.set(error instanceof Error ? error.message : String(error));
    } finally {
      this.catalogLoading.set(false);
    }
  }

  /** Descarga un archivo GGUF del repositorio seleccionado. */
  async downloadModel(fileName: string): Promise<void> {
    const repository = this.selectedRepository();
    if (!repository) return;
    this.llamaBusy.set(true);
    this.llamaError.set('');
    this.downloadProgress.set({ percent: 0, message: 'Iniciando descarga', id: fileName });
    try {
      const result = await window.electronAPI.llamaDownloadModel(repository, fileName);
      if (!result.success || !result.path) throw new Error(result.error || 'No se pudo descargar el modelo');
      await this.configService.updateAssistant(this.selectedId(), {
        model: fileName,
        localModelPath: result.path,
      });
      await this.refreshLlamaStatus();
      this.showStatus('Modelo descargado y seleccionado');
    } catch (error) {
      this.llamaError.set(error instanceof Error ? error.message : String(error));
    } finally {
      this.llamaBusy.set(false);
      this.downloadProgress.set(null);
    }
  }

  async cancelDownload(): Promise<void> {
    await window.electronAPI.llamaCancelDownload();
    this.llamaBusy.set(false);
    this.downloadProgress.set(null);
  }

  async selectLocalModel(model: { name: string; path: string }): Promise<void> {
    await this.configService.updateAssistant(this.selectedId(), { model: model.name, localModelPath: model.path });
    this.showStatus('Modelo seleccionado');
  }

  /** Arranca llama-server con el modelo local configurado en el perfil. */
  async startLlamaServer(): Promise<void> {
    const assistant = this.selectedAssistant();
    if (!assistant?.localModelPath) {
      this.llamaError.set('Selecciona o descarga un modelo GGUF primero.');
      return;
    }
    this.llamaBusy.set(true);
    this.llamaError.set('');
    try {
      const result = await window.electronAPI.llamaStart(assistant.localModelPath, assistant.endpoint);
      if (!result.success) throw new Error(result.error || 'No se pudo iniciar llama-server');
      await this.refreshLlamaStatus();
      this.showStatus('llama-server iniciado');
    } catch (error) {
      this.llamaError.set(error instanceof Error ? error.message : String(error));
    } finally {
      this.llamaBusy.set(false);
    }
  }

  /** Detiene el proceso llama-server administrado. */
  async stopLlamaServer(): Promise<void> {
    this.llamaBusy.set(true);
    const result = await window.electronAPI.llamaStop();
    if (!result.success) this.llamaError.set(result.error || 'No se pudo detener llama-server');
    await this.refreshLlamaStatus();
    this.llamaBusy.set(false);
  }

  protected formatBytes(bytes: number): string {
    if (!bytes) return 'Tamaño desconocido';
    const gb = bytes / (1024 ** 3);
    return gb >= 1 ? `${gb.toFixed(1)} GB` : `${Math.round(bytes / (1024 ** 2))} MB`;
  }

  async updateNoteInstruction(type: AssistantNoteType, value: string): Promise<void> {
    const assistant = this.selectedAssistant();
    if (!assistant) return;
    await this.updateField('noteTypeInstructions', {
      ...assistant.noteTypeInstructions,
      [type]: value,
    });
  }

  async addAction(): Promise<void> {
    await this.configService.addAction(this.selectedId());
  }

  async updateAction(actionId: string, changes: Partial<AssistantActionDefinition>): Promise<void> {
    await this.configService.updateAction(this.selectedId(), actionId, changes);
    this.showStatus('Acción guardada');
  }

  async toggleActionNoteType(action: AssistantActionDefinition, type: AssistantNoteType, enabled: boolean): Promise<void> {
    const noteTypes = enabled
      ? Array.from(new Set([...action.noteTypes, type]))
      : action.noteTypes.filter(item => item !== type);
    await this.updateAction(action.id, { noteTypes });
  }

  async deleteAction(actionId: string): Promise<void> {
    await this.configService.deleteAction(this.selectedId(), actionId);
  }

  private showStatus(message: string): void {
    this.statusMessage.set(message);
    setTimeout(() => {
      if (this.statusMessage() === message) this.statusMessage.set('');
    }, 1800);
  }
}
