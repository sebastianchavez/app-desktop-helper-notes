import http from 'http';
import https from 'https';

export type AiProvider = 'ollama' | 'llamacpp';

export interface AiRuntimeConfig {
  provider: AiProvider;
  endpoint: string;
  temperature?: number;
}

export interface AiMessage {
  role: string;
  content: string;
}

export interface AiModelInfo {
  name: string;
  size: number;
  modified_at: string;
}

const NON_CHAT_PATTERNS = /\b(embed|embedding|all-minilm|bge-|nomic-embed|mxbae|snowflake-arctic-embed)\b/i;
// Una solicitud activa por renderer permite cancelar el stream correcto sin afectar otras ventanas.
const activeRequests = new Map<number, http.ClientRequest>();

/** Resuelve una ruta relativa contra el endpoint validando que use HTTP(S). */
function endpointUrl(config: AiRuntimeConfig, pathname: string): URL {
  const endpoint = new URL(config.endpoint.trim());
  if (endpoint.protocol !== 'http:' && endpoint.protocol !== 'https:') {
    throw new Error('El endpoint debe usar HTTP o HTTPS');
  }
  const basePath = endpoint.pathname.endsWith('/') ? endpoint.pathname : `${endpoint.pathname}/`;
  endpoint.pathname = `${basePath}${pathname.replace(/^\//, '')}`.replace(/\/+/g, '/');
  endpoint.search = '';
  endpoint.hash = '';
  return endpoint;
}

/** Ejecuta una solicitud JSON corta con timeout y errores HTTP normalizados. */
function requestJson(config: AiRuntimeConfig, pathname: string, method = 'GET', body?: unknown): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const target = endpointUrl(config, pathname);
    const payload = body === undefined ? undefined : JSON.stringify(body);
    const transport = target.protocol === 'https:' ? https : http;
    const request = transport.request(target, {
      method,
      headers: payload ? {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      } : undefined,
      timeout: 8000,
    }, response => {
      let raw = '';
      response.setEncoding('utf8');
      response.on('data', chunk => raw += chunk);
      response.on('end', () => {
        if ((response.statusCode ?? 500) >= 400) {
          reject(new Error(`El proveedor respondió con HTTP ${response.statusCode}`));
          return;
        }
        try {
          resolve(raw ? JSON.parse(raw) : {});
        } catch {
          resolve(raw);
        }
      });
    });
    request.on('timeout', () => request.destroy(new Error('El proveedor no respondió a tiempo')));
    request.on('error', reject);
    if (payload) request.write(payload);
    request.end();
  });
}

/** Comprueba Ollama o llama.cpp usando los endpoints de salud disponibles. */
export async function checkAiProvider(config: AiRuntimeConfig): Promise<{ available: boolean; error?: string }> {
  try {
    if (config.provider === 'ollama') {
      await requestJson(config, '/api/tags');
    } else {
      try {
        await requestJson(config, '/health');
      } catch {
        await requestJson(config, '/v1/models');
      }
    }
    return { available: true };
  } catch (error) {
    return { available: false, error: error instanceof Error ? error.message : String(error) };
  }
}

/** Lista modelos de chat y excluye modelos de embeddings incompatibles. */
export async function listAiModels(config: AiRuntimeConfig): Promise<{ success: boolean; models: AiModelInfo[]; error?: string }> {
  try {
    const data = await requestJson(config, config.provider === 'ollama' ? '/api/tags' : '/v1/models') as Record<string, unknown>;
    const source = Array.isArray(data['models'])
      ? data['models']
      : Array.isArray(data['data']) ? data['data'] : [];
    const models = source.map(item => {
      const row = item as Record<string, unknown>;
      return {
        name: String(row['name'] ?? row['id'] ?? ''),
        size: Number(row['size'] ?? 0),
        modified_at: String(row['modified_at'] ?? ''),
      };
    }).filter(model => model.name && !NON_CHAT_PATTERNS.test(model.name));
    return { success: true, models };
  } catch (error) {
    return { success: false, models: [], error: error instanceof Error ? error.message : String(error) };
  }
}

/**
 * Unifica el streaming NDJSON de Ollama y SSE compatible con OpenAI de
 * llama.cpp. Cada fragmento de texto se entrega mediante `onToken`.
 */
export function streamAiChat(
  senderId: number,
  config: AiRuntimeConfig,
  model: string,
  messages: AiMessage[],
  onToken: (token: string) => void,
): Promise<void> {
  abortAiChat(senderId);
  return new Promise((resolve, reject) => {
    const isOllama = config.provider === 'ollama';
    const target = endpointUrl(config, isOllama ? '/api/chat' : '/v1/chat/completions');
    const payload = JSON.stringify({
      model,
      messages,
      stream: true,
      ...(isOllama
        ? { options: { temperature: config.temperature ?? 0.7 } }
        : { temperature: config.temperature ?? 0.7 }),
    });
    const transport = target.protocol === 'https:' ? https : http;
    const request = transport.request(target, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
        'Accept': isOllama ? 'application/x-ndjson' : 'text/event-stream',
      },
    }, response => {
      if ((response.statusCode ?? 500) >= 400) {
        let errorBody = '';
        response.on('data', chunk => errorBody += chunk.toString());
        response.on('end', () => reject(new Error(errorBody || `HTTP ${response.statusCode}`)));
        return;
      }

      let buffer = '';
      response.setEncoding('utf8');
      response.on('data', chunk => {
        buffer += chunk;
        const lines = buffer.split('\n');
        buffer = lines.pop() ?? '';
        for (const rawLine of lines) {
          const line = rawLine.trim();
          if (!line) continue;
          const jsonLine = isOllama ? line : line.replace(/^data:\s*/, '');
          if (jsonLine === '[DONE]') continue;
          try {
            const data = JSON.parse(jsonLine) as Record<string, unknown>;
            if (isOllama) {
              const message = data['message'] as Record<string, unknown> | undefined;
              if (message?.['content']) onToken(String(message['content']));
            } else {
              const choices = data['choices'] as Array<Record<string, unknown>> | undefined;
              const delta = choices?.[0]?.['delta'] as Record<string, unknown> | undefined;
              if (delta?.['content']) onToken(String(delta['content']));
            }
          } catch {
            // A partial or informational SSE line can be ignored safely.
          }
        }
      });
      response.on('end', () => {
        activeRequests.delete(senderId);
        resolve();
      });
      response.on('error', error => {
        activeRequests.delete(senderId);
        reject(error);
      });
    });
    request.on('error', error => {
      activeRequests.delete(senderId);
      reject(error);
    });
    activeRequests.set(senderId, request);
    request.write(payload);
    request.end();
  });
}

/** Cancela y retira la solicitud activa perteneciente a una ventana renderer. */
export function abortAiChat(senderId: number): void {
  const request = activeRequests.get(senderId);
  if (request) {
    request.destroy();
    activeRequests.delete(senderId);
  }
}
