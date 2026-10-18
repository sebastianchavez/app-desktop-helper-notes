import path from 'path';
import { isSupportedDocument } from './document-types';

let pendingPaths: string[] = [];

/**
 * Guarda solicitudes compatibles hasta que el renderer pueda consumirlas.
 * @param args Argumentos de apertura sin el ejecutable ni la ruta de desarrollo.
 * @param cwd Directorio de trabajo del proceso que envió la solicitud.
 * @returns `true` si recibió al menos una ruta compatible; ignora opciones y duplicados pendientes.
 */
export function queueOpenFiles(args: string[], cwd = process.cwd()): boolean {
  const paths = args
    .filter((arg) => !arg.startsWith('-') && isSupportedDocument(arg))
    .map((arg) => path.resolve(cwd, arg));
  pendingPaths = [...new Set([...pendingPaths, ...paths])];
  return paths.length > 0;
}

/**
 * Consume las solicitudes pendientes una sola vez.
 * @returns Rutas absolutas pendientes; vacía la cola antes de devolverlas.
 */
export function takePendingOpenFiles(): string[] {
  const paths = pendingPaths;
  pendingPaths = [];
  return paths;
}
