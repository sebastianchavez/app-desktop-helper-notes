import path from 'path';

/** Devuelve un icono compatible con el formato nativo de cada plataforma. */
export function applicationIconPath(): string {
  const fileName = process.platform === 'win32' ? 'favicon.ico' : 'favicon.png';
  return path.join(__dirname, '../assets', fileName);
}
