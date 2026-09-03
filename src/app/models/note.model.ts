/** Registro de nota compartido entre SQLite, IPC y el estado Angular. */
export interface Note {
  id: number;
  folder_id: number | null;
  title: string;
  content: string;
  /** Contenido binario Base64 para formatos como XLSX o PDF. */
  file_data?: string | null;
  /** Discriminante de documento; ausente representa texto enriquecido. */
  file_type?: string | null;
  created_at: string;
  updated_at: string;
}
