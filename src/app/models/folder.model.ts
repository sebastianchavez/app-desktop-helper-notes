/** Nodo persistido del árbol de carpetas; `parent_id = null` indica una raíz. */
export interface Folder {
  id: number;
  name: string;
  parent_id: number | null;
  created_at: string;
  updated_at: string;
}
