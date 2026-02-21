export interface Note {
  id: number;
  folder_id: number | null;
  title: string;
  content: string;
  created_at: string;
  updated_at: string;
}
