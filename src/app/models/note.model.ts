export interface Note {
  id: number;
  folder_id: number | null;
  title: string;
  content: string;
  file_data?: string | null;
  file_type?: string | null;
  created_at: string;
  updated_at: string;
}
