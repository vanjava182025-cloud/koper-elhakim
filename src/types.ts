export interface Pilgrim {
  grup: string;
  nama: string;
  paspor: string;
  bus?: string;
}

export interface LogEntry {
  timestamp: string;
  grup: string;
  nama: string;
  paspor: string;
  pos: string; // Historically 'pos', but we display as 'Tempat Cheklist'
  koperBesar: string; // Display string, e.g. "2 Pcs" or "Tidak Ada"
  koperKecil: string; // Display string, e.g. "1 Pcs" or "Tidak Ada"
  petugas: string;
  status: string;
}

export interface AdditionalItem {
  name: string;
  qty: number;
}
