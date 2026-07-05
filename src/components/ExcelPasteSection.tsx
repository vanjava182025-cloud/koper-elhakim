import React, { useState } from "react";
import { 
  Clipboard, 
  Plus, 
  Trash2, 
  Check, 
  X, 
  Loader2, 
  Grid 
} from "lucide-react";
import { Pilgrim } from "../types";

interface ExcelPasteSectionProps {
  groupName: string;
  connectionMode: "demo" | "oauth" | "gas";
  spreadsheetId: string;
  gasUrl: string;
  token: string | null;
  onSuccess: () => void;
  onCancel: () => void;
}

interface EditableRow {
  nama: string;
  paspor: string;
  bus?: string;
}

export default function ExcelPasteSection({
  groupName,
  connectionMode,
  spreadsheetId,
  gasUrl,
  token,
  onSuccess,
  onCancel
}: ExcelPasteSectionProps) {
  const [pasteText, setPasteText] = useState("");
  const [rows, setRows] = useState<EditableRow[]>([
    { nama: "", paspor: "", bus: "" }
  ]);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Parse Excel (TSV) paste content
  const handlePasteChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    const text = e.target.value;
    setPasteText(text);

    if (!text.trim()) return;

    // Excel rows are separated by newlines, columns are separated by tabs
    const lines = text.split(/\r?\n/);
    const parsedRows: EditableRow[] = [];

    lines.forEach((line) => {
      if (!line.trim()) return;
      const cols = line.split("\t");
      
      const nama = (cols[0] || "").trim();
      const paspor = (cols[1] || "").trim();
      const bus = (cols[2] || "").trim();

      // Only add rows that have at least some content
      if (nama || paspor || bus) {
        parsedRows.push({ nama, paspor, bus });
      }
    });

    if (parsedRows.length > 0) {
      setRows(parsedRows);
      setError(null);
    }
  };

  // Add empty row
  const handleAddRow = () => {
    setRows([...rows, { nama: "", paspor: "", bus: "" }]);
  };

  // Delete specific row
  const handleDeleteRow = (index: number) => {
    const updated = [...rows];
    updated.splice(index, 1);
    setRows(updated.length > 0 ? updated : [{ nama: "", paspor: "", bus: "" }]);
  };

  // Handle cell edit
  const handleCellChange = (index: number, field: keyof EditableRow, value: string) => {
    const updated = [...rows];
    updated[index][field] = value;
    setRows(updated);
  };

  // Save all rows
  const handleSaveBulk = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    const validRows = rows.filter(r => r.nama.trim() !== "");
    if (validRows.length === 0) {
      setError("Silakan masukkan setidaknya satu jemaah dengan nama lengkap.");
      return;
    }

    setIsSaving(true);

    try {
      const response = await fetch("/api/add-pilgrims-bulk", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: token ? `Bearer ${token}` : ""
        },
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          grup: groupName,
          pilgrims: validRows
        })
      });

      const resData = await response.json();

      if (!response.ok || !resData.success) {
        throw new Error(resData.error || "Gagal menyimpan data jemaah bulk.");
      }

      // Success! Clear state and invoke reload
      setPasteText("");
      setRows([{ nama: "", paspor: "", bus: "" }]);
      onSuccess();
    } catch (err: any) {
      setError(err.message || "Koneksi ke server gagal.");
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="bg-slate-50/50 border border-slate-200/80 rounded-xl p-4 space-y-4 text-left shadow-inner">
      
      {/* Header Info */}
      <div className="flex justify-between items-start border-b border-slate-100 pb-3">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-[9px] bg-teal-600 px-2 py-0.5 rounded text-white font-extrabold uppercase tracking-wider">
              Bulk Input Excel
            </span>
            <span className="text-[10px] text-slate-500 font-bold font-mono">
              Grup: {groupName}
            </span>
          </div>
          <h3 className="text-sm font-bold text-slate-800 mt-1 flex items-center gap-1.5">
            <Grid className="h-4.5 w-4.5 text-teal-600" />
            <span>Salin-Tempel dari Excel</span>
          </h3>
          <p className="text-[11px] text-slate-500 mt-0.5 leading-relaxed">
            Buka file Excel Anda, salin kolom <strong>Nama Jemaah</strong>, <strong>Nomor Paspor</strong>, dan <strong>Bus (Opsional)</strong>, lalu tempelkan di bawah ini.
          </p>
        </div>
        <button
          type="button"
          onClick={onCancel}
          className="text-slate-400 hover:text-slate-600 p-1 rounded-lg hover:bg-slate-100/80 transition-colors cursor-pointer"
        >
          <X className="h-4 w-4" />
        </button>
      </div>

      {/* Paste Box Area */}
      <div className="space-y-2">
        <label className="text-xs font-bold text-slate-600 flex items-center gap-1.5">
          <Clipboard className="h-4 w-4 text-slate-500" />
          <span>Tempel Data Excel Di Sini:</span>
        </label>
        <textarea
          rows={3}
          value={pasteText}
          onChange={handlePasteChange}
          placeholder="Contoh salinan Excel:&#10;Muhammad Ali&#9;B1234567&#9;Bus 1&#10;Siti Aminah&#9;B7654321&#9;Bus 2"
          className="w-full font-mono text-xs border border-slate-200 rounded-lg p-3 bg-slate-50 hover:bg-slate-50/50 focus:bg-white focus:ring-2 focus:ring-emerald-500 outline-none text-slate-800 placeholder-slate-400 shadow-inner"
        ></textarea>
        <p className="text-[10px] text-slate-400">
          *Sistem akan otomatis mendeteksi pembatas tab ([TAB]) antar kolom (Nama [TAB] Paspor [TAB] Bus) dan baris baru.
        </p>
      </div>

      {/* Editable Table Grid (Excel Style) */}
      <form onSubmit={handleSaveBulk} className="space-y-4">
        <div className="space-y-2">
          <label className="text-xs font-bold text-slate-600">
            Preview Tabel Jemaah ({rows.length} Baris):
          </label>
          
          <div className="border border-slate-200 rounded-lg overflow-hidden max-h-[350px] overflow-y-auto overflow-x-auto w-full">
            <table className="w-full text-left border-collapse min-w-[500px]">
              <thead>
                <tr className="bg-slate-50 text-slate-500 text-[10px] uppercase font-bold tracking-wider border-b border-slate-200">
                  <th className="py-2 px-3 w-12 text-center border-r border-slate-200">No</th>
                  <th className="py-2 px-3 border-r border-slate-200">Nama Lengkap Jemaah</th>
                  <th className="py-2 px-3 border-r border-slate-200 w-[150px]">Nomor Paspor</th>
                  <th className="py-2 px-3 border-r border-slate-200 w-[120px]">Bus</th>
                  <th className="py-2 px-3 w-12 text-center">Aksi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-xs text-slate-700">
                {rows.map((row, idx) => (
                  <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                    <td className="py-1 px-2 text-center font-mono text-[10px] text-slate-400 border-r border-slate-200 bg-slate-50/30">
                      {idx + 1}
                    </td>
                    <td className="p-1 border-r border-slate-200">
                      <input
                        type="text"
                        required
                        placeholder="Masukkan nama jemaah..."
                        className="w-full px-2 py-1.5 focus:bg-white border-0 focus:ring-1 focus:ring-emerald-500 rounded outline-none text-slate-800"
                        value={row.nama}
                        onChange={(e) => handleCellChange(idx, "nama", e.target.value)}
                      />
                    </td>
                    <td className="p-1 border-r border-slate-200">
                      <input
                        type="text"
                        placeholder="B9876543 (Opsional)"
                        className="w-full font-mono px-2 py-1.5 focus:bg-white border-0 focus:ring-1 focus:ring-emerald-500 rounded outline-none text-slate-800 uppercase"
                        value={row.paspor}
                        onChange={(e) => handleCellChange(idx, "paspor", e.target.value)}
                      />
                    </td>
                    <td className="p-1 border-r border-slate-200">
                      <input
                        type="text"
                        placeholder="Contoh: Bus 1"
                        className="w-full px-2 py-1.5 focus:bg-white border-0 focus:ring-1 focus:ring-emerald-500 rounded outline-none text-slate-800"
                        value={row.bus || ""}
                        onChange={(e) => handleCellChange(idx, "bus", e.target.value)}
                      />
                    </td>
                    <td className="p-1 text-center">
                      <button
                        type="button"
                        onClick={() => handleDeleteRow(idx)}
                        className="text-slate-400 hover:text-rose-600 p-1.5 rounded hover:bg-slate-100 transition-colors cursor-pointer"
                        title="Hapus baris"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Buttons */}
        <div className="flex flex-col sm:flex-row justify-between items-center gap-3 pt-2">
          <button
            type="button"
            onClick={handleAddRow}
            className="w-full sm:w-auto bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 px-4 py-2 rounded-lg text-xs font-bold smooth-transition flex items-center justify-center gap-1.5 cursor-pointer"
          >
            <Plus className="h-4 w-4" />
            <span>Tambah Baris Baru</span>
          </button>

          {error && (
            <div className="text-xs text-rose-600 font-semibold bg-rose-50 px-3 py-1.5 rounded-lg border border-rose-200/50 max-w-full truncate">
              {error}
            </div>
          )}

          <div className="flex items-center gap-2 w-full sm:w-auto">
            <button
              type="button"
              onClick={onCancel}
              className="w-1/2 sm:w-auto bg-slate-50 hover:bg-slate-100 text-slate-500 px-4 py-2 rounded-lg text-xs font-semibold border border-slate-200 cursor-pointer"
            >
              Batal
            </button>
            <button
              type="submit"
              disabled={isSaving}
              className="w-1/2 sm:w-auto bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-2 rounded-lg text-xs font-bold smooth-transition flex items-center justify-center gap-1.5 cursor-pointer shadow-md disabled:bg-slate-300"
            >
              {isSaving ? (
                <>
                  <Loader2 className="h-4 w-4 animate-spin" />
                  <span>Menyimpan ke Sheet...</span>
                </>
              ) : (
                <>
                  <Check className="h-4 w-4" />
                  <span>Simpan Semua Jemaah ({rows.filter(r => r.nama.trim() !== "").length})</span>
                </>
              )}
            </button>
          </div>
        </div>
      </form>

    </div>
  );
}
