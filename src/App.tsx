/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from "react";
import { 
  Luggage, 
  Database, 
  Plus, 
  CheckCircle2, 
  AlertTriangle, 
  Loader2, 
  LogOut, 
  FileSpreadsheet, 
  Check, 
  MapPin, 
  Sparkles, 
  Clock, 
  User, 
  RefreshCw, 
  Bell, 
  Camera, 
  UploadCloud, 
  Search, 
  CheckCircle,
  HelpCircle,
  Info,
  Users,
  Folder,
  ArrowLeft,
  ChevronRight,
  Bus,
  Edit2,
  Trash2,
  FileCode,
  Smartphone,
  Download,
  Chrome
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import { googleSignIn, initAuth, logout } from "./firebase";
import { User as FirebaseUser } from "firebase/auth";

// Define TypeScript Types and Component/Utility Imports
import { Pilgrim, LogEntry, AdditionalItem } from "./types";
import ExcelPasteSection from "./components/ExcelPasteSection";
import { generatePdfReport } from "./utils/pdfGenerator";

const gasCodeTemplate = `/**
 * GOOGLE APPS SCRIPT - SINKRONISASI SISTEM CHECKLIST KOPER JEMAAH
 * 
 * SCRIPT INI HARUS DI-DEPLOY SEBAGAI WEB APP DI GOOGLE APPS SCRIPT:
 * 1. Buka Google Sheets Anda.
 * 2. Klik menu Ekstensi -> Apps Script.
 * 3. Hapus semua kode default, dan paste kode di bawah ini.
 * 4. Klik ikon "Simpan" (Ctrl+S atau Cmd+S).
 * 5. Klik tombol "Terapkan" (Deploy) -> "Penerapan baru" (New deployment).
 * 6. Pilih tipe: "Aplikasi web" (Web app).
 * 7. Setel konfigurasi:
 *    - Jalankan sebagai: "Saya" (Me / akun Google Anda)
 *    - Siapa yang memiliki akses: "Siapa saja" (Anyone)
 * 8. Klik "Terapkan" (Deploy).
 * 9. Berikan izin akses penuh ke akun Google Sheets Anda jika diminta.
 * 10. Salin "URL Aplikasi Web" yang dihasilkan dan tempelkan ke aplikasi.
 */

var TAB_MANIFEST = "Manifest";
var TAB_POS = "Pos";
var TAB_LOG = "Log";
var TAB_GRUP = "Grup";

function inisialisasiSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sManifest = ss.getSheetByName(TAB_MANIFEST);
  if (!sManifest) {
    sManifest = ss.insertSheet(TAB_MANIFEST);
    sManifest.appendRow(["Nama", "Paspor", "Grup", "Bus"]);
    sManifest.getRange("A1:D1").setFontWeight("bold").setBackground("#e2e8f0");
  }
  
  var sPos = ss.getSheetByName(TAB_POS);
  if (!sPos) {
    sPos = ss.insertSheet(TAB_POS);
    sPos.appendRow(["Nama Pos"]);
    sPos.getRange("A1").setFontWeight("bold").setBackground("#e2e8f0");
    sPos.appendRow(["Pos Bandara Soekarno-Hatta"]);
  }
  
  var sLog = ss.getSheetByName(TAB_LOG);
  if (!sLog) {
    sLog = ss.insertSheet(TAB_LOG);
    sLog.appendRow(["Timestamp", "Grup", "Nama", "Paspor", "Pos", "Koper Besar", "Koper Kecil", "Petugas", "Status"]);
    sLog.getRange("A1:I1").setFontWeight("bold").setBackground("#e2e8f0");
  }

  var sGrup = ss.getSheetByName(TAB_GRUP);
  if (!sGrup) {
    sGrup = ss.insertSheet(TAB_GRUP);
    sGrup.appendRow(["Nama Grup"]);
    sGrup.getRange("A1").setFontWeight("bold").setBackground("#e2e8f0");
  }
}

function doGet(e) {
  inisialisasiSheet();
  if (!e || !e.parameter) {
    return ContentService.createTextOutput("Koneksi sukses! Web App Google Apps Script Anda telah aktif. Gunakan URL ini di aplikasi Anda untuk memulai sinkronisasi.");
  }
  var action = e.parameter.action;
  if (action === "ambilDataGrupDanPos") {
    return outputJSON(ambilDataGrupDanPos());
  }
  return outputJSON({ success: false, error: "Action GET tidak valid atau kosong" });
}

function doPost(e) {
  inisialisasiSheet();
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return outputJSON({ success: false, error: "Body POST kosong atau tidak valid." });
    }
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    if (!action) {
      return outputJSON({ success: false, error: "Parameter 'action' tidak ditemukan" });
    }
    switch (action) {
      case "tambahPosChecklist":
        return outputJSON(tambahPosChecklist(requestData));
      case "simpanLogChecklist":
        return outputJSON(simpanLogChecklist(requestData));
      case "tambahGrup":
        return outputJSON(tambahGrup(requestData));
      case "tambahJemaah":
        return outputJSON(tambahJemaah(requestData));
      case "tambahJemaahBulk":
        return outputJSON(tambahJemaahBulk(requestData));
      case "editGrup":
        return outputJSON(editGrup(requestData));
      case "hapusGrup":
        return outputJSON(hapusGrup(requestData));
      case "editPos":
        return outputJSON(editPos(requestData));
      case "hapusPos":
        return outputJSON(hapusPos(requestData));
      case "editLogChecklist":
        return outputJSON(editLogChecklist(requestData));
      case "hapusLogChecklist":
        return outputJSON(hapusLogChecklist(requestData));
      default:
        return outputJSON({ success: false, error: "Action POST tidak dikenali: " + action });
    }
  } catch (err) {
    return outputJSON({ success: false, error: "Terjadi kesalahan internal: " + err.toString() });
  }
}

function outputJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function ambilDataGrupDanPos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sGrup = ss.getSheetByName(TAB_GRUP);
  var groups = [];
  if (sGrup) {
    var data = sGrup.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) groups.push(data[i][0].toString().trim());
    }
  }
  
  var sPos = ss.getSheetByName(TAB_POS);
  var stations = [];
  if (sPos) {
    var data = sPos.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) stations.push(data[i][0].toString().trim());
    }
  }
  
  var sManifest = ss.getSheetByName(TAB_MANIFEST);
  var manifest = [];
  if (sManifest) {
    var data = sManifest.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] || data[i][1]) {
        manifest.push({
          nama: data[i][0] ? data[i][0].toString().trim() : "",
          paspor: data[i][1] ? data[i][1].toString().trim() : "",
          grup: data[i][2] ? data[i][2].toString().trim() : "",
          bus: data[i][3] ? data[i][3].toString().trim() : ""
        });
      }
    }
  }
  
  var sLog = ss.getSheetByName(TAB_LOG);
  var logs = [];
  if (sLog) {
    var data = sLog.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] || data[i][3]) {
        logs.push({
          timestamp: data[i][0] ? data[i][0].toString().trim() : "",
          grup: data[i][1] ? data[i][1].toString().trim() : "",
          nama: data[i][2] ? data[i][2].toString().trim() : "",
          paspor: data[i][3] ? data[i][3].toString().trim() : "",
          pos: data[i][4] ? data[i][4].toString().trim() : "",
          koperBesar: data[i][5] ? data[i][5].toString().trim() : "Tidak Ada",
          koperKecil: data[i][6] ? data[i][6].toString().trim() : "Tidak Ada",
          petugas: data[i][7] ? data[i][7].toString().trim() : "",
          status: data[i][8] ? data[i][8].toString().trim() : "Tidak Lengkap"
        });
      }
    }
  }
  
  logs.reverse();
  
  return {
    success: true,
    groups: groups,
    stations: stations,
    manifest: manifest,
    logs: logs
  };
}

function tambahPosChecklist(req) {
  var namaPos = req.namaPos;
  if (!namaPos) return { success: false, error: "Nama Pos tidak boleh kosong." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sPos = ss.getSheetByName(TAB_POS);
  sPos.appendRow([namaPos]);
  return { success: true, message: "Pos berhasil ditambahkan!" };
}

function simpanLogChecklist(req) {
  var rawData = req.data || req;
  var timestamp = rawData.timestamp || new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });
  var grup = rawData.grup || "";
  var nama = rawData.nama || "";
  var paspor = rawData.paspor || "";
  var pos = rawData.pos || "";
  var koperBesar = rawData.koperBesar || "Tidak Ada";
  var koperKecil = rawData.koperKecil || "Tidak Ada";
  var petugas = rawData.petugas || "GAS User";
  var status = rawData.status || "Tidak Lengkap";
  
  if (!nama && !paspor) {
    return { success: false, error: "Nama atau nomor paspor jemaah harus diisi." };
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sLog = ss.getSheetByName(TAB_LOG);
  
  sLog.appendRow([
    timestamp,
    grup,
    nama,
    paspor,
    pos,
    koperBesar,
    koperKecil,
    petugas,
    status
  ]);
  
  return { success: true, message: "Log checklist koper berhasil disimpan!" };
}

function tambahGrup(req) {
  var grupName = req.grup;
  if (!grupName) return { success: false, error: "Nama Grup tidak boleh kosong." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sGrup = ss.getSheetByName(TAB_GRUP);
  var values = sGrup.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().toLowerCase() === grupName.toLowerCase()) {
      return { success: true, message: "Grup sudah ada." };
    }
  }
  sGrup.appendRow([grupName]);
  return { success: true, message: "Grup berhasil ditambahkan!" };
}

function tambahJemaah(req) {
  var grup = req.grup || "";
  var nama = req.nama || "";
  var paspor = req.paspor || "";
  var bus = req.bus || "";
  if (!nama) return { success: false, error: "Nama jemaah tidak boleh kosong." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sManifest = ss.getSheetByName(TAB_MANIFEST);
  sManifest.appendRow([nama, paspor, grup, bus]);
  return { success: true, message: "Jemaah berhasil disimpan!" };
}

function tambahJemaahBulk(req) {
  var grup = req.grup || "";
  var pilgrims = req.pilgrims || [];
  if (pilgrims.length === 0) return { success: false, error: "Daftar jemaah kosong." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sManifest = ss.getSheetByName(TAB_MANIFEST);
  for (var i = 0; i < pilgrims.length; i++) {
    var p = pilgrims[i];
    sManifest.appendRow([p.nama || "", p.paspor || "", grup, p.bus || ""]);
  }
  return { success: true, message: pilgrims.length + " jemaah berhasil disimpan!" };
}

function editGrup(req) {
  var oldGroupName = req.oldGroupName;
  var newGroupName = req.newGroupName;
  if (!oldGroupName || !newGroupName) return { success: false, error: "Parameter tidak lengkap." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sGrup = ss.getSheetByName(TAB_GRUP);
  if (sGrup) {
    var values = sGrup.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] && values[i][0].toString() === oldGroupName) {
        sGrup.getRange(i + 1, 1).setValue(newGroupName);
        break;
      }
    }
  }
  var sManifest = ss.getSheetByName(TAB_MANIFEST);
  if (sManifest) {
    var values = sManifest.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][2] && values[i][2].toString() === oldGroupName) {
        sManifest.getRange(i + 1, 3).setValue(newGroupName);
      }
    }
  }
  var sLog = ss.getSheetByName(TAB_LOG);
  if (sLog) {
    var values = sLog.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][1] && values[i][1].toString() === oldGroupName) {
        sLog.getRange(i + 1, 2).setValue(newGroupName);
      }
    }
  }
  return { success: true, message: "Grup berhasil diedit!" };
}

function hapusGrup(req) {
  var groupName = req.groupName;
  if (!groupName) return { success: false, error: "Parameter nama grup kosong." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sGrup = ss.getSheetByName(TAB_GRUP);
  if (sGrup) {
    var values = sGrup.getDataRange().getValues();
    for (var i = values.length - 1; i >= 1; i--) {
      if (values[i][0] && values[i][0].toString() === groupName) {
        sGrup.deleteRow(i + 1);
      }
    }
  }
  var sManifest = ss.getSheetByName(TAB_MANIFEST);
  if (sManifest) {
    var values = sManifest.getDataRange().getValues();
    for (var i = values.length - 1; i >= 1; i--) {
      if (values[i][2] && values[i][2].toString() === groupName) {
        sManifest.deleteRow(i + 1);
      }
    }
  }
  var sLog = ss.getSheetByName(TAB_LOG);
  if (sLog) {
    var values = sLog.getDataRange().getValues();
    for (var i = values.length - 1; i >= 1; i--) {
      if (values[i][1] && values[i][1].toString() === groupName) {
        sLog.deleteRow(i + 1);
      }
    }
  }
  return { success: true, message: "Grup berhasil dihapus!" };
}

function editPos(req) {
  var oldStationName = req.oldStationName;
  var newStationName = req.newStationName;
  if (!oldStationName || !newStationName) return { success: false, error: "Nama pos lama/baru kosong." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sPos = ss.getSheetByName(TAB_POS);
  if (sPos) {
    var values = sPos.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] && values[i][0].toString() === oldStationName) {
        sPos.getRange(i + 1, 1).setValue(newStationName);
        break;
      }
    }
  }
  var sLog = ss.getSheetByName(TAB_LOG);
  if (sLog) {
    var values = sLog.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][4] && values[i][4].toString() === oldStationName) {
        sLog.getRange(i + 1, 5).setValue(newStationName);
      }
    }
  }
  return { success: true, message: "Pos berhasil diupdate!" };
}

function hapusPos(req) {
  var stationName = req.stationName;
  if (!stationName) return { success: false, error: "Nama pos kosong." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sPos = ss.getSheetByName(TAB_POS);
  if (sPos) {
    var values = sPos.getDataRange().getValues();
    for (var i = values.length - 1; i >= 1; i--) {
      if (values[i][0] && values[i][0].toString() === stationName) {
        sPos.deleteRow(i + 1);
      }
    }
  }
  var sLog = ss.getSheetByName(TAB_LOG);
  if (sLog) {
    var values = sLog.getDataRange().getValues();
    for (var i = values.length - 1; i >= 1; i--) {
      if (values[i][4] && values[i][4].toString() === stationName) {
        sLog.deleteRow(i + 1);
      }
    }
  }
  return { success: true, message: "Pos berhasil dihapus!" };
}

function editLogChecklist(req) {
  var originalLog = req.originalLog;
  var updatedLog = req.updatedLog;
  if (!originalLog || !updatedLog) return { success: false, error: "Data log kosong." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sLog = ss.getSheetByName(TAB_LOG);
  if (sLog) {
    var values = sLog.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (
        values[i][0].toString().trim() === originalLog.timestamp.toString().trim() &&
        values[i][1].toString().trim() === originalLog.grup.toString().trim() &&
        values[i][3].toString().trim() === originalLog.paspor.toString().trim() &&
        values[i][4].toString().trim() === originalLog.pos.toString().trim()
      ) {
        sLog.getRange(i + 1, 1).setValue(updatedLog.timestamp || originalLog.timestamp);
        sLog.getRange(i + 1, 2).setValue(updatedLog.grup);
        sLog.getRange(i + 1, 3).setValue(updatedLog.nama);
        sLog.getRange(i + 1, 4).setValue(updatedLog.paspor);
        sLog.getRange(i + 1, 5).setValue(updatedLog.pos);
        sLog.getRange(i + 1, 6).setValue(updatedLog.koperBesar);
        sLog.getRange(i + 1, 7).setValue(updatedLog.koperKecil);
        sLog.getRange(i + 1, 8).setValue(updatedLog.petugas);
        sLog.getRange(i + 1, 9).setValue(updatedLog.status);
        return { success: true, message: "Log berhasil diupdate!" };
      }
    }
  }
  return { success: false, error: "Log tidak ditemukan." };
}

function hapusLogChecklist(req) {
  var timestamp = req.timestamp;
  var grup = req.grup;
  var paspor = req.paspor;
  var pos = req.pos;
  if (!timestamp || !grup || !paspor || !pos) return { success: false, error: "Parameter tidak lengkap." };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sLog = ss.getSheetByName(TAB_LOG);
  if (sLog) {
    var values = sLog.getDataRange().getValues();
    for (var i = values.length - 1; i >= 1; i--) {
      if (
        values[i][0].toString().trim() === timestamp.toString().trim() &&
        values[i][1].toString().trim() === grup.toString().trim() &&
        values[i][3].toString().trim() === paspor.toString().trim() &&
        values[i][4].toString().trim() === pos.toString().trim()
      ) {
        sLog.deleteRow(i + 1);
        return { success: true, message: "Log berhasil dihapus!" };
      }
    }
  }
  return { success: false, error: "Log tidak ditemukan." };
}
`;

const mockManifest: Pilgrim[] = [
  { grup: "Grup Sabilal - Riyadh", nama: "Ahmad Fauzi", paspor: "A1234567", bus: "Bus 1" },
  { grup: "Grup Sabilal - Riyadh", nama: "Siti Aminah", paspor: "A7654321", bus: "Bus 1" },
  { grup: "Grup Sabilal - Riyadh", nama: "Budi Santoso", paspor: "A9876543", bus: "Bus 2" },
  { grup: "Grup Al-Haram - Mecca", nama: "Rudi Hermawan", paspor: "B2468135", bus: "Bus 1" },
  { grup: "Grup Al-Haram - Mecca", nama: "Dewi Lestari", paspor: "B1357924", bus: "Bus 2" },
  { grup: "Grup Nabawi - Medina", nama: "Anisa Rahmawati", paspor: "C9876123", bus: "Bus 1" },
  { grup: "Grup Nabawi - Medina", nama: "Muhammad Yusuf", paspor: "C1234987", bus: "Bus 1" }
];

export default function App() {
  // Auth state
  const [user, setUser] = useState<FirebaseUser | null>(null);
  const [token, setToken] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  // Connection settings
  const [gasUrl, setGasUrl] = useState<string>(() => localStorage.getItem("koperjemaah_gas_url") || "");
  const [connectionMode, setConnectionMode] = useState<"demo" | "oauth" | "gas">(
    () => (localStorage.getItem("koperjemaah_connection_mode") as "demo" | "oauth" | "gas") || (localStorage.getItem("koperjemaah_gas_url") ? "gas" : "demo")
  );
  const [spreadsheetId, setSpreadsheetId] = useState<string>(() => localStorage.getItem("koperjemaah_spreadsheet_id") || "");
  const [isInitializingSheet, setIsInitializingSheet] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [lastUpdated, setLastUpdated] = useState<string>(() => localStorage.getItem("koperjemaah_last_updated") || "");
  const [showGasInstructions, setShowGasInstructions] = useState(false);

  // App database states
  const [groups, setGroups] = useState<string[]>([]);
  const [stations, setStations] = useState<string[]>([]);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [manifest, setManifest] = useState<Pilgrim[]>([]);
  
  // Active operational selections
  const [activeGroup, setActiveGroup] = useState<string>("");
  const [activeStation, setActiveStation] = useState<string>("");
  const [newStationName, setNewStationName] = useState<string>("");
  const [isAddingStation, setIsAddingStation] = useState(false);

  // Scanner & OCR states
  const [selectedImage, setSelectedImage] = useState<string | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [isCameraActive, setIsCameraActive] = useState(false);
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const [ocrResult, setOcrResult] = useState<{
    name: string;
    passport: string;
    detectedGroup: string;
  } | null>(null);
  
  // Matching Pilgrim state
  const [matchedPilgrim, setMatchedPilgrim] = useState<Pilgrim | null>(null);
  const [foundInManifest, setFoundInManifest] = useState<boolean | null>(null);
  const [matchScore, setMatchScore] = useState<number>(0);

  // Checkbox & dynamic count luggage selections
  const [koperBesar, setKoperBesar] = useState(false);
  const [koperBesarCount, setKoperBesarCount] = useState<number>(1);
  const [koperKecil, setKoperKecil] = useState(false);
  const [koperKecilCount, setKoperKecilCount] = useState<number>(1);
  const [additionalItems, setAdditionalItems] = useState<AdditionalItem[]>([]);
  const [isSavingLog, setIsSavingLog] = useState(false);

  // Manual search states
  const [searchQuery, setSearchQuery] = useState(""); // Search for overall logs
  const [pilgrimManualSearch, setPilgrimManualSearch] = useState(""); // Search for manifest pilgrim in scanner
  const [showExcelPasteGroup, setShowExcelPasteGroup] = useState<string | null>(null);

  // Custom Item Forms
  const [customItemName, setCustomItemName] = useState("");
  const [customItemQty, setCustomItemQty] = useState<number>(1);

  // Add Group & Pilgrim states
  const [newGroupName, setNewGroupName] = useState("");
  const [isCreatingGroup, setIsCreatingGroup] = useState(false);
  const [expandedAddPilgrimGrup, setExpandedAddPilgrimGrup] = useState<string | null>(null);
  const [newPilgrimName, setNewPilgrimName] = useState("");
  const [newPilgrimPassport, setNewPilgrimPassport] = useState("");
  const [newPilgrimBus, setNewPilgrimBus] = useState("");
  const [selectedBusFilter, setSelectedBusFilter] = useState<string>("");
  const [isAddingPilgrim, setIsAddingPilgrim] = useState(false);
  const [pilgrimQuery, setPilgrimQuery] = useState("");
  const [manageTab, setManageTab] = useState<"manual" | "excel">("manual");

  // Notification and toast state
  const [mismatchNotification, setMismatchNotification] = useState<{
    type: "warning" | "success" | "info";
    title: string;
    message: string;
  } | null>(null);
  const [successToast, setSuccessToast] = useState<string | null>(null);

  // PWA states and install prompt hook
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [isStandalone, setIsStandalone] = useState(false);
  const [showPwaGuideTab, setShowPwaGuideTab] = useState<"chrome-mobile" | "chrome-desktop" | "safari-ios">("chrome-mobile");
  const [showInstallBanner, setShowInstallBanner] = useState<boolean>(
    () => localStorage.getItem("koperjemaah_dismiss_install") !== "true"
  );

  useEffect(() => {
    const isStandaloneMode = window.matchMedia("(display-mode: standalone)").matches || (window.navigator as any).standalone === true;
    setIsStandalone(isStandaloneMode);

    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };

    window.addEventListener("beforeinstallprompt", handleBeforeInstallPrompt);

    return () => {
      window.removeEventListener("beforeinstallprompt", handleBeforeInstallPrompt);
    };
  }, []);

  const handleInstallPWA = async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    if (outcome === "accepted") {
      setDeferredPrompt(null);
    }
  };

  // Edit & Delete states and handlers
  const [editingGroup, setEditingGroup] = useState<string | null>(null);
  const [newGroupNameInput, setNewGroupNameInput] = useState("");
  const [isUpdatingGroup, setIsUpdatingGroup] = useState(false);

  const [editingStation, setEditingStation] = useState<string | null>(null);
  const [newStationNameInput, setNewStationNameInput] = useState("");
  const [isUpdatingStation, setIsUpdatingStation] = useState(false);

  const [editingLog, setEditingLog] = useState<LogEntry | null>(null);
  const [editedKoperBesar, setEditedKoperBesar] = useState<string>("Ada");
  const [editedKoperKecil, setEditedKoperKecil] = useState<string>("Ada");
  const [editedStatus, setEditedStatus] = useState<string>("Lengkap");
  const [editedPetugas, setEditedPetugas] = useState<string>("");
  const [isUpdatingLog, setIsUpdatingLog] = useState(false);

  const handleEditGroup = async (oldName: string, newName: string) => {
    if (!newName.trim()) return;
    setIsUpdatingGroup(true);
    try {
      const res = await fetch("/api/edit-group", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...(token ? { Authorization: `Bearer ${token}` } : {})
        },
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          oldGroupName: oldName,
          newGroupName: newName
        })
      });
      const data = await res.json();
      if (data.success) {
        showToast("Rombongan berhasil diubah!");
        fetchDatabaseData();
        setEditingGroup(null);
      } else {
        showToast("Gagal mengubah rombongan: " + data.error);
      }
    } catch (err: any) {
      showToast("Kesalahan jaringan: " + err.message);
    } finally {
      setIsUpdatingGroup(false);
    }
  };

  const handleDeleteGroup = async (groupName: string) => {
    if (!window.confirm(`Apakah Anda yakin ingin menghapus Rombongan "${groupName}" beserta semua data manifes & lognya? Tindakan ini tidak dapat dibatalkan.`)) return;
    try {
      const res = await fetch("/api/delete-group", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...(token ? { Authorization: `Bearer ${token}` } : {})
        },
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          groupName
        })
      });
      const data = await res.json();
      if (data.success) {
        showToast("Rombongan berhasil dihapus!");
        if (activeGroup === groupName) {
          setActiveGroup("");
          setActiveStation("");
        }
        fetchDatabaseData();
      } else {
        showToast("Gagal menghapus rombongan: " + data.error);
      }
    } catch (err: any) {
      showToast("Kesalahan jaringan: " + err.message);
    }
  };

  const handleEditStation = async (oldName: string, newName: string) => {
    if (!newName.trim()) return;
    setIsUpdatingStation(true);
    try {
      const res = await fetch("/api/edit-station", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...(token ? { Authorization: `Bearer ${token}` } : {})
        },
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          oldStationName: oldName,
          newStationName: newName
        })
      });
      const data = await res.json();
      if (data.success) {
        showToast("Tempat Checklist berhasil diubah!");
        fetchDatabaseData();
        setEditingStation(null);
      } else {
        showToast("Gagal mengubah tempat Checklist: " + data.error);
      }
    } catch (err: any) {
      showToast("Kesalahan jaringan: " + err.message);
    } finally {
      setIsUpdatingStation(false);
    }
  };

  const handleDeleteStation = async (stationName: string) => {
    if (!window.confirm(`Apakah Anda yakin ingin menghapus Tempat Checklist "${stationName}"? Log Checklist yang bersangkutan juga akan disesuaikan.`)) return;
    try {
      const res = await fetch("/api/delete-station", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...(token ? { Authorization: `Bearer ${token}` } : {})
        },
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          stationName
        })
      });
      const data = await res.json();
      if (data.success) {
        showToast("Tempat Checklist berhasil dihapus!");
        if (activeStation === stationName) {
          setActiveStation("");
        }
        fetchDatabaseData();
      } else {
        showToast("Gagal menghapus tempat Checklist: " + data.error);
      }
    } catch (err: any) {
      showToast("Kesalahan jaringan: " + err.message);
    }
  };

  const handleEditLogSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingLog) return;
    setIsUpdatingLog(true);
    try {
      const res = await fetch("/api/edit-log", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...(token ? { Authorization: `Bearer ${token}` } : {})
        },
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          originalLog: editingLog,
          updatedLog: {
            koperBesar: editedKoperBesar,
            koperKecil: editedKoperKecil,
            status: editedStatus,
            petugas: editedPetugas
          }
        })
      });
      const data = await res.json();
      if (data.success) {
        showToast("Log Checklist berhasil diubah!");
        fetchDatabaseData();
        setEditingLog(null);
      } else {
        showToast("Gagal mengubah log: " + data.error);
      }
    } catch (err: any) {
      showToast("Kesalahan jaringan: " + err.message);
    } finally {
      setIsUpdatingLog(false);
    }
  };

  const handleDeleteLog = async (log: LogEntry) => {
    if (!window.confirm(`Apakah Anda yakin ingin menghapus data Checklist koper jemaah "${log.nama}"?`)) return;
    try {
      const res = await fetch("/api/delete-log", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...(token ? { Authorization: `Bearer ${token}` } : {})
        },
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          timestamp: log.timestamp,
          grup: log.grup,
          paspor: log.paspor,
          pos: log.pos
        })
      });
      const data = await res.json();
      if (data.success) {
        showToast("Log Checklist berhasil dihapus!");
        fetchDatabaseData();
      } else {
        showToast("Gagal menghapus log: " + data.error);
      }
    } catch (err: any) {
      showToast("Kesalahan jaringan: " + err.message);
    }
  };

  // Reference for file input
  const fileInputRef = useRef<HTMLInputElement>(null);

  // 1. Session Initialization - Call ambilDataGrupDanPos
  useEffect(() => {
    // Listen to Firebase Auth state
    const unsubscribe = initAuth(
      (currentUser, accessToken) => {
        setUser(currentUser);
        setToken(accessToken);
        if (!localStorage.getItem("koperjemaah_gas_url")) {
          setConnectionMode("oauth");
        }
      },
      () => {
        setUser(null);
        setToken(null);
        if (!localStorage.getItem("koperjemaah_gas_url")) {
          setConnectionMode("demo");
        }
      }
    );

    // Initial fetch
    fetchDatabaseData();

    return () => unsubscribe();
  }, [connectionMode, spreadsheetId, gasUrl, token]);

  // Persist gasUrl in localStorage on change
  useEffect(() => {
    if (gasUrl) {
      localStorage.setItem("koperjemaah_gas_url", gasUrl);
    } else {
      localStorage.removeItem("koperjemaah_gas_url");
    }
  }, [gasUrl]);

  // Persist spreadsheetId in localStorage on change
  useEffect(() => {
    if (spreadsheetId) {
      localStorage.setItem("koperjemaah_spreadsheet_id", spreadsheetId);
    } else {
      localStorage.removeItem("koperjemaah_spreadsheet_id");
    }
  }, [spreadsheetId]);

  // Persist connectionMode in localStorage on change
  useEffect(() => {
    localStorage.setItem("koperjemaah_connection_mode", connectionMode);
  }, [connectionMode]);

  const fetchDatabaseData = async () => {
    try {
      const headers: Record<string, string> = { "Content-Type": "application/json" };
      if (token) {
        headers["Authorization"] = `Bearer ${token}`;
      }

      const res = await fetch("/api/groups-and-checkpoint-stations", {
        method: "POST",
        headers,
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId: connectionMode === "oauth" ? spreadsheetId : "",
          gasUrl: connectionMode === "gas" ? gasUrl : ""
        })
      });

      const data = await res.json();
      if (data.success) {
        setGroups(data.groups || []);
        setStations(data.stations || []);
        setLogs(data.logs || []);
        setManifest(data.manifest || []);
        
        const nowStr = new Date().toLocaleString("id-ID", { 
          dateStyle: "short", 
          timeStyle: "medium" 
        });
        setLastUpdated(nowStr);
        localStorage.setItem("koperjemaah_last_updated", nowStr);
      } else {
        console.warn("Failed to fetch starting session data:", data.error);
        if (connectionMode === "oauth" && spreadsheetId) {
          showNotification("warning", "Koneksi Google Sheets Bermasalah", data.error);
        }
      }
    } catch (err: any) {
      console.error("Session init fetch error:", err);
    }
  };

  // Google Sign-in handler
  const handleGoogleSignIn = async () => {
    if (isLoggingIn) return;
    setIsLoggingIn(true);
    try {
      const result = await googleSignIn();
      if (result) {
        setUser(result.user);
        setToken(result.accessToken);
        setConnectionMode("oauth");
        showToast("Berhasil masuk dengan Google Account!");
      }
    } catch (err: any) {
      const errMsg = err?.message || String(err);
      if (errMsg.includes("cancelled-popup-request") || err?.code === "auth/cancelled-popup-request") {
        console.log("User cancelled the sign-in popup.");
        showToast("Masuk dibatalkan.");
      } else {
        console.error("Login failed:", err);
        showNotification("warning", "Otorisasi Gagal", "Gagal menghubungkan ke Google Account.");
      }
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = async () => {
    await logout();
    setUser(null);
    setToken(null);
    setConnectionMode("demo");
    showToast("Berhasil keluar.");
  };

  // 2. Tambah Pos Checklist
  const handleAddStation = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newStationName.trim()) return;

    setIsAddingStation(true);
    try {
      const headers: Record<string, string> = { "Content-Type": "application/json" };
      if (token) {
        headers["Authorization"] = `Bearer ${token}`;
      }

      const res = await fetch("/api/checkpoint-stations", {
        method: "POST",
        headers,
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          namaPos: newStationName
        })
      });

      const data = await res.json();
      if (data.success) {
        showToast(`Tempat Checklist '${newStationName}' berhasil ditambahkan!`);
        setNewStationName("");
        // Reload stations list
        fetchDatabaseData();
      } else {
        showNotification("warning", "Gagal Tambah Tempat Checklist", data.error || "Gagal menyimpan tempat Checklist baru.");
      }
    } catch (err: any) {
      console.error("Add station error:", err);
      showNotification("warning", "Gagal Tambah Tempat Checklist", "Koneksi error saat menambahkan tempat Checklist.");
    } finally {
      setIsAddingStation(false);
    }
  };

  // Create & Initialize a brand-new Spreadsheet in user's Google Drive
  const handleInitializeNewSheet = async () => {
    if (!token) {
      showNotification("info", "Otorisasi Diperlukan", "Silakan Sign in dengan Google terlebih dahulu.");
      return;
    }

    setIsInitializingSheet(true);
    try {
      const res = await fetch("/api/initialize-spreadsheet", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${token}`,
          "Content-Type": "application/json"
        }
      });

      const data = await res.json();
      if (data.success) {
        setSpreadsheetId(data.spreadsheetId);
        showNotification("success", "Template Siap!", data.message);
        showToast("Spreadsheet Baru Berhasil Dibuat!");
        // Refresh with new spreadsheet data
        setTimeout(() => {
          fetchDatabaseData();
        }, 1500);
      } else {
        showNotification("warning", "Gagal Membuat Spreadsheet", data.error);
      }
    } catch (err: any) {
      console.error("Init spreadsheet error:", err);
      showNotification("warning", "Gagal Membuat Spreadsheet", "Koneksi error.");
    } finally {
      setIsInitializingSheet(false);
    }
  };

  // Image Upload handler
  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setSelectedImage(reader.result as string);
        // Clear previous analysis
        setOcrResult(null);
        setMatchedPilgrim(null);
        setFoundInManifest(null);
      };
      reader.readAsDataURL(file);
    }
  };

  const triggerFileInput = () => {
    fileInputRef.current?.click();
  };

  // Camera handler functions
  const startCamera = async () => {
    setSelectedImage(null);
    setOcrResult(null);
    setMatchedPilgrim(null);
    setFoundInManifest(null);
    setIsCameraActive(true);
    
    try {
      const constraints = {
        video: { facingMode: "environment" }
      };
      const mediaStream = await navigator.mediaDevices.getUserMedia(constraints);
      setTimeout(() => {
        if (videoRef.current) {
          videoRef.current.srcObject = mediaStream;
          videoRef.current.play().catch(e => console.error("Video play failed:", e));
        }
      }, 300);
    } catch (err) {
      console.error("Gagal mengakses kamera:", err);
      showToast("Gagal mengakses kamera. Silakan pilih file secara manual.");
      setIsCameraActive(false);
    }
  };

  const stopCamera = () => {
    if (videoRef.current && videoRef.current.srcObject) {
      const stream = videoRef.current.srcObject as MediaStream;
      stream.getTracks().forEach(track => track.stop());
      videoRef.current.srcObject = null;
    }
    setIsCameraActive(false);
  };

  const capturePhoto = () => {
    if (videoRef.current) {
      const canvas = document.createElement("canvas");
      canvas.width = videoRef.current.videoWidth || 640;
      canvas.height = videoRef.current.videoHeight || 480;
      const ctx = canvas.getContext("2d");
      if (ctx) {
        ctx.drawImage(videoRef.current, 0, 0, canvas.width, canvas.height);
        const dataUrl = canvas.toDataURL("image/jpeg");
        setSelectedImage(dataUrl);
        stopCamera();
        showToast("Foto berhasil diambil! Klik tombol analisis di bawah.");
      }
    }
  };

  // 3. OCR Image Visual Analysis and cariJamaahDalamManifest integration
  const handleAnalyzeSticker = async () => {
    if (!selectedImage) return;
    if (!activeGroup) {
      showNotification("warning", "Grup Belum Dipilih", "Silakan pilih Grup Jemaah aktif terlebih dahulu sebelum memindai stiker.");
      return;
    }

    setIsAnalyzing(true);
    setOcrResult(null);
    setMatchedPilgrim(null);
    setFoundInManifest(null);
    setMismatchNotification(null);

    try {
      const headers: Record<string, string> = { "Content-Type": "application/json" };
      if (token) {
        headers["Authorization"] = `Bearer ${token}`;
      }

      const res = await fetch("/api/analyze-sticker", {
        method: "POST",
        headers,
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          image: selectedImage,
          grup: activeGroup
        })
      });

      const data = await res.json();
      if (data.success) {
        const ocr = data.ocrResult;
        const match = data.matchInfo;

        setOcrResult(ocr);
        setMatchedPilgrim(match.pilgrim);
        setFoundInManifest(match.foundInManifest);
        setMatchScore(match.matchScore);

        // Pre-fill checkboxes: set both to false, prompting verification
        setKoperBesar(false);
        setKoperKecil(false);

        // 6. Tampilkan Notifikasi Push / Banner jika ada ketidaksesuaian data
        if (!match.foundInManifest) {
          showNotification(
            "warning",
            "⚠️ KETIDAKSESUAIAN MANIFEST",
            `Nama "${ocr.name}" hasil pemindaian sistem tidak terdaftar di manifes "${activeGroup}". Silakan periksa kembali stiker koper.`
          );
        } else {
          // Check if the sticker contains a group, and if it differs from selected active group
          if (ocr.detectedGroup && !isGroupNameSimilar(ocr.detectedGroup, activeGroup)) {
            showNotification(
              "warning",
              "⚠️ KETIDAKSESUAIAN GRUP",
              `Grup pada stiker terdeteksi "${ocr.detectedGroup}", namun grup operasional aktif saat ini adalah "${activeGroup}".`
            );
          } else {
            showToast("OCR & Manifest Cocok! Silakan lengkapi konfirmasi koper.");
          }
        }
      } else {
        showNotification("warning", "Analisis Gagal", data.error || "Gagal menganalisis stiker koper.");
      }
    } catch (err: any) {
      console.error("Sticker analysis error:", err);
      showNotification("warning", "Analisis Gagal", "Terjadi kegagalan komunikasi dengan server.");
    } finally {
      setIsAnalyzing(false);
    }
  };

  // Helper check for similar group names
  const isGroupNameSimilar = (detected: string, active: string) => {
    const det = detected.toLowerCase();
    const act = active.toLowerCase();
    return det.includes(act) || act.includes(det) || similaritySimple(det, act) > 0.4;
  };

  const similaritySimple = (s1: string, s2: string) => {
    let matches = 0;
    const words1 = s1.split(/\s+/);
    const words2 = s2.split(/\s+/);
    words1.forEach(w => {
      if (words2.includes(w)) matches++;
    });
    return matches / Math.max(words1.length, words2.length);
  };

  // 5. Simpan Log Checklist
  const handleSaveLog = async () => {
    if (!matchedPilgrim) return;
    if (!activeStation) {
      showNotification("warning", "Tempat Checklist Belum Dipilih", "Silakan pilih Tempat Checklist aktif terlebih dahulu.");
      return;
    }

    setIsSavingLog(true);
    try {
      const headers: Record<string, string> = { "Content-Type": "application/json" };
      if (token) {
        headers["Authorization"] = `Bearer ${token}`;
      }

      // Format luggage counts for database representation
      const koperBesarStr = koperBesar ? `${koperBesarCount} Pcs` : "Tidak Ada";
      const koperKecilStr = koperKecil ? `${koperKecilCount} Pcs` : "Tidak Ada";
      
      // Compute status with additional items appended
      let baseStatus = (koperBesar && koperKecil) ? "Lengkap" : "Tidak Lengkap";
      if (additionalItems.length > 0) {
        const itemsList = additionalItems.map(item => `${item.name} (${item.qty} pcs)`).join(", ");
        baseStatus += ` | Tambahan: ${itemsList}`;
      }

      const logPayload = {
        grup: activeGroup,
        nama: matchedPilgrim.nama,
        paspor: matchedPilgrim.paspor,
        pos: activeStation,
        koperBesar,
        koperKecil,
        koperBesarStr,
        koperKecilStr,
        statusStr: baseStatus,
        petugas: user ? user.displayName || user.email : "Petugas Demo"
      };

      const res = await fetch("/api/save-log", {
        method: "POST",
        headers,
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          logData: logPayload
        })
      });

      const data = await res.json();
      if (data.success) {
        showToast(`Log berhasil disimpan! ${matchedPilgrim.nama} - ${baseStatus}`);
        
        // Clear workspace
        setSelectedImage(null);
        setOcrResult(null);
        setMatchedPilgrim(null);
        setFoundInManifest(null);
        setKoperBesar(false);
        setKoperBesarCount(1);
        setKoperKecil(false);
        setKoperKecilCount(1);
        setAdditionalItems([]);
        setPilgrimManualSearch("");

        // Reload data from database
        fetchDatabaseData();
      } else {
        showNotification("warning", "Simpan Log Gagal", data.error || "Gagal menyimpan log.");
      }
    } catch (err: any) {
      console.error("Save log error:", err);
      showNotification("warning", "Simpan Log Gagal", "Koneksi terputus saat menyimpan log.");
    } finally {
      setIsSavingLog(false);
    }
  };

  // Load a demo mock image sticker for instant testing
  const handleLoadDemoImage = (type: "valid" | "mismatch" | "mismatch_group") => {
    // We can use visual mock canvases or actual base64 of generated text.
    // To make it fully compatible and OCR-friendly, let's render a clean, high contrast canvas in base64 on-the-fly!
    const canvas = document.createElement("canvas");
    canvas.width = 600;
    canvas.height = 350;
    const ctx = canvas.getContext("2d");
    if (ctx) {
      // Draw background
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, 600, 350);

      // Draw travel company header
      ctx.fillStyle = "#047857"; // Emerald Green
      ctx.fillRect(0, 0, 600, 70);

      ctx.fillStyle = "#ffffff";
      ctx.font = "bold 24px Outfit, Arial, sans-serif";
      ctx.fillText("AL-MADINAH TOUR & TRAVEL", 30, 45);

      // Sticker border
      ctx.strokeStyle = "#e2e8f0";
      ctx.lineWidth = 10;
      ctx.strokeRect(0, 0, 600, 350);

      // Draw Sticker details
      ctx.fillStyle = "#1e293b";
      ctx.font = "bold 20px Inter, Arial, sans-serif";
      ctx.fillText("LABEL BAGASI JEMAAH", 30, 110);

      ctx.font = "16px Inter, Arial, sans-serif";
      ctx.fillStyle = "#64748b";
      ctx.fillText("NAMA / NAME:", 30, 160);
      ctx.fillText("PASPOR / PASSPORT:", 30, 220);
      ctx.fillText("GRUP / TRAVEL GROUP:", 30, 280);

      ctx.font = "bold 22px Outfit, Arial, sans-serif";
      ctx.fillStyle = "#0f172a";

      if (type === "valid") {
        ctx.fillText("Ahmad Fauzi", 30, 190);
        ctx.font = "bold 20px 'JetBrains Mono', Courier, monospace";
        ctx.fillText("A1234567", 30, 250);
        ctx.font = "bold 18px Outfit, Arial, sans-serif";
        ctx.fillText("Grup Sabilal - Riyadh", 30, 310);
      } else if (type === "mismatch") {
        ctx.fillText("Bambang Hermanto", 30, 190);
        ctx.font = "bold 20px 'JetBrains Mono', Courier, monospace";
        ctx.fillText("X9998887", 30, 250);
        ctx.font = "bold 18px Outfit, Arial, sans-serif";
        ctx.fillText("Grup Sabilal - Riyadh", 30, 310);
      } else if (type === "mismatch_group") {
        // Ahmad Fauzi, but list group is different
        ctx.fillText("Ahmad Fauzi", 30, 190);
        ctx.font = "bold 20px 'JetBrains Mono', Courier, monospace";
        ctx.fillText("A1234567", 30, 250);
        ctx.font = "bold 18px Outfit, Arial, sans-serif";
        ctx.fillText("Grup Al-Haram - Mecca", 30, 310);
      }

      // Draw a subtle stamp/barcode representation
      ctx.fillStyle = "#1e293b";
      for (let i = 0; i < 40; i++) {
        const width = Math.random() * 6 + 1;
        ctx.fillRect(450 + i * 3, 130, width, 100);
      }
      ctx.font = "10px monospace";
      ctx.fillText("*UMRAH-2026*", 465, 245);

      setSelectedImage(canvas.toDataURL("image/jpeg"));
      setOcrResult(null);
      setMatchedPilgrim(null);
      setFoundInManifest(null);
      showToast("Contoh gambar stiker berhasil dimuat!");
    }
  };

  // State toast manager
  const showToast = (message: string) => {
    setSuccessToast(message);
    setTimeout(() => {
      setSuccessToast(null);
    }, 4500);
  };

  const showNotification = (type: "warning" | "success" | "info", title: string, message: string) => {
    setMismatchNotification({ type, title, message });
  };

  const handleCreateGroup = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newGroupName.trim()) return;

    setIsCreatingGroup(true);
    try {
      const headers: Record<string, string> = { "Content-Type": "application/json" };
      if (token) {
        headers["Authorization"] = `Bearer ${token}`;
      }

      const res = await fetch("/api/add-group", {
        method: "POST",
        headers,
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          grup: newGroupName.trim()
        })
      });

      const data = await res.json();
      if (data.success) {
        showToast(`Grup '${newGroupName.trim()}' berhasil dibuat!`);
        setNewGroupName("");
        fetchDatabaseData();
      } else {
        showNotification("warning", "Gagal Buat Grup", data.error || "Gagal membuat grup.");
      }
    } catch (err: any) {
      console.error("Create group error:", err);
      showNotification("warning", "Gagal Buat Grup", "Koneksi terputus saat membuat grup.");
    } finally {
      setIsCreatingGroup(false);
    }
  };

  const handleCreatePilgrim = async (e: React.FormEvent, groupName: string) => {
    e.preventDefault();
    if (!newPilgrimName.trim() || !newPilgrimPassport.trim()) {
      showNotification("warning", "Data Tidak Lengkap", "Harap isi nama dan nomor paspor jemaah.");
      return;
    }

    setIsAddingPilgrim(true);
    try {
      const headers: Record<string, string> = { "Content-Type": "application/json" };
      if (token) {
        headers["Authorization"] = `Bearer ${token}`;
      }

      const res = await fetch("/api/add-pilgrim", {
        method: "POST",
        headers,
        body: JSON.stringify({
          mode: connectionMode,
          spreadsheetId,
          gasUrl,
          grup: groupName,
          nama: newPilgrimName.trim(),
          paspor: newPilgrimPassport.trim(),
          bus: newPilgrimBus.trim()
        })
      });

      const data = await res.json();
      if (data.success) {
        showToast(`Jemaah '${newPilgrimName.trim()}' berhasil ditambahkan ke ${groupName}!`);
        setNewPilgrimName("");
        setNewPilgrimPassport("");
        setNewPilgrimBus("");
        setExpandedAddPilgrimGrup(null);
        fetchDatabaseData();
      } else {
        showNotification("warning", "Gagal Tambah Jemaah", data.error || "Gagal menambahkan jemaah ke manifest.");
      }
    } catch (err: any) {
      console.error("Add pilgrim error:", err);
      showNotification("warning", "Gagal Tambah Jemaah", "Koneksi terputus.");
    } finally {
      setIsAddingPilgrim(false);
    }
  };

  // Computed Stats for Summary
  const groupStats = () => {
    const groupPilgrims = manifest.filter(p => p.grup === activeGroup);
    const listToUse = groupPilgrims.length > 0 
      ? groupPilgrims 
      : mockManifest.filter(p => p.grup === activeGroup);

    const totalJemaahInGroup = listToUse.length;

    // Checked-in unique keys combining name and passport
    const checkedKeys = new Set(
      logs
        .filter(l => l.grup === activeGroup && l.pos === activeStation)
        .map(l => `${l.nama.trim().toLowerCase()}_${l.paspor.trim().toLowerCase()}`)
    );

    let uniqueCheckedInGroup = 0;

    // Bus level stats
    const busBreakdown: Record<string, { total: number; checked: number; remaining: number }> = {};
    listToUse.forEach(p => {
      const busName = p.bus ? p.bus.trim() : "Tanpa Bus";
      if (!busBreakdown[busName]) {
        busBreakdown[busName] = { total: 0, checked: 0, remaining: 0 };
      }
      busBreakdown[busName].total += 1;
      
      const key = `${p.nama.trim().toLowerCase()}_${p.paspor.trim().toLowerCase()}`;
      if (checkedKeys.has(key)) {
        busBreakdown[busName].checked += 1;
        uniqueCheckedInGroup += 1;
      } else {
        busBreakdown[busName].remaining += 1;
      }
    });

    return {
      total: totalJemaahInGroup,
      checked: uniqueCheckedInGroup,
      percentage: totalJemaahInGroup > 0 ? Math.round((uniqueCheckedInGroup / totalJemaahInGroup) * 100) : 0,
      busBreakdown
    };
  };

  const stats = groupStats();

  return (
    <div id="app-root" className="min-h-screen bg-[#f8fafc] flex flex-col font-sans text-slate-800 antialiased">
      
      {/* Premium Top Navigation Bar */}
      <header className="bg-white border-b border-slate-100/80 px-6 py-4 sticky top-0 z-40 text-left shadow-sm backdrop-blur-md bg-white/95">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row justify-between items-center gap-4">
          
          {/* Brand Logo & Breadcrumb Info */}
          <div className="flex flex-col sm:flex-row items-start sm:items-center gap-4 w-full md:w-auto">
            {/* Brand Logo */}
            <div className="flex items-center gap-3 cursor-pointer" onClick={() => { setActiveGroup(""); setActiveStation(""); }}>
              <div className="bg-gradient-to-br from-emerald-500 to-teal-600 p-2.5 rounded-xl text-white shadow-md shadow-emerald-500/15 flex items-center justify-center">
                <Luggage className="h-5 w-5 text-white" />
              </div>
              <div className="text-left">
                <span className="font-display font-extrabold text-slate-900 text-base tracking-tight block">
                  KoperJemaah <span className="text-emerald-600">PRO</span>
                </span>
                <span className="block text-[9px] text-slate-400 font-bold uppercase tracking-widest mt-0.5">
                  Sistem Manajemen
                </span>
              </div>
            </div>

            {/* Breadcrumbs separator line */}
            <div className="hidden sm:block h-6 w-[1px] bg-slate-200"></div>

            {/* Context breadcrumbs */}
            <div className="flex flex-wrap items-center gap-1.5 text-xs text-slate-500 font-semibold">
              <span className="text-slate-400 font-medium">Rombongan:</span>
              <button 
                onClick={() => { setActiveGroup(""); setActiveStation(""); }}
                className={`font-bold transition-colors px-2 py-1 rounded-lg ${
                  activeGroup 
                    ? "text-emerald-700 bg-emerald-50 hover:bg-emerald-100 cursor-pointer" 
                    : "text-slate-400 bg-slate-50"
                }`}
              >
                {activeGroup || "Semua Rombongan"}
              </button>
              {activeStation && (
                <>
                  <span className="text-slate-300 font-bold">/</span>
                  <span className="text-slate-400 font-medium">Checklist:</span>
                  <button 
                    onClick={() => setActiveStation("")}
                    className="font-bold text-teal-700 bg-teal-50 hover:bg-teal-100 transition-colors px-2 py-1 rounded-lg cursor-pointer"
                  >
                    {activeStation}
                  </button>
                </>
              )}
            </div>
          </div>

          {/* Right Action & Metadata Area */}
          <div className="flex flex-wrap items-center justify-end gap-3 w-full md:w-auto">
            {/* PWA Install Button in Header */}
            {deferredPrompt && (
              <button
                onClick={handleInstallPWA}
                className="bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white px-3 py-1.5 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md shadow-emerald-600/10 cursor-pointer animate-pulse"
              >
                <Smartphone className="h-4 w-4" />
                <span>Instal Aplikasi HP</span>
              </button>
            )}
            
            {isStandalone && (
              <div className="bg-emerald-50 border border-emerald-100 rounded-xl px-2.5 py-1.5 flex items-center gap-1.5 text-emerald-800 text-[10px] font-bold">
                <Smartphone className="h-3.5 w-3.5 text-emerald-600" />
                <span>PWA AKTIF</span>
              </div>
            )}

            {gasUrl && (
              <div className="bg-emerald-50 border border-emerald-150 rounded-xl px-3 py-1.5 flex items-center gap-2 shadow-sm font-mono text-[10px] font-bold">
                <span className="relative flex h-2 w-2">
                  <span className="absolute inline-flex h-full w-full rounded-full opacity-75 bg-emerald-400 animate-ping"></span>
                  <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
                </span>
                <span className="text-emerald-700 uppercase tracking-wider">
                  SINKRONISASI LIVE
                </span>
                {lastUpdated && (
                  <span className="text-[9px] text-emerald-600 border-l border-emerald-200/60 pl-2">
                    {lastUpdated.split(", ")[1] || lastUpdated}
                  </span>
                )}
              </div>
            )}
          </div>

        </div>
      </header>

      {/* Main Panel Wrapper */}
      <div className="flex-grow flex flex-col min-w-0 relative">

        {/* Main Content Body */}
        <main className="flex-grow p-8 w-full max-w-7xl mx-auto flex flex-col gap-8">
        
        
        
        {/* Connection Setup Dashboard Bar (Config Panel) */}
        <div id="connection-panel" className="bg-white p-6 rounded-2xl border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.02)] transition-all duration-300 hover:shadow-[0_8px_30px_rgb(0,0,0,0.04)] text-left">
          <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6 pb-5 border-b border-slate-100">
            <div className="text-left flex-grow">
              <h2 className="text-base font-bold font-display text-slate-900 flex items-center gap-2">
                <Database className="h-5 w-5 text-indigo-600" /> Sinkronisasi Database Google Sheets
              </h2>
              <p className="text-xs text-slate-500 mt-1">
                Koneksikan asisten koper Anda dengan Google Sheets untuk sinkronisasi manifestasi jemaah, tempat checklist, dan histori log secara langsung.
              </p>
            </div>

            {/* Connection mode indicator status */}
            <div className="flex items-center gap-2 bg-slate-50 px-4 py-2 rounded-xl border border-slate-150">
              <span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">MODE KONEKSI:</span>
              <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[10px] font-bold ${
                connectionMode === "oauth"
                  ? "bg-emerald-50 text-emerald-700 border border-emerald-150"
                  : connectionMode === "gas"
                  ? "bg-indigo-50 text-indigo-700 border border-indigo-150"
                  : "bg-amber-50 text-amber-700 border border-amber-150"
              }`}>
                <span className={`h-1.5 w-1.5 rounded-full ${
                  connectionMode === "oauth" ? "bg-emerald-500 animate-pulse" : connectionMode === "gas" ? "bg-indigo-500" : "bg-amber-500"
                }`}></span>
                {connectionMode === "oauth" 
                  ? "Google Sheets Direct (OAuth)" 
                  : connectionMode === "gas" 
                  ? "Google Apps Script (GAS)" 
                  : "Mode Demo (Offline)"}
              </span>
            </div>
          </div>

          {/* Connection Mode Selection Tabs */}
          <div className="flex border-b border-slate-100 mt-5">
            <button
              onClick={() => {
                setConnectionMode("oauth");
              }}
              className={`pb-3 text-xs font-bold border-b-2 px-4 transition-all flex items-center gap-1.5 cursor-pointer ${
                connectionMode === "oauth" || connectionMode === "demo"
                  ? "border-indigo-600 text-indigo-600 font-extrabold"
                  : "border-transparent text-slate-400 hover:text-slate-600"
              }`}
            >
              <Chrome className="h-4 w-4" /> Gmail / Google Sign-In (Sangat Mudah)
            </button>
            <button
              onClick={() => {
                setConnectionMode("gas");
              }}
              className={`pb-3 text-xs font-bold border-b-2 px-4 transition-all flex items-center gap-1.5 cursor-pointer ${
                connectionMode === "gas"
                  ? "border-indigo-600 text-indigo-600 font-extrabold"
                  : "border-transparent text-slate-400 hover:text-slate-600"
              }`}
            >
              <FileCode className="h-4 w-4" /> Web App GAS Script (Kustom)
            </button>
          </div>

          <div className="mt-5">
            {/* TAB 1: OAuth / Google Sign-In */}
            {(connectionMode === "oauth" || connectionMode === "demo") && (
              <div className="space-y-5">
                {!user ? (
                  <div className="bg-slate-50 border border-slate-200/60 p-6 rounded-2xl text-center flex flex-col items-center justify-center space-y-4 max-w-xl mx-auto">
                    <div className="bg-white p-3 rounded-full shadow-sm border border-slate-100">
                      <Chrome className="h-8 w-8 text-indigo-600 animate-bounce" />
                    </div>
                    <div className="space-y-1">
                      <h4 className="font-display font-extrabold text-sm text-slate-900">Hubungkan Akun Google (Gmail) Anda</h4>
                      <p className="text-xs text-slate-500 max-w-sm leading-relaxed">
                        Dengan masuk menggunakan Gmail, aplikasi asisten ini dapat membuat atau memodifikasi file Google Sheets secara aman langsung di Google Drive Anda.
                      </p>
                    </div>
                    <button
                      id="btn-google-signin"
                      onClick={handleGoogleSignIn}
                      disabled={isLoggingIn}
                      className="bg-white hover:bg-slate-50 text-slate-700 font-bold px-6 py-3 rounded-xl text-xs flex items-center justify-center gap-3 border border-slate-200 shadow-sm transition-all duration-200 hover:-translate-y-0.5 active:translate-y-0 cursor-pointer"
                    >
                      {isLoggingIn ? (
                        <Loader2 className="h-4 w-4 animate-spin text-indigo-600" />
                      ) : (
                        <svg className="h-4 w-4" viewBox="0 0 24 24">
                          <path
                            fill="#EA4335"
                            d="M12 5.04c1.66 0 3.2.57 4.38 1.69l3.27-3.27C17.67 1.48 14.98 1 12 1 7.35 1 3.37 3.65 1.39 7.56l3.85 2.99c.9-2.7 3.4-4.51 6.76-4.51z"
                          />
                          <path
                            fill="#4285F4"
                            d="M23.49 12.27c0-.81-.07-1.59-.2-2.36H12v4.51h6.43c-.28 1.44-1.1 2.67-2.33 3.49l3.61 2.8c2.11-1.95 3.33-4.81 3.33-8.44z"
                          />
                          <path
                            fill="#FBBC05"
                            d="M5.24 14.55c-.23-.69-.36-1.43-.36-2.2s.13-1.51.36-2.2L1.39 7.16C.5 8.93 0 10.91 0 13s.5 4.07 1.39 5.84l3.85-3.29z"
                          />
                          <path
                            fill="#34A853"
                            d="M12 23c3.24 0 5.97-1.07 7.96-2.91l-3.61-2.8c-1.1.74-2.5 1.18-4.35 1.18-3.36 0-5.86-1.81-6.76-4.51l-3.85 2.99C3.37 20.35 7.35 23 12 23z"
                          />
                        </svg>
                      )}
                      Masuk dengan Akun Google
                    </button>
                  </div>
                ) : (
                  <div className="space-y-5">
                    {/* User Profile Card */}
                    <div className="flex flex-wrap items-center justify-between gap-4 bg-slate-50 p-4 rounded-xl border border-slate-150">
                      <div className="flex items-center gap-3">
                        {user.photoURL ? (
                          <img src={user.photoURL} alt="Google Avatar" className="h-10 w-10 rounded-full border border-slate-200" referrerPolicy="no-referrer" />
                        ) : (
                          <div className="h-10 w-10 rounded-full bg-indigo-100 text-indigo-700 flex items-center justify-center font-bold">
                            {user.displayName ? user.displayName.charAt(0) : "U"}
                          </div>
                        )}
                        <div>
                          <p className="text-xs font-bold text-slate-800">{user.displayName || "Google User"}</p>
                          <p className="text-[10px] text-slate-500">{user.email}</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-2">
                        <span className="text-[9px] bg-emerald-100 text-emerald-800 font-extrabold px-2.5 py-1 rounded-md uppercase tracking-wider">
                          Otorisasi Aktif
                        </span>
                        <button
                          onClick={handleLogout}
                          className="text-[10px] text-red-600 hover:text-red-700 font-bold px-3 py-1.5 rounded-lg border border-red-200 bg-white hover:bg-red-50 transition-colors cursor-pointer"
                        >
                          Keluar Akun
                        </button>
                      </div>
                    </div>

                    {/* Spreadsheet ID Input & Automatic Setup */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
                      <div className="space-y-3">
                        <label className="block text-xs font-bold text-slate-700 flex items-center gap-1.5">
                          <FileSpreadsheet className="h-4 w-4 text-emerald-600" /> ID Google Spreadsheet Terhubung:
                        </label>
                        <div className="relative">
                          <input
                            id="input-spreadsheet-id"
                            type="text"
                            className="w-full text-xs font-mono border border-slate-200 rounded-xl pl-4 pr-16 py-3 bg-slate-50 focus:bg-white focus:border-indigo-500 focus:ring-4 focus:ring-indigo-500/10 focus:outline-none transition-all duration-200 shadow-inner"
                            placeholder="Contoh: 1aBcDeFgHiJkLmNoPqRsTuVwXyZ"
                            value={spreadsheetId}
                            onChange={(e) => {
                              const val = e.target.value.trim();
                              setSpreadsheetId(val);
                              if (val) {
                                setConnectionMode("oauth");
                              } else {
                                setConnectionMode("demo");
                              }
                            }}
                          />
                          {spreadsheetId && (
                            <button
                              onClick={() => {
                                setSpreadsheetId("");
                                setConnectionMode("demo");
                                showToast("Spreadsheet terputus.");
                              }}
                              className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs font-bold"
                            >
                              Reset
                            </button>
                          )}
                        </div>
                        <p className="text-[10px] text-slate-500 leading-normal">
                          Dapatkan ID dari URL browser spreadsheet Anda:<br />
                          <code className="bg-slate-100 px-1 py-0.5 rounded text-[9px] text-slate-600 font-mono">
                            docs.google.com/spreadsheets/d/<span className="text-indigo-600 font-bold">SPREADSHEET_ID</span>/edit
                          </code>
                        </p>
                      </div>

                      {/* Spreadsheet Quick Actions */}
                      <div className="bg-slate-50/50 border border-slate-150 p-4.5 rounded-xl flex flex-col justify-between space-y-3">
                        <div>
                          <h4 className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                            <Sparkles className="h-4 w-4 text-amber-500" /> Belum punya Spreadsheet?
                          </h4>
                          <p className="text-[10px] text-slate-500 mt-1 leading-normal">
                            Aplikasi asisten dapat membuatkan sebuah spreadsheet baru yang sudah terformat template Manifest, Pos, dan Log langsung di Google Drive Anda.
                          </p>
                        </div>
                        <div className="flex flex-wrap gap-2 pt-1">
                          <button
                            onClick={handleInitializeNewSheet}
                            disabled={isInitializingSheet}
                            className={`px-4 py-2.5 rounded-lg text-[11px] font-bold flex items-center gap-1.5 shadow-sm transition-all cursor-pointer active:scale-95 ${
                              isInitializingSheet
                                ? "bg-slate-200 text-slate-400 cursor-not-allowed"
                                : "bg-indigo-600 hover:bg-indigo-700 text-white"
                            }`}
                          >
                            {isInitializingSheet ? (
                              <Loader2 className="h-3.5 w-3.5 animate-spin" />
                            ) : (
                              <Plus className="h-3.5 w-3.5" />
                            )}
                            {isInitializingSheet ? "Membuat..." : "Buat Spreadsheet Baru"}
                          </button>

                          {spreadsheetId && (
                            <a
                              href={`https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="px-4 py-2.5 rounded-lg text-[11px] font-bold flex items-center gap-1.5 border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 transition-colors shadow-sm"
                            >
                              Buka Spreadsheet ↗
                            </a>
                          )}
                        </div>
                      </div>
                    </div>

                    {/* Sync Trigger button */}
                    <div className="flex justify-end pt-2 border-t border-slate-100">
                      <button
                        id="btn-sync-oauth"
                        disabled={isSyncing}
                        onClick={async () => {
                          if (!spreadsheetId) {
                            showNotification("warning", "ID Spreadsheet Kosong", "Silakan buat spreadsheet baru atau tempel ID Spreadsheet Google Sheets Anda.");
                            return;
                          }
                          setIsSyncing(true);
                          showToast("Sinkronisasi database jemaah...");
                          try {
                            await fetchDatabaseData();
                            showNotification("success", "Sinkronisasi Berhasil", "Berhasil menyinkronkan data langsung dari Google Sheets via API Google.");
                          } catch (err: any) {
                            showNotification("warning", "Sinkronisasi Gagal", "Gagal membaca data dari Google Sheets. Pastikan ID valid dan berikan hak akses: " + err.message);
                          } finally {
                            setIsSyncing(false);
                          }
                        }}
                        className={`px-5 py-2.5 rounded-xl text-xs font-bold transition-all duration-200 flex items-center gap-2 cursor-pointer shadow-md active:scale-95 text-white ${
                          isSyncing
                            ? "bg-indigo-400 cursor-not-allowed"
                            : "bg-indigo-600 hover:bg-indigo-700 shadow-indigo-600/10"
                        }`}
                      >
                        <RefreshCw className={`h-4 w-4 ${isSyncing ? "animate-spin" : ""}`} />
                        {isSyncing ? "Menyinkronkan..." : "Sinkronisasi Database Sekarang"}
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* TAB 2: GAS Web App URL Setup */}
            {connectionMode === "gas" && (
              <div className="space-y-4">
                <div className="flex flex-col md:flex-row gap-3 items-stretch md:items-end">
                  <div className="flex-grow text-left">
                    <label className="block text-xs font-bold text-slate-700 mb-2 flex items-center gap-1.5">
                      <Sparkles className="h-3.5 w-3.5 text-indigo-500" /> URL Web App Google Apps Script:
                    </label>
                    <div className="relative">
                      <input
                        id="input-gas-url"
                        type="text"
                        className="w-full text-xs font-mono border border-slate-200 rounded-xl pl-4 pr-10 py-3 bg-slate-50 focus:bg-white focus:border-indigo-500 focus:ring-4 focus:ring-indigo-500/10 focus:outline-none transition-all duration-200 shadow-inner"
                        placeholder="Masukkan URL Web App GAS Anda (e.g., https://script.google.com/macros/s/.../exec)"
                        value={gasUrl}
                        onChange={(e) => {
                          const val = e.target.value.trim();
                          setGasUrl(val);
                          if (val) {
                            setConnectionMode("gas");
                          } else {
                            setConnectionMode("demo");
                          }
                        }}
                      />
                      {gasUrl && (
                        <button 
                          onClick={() => {
                            setGasUrl("");
                            setConnectionMode("demo");
                            showToast("Kembali ke Mode Demo Offline.");
                          }}
                          className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs"
                        >
                          Reset
                        </button>
                      )}
                    </div>
                  </div>

                  <div className="flex-shrink-0 flex gap-2">
                    <button
                      id="btn-sync-gas"
                      disabled={isSyncing}
                      onClick={async () => {
                        if (!gasUrl) {
                          showNotification("warning", "URL GAS Kosong", "Silakan masukkan URL Web App Google Apps Script terlebih dahulu atau gunakan Mode Demo default.");
                          return;
                        }
                        setIsSyncing(true);
                        showToast("Sedang sinkronisasi data...");
                        try {
                          await fetchDatabaseData();
                          showNotification("success", "Sinkronisasi Berhasil", "Seluruh data rombongan, Cheklist, jemaah, dan log koper berhasil ditarik dari Google Sheets.");
                        } catch (err: any) {
                          showNotification("warning", "Sinkronisasi Gagal", "Pastikan URL GAS valid dan Web App Anda sudah di-deploy dengan benar: " + err.message);
                        } finally {
                          setIsSyncing(false);
                        }
                      }}
                      className={`px-6 py-3 rounded-xl text-xs font-bold transition-all duration-200 flex items-center justify-center gap-2 cursor-pointer shadow-md active:scale-95 text-white ${
                        isSyncing 
                          ? "bg-indigo-400 cursor-not-allowed" 
                          : "bg-indigo-600 hover:bg-indigo-700 shadow-indigo-600/10"
                      }`}
                    >
                      <RefreshCw className={`h-4 w-4 ${isSyncing ? "animate-spin" : ""}`} />
                      {isSyncing ? "Sinkronisasi..." : "Sinkronisasi Sekarang"}
                    </button>
                  </div>
                </div>

                <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 text-xs text-slate-600 leading-relaxed text-left flex items-start gap-2">
                  <Info className="h-4.5 w-4.5 text-indigo-500 flex-shrink-0 mt-0.5" />
                  <div>
                    <p className="font-semibold text-slate-800">Bagaimana Cara Menggunakan Google Apps Script (GAS)?</p>
                    <p className="mt-1">
                      Gunakan tab ini jika Anda ingin sinkronisasi tanpa login Google langsung di aplikasi. Salin kode generator Apps Script yang ada di bilah sisi petunjuk aplikasi, pasang sebagai Web App di Google Sheets Anda, lalu tempel URL-nya di atas.
                    </p>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Real-time Mismatch Push Notifications / Warning Alerts */}
        <AnimatePresence>
          {mismatchNotification && (
            <motion.div
              id="mismatch-banner"
              initial={{ opacity: 0, y: -25 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -25 }}
              className="p-5 rounded-2xl border border-red-150 bg-red-50/50 text-red-900 shadow-lg flex items-start gap-4 relative overflow-hidden text-left backdrop-blur-sm"
            >
              <div className="absolute top-0 right-0 m-4 z-10">
                <span className="px-2.5 py-1 bg-red-600 text-white text-[9px] font-bold rounded uppercase tracking-wider shadow-sm">
                  System Alert
                </span>
              </div>
              <div className="absolute top-0 left-0 h-full w-1.5 bg-red-600"></div>
              <div className="bg-red-100 p-2.5 rounded-xl text-red-700 flex-shrink-0 mt-0.5 shadow-inner">
                <Bell className="h-5 w-5 animate-bounce" />
              </div>
              <div className="flex-grow pr-24">
                <h4 className="text-sm font-bold font-display tracking-tight text-red-800">
                  {mismatchNotification.title}
                </h4>
                <p className="text-xs text-red-700 mt-1.5 leading-relaxed font-semibold">
                  {mismatchNotification.message}
                </p>
              </div>
              <button 
                onClick={() => setMismatchNotification(null)}
                className="text-red-400 hover:text-red-700 text-sm font-bold p-1 self-start z-10 transition-colors cursor-pointer"
              >
                ✕
              </button>
            </motion.div>
          )}
        </AnimatePresence>

        {activeGroup ? (
          <div className="space-y-6">
            
            {/* Multi-step Navigation Banner */}
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-white p-5 rounded-2xl border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.02)]">
              <div className="flex gap-2">
                {activeStation ? (
                  <button
                    id="btn-back-to-stations"
                    onClick={() => {
                      setActiveStation("");
                      setOcrResult(null);
                      setMatchedPilgrim(null);
                      setFoundInManifest(null);
                      setKoperBesar(false);
                      setKoperKecil(false);
                      setKoperBesarCount(1);
                      setKoperKecilCount(1);
                      setAdditionalItems([]);
                      setPilgrimManualSearch("");
                    }}
                    className="bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200/60 px-4 py-2.5 rounded-xl text-xs font-bold smooth-transition flex items-center gap-2 cursor-pointer shadow-sm active:scale-95"
                  >
                    <ArrowLeft className="h-4 w-4" />
                    <span>Kembali ke Pilih Tempat Checklist</span>
                  </button>
                ) : (
                  <button
                    id="btn-back-to-groups"
                    onClick={() => {
                      setActiveGroup("");
                      setActiveStation("");
                      setOcrResult(null);
                      setMatchedPilgrim(null);
                      setFoundInManifest(null);
                    }}
                    className="bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200/60 px-4 py-2.5 rounded-xl text-xs font-bold smooth-transition flex items-center gap-2 cursor-pointer shadow-sm active:scale-95"
                  >
                    <ArrowLeft className="h-4 w-4" />
                    <span>Kembali ke Daftar Rombongan</span>
                  </button>
                )}
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <span className="text-[10px] bg-slate-900 px-3 py-1.5 rounded-xl text-white font-extrabold uppercase tracking-wider shadow-sm">
                  Rombongan: {activeGroup}
                </span>
                {activeStation && (
                  <span className="text-[10px] bg-emerald-600 px-3 py-1.5 rounded-xl text-white font-extrabold uppercase tracking-wider shadow-sm">
                    Tempat Checklist: {activeStation}
                  </span>
                )}
              </div>
            </div>

            {/* PAGE LEVEL BRANCHING */}
            {activeStation === "" ? (
              /* ========================================================
                 PAGE 2: STATION SELECTOR VIEW (PILIH TEMPAT CHECKLIST)
                 ======================================================== */
              <div className="space-y-8 text-left">
                <div className="bg-gradient-to-r from-slate-900 via-slate-800 to-emerald-950 rounded-3xl p-8 text-white border border-slate-800 shadow-xl relative overflow-hidden">
                  <div className="absolute top-0 right-0 -mr-16 -mt-16 w-64 h-64 bg-emerald-500/10 rounded-full blur-3xl pointer-events-none"></div>
                  <div className="relative z-10">
                    <span className="text-[10px] bg-emerald-500/20 text-emerald-300 border border-emerald-500/20 px-3 py-1 rounded-full font-bold uppercase tracking-wider">
                      Tahapan Lapangan
                    </span>
                    <h2 className="text-xl sm:text-2xl font-bold font-display tracking-tight mt-3">
                      Pilih Tempat Checklist Koper
                    </h2>
                    <p className="text-xs text-slate-300 mt-1.5 max-w-2xl leading-relaxed">
                      Silakan buka lokasi Checklist yang sedang berlangsung untuk melacak serta menandai status koper rombongan jemaah secara real-time.
                    </p>
                  </div>
                </div>

                {/* Section: Kelola Manifes Rombongan */}
                <div className="bg-white border border-slate-100 rounded-3xl p-6 shadow-[0_8px_30px_rgb(0,0,0,0.02)] transition-all hover:shadow-[0_8px_30px_rgb(0,0,0,0.04)] text-left mt-4 mb-8">
                  <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-slate-100 pb-5 mb-6">
                    <div>
                      <h2 className="text-sm font-extrabold uppercase tracking-wider text-slate-900 flex items-center gap-2">
                        <Users className="h-5 w-5 text-emerald-600" /> KELOLA MANIFES JAMAAH ROMBONGAN: <span className="text-emerald-700">{activeGroup}</span>
                      </h2>
                      <p className="text-xs text-slate-500 mt-1">
                        Daftar lengkap manifes rombongan ini. Anda dapat menambahkan jamaah secara manual atau impor bulk sekaligus lewat paste dari Excel.
                      </p>
                    </div>
                    
                    {/* Switch Tab buttons */}
                    <div className="flex items-center gap-1 bg-slate-150/60 p-1 rounded-xl shrink-0">
                      <button
                        type="button"
                        onClick={() => setManageTab("manual")}
                        className={`px-4 py-2 rounded-lg text-xs font-bold cursor-pointer transition-all ${
                          manageTab === "manual" 
                            ? "bg-white text-slate-900 shadow-sm" 
                            : "text-slate-500 hover:text-slate-800"
                        }`}
                      >
                        Input Manual
                      </button>
                      <button
                        type="button"
                        onClick={() => setManageTab("excel")}
                        className={`px-4 py-2 rounded-lg text-xs font-bold cursor-pointer transition-all ${
                          manageTab === "excel" 
                            ? "bg-white text-slate-900 shadow-sm" 
                            : "text-slate-500 hover:text-slate-800"
                        }`}
                      >
                        Impor Excel (Bulk)
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
                    
                    {/* Left Panel (6 cols): Manifest Roster/List */}
                    <div className="lg:col-span-6 space-y-4">
                      <div className="flex justify-between items-center gap-4">
                        <h3 className="text-xs font-extrabold text-slate-400 uppercase tracking-widest">
                          Daftar Manifes Aktif ({manifest.filter(p => p.grup === activeGroup).length} Jamaah)
                        </h3>
                        {/* Search input (Filter Cari) */}
                        <div className="relative w-48">
                          <Search className="absolute left-2.5 top-2.5 h-3.5 w-3.5 text-slate-400" />
                          <input
                            type="text"
                            placeholder="Filter Cari Nama / Paspor..."
                            className="w-full pl-8 pr-3 py-1.5 text-xs border border-slate-200 rounded-lg focus:outline-none focus:border-emerald-500 focus:ring-2 focus:ring-emerald-500/10 bg-slate-50/50 text-slate-850 font-medium transition-all"
                            value={pilgrimQuery}
                            onChange={(e) => setPilgrimQuery(e.target.value)}
                          />
                        </div>
                      </div>

                      <div className="border border-slate-100 rounded-2xl overflow-hidden max-h-[350px] overflow-y-auto bg-slate-50/20">
                        <table className="w-full text-left text-xs border-collapse">
                          <thead>
                            <tr className="bg-slate-50 text-slate-400 text-[10px] font-bold uppercase tracking-wider border-b border-slate-100/80">
                              <th className="p-3 pl-4">No</th>
                              <th className="p-3">Nama Jamaah</th>
                              <th className="p-3">No. Paspor</th>
                              <th className="p-3 pr-4">Bus / Grup</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100">
                            {(() => {
                              const filtered = manifest
                                .filter(p => p.grup === activeGroup)
                                .filter(p => {
                                  const q = pilgrimQuery.toLowerCase();
                                  return p.nama.toLowerCase().includes(q) || p.paspor.toLowerCase().includes(q);
                                });

                              if (filtered.length === 0) {
                                return (
                                  <tr>
                                    <td colSpan={4} className="p-8 text-center text-slate-400 font-bold">
                                      {pilgrimQuery ? "Tidak ada jamaah yang cocok." : "Belum ada jamaah terdaftar. Silakan tambahkan jamaah melalui form di sebelah kanan."}
                                    </td>
                                  </tr>
                                );
                              }

                              return filtered.map((p, idx) => (
                                <tr key={idx} className="hover:bg-slate-50/60 transition-colors text-slate-700">
                                  <td className="p-3 pl-4 font-mono text-[10px] text-slate-400 font-bold">{idx + 1}</td>
                                  <td className="p-3 font-bold text-slate-850">{p.nama}</td>
                                  <td className="p-3 font-mono font-bold text-[11px] text-slate-500">{p.paspor}</td>
                                  <td className="p-3 pr-4 text-slate-500 font-semibold">{p.bus || "-"}</td>
                                </tr>
                              ));
                            })()}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* Right Panel (6 cols): Selected Input Form */}
                    <div className="lg:col-span-6 bg-slate-50/40 border border-slate-100 rounded-2xl p-5 text-slate-850">
                      {manageTab === "manual" ? (
                        <form onSubmit={(e) => handleCreatePilgrim(e, activeGroup)} className="space-y-4 text-left">
                          <h4 className="text-xs font-bold text-slate-800 uppercase tracking-widest border-b border-slate-150 pb-2 mb-3">
                            INPUT JAMAAH BARU SECARA MANUAL
                          </h4>
                          <div>
                            <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Nama Lengkap Jamaah</label>
                            <input
                              type="text"
                              required
                              placeholder="Contoh: Muhammad Ali"
                              className="w-full text-xs border border-slate-200 rounded-xl px-4 py-2.5 focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 focus:outline-none bg-white text-slate-850 font-semibold transition-all shadow-sm"
                              value={newPilgrimName}
                              onChange={(e) => setNewPilgrimName(e.target.value)}
                            />
                          </div>
                          <div className="grid grid-cols-2 gap-4">
                            <div>
                              <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Nomor Paspor</label>
                              <input
                                type="text"
                                required
                                placeholder="Contoh: B9876543"
                                className="w-full text-xs border border-slate-200 rounded-xl px-4 py-2.5 focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 focus:outline-none font-mono font-bold bg-white text-slate-850 transition-all shadow-sm"
                                value={newPilgrimPassport}
                                onChange={(e) => setNewPilgrimPassport(e.target.value)}
                              />
                            </div>
                            <div>
                              <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Bus / Rombongan</label>
                              <input
                                type="text"
                                placeholder="Contoh: Bus 1"
                                className="w-full text-xs border border-slate-200 rounded-xl px-4 py-2.5 focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 focus:outline-none bg-white text-slate-850 font-semibold transition-all shadow-sm"
                                value={newPilgrimBus}
                                onChange={(e) => setNewPilgrimBus(e.target.value)}
                              />
                            </div>
                          </div>
                          <button
                            type="submit"
                            disabled={isAddingPilgrim}
                            className="w-full bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold py-3 px-4 rounded-xl smooth-transition flex items-center justify-center gap-1.5 shadow-md shadow-emerald-600/10 disabled:bg-slate-250 cursor-pointer active:scale-95"
                          >
                            {isAddingPilgrim ? (
                              <>
                                <Loader2 className="h-4 w-4 animate-spin" /> Menyimpan...
                              </>
                            ) : (
                              <>
                                <Plus className="h-4 w-4" /> Simpan Jamaah Baru
                              </>
                            )}
                          </button>
                        </form>
                      ) : (
                        <div className="text-left space-y-4">
                          <h4 className="text-xs font-bold text-slate-800 uppercase tracking-widest border-b border-slate-150 pb-2 mb-3">
                            IMPOR JAMAAH BULK (COPY PASTE DARI EXCEL)
                          </h4>
                          <ExcelPasteSection
                            groupName={activeGroup}
                            connectionMode={connectionMode}
                            spreadsheetId={spreadsheetId}
                            gasUrl={gasUrl}
                            token={token}
                            onSuccess={() => {
                              showToast(`Manifes jamaah bulk berhasil diimpor untuk rombongan ${activeGroup}!`);
                              fetchDatabaseData();
                            }}
                            onCancel={() => {}}
                          />
                        </div>
                      )}
                    </div>

                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  {[...stations].reverse().map((s, idx) => {
                    const stationLogs = logs.filter(l => l.grup === activeGroup && l.pos === s);
                    const checkedInStation = new Set(stationLogs.map(l => l.nama)).size;
                    return (
                      <div 
                        key={idx}
                        className="bg-white border border-slate-100 rounded-2xl shadow-[0_8px_30px_rgb(0,0,0,0.02)] overflow-hidden flex flex-col justify-between hover:border-emerald-500 hover:shadow-[0_8px_30px_rgb(16,185,129,0.05)] transition-all duration-300 group"
                      >
                        <div className="p-6 flex-grow text-left">
                          <div className="flex justify-between items-start mb-4">
                            <div className="bg-emerald-50 group-hover:bg-emerald-100 p-3 rounded-xl border border-emerald-100 group-hover:border-emerald-200 transition-all duration-300 self-start inline-block">
                              <MapPin className="h-5 w-5 text-emerald-600" />
                            </div>
                            
                            {editingStation !== s && (
                              <div className="flex gap-1 items-center opacity-40 group-hover:opacity-100 transition-all duration-300">
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setEditingStation(s);
                                    setNewStationNameInput(s);
                                  }}
                                  className="text-slate-400 hover:text-emerald-600 p-1.5 rounded-lg hover:bg-slate-50 cursor-pointer transition-colors"
                                  title="Edit Nama Tempat Checklist"
                                >
                                  <Edit2 className="h-3.5 w-3.5" />
                                </button>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleDeleteStation(s);
                                  }}
                                  className="text-slate-400 hover:text-red-600 p-1.5 rounded-lg hover:bg-slate-50 cursor-pointer transition-colors"
                                  title="Hapus Tempat Checklist"
                                >
                                  <Trash2 className="h-3.5 w-3.5" />
                                </button>
                              </div>
                            )}
                          </div>

                          {editingStation === s ? (
                            <div className="flex items-center gap-1.5 mb-2" onClick={(e) => e.stopPropagation()}>
                              <input
                                type="text"
                                className="text-xs border border-slate-200 rounded-xl px-3 py-2 bg-slate-50 focus:bg-white focus:ring-4 focus:ring-emerald-500/10 focus:border-emerald-500 focus:outline-none w-full text-slate-850 font-semibold"
                                value={newStationNameInput}
                                onChange={(e) => setNewStationNameInput(e.target.value)}
                                autoFocus
                              />
                              <button
                                onClick={() => handleEditStation(s, newStationNameInput)}
                                className="bg-emerald-600 hover:bg-emerald-700 text-white p-2 rounded-xl transition-all duration-200 cursor-pointer shadow-sm active:scale-95 flex items-center justify-center"
                                title="Simpan"
                                disabled={isUpdatingStation}
                              >
                                {isUpdatingStation ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Check className="h-3.5 w-3.5" />}
                              </button>
                              <button
                                onClick={() => setEditingStation(null)}
                                className="bg-slate-100 hover:bg-slate-200 text-slate-600 p-2 rounded-xl transition-colors cursor-pointer"
                                title="Batal"
                              >
                                ✕
                              </button>
                            </div>
                          ) : (
                            <>
                              <h3 className="text-sm font-bold font-display text-slate-800 group-hover:text-emerald-700 transition-colors mb-1 leading-tight">
                                {s}
                              </h3>
                              <p className="text-[11px] text-slate-400 font-medium">Checklist koper aktif di tempat Checklist ini</p>
                            </>
                          )}
                          
                          <div className="mt-4 pt-4 border-t border-slate-100 text-[11px] text-slate-500 space-y-1">
                            <div className="flex justify-between items-center">
                              <span className="font-medium text-slate-400">Terverifikasi di Lokasi Ini:</span>
                              <strong className="text-slate-800 bg-slate-50 px-2.5 py-1 rounded-lg border border-slate-100">{checkedInStation} jamaah</strong>
                            </div>
                          </div>
                        </div>

                        <div className="p-4 bg-slate-50 border-t border-slate-100/80 flex items-center justify-between">
                          <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Aksi Checklist</span>
                          <button
                            onClick={() => {
                              setActiveStation(s);
                              setOcrResult(null);
                              setMatchedPilgrim(null);
                              setFoundInManifest(null);
                            }}
                            className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-4 py-2.5 rounded-xl shadow-sm hover:shadow-md smooth-transition flex items-center gap-1 cursor-pointer active:scale-95"
                          >
                            Buka Checklist <ChevronRight className="h-3.5 w-3.5" />
                          </button>
                        </div>
                      </div>
                    );
                  })}

                  {/* Add New Station Card */}
                  <div className="bg-slate-50/50 border-2 border-dashed border-slate-200 rounded-2xl p-6 flex flex-col justify-between hover:border-emerald-300 hover:bg-white transition-all duration-300 text-left">
                    <div>
                      <div className="bg-slate-100 p-2.5 rounded-xl border border-slate-200 self-start inline-block mb-3">
                        <Plus className="h-4.5 w-4.5 text-slate-500" />
                      </div>
                      <h3 className="text-xs font-bold text-slate-700 mb-1">Tambah Tempat Checklist</h3>
                      <p className="text-[11px] text-slate-400 leading-relaxed mb-4 font-medium">
                        Tambahkan lokasi pemantauan baru seperti bandara, hotel, lobby bus, asrama jamaah, atau checkpoint transit.
                      </p>
                    </div>

                    <form onSubmit={handleAddStation} className="space-y-3">
                      <input
                        type="text"
                        required
                        placeholder="Contoh: Hotel Mekkah Dar Al-Iman"
                        className="w-full text-xs border border-slate-200 rounded-xl px-4 py-2.5 bg-white focus:ring-4 focus:ring-emerald-500/10 focus:border-emerald-500 focus:outline-none text-slate-850 font-medium transition-all duration-200"
                        value={newStationName}
                        onChange={(e) => setNewStationName(e.target.value)}
                      />
                      <button
                        type="submit"
                        disabled={isAddingStation || !newStationName.trim()}
                        className="w-full bg-slate-950 hover:bg-slate-900 text-white text-xs font-bold py-2.5 px-4 rounded-xl smooth-transition flex items-center justify-center gap-2 shadow-md disabled:bg-slate-200 disabled:text-slate-400"
                      >
                        {isAddingStation ? (
                          <>
                            <Loader2 className="h-4 w-4 animate-spin" /> Menyimpan...
                          </>
                        ) : (
                          <>
                            <Plus className="h-4.5 w-4.5" /> Tambah Cheklist Baru
                          </>
                        )}
                      </button>
                    </form>
                  </div>
                </div>
              </div>
            ) : (
              /* ========================================================
                 PAGE 3: DEDICATED EXAMINATION VIEW (TEMPAT CHECKLIST)
                 ======================================================== */
              <div className="space-y-8">
                
                <div className="max-w-2xl mx-auto w-full flex flex-col gap-8">
                  
                  {/* Verification and Checklist Card */}
                  <div id="verification-card" className="bg-white p-6 rounded-2xl border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.02)] min-h-[300px] flex flex-col transition-all duration-300 hover:shadow-[0_8px_30px_rgb(0,0,0,0.04)] text-left">
                    <h2 className="mb-4 text-xs font-bold uppercase tracking-widest text-slate-400 flex items-center gap-2">
                      <CheckCircle2 className="h-4.5 w-4.5 text-emerald-600" /> DETAIL KONFIRMASI JEMAAH
                    </h2>

                    {/* Manual Jamaah Selection (FALLBACK & MANUAL LOOKUP) */}
                    {!matchedPilgrim && (
                      <div className="mb-5 p-4 rounded-xl border border-slate-100 bg-slate-50/50 text-left">
                        <span className="text-xs font-bold text-slate-700 block mb-1">Pilih Jemaah:</span>
                        <p className="text-[10px] text-slate-400 mb-3">Ketik nama jemaah atau nomor paspor untuk memulai pengisian Checklist.</p>
                        <div className="relative mb-2">
                          <Search className="absolute left-3.5 top-3 h-4 w-4 text-slate-400" />
                          <input
                            type="text"
                            placeholder="Ketik Nama Lengkap atau No Paspor..."
                            className="w-full pl-10 pr-4 py-2.5 text-xs border border-slate-200 rounded-xl bg-white focus:outline-none focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 text-slate-800 transition-all duration-200 font-medium"
                            value={pilgrimManualSearch}
                            onChange={(e) => setPilgrimManualSearch(e.target.value)}
                          />
                        </div>
                        
                        {pilgrimManualSearch.trim().length > 0 && (
                          <div className="max-h-36 overflow-y-auto border border-slate-150 rounded-xl bg-white divide-y divide-slate-100 shadow-lg">
                            {manifest
                              .filter(p => p.grup === activeGroup)
                              .filter(p => 
                                p.nama.toLowerCase().includes(pilgrimManualSearch.toLowerCase()) || 
                                p.paspor.toLowerCase().includes(pilgrimManualSearch.toLowerCase())
                              )
                              .map((p, idx) => (
                                <button
                                  key={idx}
                                  type="button"
                                  onClick={() => {
                                    setMatchedPilgrim(p);
                                    setFoundInManifest(true);
                                    setOcrResult({ name: p.nama, passport: p.paspor });
                                    setPilgrimManualSearch("");
                                    setKoperBesar(true);
                                    setKoperBesarCount(1);
                                    setKoperKecil(true);
                                    setKoperKecilCount(1);
                                  }}
                                  className="w-full text-left px-4 py-2.5 text-xs hover:bg-emerald-50/50 flex justify-between items-center text-slate-700 font-bold transition-colors cursor-pointer"
                                >
                                  <span>{p.nama}</span>
                                  <span className="font-mono text-[10px] text-slate-400 font-bold bg-slate-50 px-2 py-0.5 rounded border border-slate-100">{p.paspor}</span>
                                </button>
                              ))
                            }
                            {manifest
                              .filter(p => p.grup === activeGroup)
                              .filter(p => 
                                p.nama.toLowerCase().includes(pilgrimManualSearch.toLowerCase()) || 
                                p.paspor.toLowerCase().includes(pilgrimManualSearch.toLowerCase())
                              ).length === 0 && (
                                <div className="p-3.5 text-center text-xs text-slate-400 font-semibold">Jemaah tidak ditemukan di manifest grup ini.</div>
                              )
                            }
                          </div>
                        )}
                      </div>
                    )}

                    {!matchedPilgrim && (
                      <div className="flex-grow flex flex-col items-center justify-center text-center p-8 bg-slate-50/50 border-2 border-dashed border-slate-200 rounded-2xl">
                        <div className="bg-slate-100 p-3 rounded-full text-slate-400 mb-3 border border-slate-200">
                          <User className="h-6 w-6 text-slate-500" />
                        </div>
                        <p className="text-xs font-bold text-slate-700">Belum Ada Jemaah Dipilih</p>
                        <p className="text-[10px] text-slate-400 mt-1 max-w-[220px] leading-relaxed font-medium">
                          Silakan cari nama jemaah pada kolom pencarian di atas untuk memulai Checklist Koper.
                        </p>
                      </div>
                    )}

                    {matchedPilgrim && (
                      <motion.div
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        className="flex-grow flex flex-col"
                      >
                          {/* Pilgrim Identity Banner */}
                          <div className="mb-5 flex items-center gap-4 border-b pb-5 border-slate-100">
                            <div className="h-14 w-14 rounded-full bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center text-white font-extrabold text-base shrink-0 shadow-md">
                              {matchedPilgrim?.nama ? matchedPilgrim.nama.split(' ').map((n: string) => n[0]).join('').slice(0, 2).toUpperCase() : "JM"}
                            </div>
                            <div className="text-left">
                              <h4 className="text-md font-bold text-slate-900 tracking-tight leading-snug">{matchedPilgrim?.nama || "Tidak Dikenal"}</h4>
                              <p className="text-xs font-mono font-bold text-slate-400 mt-0.5">Paspor: {matchedPilgrim?.paspor || "Tidak Ditemukan"}</p>
                              <div className="flex flex-wrap items-center gap-2 mt-1.5">
                                <span className={`inline-block px-2.5 py-0.5 text-[9px] font-bold rounded-full ${
                                  foundInManifest 
                                    ? "bg-emerald-50 text-emerald-700 border border-emerald-200" 
                                    : "bg-red-50 text-red-700 border border-red-200"
                                }`}>
                                  {foundInManifest ? `MANIFES COCOK (${matchedPilgrim?.grup})` : "DATA TIDAK COCOK"}
                                </span>
                                {matchedPilgrim?.bus && (
                                  <span className="inline-flex items-center gap-1 bg-indigo-50 text-indigo-700 border border-indigo-200 px-2.5 py-0.5 text-[9px] font-bold rounded-full">
                                    <Bus className="h-3 w-3" /> BUS {matchedPilgrim.bus}
                                  </span>
                                )}
                              </div>
                            </div>
                          </div>

                          {/* Interactive suitcase checklist validation with quantity modifier */}
                          <div className="space-y-5 flex-grow text-left">
                            <p className="text-xs font-extrabold text-slate-700">Checklist Kelengkapan Fisik Koper:</p>

                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                              
                              {/* Koper Besar */}
                              <div className={`p-4 rounded-xl border flex flex-col justify-between transition-all duration-200 ${
                                koperBesar 
                                  ? "bg-emerald-50/50 border-emerald-500 text-emerald-950 ring-2 ring-emerald-500/10" 
                                  : "bg-slate-50/50 border-slate-200 text-slate-500"
                              }`}>
                                <div className="flex items-center justify-between">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const nextVal = !koperBesar;
                                      setKoperBesar(nextVal);
                                      if (nextVal && koperBesarCount === 0) setKoperBesarCount(1);
                                    }}
                                    className="flex items-center gap-2.5 cursor-pointer focus:outline-none text-left"
                                  >
                                    <input type="checkbox" className="h-4.5 w-4.5 rounded-lg border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer" checked={koperBesar} onChange={() => {}} />
                                    <span className="text-xs font-extrabold">Koper Besar Bagasi</span>
                                  </button>
                                </div>
                                {koperBesar && (
                                  <div className="mt-3 pt-3 border-t border-emerald-100 flex items-center justify-between">
                                    <span className="text-[10px] text-emerald-700 font-bold">Jumlah:</span>
                                    <div className="flex items-center gap-1.5">
                                      <button
                                        type="button"
                                        onClick={() => {
                                          if (koperBesarCount > 1) {
                                            setKoperBesarCount(koperBesarCount - 1);
                                          } else {
                                            setKoperBesar(false);
                                          }
                                        }}
                                        className="bg-white hover:bg-emerald-50 border border-slate-200 text-slate-700 h-6.5 w-6.5 rounded-lg flex items-center justify-center font-bold text-xs shadow-sm transition-all cursor-pointer"
                                      >
                                        -
                                      </button>
                                      <span className="text-xs font-bold text-emerald-900 font-mono w-6 text-center">{koperBesarCount}</span>
                                      <button
                                        type="button"
                                        onClick={() => setKoperBesarCount(koperBesarCount + 1)}
                                        className="bg-white hover:bg-emerald-50 border border-slate-200 text-slate-700 h-6.5 w-6.5 rounded-lg flex items-center justify-center font-bold text-xs shadow-sm transition-all cursor-pointer"
                                      >
                                        +
                                      </button>
                                    </div>
                                  </div>
                                )}
                              </div>

                              {/* Koper Kecil */}
                              <div className={`p-4 rounded-xl border flex flex-col justify-between transition-all duration-200 ${
                                koperKecil 
                                  ? "bg-emerald-50/50 border-emerald-500 text-emerald-950 ring-2 ring-emerald-500/10" 
                                  : "bg-slate-50/50 border-slate-200 text-slate-500"
                              }`}>
                                <div className="flex items-center justify-between">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const nextVal = !koperKecil;
                                      setKoperKecil(nextVal);
                                      if (nextVal && koperKecilCount === 0) setKoperKecilCount(1);
                                    }}
                                    className="flex items-center gap-2.5 cursor-pointer focus:outline-none text-left"
                                  >
                                    <input type="checkbox" className="h-4.5 w-4.5 rounded-lg border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer" checked={koperKecil} onChange={() => {}} />
                                    <span className="text-xs font-extrabold">Koper Kecil Kabin</span>
                                  </button>
                                </div>
                                {koperKecil && (
                                  <div className="mt-3 pt-3 border-t border-emerald-100 flex items-center justify-between">
                                    <span className="text-[10px] text-emerald-700 font-bold">Jumlah:</span>
                                    <div className="flex items-center gap-1.5">
                                      <button
                                        type="button"
                                        onClick={() => {
                                          if (koperKecilCount > 1) {
                                            setKoperKecilCount(koperKecilCount - 1);
                                          } else {
                                            setKoperKecil(false);
                                          }
                                        }}
                                        className="bg-white hover:bg-emerald-50 border border-slate-200 text-slate-700 h-6.5 w-6.5 rounded-lg flex items-center justify-center font-bold text-xs shadow-sm transition-all cursor-pointer"
                                      >
                                        -
                                      </button>
                                      <span className="text-xs font-bold text-emerald-900 font-mono w-6 text-center">{koperKecilCount}</span>
                                      <button
                                        type="button"
                                        onClick={() => setKoperKecilCount(koperKecilCount + 1)}
                                        className="bg-white hover:bg-emerald-50 border border-slate-200 text-slate-700 h-6.5 w-6.5 rounded-lg flex items-center justify-center font-bold text-xs shadow-sm transition-all cursor-pointer"
                                      >
                                        +
                                      </button>
                                    </div>
                                  </div>
                                )}
                              </div>

                            </div>

                            {/* Warning if koper kecil is not checked */}
                            {!koperKecil && (
                              <div className="rounded-xl bg-amber-50 p-3.5 border border-amber-200/60 flex items-start gap-2.5 text-left">
                                <AlertTriangle className="h-4 w-4 text-amber-600 mt-0.5 shrink-0" />
                                <p className="text-[11px] text-amber-850 leading-relaxed font-semibold">
                                  <b>Checklist Mandiri:</b> Koper kecil kabin belum ditandai. Harap ingatkan jemaah untuk selalu memegang koper kabinnya sendiri selama transit.
                                </p>
                              </div>
                            )}

                            {/* Additional Items Section */}
                            <div className="p-4 rounded-xl border border-slate-100 bg-slate-50 text-left space-y-3">
                              <span className="text-xs font-bold text-slate-700 block">Barang Bawaan Tambahan:</span>
                              
                              {additionalItems.length > 0 ? (
                                <div className="flex flex-wrap gap-1.5 mb-2.5">
                                  {additionalItems.map((item, i) => (
                                    <span key={i} className="inline-flex items-center gap-1.5 text-[10px] bg-emerald-50 text-emerald-800 font-bold px-2.5 py-1 rounded-full border border-emerald-100">
                                      <span>{item.name} ({item.qty})</span>
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setAdditionalItems(additionalItems.filter((_, idx) => idx !== i));
                                        }}
                                        className="text-emerald-500 hover:text-emerald-800 font-extrabold ml-1 cursor-pointer"
                                      >
                                        ×
                                      </button>
                                    </span>
                                  ))}
                                </div>
                              ) : (
                                <p className="text-[10px] text-slate-400 italic mb-2.5 font-medium">Belum ada barang tambahan (misal: kursi roda, air zamzam, kerdus, dll).</p>
                              )}

                              {/* Quick suggestions */}
                              <div className="mb-2.5">
                                <span className="text-[10px] text-slate-400 block mb-1.5 font-bold uppercase tracking-wide">Rekomendasi Cepat:</span>
                                <div className="flex flex-wrap gap-1.5">
                                  {[
                                    { label: "♿ Kursi Roda", name: "Kursi Roda" },
                                    { label: "🎒 Tas Tambahan", name: "Tas Tambahan" },
                                    { label: "💧 Air Zamzam", name: "Air Zamzam" },
                                    { label: "📦 Kardus Air", name: "Kardus Air" },
                                    { label: "💼 Tas Kabin", name: "Tas Kabin" }
                                  ].map((t, idx) => (
                                    <button
                                      key={idx}
                                      type="button"
                                      onClick={() => {
                                        const existing = additionalItems.find(item => item.name === t.name);
                                        if (existing) {
                                          setAdditionalItems(additionalItems.map(item => 
                                            item.name === t.name ? { ...item, qty: item.qty + 1 } : item
                                          ));
                                        } else {
                                          setAdditionalItems([...additionalItems, { name: t.name, qty: 1 }]);
                                        }
                                      }}
                                      className="text-[10px] bg-white hover:bg-slate-100 text-slate-600 border border-slate-200/80 px-2.5 py-1.5 rounded-lg transition-colors cursor-pointer font-medium"
                                    >
                                      {t.label}
                                    </button>
                                  ))}
                                </div>
                              </div>

                              {/* Custom Item Form */}
                              <div className="flex gap-2 items-center">
                                <input
                                  type="text"
                                  placeholder="Contoh: Kerdus Oleh-oleh"
                                  className="flex-grow text-[11px] border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-emerald-500 bg-white text-slate-800 font-medium"
                                  value={customItemName}
                                  onChange={(e) => setCustomItemName(e.target.value)}
                                />
                                <input
                                  type="number"
                                  min="1"
                                  className="w-12 text-center text-[11px] border border-slate-200 rounded-xl py-1.5 focus:outline-none focus:border-emerald-500 bg-white text-slate-850 font-mono font-bold"
                                  value={customItemQty}
                                  onChange={(e) => setCustomItemQty(parseInt(e.target.value) || 1)}
                                />
                                <button
                                  type="button"
                                  disabled={!customItemName.trim()}
                                  onClick={() => {
                                    if (!customItemName.trim()) return;
                                    const name = customItemName.trim();
                                    const qty = customItemQty;
                                    const existing = additionalItems.find(item => item.name.toLowerCase() === name.toLowerCase());
                                    if (existing) {
                                      setAdditionalItems(additionalItems.map(item => 
                                        item.name.toLowerCase() === name.toLowerCase() ? { ...item, qty: item.qty + qty } : item
                                      ));
                                    } else {
                                      setAdditionalItems([...additionalItems, { name, qty }]);
                                    }
                                    setCustomItemName("");
                                    setCustomItemQty(1);
                                  }}
                                  className="bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 text-white font-bold text-[11px] px-3.5 py-1.5 rounded-xl transition-colors cursor-pointer"
                                >
                                  Tambah
                                </button>
                              </div>
                            </div>

                            {/* Quick Select: Keduanya */}
                            <button
                              id="btn-select-both"
                              onClick={() => {
                                setKoperBesar(true);
                                setKoperBesarCount(1);
                                setKoperKecil(true);
                                setKoperKecilCount(1);
                              }}
                              className="w-full bg-slate-50 hover:bg-slate-100 text-slate-700 py-2.5 rounded-xl text-xs font-bold transition-colors font-sans cursor-pointer border border-slate-200/60"
                            >
                              ✨ Checklist Lengkap (1 Besar, 1 Kecil)
                            </button>

                            {/* Submit Log Checklist */}
                            <button
                              id="btn-save-checklist"
                              onClick={handleSaveLog}
                              disabled={isSavingLog}
                              className="w-full rounded-xl bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 py-3.5 font-bold text-white shadow-lg shadow-emerald-600/10 active:scale-[0.98] transition-all text-xs flex items-center justify-center gap-2 cursor-pointer disabled:from-slate-400 disabled:to-slate-500"
                            >
                              {isSavingLog ? (
                                <>
                                  <Loader2 className="h-4 w-4 animate-spin text-white" /> MENYIMPAN KE GOOGLE SHEETS...
                                </>
                              ) : (
                                <>
                                  <Check className="h-4.5 w-4.5" /> SIMPAN CHECKLIST KOPER JAMAAH
                                </>
                              )}
                            </button>

                            {/* Cancel active check button */}
                            <button
                              type="button"
                              onClick={() => {
                                setOcrResult(null);
                                setMatchedPilgrim(null);
                                setFoundInManifest(null);
                                setKoperBesar(false);
                                setKoperBesarCount(1);
                                setKoperKecil(false);
                                setKoperKecilCount(1);
                                setAdditionalItems([]);
                              }}
                              className="w-full text-slate-400 hover:text-slate-600 text-[10px] py-1 cursor-pointer font-bold uppercase tracking-wider transition-colors"
                            >
                              Batal & Reset Verifikasi
                            </button>

                          </div>
                        </motion.div>
                      )}
                    </div>
                  </div>

                {/* Real-time statistics counters and graphical summary bar */}
                <div id="stats-summary-card" className="bg-white p-6 rounded-2xl border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.02)] grid grid-cols-1 md:grid-cols-3 gap-6">
                  
                  {/* Stat 1 */}
                  <div className="flex items-center gap-4 text-left">
                    <div className="bg-slate-50 p-3.5 rounded-xl text-slate-700 border border-slate-100 shadow-inner shrink-0">
                      <Luggage className="h-6 w-6 text-emerald-600" />
                    </div>
                    <div className="truncate">
                      <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block">Grup Terpilih</span>
                      <span className="text-sm font-extrabold text-slate-900 block truncate max-w-[220px]" title={activeGroup}>
                        {activeGroup}
                      </span>
                      <span className="text-[10px] font-bold text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded-md inline-block mt-0.5 truncate max-w-[220px]">
                        Checklist: {activeStation}
                      </span>
                    </div>
                  </div>

                  {/* Stat 2 */}
                  <div className="flex items-center gap-4 text-left">
                    <div className="bg-slate-50 p-3.5 rounded-xl text-slate-700 border border-slate-100 shadow-inner shrink-0">
                      <CheckCircle2 className="h-6 w-6 text-emerald-600" />
                    </div>
                    <div className="flex-grow">
                      <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block">Kelengkapan Log</span>
                      <span className="text-lg font-extrabold text-slate-900 leading-tight block">
                        {stats.checked} <span className="text-xs text-slate-400 font-bold">dari {stats.total} jamaah</span>
                      </span>
                      <span className="text-[10px] text-slate-400 block mt-0.5 font-bold">Terverifikasi di lapangan</span>
                    </div>
                  </div>

                  {/* Stat 3 - Visual progress bar */}
                  <div className="flex flex-col justify-center text-left">
                    <div className="flex justify-between items-center mb-2 text-[9px] font-bold text-slate-400 tracking-wider">
                      <span>PERSENTASE PENGECEKAN</span>
                      <span className="text-emerald-600 font-extrabold">{stats.percentage}%</span>
                    </div>
                    <div className="w-full bg-slate-50 border border-slate-100 rounded-full h-3 overflow-hidden p-0.5 shadow-inner">
                      <div 
                        className="bg-gradient-to-r from-emerald-500 to-teal-500 h-full rounded-full transition-all duration-500" 
                        style={{ width: `${stats.percentage}%` }}
                      ></div>
                    </div>
                  </div>

                </div>

                {/* 3.5. Bus Breakdown Panel (Mengecek kurang berapa per bus) */}
                {stats.busBreakdown && Object.keys(stats.busBreakdown).length > 0 && (
                  <div id="bus-status-grid-card" className="bg-white p-6 rounded-2xl border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.02)] text-left">
                    <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 border-b border-slate-100 pb-4 mb-5">
                      <div>
                        <h3 className="text-xs font-bold uppercase tracking-widest text-slate-400 flex items-center gap-2">
                          <Bus className="h-4.5 w-4.5 text-emerald-600 animate-pulse" /> STATUS CHECKLIST PER BUS / ROMBONGAN KECIL
                        </h3>
                        <p className="text-xs text-slate-500 mt-1">
                          Klik salah satu Bus di bawah ini untuk menyaring log jamaah atau melihat jamaah yang belum scan secara real-time.
                        </p>
                      </div>
                      {selectedBusFilter && (
                        <button
                          onClick={() => setSelectedBusFilter("")}
                          className="text-[10px] font-bold text-rose-600 hover:text-white hover:bg-rose-600 bg-rose-50 px-3 py-1.5 rounded-xl border border-rose-100 smooth-transition cursor-pointer"
                        >
                          Reset Filter Bus ({selectedBusFilter})
                        </button>
                      )}
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
                      {Object.entries(stats.busBreakdown).map(([busName, bStat]) => {
                        const isFiltered = selectedBusFilter === busName;
                        const pct = bStat.total > 0 ? Math.round((bStat.checked / bStat.total) * 100) : 0;
                        return (
                          <div
                            key={busName}
                            onClick={() => setSelectedBusFilter(isFiltered ? "" : busName)}
                            className={`p-4 rounded-xl border transition-all duration-200 cursor-pointer text-left relative overflow-hidden ${
                              isFiltered 
                                ? "bg-emerald-50/50 border-emerald-400 ring-2 ring-emerald-500/20 shadow-md" 
                                : bStat.remaining === 0 
                                  ? "bg-slate-50/50 border-slate-200 hover:border-slate-300"
                                  : "bg-white border-slate-100 hover:border-emerald-300 hover:shadow-md"
                            }`}
                          >
                            {/* Accent indicator */}
                            <div className={`absolute top-0 left-0 w-1.5 h-full ${
                              bStat.remaining === 0 ? "bg-emerald-500" : "bg-teal-500"
                            }`}></div>

                            <div className="pl-2">
                              <div className="flex justify-between items-start mb-1.5">
                                <span className="font-bold text-slate-800 text-xs sm:text-sm flex items-center gap-1.5">
                                  <Bus className="h-4 w-4 text-slate-400" />
                                  {busName}
                                </span>
                                {bStat.remaining === 0 ? (
                                  <span className="text-[9px] bg-emerald-50 text-emerald-800 font-extrabold px-2 py-0.5 rounded-full border border-emerald-100 uppercase tracking-wider">
                                    Lengkap
                                  </span>
                                ) : (
                                  <span className="text-[9px] bg-amber-50 text-amber-800 font-extrabold px-2 py-0.5 rounded-full border border-amber-100 uppercase tracking-wider">
                                    Kurang {bStat.remaining}
                                  </span>
                                )}
                              </div>

                              <div className="mt-3 flex items-baseline gap-1">
                                <span className="text-xl font-extrabold text-slate-950">{bStat.checked}</span>
                                <span className="text-[10px] text-slate-400 font-bold">/ {bStat.total} Jamaah</span>
                              </div>

                              {/* Small progress bar */}
                              <div className="mt-3 space-y-1">
                                <div className="flex justify-between items-center text-[9px] font-bold text-slate-400">
                                  <span>PROGRES</span>
                                  <span>{pct}%</span>
                                </div>
                                <div className="w-full bg-slate-50 border border-slate-100 rounded-full h-1.5 overflow-hidden p-0.5">
                                  <div
                                    className={`h-full rounded-full transition-all duration-300 ${
                                      bStat.remaining === 0 ? "bg-emerald-500" : "bg-emerald-600"
                                    }`}
                                    style={{ width: `${pct}%` }}
                                  ></div>
                                </div>
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    {/* Displaying missing list specifically if there is remaining */}
                    {selectedBusFilter && stats.busBreakdown[selectedBusFilter]?.remaining > 0 && (
                      <div className="mt-5 p-4 bg-amber-50/50 border border-amber-100/80 rounded-xl">
                        <span className="text-[10px] font-extrabold text-amber-800 block mb-2.5 uppercase tracking-widest flex items-center gap-1.5">
                          <AlertTriangle className="h-4 w-4 text-amber-600 shrink-0" />
                          Daftar Jamaah {selectedBusFilter} yang Belum Scan Koper ({stats.busBreakdown[selectedBusFilter].remaining} Orang):
                        </span>
                        <div className="flex flex-wrap gap-2">
                          {(manifest.filter(p => p.grup === activeGroup).length > 0 
                            ? manifest.filter(p => p.grup === activeGroup) 
                            : mockManifest.filter(p => p.grup === activeGroup)
                          )
                            .filter(p => {
                              const bName = p.bus ? p.bus.trim() : "Tanpa Bus";
                              return bName === selectedBusFilter;
                            })
                            .filter(p => !logs.some(l => 
                              l.grup === activeGroup && 
                              l.pos === activeStation && 
                              l.nama.trim().toLowerCase() === p.nama.trim().toLowerCase() && 
                              l.paspor.trim().toLowerCase() === p.paspor.trim().toLowerCase()
                            ))
                            .map((p, pIdx) => (
                              <span 
                                key={pIdx} 
                                className="inline-flex items-center gap-1.5 bg-white border border-amber-200/80 px-3 py-1.5 rounded-xl text-xs text-slate-700 font-semibold shadow-sm"
                              >
                                <span className="w-1.5 h-1.5 bg-amber-500 rounded-full animate-pulse"></span>
                                {p.nama} <span className="text-[10px] text-slate-400 font-mono font-bold">{p.paspor}</span>
                              </span>
                            ))
                          }
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {/* 4. Activity Logs Ledger for ACTIVE GROUP & STATION */}
                {/* 4. Activity Logs Ledger for ACTIVE GROUP & STATION */}
                <div id="activity-log-ledger" className="bg-white rounded-2xl border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.02)] overflow-hidden transition-all duration-300 hover:shadow-[0_8px_30px_rgb(0,0,0,0.04)]">
                  
                  {/* Header */}
                  <div className="p-6 border-b border-slate-100 flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 text-left">
                    <div>
                      <h2 className="text-xs font-bold uppercase tracking-widest text-slate-400 flex items-center gap-2">
                        <FileSpreadsheet className="h-4.5 w-4.5 text-emerald-600" /> DATA TERCATAT LOKASI CHECKLIST INI
                      </h2>
                      <p className="text-xs text-slate-500 mt-1">
                        Daftar lengkap koper jemaah rombongan {activeGroup} yang telah diverifikasi di Checklist {activeStation}.
                      </p>
                    </div>

                    {/* Actions Panel (Search, PDF report, Refresh) */}
                    <div className="flex flex-wrap items-center gap-3 self-stretch lg:self-auto w-full lg:w-auto">
                      
                      {/* Search box */}
                      <div className="relative flex-grow sm:w-60">
                        <Search className="absolute left-3 top-2.5 h-4 w-4 text-slate-400" />
                        <input
                          type="text"
                          placeholder="Filter Cari Nama / Paspor..."
                          className="w-full pl-9 pr-4 py-2 text-xs border border-slate-200 rounded-xl focus:outline-none focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 bg-slate-50/50 text-slate-800 transition-all font-medium"
                          value={searchQuery}
                          onChange={(e) => setSearchQuery(e.target.value)}
                        />
                      </div>

                      {/* Download PDF Report for active station logs */}
                      <button
                        onClick={() => {
                          const activeStationLogs = logs.filter(l => l.grup === activeGroup && l.pos === activeStation);
                          const groupPilgrims = manifest.filter(p => p.grup === activeGroup);
                          const manifestToUse = groupPilgrims.length > 0 ? groupPilgrims : mockManifest.filter(p => p.grup === activeGroup);
                          generatePdfReport(activeGroup, activeStation, activeStationLogs, manifestToUse);
                          showToast(`Laporan PDF '${activeGroup} - ${activeStation}' berhasil diunduh!`);
                        }}
                        className="bg-rose-600 hover:bg-rose-700 text-white px-4 py-2 rounded-xl text-xs font-bold smooth-transition flex items-center gap-1.5 shadow-md shadow-rose-600/10 cursor-pointer"
                      >
                        <FileSpreadsheet className="h-4 w-4" />
                        <span>Unduh PDF</span>
                      </button>

                      {/* Refresh logs manually */}
                      <button
                        id="btn-refresh-logs"
                        onClick={fetchDatabaseData}
                        className="bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-200/60 px-4 py-2 rounded-xl text-xs font-bold smooth-transition flex items-center gap-1.5 cursor-pointer"
                      >
                        <RefreshCw className="h-4 w-4 text-slate-500 animate-spin-slow" />
                        <span>Sinkronkan</span>
                      </button>

                    </div>
                  </div>

                  {/* Table Container */}
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse text-slate-650">
                      <thead>
                        <tr className="bg-slate-50/70 text-slate-400 text-[10px] font-bold uppercase tracking-wider border-b border-slate-100">
                          <th className="p-4 pl-6">Waktu Scan</th>
                          <th className="p-4">Nama Jamaah</th>
                          <th className="p-4">Nomor Paspor</th>
                          <th className="p-4 text-center">Koper Besar</th>
                          <th className="p-4 text-center">Koper Kecil</th>
                          <th className="p-4">Keterangan / Tambahan</th>
                          <th className="p-4">Petugas Lapangan</th>
                          <th className="p-4 text-center">Status</th>
                          <th className="p-4 text-center pr-6">Aksi</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {(() => {
                          const filtered = logs
                            .filter(l => l.grup === activeGroup && l.pos === activeStation)
                            .filter(l => {
                              if (!selectedBusFilter) return true;
                              const pilgrim = manifest.find(p => 
                                p.nama.trim().toLowerCase() === l.nama.trim().toLowerCase() && 
                                p.paspor.trim().toLowerCase() === l.paspor.trim().toLowerCase()
                              ) || mockManifest.find(p => 
                                p.nama.trim().toLowerCase() === l.nama.trim().toLowerCase() && 
                                p.paspor.trim().toLowerCase() === l.paspor.trim().toLowerCase()
                              );
                              const bName = pilgrim?.bus ? pilgrim.bus.trim() : "Tanpa Bus";
                              return bName === selectedBusFilter;
                            })
                            .filter(l => {
                              const q = searchQuery.toLowerCase();
                              return (
                                l.nama.toLowerCase().includes(q) ||
                                l.paspor.toLowerCase().includes(q) ||
                                (l.status && l.status.toLowerCase().includes(q))
                              );
                            });

                          if (filtered.length === 0) {
                            return (
                              <tr>
                                <td colSpan={9} className="p-10 text-center text-slate-400 font-bold leading-relaxed">
                                  {searchQuery ? "Tidak ditemukan hasil log yang cocok dengan pencarian." : "Belum ada jemaah yang diverifikasi di tempat Checklist ini. Silakan lakukan Checklist koper."}
                                </td>
                              </tr>
                            );
                          }

                          return [...filtered].reverse().map((log, idx) => {
                            // Extract custom items if written in status
                            const hasAdditional = log.status && log.status.includes("| Tambahan:");
                            const additionalText = hasAdditional ? log.status.split("| Tambahan:")[1].trim() : "-";
                            const cleanStatus = log.status && log.status.includes("|") ? log.status.split("|")[0].trim() : log.status;

                            return (
                              <tr key={idx} className="hover:bg-slate-50/40 transition-colors text-slate-700">
                                <td className="p-4 pl-6 font-mono text-[10px] text-slate-400 font-bold">{log.timestamp}</td>
                                <td className="p-4 font-bold text-slate-900">{log.nama}</td>
                                <td className="p-4 font-mono font-bold text-[11px] text-slate-500 bg-slate-50 px-2 py-0.5 rounded border border-slate-100/50 inline-block mt-3">{log.paspor}</td>
                                <td className="p-4 text-center font-mono font-extrabold text-slate-800">{log.koperBesar === "Ada" ? "✓ 1 Pcs" : log.koperBesar || "Tidak Ada"}</td>
                                <td className="p-4 text-center font-mono font-extrabold text-slate-800">{log.koperKecil === "Ada" ? "✓ 1 Pcs" : log.koperKecil || "Tidak Ada"}</td>
                                <td className="p-4 text-slate-500 font-semibold">{additionalText}</td>
                                <td className="p-4 text-slate-500 font-medium">{log.petugas}</td>
                                <td className="p-4 text-center">
                                  <span className={`inline-flex px-2.5 py-1 rounded-full text-[9px] font-bold border ${
                                    cleanStatus === "Lengkap" 
                                      ? "bg-emerald-50 text-emerald-800 border-emerald-200" 
                                      : "bg-amber-50 text-amber-800 border-amber-200"
                                  }`}>
                                    {cleanStatus.toUpperCase()}
                                  </span>
                                </td>
                                <td className="p-4 text-center pr-6">
                                  <div className="flex items-center justify-center gap-2">
                                    <button
                                      onClick={() => {
                                        setEditingLog(log);
                                        setEditedKoperBesar(log.koperBesar);
                                        setEditedKoperKecil(log.koperKecil);
                                        setEditedStatus(cleanStatus);
                                        setEditedPetugas(log.petugas || "");
                                      }}
                                      className="p-1.5 text-teal-600 hover:bg-teal-50 hover:text-teal-700 rounded-lg transition-colors cursor-pointer border border-transparent hover:border-teal-100"
                                      title="Edit Log Bagasi"
                                    >
                                      <Edit2 className="h-4 w-4" />
                                    </button>
                                    <button
                                      onClick={() => handleDeleteLog(log)}
                                      className="p-1.5 text-rose-600 hover:bg-rose-50 hover:text-rose-700 rounded-lg transition-colors cursor-pointer border border-transparent hover:border-rose-100"
                                      title="Hapus Log Bagasi"
                                    >
                                      <Trash2 className="h-4 w-4" />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            );
                          });
                        })()}
                      </tbody>
                    </table>
                  </div>

                </div>

              </div>
            )}

          </div>
        ) : (
          <div className="space-y-6">
            
            {/* Header / Intro */}
            <div className="bg-gradient-to-br from-slate-900 to-slate-800 rounded-2xl p-6 text-white border border-slate-700/50 shadow-xl relative overflow-hidden">
              <div className="absolute top-0 right-0 -mr-16 -mt-16 w-64 h-64 bg-teal-600/10 rounded-full blur-3xl pointer-events-none"></div>
              <div className="relative z-10 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                <div>
                  <span className="text-[10px] bg-teal-600 px-2.5 py-1 rounded-full font-bold uppercase tracking-wider text-white shadow-sm">
                    Sesi Operasional Aktif
                  </span>
                  <h2 className="text-xl sm:text-2xl font-bold font-display tracking-tight mt-2">
                    Daftar Grup Jamaah Operasional
                  </h2>
                  <p className="text-xs text-slate-300 mt-1 max-w-2xl">
                    Pilih grup jamaah di bawah ini untuk mengelola checklist koper, memindai barcode stiker manifest, dan melihat progres verifikasi real-time.
                  </p>
                </div>
                <div className="flex items-center gap-3 bg-slate-800/80 border border-slate-700/60 p-3 rounded-xl font-mono">
                  <div className="text-left">
                    <p className="text-[10px] text-slate-400">Total Grup</p>
                    <p className="text-lg font-bold text-teal-400">{groups.length} Rombongan</p>
                  </div>
                </div>
              </div>
            </div>

            {/* Bulk Excel Paste Section has been moved inside each card */}

            {/* Grid of Group Cards */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {[...groups].reverse().map((groupName, idx) => {
                const groupPilgrims = manifest.filter(p => p.grup === groupName);
                const groupLogs = logs.filter(l => l.grup === groupName);
                const checkedNames = new Set(groupLogs.map(l => l.nama));
                const checkedCount = checkedNames.size;
                const totalCount = groupPilgrims.length;
                const percentage = totalCount > 0 ? Math.round((checkedCount / totalCount) * 100) : 0;
                const isFormOpen = expandedAddPilgrimGrup === groupName;

                return (
                  <div 
                    key={idx}
                    className="bg-white border border-slate-100 rounded-2xl shadow-[0_8px_30px_rgb(0,0,0,0.02)] overflow-hidden flex flex-col justify-between hover:border-emerald-500 hover:shadow-[0_12px_40px_rgba(16,185,129,0.06)] transition-all duration-300 group"
                  >
                    <div className="p-5 flex-grow">
                      <div className="flex justify-between items-start mb-4">
                        <div className="bg-slate-50 group-hover:bg-emerald-50 p-2.5 rounded-xl border border-slate-100 group-hover:border-emerald-100 transition-colors">
                          <Users className="h-5 w-5 text-slate-600 group-hover:text-emerald-600" />
                        </div>
                        <span className={`text-[10px] font-extrabold px-2.5 py-1 rounded-full border ${
                          percentage === 100 
                            ? "bg-emerald-50 text-emerald-800 border-emerald-100" 
                            : percentage > 0 
                              ? "bg-teal-50 text-teal-800 border-teal-100" 
                              : "bg-slate-50 text-slate-700 border-slate-100"
                        }`}>
                          {percentage}% Selesai
                        </span>
                      </div>

                      {editingGroup === groupName ? (
                        <div className="flex items-center gap-1.5 mt-2" onClick={(e) => e.stopPropagation()}>
                          <input
                            type="text"
                            className="text-xs border border-slate-200 rounded-xl px-3 py-1.5 w-full text-slate-800 focus:outline-none focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 transition-all bg-white font-semibold"
                            value={newGroupNameInput}
                            onChange={(e) => setNewGroupNameInput(e.target.value)}
                            autoFocus
                          />
                          <button
                            onClick={() => handleEditGroup(groupName, newGroupNameInput)}
                            className="bg-emerald-600 hover:bg-emerald-700 text-white p-2 rounded-xl transition-colors cursor-pointer"
                            title="Simpan"
                            disabled={isUpdatingGroup}
                          >
                            {isUpdatingGroup ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Check className="h-3.5 w-3.5" />}
                          </button>
                          <button
                            onClick={() => setEditingGroup(null)}
                            className="bg-slate-100 hover:bg-slate-200 text-slate-700 p-2 rounded-xl transition-colors cursor-pointer"
                            title="Batal"
                          >
                            ✕
                          </button>
                        </div>
                      ) : (
                        <div className="flex justify-between items-center mb-2">
                          <h3 className="text-sm font-extrabold font-display text-slate-900 line-clamp-1 group-hover:text-emerald-600 transition-colors">
                            {groupName}
                          </h3>
                          <div className="flex gap-1 items-center opacity-70 group-hover:opacity-100 transition-opacity">
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                setEditingGroup(groupName);
                                setNewGroupNameInput(groupName);
                              }}
                              className="text-slate-400 hover:text-emerald-600 p-1.5 rounded-lg hover:bg-slate-50 cursor-pointer transition-colors"
                              title="Edit Nama Rombongan"
                            >
                              <Edit2 className="h-3.5 w-3.5" />
                            </button>
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                handleDeleteGroup(groupName);
                              }}
                              className="text-slate-400 hover:text-rose-600 p-1.5 rounded-lg hover:bg-slate-50 cursor-pointer transition-colors"
                              title="Hapus Rombongan"
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                            </button>
                          </div>
                        </div>
                      )}

                      <div className="space-y-2 text-xs text-slate-500 mt-4 border-t border-slate-100/80 pt-3">
                        <div className="flex justify-between">
                          <span className="font-semibold text-slate-400">Manifes Jamaah:</span>
                          <span className="font-extrabold text-slate-850">{totalCount} orang</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="font-semibold text-slate-400">Telah Terverifikasi:</span>
                          <span className="font-extrabold text-slate-850">{checkedCount} jamaah</span>
                        </div>
                        <div className="w-full bg-slate-50 border border-slate-100 rounded-full h-2 mt-1.5 overflow-hidden p-0.5">
                          <div 
                            className="bg-gradient-to-r from-emerald-500 to-teal-500 h-full rounded-full transition-all duration-500"
                            style={{ width: `${percentage}%` }}
                          ></div>
                        </div>
                      </div>

                      {/* Manifest Preview list */}
                      {groupPilgrims.length > 0 && (
                        <div className="mt-4 bg-slate-50/50 p-3 rounded-xl border border-slate-100">
                          <p className="text-[10px] font-extrabold text-slate-400 uppercase tracking-widest mb-2">Manifes Terdaftar</p>
                          <div className="space-y-1.5">
                            {groupPilgrims.slice(0, 3).map((p, pIdx) => (
                              <div key={pIdx} className="text-[10px] text-slate-700 flex justify-between font-semibold">
                                <span className="truncate max-w-[130px]">{p.nama}</span>
                                <span className="font-mono text-slate-400 font-bold">{p.paspor}</span>
                              </div>
                            ))}
                            {groupPilgrims.length > 3 && (
                              <div className="text-[9px] text-slate-450 text-right mt-1.5 font-bold">
                                + {groupPilgrims.length - 3} jamaah lainnya
                              </div>
                            )}
                          </div>
                        </div>
                      )}

                    </div>

                    <div className="p-4 bg-slate-50 border-t border-slate-100/80 flex items-center justify-between">
                      <span className="text-[10px] text-slate-450 font-bold">Mulai pemindaian koper grup</span>
                      <button
                        onClick={() => {
                          setActiveGroup(groupName);
                          setOcrResult(null);
                          setMatchedPilgrim(null);
                          setFoundInManifest(null);
                        }}
                        className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-4 py-2 rounded-xl shadow-md shadow-emerald-600/10 smooth-transition flex items-center gap-1 cursor-pointer"
                      >
                        Buka Cheklist <ChevronRight className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  </div>
                );
              })}

              {/* Dynamic Card: Create Group */}
              <div className="bg-emerald-50/10 border-2 border-dashed border-emerald-300/60 rounded-2xl p-6 flex flex-col justify-between hover:border-emerald-500 hover:bg-emerald-50/20 transition-all duration-300">
                <div>
                  <div className="bg-emerald-50 p-2.5 rounded-xl border border-emerald-100 self-start inline-block mb-3 text-emerald-600">
                    <Folder className="h-5 w-5" />
                  </div>
                  <h3 className="text-sm font-extrabold font-display text-slate-900 mb-1">
                    Buat Rombongan Baru
                  </h3>
                  <p className="text-xs text-slate-500 leading-relaxed mb-5">
                    Tambahkan grup jemaah baru ke database Google Sheets Anda untuk memisahkan manifestasi jemaah per rombongan operasional.
                  </p>
                </div>

                <form onSubmit={handleCreateGroup} className="space-y-3">
                  <div>
                    <input
                      type="text"
                      required
                      placeholder="Contoh: Rombongan Haji Banjar"
                      className="w-full text-xs border border-slate-200 rounded-xl px-3 py-2 bg-white focus:outline-none focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 text-slate-800 font-semibold transition-all"
                      value={newGroupName}
                      onChange={(e) => setNewGroupName(e.target.value)}
                    />
                  </div>
                  <button
                    type="submit"
                    disabled={isCreatingGroup || !newGroupName.trim()}
                    className="w-full bg-slate-900 hover:bg-slate-950 text-white text-xs font-bold py-2.5 px-3 rounded-xl smooth-transition flex items-center justify-center gap-1.5 shadow-md shadow-slate-900/10 disabled:bg-slate-100 disabled:text-slate-400 cursor-pointer"
                  >
                    {isCreatingGroup ? (
                      <>
                        <Loader2 className="h-3.5 w-3.5 animate-spin" /> Membuat...
                      </>
                    ) : (
                      <>
                        <Plus className="h-4 w-4" /> Buat Grup Baru
                      </>
                    )}
                  </button>
                </form>
              </div>
            </div>
          </div>
        )}

        {/* PWA Promotion & Quick Install Panel (Placed neatly at the bottom) */}
        {showInstallBanner && !isStandalone && (
          <div id="pwa-install-banner" className="bg-white border border-slate-150 p-5 rounded-2xl shadow-[0_4px_20px_rgb(0,0,0,0.02)] text-left flex flex-col sm:flex-row justify-between items-center gap-4 mt-8">
            <div className="flex items-center gap-3">
              <div className="bg-indigo-50 text-indigo-600 p-2.5 rounded-xl border border-indigo-100 flex-shrink-0">
                <Download className="h-5 w-5" />
              </div>
              <div className="text-left">
                <h3 className="font-display font-bold text-xs text-slate-900">
                  Instal Aplikasi Koper Elhakim
                </h3>
                <p className="text-[10px] text-slate-500 mt-0.5">
                  Akses asisten koper lebih cepat dan hemat kuota langsung dari beranda HP atau desktop Anda.
                </p>
              </div>
            </div>
            <div className="flex items-center gap-2 w-full sm:w-auto shrink-0 justify-end">
              <button
                onClick={handleInstallPWA}
                className="bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold px-4.5 py-2.5 rounded-xl text-xs flex items-center gap-1.5 cursor-pointer shadow-sm shadow-indigo-600/10 hover:shadow-md transition-all active:scale-95 w-full sm:w-auto justify-center"
              >
                <Download className="h-4 w-4" />
                Instal Aplikasi
              </button>
              <button
                onClick={() => {
                  localStorage.setItem("koperjemaah_dismiss_install", "true");
                  setShowInstallBanner(false);
                }}
                className="text-slate-400 hover:text-slate-600 px-3 py-2 text-xs font-semibold cursor-pointer transition-colors"
              >
                Nanti Saja
              </button>
            </div>
          </div>
        )}
      </main>

      {/* Floating alert toasts */}
      <AnimatePresence>
        {successToast && (
          <motion.div
            id="toast-notification"
            initial={{ opacity: 0, scale: 0.9, y: 30 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.9, y: 30 }}
            className="fixed bottom-6 right-6 bg-slate-900/95 backdrop-blur-md text-white text-xs font-extrabold px-5 py-3.5 rounded-2xl shadow-2xl flex items-center gap-3 border border-slate-800 z-50"
          >
            <div className="bg-emerald-500/10 p-1 rounded-lg">
              <CheckCircle2 className="h-4.5 w-4.5 text-emerald-400 animate-bounce" />
            </div>
            <span>{successToast}</span>
          </motion.div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {editingLog && (
          <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 z-50 transition-opacity">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white border border-slate-100 rounded-3xl max-w-md w-full shadow-2xl p-6 relative text-left"
            >
              <h3 className="text-base font-extrabold text-slate-900 mb-1">
                Edit Verifikasi Bagasi Jamaah
              </h3>
              <p className="text-xs text-slate-500 mb-5">
                Ubah log pencatatan koper jamaah <strong>{editingLog.nama}</strong> ({editingLog.paspor})
              </p>

              <form onSubmit={handleEditLogSubmit} className="space-y-4">
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Status Koper Besar</label>
                  <select
                    className="w-full text-xs border border-slate-200 rounded-xl px-3 py-2.5 bg-white text-slate-800 font-semibold focus:outline-none focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 transition-all"
                    value={editedKoperBesar}
                    onChange={(e) => setEditedKoperBesar(e.target.value)}
                  >
                    <option value="Ada">Ada (1 Pcs)</option>
                    <option value="Tidak Ada">Tidak Ada</option>
                    <option value="2 Pcs">2 Pcs</option>
                    <option value="3 Pcs">3 Pcs</option>
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Status Koper Kecil</label>
                  <select
                    className="w-full text-xs border border-slate-200 rounded-xl px-3 py-2.5 bg-white text-slate-800 font-semibold focus:outline-none focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 transition-all"
                    value={editedKoperKecil}
                    onChange={(e) => setEditedKoperKecil(e.target.value)}
                  >
                    <option value="Ada">Ada (1 Pcs)</option>
                    <option value="Tidak Ada">Tidak Ada</option>
                    <option value="2 Pcs">2 Pcs</option>
                    <option value="3 Pcs">3 Pcs</option>
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Status Checklist</label>
                  <select
                    className="w-full text-xs border border-slate-200 rounded-xl px-3 py-2.5 bg-white text-slate-800 font-semibold focus:outline-none focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 transition-all"
                    value={editedStatus}
                    onChange={(e) => setEditedStatus(e.target.value)}
                  >
                    <option value="Lengkap">Lengkap</option>
                    <option value="Belum Lengkap">Belum Lengkap</option>
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Nama Petugas Lapangan</label>
                  <input
                    type="text"
                    className="w-full text-xs border border-slate-200 rounded-xl px-3 py-2.5 bg-white text-slate-800 font-semibold focus:outline-none focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 transition-all"
                    placeholder="Contoh: Budi Santoso"
                    value={editedPetugas}
                    onChange={(e) => setEditedPetugas(e.target.value)}
                  />
                </div>

                <div className="flex justify-end gap-2.5 pt-4 border-t border-slate-100 mt-6">
                  <button
                    type="button"
                    onClick={() => setEditingLog(null)}
                    className="bg-slate-50 hover:bg-slate-100 text-slate-700 text-xs font-bold px-4 py-2.5 rounded-xl border border-slate-200/40 smooth-transition cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    type="submit"
                    disabled={isUpdatingLog}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-5 py-2.5 rounded-xl smooth-transition shadow-lg shadow-emerald-600/10 flex items-center gap-1.5 cursor-pointer"
                  >
                    {isUpdatingLog ? (
                      <>
                        <Loader2 className="h-3.5 w-3.5 animate-spin" /> Menyimpan...
                      </>
                    ) : (
                      <>
                        <Check className="h-3.5 w-3.5" /> Simpan Perubahan
                      </>
                    )}
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <footer className="w-full bg-white border-t border-slate-100 py-5 mt-auto text-center text-xs text-slate-400 shrink-0">
        <p>© 2026 Checklist Koper jamaah Elhakim. Sistem Manajemen Koper Terintegrasi AI.</p>
      </footer>

      </div>

    </div>
  );
}

