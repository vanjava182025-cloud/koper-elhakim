/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3000;

// Increase payload limit for base64 image uploads
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ limit: "50mb", extended: true }));

// Initialize Gemini client (Only if key is available)
let ai: GoogleGenAI | null = null;
if (process.env.GEMINI_API_KEY) {
  ai = new GoogleGenAI({
    apiKey: process.env.GEMINI_API_KEY,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
} else {
  console.warn("WARNING: GEMINI_API_KEY is not defined. Image analysis features will fail.");
}

// In-memory mock database for Demo Mode
let mockCheckpointStations = [
  "Pemberangkatan Jakarta (CGK)",
  "Kedatangan Bandara Jeddah (JED)",
  "Penerimaan Hotel Mekkah",
  "Pelepasan Hotel Madinah",
  "Kepulangan Bandara Medina (MED)"
];

let mockGroups = [
  "Grup Sabilal - Riyadh",
  "Grup Al-Haram - Mecca",
  "Grup Nabawi - Medina"
];

let mockManifest = [
  { grup: "Grup Sabilal - Riyadh", nama: "Ahmad Fauzi", paspor: "A1234567", bus: "Bus 1" },
  { grup: "Grup Sabilal - Riyadh", nama: "Siti Aminah", paspor: "A7654321", bus: "Bus 1" },
  { grup: "Grup Sabilal - Riyadh", nama: "Budi Santoso", paspor: "A9876543", bus: "Bus 2" },
  { grup: "Grup Al-Haram - Mecca", nama: "Rudi Hermawan", paspor: "B2468135", bus: "Bus 1" },
  { grup: "Grup Al-Haram - Mecca", nama: "Dewi Lestari", paspor: "B1357924", bus: "Bus 2" },
  { grup: "Grup Nabawi - Medina", nama: "Anisa Rahmawati", paspor: "C9876123", bus: "Bus 1" },
  { grup: "Grup Nabawi - Medina", nama: "Muhammad Yusuf", paspor: "C1234987", bus: "Bus 1" }
];

let mockLogs: Array<{
  timestamp: string;
  grup: string;
  nama: string;
  paspor: string;
  pos: string;
  koperBesar: string;
  koperKecil: string;
  petugas: string;
  status: string;
}> = [];

// Helper function to calculate String Similarity (Fuzzy Match)
function similarity(s1: string, s2: string): number {
  let longer = s1.toLowerCase();
  let shorter = s2.toLowerCase();
  if (s1.length < s2.length) {
    longer = s2.toLowerCase();
    shorter = s1.toLowerCase();
  }
  const longerLength = longer.length;
  if (longerLength === 0) {
    return 1.0;
  }
  return (longerLength - editDistance(longer, shorter)) / longerLength;
}

function editDistance(s1: string, s2: string): number {
  const costs = [];
  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= s2.length; j++) {
      if (i === 0) {
        costs[j] = j;
      } else {
        if (j > 0) {
          let newValue = costs[j - 1];
          if (s1.charAt(i - 1) !== s2.charAt(j - 1)) {
            newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
          }
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0) {
      costs[s2.length] = lastValue;
    }
  }
  return costs[s2.length];
}

// Helper to perform safe fetch calls to Google Apps Script Web Apps
async function fetchGasSafe(url: string, options?: RequestInit) {
  const response = await fetch(url, options);
  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`HTTP ${response.status}: ${errText || response.statusText}`);
  }
  const text = await response.text();
  const trimmed = text.trim();
  
  // Detect Google Login screen or generic HTML error pages
  if (
    trimmed.startsWith("<") || 
    trimmed.includes("<!DOCTYPE") || 
    trimmed.includes("Sign in - Google Accounts") || 
    trimmed.includes("google.com/accounts") ||
    trimmed.includes("Service Login")
  ) {
    throw new Error("GAS Web App mengembalikan halaman login Google. Pastikan Anda telah men-deploy Web App di Apps Script dengan setelan 'Who has access: Anyone' (Siapa saja) dan sudah melakukan otorisasi penuh pada akun Google Anda.");
  }

  try {
    return JSON.parse(trimmed);
  } catch (err) {
    return { success: true, message: trimmed };
  }
}

// 1. Endpoint: Ambil Data Grup dan Pos
app.post("/api/groups-and-checkpoint-stations", async (req, res) => {
  const { mode, spreadsheetId, gasUrl } = req.body;
  const authHeader = req.headers.authorization;

  if (mode === "demo") {
    const groups = Array.from(new Set([...mockGroups, ...mockManifest.map(m => m.grup)]));
    return res.json({
      success: true,
      groups,
      stations: mockCheckpointStations,
      logs: mockLogs,
      manifest: mockManifest
    });
  }

  if (mode === "gas") {
    if (!gasUrl) {
      return res.status(400).json({ success: false, error: "GAS URL tidak boleh kosong." });
    }
    try {
      const data = await fetchGasSafe(`${gasUrl}?action=ambilDataGrupDanPos`);
      return res.json({ success: true, ...data });
    } catch (err: any) {
      console.error("GAS error:", err);
      return res.status(500).json({ success: false, error: "Gagal mengambil data dari GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) {
      return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    }
    if (!authHeader) {
      return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });
    }

    try {
      // 1. Get groups from Manifest tab
      const manifestRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A:D`,
        { headers: { Authorization: authHeader } }
      );
      
      if (!manifestRes.ok) {
        const errText = await manifestRes.text();
        return res.status(manifestRes.status).json({ 
          success: false, 
          error: `Gagal membaca worksheet Manifest. Silakan pastikan format spreadsheet sesuai atau inisialisasi ulang. Detail: ${errText}` 
        });
      }

      const manifestData = await manifestRes.json();
      const manifestRows = manifestData.values || [];
      
      // Rows: [ [Grup, Nama, Paspor], ... ] - Skip header
      const groupsSet = new Set<string>();
      if (manifestRows.length > 1) {
        for (let i = 1; i < manifestRows.length; i++) {
          if (manifestRows[i][0]) {
            groupsSet.add(manifestRows[i][0]);
          }
        }
      }
      const groups = Array.from(groupsSet);

      // 2. Get active checkpoint stations from Cheklist or Pos tab
      let posRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Cheklist!A:A`,
        { headers: { Authorization: authHeader } }
      );
      if (!posRes.ok) {
        posRes = await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Pos!A:A`,
          { headers: { Authorization: authHeader } }
        );
      }
      const posData = await posRes.json();
      const posRows = posData.values || [];
      const stations = posRows.slice(1).map((row: any) => row[0]).filter(Boolean);

      // 3. Get Logs
      const logsRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I`,
        { headers: { Authorization: authHeader } }
      );
      let logs: any[] = [];
      if (logsRes.ok) {
        const logsData = await logsRes.json();
        const logsRows = logsData.values || [];
        if (logsRows.length > 1) {
          logs = logsRows.slice(1).map((row: any) => ({
            timestamp: row[0] || "",
            grup: row[1] || "",
            nama: row[2] || "",
            paspor: row[3] || "",
            pos: row[4] || "",
            koperBesar: row[5] || "",
            koperKecil: row[6] || "",
            petugas: row[7] || "",
            status: row[8] || "Lengkap"
          }));
        }
      }

      const manifestList = manifestRows.slice(1).map((r: any) => ({
        grup: r[0] || "",
        nama: r[1] || "",
        paspor: r[2] || "",
        bus: r[3] || ""
      })).filter((p: any) => p.grup && p.nama);

      return res.json({
        success: true,
        groups,
        stations,
        logs,
        manifest: manifestList
      });
    } catch (err: any) {
      console.error("OAuth error:", err);
      return res.status(500).json({ success: false, error: "Gagal berinteraksi dengan Google Sheets: " + err.message });
    }
  }

  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// 2. Endpoint: Tambah Pos Checklist
app.post("/api/checkpoint-stations", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, namaPos } = req.body;
  const authHeader = req.headers.authorization;

  if (!namaPos || namaPos.trim() === "") {
    return res.status(400).json({ success: false, error: "Nama pos baru tidak boleh kosong." });
  }

  if (mode === "demo") {
    if (!mockCheckpointStations.includes(namaPos)) {
      mockCheckpointStations.push(namaPos);
    }
    return res.json({ success: true, stations: mockCheckpointStations });
  }

  if (mode === "gas") {
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "tambahPosChecklist", namaPos })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal menyimpan pos ke GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) {
      return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    }
    if (!authHeader) {
      return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });
    }

    try {
      let appendRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Cheklist!A:A:append?valueInputOption=USER_ENTERED`,
        {
          method: "POST",
          headers: {
            Authorization: authHeader,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            values: [[namaPos]]
          })
        }
      );

      if (!appendRes.ok) {
        appendRes = await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Pos!A:A:append?valueInputOption=USER_ENTERED`,
          {
            method: "POST",
            headers: {
              Authorization: authHeader,
              "Content-Type": "application/json"
            },
            body: JSON.stringify({
              values: [[namaPos]]
            })
          }
        );
      }

      if (!appendRes.ok) {
        throw new Error(await appendRes.text());
      }

      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal menyimpan pos ke Google Sheets: " + err.message });
    }
  }

  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// 3. Endpoint: Analisis Stiker dan Cari Jamaah
app.post("/api/analyze-sticker", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, image, grup } = req.body;
  const authHeader = req.headers.authorization;

  if (!image) {
    return res.status(400).json({ success: false, error: "Gambar stiker koper diperlukan." });
  }
  if (!grup) {
    return res.status(400).json({ success: false, error: "Silakan pilih Grup terlebih dahulu." });
  }

  try {
    // Process base64 image
    let base64Data = image;
    let mimeType = "image/jpeg";

    if (image.includes(";base64,")) {
      const parts = image.split(";base64,");
      mimeType = parts[0].split(":")[1] || "image/jpeg";
      base64Data = parts[1];
    }

    if (!ai) {
      return res.status(500).json({ 
        success: false, 
        error: "Gemini AI client tidak terinisialisasi. Silakan pastikan GEMINI_API_KEY telah dikonfigurasi di panel Secrets." 
      });
    }

    const imagePart = {
      inlineData: {
        mimeType,
        data: base64Data
      }
    };

    const textPart = {
      text: `Analyze this image which is a luggage sticker of an Islamic pilgrim (stiker koper jemaah umrah/haji). Identify the full name of the pilgrim (Nama Jemaah) and passport number (Nomor Paspor) if visible. Indonesian pilgrim stickers usually clearly have 'NAMA', 'PASPOR' or 'PASSPORT', and 'GRUP' or 'TRAVEL'. Identify these carefully.
      Return your response as a valid, raw JSON object with exactly the following keys:
      {
        "name": "Full Name Found",
        "passport": "Passport Number Found or null",
        "group": "Group or Travel Name Found or null"
      }
      Do NOT include any markdown code blocks, backticks, or extra text. Just output raw, pure JSON.`
    };

    // Use gemini-3.1-pro-preview as requested!
    const geminiResponse = await ai.models.generateContent({
      model: "gemini-3.1-pro-preview",
      contents: { parts: [imagePart, textPart] }
    });

    let rawText = geminiResponse.text || "{}";
    rawText = rawText.trim();
    // Strip markdown JSON codeblock formatting if present
    if (rawText.startsWith("```json")) {
      rawText = rawText.replace(/^```json\s*/, "").replace(/\s*```$/, "");
    } else if (rawText.startsWith("```")) {
      rawText = rawText.replace(/^```\s*/, "").replace(/\s*```$/, "");
    }

    let parsedResult = { name: "", passport: "", group: "" };
    try {
      parsedResult = JSON.parse(rawText);
    } catch (parseErr) {
      console.error("Failed to parse Gemini JSON:", rawText, parseErr);
      // Fallback regex parsing if JSON format is slightly off
      const nameMatch = rawText.match(/"name":\s*"([^"]+)"/);
      const passportMatch = rawText.match(/"passport":\s*"([^"]+)"/);
      parsedResult = {
        name: nameMatch ? nameMatch[1] : "Tidak Terdeteksi",
        passport: passportMatch ? passportMatch[1] : "",
        group: ""
      };
    }

    const ocrName = parsedResult.name || "Tidak Terdeteksi";
    const ocrPassport = parsedResult.passport || "";

    // Now, run the 'cariJamaahDalamManifest' logic with the active group
    let matchedPilgrim: any = null;
    let foundInManifest = false;
    let matchScore = 0;
    let pilgrimsList: any[] = [];

    if (mode === "demo") {
      pilgrimsList = mockManifest.filter(p => p.grup === grup);
    } else if (mode === "gas") {
      try {
        const gasRes = await fetch(`${gasUrl}?action=ambilDataGrupDanPos`);
        const gasData = await gasRes.json();
        // Assume GAS return manifest or can be searched
        if (gasData.manifest) {
          pilgrimsList = gasData.manifest.filter((p: any) => p.grup === grup);
        }
      } catch (e) {
        console.error("Failed to fetch manifest from GAS for matching:", e);
      }
    } else if (mode === "oauth") {
      const manifestRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A:D`,
        { headers: { Authorization: authHeader } }
      );
      if (manifestRes.ok) {
        const manifestData = await manifestRes.json();
        const rows = manifestData.values || [];
        if (rows.length > 1) {
          pilgrimsList = rows.slice(1).map((r: any) => ({
            grup: r[0] || "",
            nama: r[1] || "",
            paspor: r[2] || "",
            bus: r[3] || ""
          })).filter((p: any) => p.grup === grup);
        }
      }
    }

    // Fuzzy matching name within the active group
    if (pilgrimsList.length > 0 && ocrName && ocrName !== "Tidak Terdeteksi") {
      let bestMatch: any = null;
      let highestSim = 0;

      for (const pilgrim of pilgrimsList) {
        const sim = similarity(ocrName, pilgrim.nama);
        if (sim > highestSim) {
          highestSim = sim;
          bestMatch = pilgrim;
        }
      }

      // Threshold of 0.6 for similarity
      if (highestSim >= 0.6) {
        matchedPilgrim = bestMatch;
        foundInManifest = true;
        matchScore = highestSim;
      }
    }

    return res.json({
      success: true,
      ocrResult: {
        name: ocrName,
        passport: ocrPassport,
        detectedGroup: parsedResult.group || ""
      },
      matchInfo: {
        foundInManifest,
        matchScore,
        pilgrim: matchedPilgrim || {
          nama: ocrName,
          paspor: ocrPassport || "Tidak Ditemukan",
          grup: grup
        }
      }
    });

  } catch (err: any) {
    console.error("Image analysis failed:", err);
    return res.status(500).json({ success: false, error: "Gagal menganalisis gambar: " + err.message });
  }
});

// 4. Endpoint: Simpan Log Checklist
app.post("/api/save-log", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, logData } = req.body;
  const authHeader = req.headers.authorization;

  const { grup, nama, paspor, pos, koperBesar, koperKecil, petugas, koperBesarStr: customKoperBesarStr, koperKecilStr: customKoperKecilStr, statusStr: customStatusStr } = logData;

  const timestamp = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });
  
  const koperBesarStr = customKoperBesarStr !== undefined ? customKoperBesarStr : (koperBesar ? "Ada" : "Tidak Ada");
  const koperKecilStr = customKoperKecilStr !== undefined ? customKoperKecilStr : (koperKecil ? "Ada" : "Tidak Ada");
  const statusStr = customStatusStr !== undefined ? customStatusStr : ((koperBesar && koperKecil) ? "Lengkap" : "Tidak Lengkap");

  if (mode === "demo") {
    const newLog = {
      timestamp,
      grup,
      nama,
      paspor,
      pos,
      koperBesar: koperBesarStr,
      koperKecil: koperKecilStr,
      petugas: petugas || "Demo User",
      status: statusStr
    };
    mockLogs.unshift(newLog); // Put latest on top
    return res.json({ success: true, logs: mockLogs });
  }

  if (mode === "gas") {
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "simpanLogChecklist",
          data: {
            timestamp,
            grup,
            nama,
            paspor,
            pos,
            koperBesar: koperBesarStr,
            koperKecil: koperKecilStr,
            petugas: petugas || "GAS User",
            status: statusStr
          }
        })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal menyimpan log ke GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) {
      return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    }
    if (!authHeader) {
      return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });
    }

    try {
      const appendRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I:append?valueInputOption=USER_ENTERED`,
        {
          method: "POST",
          headers: {
            Authorization: authHeader,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            values: [[
              timestamp,
              grup,
              nama,
              paspor,
              pos,
              koperBesarStr,
              koperKecilStr,
              petugas || "Petugas Lapangan",
              statusStr
            ]]
          })
        }
      );

      if (!appendRes.ok) {
        throw new Error(await appendRes.text());
      }

      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal menyimpan log ke Google Sheets: " + err.message });
    }
  }

  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// Endpoint: Tambah Grup Baru
app.post("/api/add-group", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, grup } = req.body;
  const authHeader = req.headers.authorization;

  if (!grup || grup.trim() === "") {
    return res.status(400).json({ success: false, error: "Nama grup tidak boleh kosong." });
  }

  if (mode === "demo") {
    if (!mockGroups.includes(grup)) {
      mockGroups.push(grup);
    }
    return res.json({ success: true, groups: Array.from(new Set([...mockGroups, ...mockManifest.map(m => m.grup)])) });
  }

  if (mode === "gas") {
    if (!gasUrl) return res.status(400).json({ success: false, error: "GAS URL tidak boleh kosong." });
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "tambahGrup", grup })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal menyimpan grup baru ke GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) {
      return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    }
    if (!authHeader) {
      return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });
    }

    try {
      const appendRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A:C:append?valueInputOption=USER_ENTERED`,
        {
          method: "POST",
          headers: {
            Authorization: authHeader,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            values: [[grup, "", ""]]
          })
        }
      );

      if (!appendRes.ok) {
        throw new Error(await appendRes.text());
      }

      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal menyimpan grup baru ke Google Sheets: " + err.message });
    }
  }

  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// Endpoint: Tambah Jemaah ke Manifest Grup
app.post("/api/add-pilgrim", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, grup, nama, paspor, bus } = req.body;
  const authHeader = req.headers.authorization;

  if (!grup || grup.trim() === "") {
    return res.status(400).json({ success: false, error: "Grup wajib ditentukan." });
  }
  if (!nama || nama.trim() === "") {
    return res.status(400).json({ success: false, error: "Nama jemaah tidak boleh kosong." });
  }
  if (!paspor || paspor.trim() === "") {
    return res.status(400).json({ success: false, error: "Nomor paspor jemaah tidak boleh kosong." });
  }

  if (mode === "demo") {
    mockManifest.push({ grup, nama, paspor, bus: bus || "" });
    return res.json({ success: true, manifest: mockManifest });
  }

  if (mode === "gas") {
    if (!gasUrl) return res.status(400).json({ success: false, error: "GAS URL tidak boleh kosong." });
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "tambahJemaah", grup, nama, paspor, bus })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal menyimpan jemaah ke GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) {
      return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    }
    if (!authHeader) {
      return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });
    }

    try {
      const appendRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A:D:append?valueInputOption=USER_ENTERED`,
        {
          method: "POST",
          headers: {
            Authorization: authHeader,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            values: [[grup, nama, paspor, bus || ""]]
          })
        }
      );

      if (!appendRes.ok) {
        throw new Error(await appendRes.text());
      }

      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal menyimpan jemaah ke Google Sheets: " + err.message });
    }
  }

  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// Endpoint: Tambah Banyak Jemaah sekaligus (Excel bulk paste)
app.post("/api/add-pilgrims-bulk", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, grup, pilgrims } = req.body;
  const authHeader = req.headers.authorization;

  if (!grup || grup.trim() === "") {
    return res.status(400).json({ success: false, error: "Grup wajib ditentukan." });
  }
  if (!pilgrims || !Array.isArray(pilgrims) || pilgrims.length === 0) {
    return res.status(400).json({ success: false, error: "Daftar jemaah kosong." });
  }

  // Validate pilgrim records
  const validPilgrims = pilgrims.filter(p => p.nama && p.nama.trim() !== "");
  if (validPilgrims.length === 0) {
    return res.status(400).json({ success: false, error: "Tidak ada jemaah dengan nama valid untuk ditambahkan." });
  }

  if (mode === "demo") {
    validPilgrims.forEach(p => {
      mockManifest.push({ grup, nama: p.nama.trim(), paspor: (p.paspor || "").trim(), bus: (p.bus || "").trim() });
    });
    return res.json({ success: true, manifest: mockManifest });
  }

  if (mode === "gas") {
    if (!gasUrl) return res.status(400).json({ success: false, error: "GAS URL tidak boleh kosong." });
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "tambahJemaahBulk", grup, pilgrims: validPilgrims })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal menyimpan jemaah bulk ke GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) {
      return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    }
    if (!authHeader) {
      return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });
    }

    try {
      const rows = validPilgrims.map(p => [grup, p.nama.trim(), (p.paspor || "").trim(), (p.bus || "").trim()]);
      const appendRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A:D:append?valueInputOption=USER_ENTERED`,
        {
          method: "POST",
          headers: {
            Authorization: authHeader,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            values: rows
          })
        }
      );

      if (!appendRes.ok) {
        throw new Error(await appendRes.text());
      }

      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal menyimpan jemaah ke Google Sheets: " + err.message });
    }
  }

  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// 5. Endpoint: Inisialisasi Template Google Sheets Baru
app.post("/api/initialize-spreadsheet", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (!authHeader) {
    return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });
  }

  try {
    // Step A: Create Spreadsheet
    const createRes = await fetch("https://sheets.googleapis.com/v4/spreadsheets", {
      method: "POST",
      headers: {
        Authorization: authHeader,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        properties: {
          title: "Sistem Manajemen Koper Jemaah"
        },
        sheets: [
          { properties: { title: "Manifest" } },
          { properties: { title: "Pos" } },
          { properties: { title: "Log" } }
        ]
      })
    });

    if (!createRes.ok) {
      const errTxt = await createRes.text();
      throw new Error(`Gagal membuat Spreadsheet: ${errTxt}`);
    }

    const sheetInfo = await createRes.json();
    const spreadsheetId = sheetInfo.spreadsheetId;

    // Step B: Seed Manifest Table
    const manifestRows = [
      ["Grup", "Nama Jemaah", "Nomor Paspor", "Bus/Rombongan Kecil"],
      ["Grup Sabilal - Riyadh", "Ahmad Fauzi", "A1234567", "Bus 1"],
      ["Grup Sabilal - Riyadh", "Siti Aminah", "A7654321", "Bus 1"],
      ["Grup Sabilal - Riyadh", "Budi Santoso", "A9876543", "Bus 2"],
      ["Grup Al-Haram - Mecca", "Rudi Hermawan", "B2468135", "Bus 1"],
      ["Grup Al-Haram - Mecca", "Dewi Lestari", "B1357924", "Bus 2"],
      ["Grup Nabawi - Medina", "Anisa Rahmawati", "C9876123", "Bus 1"],
      ["Grup Nabawi - Medina", "Muhammad Yusuf", "C1234987", "Bus 1"]
    ];

    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A1?valueInputOption=USER_ENTERED`,
      {
        method: "PUT",
        headers: {
          Authorization: authHeader,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ values: manifestRows })
      }
    );

    // Step C: Seed Pos Table
    const posRows = [
      ["Nama Pos Checklist"],
      ["Pemberangkatan Jakarta (CGK)"],
      ["Kedatangan Bandara Jeddah (JED)"],
      ["Penerimaan Hotel Mekkah"],
      ["Pelepasan Hotel Madinah"],
      ["Kepulangan Bandara Medina (MED)"]
    ];

    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Pos!A1?valueInputOption=USER_ENTERED`,
      {
        method: "PUT",
        headers: {
          Authorization: authHeader,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ values: posRows })
      }
    );

    // Step D: Seed Log Table headers
    const logHeaders = [
      ["Waktu Pencatatan", "Grup Jemaah", "Nama Jemaah", "Nomor Paspor", "Pos Checklist", "Koper Besar", "Koper Kecil", "Petugas", "Status Kelengkapan"]
    ];

    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A1?valueInputOption=USER_ENTERED`,
      {
        method: "PUT",
        headers: {
          Authorization: authHeader,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ values: logHeaders })
      }
    );

    return res.json({
      success: true,
      spreadsheetId,
      message: "Spreadsheet 'Sistem Manajemen Koper Jemaah' berhasil dibuat dan diisi template data!"
    });

  } catch (err: any) {
    console.error("Spreadsheet initialization error:", err);
    return res.status(500).json({ success: false, error: "Gagal membuat spreadsheet: " + err.message });
  }
});

// 6. Endpoint: Edit Rombongan (Grup)
app.post("/api/edit-group", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, oldGroupName, newGroupName } = req.body;
  const authHeader = req.headers.authorization;

  if (!newGroupName || newGroupName.trim() === "") {
    return res.status(400).json({ success: false, error: "Nama grup baru tidak boleh kosong." });
  }

  if (mode === "demo") {
    const gIdx = mockGroups.indexOf(oldGroupName);
    if (gIdx !== -1) mockGroups[gIdx] = newGroupName;
    
    mockManifest.forEach(p => {
      if (p.grup === oldGroupName) p.grup = newGroupName;
    });

    mockLogs.forEach(l => {
      if (l.grup === oldGroupName) l.grup = newGroupName;
    });

    return res.json({ success: true });
  }

  if (mode === "gas") {
    if (!gasUrl) return res.status(400).json({ success: false, error: "GAS URL tidak boleh kosong." });
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "editGrup", oldGroupName, newGroupName })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal edit grup di GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    if (!authHeader) return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });

    try {
      // Edit Manifest
      const manifestRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A:D`,
        { headers: { Authorization: authHeader } }
      );
      if (manifestRes.ok) {
        const manifestData = await manifestRes.json();
        const rows = manifestData.values || [];
        if (rows.length > 1) {
          for (let i = 1; i < rows.length; i++) {
            if (rows[i][0] === oldGroupName) {
              rows[i][0] = newGroupName;
            }
          }
          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A:D?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers: { Authorization: authHeader, "Content-Type": "application/json" },
              body: JSON.stringify({ values: rows })
            }
          );
        }
      }

      // Edit Log
      const logsRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I`,
        { headers: { Authorization: authHeader } }
      );
      if (logsRes.ok) {
        const logsData = await logsRes.json();
        const rows = logsData.values || [];
        if (rows.length > 1) {
          for (let i = 1; i < rows.length; i++) {
            if (rows[i][1] === oldGroupName) {
              rows[i][1] = newGroupName;
            }
          }
          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers: { Authorization: authHeader, "Content-Type": "application/json" },
              body: JSON.stringify({ values: rows })
            }
          );
        }
      }

      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal edit grup di Sheets: " + err.message });
    }
  }
  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// 7. Endpoint: Hapus Rombongan (Grup)
app.post("/api/delete-group", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, groupName } = req.body;
  const authHeader = req.headers.authorization;

  if (mode === "demo") {
    mockGroups = mockGroups.filter(g => g !== groupName);
    mockManifest = mockManifest.filter(p => p.grup !== groupName);
    mockLogs = mockLogs.filter(l => l.grup !== groupName);
    return res.json({ success: true });
  }

  if (mode === "gas") {
    if (!gasUrl) return res.status(400).json({ success: false, error: "GAS URL tidak boleh kosong." });
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "hapusGrup", groupName })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal hapus grup di GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    if (!authHeader) return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });

    try {
      // Filter Manifest Sheet
      const manifestRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A:D`,
        { headers: { Authorization: authHeader } }
      );
      if (manifestRes.ok) {
        const manifestData = await manifestRes.json();
        const rows = manifestData.values || [];
        if (rows.length > 0) {
          const header = rows[0];
          const filteredRows = [header, ...rows.slice(1).filter((r: any) => r[0] !== groupName)];
          
          await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A:D:clear`, {
            method: "POST",
            headers: { Authorization: authHeader }
          });
          
          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Manifest!A1?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers: { Authorization: authHeader, "Content-Type": "application/json" },
              body: JSON.stringify({ values: filteredRows })
            }
          );
        }
      }

      // Filter Log Sheet
      const logsRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I`,
        { headers: { Authorization: authHeader } }
      );
      if (logsRes.ok) {
        const logsData = await logsRes.json();
        const rows = logsData.values || [];
        if (rows.length > 0) {
          const header = rows[0];
          const filteredRows = [header, ...rows.slice(1).filter((r: any) => r[1] !== groupName)];
          
          await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I:clear`, {
            method: "POST",
            headers: { Authorization: authHeader }
          });

          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A1?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers: { Authorization: authHeader, "Content-Type": "application/json" },
              body: JSON.stringify({ values: filteredRows })
            }
          );
        }
      }

      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal hapus grup di Sheets: " + err.message });
    }
  }
  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// 8. Endpoint: Edit Tempat Checklist (Station)
app.post("/api/edit-station", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, oldStationName, newStationName } = req.body;
  const authHeader = req.headers.authorization;

  if (!newStationName || newStationName.trim() === "") {
    return res.status(400).json({ success: false, error: "Nama tempat Checklist baru tidak boleh kosong." });
  }

  if (mode === "demo") {
    const sIdx = mockCheckpointStations.indexOf(oldStationName);
    if (sIdx !== -1) mockCheckpointStations[sIdx] = newStationName;

    mockLogs.forEach(l => {
      if (l.pos === oldStationName) l.pos = newStationName;
    });

    return res.json({ success: true });
  }

  if (mode === "gas") {
    if (!gasUrl) return res.status(400).json({ success: false, error: "GAS URL tidak boleh kosong." });
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "editPos", oldStationName, newStationName })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal edit pos di GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    if (!authHeader) return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });

    try {
      // Edit Pos/Cheklist List
      let sheetName = "Cheklist";
      let posRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Cheklist!A:A`,
        { headers: { Authorization: authHeader } }
      );
      if (!posRes.ok) {
        sheetName = "Pos";
        posRes = await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Pos!A:A`,
          { headers: { Authorization: authHeader } }
        );
      }
      if (posRes.ok) {
        const posData = await posRes.json();
        const rows = posData.values || [];
        if (rows.length > 1) {
          for (let i = 1; i < rows.length; i++) {
            if (rows[i][0] === oldStationName) {
              rows[i][0] = newStationName;
            }
          }
          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${sheetName}!A:A?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers: { Authorization: authHeader, "Content-Type": "application/json" },
              body: JSON.stringify({ values: rows })
            }
          );
        }
      }

      // Edit Log References
      const logsRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I`,
        { headers: { Authorization: authHeader } }
      );
      if (logsRes.ok) {
        const logsData = await logsRes.json();
        const rows = logsData.values || [];
        if (rows.length > 1) {
          for (let i = 1; i < rows.length; i++) {
            if (rows[i][4] === oldStationName) {
              rows[i][4] = newStationName;
            }
          }
          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers: { Authorization: authHeader, "Content-Type": "application/json" },
              body: JSON.stringify({ values: rows })
            }
          );
        }
      }

      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal edit pos di Sheets: " + err.message });
    }
  }
  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// 9. Endpoint: Hapus Tempat Checklist (Station)
app.post("/api/delete-station", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, stationName } = req.body;
  const authHeader = req.headers.authorization;

  if (mode === "demo") {
    mockCheckpointStations = mockCheckpointStations.filter(s => s !== stationName);
    mockLogs = mockLogs.filter(l => l.pos !== stationName);
    return res.json({ success: true });
  }

  if (mode === "gas") {
    if (!gasUrl) return res.status(400).json({ success: false, error: "GAS URL tidak boleh kosong." });
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "hapusPos", stationName })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal hapus pos di GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    if (!authHeader) return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });

    try {
      // Filter Pos/Cheklist
      let sheetName = "Cheklist";
      let posRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Cheklist!A:A`,
        { headers: { Authorization: authHeader } }
      );
      if (!posRes.ok) {
        sheetName = "Pos";
        posRes = await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Pos!A:A`,
          { headers: { Authorization: authHeader } }
        );
      }
      if (posRes.ok) {
        const posData = await posRes.json();
        const rows = posData.values || [];
        if (rows.length > 0) {
          const header = rows[0];
          const filteredRows = [header, ...rows.slice(1).filter((r: any) => r[0] !== stationName)];
          
          await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${sheetName}!A:A:clear`, {
            method: "POST",
            headers: { Authorization: authHeader }
          });

          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${sheetName}!A1?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers: { Authorization: authHeader, "Content-Type": "application/json" },
              body: JSON.stringify({ values: filteredRows })
            }
          );
        }
      }

      // Filter Logs that reference this station
      const logsRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I`,
        { headers: { Authorization: authHeader } }
      );
      if (logsRes.ok) {
        const logsData = await logsRes.json();
        const rows = logsData.values || [];
        if (rows.length > 0) {
          const header = rows[0];
          const filteredRows = [header, ...rows.slice(1).filter((r: any) => r[4] !== stationName)];
          
          await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I:clear`, {
            method: "POST",
            headers: { Authorization: authHeader }
          });

          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A1?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers: { Authorization: authHeader, "Content-Type": "application/json" },
              body: JSON.stringify({ values: filteredRows })
            }
          );
        }
      }

      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal hapus pos di Sheets: " + err.message });
    }
  }
  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// 10. Endpoint: Edit Log Checklist (Data Tercatat)
app.post("/api/edit-log", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, originalLog, updatedLog } = req.body;
  const authHeader = req.headers.authorization;

  if (mode === "demo") {
    const lIdx = mockLogs.findIndex(l => 
      l.timestamp === originalLog.timestamp &&
      l.grup === originalLog.grup &&
      l.paspor === originalLog.paspor &&
      l.pos === originalLog.pos
    );
    if (lIdx !== -1) {
      mockLogs[lIdx] = {
        ...mockLogs[lIdx],
        koperBesar: updatedLog.koperBesar,
        koperKecil: updatedLog.koperKecil,
        status: updatedLog.status,
        petugas: updatedLog.petugas || "Demo User"
      };
    }
    return res.json({ success: true, logs: mockLogs });
  }

  if (mode === "gas") {
    if (!gasUrl) return res.status(400).json({ success: false, error: "GAS URL tidak boleh kosong." });
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "editLogChecklist", originalLog, updatedLog })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal edit log di GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    if (!authHeader) return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });

    try {
      const logsRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I`,
        { headers: { Authorization: authHeader } }
      );
      if (logsRes.ok) {
        const logsData = await logsRes.json();
        const rows = logsData.values || [];
        let updatedRowIndex = -1;
        for (let i = 1; i < rows.length; i++) {
          if (
            rows[i][0] === originalLog.timestamp &&
            rows[i][1] === originalLog.grup &&
            rows[i][3] === originalLog.paspor &&
            rows[i][4] === originalLog.pos
          ) {
            updatedRowIndex = i + 1; // 1-based index in sheets
            break;
          }
        }

        if (updatedRowIndex !== -1) {
          const r = updatedRowIndex;
          const updatedRowValues = [
            originalLog.timestamp,
            originalLog.grup,
            originalLog.nama,
            originalLog.paspor,
            originalLog.pos,
            updatedLog.koperBesar,
            updatedLog.koperKecil,
            updatedLog.petugas || "Petugas Lapangan",
            updatedLog.status
          ];

          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A${r}:I${r}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers: { Authorization: authHeader, "Content-Type": "application/json" },
              body: JSON.stringify({ values: [updatedRowValues] })
            }
          );
        }
      }
      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal edit log di Sheets: " + err.message });
    }
  }
  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// 11. Endpoint: Hapus Log Checklist (Data Tercatat)
app.post("/api/delete-log", async (req, res) => {
  const { mode, spreadsheetId, gasUrl, timestamp, grup, paspor, pos } = req.body;
  const authHeader = req.headers.authorization;

  if (mode === "demo") {
    mockLogs = mockLogs.filter(l => 
      !(l.timestamp === timestamp &&
        l.grup === grup &&
        l.paspor === paspor &&
        l.pos === pos)
    );
    return res.json({ success: true, logs: mockLogs });
  }

  if (mode === "gas") {
    if (!gasUrl) return res.status(400).json({ success: false, error: "GAS URL tidak boleh kosong." });
    try {
      const data = await fetchGasSafe(gasUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "hapusLogChecklist", timestamp, grup, paspor, pos })
      });
      return res.json({ success: true, ...data });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal hapus log di GAS: " + err.message });
    }
  }

  if (mode === "oauth") {
    if (!spreadsheetId) return res.status(400).json({ success: false, error: "Spreadsheet ID wajib diisi." });
    if (!authHeader) return res.status(401).json({ success: false, error: "Token otorisasi diperlukan." });

    try {
      const logsRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I`,
        { headers: { Authorization: authHeader } }
      );
      if (logsRes.ok) {
        const logsData = await logsRes.json();
        const rows = logsData.values || [];
        if (rows.length > 0) {
          const header = rows[0];
          const filteredRows = [
            header,
            ...rows.slice(1).filter((r: any) => 
              !(r[0] === timestamp &&
                r[1] === grup &&
                r[3] === paspor &&
                r[4] === pos)
            )
          ];
          
          await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A:I:clear`, {
            method: "POST",
            headers: { Authorization: authHeader }
          });

          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Log!A1?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers: { Authorization: authHeader, "Content-Type": "application/json" },
              body: JSON.stringify({ values: filteredRows })
            }
          );
        }
      }
      return res.json({ success: true });
    } catch (err: any) {
      return res.status(500).json({ success: false, error: "Gagal hapus log di Sheets: " + err.message });
    }
  }
  return res.status(400).json({ success: false, error: "Mode tidak valid." });
});

// Vite Middleware for development, or serving compiled React build in production
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Express server running on http://0.0.0.0:${PORT}`);
  });
}

if (!process.env.VERCEL) {
  startServer();
}

export default app;
