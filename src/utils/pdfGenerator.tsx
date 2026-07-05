import React from "react";
import { Document, Page, Text, View, StyleSheet, pdf } from "@react-pdf/renderer";
import { LogEntry, Pilgrim } from "../types";

// Helper to parse luggage count
function parseLuggageCount(val: string | number | undefined): number {
  if (val === undefined || val === null) return 0;
  if (typeof val === "number") return val;
  const valStr = String(val);
  const lower = valStr.toLowerCase();
  if (lower.includes("tidak ada") || lower.includes("tidak") || lower.includes("0")) return 0;
  const match = valStr.match(/\d+/);
  if (match) return parseInt(match[0], 10);
  if (lower.includes("ada")) return 1;
  return 0;
}

// Design Styles for react-pdf/renderer (React Native CSS subset)
const styles = StyleSheet.create({
  page: {
    padding: 30,
    backgroundColor: "#ffffff",
    fontFamily: "Helvetica",
    fontSize: 9,
    color: "#334155",
  },
  header: {
    backgroundColor: "#0f172a", // Deep Slate 900
    padding: 15,
    borderRadius: 6,
    marginBottom: 15,
  },
  headerTitle: {
    color: "#ffffff",
    fontSize: 14,
    fontWeight: "bold",
    letterSpacing: 0.5,
    marginBottom: 2,
  },
  headerSubtitle: {
    color: "#94a3b8", // Slate 400
    fontSize: 7.5,
    textTransform: "uppercase",
    marginBottom: 6,
    fontWeight: "bold",
  },
  headerDate: {
    color: "#cbd5e1", // Slate 300
    fontSize: 8,
  },
  accentLine: {
    height: 3,
    backgroundColor: "#10b981", // Emerald 500
    marginTop: 8,
    borderRadius: 1.5,
  },
  infoSection: {
    flexDirection: "row",
    justifyContent: "space-between",
    marginBottom: 15,
  },
  metaCol: {
    width: "55%",
    flexDirection: "column",
  },
  metaTitle: {
    fontSize: 9,
    fontWeight: "bold",
    color: "#0f172a",
    borderBottomWidth: 1,
    borderBottomColor: "#e2e8f0",
    borderBottomStyle: "solid",
    paddingBottom: 4,
    marginBottom: 6,
    textTransform: "uppercase",
  },
  metaRow: {
    flexDirection: "row",
    marginBottom: 4,
    alignItems: "center",
  },
  metaLabel: {
    width: 110,
    fontSize: 8,
    color: "#64748b", // Slate 500
  },
  metaValue: {
    fontSize: 8.5,
    fontWeight: "bold",
    color: "#0f172a",
  },
  rekapCard: {
    width: "40%",
    backgroundColor: "#f8fafc", // Soft Slate 50
    borderWidth: 1,
    borderColor: "#e2e8f0",
    borderStyle: "solid",
    borderRadius: 6,
    padding: 8,
  },
  rekapTitle: {
    fontSize: 8,
    fontWeight: "bold",
    color: "#2563eb", // Blue 600
    marginBottom: 5,
    textTransform: "uppercase",
    borderBottomWidth: 1,
    borderBottomColor: "#e2e8f0",
    borderBottomStyle: "solid",
    paddingBottom: 2,
  },
  rekapRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    marginBottom: 3,
  },
  rekapLabel: {
    fontSize: 7.5,
    color: "#475569",
  },
  rekapValue: {
    fontSize: 8,
    fontWeight: "bold",
    color: "#0f172a",
  },
  tableHeader: {
    flexDirection: "row",
    backgroundColor: "#1e293b", // Slate 800
    borderRadius: 3,
    padding: 6,
    alignItems: "center",
    marginTop: 5,
  },
  tableHeaderColNo: {
    width: "6%",
    color: "#ffffff",
    fontSize: 7.5,
    fontWeight: "bold",
  },
  tableHeaderColName: {
    width: "36%",
    color: "#ffffff",
    fontSize: 7.5,
    fontWeight: "bold",
  },
  tableHeaderColPaspor: {
    width: "16%",
    color: "#ffffff",
    fontSize: 7.5,
    fontWeight: "bold",
  },
  tableHeaderColKoperB: {
    width: "13%",
    color: "#ffffff",
    fontSize: 7.5,
    fontWeight: "bold",
    textAlign: "center",
  },
  tableHeaderColKoperK: {
    width: "13%",
    color: "#ffffff",
    fontSize: 7.5,
    fontWeight: "bold",
    textAlign: "center",
  },
  tableHeaderColStatus: {
    width: "16%",
    color: "#ffffff",
    fontSize: 7.5,
    fontWeight: "bold",
  },
  tableRow: {
    flexDirection: "row",
    borderBottomWidth: 1,
    borderBottomColor: "#f1f5f9",
    borderBottomStyle: "solid",
    padding: "5 6",
    alignItems: "center",
  },
  tableRowAlt: {
    flexDirection: "row",
    borderBottomWidth: 1,
    borderBottomColor: "#f1f5f9",
    borderBottomStyle: "solid",
    padding: "5 6",
    alignItems: "center",
    backgroundColor: "#f8fafc",
  },
  colNo: {
    width: "6%",
    fontSize: 7.5,
    color: "#334155",
  },
  colName: {
    width: "36%",
    fontSize: 7.5,
    fontWeight: "bold",
    color: "#0f172a",
  },
  colPaspor: {
    width: "16%",
    fontSize: 7.5,
    color: "#475569",
  },
  colKoperB: {
    width: "13%",
    fontSize: 7.5,
    color: "#334155",
    textAlign: "center",
  },
  colKoperK: {
    width: "13%",
    fontSize: 7.5,
    color: "#334155",
    textAlign: "center",
  },
  colStatus: {
    width: "16%",
    fontSize: 7.5,
    color: "#475569",
  },
  emptyTableContainer: {
    padding: 15,
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: "#f8fafc",
    borderRadius: 4,
    borderWidth: 1,
    borderColor: "#e2e8f0",
    borderStyle: "dashed",
    marginTop: 5,
  },
  emptyTableText: {
    fontSize: 8,
    color: "#64748b",
    fontStyle: "italic",
  },
  summaryBox: {
    backgroundColor: "#fef3c7", // Amber 100
    borderWidth: 1,
    borderColor: "#d97706", // Amber 600
    borderStyle: "solid",
    borderRadius: 6,
    padding: 10,
    marginTop: 15,
    flexDirection: "row",
  },
  summaryAccent: {
    width: 4,
    backgroundColor: "#d97706",
    borderRadius: 2,
    marginRight: 8,
  },
  summaryContent: {
    flex: 1,
  },
  summaryTitle: {
    fontSize: 9,
    fontWeight: "bold",
    color: "#0f172a",
    marginBottom: 4,
    textTransform: "uppercase",
  },
  summaryDivider: {
    height: 1,
    backgroundColor: "#fbe183",
    marginBottom: 6,
  },
  summaryGrid: {
    flexDirection: "row",
    flexWrap: "wrap",
  },
  summaryCol: {
    width: "50%",
    marginBottom: 3,
    flexDirection: "row",
    alignItems: "center",
  },
  summaryBullet: {
    fontSize: 10,
    color: "#d97706",
    marginRight: 4,
    fontWeight: "bold",
  },
  summaryText: {
    fontSize: 7.5,
    color: "#1e293b",
  },
  summaryTextHighlight: {
    fontSize: 7.5,
    fontWeight: "bold",
    color: "#b45309",
  },
  signatureSection: {
    marginTop: 20,
    flexDirection: "row",
    justifyContent: "flex-end",
    paddingRight: 10,
  },
  signatureBlock: {
    width: 180,
    alignItems: "center",
  },
  signatureLabel: {
    fontSize: 8,
    color: "#334155",
    marginBottom: 35,
  },
  signatureLine: {
    borderBottomWidth: 1,
    borderBottomColor: "#94a3b8",
    borderBottomStyle: "solid",
    width: "100%",
    marginBottom: 3,
  },
  signatureSub: {
    fontSize: 7,
    color: "#64748b",
    fontStyle: "italic",
  },
});

// React PDF Document Component
interface DocumentProps {
  groupName: string;
  stationName: string;
  logs: LogEntry[];
  manifest: Pilgrim[];
}

export const PDFReportDocument: React.FC<DocumentProps> = ({
  groupName,
  stationName,
  logs,
  manifest,
}) => {
  // Filter logs for this specific group and station
  const filteredLogs = logs.filter(
    (l) => l.grup === groupName && (l.pos === stationName || stationName === "Semua")
  );

  // Group logs by bus using manifest lookup
  const logsByBus: Record<string, LogEntry[]> = {};
  filteredLogs.forEach((l) => {
    const pilgrim = manifest.find((p) => 
      p.nama.trim().toLowerCase() === l.nama.trim().toLowerCase() && 
      p.paspor.trim().toLowerCase() === l.paspor.trim().toLowerCase()
    );
    const busName = pilgrim && pilgrim.bus ? pilgrim.bus.trim() : "Tanpa Bus";
    if (!logsByBus[busName]) {
      logsByBus[busName] = [];
    }
    logsByBus[busName].push(l);
  });

  const busNames = Object.keys(logsByBus).sort();
  const dateStr = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" }) + " WIB";

  // If there are no logs, create at least one page for "Semua Bus" showing an empty state
  const pagesData = busNames.length === 0 ? [{ name: "Semua Bus", list: [] }] : busNames.map(name => ({
    name,
    list: logsByBus[name]
  }));

  return (
    <Document>
      {pagesData.map((bus, busIdx) => {
        // Calculate totals for this bus
        let totalKB = 0;
        let totalKK = 0;
        bus.list.forEach((l) => {
          totalKB += parseLuggageCount(l.koperBesar);
          totalKK += parseLuggageCount(l.koperKecil);
        });

        return (
          <Page key={busIdx} size="A4" style={styles.page}>
            {/* 1. Header Banner */}
            <View style={styles.header}>
              <Text style={styles.headerTitle}>LAPORAN DATA CHECKLIST KOPER JEMAAH</Text>
              <Text style={styles.headerSubtitle}>SISTEM MANAJEMEN MANIFEST DAN KOPER ROMBONGAN REAL-TIME</Text>
              <Text style={styles.headerDate}>Waktu Cetak Dokumen: {dateStr}</Text>
              <View style={styles.accentLine} />
            </View>

            {/* 2. Metadata Section */}
            <View style={styles.infoSection}>
              {/* Left Details */}
              <View style={styles.metaCol}>
                <Text style={styles.metaTitle}>INFORMASI KEGIATAN & OPERASIONAL</Text>
                <View style={styles.metaRow}>
                  <Text style={styles.metaLabel}>Nama Rombongan / Grup:</Text>
                  <Text style={styles.metaValue}>{groupName}</Text>
                </View>
                <View style={styles.metaRow}>
                  <Text style={styles.metaLabel}>Tempat Checklist:</Text>
                  <Text style={styles.metaValue}>{stationName}</Text>
                </View>
                <View style={styles.metaRow}>
                  <Text style={styles.metaLabel}>Rombongan Armada:</Text>
                  <Text style={styles.metaValue}>{bus.name}</Text>
                </View>
              </View>

              {/* Right Recapitulation Card */}
              <View style={styles.rekapCard}>
                <Text style={styles.rekapTitle}>REKAPITULASI {bus.name.toUpperCase()}</Text>
                <View style={styles.rekapRow}>
                  <Text style={styles.rekapLabel}>Jemaah Terperiksa:</Text>
                  <Text style={styles.rekapValue}>{bus.list.length} Orang</Text>
                </View>
                <View style={styles.rekapRow}>
                  <Text style={styles.rekapLabel}>Total Koper Besar:</Text>
                  <Text style={styles.rekapValue}>{totalKB} Pcs</Text>
                </View>
                <View style={styles.rekapRow}>
                  <Text style={styles.rekapLabel}>Total Koper Kecil:</Text>
                  <Text style={styles.rekapValue}>{totalKK} Pcs</Text>
                </View>
              </View>
            </View>

            {/* 3. Table Header */}
            <View style={styles.tableHeader}>
              <Text style={styles.tableHeaderColNo}>NO</Text>
              <Text style={styles.tableHeaderColName}>NAMA JEMAAH</Text>
              <Text style={styles.tableHeaderColPaspor}>NO. PASPOR</Text>
              <Text style={styles.tableHeaderColKoperB}>KOPER BESAR</Text>
              <Text style={styles.tableHeaderColKoperK}>KOPER KECIL</Text>
              <Text style={styles.tableHeaderColStatus}>STATUS & CATATAN</Text>
            </View>

            {/* 4. Table Rows */}
            {bus.list.length === 0 ? (
              <View style={styles.emptyTableContainer}>
                <Text style={styles.emptyTableText}>Belum ada data checklist koper untuk armada ini.</Text>
              </View>
            ) : (
              bus.list.map((l, idx) => {
                const rowStyle = idx % 2 === 1 ? styles.tableRowAlt : styles.tableRow;
                const displayName = l.nama.length > 30 ? l.nama.substring(0, 28) + "..." : l.nama;
                const displayStatus = l.status || "Lengkap";
                const truncatedStatus = displayStatus.length > 22 ? displayStatus.substring(0, 20) + "..." : displayStatus;

                return (
                  <View key={idx} style={rowStyle}>
                    <Text style={styles.colNo}>{idx + 1}</Text>
                    <Text style={styles.colName}>{displayName}</Text>
                    <Text style={styles.colPaspor}>{l.paspor}</Text>
                    <Text style={styles.colKoperB}>{l.koperBesar || "Tidak Ada"}</Text>
                    <Text style={styles.colKoperK}>{l.koperKecil || "Tidak Ada"}</Text>
                    <Text style={styles.colStatus}>{truncatedStatus}</Text>
                  </View>
                );
              })
            )}

            {/* 5. Final Summary Box */}
            <View style={styles.summaryBox}>
              <View style={styles.summaryAccent} />
              <View style={styles.summaryContent}>
                <Text style={styles.summaryTitle}>RINGKASAN AKHIR VERIFIKASI BAGASI ({bus.name.toUpperCase()})</Text>
                <View style={styles.summaryDivider} />
                <View style={styles.summaryGrid}>
                  <View style={styles.summaryCol}>
                    <Text style={styles.summaryBullet}>•</Text>
                    <Text style={styles.summaryText}>
                      Jumlah Jemaah Terverifikasi: <Text style={styles.summaryTextHighlight}>{bus.list.length} Orang</Text>
                    </Text>
                  </View>
                  <View style={styles.summaryCol}>
                    <Text style={styles.summaryBullet}>•</Text>
                    <Text style={styles.summaryText}>
                      Total Akumulasi Koper Besar: <Text style={styles.summaryTextHighlight}>{totalKB} Koper</Text>
                    </Text>
                  </View>
                  <View style={styles.summaryCol}>
                    <Text style={styles.summaryBullet}>•</Text>
                    <Text style={styles.summaryText}>
                      Total Akumulasi Koper Kecil: <Text style={styles.summaryTextHighlight}>{totalKK} Koper</Text>
                    </Text>
                  </View>
                  <View style={styles.summaryCol}>
                    <Text style={styles.summaryBullet}>•</Text>
                    <Text style={styles.summaryText}>
                      Total Seluruh Barang Bawaan: <Text style={styles.summaryTextHighlight}>{totalKB + totalKK} Bagasi</Text>
                    </Text>
                  </View>
                </View>
              </View>
            </View>

            {/* 6. Signatures Section */}
            <View style={styles.signatureSection}>
              <View style={styles.signatureBlock}>
                <Text style={styles.signatureLabel}>Petugas Pemeriksa {bus.name},</Text>
                <View style={styles.signatureLine} />
                <Text style={styles.signatureSub}>Tanda Tangan & Nama Terang</Text>
              </View>
            </View>
          </Page>
        );
      })}
    </Document>
  );
};

// Orchestrator Function to build and save the PDF document
export async function generatePdfReport(
  groupName: string,
  stationName: string,
  logs: LogEntry[],
  manifest: Pilgrim[]
) {
  try {
    const docElement = (
      <PDFReportDocument
        groupName={groupName}
        stationName={stationName}
        logs={logs}
        manifest={manifest}
      />
    );
    
    // Generate blob on-the-fly using react-pdf/renderer's pdf function
    const blob = await pdf(docElement).toBlob();
    
    const safeGroup = groupName.replace(/[^a-z0-9]/gi, "_").toLowerCase();
    const safeStation = stationName.replace(/[^a-z0-9]/gi, "_").toLowerCase();
    const filename = `Laporan_Koper_${safeGroup}_${safeStation}.pdf`;

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  } catch (error) {
    console.error("Gagal membuat laporan PDF dengan react-pdf:", error);
  }
}
