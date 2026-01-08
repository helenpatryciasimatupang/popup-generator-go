// =====================================================
// POP UP CSV GENERATOR — FINAL TEMPLATE FIX (FULL)
// FIX UTAMA:
// - HOME vs HOME-BIZ ditentukan dari KOLOM TERAKHIR EXCEL (AKURAT)
// - Category BizPass TETAP nilai ASLI master (TIDAK diubah)
// - FDT = HANYA 1 BARIS: ambil data POLE "1A" dari Master (lengkap printilan)
// =====================================================

const $ = (id) => document.getElementById(id);

const fileInput = $("file");
const btn = $("btn");
const statusEl = $("status");

fileInput.addEventListener("change", () => {
  btn.disabled = !fileInput.files.length;
});

const SEP = ";";
const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;

// ================= TEMPLATE HEADER =================

// FAT & FDT (HEADER SAMA)
const FAT_FDT_HEADERS = [
  "Pole ID (New)",
  "Coordinate (Lat) NEW",
  "Coordinate (Long) NEW",
  "Pole Provider (New)",
  "Pole Type",
  "FAT ID/NETWORK ID",
];

// HOME & HOME-BIZ (HEADER SAMA)
const HOME_HEADERS = [
  "HOMEPASS_ID",
  "CLUSTER_NAME",
  "PREFIX_ADDRESS",
  "STREET_NAME",
  "HOUSE_NUMBER",
  "BLOCK",
  "FLOOR",
  "RT",
  "RW",
  "DISTRICT",
  "SUB_DISTRICT",
  "FDT_CODE",
  "FAT_CODE",
  "BUILDING_LATITUDE",
  "BUILDING_LONGITUDE",
  "Category BizPass",
  "POST CODE",
  "ADDRESS POLE / FAT",
  "OV_UG",
  "HOUSE_COMMENT_",
  "BUILDING_NAME",
  "TOWER",
  "APTN",
  "FIBER_NODE__HFC_",
  "ID_Area",
  "Clamp_Hook_ID",
  "DEPLOYMENT_TYPE",
  "NEED_SURVEY",
];

// POLE
const POLE_HEADERS = [
  "Pole ID (New)",
  "Coordinate (Lat) NEW",
  "Coordinate (Long) NEW",
  "Pole Provider (New)",
  "Pole Type",
  "LINE",
];

// ================= CSV BUILDER =================
function toCSV(headers, rows) {
  const out = [];
  out.push(headers.map(esc).join(SEP));
  rows.forEach((r) => {
    out.push(headers.map((h) => esc(r[h] || "")).join(SEP));
  });
  return out.join("\n");
}

// ================= MAIN =================
btn.addEventListener("click", async () => {
  const file = fileInput.files[0];
  if (!file) return;

  statusEl.textContent = "Memproses Master Excel...";
  btn.disabled = true;

  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const sheetName =
      wb.SheetNames.find((s) => s === "Master Data") || wb.SheetNames[0];

    const worksheet = wb.Sheets[sheetName];

    // ================== AMBIL HEADER ASLI EXCEL ==================
    const headerRow = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
    })[0];

    if (!headerRow || !headerRow.length) {
      throw new Error("Header Excel tidak terbaca");
    }

    // NAMA KOLOM TERAKHIR
    const LAST_COL_NAME = headerRow[headerRow.length - 1];

    // DATA MASTER
    const master = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    if (!master.length) throw new Error("Master kosong");

    const area =
      $("areaName").value ||
      master[0]["ID_Area"] ||
      file.name.replace(/\.(xlsx|xls)$/i, "");

    // ================= HOME / HOME-BIZ =================
    const HOME = [];
    const HOME_BIZ = [];

    master.forEach((r) => {
      const lastVal = String(r[LAST_COL_NAME] || "")
        .trim()
        .toUpperCase()
        .replace(/\s+/g, "");

      const isBiz =
        lastVal.includes("BIZ") ||
        lastVal.includes("BIS") ||
        lastVal.includes("BUSINESS");

      const row = Object.fromEntries(
        HOME_HEADERS.map((h) => [h, r[h] || ""])
      );

      // Category BizPass TETAP nilai ASLI
      row["Category BizPass"] = r["Category BizPass"] || "";

      if (isBiz) HOME_BIZ.push(row);
      else HOME.push(row);
    });

    // ================= FAT =================
    const FAT = [];
    master.forEach((r) => {
      FAT.push(
        Object.fromEntries(
          FAT_FDT_HEADERS.map((h) => [h, r[h] || ""])
        )
      );
    });

    // ================= FDT (FINAL: ambil data POLE 1A dari Master) =================
    const fdtSource = master.find(
      (r) =>
        String(r["Pole ID (New)"] || "")
          .trim()
          .toUpperCase() === "1A"
    );

    if (!fdtSource) {
      throw new Error('Data POLE "1A" tidak ditemukan di Master (Pole ID (New))');
    }

    const FDT = [
      Object.fromEntries(
        FAT_FDT_HEADERS.map((h) => [h, fdtSource[h] || ""])
      ),
    ];

    // ================= POLE =================
    const POLE = [];
    master.forEach((r) => {
      POLE.push(
        Object.fromEntries(
          POLE_HEADERS.map((h) => [h, r[h] || ""])
        )
      );
    });

    // ================= ZIP =================
    const zip = new JSZip();
    const folder = zip.folder(area);

    folder.file("HOME.csv", toCSV(HOME_HEADERS, HOME));
    folder.file("HOME-BIZ.csv", toCSV(HOME_HEADERS, HOME_BIZ));
    folder.file("FAT.csv", toCSV(FAT_FDT_HEADERS, FAT));
    folder.file("FDT.csv", toCSV(FAT_FDT_HEADERS, FDT));
    folder.file("POLE.csv", toCSV(POLE_HEADERS, POLE));

    const blob = await zip.generateAsync({ type: "blob" });
    saveAs(blob, `${area}_POPUP.zip`);

    statusEl.textContent =
      `SELESAI ✔ HOME=${HOME.length} | HOME-BIZ=${HOME_BIZ.length} | ` +
      `FAT=${FAT.length} | FDT=1 (Pole 1A) | POLE=${POLE.length}`;
  } catch (e) {
    console.error(e);
    statusEl.textContent = "ERROR: " + e.message;
  } finally {
    btn.disabled = false;
  }
});
