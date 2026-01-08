// =====================================================
// POP UP CSV GENERATOR
// FINAL – ISI SAMA, URUTAN BEBAS
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

// ===== TEMPLATE HEADERS (PERSIS CSV SEHARUSNYA) =====
const HOME_HEADERS = [
  "HOME_PASS_ID","ADDRESS","RT","RW","KELURAHAN","KECAMATAN","KABUPATEN",
  "PROVINSI","POSTAL_CODE","CLUSTER_NAME","ID_Area","LATITUDE","LONGITUDE",
  "FDT_CODE","FAT ID/NETWORK ID","Pole ID (New)","Clamp_Hook_ID",
  "HOME/HOME-BIZ"
];
const FDT_HEADERS = ["FDT_CODE","CLUSTER_NAME","ID_Area"];
const FAT_HEADERS = ["FAT_CODE","FDT_CODE","CLUSTER_NAME","ID_Area"];
const POLE_HEADERS = [
  "Pole ID (New)","Coordinate (Lat) NEW","Coordinate (Long) NEW",
  "Pole Provider (New)","Pole Type","LINE","CLUSTER_NAME","ID_Area"
];

// ===== CSV BUILDER =====
function toCSV(headers, rows) {
  const out = [];
  out.push(headers.map(esc).join(SEP));
  rows.forEach(r => {
    out.push(headers.map(h => esc(r[h] || "")).join(SEP));
  });
  return out.join("\n");
}

// ===== MAIN =====
btn.addEventListener("click", async () => {
  const file = fileInput.files[0];
  if (!file) return;

  statusEl.textContent = "Memproses Master Excel...";
  btn.disabled = true;

  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const sheet =
      wb.SheetNames.find(s => s === "Master Data") ||
      wb.SheetNames[0];

    const master = XLSX.utils.sheet_to_json(
      wb.Sheets[sheet],
      { defval: "" }
    );

    if (!master.length) throw new Error("Master kosong");

    const area =
      $("areaName").value ||
      master[0]["ID_Area"] ||
      file.name.replace(/\.(xlsx|xls)$/i, "");

    // ================= HOME / HOME-BIZ =================
    const HOME = [];
    const HOME_BIZ = [];

    master.forEach(r => {
      const row = Object.fromEntries(
        HOME_HEADERS.map(h => [h, r[h] || ""])
      );
      if (r["HOME/HOME-BIZ"] === "HOME") HOME.push(row);
      if (r["HOME/HOME-BIZ"] === "HOME-BIZ") HOME_BIZ.push(row);
    });

    // ================= FDT =================
    const fdtMap = {};
    master.forEach(r => {
      if (!r["FDT_CODE"]) return;
      const key = r["FDT_CODE"];
      if (!fdtMap[key]) {
        fdtMap[key] = {
          "FDT_CODE": r["FDT_CODE"],
          "CLUSTER_NAME": r["CLUSTER_NAME"] || "",
          "ID_Area": r["ID_Area"] || ""
        };
      }
    });
    const FDT = Object.values(fdtMap);

    // ================= FAT =================
    const FAT = [];
    master.forEach(r => {
      const raw = r["FAT ID/NETWORK ID"];
      if (!raw) return;

      raw.split("&").map(x => x.trim()).filter(Boolean).forEach(fat => {
        FAT.push({
          "FAT_CODE": fat,
          "FDT_CODE": r["FDT_CODE"] || "",
          "CLUSTER_NAME": r["CLUSTER_NAME"] || "",
          "ID_Area": r["ID_Area"] || ""
        });
      });
    });

    // ================= POLE =================
    const poleMap = {};
    master.forEach(r => {
      if (!r["Pole ID (New)"]) return;
      const key = r["Pole ID (New)"] + "|" + r["CLUSTER_NAME"] + "|" + r["ID_Area"];
      if (!poleMap[key]) {
        poleMap[key] = {
          "Pole ID (New)": r["Pole ID (New)"],
          "Coordinate (Lat) NEW": r["Coordinate (Lat) NEW"] || "",
          "Coordinate (Long) NEW": r["Coordinate (Long) NEW"] || "",
          "Pole Provider (New)": r["Pole Provider (New)"] || "",
          "Pole Type": r["Pole Type"] || "",
          "LINE": r["LINE"] || "",
          "CLUSTER_NAME": r["CLUSTER_NAME"] || "",
          "ID_Area": r["ID_Area"] || ""
        };
      }
    });
    const POLE = Object.values(poleMap);

    // ================= ZIP =================
    const zip = new JSZip();
    const folder = zip.folder(area);

    folder.file("HOME.csv", toCSV(HOME_HEADERS, HOME));
    folder.file("HOME-BIZ.csv", toCSV(HOME_HEADERS, HOME_BIZ));
    folder.file("FDT.csv", toCSV(FDT_HEADERS, FDT));
    folder.file("FAT.csv", toCSV(FAT_HEADERS, FAT));
    folder.file("POLE.csv", toCSV(POLE_HEADERS, POLE));

    const blob = await zip.generateAsync({ type: "blob" });
    saveAs(blob, `${area}_popups.zip`);

    statusEl.textContent = "SELESAI ✔ ISI SESUAI CSV SEHARUSNYA";
  } catch (e) {
    console.error(e);
    statusEl.textContent = "ERROR: " + e.message;
  } finally {
    btn.disabled = false;
  }
});
