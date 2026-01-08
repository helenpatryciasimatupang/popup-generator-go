// =======================================================
// POP UP CSV GENERATOR
// FINAL – HARDCODE TEMPLATE (IKUT CSV CONTOH)
// =======================================================

const $ = (id) => document.getElementById(id);

const fileInput = $("file");
const btn = $("btn");
const statusEl = $("status");

const cHome = $("cHome");
const cHomeBiz = $("cHomeBiz");
const cFdt = $("cFdt");
const cFat = $("cFat");
const cHook = $("cHook");
const cPole = $("cPole");

fileInput.addEventListener("change", () => {
  btn.disabled = !fileInput.files.length;
});

const SEP = ";";
const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;

// =====================
// TEMPLATE HEADERS (FIX)
// =====================
const HOME_HEADERS = [
  "HOME_PASS_ID","ADDRESS","RT","RW","KELURAHAN","KECAMATAN","KABUPATEN",
  "PROVINSI","POSTAL_CODE","CLUSTER_NAME","ID_Area","LATITUDE","LONGITUDE",
  "FDT_CODE","FAT ID/NETWORK ID","Pole ID (New)","Clamp_Hook_ID",
  "HOME/HOME-BIZ"
];

const FDT_HEADERS = ["FDT_CODE","CLUSTER_NAME","ID_Area"];
const FAT_HEADERS = ["FAT_CODE","FDT_CODE","CLUSTER_NAME","ID_Area"];
const HOOK_HEADERS = [
  "Clamp_Hook_ID","Clamp_Hook_LATITUDE","Clamp_Hook_LONGITUDE",
  "CLUSTER_NAME","ID_Area"
];
const POLE_HEADERS = [
  "Pole ID (New)","Coordinate (Lat) NEW","Coordinate (Long) NEW",
  "Pole Provider (New)","Pole Type","LINE","CLUSTER_NAME","ID_Area"
];

// =====================
function toCSV(headers, rows) {
  const out = [];
  out.push(headers.map(esc).join(SEP));
  rows.forEach(r => {
    out.push(headers.map(h => esc(r[h])).join(SEP));
  });
  return out.join("\n");
}

function uniqBy(rows, key) {
  const map = {};
  rows.forEach(r => {
    const v = r[key];
    if (v && v !== "-") map[v] = r;
  });
  return Object.values(map);
}

// =====================
btn.addEventListener("click", async () => {
  const file = fileInput.files[0];
  if (!file) return;

  statusEl.textContent = "Memproses Excel...";
  btn.disabled = true;

  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const sheet =
      wb.SheetNames.find(s => s === "Master Data") ||
      wb.SheetNames[0];

    const rows = XLSX.utils.sheet_to_json(
      wb.Sheets[sheet],
      { defval: "" }
    );

    if (!rows.length) throw new Error("Sheet kosong");

    const area =
      $("areaName").value ||
      rows[0]["ID_Area"] ||
      file.name.replace(/\.(xlsx|xls)$/i, "");

    // ================= HOME / HOME-BIZ =================
    const HOME = rows.filter(
      r => String(r["HOME/HOME-BIZ"]).trim() === "HOME"
    );
    const HOME_BIZ = rows.filter(
      r => String(r["HOME/HOME-BIZ"]).trim() === "HOME-BIZ"
    );

    // ================= FDT =================
    const FDT = uniqBy(rows, "FDT_CODE").map(r => ({
      "FDT_CODE": r["FDT_CODE"],
      "CLUSTER_NAME": r["CLUSTER_NAME"],
      "ID_Area": r["ID_Area"]
    }));

    // ================= FAT =================
    const fatMap = {};
    rows.forEach(r => {
      const raw = r["FAT ID/NETWORK ID"];
      if (!raw) return;
      raw.split(/[,&]/).forEach(x => {
        const code = x.trim();
        if (!code) return;
        fatMap[code] = {
          "FAT_CODE": code,
          "FDT_CODE": r["FDT_CODE"],
          "CLUSTER_NAME": r["CLUSTER_NAME"],
          "ID_Area": r["ID_Area"]
        };
      });
    });
    const FAT = Object.values(fatMap);

    // ================= HOOK =================
    const HOOK = uniqBy(rows, "Clamp_Hook_ID").map(r => ({
      "Clamp_Hook_ID": r["Clamp_Hook_ID"],
      "Clamp_Hook_LATITUDE": r["Clamp_Hook_LATITUDE"],
      "Clamp_Hook_LONGITUDE": r["Clamp_Hook_LONGITUDE"],
      "CLUSTER_NAME": r["CLUSTER_NAME"],
      "ID_Area": r["ID_Area"]
    }));

    // ================= POLE =================
    const POLE = uniqBy(rows, "Pole ID (New)").map(r => ({
      "Pole ID (New)": r["Pole ID (New)"],
      "Coordinate (Lat) NEW": r["Coordinate (Lat) NEW"],
      "Coordinate (Long) NEW": r["Coordinate (Long) NEW"],
      "Pole Provider (New)": r["Pole Provider (New)"],
      "Pole Type": r["Pole Type"],
      "LINE": r["LINE"],
      "CLUSTER_NAME": r["CLUSTER_NAME"],
      "ID_Area": r["ID_Area"]
    }));

    // ================= ZIP =================
    const zip = new JSZip();
    const folder = zip.folder(area);

    folder.file("HOME.csv", toCSV(HOME_HEADERS, HOME));
    folder.file("HOME-BIZ.csv", toCSV(HOME_HEADERS, HOME_BIZ));
    folder.file("FDT.csv", toCSV(FDT_HEADERS, FDT));
    folder.file("FAT.csv", toCSV(FAT_HEADERS, FAT));
    folder.file("HOOK.csv", toCSV(HOOK_HEADERS, HOOK));
    folder.file("POLE.csv", toCSV(POLE_HEADERS, POLE));

    cHome.textContent = HOME.length;
    cHomeBiz.textContent = HOME_BIZ.length;
    cFdt.textContent = FDT.length;
    cFat.textContent = FAT.length;
    cHook.textContent = HOOK.length;
    cPole.textContent = POLE.length;

    const blob = await zip.generateAsync({ type: "blob" });
    saveAs(blob, `${area}_popups.zip`);

    statusEl.textContent = "SELESAI ✔ CSV PERSIS TEMPLATE";
  } catch (e) {
    console.error(e);
    statusEl.textContent = "ERROR: " + e.message;
  } finally {
    btn.disabled = false;
  }
});
