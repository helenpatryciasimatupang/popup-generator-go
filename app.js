// =====================================================
// POP UP CSV GENERATOR — FINAL TEMPLATE FIX (FULL)
// + KMZ PATCHER (CLIENT-SIDE)
//
// CSV GENERATOR:
// - HOME vs HOME-BIZ pakai KOLOM TERAKHIR
// - Category BizPass tetap ASLI
// - FDT = hanya POLE 1A
// - FAT = semua POLE kecuali 1A
//
// KMZ PATCHER:
// - Update popup dari CSV
// - Mapping BERDASARKAN FOLDER
// - ID diambil dari <Placemark><name>
// =====================================================

const $ = (id) => document.getElementById(id);

// ================= ELEMENT =================
const fileInput = $("file");
const btn = $("btn");
const statusEl = $("status");

const kmzFileInput = $("kmzFile");
const csvZipInput = $("csvZip");
const btnPatchKmz = $("btnPatchKmz");
const statusKmz = $("statusKmz");

// ================= CSV GENERATOR =================
fileInput.addEventListener("change", () => {
  btn.disabled = !fileInput.files.length;
});

const SEP = ";";
const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;

// ================= HEADERS =================
const FAT_FDT_HEADERS = [
  "Pole ID (New)",
  "Coordinate (Lat) NEW",
  "Coordinate (Long) NEW",
  "Pole Provider (New)",
  "Pole Type",
  "FAT ID/NETWORK ID",
];

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

// ================= GENERATE CSV =================
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

    const ws = wb.Sheets[sheetName];

    const headerRow = XLSX.utils.sheet_to_json(ws, { header: 1 })[0];
    const LAST_COL = headerRow[headerRow.length - 1];

    const master = XLSX.utils.sheet_to_json(ws, { defval: "" });

    const area =
      $("areaName").value ||
      master[0]["ID_Area"] ||
      file.name.replace(/\.(xlsx|xls)$/i, "");

    const HOME = [];
    const HOME_BIZ = [];

    master.forEach((r) => {
      const lastVal = String(r[LAST_COL] || "").toUpperCase();
      const isBiz = lastVal.includes("BIZ") || lastVal.includes("BIS");

      const row = Object.fromEntries(HOME_HEADERS.map((h) => [h, r[h] || ""]));
      row["Category BizPass"] = r["Category BizPass"] || "";

      (isBiz ? HOME_BIZ : HOME).push(row);
    });

    const FAT = [];
    master.forEach((r) => {
      if (String(r["Pole ID (New)"]).toUpperCase() === "1A") return;
      FAT.push(Object.fromEntries(FAT_FDT_HEADERS.map((h) => [h, r[h] || ""])));
    });

    const fdtSource = master.find(
      (r) => String(r["Pole ID (New)"]).toUpperCase() === "1A"
    );
    if (!fdtSource) throw new Error("POLE 1A tidak ditemukan");

    const FDT = [
      Object.fromEntries(
        FAT_FDT_HEADERS.map((h) => [h, fdtSource[h] || ""])
      ),
    ];

    const POLE = master.map((r) =>
      Object.fromEntries(POLE_HEADERS.map((h) => [h, r[h] || ""]))
    );

    const zip = new JSZip();
    const folder = zip.folder(area);

    folder.file("HOME.csv", toCSV(HOME_HEADERS, HOME));
    folder.file("HOME-BIZ.csv", toCSV(HOME_HEADERS, HOME_BIZ));
    folder.file("FAT.csv", toCSV(FAT_FDT_HEADERS, FAT));
    folder.file("FDT.csv", toCSV(FAT_FDT_HEADERS, FDT));
    folder.file("POLE.csv", toCSV(POLE_HEADERS, POLE));

    saveAs(await zip.generateAsync({ type: "blob" }), `${area}_POPUP.zip`);

    statusEl.textContent = "SELESAI ✔ CSV siap";
  } catch (e) {
    statusEl.textContent = "ERROR: " + e.message;
  } finally {
    btn.disabled = false;
  }
});

// =====================================================
// KMZ PATCHER (BARU)
// =====================================================
function enablePatchBtn() {
  btnPatchKmz.disabled = !(kmzFileInput.files.length && csvZipInput.files.length);
}
kmzFileInput.addEventListener("change", enablePatchBtn);
csvZipInput.addEventListener("change", enablePatchBtn);

const norm = (s) => String(s || "").trim().toUpperCase();

function parseCSV(text) {
  const [h, ...lines] = text.split(/\r?\n/).filter(Boolean);
  const headers = h.split(";");
  const rows = lines.map((l) => {
    const o = {};
    l.split(";").forEach((v, i) => (o[headers[i]] = v.replace(/^"|"$/g, "")));
    return o;
  });
  return { headers, rows };
}

function indexBy(rows, key) {
  const m = new Map();
  rows.forEach((r) => m.set(norm(r[key]), r));
  return m;
}

function patchFolder(doc, path, idx, headers) {
  let cur = doc.querySelector("Document");
  for (const p of path) {
    cur = [...cur.querySelectorAll(":scope>Folder")].find(
      (f) => norm(f.querySelector("name")?.textContent) === norm(p)
    );
    if (!cur) return;
  }

  cur.querySelectorAll("Placemark").forEach((pm) => {
    const id = norm(pm.querySelector("name")?.textContent);
    const row = idx.get(id);
    if (!row) return;

    pm.querySelector("ExtendedData")?.remove();
    const ext = doc.createElement("ExtendedData");

    headers.forEach((h) => {
      const d = doc.createElement("Data");
      d.setAttribute("name", h);
      const v = doc.createElement("value");
      v.textContent = row[h] || "";
      d.appendChild(v);
      ext.appendChild(d);
    });
    pm.appendChild(ext);
  });
}

btnPatchKmz.addEventListener("click", async () => {
  statusKmz.textContent = "Memproses KMZ...";
  btnPatchKmz.disabled = true;

  try {
    const csvZip = await JSZip.loadAsync(await csvZipInput.files[0].arrayBuffer());

    const read = async (n) => parseCSV(await csvZip.file(n).async("string"));

    const HOME = await read("HOME.csv");
    const HOME_BIZ = await read("HOME-BIZ.csv");
    const POLE = await read("POLE.csv");
    const FDT = await read("FDT.csv");
    const FAT = await read("FAT.csv");

    const kmz = await JSZip.loadAsync(await kmzFileInput.files[0].arrayBuffer());
    const kmlName = Object.keys(kmz.files).find((f) => f.endsWith(".kml"));
    const kmlText = await kmz.file(kmlName).async("string");

    const doc = new DOMParser().parseFromString(kmlText, "text/xml");

    patchFolder(doc, ["DISTRIBUSI", "HP", "HOME"], indexBy(HOME.rows, "HOMEPASS_ID"), HOME.headers);
    patchFolder(doc, ["DISTRIBUSI", "HP", "HOME-BIZ"], indexBy(HOME_BIZ.rows, "HOMEPASS_ID"), HOME_BIZ.headers);
    patchFolder(doc, ["DISTRIBUSI", "POLE"], indexBy(POLE.rows, "Pole ID (New)"), POLE.headers);
    patchFolder(doc, ["DISTRIBUSI", "FDT"], indexBy(FDT.rows, "Pole ID (New)"), FDT.headers);
    patchFolder(doc, ["DISTRIBUSI", "FAT"], indexBy(FAT.rows, "Pole ID (New)"), FAT.headers);

    kmz.file(kmlName, new XMLSerializer().serializeToString(doc));

    saveAs(await kmz.generateAsync({ type: "blob" }), "KMZ_POPUP_FINAL.kmz");
    statusKmz.textContent = "SELESAI ✔ KMZ siap dipakai";
  } catch (e) {
    statusKmz.textContent = "ERROR: " + e.message;
  } finally {
    enablePatchBtn();
  }
});
