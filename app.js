// ===== Pop Up CSV Generator (FINAL FIX) =====

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

function toCSV(rows) {
  if (!rows.length) return "";
  const headers = Object.keys(rows[0]);
  const esc = (v) =>
    `"${String(v ?? "").replace(/"/g, '""')}"`;
  const sep = ";";

  const lines = [];
  lines.push(headers.map(esc).join(sep));
  rows.forEach((r) =>
    lines.push(headers.map((h) => esc(r[h])).join(sep))
  );
  return lines.join("\n");
}

function uniqBy(arr, keyFn) {
  const map = {};
  arr.forEach((x) => {
    const k = keyFn(x);
    if (k) map[k] = x;
  });
  return Object.values(map);
}

btn.addEventListener("click", async () => {
  const file = fileInput.files[0];
  if (!file) return;

  statusEl.textContent = "Membaca Excel...";
  btn.disabled = true;

  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array" });

    const sheetName =
      wb.SheetNames.find((n) => n === "Master Data") ||
      wb.SheetNames[0];

    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, {
      defval: "",
      raw: false,
    });

    if (!rows.length) throw new Error("Sheet kosong");

    const headers = Object.keys(rows[0]);
    const colType = headers.find((h) => h.trim() === "HOME/HOME-BIZ");
    if (!colType) throw new Error("Kolom HOME/HOME-BIZ tidak ditemukan");

    const HOME = [];
    const HOME_BIZ = [];

    rows.forEach((r) => {
      const v = String(r[colType]).toUpperCase().trim();
      if (v === "HOME") HOME.push(r);
      if (v === "HOME-BIZ" || v === "HOMEBIZ") HOME_BIZ.push(r);
    });

    const FAT = uniqBy(rows, (r) => r["FAT ID/NETWORK ID"]);
    const FDT = uniqBy(rows, (r) => r["FDT_CODE"]);
    const POLE = uniqBy(rows, (r) => r["Pole ID (New)"]);
    const HOOK = uniqBy(rows, (r) => r["Clamp_Hook_ID"]);

    // update preview
    cHome.textContent = HOME.length;
    cHomeBiz.textContent = HOME_BIZ.length;
    cFat.textContent = FAT.length;
    cFdt.textContent = FDT.length;
    cPole.textContent = POLE.length;
    cHook.textContent = HOOK.length;

    const zip = new JSZip();

    const area =
      $("areaName").value ||
      rows[0]["ID_Area"] ||
      file.name.replace(/\.(xlsx|xls)$/i, "");

    const folder = zip.folder(area);

    folder.file("HOME.csv", toCSV(HOME));
    folder.file("HOME-BIZ.csv", toCSV(HOME_BIZ));
    folder.file("FAT.csv", toCSV(FAT));
    folder.file("FDT.csv", toCSV(FDT));
    folder.file("POLE.csv", toCSV(POLE));
    folder.file("HOOK.csv", toCSV(HOOK));

    const blob = await zip.generateAsync({ type: "blob" });
    saveAs(blob, `${area}_popups.zip`);

    statusEl.textContent = "Selesai ✔ ZIP terdownload";
  } catch (err) {
    console.error(err);
    statusEl.textContent = "ERROR: " + err.message;
  } finally {
    btn.disabled = false;
  }
});
