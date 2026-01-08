/* Pop Up CSV Generator - FIX VERSION
   Support Excel besar & kompleks
   HOME vs HOME-BIZ dari kolom "HOME/HOME-BIZ"
*/

const $ = (id) => document.getElementById(id);

function normalizeHeader(h) {
  return String(h || "").trim();
}

function toCSV(rows, headers) {
  const sep = ";";
  const escape = (v) =>
    `"${String(v ?? "").replace(/"/g, '""')}"`;
  const lines = [];
  lines.push(headers.map(escape).join(sep));
  rows.forEach((r) => {
    lines.push(headers.map((h) => escape(r[h])).join(sep));
  });
  return lines.join("\n");
}

async function readExcel(file) {
  const data = await file.arrayBuffer();
  let wb;
  try {
    wb = XLSX.read(data, { type: "array" });
  } catch (e) {
    // fallback
    const binary = new Uint8Array(data)
      .reduce((acc, b) => acc + String.fromCharCode(b), "");
    wb = XLSX.read(binary, { type: "binary" });
  }
  return wb;
}

$("generateBtn").onclick = async () => {
  const file = $("excelFile").files[0];
  if (!file) {
    alert("Pilih file Excel dulu");
    return;
  }

  $("status").textContent = "Membaca Excel...";
  try {
    const wb = await readExcel(file);

    const sheetName =
      wb.SheetNames.find((n) => n === "Master Data") ||
      wb.SheetNames[0];

    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, {
      defval: "",
      raw: false,
    });

    if (!json.length) throw new Error("Sheet kosong");

    const headers = Object.keys(json[0]).map(normalizeHeader);
    const colType = headers.find((h) => h === "HOME/HOME-BIZ");
    if (!colType)
      throw new Error("Kolom HOME/HOME-BIZ tidak ditemukan");

    const HOME = [];
    const HOMEBIZ = [];

    json.forEach((r) => {
      const t = String(r[colType]).toUpperCase().trim();
      if (t === "HOME") HOME.push(r);
      if (t === "HOME-BIZ" || t === "HOMEBIZ") HOMEBIZ.push(r);
    });

    const uniq = (arr, key) =>
      Object.values(
        arr.reduce((a, x) => {
          if (x[key]) a[x[key]] = x;
          return a;
        }, {})
      );

    const FDT = uniq(json, "FDT_CODE");
    const FAT = uniq(json, "FAT ID/NETWORK ID");
    const HOOK = uniq(json, "Clamp_Hook_ID");
    const POLE = uniq(json, "Pole ID (New)");

    const zip = new JSZip();
    const area =
      $("areaName").value ||
      json[0]["ID_Area"] ||
      file.name.replace(".xlsx", "");

    const folder = zip.folder(area);

    folder.file("HOME.csv", toCSV(HOME, headers));
    folder.file("HOME-BIZ.csv", toCSV(HOMEBIZ, headers));
    folder.file("FDT.csv", toCSV(FDT, Object.keys(FDT[0] || {})));
    folder.file("FAT.csv", toCSV(FAT, Object.keys(FAT[0] || {})));
    folder.file("HOOK.csv", toCSV(HOOK, Object.keys(HOOK[0] || {})));
    folder.file("POLE.csv", toCSV(POLE, Object.keys(POLE[0] || {})));

    const blob = await zip.generateAsync({ type: "blob" });
    saveAs(blob, `${area}_popups.zip`);

    $("status").textContent = "Selesai ✔ ZIP terdownload";
  } catch (err) {
    console.error(err);
    $("status").textContent =
      "ERROR: " + (err.message || "Gagal baca Excel");
  }
};
