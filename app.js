btnPatchKmz.addEventListener("click", async () => {
  statusKmz.textContent = "Memproses KMZ...";
  btnPatchKmz.disabled = true;

  try {
    // ================= LOAD CSV ZIP =================
    const csvZip = await JSZip.loadAsync(
      await csvZipInput.files[0].arrayBuffer()
    );

    // ===== DETEKSI CSV DI ROOT ATAU DALAM FOLDER =====
    const csvFiles = Object.keys(csvZip.files).filter(
      (f) => f.toUpperCase().endsWith(".CSV")
    );

    if (!csvFiles.length) {
      throw new Error("ZIP tidak berisi file CSV");
    }

    // Ambil base path (misal: "BABAT JERAWAT/")
    const basePath = csvFiles[0].includes("/")
      ? csvFiles[0].substring(0, csvFiles[0].lastIndexOf("/") + 1)
      : "";

    const readCSV = async (name) => {
      const file = csvZip.file(basePath + name);
      if (!file) throw new Error(`CSV tidak ditemukan: ${name}`);
      return parseCSV(await file.async("string"));
    };

    // ================= LOAD CSV =================
    const HOME = await readCSV("HOME.csv");
    const HOME_BIZ = await readCSV("HOME-BIZ.csv");
    const POLE = await readCSV("POLE.csv");
    const FDT = await readCSV("FDT.csv");
    const FAT = await readCSV("FAT.csv");

    // ================= LOAD KMZ =================
    const kmz = await JSZip.loadAsync(
      await kmzFileInput.files[0].arrayBuffer()
    );

    const kmlName = Object.keys(kmz.files).find((f) =>
      f.toLowerCase().endsWith(".kml")
    );

    if (!kmlName) {
      throw new Error("File KML tidak ditemukan di dalam KMZ");
    }

    const kmlText = await kmz.file(kmlName).async("string");
    const doc = new DOMParser().parseFromString(kmlText, "text/xml");

    // ================= PATCH FOLDER =================
    patchFolder(
      doc,
      ["DISTRIBUSI", "HP", "HOME"],
      indexBy(HOME.rows, "HOMEPASS_ID"),
      HOME.headers
    );

    patchFolder(
      doc,
      ["DISTRIBUSI", "HP", "HOME-BIZ"],
      indexBy(HOME_BIZ.rows, "HOMEPASS_ID"),
      HOME_BIZ.headers
    );

    patchFolder(
      doc,
      ["DISTRIBUSI", "POLE"],
      indexBy(POLE.rows, "Pole ID (New)"),
      POLE.headers
    );

    patchFolder(
      doc,
      ["DISTRIBUSI", "FDT"],
      indexBy(FDT.rows, "Pole ID (New)"),
      FDT.headers
    );

    patchFolder(
      doc,
      ["DISTRIBUSI", "FAT"],
      indexBy(FAT.rows, "Pole ID (New)"),
      FAT.headers
    );

    // ================= SAVE KMZ =================
    kmz.file(kmlName, new XMLSerializer().serializeToString(doc));

    saveAs(
      await kmz.generateAsync({ type: "blob" }),
      "KMZ_POPUP_FINAL.kmz"
    );

    statusKmz.textContent = "SELESAI ✔ KMZ siap dipakai di Google Earth";
  } catch (e) {
    console.error(e);
    statusKmz.textContent = "ERROR: " + e.message;
  } finally {
    enablePatchBtn();
  }
});
// ================= FORCE ENABLE PATCH BUTTON =================
function enablePatchBtn() {
  btnPatchKmz.disabled = !(
    kmzFileInput.files.length &&
    csvZipInput.files.length
  );
}

// PENTING: bind ulang event
kmzFileInput.addEventListener("change", enablePatchBtn);
csvZipInput.addEventListener("change", enablePatchBtn);

// PENTING: paksa cek saat page load
setTimeout(enablePatchBtn, 100);
