btnPatchKmz.addEventListener("click", async () => {
  statusKmz.textContent = "Memproses KMZ...";
  btnPatchKmz.disabled = true;

  try {
    // ================= LOAD CSV ZIP =================
    const csvZip = await JSZip.loadAsync(
      await csvZipInput.files[0].arrayBuffer()
    );

    // 🔍 CARI FOLDER AREA (misal: BABAT JERAWAT/)
    const folderName = Object.keys(csvZip.files).find(
      (k) => csvZip.files[k].dir
    );

    if (!folderName) {
      throw new Error("Folder area tidak ditemukan di ZIP CSV");
    }

    const readCSV = async (name) => {
      const file = csvZip.file(folderName + name);
      if (!file) throw new Error(`CSV tidak ditemukan: ${name}`);
      return parseCSV(await file.async("string"));
    };

    const HOME = await readCSV("HOME.csv");
    const HOME_BIZ = await readCSV("HOME-BIZ.csv");
    const POLE = await readCSV("POLE.csv");
    const FDT = await readCSV("FDT.csv");
    const FAT = await readCSV("FAT.csv");

    // ================= LOAD KMZ =================
    const kmz = await JSZip.loadAsync(
      await kmzFileInput.files[0].arrayBuffer()
    );

    const kmlName = Object.keys(kmz.files).find((f) => f.endsWith(".kml"));
    if (!kmlName) throw new Error("File KML tidak ditemukan di KMZ");

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
