document.addEventListener("DOMContentLoaded", () => {

  // ================= ELEMENT =================
  const kmzFileInput = document.getElementById("kmzFile");
  const csvZipInput  = document.getElementById("csvZip");
  const btnPatchKmz  = document.getElementById("btnPatchKmz");
  const statusKmz    = document.getElementById("statusKmz");

  if (!kmzFileInput || !csvZipInput || !btnPatchKmz) {
    alert("ERROR: Elemen KMZ tidak ditemukan. Cek id HTML.");
    return;
  }

  // ================= ENABLE BUTTON =================
  function enablePatchBtn() {
    btnPatchKmz.disabled = !(
      kmzFileInput.files.length > 0 &&
      csvZipInput.files.length > 0
    );
  }

  kmzFileInput.addEventListener("change", enablePatchBtn);
  csvZipInput.addEventListener("change", enablePatchBtn);

  // ================= HELPERS =================
  const norm = (s) => String(s || "").trim().toUpperCase();

  function parseCSV(text) {
    const [h, ...lines] = text.split(/\r?\n/).filter(Boolean);
    const headers = h.split(";").map(x => x.replace(/^"|"$/g, ""));
    const rows = lines.map((l) => {
      const o = {};
      l.split(";").forEach((v, i) => {
        o[headers[i]] = v.replace(/^"|"$/g, "");
      });
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
      cur = [...cur.querySelectorAll(":scope > Folder")].find(
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

  // ================= PATCH KMZ =================
  btnPatchKmz.addEventListener("click", async () => {
    statusKmz.textContent = "Memproses KMZ...";
    btnPatchKmz.disabled = true;

    try {
      const csvZip = await JSZip.loadAsync(
        await csvZipInput.files[0].arrayBuffer()
      );

      const csvFiles = Object.keys(csvZip.files).filter(f =>
        f.toUpperCase().endsWith(".CSV")
      );
      if (!csvFiles.length) throw new Error("ZIP tidak berisi CSV");

      const basePath = csvFiles[0].includes("/")
        ? csvFiles[0].split("/")[0] + "/"
        : "";

      const readCSV = async (name) => {
        const f = csvZip.file(basePath + name);
        if (!f) throw new Error("CSV tidak ditemukan: " + name);
        return parseCSV(await f.async("string"));
      };

      const HOME     = await readCSV("HOME.csv");
      const HOME_BIZ = await readCSV("HOME-BIZ.csv");
      const POLE     = await readCSV("POLE.csv");
      const FDT      = await readCSV("FDT.csv");
      const FAT      = await readCSV("FAT.csv");

      const kmz = await JSZip.loadAsync(
        await kmzFileInput.files[0].arrayBuffer()
      );

      const kmlName = Object.keys(kmz.files).find(f => f.endsWith(".kml"));
      if (!kmlName) throw new Error("KML tidak ditemukan");

      const kmlText = await kmz.file(kmlName).async("string");
      const doc = new DOMParser().parseFromString(kmlText, "text/xml");

      patchFolder(doc, ["DISTRIBUSI","HP","HOME"], indexBy(HOME.rows,"HOMEPASS_ID"), HOME.headers);
      patchFolder(doc, ["DISTRIBUSI","HP","HOME-BIZ"], indexBy(HOME_BIZ.rows,"HOMEPASS_ID"), HOME_BIZ.headers);
      patchFolder(doc, ["DISTRIBUSI","POLE"], indexBy(POLE.rows,"Pole ID (New)"), POLE.headers);
      patchFolder(doc, ["DISTRIBUSI","FDT"], indexBy(FDT.rows,"Pole ID (New)"), FDT.headers);
      patchFolder(doc, ["DISTRIBUSI","FAT"], indexBy(FAT.rows,"Pole ID (New)"), FAT.headers);

      kmz.file(kmlName, new XMLSerializer().serializeToString(doc));
      saveAs(await kmz.generateAsync({ type: "blob" }), "KMZ_POPUP_FINAL.kmz");

      statusKmz.textContent = "SELESAI ✔ KMZ siap dipakai";
    } catch (e) {
      console.error(e);
      statusKmz.textContent = "ERROR: " + e.message;
    } finally {
      enablePatchBtn();
    }
  });

});
