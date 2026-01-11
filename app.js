document.addEventListener("DOMContentLoaded", () => {

  // =====================================================
  // ELEMENT
  // =====================================================
  const $ = (id) => document.getElementById(id);

  const kmzFileInput = $("kmzFile");
  const csvZipInput  = $("csvZip");
  const btnPatchKmz  = $("btnPatchKmz");
  const statusKmz    = $("statusKmz");

  function enablePatchBtn() {
    btnPatchKmz.disabled = !(
      kmzFileInput.files.length &&
      csvZipInput.files.length
    );
  }
  kmzFileInput.addEventListener("change", enablePatchBtn);
  csvZipInput.addEventListener("change", enablePatchBtn);

  // =====================================================
  // HELPERS
  // =====================================================
  const norm = (s) => String(s || "").trim().toUpperCase();

  function parseCSV(text) {
    const [h, ...lines] = text.split(/\r?\n/).filter(Boolean);
    const headers = h.split(";").map(x => x.replace(/^"|"$/g, ""));
    const rows = lines.map(l => {
      const o = {};
      l.split(";").forEach((v,i)=> o[headers[i]] = v.replace(/^"|"$/g,""));
      return o;
    });
    return { headers, rows };
  }

  function indexBy(rows, key) {
    const m = new Map();
    rows.forEach(r => m.set(norm(r[key]), r));
    return m;
  }

  // =====================================================
  // SCHEMA FIELD MAP (HARUS SAMA DENGAN KMZ CONTOH)
  // =====================================================
  const SCHEMA_MAP = {
    HOME: "HOME",
    "HOME-BIZ": "HOME_BIZ",
    POLE: "POLE",
    FDT: "FDT",
    FAT: "FAT"
  };

  // =====================================================
  // PATCH FOLDER (SCHEMA-BASED, GOOGLE EARTH NATIVE)
  // =====================================================
  function patchFolder(doc, path, idx, schemaId) {

    let cur = doc.querySelector("Document");
    for (const p of path) {
      cur = [...cur.querySelectorAll(":scope > Folder")].find(
        f => norm(f.querySelector("name")?.textContent) === norm(p)
      );
      if (!cur) return;
    }

    cur.querySelectorAll("Placemark").forEach(pm => {
      const id = norm(pm.querySelector("name")?.textContent);
      const row = idx.get(id);
      if (!row) return;

      // HAPUS SEMUA POPUP LAMA
      pm.querySelector("description")?.remove();
      pm.querySelector("ExtendedData")?.remove();

      // BANGUN SchemaData (INI KUNCI)
      const ext = doc.createElement("ExtendedData");
      const schemaData = doc.createElement("SchemaData");
      schemaData.setAttribute("schemaUrl", `#${schemaId}`);

      Object.entries(row).forEach(([k,v])=>{
        const sd = doc.createElement("SimpleData");
        sd.setAttribute("name", k.replace(/[^A-Za-z0-9_]/g,"_"));
        sd.textContent = v ?? "";
        schemaData.appendChild(sd);
      });

      ext.appendChild(schemaData);
      pm.appendChild(ext);
    });
  }

  // =====================================================
  // PATCH KMZ
  // =====================================================
  btnPatchKmz.addEventListener("click", async ()=> {
    statusKmz.textContent = "Memproses KMZ...";
    btnPatchKmz.disabled = true;

    try {
      // ===== LOAD CSV ZIP =====
      const csvZip = await JSZip.loadAsync(
        await csvZipInput.files[0].arrayBuffer()
      );

      const csvFiles = Object.keys(csvZip.files)
        .filter(f => f.toUpperCase().endsWith(".CSV"));
      if (!csvFiles.length) throw new Error("ZIP tidak berisi CSV");

      const basePath = csvFiles[0].includes("/")
        ? csvFiles[0].split("/")[0] + "/"
        : "";

      const readCSV = async (n) =>
        parseCSV(await csvZip.file(basePath+n).async("string"));

      const HOME     = await readCSV("HOME.csv");
      const HOME_BIZ = await readCSV("HOME-BIZ.csv");
      const POLE     = await readCSV("POLE.csv");
      const FDT      = await readCSV("FDT.csv");
      const FAT      = await readCSV("FAT.csv");

      // ===== LOAD KMZ =====
      const kmz = await JSZip.loadAsync(
        await kmzFileInput.files[0].arrayBuffer()
      );

      const kmlName = Object.keys(kmz.files).find(f => f.endsWith(".kml"));
      if (!kmlName) throw new Error("KML tidak ditemukan");

      const doc = new DOMParser().parseFromString(
        await kmz.file(kmlName).async("string"),
        "text/xml"
      );

      // ===== PATCH (SCHEMA MODE) =====
      patchFolder(doc, ["DISTRIBUSI","HP","HOME"],
        indexBy(HOME.rows,"HOMEPASS_ID"),
        SCHEMA_MAP.HOME
      );

      patchFolder(doc, ["DISTRIBUSI","HP","HOME-BIZ"],
        indexBy(HOME_BIZ.rows,"HOMEPASS_ID"),
        SCHEMA_MAP["HOME-BIZ"]
      );

      patchFolder(doc, ["DISTRIBUSI","POLE"],
        indexBy(POLE.rows,"Pole ID (New)"),
        SCHEMA_MAP.POLE
      );

      patchFolder(doc, ["DISTRIBUSI","FDT"],
        indexBy(FDT.rows,"Pole ID (New)"),
        SCHEMA_MAP.FDT
      );

      patchFolder(doc, ["DISTRIBUSI","FAT"],
        indexBy(FAT.rows,"Pole ID (New)"),
        SCHEMA_MAP.FAT
      );

      // ===== SAVE =====
      kmz.file(kmlName, new XMLSerializer().serializeToString(doc));
      saveAs(await kmz.generateAsync({type:"blob"}), "KMZ_POPUP_FINAL.kmz");

      statusKmz.textContent = "SELESAI ✔ POPUP IDENTIK DESAIN";
    } catch(e) {
      console.error(e);
      statusKmz.textContent = "ERROR: " + e.message;
    } finally {
      enablePatchBtn();
    }
  });

});
