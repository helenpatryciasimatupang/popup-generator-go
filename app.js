document.addEventListener("DOMContentLoaded", () => {

  // =====================================================
  // ELEMENT
  // =====================================================
  const $ = (id) => document.getElementById(id);

  const fileInput   = $("file");
  const areaInput   = $("areaName");
  const btnGenCSV   = $("btn");
  const statusCSV   = $("status");

  const kmzFileInput = $("kmzFile");
  const csvZipInput  = $("csvZip");
  const btnPatchKmz  = $("btnPatchKmz");
  const statusKmz    = $("statusKmz");

  // =====================================================
  // ENABLE BUTTON
  // =====================================================
  if (fileInput && btnGenCSV) {
    fileInput.addEventListener("change", () => {
      btnGenCSV.disabled = !fileInput.files.length;
    });
  }

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
  const SEP = ";";
  const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const norm = (s) => String(s || "").trim().toUpperCase();

  function toCSV(headers, rows) {
    const out = [];
    out.push(headers.map(esc).join(SEP));
    rows.forEach(r => {
      out.push(headers.map(h => esc(r[h] || "")).join(SEP));
    });
    return out.join("\n");
  }

  function parseCSV(text) {
    const [h, ...lines] = text.split(/\r?\n/).filter(Boolean);
    const headers = h.split(";").map(x => x.replace(/^"|"$/g, ""));
    const rows = lines.map(l => {
      const o = {};
      l.split(";").forEach((v,i)=>{
        o[headers[i]] = v.replace(/^"|"$/g,"");
      });
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
  // HEADERS
  // =====================================================
  const FAT_FDT_HEADERS = [
    "Pole ID (New)",
    "Coordinate (Lat) NEW",
    "Coordinate (Long) NEW",
    "Pole Provider (New)",
    "Pole Type",
    "FAT ID/NETWORK ID",
  ];

  const HOME_HEADERS = [
    "HOMEPASS_ID","CLUSTER_NAME","PREFIX_ADDRESS","STREET_NAME","HOUSE_NUMBER",
    "BLOCK","FLOOR","RT","RW","DISTRICT","SUB_DISTRICT","FDT_CODE","FAT_CODE",
    "BUILDING_LATITUDE","BUILDING_LONGITUDE","Category BizPass","POST CODE",
    "ADDRESS POLE / FAT","OV_UG","HOUSE_COMMENT_","BUILDING_NAME","TOWER","APTN",
    "FIBER_NODE__HFC_","ID_Area","Clamp_Hook_ID","DEPLOYMENT_TYPE","NEED_SURVEY",
  ];

  const POLE_HEADERS = [
    "Pole ID (New)",
    "Coordinate (Lat) NEW",
    "Coordinate (Long) NEW",
    "Pole Provider (New)",
    "Pole Type",
    "LINE",
  ];

  // =====================================================
  // PATCH FOLDER — GOOGLE EARTH NATIVE POPUP (FIX)
  // =====================================================
  function patchFolder(doc, path, idx, headers) {
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

      // bersihkan popup lama
      pm.querySelector("description")?.remove();
      pm.querySelector("styleUrl")?.remove();
      pm.querySelector("Style")?.remove();
      pm.querySelector("ExtendedData")?.remove();

      const ext = doc.createElement("ExtendedData");

      headers.forEach(h => {
        const d = doc.createElement("Data");
        d.setAttribute("name", h);

        // 🔴 INI YANG PALING PENTING
        const display = doc.createElement("displayName");
        display.textContent = h;

        const v = doc.createElement("value");
        v.textContent = row[h] ?? "";

        d.appendChild(display);
        d.appendChild(v);
        ext.appendChild(d);
      });

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
      const csvZip = await JSZip.loadAsync(
        await csvZipInput.files[0].arrayBuffer()
      );

      const csvFiles = Object.keys(csvZip.files)
        .filter(f => f.toUpperCase().endsWith(".CSV"));

      if (!csvFiles.length) throw new Error("ZIP tidak berisi CSV");

      const basePath = csvFiles[0].includes("/")
        ? csvFiles[0].split("/")[0] + "/"
        : "";

      const readCSV = async n =>
        parseCSV(await csvZip.file(basePath + n).async("string"));

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

      const doc = new DOMParser().parseFromString(
        await kmz.file(kmlName).async("string"),
        "text/xml"
      );

      patchFolder(doc, ["DISTRIBUSI","HP","HOME"], indexBy(HOME.rows,"HOMEPASS_ID"), HOME.headers);
      patchFolder(doc, ["DISTRIBUSI","HP","HOME-BIZ"], indexBy(HOME_BIZ.rows,"HOMEPASS_ID"), HOME_BIZ.headers);
      patchFolder(doc, ["DISTRIBUSI","POLE"], indexBy(POLE.rows,"Pole ID (New)"), POLE.headers);
      patchFolder(doc, ["DISTRIBUSI","FDT"], indexBy(FDT.rows,"Pole ID (New)"), FDT.headers);
      patchFolder(doc, ["DISTRIBUSI","FAT"], indexBy(FAT.rows,"Pole ID (New)"), FAT.headers);

      kmz.file(kmlName, new XMLSerializer().serializeToString(doc));
      saveAs(await kmz.generateAsync({ type:"blob" }), "KMZ_POPUP_FINAL.kmz");

      statusKmz.textContent = "SELESAI ✔ KMZ siap dipakai";
    } catch(e) {
      console.error(e);
      statusKmz.textContent = "ERROR: " + e.message;
    } finally {
      enablePatchBtn();
    }
  });

});
