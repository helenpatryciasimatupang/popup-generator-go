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
  // ENABLE BUTTONS
  // =====================================================
  if (fileInput && btnGenCSV) {
    fileInput.addEventListener("change", () => {
      btnGenCSV.disabled = !fileInput.files.length;
    });
  }

  function enablePatchBtn() {
    btnPatchKmz.disabled = !(
      kmzFileInput.files.length > 0 &&
      csvZipInput.files.length > 0
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
    rows.forEach(r => out.push(headers.map(h => esc(r[h] || "")).join(SEP)));
    return out.join("\n");
  }

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
  // CSV GENERATOR
  // =====================================================
  if (btnGenCSV) {
    btnGenCSV.addEventListener("click", async () => {
      statusCSV.textContent = "Memproses Master Excel...";
      btnGenCSV.disabled = true;

      try {
        const buf = await fileInput.files[0].arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const sheetName = wb.SheetNames.includes("Master Data") ? "Master Data" : wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];

        const headerRow = XLSX.utils.sheet_to_json(ws, { header:1 })[0];
        const LAST_COL = headerRow[headerRow.length-1];

        const master = XLSX.utils.sheet_to_json(ws, { defval:"" });
        const area = areaInput.value || master[0]["ID_Area"] || "AREA";

        // HOME / HOME-BIZ
        const HOME = [], HOME_BIZ = [];
        master.forEach(r=>{
          const isBiz = String(r[LAST_COL]).toUpperCase().includes("BIZ");
          const row = Object.fromEntries(HOME_HEADERS.map(h=>[h,r[h]||""]));
          row["Category BizPass"] = r["Category BizPass"] || "";
          (isBiz?HOME_BIZ:HOME).push(row);
        });

        // FAT (tanpa 1A)
        const FAT = [];
        master.forEach(r=>{
          if (norm(r["Pole ID (New)"])==="1A") return;
          FAT.push(Object.fromEntries(FAT_FDT_HEADERS.map(h=>[h,r[h]||""])));
        });

        // FDT (hanya 1A)
        const fdtSource = master.find(r=>norm(r["Pole ID (New)"])==="1A");
        if (!fdtSource) throw new Error("POLE 1A tidak ditemukan");
        const FDT = [Object.fromEntries(FAT_FDT_HEADERS.map(h=>[h,fdtSource[h]||""]))];

        // POLE
        const POLE = master.map(r=>Object.fromEntries(POLE_HEADERS.map(h=>[h,r[h]||""])));

        const zip = new JSZip();
        const folder = zip.folder(area);
        folder.file("HOME.csv", toCSV(HOME_HEADERS, HOME));
        folder.file("HOME-BIZ.csv", toCSV(HOME_HEADERS, HOME_BIZ));
        folder.file("FAT.csv", toCSV(FAT_FDT_HEADERS, FAT));
        folder.file("FDT.csv", toCSV(FAT_FDT_HEADERS, FDT));
        folder.file("POLE.csv", toCSV(POLE_HEADERS, POLE));

        saveAs(await zip.generateAsync({type:"blob"}), `${area}_POPUP.zip`);
        statusCSV.textContent = "SELESAI ✔ CSV siap";
      } catch(e){
        statusCSV.textContent = "ERROR: " + e.message;
      } finally {
        btnGenCSV.disabled = false;
      }
    });
  }

  // =====================================================
  // PATCH FOLDER (POP-UP FINAL TANPA GAMBAR)
  // =====================================================
  function patchFolder(doc, path, idx, headers) {
    let cur = doc.querySelector("Document");
    for (const p of path) {
      cur = [...cur.querySelectorAll(":scope > Folder")].find(
        f => f.querySelector("name")?.textContent.trim().toUpperCase() === p.toUpperCase()
      );
      if (!cur) return;
    }

    cur.querySelectorAll("Placemark").forEach(pm=>{
      const id = norm(pm.querySelector("name")?.textContent);
      const row = idx.get(id);
      if (!row) return;

      pm.querySelector("description")?.remove();
      pm.querySelector("styleUrl")?.remove();
      pm.querySelector("ExtendedData")?.remove();

      let html = `<table border="1" style="border-collapse:collapse;font-size:12px">`;
      headers.forEach(h=>{
        html += `<tr><td><b>${h}</b></td><td>${row[h]??""}</td></tr>`;
      });
      html += `</table>`;

      const style = doc.createElement("Style");
      const balloon = doc.createElement("BalloonStyle");
      const text = doc.createElement("text");
      text.textContent = `<![CDATA[${html}]]>`;
      balloon.appendChild(text);
      style.appendChild(balloon);
      pm.appendChild(style);
    });
  }

  // =====================================================
  // PATCH KMZ
  // =====================================================
  btnPatchKmz.addEventListener("click", async ()=>{
    statusKmz.textContent = "Memproses KMZ...";
    btnPatchKmz.disabled = true;

    try {
      const csvZip = await JSZip.loadAsync(await csvZipInput.files[0].arrayBuffer());
      const csvFiles = Object.keys(csvZip.files).filter(f=>f.toUpperCase().endsWith(".CSV"));
      const basePath = csvFiles[0].includes("/") ? csvFiles[0].split("/")[0]+"/" : "";

      const readCSV = async n => parseCSV(await csvZip.file(basePath+n).async("string"));
      const HOME = await readCSV("HOME.csv");
      const HOME_BIZ = await readCSV("HOME-BIZ.csv");
      const POLE = await readCSV("POLE.csv");
      const FDT = await readCSV("FDT.csv");
      const FAT = await readCSV("FAT.csv");

      const kmz = await JSZip.loadAsync(await kmzFileInput.files[0].arrayBuffer());
      const kmlName = Object.keys(kmz.files).find(f=>f.endsWith(".kml"));
      const doc = new DOMParser().parseFromString(await kmz.file(kmlName).async("string"),"text/xml");

      patchFolder(doc,["DISTRIBUSI","HP","HOME"],indexBy(HOME.rows,"HOMEPASS_ID"),HOME.headers);
      patchFolder(doc,["DISTRIBUSI","HP","HOME-BIZ"],indexBy(HOME_BIZ.rows,"HOMEPASS_ID"),HOME_BIZ.headers);
      patchFolder(doc,["DISTRIBUSI","POLE"],indexBy(POLE.rows,"Pole ID (New)"),POLE.headers);
      patchFolder(doc,["DISTRIBUSI","FDT"],indexBy(FDT.rows,"Pole ID (New)"),FDT.headers);
      patchFolder(doc,["DISTRIBUSI","FAT"],indexBy(FAT.rows,"Pole ID (New)"),FAT.headers);

      kmz.file(kmlName,new XMLSerializer().serializeToString(doc));
      saveAs(await kmz.generateAsync({type:"blob"}),"KMZ_POPUP_FINAL.kmz");
      statusKmz.textContent = "SELESAI ✔ KMZ siap dipakai";
    } catch(e){
      statusKmz.textContent = "ERROR: "+e.message;
    } finally {
      enablePatchBtn();
    }
  });

});
