document.addEventListener("DOMContentLoaded", () => {

  // =====================================================
  // ELEMENT HELPER
  // =====================================================
  const $ = (id) => document.getElementById(id);

  // ================= DATA BASE → MASTER POP UP =========
  const dbFile        = $("dbFile");
  const dbArea        = $("dbArea");
  const dbDistrict    = $("dbDistrict");
  const dbSubDistrict = $("dbSubDistrict");
  const dbPostCode    = $("dbPostCode");
  const dbFdtCode     = $("dbFdtCode");
  const dbIdArea      = $("dbIdArea");
  const btnGenMaster  = $("btnGenMaster");
  const statusMaster  = $("statusMaster");

  // ================= MASTER POP UP → CSV ===============
  const fileMasterCSV = $("file");
  const areaInput     = $("areaName");
  const btnCSV        = $("btn");
  const statusCSV     = $("status");

  // ================= CSV → KMZ =========================
  const kmzFileInput  = $("kmzFile");
  const csvZipInput   = $("csvZip");
  const btnPatchKmz   = $("btnPatchKmz");
  const statusKmz     = $("statusKmz");

  // =====================================================
  // UTIL
  // =====================================================
  const norm = (s) => String(s || "").trim().toUpperCase();
  const raw  = (s) => String(s || "").trim();

  // =====================================================
  // ENABLE BUTTONS (INI YANG TADI HILANG)
  // =====================================================

  // MASTER POP UP → CSV
  if (fileMasterCSV && btnCSV) {
    fileMasterCSV.addEventListener("change", () => {
      btnCSV.disabled = !fileMasterCSV.files.length;
    });
  }

  // CSV → KMZ
  function enablePatchBtn() {
    btnPatchKmz.disabled = !(
      kmzFileInput.files.length &&
      csvZipInput.files.length
    );
  }
  kmzFileInput?.addEventListener("change", enablePatchBtn);
  csvZipInput?.addEventListener("change", enablePatchBtn);

  // =====================================================
  // 1️⃣ DATA BASE → MASTER POP UP
  // =====================================================
  btnGenMaster?.addEventListener("click", async () => {
    statusMaster.textContent = "Memproses DATA BASE → MASTER POP UP...";

    try {
      if (!dbFile.files.length) throw new Error("File DATA BASE belum dipilih");

      const buf = await dbFile.files[0].arrayBuffer();
      const wb  = XLSX.read(buf, { type: "array" });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

      const HOME = [];
      const HOME_BIZ = [];
      const POLE = [];

      rows.forEach(r => {
        const folder = raw(r.FolderPath);
        const name   = raw(r.Name);

        // ================= HOME / HOME-BIZ =================
        if (folder.includes("/HOME")) {

          const isBiz = folder.includes("HOME-BIZ");
          const street = folder.split("/").slice(-2)[0];

          const blockMatch = name.match(/^([A-Z]+)\d+/);
          const BLOCK = blockMatch ? blockMatch[1] : "-";

          const row = {
            HOMEPASS_ID: "-",
            CLUSTER_NAME: dbArea.value,
            PREFIX_ADDRESS: "JL.",
            STREET_NAME: street,
            HOUSE_NUMBER: name,
            BLOCK,
            FLOOR: "-",
            RT: "-",
            RW: "-",
            DISTRICT: dbDistrict.value,
            SUB_DISTRICT: dbSubDistrict.value,
            FDT_CODE: dbFdtCode.value,
            FAT_CODE: "FAT",
            BUILDING_LATITUDE: r.Y,
            BUILDING_LONGITUDE: r.X,
            "Category BizPass": "",
            "POST CODE": dbPostCode.value,
            "ADDRESS POLE / FAT": "-",
            OV_UG: "O",
            HOUSE_COMMENT_: "NEED SURVEY",
            BUILDING_NAME: "-",
            TOWER: "-",
            APTN: "-",
            FIBER_NODE__HFC_: "-",
            ID_Area: dbIdArea.value,
            Clamp_Hook_ID: "CLAIM_HOOK",
            DEPLOYMENT_TYPE: "FAT EXT",
            NEED_SURVEY: "YES"
          };

          isBiz ? HOME_BIZ.push(row) : HOME.push(row);
        }

        // ================= POLE =================
        if (folder.includes("/POLE")) {
          POLE.push({
            "Pole ID (New)": name,
            "Coordinate (Lat) NEW": r.Y,
            "Coordinate (Long) NEW": r.X,
            "Pole Provider (New)": "LN (Existing)",
            "Pole Type": "Pole 7/250",
            LINE: "",
            "FAT ID/NETWORK ID": "FAT"
          });
        }
      });

      const wbOut = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wbOut, XLSX.utils.json_to_sheet(HOME), "HOME");
      XLSX.utils.book_append_sheet(wbOut, XLSX.utils.json_to_sheet(HOME_BIZ), "HOME-BIZ");
      XLSX.utils.book_append_sheet(wbOut, XLSX.utils.json_to_sheet(POLE), "POLE");

      XLSX.writeFile(wbOut, `MASTER_POP_UP_${dbArea.value}.xlsx`);
      statusMaster.textContent = "SELESAI ✔ MASTER POP UP TERBENTUK";

    } catch (e) {
      console.error(e);
      statusMaster.textContent = "ERROR: " + e.message;
    }
  });

  // =====================================================
  // 2️⃣ MASTER POP UP → CSV ZIP
  // =====================================================
  btnCSV?.addEventListener("click", async () => {
    statusCSV.textContent = "Memproses MASTER POP UP → CSV...";

    try {
      const buf = await fileMasterCSV.files[0].arrayBuffer();
      const wb  = XLSX.read(buf, { type: "array" });

      const zip = new JSZip();
      const area = areaInput.value || "AREA";

      wb.SheetNames.forEach(name => {
        const ws = wb.Sheets[name];
        const csv = XLSX.utils.sheet_to_csv(ws, { FS: ";" });
        zip.file(`${area}/${name}.csv`, csv);
      });

      saveAs(await zip.generateAsync({ type: "blob" }), `${area}_POPUP.zip`);
      statusCSV.textContent = "SELESAI ✔ CSV siap";

    } catch (e) {
      console.error(e);
      statusCSV.textContent = "ERROR: " + e.message;
    }
  });

  // =====================================================
  // 3️⃣ CSV → KMZ (SCHEMA MODE – IDENTIK MANUAL)
  // =====================================================
  function parseCSV(text) {
    const [h, ...lines] = text.split(/\r?\n/).filter(Boolean);
    const headers = h.split(";").map(x => x.replace(/^"|"$/g,""));
    const rows = lines.map(l => {
      const o = {};
      l.split(";").forEach((v,i)=> o[headers[i]] = v.replace(/^"|"$/g,""));
      return o;
    });
    return { headers, rows };
  }

  function indexBy(rows,key){
    const m=new Map();
    rows.forEach(r=>m.set(norm(r[key]),r));
    return m;
  }

  function patchFolder(doc,path,idx,schemaId){
    let cur=doc.querySelector("Document");
    for(const p of path){
      cur=[...cur.querySelectorAll(":scope>Folder")]
        .find(f=>norm(f.querySelector("name")?.textContent)===norm(p));
      if(!cur) return;
    }

    cur.querySelectorAll("Placemark").forEach(pm=>{
      const id=norm(pm.querySelector("name")?.textContent);
      const row=idx.get(id);
      if(!row) return;

      pm.querySelector("ExtendedData")?.remove();
      pm.querySelector("description")?.remove();

      const ext=doc.createElement("ExtendedData");
      const sd=doc.createElement("SchemaData");
      sd.setAttribute("schemaUrl","#"+schemaId);

      Object.entries(row).forEach(([k,v])=>{
        const s=doc.createElement("SimpleData");
        s.setAttribute("name",k.replace(/[^A-Za-z0-9_]/g,"_"));
        s.textContent=v;
        sd.appendChild(s);
      });

      ext.appendChild(sd);
      pm.appendChild(ext);
    });
  }

  btnPatchKmz?.addEventListener("click", async ()=>{
    statusKmz.textContent="Memproses KMZ...";

    try {
      const csvZip=await JSZip.loadAsync(await csvZipInput.files[0].arrayBuffer());
      const base=Object.keys(csvZip.files).find(f=>f.endsWith(".csv")).split("/")[0]+"/";

      const HOME=parseCSV(await csvZip.file(base+"HOME.csv").async("string"));
      const HOME_BIZ=parseCSV(await csvZip.file(base+"HOME-BIZ.csv").async("string"));
      const POLE=parseCSV(await csvZip.file(base+"POLE.csv").async("string"));

      const kmz=await JSZip.loadAsync(await kmzFileInput.files[0].arrayBuffer());
      const kmlName=Object.keys(kmz.files).find(f=>f.endsWith(".kml"));
      const doc=new DOMParser().parseFromString(await kmz.file(kmlName).async("string"),"text/xml");

      patchFolder(doc,["DISTRIBUSI","HP","HOME"],indexBy(HOME.rows,"HOMEPASS_ID"),"HOME");
      patchFolder(doc,["DISTRIBUSI","HP","HOME-BIZ"],indexBy(HOME_BIZ.rows,"HOMEPASS_ID"),"HOME_BIZ");
      patchFolder(doc,["DISTRIBUSI","POLE"],indexBy(POLE.rows,"Pole ID (New)"),"POLE");

      kmz.file(kmlName,new XMLSerializer().serializeToString(doc));
      saveAs(await kmz.generateAsync({type:"blob"}),"KMZ_POPUP_FINAL.kmz");

      statusKmz.textContent="SELESAI ✔ POPUP IDENTIK GOOGLE EARTH";

    } catch(e){
      console.error(e);
      statusKmz.textContent="ERROR: "+e.message;
    }
  });

});
