document.addEventListener("DOMContentLoaded", () => {

  // =====================================================
  // ELEMENT
  // =====================================================
  const $ = (id) => document.getElementById(id);

  const fileDB       = $("file");          // DATA BASE
  const btnMaster    = $("btnMaster");     // Generate MASTER POP UP
  const btnCSV       = $("btn");            // Generate CSV
  const status       = $("status");

  const areaInput    = $("areaName");
  const districtIn   = $("district");
  const subDistrictIn= $("subdistrict");
  const postCodeIn   = $("postcode");
  const fdtIn        = $("fdtcode");
  const idAreaIn     = $("idarea");

  const kmzFileInput = $("kmzFile");
  const csvZipInput  = $("csvZip");
  const btnPatchKmz  = $("btnPatchKmz");
  const statusKmz    = $("statusKmz");

  // =====================================================
  // HELPERS
  // =====================================================
  const norm = (s) => String(s || "").trim().toUpperCase();
  const raw  = (s) => String(s || "").trim();

  // =====================================================
  // ================== 1️⃣ DATA BASE → MASTER POP UP
  // =====================================================
  btnMaster?.addEventListener("click", async () => {

    status.textContent = "Memproses DATA BASE → MASTER POP UP...";

    try {
      const buf = await fileDB.files[0].arrayBuffer();
      const wb  = XLSX.read(buf, { type: "array" });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const rows= XLSX.utils.sheet_to_json(ws, { defval:"" });

      const HOME = [];
      const HOME_BIZ = [];
      const POLE = [];

      rows.forEach(r => {
        const folder = raw(r.FolderPath);
        const name   = raw(r.Name);

        // ========= HOME / HOME-BIZ =========
        if (folder.includes("/HOME")) {

          const isBiz = folder.includes("HOME-BIZ");
          const street = folder.split("/").slice(-2)[0];

          const blockMatch = name.match(/^([A-Z]+)\d+/);
          const BLOCK = blockMatch ? blockMatch[1] : "-";

          const row = {
            HOMEPASS_ID: "-",
            CLUSTER_NAME: areaInput.value,
            PREFIX_ADDRESS: "JL.",
            STREET_NAME: street,
            HOUSE_NUMBER: name,
            BLOCK,
            FLOOR: "-",
            RT: "-",
            RW: "-",
            DISTRICT: districtIn.value,
            SUB_DISTRICT: subDistrictIn.value,
            FDT_CODE: fdtIn.value,
            FAT_CODE: "FAT",
            BUILDING_LATITUDE: r.Y,
            BUILDING_LONGITUDE: r.X,
            "Category BizPass": "",
            "POST CODE": postCodeIn.value,
            "ADDRESS POLE / FAT": "-",
            OV_UG: "O",
            HOUSE_COMMENT_: "NEED SURVEY",
            BUILDING_NAME: "-",
            TOWER: "-",
            APTN: "-",
            FIBER_NODE__HFC_: "-",
            ID_Area: idAreaIn.value,
            Clamp_Hook_ID: "CLAIM_HOOK",
            DEPLOYMENT_TYPE: "FAT EXT",
            NEED_SURVEY: "YES"
          };

          isBiz ? HOME_BIZ.push(row) : HOME.push(row);
        }

        // ========= POLE =========
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

      XLSX.writeFile(wbOut, `MASTER_POP_UP_${areaInput.value}.xlsx`);
      status.textContent = "SELESAI ✔ MASTER POP UP TERBENTUK";

    } catch (e) {
      console.error(e);
      status.textContent = "ERROR: " + e.message;
    }
  });

  // =====================================================
  // ================== 2️⃣ CSV → KMZ (SCHEMA MODE)
  // =====================================================
  function parseCSV(text) {
    const [h, ...lines] = text.split(/\r?\n/).filter(Boolean);
    const headers = h.split(";").map(x => x.replace(/^"|"$/g,""));
    const rows = lines.map(l=>{
      const o={};
      l.split(";").forEach((v,i)=>o[headers[i]]=v.replace(/^"|"$/g,""));
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
  });

});
