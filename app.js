document.addEventListener("DOMContentLoaded", () => {

  // =====================================================
  // ELEMENT (FIXED – SESUAI index.html)
  // =====================================================
  const $ = (id) => document.getElementById(id);

  // === DATA BASE → MASTER POP UP
  const fileDB        = $("dbFile");
  const btnMaster     = $("btnGenMaster");
  const statusMaster  = $("statusMaster");

  const areaInput     = $("dbArea");
  const districtIn    = $("dbDistrict");
  const subDistrictIn = $("dbSubDistrict");
  const postCodeIn    = $("dbPostCode");
  const fdtIn         = $("dbFdtCode");
  const idAreaIn      = $("dbIdArea");

  // === CSV GENERATOR (LAMA)
  const fileMasterCSV = $("file");
  const btnCSV        = $("btn");
  const statusCSV     = $("status");

  // === PATCH KMZ
  const kmzFileInput  = $("kmzFile");
  const csvZipInput   = $("csvZip");
  const btnPatchKmz   = $("btnPatchKmz");
  const statusKmz     = $("statusKmz");

  // =====================================================
  // HELPERS
  // =====================================================
  const norm = (s) => String(s || "").trim().toUpperCase();
  const raw  = (s) => String(s || "").trim();

  // =====================================================
  // 1️⃣ DATA BASE → MASTER POP UP
  // =====================================================
  btnMaster.addEventListener("click", async () => {
    if (!fileDB.files.length) {
      statusMaster.textContent = "ERROR: File DATA BASE belum dipilih";
      return;
    }

    statusMaster.textContent = "Memproses DATA BASE → MASTER POP UP...";

    try {
      const buf  = await fileDB.files[0].arrayBuffer();
      const wb   = XLSX.read(buf, { type: "array" });
      const ws   = wb.Sheets[wb.SheetNames[0]];
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

      XLSX.writeFile(wbOut, `MASTER_POP_UP_${areaInput.value || "AREA"}.xlsx`);

      statusMaster.textContent = "SELESAI ✔ MASTER POP UP TERBENTUK";

    } catch (e) {
      console.error(e);
      statusMaster.textContent = "ERROR: " + e.message;
    }
  });

});
