/* Pop Up CSV Generator - client-side (no upload server)
   - Reads Excel (Master Pop Up)
   - Generates 6 CSV (HOME, HOME-BIZ, FDT, FAT, HOOK, POLE)
   - Output in ZIP
*/

const elFile = document.getElementById('file');
const elBtn = document.getElementById('btn');
const elArea = document.getElementById('areaName');
const elStatus = document.getElementById('status');

const elCHome = document.getElementById('cHome');
const elCHomeBiz = document.getElementById('cHomeBiz');
const elCFdt = document.getElementById('cFdt');
const elCFat = document.getElementById('cFat');
const elCHook = document.getElementById('cHook');
const elCPole = document.getElementById('cPole');

let lastResult = null;

const TPL = {
  HOME: ["HOMEPASS_ID","CLUSTER_NAME","PREFIX_ADDRESS","STREET_NAME","HOUSE_NUMBER","BLOCK","FLOOR","RT","RW","DISTRICT","SUB_DISTRICT","FDT_CODE","FAT_CODE","BUILDING_LATITUDE","BUILDING_LONGITUDE","Category BizPass","POST CODE","ADDRESS POLE / FAT","OV_UG","HOUSE_COMMENT_","BUILDING_NAME","TOWER","APTN","FIBER_NODE__HFC_","ID_Area","Clamp_Hook_ID","DEPLOYMENT_TYPE","NEED_SURVEY"],
  HOME_BIZ: ["HOMEPASS_ID","CLUSTER_NAME","PREFIX_ADDRESS","STREET_NAME","HOUSE_NUMBER","BLOCK","FLOOR","RT","RW","DISTRICT","SUB_DISTRICT","FDT_CODE","FAT_CODE","BUILDING_LATITUDE","BUILDING_LONGITUDE","Category BizPass","POST CODE","ADDRESS POLE / FAT","OV_UG","HOUSE_COMMENT_","BUILDING_NAME","TOWER","APTN","FIBER_NODE__HFC_","ID_Area","Clamp_Hook_ID","DEPLOYMENT_TYPE","NEED_SURVEY"],
  FDT: ["Pole ID (New)","Coordinate (Lat) NEW","Coordinate (Long) NEW","Pole Provider (New)","Pole Type","FAT ID/NETWORK ID"],
  FAT: ["Pole ID (New)","Coordinate (Lat) NEW","Coordinate (Long) NEW","Pole Provider (New)","Pole Type","FAT ID/NETWORK ID"],
  HOOK: ["Clamp_Hook_ID","Clamp_Hook_LATITUDE","Clamp_Hook_LONGITUDE"],
  POLE: ["Pole ID (New)","Coordinate (Lat) NEW","Coordinate (Long) NEW","Pole Provider (New)","Pole Type","LINE"],
};

function normStr(v){
  if (v === null || v === undefined) return "";
  return String(v).replace(/\u00A0/g,' ').trim();
}

function normalizeHomeType(v){
  const s = normStr(v).toUpperCase().replace(/\s+/g,'');
  if (!s) return "";
  if (s === "HOME") return "HOME";
  if (s === "HOMEBIZ" || s === "HOME-BIZ" || s === "HOME_BIZ" || s === "HOME BIZ") return "HOME-BIZ";
  // fallback: contains BIZ
  if (s.includes("BIZ")) return "HOME-BIZ";
  return "HOME";
}

function splitIds(raw){
  const s = normStr(raw);
  if (!s || s === "-" ) return [];
  // Split on & , ; / newline and also the word 'dan' / 'and'
  return s
    .split(/\s*(?:&|,|;|\n|\r|\/|\band\b|\bdan\b)\s*/i)
    .map(x => normStr(x))
    .filter(x => x && x !== "-");
}

function csvEscape(val){
  const s = normStr(val);
  if (s.includes(';') || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return '"' + s.replace(/"/g,'""') + '"';
  }
  return s;
}

function rowsToCsv(headers, rows){
  const lines = [];
  lines.push(headers.map(csvEscape).join(';'));
  for (const r of rows){
    const line = headers.map(h => csvEscape(r[h] ?? "")).join(';');
    lines.push(line);
  }
  return lines.join('\n');
}

function pick(row, headers){
  const out = {};
  for (const h of headers) out[h] = normStr(row[h]);
  return out;
}

function countSet(s){ return s ? s.size : 0; }

function setStatus(msg){ elStatus.textContent = msg || ""; }

function updatePreview(res){
  elCHome.textContent = res.homeRows.length;
  elCHomeBiz.textContent = res.homeBizRows.length;
  elCFdt.textContent = res.fdtRows.length;
  elCFat.textContent = res.fatRows.length;
  elCHook.textContent = res.hookRows.length;
  elCPole.textContent = res.poleRows.length;
}

function deriveAreaName(workbookName, firstRow){
  const manual = normStr(elArea.value);
  if (manual) return manual;

  const idArea = firstRow ? normStr(firstRow["ID_Area"]) : "";
  if (idArea) return idArea;

  // fallback to filename without extension
  const base = workbookName.replace(/\.[^.]+$/,'').trim();
  return base || "AREA";
}

function generateFromRows(allRows, workbookName){
  // HOME / HOME-BIZ by column "HOME/HOME-BIZ"
  const homeRows = [];
  const homeBizRows = [];

  for (const r of allRows){
    const ht = normalizeHomeType(r["HOME/HOME-BIZ"]);
    if (ht === "HOME-BIZ") homeBizRows.push(pick(r, TPL.HOME_BIZ));
    else if (ht === "HOME") homeRows.push(pick(r, TPL.HOME));
  }

  // HOOK unique by Clamp_Hook_ID (ignore '-')
  const hookMap = new Map();
  for (const r of allRows){
    const id = normStr(r["Clamp_Hook_ID"]);
    if (!id || id === "-") continue;
    if (!hookMap.has(id)){
      hookMap.set(id, pick(r, TPL.HOOK));
    }
  }
  const hookRows = Array.from(hookMap.values()).sort((a,b)=>normStr(a["Clamp_Hook_ID"]).localeCompare(normStr(b["Clamp_Hook_ID"]), undefined, {numeric:true}));

  // POLE unique by Pole ID (New)
  const poleMap = new Map();
  for (const r of allRows){
    const pid = normStr(r["Pole ID (New)"]);
    if (!pid) continue;
    if (!poleMap.has(pid)){
      poleMap.set(pid, pick(r, TPL.POLE));
    }
  }
  const poleRows = Array.from(poleMap.values()).sort((a,b)=>normStr(a["Pole ID (New)"]).localeCompare(normStr(b["Pole ID (New)"]), undefined, {numeric:true}));

  // FAT / FDT based on FAT ID/NETWORK ID tokens
  const fatRows = [];
  const fdtRows = [];
  const seenFat = new Set();
  const seenFdt = new Set();

  for (const r of allRows){
    const pid = normStr(r["Pole ID (New)"]);
    const base = pick(r, ["Pole ID (New)","Coordinate (Lat) NEW","Coordinate (Long) NEW","Pole Provider (New)","Pole Type","FAT ID/NETWORK ID"]);
    if (!pid) continue;

    const tokens = splitIds(r["FAT ID/NETWORK ID"]);
    for (const t of tokens){
      const rowOut = {...base, "FAT ID/NETWORK ID": t};
      if (t.toUpperCase().includes("S")) {
        const key = pid + "|" + t;
        if (!seenFat.has(key)){
          seenFat.add(key);
          fatRows.push(pick(rowOut, TPL.FAT));
        }
      } else {
        const key = pid + "|" + t;
        if (!seenFdt.has(key)){
          seenFdt.add(key);
          fdtRows.push(pick(rowOut, TPL.FDT));
        }
      }
    }
  }

  // Sort FAT / FDT by FAT ID/NETWORK ID then pole
  fatRows.sort((a,b)=> normStr(a["FAT ID/NETWORK ID"]).localeCompare(normStr(b["FAT ID/NETWORK ID"])) || normStr(a["Pole ID (New)"]).localeCompare(normStr(b["Pole ID (New)"]), undefined, {numeric:true}));
  fdtRows.sort((a,b)=> normStr(a["FAT ID/NETWORK ID"]).localeCompare(normStr(b["FAT ID/NETWORK ID"])) || normStr(a["Pole ID (New)"]).localeCompare(normStr(b["Pole ID (New)"]), undefined, {numeric:true}));

  const areaName = deriveAreaName(workbookName, allRows[0]);

  return { areaName, homeRows, homeBizRows, fdtRows, fatRows, hookRows, poleRows };
}

async function buildZip(res){
  const zip = new JSZip();
  const folder = zip.folder(res.areaName);

  folder.file("HOME.csv", rowsToCsv(TPL.HOME, res.homeRows));
  folder.file("HOME-BIZ.csv", rowsToCsv(TPL.HOME_BIZ, res.homeBizRows));
  folder.file("FDT.csv", rowsToCsv(TPL.FDT, res.fdtRows));
  folder.file("FAT.csv", rowsToCsv(TPL.FAT, res.fatRows));
  folder.file("HOOK.csv", rowsToCsv(TPL.HOOK, res.hookRows));
  folder.file("POLE.csv", rowsToCsv(TPL.POLE, res.poleRows));

  return await zip.generateAsync({ type: "blob" });
}

function readWorkbook(file){
  return new Promise((resolve,reject)=>{
    const reader = new FileReader();
    reader.onload = (e)=>{
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        resolve(wb);
      }catch(err){ reject(err); }
    };
    reader.onerror = ()=>reject(new Error("Gagal membaca file."));
    reader.readAsArrayBuffer(file);
  });
}

function sheetToRows(wb){
  const target = wb.Sheets["Master Data"] ? "Master Data" : wb.SheetNames[0];
  const ws = wb.Sheets[target];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
  return { sheetName: target, rows };
}

elFile.addEventListener('change', async ()=>{
  lastResult = null;
  elBtn.disabled = true;
  setStatus("");
  elCHome.textContent = elCHomeBiz.textContent = elCFdt.textContent = elCFat.textContent = elCHook.textContent = elCPole.textContent = "-";

  const f = elFile.files?.[0];
  if (!f) return;

  setStatus("Membaca Excel…");
  try{
    const wb = await readWorkbook(f);
    const { sheetName, rows } = sheetToRows(wb);
    if (!rows || rows.length === 0){
      setStatus("Sheet kosong. Pastikan ada data di sheet Master Data.");
      return;
    }
    const res = generateFromRows(rows, f.name);
    lastResult = res;
    updatePreview(res);
    elBtn.disabled = false;
    setStatus(`Siap generate. Sheet: ${sheetName}. Baris terbaca: ${rows.length}.`);
  }catch(err){
    console.error(err);
    setStatus("Gagal membaca Excel. Coba pastikan file .xlsx dan tidak corrupt.");
  }
});

elBtn.addEventListener('click', async ()=>{
  if (!lastResult) return;
  elBtn.disabled = true;
  setStatus("Membuat ZIP…");
  try{
    const blob = await buildZip(lastResult);
    const safeArea = lastResult.areaName.replace(/[^a-z0-9 _-]/gi,'_');
    saveAs(blob, `${safeArea}_popups.zip`);
    setStatus("Selesai. ZIP sudah di-download.");
  }catch(err){
    console.error(err);
    setStatus("Gagal membuat ZIP.");
  }finally{
    elBtn.disabled = false;
  }
});
