"use strict";
/* =========================================================
   Meter IP Communication Dashboard
   - Parses ip_year_data-style workbooks (any number of day columns)
   - Dedup rule: for IPs with multiple rows, keep the row with the
     most TRUE days; if no row has any TRUE, keep the row with the
     most FALSE (i.e. most reported data); final fallback = first row.
   - Day cell states: 1 = TRUE (communicated), 0 = FALSE (silent),
     -1 = NULL / blank (no data — excluded from rate denominators)
   ========================================================= */

// ---------- state ----------
let DATA = null;      // { meters:[], dates:[], meta:{} }
let FILTER = { company: "", tsp: "", region: "", circle: "", po: "", range: "all" };
let tableSort = { key: "uptime", dir: 1 };
let tableLimit = 25;
let tableSearch = "";
let trendMode = "overall";
const tableModes = {}; // per-card chart/table toggle
let currentTab = "overview";
let weTableSort = { key: "gap", dir: -1 };
let weTableLimit = 25;
let weTableSearch = "";
let patternCategory = "Weekday-only";
let patTableSort = { key: "gap", dir: -1 };
let patTableLimit = 25;
let patTableSearch = "";
let drTableSort = { key: "uptime", dir: 1 };
let drTableLimit = 25;
let drTableSearch = "";
let searchedMeter = null;
let selectedMonths = new Set();
let MOSTATS = null;
let moTableSort = { key: "uptime", dir: 1 };
let moTableLimit = 25;
let moTableSearch = "";
let MO_PSTATS = null;
let moPatternCategory = "Weekday-only";
let moPatTableSort = { key: "gap", dir: -1 };
let moPatTableLimit = 25;
let moPatTableSearch = "";
let MO_RSTATS = null;
let moRhythmCategory = "Alternating";
let moRhTableSort = { key: "avgRunLen", dir: -1 };
let moRhTableLimit = 25;
let moRhTableSearch = "";
let lossTableSearch = "";
let lossRows = []; // last-rendered top-50, kept for CSV export
let SMDATA = null; // { sheetName, rows:[{kk,circle,region,cuscode,mserial,status,make,remarks}], statuses:[], totalRawRows, dupResolved }
let SM_FILTER = { region: "", circle: "", make: "" };
let smInsightStatus = "";
let smTableSort = { key: "kk", dir: 1 };
let smTableLimit = 25;
let smTableSearch = "";

const $ = (id) => document.getElementById(id);
const fmt = (n) => n.toLocaleString("en-IN");
const pct = (n, d = 1) => (n * 100).toFixed(d) + "%";

// ---------- upload handling ----------
const dropZone = $("dropZone"), fileInput = $("fileInput");

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.classList.add("dragover"); });
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
dropZone.addEventListener("drop", (e) => {
  e.preventDefault(); dropZone.classList.remove("dragover");
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", () => { if (fileInput.files.length) handleFile(fileInput.files[0]); });
$("resetBtn").addEventListener("click", () => {
  DATA = null;
  $("dashView").classList.add("hidden");
  $("uploadView").classList.remove("hidden");
  $("resetBtn").classList.add("hidden");
  $("parseError").classList.add("hidden");
  fileInput.value = "";
});

function setStatus(msg) {
  $("parseStatus").classList.remove("hidden");
  $("parseMsg").textContent = msg;
}
function failParse(msg) {
  $("parseStatus").classList.add("hidden");
  const el = $("parseError");
  el.textContent = msg;
  el.classList.remove("hidden");
}

function handleFile(file) {
  $("parseError").classList.add("hidden");
  setStatus(`Reading ${file.name} (${(file.size / 1048576).toFixed(1)} MB)…`);
  const reader = new FileReader();
  reader.onerror = () => failParse("Could not read the file from disk.");
  reader.onload = (e) => {
    setStatus("Parsing workbook — large files can take up to a minute…");
    // let the status paint before the heavy synchronous parse
    setTimeout(() => {
      try {
        parseWorkbook(e.target.result);
      } catch (err) {
        console.error(err);
        failParse("Parse failed: " + err.message);
      }
    }, 60);
  };
  reader.readAsArrayBuffer(file);
}

// ---------- parsing ----------
function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: "array", dense: true, cellDates: true });
  // prefer a sheet whose name looks like ip data, else first sheet
  let sheetName = wb.SheetNames.find((n) => /ip/i.test(n) && /data/i.test(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = ws["!data"];
  if (!rows || rows.length < 2) return failParse(`Sheet "${sheetName}" is empty or has no data rows.`);

  // --- header detection ---
  const header = rows[0].map((c) => (c ? c.v : null));
  const findCol = (re) => header.findIndex((h) => typeof h === "string" && re.test(h.trim()));
  const col = {
    ip: findCol(/^ip$/i),
    msn: findCol(/^msn$/i),
    tsp: findCol(/^tsp$/i),
    dup: findCol(/^dup/i),
    po: findCol(/^po/i),
    region: findCol(/^region$/i),
    circle: findCol(/^circle$/i),
  };
  if (col.ip === -1) col.ip = findCol(/^ip\b/i);
  if (col.ip === -1) return failParse('Could not find the "IP" column (second column in the expected layout).');

  // date columns: header is a Date (cellDates) or a date-like string/number
  const dateCols = [], dates = [];
  header.forEach((h, i) => {
    let d = null;
    if (h instanceof Date) d = h;
    else if (typeof h === "number" && h > 20000 && h < 80000) {
      const p = XLSX.SSF.parse_date_code(h);
      if (p) d = new Date(p.y, p.m - 1, p.d);
    } else if (typeof h === "string" && /^\d{4}-\d{2}-\d{2}/.test(h)) d = new Date(h);
    // normalize to local midnight — SheetJS "Date" header cells can carry a UTC-based
    // time-of-day (e.g. 05:30 IST) that otherwise breaks exact-time date comparisons
    if (d && !isNaN(d)) { dateCols.push(i); dates.push(new Date(d.getFullYear(), d.getMonth(), d.getDate())); }
  });
  if (dateCols.length === 0) return failParse("No day columns detected. Headers after the meta columns must be dates.");

  const nDays = dateCols.length;
  setStatus(`Processing ${fmt(rows.length - 1)} rows × ${nDays} day columns…`);

  // --- read rows ---
  const raw = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    // note: t === "e" is an Excel error cell (#N/A etc.) — its .v is a numeric error code, not data
    const cell = (c) => (c >= 0 && row[c] !== undefined && row[c] !== null && row[c].t !== "e" ? row[c].v : null);
    const ip = cell(col.ip);
    if (ip === null || ip === "") continue;
    const days = new Int8Array(nDays);
    let t = 0, f = 0, lastSeen = -1;
    for (let j = 0; j < nDays; j++) {
      const c = row[dateCols[j]];
      const v = c === undefined || c === null ? null : c.v;
      if (v === true || v === "TRUE" || v === "True" || v === "true") { days[j] = 1; t++; lastSeen = j; }
      else if (v === false || v === "FALSE" || v === "False" || v === "false") { days[j] = 0; f++; }
      else days[j] = -1; // NULL / blank / anything else = no data
    }
    const msn = str(cell(col.msn));
    raw.push({
      ip: String(ip),
      msn, msnPrefix: msnPrefixOf(msn),
      tsp: str(cell(col.tsp)), dupCol: cell(col.dup),
      po: str(cell(col.po)), region: str(cell(col.region)), circle: str(cell(col.circle)),
      days, t, f, lastSeen,
    });
  }
  if (raw.length === 0) return failParse("No usable data rows found (empty IP column?).");

  // --- dedup by IP: MERGE all duplicate rows into one, day by day ---
  // per-day precedence TRUE > FALSE > NULL: if any duplicate is TRUE that day -> TRUE,
  // else if any is FALSE -> FALSE, else NULL (no data). Every IP ends up as a single record.
  const rawByIp = new Map(); // ip -> all raw records (kept for the Search tab's duplicate view)
  for (const rec of raw) {
    let list = rawByIp.get(rec.ip);
    if (!list) { list = []; rawByIp.set(rec.ip, list); }
    list.push(rec);
  }
  const meters = [];
  for (const [ip, list] of rawByIp) {
    // company/companies = distinct MSN prefixes across this IP's raw rows (a swapped
    // meter can belong to more than one company over the year)
    const msnPrefixes = [...new Set(list.map((rec) => rec.msnPrefix))].sort();
    if (list.length === 1) { list[0].msnPrefixes = msnPrefixes; meters.push(list[0]); continue; }
    const days = new Int8Array(nDays);
    let t = 0, f = 0, lastSeen = -1;
    for (let j = 0; j < nDays; j++) {
      let v = -1; // NULL until proven otherwise
      for (const rec of list) {
        const d = rec.days[j];
        if (d === 1) { v = 1; break; }      // TRUE wins outright
        if (d === 0) v = 0;                 // FALSE beats NULL, keep scanning for a TRUE
      }
      days[j] = v;
      if (v === 1) { t++; lastSeen = j; }
      else if (v === 0) f++;
    }
    // meta fields come from the first row for this IP (identical across duplicates)
    const base = list[0];
    meters.push({ ip, msn: base.msn, msnPrefix: base.msnPrefix, msnPrefixes, tsp: base.tsp, dupCol: base.dupCol, po: base.po, region: base.region, circle: base.circle, days, t, f, lastSeen });
  }

  DATA = {
    meters, dates, rawByIp,
    meta: {
      fileRows: raw.length,
      unique: meters.length,
      dupResolved: raw.length - meters.length,
      dupIps: raw.length - meters.length === 0 ? 0 : countDupIps(raw),
      sheetName,
    },
  };

  // build filter dropdowns
  fillSelect("fCompany", meters.flatMap((m) => m.msnPrefixes || [m.msnPrefix]));
  fillSelect("fTsp", meters.map((m) => m.tsp));
  fillSelect("fRegion", meters.map((m) => m.region));
  fillSelect("fCircle", meters.map((m) => m.circle));
  fillSelect("fPo", meters.map((m) => m.po));
  FILTER = { company: "", tsp: "", region: "", circle: "", po: "", range: "all" };
  $("fRange").value = "all";

  // date-range tab: bound the two date pickers to the loaded data's span
  const isoFirst = isoDate(dates[0]), isoLast = isoDate(dates[dates.length - 1]);
  ["drStart", "drEnd"].forEach((id) => { $(id).min = isoFirst; $(id).max = isoLast; });
  $("drStart").value = isoFirst;
  $("drEnd").value = isoLast;

  // search tab: bound its own date pickers the same way, reset any previous search
  ["srStart", "srEnd"].forEach((id) => { $(id).min = isoFirst; $(id).max = isoLast; });
  $("srStart").value = isoFirst;
  $("srEnd").value = isoLast;
  searchedMeter = null;
  $("srIpInput").value = "";
  $("srResults").classList.add("hidden");
  $("srNotFound").classList.add("hidden");
  $("srSuggestions").classList.add("hidden");

  // monthly-analytics tab: bucket every loaded day into its calendar month
  const monthMap = new Map();
  dates.forEach((d, idx) => {
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    let e = monthMap.get(key);
    if (!e) { e = { key, label: d.toLocaleDateString("en-IN", { month: "short", year: "numeric" }), indices: [] }; monthMap.set(key, e); }
    e.indices.push(idx);
  });
  DATA.monthList = [...monthMap.values()];
  selectedMonths = new Set();
  $("moMonthChips").innerHTML = DATA.monthList.map((mo) =>
    `<button class="seg-btn month-chip" data-month="${mo.key}">${esc(mo.label)}</button>`
  ).join("");
  $("moMonthChips").querySelectorAll(".month-chip").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.dataset.month;
      if (selectedMonths.has(key)) selectedMonths.delete(key); else selectedMonths.add(key);
      btn.classList.toggle("active", selectedMonths.has(key));
      renderMonthly();
    });
  });
  moTableLimit = 25; moTableSearch = ""; $("moMeterSearch").value = "";
  renderMonthly();

  $("parseStatus").classList.add("hidden");
  $("uploadView").classList.add("hidden");
  $("dashView").classList.remove("hidden");
  $("resetBtn").classList.remove("hidden");

  const d0 = dates[0], d1 = dates[dates.length - 1];
  $("dedupNote").innerHTML =
    `Loaded <b>${fmt(raw.length)}</b> rows from sheet “${sheetName}” · <b>${fmt(meters.length)}</b> unique meters after duplicate resolution ` +
    `(<b>${fmt(raw.length - meters.length)}</b> duplicate rows across <b>${fmt(DATA.meta.dupIps)}</b> IPs merged — for each IP all rows are combined day-by-day, ` +
    `taking TRUE if any row was TRUE that day, else FALSE if any was FALSE, else no data) · ` +
    `Date range <b>${dstr(d0)} → ${dstr(d1)}</b> (${dates.length} days) · NULL cells are treated as “no data” and excluded from rate calculations.`;

  loadSmartMeters(wb);

  render();
}

// ---------- smartmeters sheet ----------
const SM_STATUS_RANK = { "Communicating": 0, "SIM Installation pending": 1, "DC": 2, ">1 Month Non Comm": 3 };
const SM_STATUS_COLORS = {
  "Communicating": "var(--good-text)",
  "SIM Installation pending": "#d68a1e",
  "DC": "var(--critical)",
  ">1 Month Non Comm": "#8a3fd0",
};
function smStatusColor(status) { return SM_STATUS_COLORS[status] || "var(--deemph)"; }

function parseSmartMetersSheet(wb) {
  const sheetName = wb.SheetNames.find((n) => /smart\s*meter/i.test(n));
  if (!sheetName) return null;
  const ws = wb.Sheets[sheetName];
  const rows = ws["!data"];
  if (!rows || rows.length < 2) return null;

  const header = rows[0].map((c) => (c ? c.v : null));
  const findCol = (re) => header.findIndex((h) => typeof h === "string" && re.test(h.trim()));
  const col = {
    circle: findCol(/^circle\s*name$/i),
    region: findCol(/^region$/i),
    cuscode: findCol(/^cuscode$/i),
    mserial: findCol(/^mserial$/i),
    status: findCol(/^comm\s*status$/i),
    kk: findCol(/^kk$/i),
    make: findCol(/^make$/i),
    remarks: findCol(/^remarks$/i),
  };
  if (col.circle === -1) col.circle = findCol(/circle/i);
  if (col.status === -1) col.status = findCol(/status/i);
  if (col.kk === -1 || col.status === -1) return null;

  const seenKk = new Set();
  const smRows = [];
  let dataRowCount = 0;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const cell = (c) => (c >= 0 && row[c] !== undefined && row[c] !== null && row[c].t !== "e" ? row[c].v : null);
    const kk = str(cell(col.kk));
    if (kk === "Unknown" && str(cell(col.status)) === "Unknown") continue; // fully blank row
    dataRowCount++;
    if (kk !== "Unknown") {
      if (seenKk.has(kk)) continue; // de-dupe by MSN (kk) — a handful of rows repeat the same meter
      seenKk.add(kk);
    }
    smRows.push({
      kk, circle: str(cell(col.circle)), region: str(cell(col.region)),
      cuscode: str(cell(col.cuscode)), mserial: str(cell(col.mserial)),
      status: str(cell(col.status)), make: str(cell(col.make)), remarks: str(cell(col.remarks)),
    });
  }
  if (smRows.length === 0) return null;

  const statuses = [...new Set(smRows.map((r) => r.status))]
    .sort((a, b) => (SM_STATUS_RANK[a] ?? 99) - (SM_STATUS_RANK[b] ?? 99) || a.localeCompare(b));

  return { sheetName, rows: smRows, statuses, totalRawRows: dataRowCount, dupResolved: dataRowCount - smRows.length };
}

function loadSmartMeters(wb) {
  SMDATA = parseSmartMetersSheet(wb);
  SM_FILTER = { region: "", circle: "", make: "" };
  $("fSmRegion").value = ""; $("fSmCircle").value = ""; $("fSmMake").value = "";
  smTableSort = { key: "kk", dir: 1 }; smTableLimit = 25; smTableSearch = ""; $("smMeterSearch").value = "";

  if (!SMDATA) {
    $("smNoSheet").classList.remove("hidden");
    $("smContent").classList.add("hidden");
    return;
  }
  $("smNoSheet").classList.add("hidden");
  $("smContent").classList.remove("hidden");

  fillSelect("fSmRegion", SMDATA.rows.map((r) => r.region));
  fillSelect("fSmCircle", SMDATA.rows.map((r) => r.circle));
  fillSelect("fSmMake", SMDATA.rows.map((r) => r.make));

  smInsightStatus = SMDATA.statuses[0] || "";
  $("smInsightStatus").innerHTML = SMDATA.statuses.map((s) => `<option value="${esc(s)}">${esc(s)}</option>`).join("");

  $("smNote").innerHTML =
    `Loaded <b>${fmt(SMDATA.totalRawRows)}</b> rows from sheet “${esc(SMDATA.sheetName)}” · <b>${fmt(SMDATA.rows.length)}</b> unique smart meters by MSN (kk)` +
    `${SMDATA.dupResolved > 0 ? ` (<b>${fmt(SMDATA.dupResolved)}</b> duplicate kk rows dropped)` : ""}.`;
}

function countDupIps(raw) {
  const c = new Map();
  raw.forEach((r) => c.set(r.ip, (c.get(r.ip) || 0) + 1));
  let n = 0; c.forEach((v) => { if (v > 1) n++; });
  return n;
}
function str(v) {
  if (v === null || v === undefined) return "Unknown";
  const s = String(v).trim();
  return s === "" || s === "#N/A" || s === "NULL" ? "Unknown" : s;
}
// MSN = meter serial number. The first 3 letters identify the meter make/vendor
// (known set: ISE, CPS, JPM, VTK). Used to show which physical meter was installed
// on this IP over time — a meter swap shows up as a new prefix owning a later date range.
const MSN_COLORS = { ISE: "#2a78d6", CPS: "#e08a1e", JPM: "#7c4dd6", VTK: "#1baf7a" };
function msnPrefixOf(msn) {
  if (!msn || msn === "Unknown") return "Unknown";
  const p = String(msn).trim().slice(0, 3).toUpperCase();
  return /^[A-Z]{3}$/.test(p) ? p : "Unknown";
}
function msnColor(prefix) { return MSN_COLORS[prefix] || (prefix === "Unknown" ? "var(--grid)" : "#8a8a80"); }
// base hex for a prefix (Unknown falls back to a concrete grey so shade variants work)
function msnBaseHex(prefix) { return MSN_COLORS[prefix] || (prefix === "Unknown" ? "#9a9a90" : "#8a8a80"); }

// --- tiny hex<->hsl helpers, used to derive distinguishable shades of one base colour ---
function hexToHsl(hex) {
  const n = parseInt(hex.slice(1), 16);
  let r = ((n >> 16) & 255) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2;
  let h = 0, s = 0;
  if (mx !== mn) {
    const d = mx - mn;
    s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  return [h * 360, s * 100, l * 100];
}
function hslToCss(h, s, l) { return `hsl(${h.toFixed(0)} ${Math.max(0, Math.min(100, s)).toFixed(0)}% ${Math.max(0, Math.min(100, l)).toFixed(0)}%)`; }

// Build a { msn -> colour } map for ONE IP's records. Different make (prefix) => the
// prefix's own base colour. When several distinct MSNs share a prefix (e.g. two ISE
// meters on one IP), spread them across lightness + a small hue shift so each lifecycle
// block is visually distinct while still reading as the same colour family.
function buildMsnColorMap(records) {
  const membersByPrefix = new Map(); // prefix -> [distinct msn in first-seen order]
  for (const rec of records) {
    const list = membersByPrefix.get(rec.msnPrefix) || (membersByPrefix.set(rec.msnPrefix, []).get(rec.msnPrefix));
    if (!list.includes(rec.msn)) list.push(rec.msn);
  }
  const map = new Map();
  for (const [prefix, msns] of membersByPrefix) {
    if (msns.length === 1) { map.set(msns[0], msnColor(prefix)); continue; }
    const [h, s, l] = hexToHsl(msnBaseHex(prefix));
    const n = msns.length, spread = 34; // total lightness spread across the group
    msns.forEach((msn, i) => {
      const t = n === 1 ? 0 : i / (n - 1) - 0.5; // -0.5 .. +0.5
      map.set(msn, hslToCss(h + t * 24, s, l + t * spread)); // shift hue slightly + lightness
    });
  }
  return map;
}

function dstr(d) { return d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" }); }
function dstrShort(d) { return d.toLocaleDateString("en-IN", { day: "2-digit", month: "short" }); }
function isoDate(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; }

function fillSelect(id, values) {
  const uniq = [...new Set(values)].sort();
  const sel = $(id);
  sel.innerHTML = `<option value="">All</option>` + uniq.map((v) => `<option>${esc(v)}</option>`).join("");
}
function esc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }

// ---------- filters ----------
["fCompany", "fTsp", "fRegion", "fCircle", "fPo", "fRange"].forEach((id) => {
  $(id).addEventListener("change", () => {
    FILTER = {
      company: $("fCompany").value, tsp: $("fTsp").value, region: $("fRegion").value,
      circle: $("fCircle").value, po: $("fPo").value, range: $("fRange").value,
    };
    tableLimit = 25;
    render();
  });
});
$("fSatWeekend").addEventListener("change", () => {
  satIsWeekend = $("fSatWeekend").value === "weekend";
  render();
});

function currentWindow() {
  const n = DATA.dates.length;
  if (FILTER.range === "all") return [0, n - 1];
  const days = Math.min(parseInt(FILTER.range, 10), n);
  return [n - days, n - 1];
}
function filteredMeters() {
  return DATA.meters.filter((m) =>
    (!FILTER.company || (m.msnPrefixes || [m.msnPrefix]).includes(FILTER.company)) &&
    (!FILTER.tsp || m.tsp === FILTER.tsp) &&
    (!FILTER.region || m.region === FILTER.region) &&
    (!FILTER.circle || m.circle === FILTER.circle) &&
    (!FILTER.po || m.po === FILTER.po)
  );
}

// ---------- aggregation ----------
function computeStats(meters, i0, i1) {
  const nDays = i1 - i0 + 1;
  const dayTrue = new Float64Array(nDays), dayFalse = new Float64Array(nDays), dayNull = new Float64Array(nDays);
  const perMeter = [];
  const byGroup = { tsp: new Map(), region: new Map(), circle: new Map(), po: new Map() };
  const dayTspTrue = new Map(), dayTspData = new Map(); // tsp -> Float64Array

  let totT = 0, totF = 0, totN = 0, never = 0, silent7 = 0, silent30 = 0, full = 0;
  const buckets = [0, 0, 0, 0, 0, 0, 0, 0]; // 0%, >0-25, 25-50, 50-75, 75-85, 85-<95, 95-<100, 100

  for (const m of meters) {
    let t = 0, f = 0, nl = 0, lastSeen = -1;
    const d = m.days;
    for (let j = i0; j <= i1; j++) {
      const v = d[j], k = j - i0;
      if (v === 1) { t++; dayTrue[k]++; lastSeen = j; }
      else if (v === 0) { f++; dayFalse[k]++; }
      else { nl++; dayNull[k]++; }
    }
    totT += t; totF += f; totN += nl;
    const dataDays = t + f;
    const uptime = dataDays > 0 ? t / dataDays : -1; // -1 = no data at all in window
    if (dataDays > 0 && t === 0) never++;
    if (uptime === 1) full++;
    // silent streak from window end (only meaningful if meter has data)
    let bucket = -1;
    if (dataDays > 0) {
      const gap = i1 - lastSeen; // lastSeen=-1 -> huge gap
      const silentDays = lastSeen === -1 ? nDays : gap;
      if (silentDays >= 7 && nDays >= 7) silent7++;
      if (silentDays >= 30 && nDays >= 30) silent30++;
      if (uptime === 0) bucket = 0;
      else if (uptime === 1) bucket = 7;
      else if (uptime <= 0.25) bucket = 1;
      else if (uptime <= 0.5) bucket = 2;
      else if (uptime <= 0.75) bucket = 3;
      else if (uptime <= 0.85) bucket = 4;
      else if (uptime < 0.95) bucket = 5;
      else bucket = 6;
      buckets[bucket]++;
    }
    perMeter.push({ m, t, f, nl, uptime, lastSeen, bucket, silentDays: dataDays > 0 ? (lastSeen === -1 ? nDays : i1 - lastSeen) : -1 });

    for (const g of ["tsp", "region", "circle", "po"]) {
      const key = m[g];
      let e = byGroup[g].get(key);
      if (!e) { e = { t: 0, f: 0, n: 0, meters: 0 }; byGroup[g].set(key, e); }
      e.t += t; e.f += f; e.n += nl; e.meters++;
    }
    // per-day by TSP (for the trend split)
    let tt = dayTspTrue.get(m.tsp), td = dayTspData.get(m.tsp);
    if (!tt) { tt = new Float64Array(nDays); td = new Float64Array(nDays); dayTspTrue.set(m.tsp, tt); dayTspData.set(m.tsp, td); }
    for (let j = i0; j <= i1; j++) {
      const v = d[j], k = j - i0;
      if (v === 1) { tt[k]++; td[k]++; } else if (v === 0) td[k]++;
    }
  }
  return { dayTrue, dayFalse, dayNull, perMeter, byGroup, dayTspTrue, dayTspData, totT, totF, totN, never, silent7, silent30, full, buckets, i0, i1, nDays, meterCount: meters.length };
}

// ---------- weekend vs weekday aggregation ----------
// JS Date.getDay(): 0=Sun..6=Sat. Remap so 0=Mon..6=Sun.
const DOW_LABELS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
function dowIndex(d) { return (d.getDay() + 6) % 7; }
// global setting: whether Saturday (dow index 5) counts as weekend (default) or weekday.
// Sunday (6) is always weekend. Every weekday/weekend computation routes through this.
let satIsWeekend = true;
function isWeekendDow(dow) { return satIsWeekend ? dow >= 5 : dow === 6; }

function computeWeekendStats(meters, i0, i1) {
  const dowT = new Array(7).fill(0), dowF = new Array(7).fill(0);
  let weekdayT = 0, weekdayF = 0, weekendT = 0, weekendF = 0;
  const byTsp = new Map(), byRegion = new Map(), byCircle = new Map(), byPo = new Map();
  const perMeter = [];

  // precompute which window index is weekend
  const isWeekend = new Array(i1 - i0 + 1);
  for (let j = i0; j <= i1; j++) isWeekend[j - i0] = isWeekendDow(dowIndex(DATA.dates[j]));

  for (const m of meters) {
    let wdT = 0, wdF = 0, weT = 0, weF = 0;
    const d = m.days;
    for (let j = i0; j <= i1; j++) {
      const v = d[j];
      if (v === -1) continue;
      const k = j - i0, dow = dowIndex(DATA.dates[j]);
      if (v === 1) { dowT[dow]++; if (isWeekend[k]) weT++; else wdT++; }
      else { dowF[dow]++; if (isWeekend[k]) weF++; else wdF++; }
    }
    weekdayT += wdT; weekdayF += wdF; weekendT += weT; weekendF += weF;
    const wdRate = wdT + wdF > 0 ? wdT / (wdT + wdF) : null;
    const weRate = weT + weF > 0 ? weT / (weT + weF) : null;
    const gap = wdRate !== null && weRate !== null ? wdRate - weRate : null;
    const uptime = wdT + weT + wdF + weF > 0 ? (wdT + weT) / (wdT + wdF + weT + weF) : null;
    perMeter.push({ m, wdT, wdF, weT, weF, wdRate, weRate, gap, uptime });

    for (const [map, key] of [[byTsp, m.tsp], [byRegion, m.region], [byCircle, m.circle], [byPo, m.po]]) {
      let e = map.get(key);
      if (!e) { e = { wdT: 0, wdF: 0, weT: 0, weF: 0, meters: 0 }; map.set(key, e); }
      e.wdT += wdT; e.wdF += wdF; e.weT += weT; e.weF += weF; e.meters++;
    }
  }
  const dowRate = dowT.map((t, i) => (t + dowF[i] > 0 ? t / (t + dowF[i]) : null));
  return { dowT, dowF, dowRate, weekdayT, weekdayF, weekendT, weekendF, byTsp, byRegion, byCircle, byPo, perMeter };
}

// same weekday/weekend split as computeWeekendStats(), but driven by an arbitrary index list
// (used by Monthly Analytics, where selected months need not be contiguous)
function computeWeekendStatsForIndices(meters, indices) {
  const nDays = indices.length;
  const isWeekend = indices.map((j) => isWeekendDow(dowIndex(DATA.dates[j])));
  const byTsp = new Map(), byRegion = new Map(), byCircle = new Map(), byPo = new Map();
  const perMeter = [];
  for (const m of meters) {
    let wdT = 0, wdF = 0, weT = 0, weF = 0;
    const d = m.days;
    for (let k = 0; k < nDays; k++) {
      const v = d[indices[k]];
      if (v === -1) continue;
      if (v === 1) { if (isWeekend[k]) weT++; else wdT++; }
      else { if (isWeekend[k]) weF++; else wdF++; }
    }
    const wdRate = wdT + wdF > 0 ? wdT / (wdT + wdF) : null;
    const weRate = weT + weF > 0 ? weT / (weT + weF) : null;
    const gap = wdRate !== null && weRate !== null ? wdRate - weRate : null;
    const uptime = wdT + weT + wdF + weF > 0 ? (wdT + weT) / (wdT + wdF + weT + weF) : null;
    perMeter.push({ m, wdT, wdF, weT, weF, wdRate, weRate, gap, uptime });
    for (const [map, key] of [[byTsp, m.tsp], [byRegion, m.region], [byCircle, m.circle], [byPo, m.po]]) {
      let e = map.get(key);
      if (!e) { e = { wdT: 0, wdF: 0, weT: 0, weF: 0, meters: 0 }; map.set(key, e); }
      e.wdT += wdT; e.wdF += wdF; e.weT += weT; e.weF += weF; e.meters++;
    }
  }
  return { byTsp, byRegion, byCircle, byPo, perMeter };
}

// ---------- tabs ----------
const TAB_PANEL = { overview: "tabOverview", weekend: "tabWeekend", patterns: "tabPatterns", daterange: "tabDateRange", search: "tabSearch", monthly: "tabMonthly", lossrank: "tabLossRank", smartmeters: "tabSmartMeters" };
const TAB_RENDER = {
  weekend: () => renderWeekend(), patterns: () => renderPatterns(), daterange: () => renderDateRange(),
  search: () => { if (searchedMeter) renderSearchTab(); },
  monthly: () => renderMonthly(),
  lossrank: () => renderLossTab(),
  smartmeters: () => renderSmartMeters(),
};
document.querySelectorAll(".tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    currentTab = btn.dataset.tab;
    Object.values(TAB_PANEL).forEach((id) => $(id).classList.add("hidden"));
    $(TAB_PANEL[currentTab]).classList.remove("hidden");
    if (TAB_RENDER[currentTab]) TAB_RENDER[currentTab]();
  });
});

// ---------- render ----------
let STATS = null;
let WSTATS = null;

function render() {
  const [i0, i1] = currentWindow();
  const meters = filteredMeters();
  STATS = computeStats(meters, i0, i1);
  WSTATS = computeWeekendStats(meters, i0, i1);
  $("filterMeta").textContent =
    `${fmt(meters.length)} meters in view · ${dstr(DATA.dates[i0])} → ${dstr(DATA.dates[i1])}`;
  renderKpis();
  renderTrend();
  renderTspChart();
  renderHealth();
  renderGroupBar("region", "chartRegion", 99);
  renderGroupBar("circle", "chartCircle", 15);
  renderGroupBar("po", "chartPo", 99);
  renderMeterTable();
  if (TAB_RENDER[currentTab]) TAB_RENDER[currentTab]();
}

// Revenue-loss model: each SIM costs SIM_COST regardless of comm rate. A meter is
// credited by its EXACT communication rate (uptime = TRUE ÷ (TRUE+FALSE)) — e.g. a
// meter at 17% comm realizes 0.17 of its SIM's worth and loses the other 0.83. So
// every IP carries its own precise weight in the fleet-level loss (no bucket rounding).
// A meter with no data at all in the window realizes nothing → full loss.
const SIM_COST = 19;
function realizedFraction(r) { return r.uptime >= 0 ? r.uptime : 0; } // uptime -1 = no data in window
function computeLoss(statsSrc) {
  let x = 0;
  for (const r of statsSrc.perMeter) x += realizedFraction(r) * SIM_COST;
  const total = statsSrc.meterCount * SIM_COST;
  const lossAmount = total - x;
  const lossPct = total > 0 ? (lossAmount / total) * 100 : 0;
  return { x, total, lossAmount, lossPct };
}
function inr(n) { return "₹" + Math.round(n).toLocaleString("en-IN"); }

// per-meter loss, using the same exact-rate model as computeLoss()
function computeMeterLoss(r) {
  const fraction = realizedFraction(r);
  const lossAmount = (1 - fraction) * SIM_COST;
  const lossPct = (1 - fraction) * 100;
  return { lossAmount, lossPct, fraction };
}

function renderKpisInto(s, rowId, improvement) {
  const rate = s.totT + s.totF > 0 ? s.totT / (s.totT + s.totF) : 0;
  const noData = s.totT + s.totF + s.totN > 0 ? s.totN / (s.totT + s.totF + s.totN) : 0;
  const loss = computeLoss(s);
  const tiles = [
    tile("Overall communication rate", pct(rate), `TRUE ÷ (TRUE+FALSE) over the selected window`, true, "all"),
    tile("Meters in view", fmt(s.meterCount), `${fmt(DATA.meta.unique)} unique in file · ${fmt(DATA.meta.dupResolved)} duplicate rows resolved`, undefined, "all"),
    tile("Never communicated", fmt(s.never), "have data but zero TRUE days in window", s.never > 0 ? "bad" : "good", "never"),
    tile("100% uptime meters", fmt(s.full), "TRUE on every day it reported data (NULL gaps allowed)", "good", "full"),
    tile("Silent ≥ 30 days", s.nDays >= 30 ? fmt(s.silent30) : "—", "no TRUE in the last 30 days of window", s.silent30 > 0 ? "bad" : "good", s.nDays >= 30 ? "silent30" : null),
    tile("No-data share", pct(noData), "NULL cells ÷ all cells (provisioning gaps)", undefined, "all"),
    tile("Revenue loss", `${inr(loss.lossAmount)} (${loss.lossPct.toFixed(1)}%)`, `SIM @ ₹${SIM_COST} · exact-rate weighted vs. ${inr(loss.total)} at 100% comm`, loss.lossPct > 0 ? "bad" : "good", "loss"),
  ];
  if (improvement) {
    const { firstRate, secondRate, deltaPp } = improvement;
    const hasSplit = deltaPp !== null;
    tiles.push(tile(
      "Improvement (1st → 2nd half)",
      hasSplit ? (deltaPp >= 0 ? "+" : "") + deltaPp.toFixed(1) + " pp" : "—",
      hasSplit ? `${pct(firstRate)} → ${pct(secondRate)} comm rate, split at the window's midpoint` : "range too short to split into two halves",
      hasSplit ? (deltaPp >= 0.5 ? "good" : deltaPp <= -0.5 ? "bad" : undefined) : undefined,
      hasSplit ? "improvement" : null
    ));
  }
  $(rowId).innerHTML = tiles.join("");
}
function renderKpis() { renderKpisInto(STATS, "kpiRow"); }

function kpiMeterSubset(s, key) {
  switch (key) {
    case "never": return s.perMeter.filter((r) => r.uptime === 0);
    case "full": return s.perMeter.filter((r) => r.uptime === 1);
    case "silent7": return s.perMeter.filter((r) => r.silentDays >= 7);
    case "silent30": return s.perMeter.filter((r) => r.silentDays >= 30);
    default: return s.perMeter; // "all"
  }
}
const KPI_FILENAMES = {
  never: "never_communicated.csv", full: "full_uptime_meters.csv",
  silent7: "silent_7plus_days.csv", silent30: "silent_30plus_days.csv", all: "meters_in_view.csv",
};
// per-bucket loss breakdown, summed from each meter's EXACT rate (not the bucket's
// upper bound) — shared by the CSV export and the "loss by bucket" chart
function computeLossByBucket(statsSrc) {
  const byBucket = BUCKET_LABELS.map(() => ({ meters: 0, loss: 0 }));
  const noData = { meters: 0, loss: 0 };
  for (const r of statsSrc.perMeter) {
    const l = computeMeterLoss(r);
    if (r.bucket === -1) { noData.meters++; noData.loss += l.lossAmount; }
    else { byBucket[r.bucket].meters++; byBucket[r.bucket].loss += l.lossAmount; }
  }
  return { byBucket, noData };
}
function exportLossCsv(statsSrc, filenamePrefix) {
  const loss = computeLoss(statsSrc);
  const { byBucket, noData } = computeLossByBucket(statsSrc);
  const rows = BUCKET_LABELS.map((label, i) => [label, byBucket[i].meters, byBucket[i].loss.toFixed(2)]);
  rows.push(["No data (never reported)", noData.meters, noData.loss.toFixed(2)]);
  rows.push(["TOTAL realized value", "", loss.x.toFixed(2)]);
  rows.push(["Value at 100% comm", statsSrc.meterCount, loss.total.toFixed(2)]);
  rows.push(["Loss amount", "", loss.lossAmount.toFixed(2)]);
  rows.push(["Loss %", "", loss.lossPct.toFixed(2)]);
  downloadCsv(`${filenamePrefix}loss_breakdown.csv`, ["Uptime bucket", "Meters", "Loss (₹)"], rows);
}
wireKpiRow("kpiRow", (key) => {
  if (!STATS) return;
  if (key === "loss") return exportLossCsv(STATS, "");
  downloadMeterCsv(KPI_FILENAMES[key] || "meters_in_view.csv", kpiMeterSubset(STATS, key));
});
wireKpiRow("drKpiRow", (key) => {
  if (!DRSTATS) return;
  if (key === "improvement") return exportImprovementCsv();
  if (key === "loss") return exportLossCsv(DRSTATS, "daterange_");
  downloadMeterCsv(`daterange_${KPI_FILENAMES[key] || "meters_in_view.csv"}`, kpiMeterSubset(DRSTATS, key));
});
wireKpiRow("moKpiRow", (key) => {
  if (!MOSTATS) return;
  if (key === "loss") return exportLossCsv(MOSTATS, "monthly_");
  downloadMeterCsv(`monthly_${KPI_FILENAMES[key] || "meters_in_view.csv"}`, kpiMeterSubset(MOSTATS, key));
});

// ==========================================================
// Revenue Loss tab — reuses the same STATS (filteredMeters() + currentWindow())
// as Overview, so it inherits every slicer (Company/TSP/Region/Circle/PO/date range)
// ==========================================================
function renderLossTab() {
  if (!STATS) return;
  const loss = computeLoss(STATS);
  $("lossKpiRow").innerHTML = [
    tile("Revenue loss", `${inr(loss.lossAmount)} (${loss.lossPct.toFixed(1)}%)`, `SIM @ ₹${SIM_COST} · exact-rate weighted vs. ${inr(loss.total)} at 100% comm`, loss.lossPct > 0 ? "bad" : "good", "loss"),
    tile("Meters in view", fmt(STATS.meterCount), "reflects the filters above"),
    tile("At-risk value", inr(loss.total), `${fmt(STATS.meterCount)} meters × ₹${SIM_COST}`),
  ].join("");
  renderLossByBucket(STATS, "chartLossBucket", "lossbucket");
  renderLossTable();
  renderLossTrend();
}

// ---------- Revenue loss trend: one point per calendar month, Jul 2025 -> Jul 2026 ----------
const LOSS_TREND_START_KEY = "2025-07", LOSS_TREND_END_KEY = "2026-07";
function lossTrendPoints() {
  const meters = filteredMeters(); // company/tsp/region/circle/po only — deliberately ignores the Date range dropdown
  const months = DATA.monthList.filter((mo) => mo.key >= LOSS_TREND_START_KEY && mo.key <= LOSS_TREND_END_KEY);
  return months.map((mo) => {
    const stats = computeStatsForIndices(meters, mo.indices);
    return { label: mo.label, key: mo.key, loss: computeLoss(stats), meterCount: stats.meterCount };
  });
}
function renderLossTrend() {
  if (!DATA) return;
  const points = lossTrendPoints();
  if (tableModes["losstrend"]) {
    $("chartLossTrend").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Month</th><th class="num">Loss %</th><th class="num">Loss amount</th><th class="num">Meters</th></tr></thead>
      <tbody>${points.map((p) => `<tr><td>${esc(p.label)}</td><td class="num">${p.loss.lossPct.toFixed(1)}%</td><td class="num">${inr(p.loss.lossAmount)}</td><td class="num">${fmt(p.meterCount)}</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  if (points.length === 0) {
    $("chartLossTrend").innerHTML = `<div class="note">No months between Jul 2025 and Jul 2026 were found in the uploaded file.</div>`;
    return;
  }

  const W = Math.max(680, Math.min(1300, $("chartLossTrend").clientWidth || 900));
  const H = 260, mL = 50, mR = 16, mT = 12, mB = 40;
  const iw = W - mL - mR, ih = H - mT - mB;
  const n = points.length;
  const x = (k) => mL + (n === 1 ? iw / 2 : (k / (n - 1)) * iw);
  const ticks = niceTicks(Math.max(...points.map((p) => p.loss.lossPct), 1));
  const yMax = ticks[ticks.length - 1];
  const y = (v) => mT + (1 - v / yMax) * ih;

  let g = "";
  for (const t of ticks) {
    g += `<line x1="${mL}" y1="${y(t)}" x2="${W - mR}" y2="${y(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${y(t) + 4}" text-anchor="end">${fmt(t)}%</text>`;
  }
  points.forEach((p, k) => { g += `<text class="axis-text" x="${x(k)}" y="${H - 10}" text-anchor="middle">${esc(p.label)}</text>`; });
  g += `<line x1="${mL}" y1="${y(0)}" x2="${W - mR}" y2="${y(0)}" stroke="var(--baseline)" stroke-width="1"/>`;

  let d = "";
  points.forEach((p, k) => { d += (k === 0 ? "M" : "L") + x(k).toFixed(1) + " " + y(p.loss.lossPct).toFixed(1); });
  let area = "M" + x(0).toFixed(1) + " " + y(0);
  points.forEach((p, k) => { area += "L" + x(k).toFixed(1) + " " + y(p.loss.lossPct).toFixed(1); });
  area += "L" + x(n - 1).toFixed(1) + " " + y(0) + "Z";

  let dots = "";
  points.forEach((p, k) => {
    dots += `<circle cx="${x(k).toFixed(1)}" cy="${y(p.loss.lossPct).toFixed(1)}" r="4.5" fill="var(--series-1)" stroke="var(--surface-1)" stroke-width="2" class="lossPt" data-i="${k}"/>`;
  });

  $("chartLossTrend").innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">
    ${g}<path d="${area}" fill="var(--series-area)"/><path d="${d}" fill="none" stroke="var(--series-1)" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>${dots}
  </svg>`;

  $("chartLossTrend").querySelectorAll(".lossPt").forEach((el) => {
    const i = +el.dataset.i;
    el.style.cursor = "pointer";
    el.addEventListener("mousemove", (e) => {
      const p = points[i];
      showTip(`<div class="tt-title">${esc(p.label)}</div>
        <div class="tt-row"><span>Loss</span><b>${inr(p.loss.lossAmount)} (${p.loss.lossPct.toFixed(1)}%)</b></div>
        <div class="tt-row"><span>Meters</span><b>${fmt(p.meterCount)}</b></div>`, e.clientX, e.clientY);
    });
    el.addEventListener("mouseleave", hideTip);
  });
}
wireKpiRow("lossKpiRow", (key) => {
  if (!STATS) return;
  if (key === "loss") exportLossCsv(STATS, "revenue_loss_");
});

function lossTableRows() {
  let rows = STATS.perMeter
    .map((r) => ({ r, loss: computeMeterLoss(r) }))
    .sort((a, b) => b.loss.lossAmount - a.loss.lossAmount);
  if (lossTableSearch) rows = rows.filter(({ r }) => r.m.ip.toLowerCase().includes(lossTableSearch));
  return rows;
}
function renderLossTable() {
  lossRows = lossTableRows().slice(0, 50);
  $("lossTableBody").innerHTML = lossRows.map(({ r, loss }, i) => `<tr>
    <td>${i + 1}</td><td>${esc(r.m.ip)}</td><td>${esc(r.m.tsp)}</td><td>${esc(r.m.region)}</td><td>${esc(r.m.circle)}</td><td>${esc(r.m.po)}</td>
    <td class="num">${r.uptime < 0 ? "—" : pct(r.uptime)}</td>
    <td class="num pct-bad">${loss.lossPct.toFixed(1)}%</td>
    <td class="num">${inr(loss.lossAmount)}</td>
  </tr>`).join("");
  $("lossTableCount").textContent = `Showing top ${fmt(lossRows.length)} of ${fmt(STATS.meterCount)} meters in view`;
}
$("lossMeterSearch").addEventListener("input", (e) => { lossTableSearch = e.target.value.trim().toLowerCase(); renderLossTable(); });
$("lossExportCsv").addEventListener("click", () => {
  const rows = lossRows.map(({ r, loss }) => [
    r.m.ip, r.m.tsp, r.m.region, r.m.circle, r.m.po,
    r.uptime < 0 ? "" : (r.uptime * 100).toFixed(2), loss.lossPct.toFixed(2), loss.lossAmount.toFixed(2),
  ]);
  downloadCsv("top_revenue_loss.csv", ["IP", "TSP", "Region", "Circle", "PO", "Uptime %", "Loss %", "Loss amount (₹)"], rows);
});

function tile(label, value, note, mood, key) {
  const hero = mood === true;
  const cls = mood === "bad" ? "bad" : mood === "good" ? "good" : "";
  const clickAttr = key ? ` data-kpi="${key}" tabindex="0" role="button"` : "";
  return `<div class="tile ${hero ? "hero" : ""} ${key ? "clickable" : ""}"${clickAttr}>
    <div class="t-label">${label}${key ? `<span class="csv-badge">CSV &#8681;</span>` : ""}</div>
    <div class="t-value">${value}</div>
    ${note ? `<div class="t-note ${cls}">${note}</div>` : ""}
  </div>`;
}

function slug(s) { return String(s).trim().replace(/[^\w]+/g, "_"); }

// generic CSV download helper
function downloadCsv(filename, headerLabels, rowsOfCells) {
  const head = headerLabels.join(",");
  const lines = rowsOfCells.map((cells) => cells.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","));
  // UTF-8 BOM so Excel reads non-ASCII chars (–, ₹) correctly instead of via the system codepage
  const blob = new Blob(["﻿" + head + "\n" + lines.join("\n")], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}
// enable Enter/Space activation + delegated click handling for a KPI row
function wireKpiRow(id, onKpiClick) {
  const el = $(id);
  el.addEventListener("click", (e) => {
    const t = e.target.closest("[data-kpi]");
    if (t) onKpiClick(t.dataset.kpi);
  });
  el.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" && e.key !== " ") return;
    const t = e.target.closest("[data-kpi]");
    if (t) { e.preventDefault(); onKpiClick(t.dataset.kpi); }
  });
}

// ---------- charts: shared helpers ----------
const tooltip = $("tooltip");
function showTip(html, x, y) {
  tooltip.innerHTML = html;
  tooltip.classList.remove("hidden");
  const r = tooltip.getBoundingClientRect();
  let left = x + 14, top = y + 14;
  if (left + r.width > window.innerWidth - 8) left = x - r.width - 14;
  if (top + r.height > window.innerHeight - 8) top = y - r.height - 14;
  tooltip.style.left = left + "px";
  tooltip.style.top = top + "px";
}
function hideTip() { tooltip.classList.add("hidden"); }

function niceTicks(max, count = 4) {
  if (max <= 0) return [0, 1];
  const step = Math.pow(10, Math.floor(Math.log10(max / count)));
  const mult = [1, 2, 5, 10].find((m) => (max / (step * m)) <= count) || 10;
  const s = step * mult, ticks = [0];
  // guarantee the final tick is >= max — otherwise the tallest bar's label clips off-chart
  while (ticks[ticks.length - 1] < max - 1e-9) ticks.push(ticks[ticks.length - 1] + s);
  return ticks;
}

// ---------- trend line chart ----------
document.querySelectorAll("[data-trend]").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll("[data-trend]").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    trendMode = btn.dataset.trend;
    renderTrend();
  });
});

function trendSeries() {
  const s = STATS;
  if (trendMode === "overall") {
    const vals = [];
    for (let k = 0; k < s.nDays; k++) {
      const denom = s.dayTrue[k] + s.dayFalse[k];
      vals.push(denom > 0 ? s.dayTrue[k] / denom : null);
    }
    return [{ name: "All meters", color: "var(--series-1)", vals }];
  }
  const out = [];
  const names = [...s.dayTspTrue.keys()].sort();
  // fixed entity->slot mapping: Airtel always slot 1, Vodafone always slot 2
  const slot = (n) => (/airtel/i.test(n) ? "var(--series-1)" : /voda/i.test(n) ? "var(--series-2)" : "var(--deemph)");
  for (const name of names) {
    const tt = s.dayTspTrue.get(name), td = s.dayTspData.get(name), vals = [];
    for (let k = 0; k < s.nDays; k++) vals.push(td[k] > 0 ? tt[k] / td[k] : null);
    out.push({ name, color: slot(name), vals });
  }
  return out;
}

function renderTrend() {
  if (tableModes["trend"]) return renderTrendTable();
  const series = trendSeries();
  const s = STATS, W = Math.max(680, Math.min(1300, $("chartTrend").clientWidth || 900));
  const H = 300, mL = 46, mR = 16, mT = 12, mB = 34;
  const iw = W - mL - mR, ih = H - mT - mB;
  const n = s.nDays;
  const x = (k) => mL + (n === 1 ? iw / 2 : (k / (n - 1)) * iw);
  const y = (v) => mT + (1 - v) * ih;

  let g = "";
  for (const t of [0, 0.25, 0.5, 0.75, 1]) {
    g += `<line x1="${mL}" y1="${y(t)}" x2="${W - mR}" y2="${y(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${y(t) + 4}" text-anchor="end">${t * 100}%</text>`;
  }
  // x labels ~6
  const stepX = Math.max(1, Math.round(n / 6));
  for (let k = 0; k < n; k += stepX) {
    g += `<text class="axis-text" x="${x(k)}" y="${H - 10}" text-anchor="middle">${dstrShort(DATA.dates[s.i0 + k])}</text>`;
  }
  g += `<line x1="${mL}" y1="${y(0)}" x2="${W - mR}" y2="${y(0)}" stroke="var(--baseline)" stroke-width="1"/>`;

  let paths = "";
  for (const ser of series) {
    let d = "", pen = false;
    for (let k = 0; k < n; k++) {
      const v = ser.vals[k];
      if (v === null) { pen = false; continue; }
      d += (pen ? "L" : "M") + x(k).toFixed(1) + " " + y(v).toFixed(1);
      pen = true;
    }
    if (series.length === 1) {
      // single series: soft area wash under the line
      let a = "", first = null, last = null;
      for (let k = 0; k < n; k++) if (ser.vals[k] !== null) { if (first === null) first = k; last = k; }
      if (first !== null) {
        a = "M" + x(first).toFixed(1) + " " + y(0);
        for (let k = first; k <= last; k++) if (ser.vals[k] !== null) a += "L" + x(k).toFixed(1) + " " + y(ser.vals[k]).toFixed(1);
        a += "L" + x(last).toFixed(1) + " " + y(0) + "Z";
        paths += `<path d="${a}" fill="var(--series-area)"/>`;
      }
    }
    paths += `<path d="${d}" fill="none" stroke="${ser.color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
  }

  const legend = series.length > 1
    ? `<div class="legend">` + series.map((s2) => `<span class="key"><span class="swatch" style="background:${s2.color}"></span>${esc(s2.name)}</span>`).join("") + `</div>`
    : "";

  $("chartTrend").innerHTML = legend +
    `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">
      ${g}${paths}
      <line id="crosshair" x1="0" y1="${mT}" x2="0" y2="${mT + ih}" stroke="var(--baseline)" stroke-width="1" visibility="hidden"/>
      <g id="hoverDots"></g>
      <rect x="${mL}" y="${mT}" width="${iw}" height="${ih}" fill="transparent" id="trendHover"/>
    </svg>`;

  const svg = $("chartTrend").querySelector("svg");
  const hover = svg.querySelector("#trendHover");
  const cross = svg.querySelector("#crosshair");
  const dots = svg.querySelector("#hoverDots");
  hover.addEventListener("mousemove", (e) => {
    const rect = svg.getBoundingClientRect();
    const scale = W / rect.width;
    const px = (e.clientX - rect.left) * scale;
    const k = Math.max(0, Math.min(n - 1, Math.round(((px - mL) / iw) * (n - 1))));
    cross.setAttribute("x1", x(k)); cross.setAttribute("x2", x(k));
    cross.setAttribute("visibility", "visible");
    let dotHtml = "", rowsHtml = "";
    for (const ser of series) {
      const v = ser.vals[k];
      if (v === null) continue;
      dotHtml += `<circle cx="${x(k)}" cy="${y(v)}" r="4.5" fill="${ser.color}" stroke="var(--surface-1)" stroke-width="2"/>`;
      rowsHtml += `<div class="tt-row"><span>${esc(ser.name)}</span><b>${pct(v)}</b></div>`;
    }
    const denom = s.dayTrue[k] + s.dayFalse[k];
    rowsHtml += `<div class="tt-row"><span>Meters with data</span><b>${fmt(denom)}</b></div>`;
    rowsHtml += `<div class="tt-row"><span>Communicated</span><b>${fmt(s.dayTrue[k])}</b></div>`;
    dots.innerHTML = dotHtml;
    showTip(`<div class="tt-title">${dstr(DATA.dates[s.i0 + k])}</div>${rowsHtml}`, e.clientX, e.clientY);
  });
  hover.addEventListener("mouseleave", () => { cross.setAttribute("visibility", "hidden"); dots.innerHTML = ""; hideTip(); });
}

function renderTrendTable() {
  const series = trendSeries();
  const s = STATS;
  let rows = "";
  for (let k = 0; k < s.nDays; k++) {
    rows += `<tr><td>${dstr(DATA.dates[s.i0 + k])}</td>` +
      series.map((ser) => `<td class="num">${ser.vals[k] === null ? "—" : pct(ser.vals[k])}</td>`).join("") +
      `<td class="num">${fmt(s.dayTrue[k] + s.dayFalse[k])}</td></tr>`;
  }
  $("chartTrend").innerHTML = `<div class="table-scroll mini-table" style="max-height:320px;overflow-y:auto"><table>
    <thead><tr><th>Date</th>${series.map((x2) => `<th class="num">${esc(x2.name)}</th>`).join("")}<th class="num">Meters with data</th></tr></thead>
    <tbody>${rows}</tbody></table></div>`;
}

// ---------- horizontal bar chart (single hue) ----------
function hBars(container, items, opts = {}) {
  // items: {label, value (0..1), sub, tip}
  const W = Math.max(420, Math.min(660, $(container).clientWidth || 560));
  const barH = 18, gap = 14, labelW = opts.labelW || 130;
  const mT = 6, mB = 24, mR = 60;
  const H = mT + items.length * (barH + gap) + mB;
  const iw = W - labelW - mR;
  const xv = (v) => labelW + v * iw;

  let g = "";
  for (const t of [0, 0.25, 0.5, 0.75, 1]) {
    g += `<line x1="${xv(t)}" y1="${mT}" x2="${xv(t)}" y2="${H - mB}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${xv(t)}" y="${H - 8}" text-anchor="middle">${t * 100}%</text>`;
  }
  let bars = "";
  items.forEach((it, i) => {
    const yTop = mT + i * (barH + gap);
    const w = Math.max(0, it.value * iw);
    // rounded data-end, square baseline end
    const r = Math.min(4, w);
    bars += `<path d="M${labelW} ${yTop} h${(w - r).toFixed(1)} a${r} ${r} 0 0 1 ${r} ${r} v${barH - 2 * r} a${r} ${r} 0 0 1 -${r} ${r} h-${(w - r).toFixed(1)} Z"
      fill="var(--series-1)" data-i="${i}" class="hbar"/>`;
    if (it.sub) {
      bars += `<text class="bar-label" x="${labelW - 8}" y="${yTop + barH / 2}" text-anchor="end">${esc(trunc(it.label, 18))}</text>`;
      bars += `<text class="bar-sub" x="${labelW - 8}" y="${yTop + barH / 2 + 11}" text-anchor="end">${esc(it.sub)}</text>`;
    } else {
      bars += `<text class="bar-label" x="${labelW - 8}" y="${yTop + barH / 2 + 4}" text-anchor="end">${esc(trunc(it.label, 18))}</text>`;
    }
    bars += `<text class="bar-label" x="${(labelW + w + 6).toFixed(1)}" y="${yTop + barH / 2 + 4}">${pct(it.value)}</text>`;
    // full-row invisible hover target (≥24px tall)
    bars += `<rect x="0" y="${yTop - gap / 2}" width="${W}" height="${barH + gap}" fill="transparent" data-i="${i}" class="hrow"/>`;
  });
  $(container).innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${g}${bars}</svg>`;

  $(container).querySelectorAll(".hrow").forEach((el) => {
    const it = items[+el.dataset.i];
    el.addEventListener("mousemove", (e) => showTip(it.tip + (opts.onClick ? `<div class="tt-row"><i>Click bar to export CSV</i></div>` : ""), e.clientX, e.clientY));
    el.addEventListener("mouseleave", hideTip);
    if (opts.onClick) { el.style.cursor = "pointer"; el.addEventListener("click", () => opts.onClick(it)); }
  });
}
function trunc(s, n) { return s.length > n ? s.slice(0, n - 1) + "…" : s; }

function groupItems(g, limit, statsSrc = STATS) {
  const map = statsSrc.byGroup[g];
  const items = [];
  map.forEach((e, key) => {
    const denom = e.t + e.f;
    if (denom === 0 && e.meters === 0) return;
    items.push({
      label: key,
      value: denom > 0 ? e.t / denom : 0,
      meters: e.meters, t: e.t, f: e.f, n: e.n,
    });
  });
  items.sort((a, b) => a.value - b.value); // worst first
  return items.slice(0, limit);
}

function renderGroupBar(g, container, limit, statsSrc = STATS, cardKey = g) {
  const items = groupItems(g, limit, statsSrc).map((it) => ({
    ...it,
    sub: `${fmt(it.meters)} meters`,
    tip: `<div class="tt-title">${esc(it.label)}</div>
      <div class="tt-row"><span>Comm rate</span><b>${pct(it.value)}</b></div>
      <div class="tt-row"><span>Meters</span><b>${fmt(it.meters)}</b></div>
      <div class="tt-row"><span>TRUE days</span><b>${fmt(it.t)}</b></div>
      <div class="tt-row"><span>FALSE days</span><b>${fmt(it.f)}</b></div>
      <div class="tt-row"><span>No-data days</span><b>${fmt(it.n)}</b></div>`,
  }));
  if (tableModes[cardKey]) {
    $(container).innerHTML = groupTable(items);
  } else if (items.length === 0) {
    $(container).innerHTML = `<p class="card-sub">No data for this slice.</p>`;
  } else {
    hBars(container, items, {
      onClick: (it) => downloadMeterCsv(`${g}_${slug(it.label)}.csv`, statsSrc.perMeter.filter((r) => r.m[g] === it.label)),
    });
  }
}
function groupTable(items) {
  return `<div class="table-scroll mini-table"><table>
    <thead><tr><th>Group</th><th class="num">Comm rate</th><th class="num">Meters</th><th class="num">TRUE</th><th class="num">FALSE</th><th class="num">No data</th></tr></thead>
    <tbody>${items.map((it) => `<tr><td>${esc(it.label)}</td><td class="num">${pct(it.value)}</td><td class="num">${fmt(it.meters)}</td><td class="num">${fmt(it.t)}</td><td class="num">${fmt(it.f)}</td><td class="num">${fmt(it.n)}</td></tr>`).join("")}</tbody>
  </table></div>`;
}

// ---------- TSP card ----------
function renderTspChart() { renderGroupBar("tsp", "chartTsp", 99); }

// ---------- health histogram ----------
const BUCKET_LABELS = ["0%", "0–25%", "25–50%", "50–75%", "75–85%", "85–95%", "95–99%", "100%"];
function renderHealth(statsSrc = STATS, container = "chartHealth", cardKey = "health") {
  const b = statsSrc.buckets;
  if (tableModes[cardKey]) {
    $(container).innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Uptime bucket</th><th class="num">Meters</th></tr></thead>
      <tbody>${BUCKET_LABELS.map((l, i) => `<tr><td>${l}</td><td class="num">${fmt(b[i])}</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  const W = Math.max(420, Math.min(660, $(container).clientWidth || 560));
  const H = 240, mL = 52, mR = 10, mT = 12, mB = 30;
  const iw = W - mL - mR, ih = H - mT - mB;
  const max = Math.max(...b, 1);
  const ticks = niceTicks(max);
  const yv = (v) => mT + ih - (v / ticks[ticks.length - 1]) * ih;
  const slotW = iw / b.length;
  const barW = Math.min(24, slotW * 0.55);

  let g = "";
  for (const t of ticks) {
    g += `<line x1="${mL}" y1="${yv(t)}" x2="${W - mR}" y2="${yv(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${yv(t) + 4}" text-anchor="end">${fmt(t)}</text>`;
  }
  let bars = "";
  b.forEach((v, i) => {
    const cx = mL + slotW * i + slotW / 2;
    const yTop = yv(v), h = mT + ih - yTop;
    const r = Math.min(4, h);
    bars += `<path d="M${(cx - barW / 2).toFixed(1)} ${(mT + ih).toFixed(1)} v-${(h - r).toFixed(1)} a${r} ${r} 0 0 1 ${r} -${r} h${(barW - 2 * r).toFixed(1)} a${r} ${r} 0 0 1 ${r} ${r} v${(h - r).toFixed(1)} Z" fill="var(--series-1)"/>`;
    bars += `<text class="axis-text" x="${cx}" y="${H - 8}" text-anchor="middle">${BUCKET_LABELS[i]}</text>`;
    bars += `<rect x="${(cx - slotW / 2).toFixed(1)}" y="${mT}" width="${slotW.toFixed(1)}" height="${ih}" fill="transparent" class="hcol" data-i="${i}"/>`;
  });
  bars += `<line x1="${mL}" y1="${mT + ih}" x2="${W - mR}" y2="${mT + ih}" stroke="var(--baseline)" stroke-width="1"/>`;

  $(container).innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${g}${bars}</svg>`;
  $(container).querySelectorAll(".hcol").forEach((el) => {
    el.style.cursor = "pointer";
    const i = +el.dataset.i;
    el.addEventListener("mousemove", (e) => {
      showTip(`<div class="tt-title">Uptime ${BUCKET_LABELS[i]}</div><div class="tt-row"><span>Meters</span><b>${fmt(b[i])}</b></div><div class="tt-row"><i>Click bar to export CSV</i></div>`, e.clientX, e.clientY);
    });
    el.addEventListener("mouseleave", hideTip);
    el.addEventListener("click", () => downloadMeterCsv(`uptime_bucket_${slug(BUCKET_LABELS[i])}.csv`, statsSrc.perMeter.filter((r) => r.bucket === i)));
  });
}

// same layout as renderHealth(), but each bar is that bucket's LOSS AMOUNT IN ₹ —
// i.e. which uptime band is actually driving the revenue loss in rupees. Includes a "No data"
// bar for meters with zero data days.
const LOSS_BUCKET_LABELS = [...BUCKET_LABELS, "No data"];
function renderLossByBucket(statsSrc = STATS, container = "chartLossBucket", cardKey = "lossbucket") {
  const { byBucket, noData } = computeLossByBucket(statsSrc);
  const losses = [...byBucket.map((b) => b.loss), noData.loss];
  const meterCounts = [...byBucket.map((b) => b.meters), noData.meters];
  const totalLoss = losses.reduce((a, v) => a + v, 0) || 1;
  const pctOfLoss = losses.map((v) => (v / totalLoss) * 100);

  if (tableModes[cardKey]) {
    $(container).innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Uptime bucket</th><th class="num">Meters</th><th class="num">Loss (₹)</th><th class="num">% of total</th></tr></thead>
      <tbody>${LOSS_BUCKET_LABELS.map((l, i) => `<tr><td>${l}</td><td class="num">${fmt(meterCounts[i])}</td><td class="num">${inr(losses[i])}</td><td class="num">${pctOfLoss[i].toFixed(1)}%</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  const W = Math.max(420, Math.min(660, $(container).clientWidth || 560));
  const H = 240, mL = 52, mR = 10, mT = 12, mB = 30;
  const iw = W - mL - mR, ih = H - mT - mB;
  const max = Math.max(...losses, 1);
  const ticks = niceTicks(max);
  const yv = (v) => mT + ih - (v / ticks[ticks.length - 1]) * ih;
  const slotW = iw / losses.length;
  const barW = Math.min(24, slotW * 0.55);

  let g = "";
  for (const t of ticks) {
    g += `<line x1="${mL}" y1="${yv(t)}" x2="${W - mR}" y2="${yv(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${yv(t) + 4}" text-anchor="end">${inr(t)}</text>`;
  }
  let bars = "";
  losses.forEach((v, i) => {
    const cx = mL + slotW * i + slotW / 2;
    const yTop = yv(losses[i]), h = mT + ih - yTop;
    const r = Math.min(4, h);
    bars += `<path d="M${(cx - barW / 2).toFixed(1)} ${(mT + ih).toFixed(1)} v-${(h - r).toFixed(1)} a${r} ${r} 0 0 1 ${r} -${r} h${(barW - 2 * r).toFixed(1)} a${r} ${r} 0 0 1 ${r} ${r} v${(h - r).toFixed(1)} Z" fill="var(--critical)"/>`;
    bars += `<text class="axis-text" x="${cx}" y="${H - 8}" text-anchor="middle">${LOSS_BUCKET_LABELS[i]}</text>`;
    bars += `<rect x="${(cx - slotW / 2).toFixed(1)}" y="${mT}" width="${slotW.toFixed(1)}" height="${ih}" fill="transparent" class="lbcol" data-i="${i}"/>`;
  });
  bars += `<line x1="${mL}" y1="${mT + ih}" x2="${W - mR}" y2="${mT + ih}" stroke="var(--baseline)" stroke-width="1"/>`;

  $(container).innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${g}${bars}</svg>`;
  $(container).querySelectorAll(".lbcol").forEach((el) => {
    el.style.cursor = "pointer";
    const i = +el.dataset.i;
    el.addEventListener("mousemove", (e) => {
      showTip(`<div class="tt-title">Uptime ${LOSS_BUCKET_LABELS[i]}</div>
        <div class="tt-row"><span>Share of total loss</span><b>${pctOfLoss[i].toFixed(1)}%</b></div>
        <div class="tt-row"><span>Loss amount</span><b>${inr(losses[i])}</b></div>
        <div class="tt-row"><span>Meters</span><b>${fmt(meterCounts[i])}</b></div>
        <div class="tt-row"><i>Click bar to export CSV</i></div>`, e.clientX, e.clientY);
    });
    el.addEventListener("mouseleave", hideTip);
    el.addEventListener("click", () => downloadMeterCsv(`loss_bucket_${slug(LOSS_BUCKET_LABELS[i])}.csv`,
      i === LOSS_BUCKET_LABELS.length - 1 ? statsSrc.perMeter.filter((r) => r.bucket === -1) : statsSrc.perMeter.filter((r) => r.bucket === i)));
  });
}

// ---------- chart/table toggles ----------
document.querySelectorAll(".tbl-toggle").forEach((btn) => {
  btn.addEventListener("click", () => {
    const key = btn.dataset.card;
    tableModes[key] = !tableModes[key];
    btn.classList.toggle("active", !!tableModes[key]);
    btn.textContent = tableModes[key] ? "Chart" : "Table";
    if (key === "trend") renderTrend();
    else if (key === "tsp") renderTspChart();
    else if (key === "health") renderHealth();
    else if (key === "region") renderGroupBar("region", "chartRegion", 99);
    else if (key === "circle") renderGroupBar("circle", "chartCircle", 15);
    else if (key === "po") renderGroupBar("po", "chartPo", 99);
    else if (key === "dow") renderDowChart();
    else if (key === "wetsp") renderWeekendPaired("tsp", "chartWeTsp", 99);
    else if (key === "weregion") renderWeekendPaired("region", "chartWeRegion", 99);
    else if (key === "wecircle") renderWeekendPaired("circle", "chartWeCircle", 10);
    else if (key === "wepo") renderWeekendPaired("po", "chartWePo", 99);
    else if (key === "patdist") renderPatternDist();
    else if (key === "patregion") renderPatternRegion();
    else if (key === "rhdist") renderRhythmDist();
    else if (key === "rhregion") renderRhythmRegion();
    else if (key === "drtrend") renderTrendFor(DRSTATS, "chartDrTrend", "drtrend");
    else if (key === "drtsp") renderGroupBar("tsp", "chartDrTsp", 99, DRSTATS, "drtsp");
    else if (key === "drhealth") renderHealth(DRSTATS, "chartDrHealth", "drhealth");
    else if (key === "drregion") renderGroupBar("region", "chartDrRegion", 99, DRSTATS, "drregion");
    else if (key === "drcircle") renderGroupBar("circle", "chartDrCircle", 15, DRSTATS, "drcircle");
    else if (key === "drpo") renderGroupBar("po", "chartDrPo", 99, DRSTATS, "drpo");
    else if (key === "motrend") renderTrendFor(MOSTATS, "chartMoTrend", "motrend", (k) => DATA.dates[MOSTATS.indices[k]]);
    else if (key === "motsp") renderGroupBar("tsp", "chartMoTsp", 99, MOSTATS, "motsp");
    else if (key === "mohealth") renderHealth(MOSTATS, "chartMoHealth", "mohealth");
    else if (key === "moregion") renderGroupBar("region", "chartMoRegion", 99, MOSTATS, "moregion");
    else if (key === "mocircle") renderGroupBar("circle", "chartMoCircle", 15, MOSTATS, "mocircle");
    else if (key === "mopo") renderGroupBar("po", "chartMoPo", 99, MOSTATS, "mopo");
    else if (key === "mopatdist") renderMoPatternDist();
    else if (key === "mopatregion") renderMoPatternRegion();
    else if (key === "morhdist") renderMoRhythmDist();
    else if (key === "morhregion") renderMoRhythmRegion();
    else if (key === "losstrend") renderLossTrend();
    else if (key === "lossbucket") renderLossByBucket(STATS, "chartLossBucket", "lossbucket");
    else if (key === "smregion") renderSmRegionChart();
    else if (key === "smcircle") renderSmCircleChart();
    else if (key === "smstatus") renderSmStatusChart();
    else if (key === "smmake") renderSmMakeChart();
    else if (key === "sminsightregion") renderSmInsightRegion();
    else if (key === "sminsightcircle") renderSmInsightCircle();
  });
});

// ---------- problem meters table ----------
const TABLE_COLS = [
  { key: "ip", label: "IP", num: false },
  { key: "tsp", label: "TSP", num: false },
  { key: "region", label: "Region", num: false },
  { key: "circle", label: "Circle", num: false },
  { key: "po", label: "PO", num: false },
  { key: "uptime", label: "Uptime %", num: true },
  { key: "t", label: "TRUE", num: true },
  { key: "f", label: "FALSE", num: true },
  { key: "nl", label: "No data", num: true },
  { key: "silentDays", label: "Silent days", num: true },
  { key: "lastSeen", label: "Last communicated", num: false },
];

$("meterSearch").addEventListener("input", (e) => { tableSearch = e.target.value.trim().toLowerCase(); tableLimit = 25; renderMeterTable(); });
$("loadMore").addEventListener("click", () => { tableLimit += 50; renderMeterTable(); });
$("exportCsv").addEventListener("click", exportCsv);

function tableRows() {
  let rows = STATS.perMeter.filter((r) => r.uptime >= 0); // meters with at least one data day
  if (tableSearch) rows = rows.filter((r) => r.m.ip.toLowerCase().includes(tableSearch));
  const k = tableSort.key, dir = tableSort.dir;
  const val = (r) => {
    if (k === "ip" || k === "tsp" || k === "region" || k === "circle" || k === "po") return r.m[k];
    if (k === "lastSeen") return r.lastSeen;
    return r[k];
  };
  rows.sort((a, b) => {
    const va = val(a), vb = val(b);
    if (typeof va === "string") return dir * va.localeCompare(vb);
    return dir * (va - vb);
  });
  return rows;
}

function renderMeterTable() {
  $("meterTableHead").innerHTML = TABLE_COLS.map((c) =>
    `<th class="${c.num ? "num" : ""} ${tableSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${tableSort.key === c.key ? (tableSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  $("meterTableHead").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (tableSort.key === key) tableSort.dir *= -1;
      else tableSort = { key, dir: 1 };
      renderMeterTable();
    });
  });
  const rows = tableRows();
  const shown = rows.slice(0, tableLimit);
  $("meterTableBody").innerHTML = shown.map((r) => `<tr>
    <td>${esc(r.m.ip)}</td><td>${esc(r.m.tsp)}</td><td>${esc(r.m.region)}</td>
    <td>${esc(r.m.circle)}</td><td>${esc(r.m.po)}</td>
    <td class="num ${r.uptime < 0.5 ? "pct-bad" : ""}">${pct(r.uptime)}</td>
    <td class="num">${fmt(r.t)}</td><td class="num">${fmt(r.f)}</td><td class="num">${fmt(r.nl)}</td>
    <td class="num">${r.silentDays < 0 ? "—" : fmt(r.silentDays)}</td>
    <td>${r.lastSeen >= 0 ? dstr(DATA.dates[r.lastSeen]) : "never"}</td>
  </tr>`).join("");
  $("tableCount").textContent = `Showing ${fmt(shown.length)} of ${fmt(rows.length)} meters`;
  $("loadMore").style.visibility = shown.length < rows.length ? "visible" : "hidden";
}

function downloadMeterCsv(filename, rows) {
  const cells = rows.map((r) => [
    r.m.ip, r.m.tsp, r.m.region, r.m.circle, r.m.po,
    (r.uptime * 100).toFixed(2), r.t, r.f, r.nl,
    r.silentDays < 0 ? "" : r.silentDays,
    r.lastSeen >= 0 ? DATA.dates[r.lastSeen].toISOString().slice(0, 10) : "never",
  ]);
  downloadCsv(filename, TABLE_COLS.map((c) => c.label), cells);
}
function exportCsv() {
  downloadMeterCsv("problem_meters.csv", tableRows());
}

// ==========================================================
// Weekend vs Weekday tab
// ==========================================================
function renderWeekend() {
  renderWeekendKpis();
  renderDowChart();
  renderWeekendPaired("tsp", "chartWeTsp", 99);
  renderWeekendPaired("region", "chartWeRegion", 99);
  renderWeekendPaired("circle", "chartWeCircle", 10);
  renderWeekendPaired("po", "chartWePo", 99);
  renderWeMeterTable();
}

function renderWeekendKpis() {
  const w = WSTATS;
  const wdRate = w.weekdayT + w.weekdayF > 0 ? w.weekdayT / (w.weekdayT + w.weekdayF) : 0;
  const weRate = w.weekendT + w.weekendF > 0 ? w.weekendT / (w.weekendT + w.weekendF) : 0;
  const gapPp = (wdRate - weRate) * 100;
  let bestI = 0, worstI = 0;
  w.dowRate.forEach((r, i) => {
    if (r === null) return;
    if (w.dowRate[bestI] === null || r > w.dowRate[bestI]) bestI = i;
    if (w.dowRate[worstI] === null || r < w.dowRate[worstI]) worstI = i;
  });
  $("weKpiRow").innerHTML = [
    tile("Weekday communication rate", pct(wdRate), `${satIsWeekend ? "Mon–Fri" : "Mon–Sat"}, TRUE ÷ (TRUE+FALSE)`, true, "all"),
    tile("Weekend communication rate", pct(weRate), `${satIsWeekend ? "Sat–Sun" : "Sun only"}, TRUE ÷ (TRUE+FALSE)`, true, "all"),
    tile("Weekday − weekend gap", (gapPp >= 0 ? "+" : "") + gapPp.toFixed(1) + " pp",
      gapPp > 0.5 ? "meters communicate less on weekends" : gapPp < -0.5 ? "meters communicate more on weekends" : "no meaningful gap",
      Math.abs(gapPp) > 0.5 ? "bad" : "good", "all"),
    tile("Best day", w.dowRate[bestI] === null ? "—" : DOW_LABELS[bestI], w.dowRate[bestI] === null ? "" : pct(w.dowRate[bestI]) + " comm rate"),
    tile("Worst day", w.dowRate[worstI] === null ? "—" : DOW_LABELS[worstI], w.dowRate[worstI] === null ? "" : pct(w.dowRate[worstI]) + " comm rate"),
  ].join("");
}
wireKpiRow("weKpiRow", () => {
  if (!WSTATS) return;
  const rows = WSTATS.perMeter.filter((r) => r.gap !== null).sort((a, b) => b.gap - a.gap);
  downloadWeCsv("all_meters_weekday_weekend.csv", rows);
});

function renderDowChart() {
  const w = WSTATS;
  if (tableModes["dow"]) {
    $("chartDow").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Day</th><th class="num">Comm rate</th><th class="num">TRUE</th><th class="num">FALSE</th></tr></thead>
      <tbody>${DOW_LABELS.map((l, i) => `<tr><td>${l}${isWeekendDow(i) ? " (weekend)" : ""}</td><td class="num">${w.dowRate[i] === null ? "—" : pct(w.dowRate[i])}</td><td class="num">${fmt(w.dowT[i])}</td><td class="num">${fmt(w.dowF[i])}</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  const W = Math.max(500, Math.min(900, $("chartDow").clientWidth || 700));
  const H = 260, mL = 46, mR = 14, mT = 12, mB = 30;
  const iw = W - mL - mR, ih = H - mT - mB;
  const max = 1;
  const slotW = iw / 7, barW = Math.min(24, slotW * 0.55);
  const yv = (v) => mT + ih - v * ih;

  let g = "";
  for (const t of [0, 0.25, 0.5, 0.75, 1]) {
    g += `<line x1="${mL}" y1="${yv(t)}" x2="${W - mR}" y2="${yv(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${yv(t) + 4}" text-anchor="end">${t * 100}%</text>`;
  }
  let bars = "";
  for (let i = 0; i < 7; i++) {
    const v = w.dowRate[i] || 0;
    const cx = mL + slotW * i + slotW / 2;
    const yTop = yv(v), h = mT + ih - yTop;
    const r = Math.min(4, h);
    const color = isWeekendDow(i) ? "var(--series-2)" : "var(--series-1)";
    bars += `<path d="M${(cx - barW / 2).toFixed(1)} ${(mT + ih).toFixed(1)} v-${(h - r).toFixed(1)} a${r} ${r} 0 0 1 ${r} -${r} h${(barW - 2 * r).toFixed(1)} a${r} ${r} 0 0 1 ${r} ${r} v${(h - r).toFixed(1)} Z" fill="${color}"/>`;
    bars += `<text class="axis-text" x="${cx}" y="${H - 8}" text-anchor="middle">${DOW_LABELS[i]}</text>`;
    bars += `<rect x="${(cx - slotW / 2).toFixed(1)}" y="${mT}" width="${slotW.toFixed(1)}" height="${ih}" fill="transparent" class="dowcol" data-i="${i}"/>`;
  }
  bars += `<line x1="${mL}" y1="${mT + ih}" x2="${W - mR}" y2="${mT + ih}" stroke="var(--baseline)" stroke-width="1"/>`;

  const legend = `<div class="legend">
    <span class="key"><span class="swatch" style="background:var(--series-1)"></span>Weekday</span>
    <span class="key"><span class="swatch" style="background:var(--series-2)"></span>Weekend</span></div>`;

  $("chartDow").innerHTML = legend + `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${g}${bars}</svg>`;
  $("chartDow").querySelectorAll(".dowcol").forEach((el) => {
    el.addEventListener("mousemove", (e) => {
      const i = +el.dataset.i;
      showTip(`<div class="tt-title">${DOW_LABELS[i]}${isWeekendDow(i) ? " (weekend)" : ""}</div>
        <div class="tt-row"><span>Comm rate</span><b>${w.dowRate[i] === null ? "—" : pct(w.dowRate[i])}</b></div>
        <div class="tt-row"><span>TRUE</span><b>${fmt(w.dowT[i])}</b></div>
        <div class="tt-row"><span>FALSE</span><b>${fmt(w.dowF[i])}</b></div>`, e.clientX, e.clientY);
    });
    el.addEventListener("mouseleave", hideTip);
  });
}

// paired horizontal bars: weekday (series-1) vs weekend (series-2) per group, sorted by |gap| desc
function weekendGroupItems(g, limit) {
  const src = g === "tsp" ? WSTATS.byTsp : g === "region" ? WSTATS.byRegion : g === "circle" ? WSTATS.byCircle : WSTATS.byPo;
  const items = [];
  src.forEach((e, key) => {
    const wd = e.wdT + e.wdF > 0 ? e.wdT / (e.wdT + e.wdF) : null;
    const we = e.weT + e.weF > 0 ? e.weT / (e.weT + e.weF) : null;
    if (wd === null && we === null) return;
    items.push({ label: key, wd, we, gap: wd !== null && we !== null ? wd - we : 0, meters: e.meters, wdT: e.wdT, wdF: e.wdF, weT: e.weT, weF: e.weF });
  });
  items.sort((a, b) => Math.abs(b.gap) - Math.abs(a.gap));
  return items.slice(0, limit);
}

function pairedBars(container, items) {
  const W = Math.max(420, Math.min(660, $(container).clientWidth || 560));
  const rowH = 34, gap = 10, labelW = 130, mT = 6, mB = 24, mR = 60;
  const H = mT + items.length * (rowH + gap) + mB;
  const iw = W - labelW - mR;
  const xv = (v) => labelW + Math.max(0, v) * iw;
  const barH = 12;

  let g2 = "";
  for (const t of [0, 0.25, 0.5, 0.75, 1]) {
    g2 += `<line x1="${xv(t)}" y1="${mT}" x2="${xv(t)}" y2="${H - mB}" stroke="var(--grid)" stroke-width="1"/>`;
    g2 += `<text class="axis-text" x="${xv(t)}" y="${H - 8}" text-anchor="middle">${t * 100}%</text>`;
  }
  let bars = "";
  items.forEach((it, i) => {
    const yTop = mT + i * (rowH + gap);
    bars += `<text class="bar-label" x="${labelW - 8}" y="${yTop + rowH / 2 + 4}" text-anchor="end">${esc(trunc(it.label, 16))}</text>`;
    if (it.wd !== null) {
      const w1 = it.wd * iw;
      bars += `<rect x="${labelW}" y="${yTop}" width="${w1.toFixed(1)}" height="${barH}" rx="3" fill="var(--series-1)"/>`;
      bars += `<text class="bar-label" x="${(labelW + w1 + 6).toFixed(1)}" y="${yTop + barH - 1}">${pct(it.wd)}</text>`;
    }
    if (it.we !== null) {
      const w2 = it.we * iw;
      bars += `<rect x="${labelW}" y="${yTop + barH + 3}" width="${w2.toFixed(1)}" height="${barH}" rx="3" fill="var(--series-2)"/>`;
      bars += `<text class="bar-label" x="${(labelW + w2 + 6).toFixed(1)}" y="${yTop + 2 * barH + 2}">${pct(it.we)}</text>`;
    }
    bars += `<rect x="0" y="${yTop - gap / 2}" width="${W}" height="${rowH + gap}" fill="transparent" data-i="${i}" class="hrow"/>`;
  });

  const legend = `<div class="legend">
    <span class="key"><span class="swatch" style="background:var(--series-1)"></span>Weekday</span>
    <span class="key"><span class="swatch" style="background:var(--series-2)"></span>Weekend</span></div>`;

  $(container).innerHTML = legend + `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${g2}${bars}</svg>`;
  $(container).querySelectorAll(".hrow").forEach((el) => {
    el.addEventListener("mousemove", (e) => {
      const it = items[+el.dataset.i];
      showTip(`<div class="tt-title">${esc(it.label)}</div>
        <div class="tt-row"><span>Weekday</span><b>${it.wd === null ? "—" : pct(it.wd)}</b></div>
        <div class="tt-row"><span>Weekend</span><b>${it.we === null ? "—" : pct(it.we)}</b></div>
        <div class="tt-row"><span>Gap</span><b>${it.wd !== null && it.we !== null ? ((it.wd - it.we) * 100).toFixed(1) + " pp" : "—"}</b></div>
        <div class="tt-row"><span>Meters</span><b>${fmt(it.meters)}</b></div>`, e.clientX, e.clientY);
    });
    el.addEventListener("mouseleave", hideTip);
  });
}

function renderWeekendPaired(g, container, limit) {
  const cardKey = g === "tsp" ? "wetsp" : g === "region" ? "weregion" : g === "circle" ? "wecircle" : "wepo";
  const items = weekendGroupItems(g, limit);
  if (tableModes[cardKey]) {
    $(container).innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Group</th><th class="num">Weekday</th><th class="num">Weekend</th><th class="num">Gap (pp)</th><th class="num">Meters</th></tr></thead>
      <tbody>${items.map((it) => `<tr><td>${esc(it.label)}</td><td class="num">${it.wd === null ? "—" : pct(it.wd)}</td><td class="num">${it.we === null ? "—" : pct(it.we)}</td><td class="num">${it.wd !== null && it.we !== null ? ((it.wd - it.we) * 100).toFixed(1) : "—"}</td><td class="num">${fmt(it.meters)}</td></tr>`).join("")}</tbody></table></div>`;
  } else if (items.length === 0) {
    $(container).innerHTML = `<p class="card-sub">No data for this slice.</p>`;
  } else {
    pairedBars(container, items);
  }
}

// ---------- weekend gap meter table ----------
const WE_TABLE_COLS = [
  { key: "ip", label: "IP", num: false },
  { key: "tsp", label: "TSP", num: false },
  { key: "region", label: "Region", num: false },
  { key: "circle", label: "Circle", num: false },
  { key: "po", label: "PO", num: false },
  { key: "wdRate", label: "Weekday %", num: true },
  { key: "weRate", label: "Weekend %", num: true },
  { key: "gap", label: "Gap (pp)", num: true },
];

$("weMeterSearch").addEventListener("input", (e) => { weTableSearch = e.target.value.trim().toLowerCase(); weTableLimit = 25; renderWeMeterTable(); });
$("weLoadMore").addEventListener("click", () => { weTableLimit += 50; renderWeMeterTable(); });
$("weExportCsv").addEventListener("click", weExportCsv);

function weTableRows() {
  let rows = WSTATS.perMeter.filter((r) => r.gap !== null);
  if (weTableSearch) rows = rows.filter((r) => r.m.ip.toLowerCase().includes(weTableSearch));
  const k = weTableSort.key, dir = weTableSort.dir;
  const val = (r) => {
    if (k === "ip" || k === "tsp" || k === "region" || k === "circle" || k === "po") return r.m[k];
    return r[k];
  };
  rows.sort((a, b) => {
    const va = val(a), vb = val(b);
    if (typeof va === "string") return dir * va.localeCompare(vb);
    return dir * (va - vb);
  });
  return rows;
}

function renderWeMeterTable() {
  $("weMeterTableHead").innerHTML = WE_TABLE_COLS.map((c) =>
    `<th class="${c.num ? "num" : ""} ${weTableSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${weTableSort.key === c.key ? (weTableSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  $("weMeterTableHead").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (weTableSort.key === key) weTableSort.dir *= -1;
      else weTableSort = { key, dir: -1 };
      renderWeMeterTable();
    });
  });
  const rows = weTableRows();
  const shown = rows.slice(0, weTableLimit);
  $("weMeterTableBody").innerHTML = shown.map((r) => `<tr>
    <td>${esc(r.m.ip)}</td><td>${esc(r.m.tsp)}</td><td>${esc(r.m.region)}</td>
    <td>${esc(r.m.circle)}</td><td>${esc(r.m.po)}</td>
    <td class="num">${r.wdRate === null ? "—" : pct(r.wdRate)}</td>
    <td class="num">${r.weRate === null ? "—" : pct(r.weRate)}</td>
    <td class="num ${Math.abs(r.gap) > 0.3 ? "pct-bad" : ""}">${(r.gap * 100).toFixed(1)}</td>
  </tr>`).join("");
  $("weTableCount").textContent = `Showing ${fmt(shown.length)} of ${fmt(rows.length)} meters`;
  $("weLoadMore").style.visibility = shown.length < rows.length ? "visible" : "hidden";
}

function downloadWeCsv(filename, rows) {
  const cells = rows.map((r) => [
    r.m.ip, r.m.tsp, r.m.region, r.m.circle, r.m.po,
    r.wdRate === null ? "" : (r.wdRate * 100).toFixed(2),
    r.weRate === null ? "" : (r.weRate * 100).toFixed(2),
    (r.gap * 100).toFixed(2),
  ]);
  downloadCsv(filename, WE_TABLE_COLS.map((c) => c.label), cells);
}
function weExportCsv() {
  downloadWeCsv("weekend_weekday_gap.csv", weTableRows());
}

// ==========================================================
// Patterns tab — classify every meter by its weekday/weekend behavior
// ==========================================================
const PATTERN_ORDER = ["Weekday-only", "Weekend-only", "Always-on", "Never active", "Weekend-drop", "Weekend-boost", "Insufficient data"];
const PATTERN_SHORT = {
  "Weekday-only": "Wkday-only", "Weekend-only": "Wkend-only", "Always-on": "Always-on",
  "Never active": "Never", "Weekend-drop": "Wkend-drop", "Weekend-boost": "Wkend-boost",
  "Insufficient data": "No data",
};
const GAP_THRESHOLD = 0.3; // 30 percentage points

// classify a single meter's weekday/weekend record (from WSTATS.perMeter)
const PURE_SIDE_MIN_RATE = 0.75; // Weekday-only/Weekend-only requires the active side to be meaningfully high, not just nonzero
function classifyMeter(r) {
  const wd = r.wdT + r.wdF, we = r.weT + r.weF;
  // Check 100% uptime first, regardless of weekday/weekend data distribution
  if (r.uptime === 1) return "Always-on";
  if (wd === 0 || we === 0) return "Insufficient data"; // no data on one side within window
  if (r.wdRate === 0 && r.weRate === 0) return "Never active";
  if (r.wdRate >= PURE_SIDE_MIN_RATE && r.weRate === 0) return "Weekday-only";
  if (r.weRate >= PURE_SIDE_MIN_RATE && r.wdRate === 0) return "Weekend-only";
  if (r.gap >= GAP_THRESHOLD) return "Weekend-drop";
  if (r.gap <= -GAP_THRESHOLD) return "Weekend-boost";
  return "Insufficient data"; // fallback for mixed patterns that don't fit above categories
}

let PSTATS = null;
function computePatternStats(weekStats = WSTATS) {
  const perMeter = weekStats.perMeter.map((r) => ({ ...r, category: classifyMeter(r) }));
  const counts = {};
  PATTERN_ORDER.forEach((c) => (counts[c] = 0));
  const byCategoryRegion = new Map();
  for (const r of perMeter) {
    counts[r.category]++;
    let m = byCategoryRegion.get(r.category);
    if (!m) { m = new Map(); byCategoryRegion.set(r.category, m); }
    m.set(r.m.region, (m.get(r.m.region) || 0) + 1);
  }
  return { perMeter, counts, byCategoryRegion };
}

function renderPatterns() {
  PSTATS = computePatternStats();
  renderPatternKpis();
  renderPatternDist();
  renderPatternRegion();
  renderPatternMeterTable();
  renderRhythm();
}

function renderPatternKpis() {
  const c = PSTATS.counts;
  const classified = PATTERN_ORDER.reduce((a, k) => a + (k === "Insufficient data" ? 0 : c[k]), 0);
  const pctOf = (n) => (classified > 0 ? pct(n / classified) : "—");
  $("patKpiRow").innerHTML = [
    tile("Weekday-only meters", fmt(c["Weekday-only"]), `${pctOf(c["Weekday-only"])} of classified meters — ≥${PURE_SIDE_MIN_RATE * 100}% weekday comm rate, silent every weekend`, c["Weekday-only"] > 0 ? "bad" : "good", "Weekday-only"),
    tile("Weekend-only meters", fmt(c["Weekend-only"]), `${pctOf(c["Weekend-only"])} — ≥${PURE_SIDE_MIN_RATE * 100}% weekend comm rate, silent all week`, c["Weekend-only"] > 0 ? "bad" : "good", "Weekend-only"),
    tile("Always-on (7/7)", fmt(c["Always-on"]), `${pctOf(c["Always-on"])} — 100% uptime every weekday and weekend`, "good", "Always-on"),
    tile("Never active", fmt(c["Never active"]), `${pctOf(c["Never active"])} — zero TRUE days in the window`, c["Never active"] > 0 ? "bad" : "good", "Never active"),
    tile(`Weekend-drop (≥${GAP_THRESHOLD * 100}pp)`, fmt(c["Weekend-drop"]), `${pctOf(c["Weekend-drop"])} — meaningfully worse on weekends`, c["Weekend-drop"] > 0 ? "bad" : "good", "Weekend-drop"),
    tile(`Weekend-boost (≥${GAP_THRESHOLD * 100}pp)`, fmt(c["Weekend-boost"]), `${pctOf(c["Weekend-boost"])} — meaningfully better on weekends`, undefined, "Weekend-boost"),
  ].join("");
}
function selectPatternCategory(category) {
  patternCategory = category;
  $("patCategorySelect").value = patternCategory;
  patTableLimit = 25; patTableSearch = ""; $("patMeterSearch").value = "";
  renderPatternRegion();
  renderPatternMeterTable();
  patExportCsv();
}
wireKpiRow("patKpiRow", (key) => {
  if (!PSTATS) return;
  selectPatternCategory(key);
});

function renderPatternDist() {
  const counts = PATTERN_ORDER.map((c) => PSTATS.counts[c] || 0);
  if (tableModes["patdist"]) {
    $("chartPatDist").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Pattern</th><th class="num">Meters</th></tr></thead>
      <tbody>${PATTERN_ORDER.map((c, i) => `<tr><td>${c}</td><td class="num">${fmt(counts[i])}</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  const W = Math.max(560, Math.min(1000, $("chartPatDist").clientWidth || 800));
  const H = 260, mL = 52, mR = 14, mT = 12, mB = 46;
  const iw = W - mL - mR, ih = H - mT - mB;
  const max = Math.max(...counts, 1);
  const ticks = niceTicks(max);
  const yv = (v) => mT + ih - (v / ticks[ticks.length - 1]) * ih;
  const slotW = iw / counts.length, barW = Math.min(40, slotW * 0.55);

  let g = "";
  for (const t of ticks) {
    g += `<line x1="${mL}" y1="${yv(t)}" x2="${W - mR}" y2="${yv(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${yv(t) + 4}" text-anchor="end">${fmt(t)}</text>`;
  }
  let bars = "";
  counts.forEach((v, i) => {
    const cx = mL + slotW * i + slotW / 2;
    const yTop = yv(v), h = mT + ih - yTop;
    const r = Math.min(4, h);
    bars += `<path d="M${(cx - barW / 2).toFixed(1)} ${(mT + ih).toFixed(1)} v-${(h - r).toFixed(1)} a${r} ${r} 0 0 1 ${r} -${r} h${(barW - 2 * r).toFixed(1)} a${r} ${r} 0 0 1 ${r} ${r} v${(h - r).toFixed(1)} Z" fill="var(--series-1)"/>`;
    bars += `<text class="axis-text" x="${cx}" y="${H - 28}" text-anchor="middle">${PATTERN_SHORT[PATTERN_ORDER[i]]}</text>`;
    bars += `<rect x="${(cx - slotW / 2).toFixed(1)}" y="${mT}" width="${slotW.toFixed(1)}" height="${ih}" fill="transparent" class="patcol" data-i="${i}"/>`;
  });
  bars += `<line x1="${mL}" y1="${mT + ih}" x2="${W - mR}" y2="${mT + ih}" stroke="var(--baseline)" stroke-width="1"/>`;

  $("chartPatDist").innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${g}${bars}</svg>`;
  $("chartPatDist").querySelectorAll(".patcol").forEach((el) => {
    el.style.cursor = "pointer";
    const i = +el.dataset.i;
    el.addEventListener("mousemove", (e) => {
      showTip(`<div class="tt-title">${PATTERN_ORDER[i]}</div><div class="tt-row"><span>Meters</span><b>${fmt(counts[i])}</b></div><div class="tt-row"><i>Click bar to export CSV</i></div>`, e.clientX, e.clientY);
    });
    el.addEventListener("mouseleave", hideTip);
    el.addEventListener("click", () => selectPatternCategory(PATTERN_ORDER[i]));
  });
}

// count-scaled horizontal bars (labels + absolute counts, not percentages)
function countBars(container, items, onClick) {
  const max = Math.max(1, ...items.map((i) => i.count));
  const W = Math.max(420, Math.min(660, $(container).clientWidth || 560));
  const barH = 18, gap = 14, labelW = 130, mT = 6, mB = 24, mR = 60;
  const H = mT + items.length * (barH + gap) + mB;
  const iw = W - labelW - mR;
  const ticks = niceTicks(max);
  const xv = (v) => labelW + (v / ticks[ticks.length - 1]) * iw;

  let g = "";
  for (const t of ticks) {
    g += `<line x1="${xv(t)}" y1="${mT}" x2="${xv(t)}" y2="${H - mB}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${xv(t)}" y="${H - 8}" text-anchor="middle">${fmt(t)}</text>`;
  }
  let bars = "";
  items.forEach((it, i) => {
    const yTop = mT + i * (barH + gap);
    const w = Math.max(0, xv(it.count) - labelW);
    const r = Math.min(4, w);
    bars += `<path d="M${labelW} ${yTop} h${(w - r).toFixed(1)} a${r} ${r} 0 0 1 ${r} ${r} v${barH - 2 * r} a${r} ${r} 0 0 1 -${r} ${r} h-${(w - r).toFixed(1)} Z" fill="var(--series-1)"/>`;
    bars += `<text class="bar-label" x="${labelW - 8}" y="${yTop + barH / 2 + 4}" text-anchor="end">${esc(trunc(it.label, 18))}</text>`;
    bars += `<text class="bar-label" x="${(labelW + w + 6).toFixed(1)}" y="${yTop + barH / 2 + 4}">${fmt(it.count)}</text>`;
    bars += `<rect x="0" y="${yTop - gap / 2}" width="${W}" height="${barH + gap}" fill="transparent" data-i="${i}" class="hrow"/>`;
  });
  $(container).innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${g}${bars}</svg>`;
  $(container).querySelectorAll(".hrow").forEach((el) => {
    const it = items[+el.dataset.i];
    el.addEventListener("mousemove", (e) => {
      showTip(`<div class="tt-title">${esc(it.label)}</div><div class="tt-row"><span>Meters</span><b>${fmt(it.count)}</b></div>${onClick ? `<div class="tt-row"><i>Click bar to export CSV</i></div>` : ""}`, e.clientX, e.clientY);
    });
    el.addEventListener("mouseleave", hideTip);
    if (onClick) { el.style.cursor = "pointer"; el.addEventListener("click", () => onClick(it)); }
  });
}

function renderPatternRegion() {
  const map = PSTATS.byCategoryRegion.get(patternCategory) || new Map();
  const items = [...map.entries()].map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count).slice(0, 10);
  $("cardPatRegionSub").textContent = `${fmt(PSTATS.counts[patternCategory] || 0)} meters classified as "${patternCategory}" — top regions`;
  if (tableModes["patregion"]) {
    $("chartPatRegion").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Region</th><th class="num">Meters</th></tr></thead>
      <tbody>${items.map((it) => `<tr><td>${esc(it.label)}</td><td class="num">${fmt(it.count)}</td></tr>`).join("")}</tbody></table></div>`;
  } else if (items.length === 0) {
    $("chartPatRegion").innerHTML = `<p class="card-sub">No meters in this pattern for the current filters.</p>`;
  } else {
    countBars("chartPatRegion", items, (it) => {
      const rows = PSTATS.perMeter.filter((r) => r.category === patternCategory && r.m.region === it.label);
      downloadCsv(`pattern_${slug(patternCategory)}_${slug(it.label)}.csv`, PAT_TABLE_COLS.map((c) => c.label).concat("Pattern"), patRowsToCells(rows));
    });
  }
}

$("patCategorySelect").addEventListener("change", () => {
  patternCategory = $("patCategorySelect").value;
  patTableLimit = 25; patTableSearch = ""; $("patMeterSearch").value = "";
  renderPatternRegion();
  renderPatternMeterTable();
});

// ---------- pattern meter table ----------
const PAT_TABLE_COLS = [
  { key: "ip", label: "IP", num: false },
  { key: "tsp", label: "TSP", num: false },
  { key: "region", label: "Region", num: false },
  { key: "circle", label: "Circle", num: false },
  { key: "po", label: "PO", num: false },
  { key: "wdRate", label: "Weekday %", num: true },
  { key: "weRate", label: "Weekend %", num: true },
  { key: "gap", label: "Gap (pp)", num: true },
];

$("patMeterSearch").addEventListener("input", (e) => { patTableSearch = e.target.value.trim().toLowerCase(); patTableLimit = 25; renderPatternMeterTable(); });
$("patLoadMore").addEventListener("click", () => { patTableLimit += 50; renderPatternMeterTable(); });
$("patExportCsv").addEventListener("click", patExportCsv);

function patTableRows() {
  let rows = PSTATS.perMeter.filter((r) => r.category === patternCategory);
  if (patTableSearch) rows = rows.filter((r) => r.m.ip.toLowerCase().includes(patTableSearch));
  const k = patTableSort.key, dir = patTableSort.dir;
  const val = (r) => {
    if (k === "ip" || k === "tsp" || k === "region" || k === "circle" || k === "po") return r.m[k];
    return r[k] === null ? -Infinity : r[k];
  };
  rows.sort((a, b) => {
    const va = val(a), vb = val(b);
    if (typeof va === "string") return dir * va.localeCompare(vb);
    return dir * (va - vb);
  });
  return rows;
}

function renderPatternMeterTable() {
  $("patMeterTableHead").innerHTML = PAT_TABLE_COLS.map((c) =>
    `<th class="${c.num ? "num" : ""} ${patTableSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${patTableSort.key === c.key ? (patTableSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  $("patMeterTableHead").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (patTableSort.key === key) patTableSort.dir *= -1;
      else patTableSort = { key, dir: -1 };
      renderPatternMeterTable();
    });
  });
  const rows = patTableRows();
  const shown = rows.slice(0, patTableLimit);
  $("patMeterTableBody").innerHTML = shown.map((r) => `<tr>
    <td>${esc(r.m.ip)}</td><td>${esc(r.m.tsp)}</td><td>${esc(r.m.region)}</td>
    <td>${esc(r.m.circle)}</td><td>${esc(r.m.po)}</td>
    <td class="num">${r.wdRate === null ? "—" : pct(r.wdRate)}</td>
    <td class="num">${r.weRate === null ? "—" : pct(r.weRate)}</td>
    <td class="num">${r.gap === null ? "—" : (r.gap * 100).toFixed(1)}</td>
  </tr>`).join("");
  $("patTableCount").textContent = `Showing ${fmt(shown.length)} of ${fmt(rows.length)} meters in "${patternCategory}"`;
  $("patLoadMore").style.visibility = shown.length < rows.length ? "visible" : "hidden";
}

// ==========================================================
// Communication rhythm — for meters with both TRUE and FALSE days,
// classify the shape of their on/off sequence (bursty, sporadic, alternating...)
// ==========================================================
const RHYTHM_ORDER = ["Alternating", "Long burst", "Sporadic", "Stable (occasional drop)", "Intermittent null", "Irregular"];
const RHYTHM_SHORT = { "Alternating": "Alternating", "Long burst": "Long burst", "Sporadic": "Sporadic", "Stable (occasional drop)": "Stable+drop", "Intermittent null": "Interm. null", "Irregular": "Irregular" };
const ALT_RUN_THRESHOLD = 1.8;   // avg run length <= this => flips almost every day
const BURST_RUN_THRESHOLD = 4;   // avg run length >= this => long contiguous on/off blocks
const BLIP_RUN_THRESHOLD = 2;    // isolated blips are <= this many days long
const SPORADIC_RATE = 0.35;      // mostly-off threshold
const STABLE_RATE = 0.65;        // mostly-on threshold

// run-length encode a meter's data-days (skipping NULLs) over [i0,i1]
function computeRunStats(days, i0, i1) {
  let curVal = null, curLen = 0, seqLen = 0, transitions = 0;
  let trueRunSum = 0, trueRunCount = 0, falseRunSum = 0, falseRunCount = 0, maxTrueRun = 0, maxFalseRun = 0;
  const closeRun = () => {
    if (curVal === 1) { trueRunSum += curLen; trueRunCount++; if (curLen > maxTrueRun) maxTrueRun = curLen; }
    else if (curVal === 0) { falseRunSum += curLen; falseRunCount++; if (curLen > maxFalseRun) maxFalseRun = curLen; }
  };
  for (let j = i0; j <= i1; j++) {
    const v = days[j];
    if (v === -1) continue;
    seqLen++;
    if (curVal === null) { curVal = v; curLen = 1; }
    else if (v === curVal) curLen++;
    else { closeRun(); transitions++; curVal = v; curLen = 1; }
  }
  if (curVal !== null) closeRun();
  const nRuns = trueRunCount + falseRunCount;

  // separately: find NULL gaps sitting between two data days (not a leading/trailing
  // gap). The run-length pass above skips NULLs entirely, so a TRUE...NULL...TRUE
  // sequence reads as one continuous TRUE run there — invisible to the categories
  // above. This is exactly what "Intermittent null" looks for. Only gaps starting
  // after INTERMITTENT_NULL_CUTOFF count, per the Dec-2025 cutoff for this rhythm.
  let sawData = false, nullRunLen = 0, nullRunStart = -1, maxInteriorNullRun = 0, interiorNullRunCount = 0;
  for (let j = i0; j <= i1; j++) {
    const v = days[j];
    if (v === -1) { if (sawData) { if (nullRunLen === 0) nullRunStart = j; nullRunLen++; } }
    else {
      if (nullRunLen > 0 && DATA.dates[nullRunStart] > INTERMITTENT_NULL_CUTOFF) {
        interiorNullRunCount++; if (nullRunLen > maxInteriorNullRun) maxInteriorNullRun = nullRunLen;
      }
      nullRunLen = 0; sawData = true;
    }
  }

  return {
    seqLen, transitions,
    avgRunLen: nRuns > 0 ? seqLen / nRuns : 0,
    avgTrueRun: trueRunCount ? trueRunSum / trueRunCount : 0,
    avgFalseRun: falseRunCount ? falseRunSum / falseRunCount : 0,
    maxTrueRun, maxFalseRun,
    maxInteriorNullRun, interiorNullRunCount,
  };
}

// same run-length encoding as computeRunStats(), but driven by an arbitrary index list
function computeRunStatsForIndices(days, indices) {
  let curVal = null, curLen = 0, seqLen = 0, transitions = 0;
  let trueRunSum = 0, trueRunCount = 0, falseRunSum = 0, falseRunCount = 0, maxTrueRun = 0, maxFalseRun = 0;
  const closeRun = () => {
    if (curVal === 1) { trueRunSum += curLen; trueRunCount++; if (curLen > maxTrueRun) maxTrueRun = curLen; }
    else if (curVal === 0) { falseRunSum += curLen; falseRunCount++; if (curLen > maxFalseRun) maxFalseRun = curLen; }
  };
  for (const j of indices) {
    const v = days[j];
    if (v === -1) continue;
    seqLen++;
    if (curVal === null) { curVal = v; curLen = 1; }
    else if (v === curVal) curLen++;
    else { closeRun(); transitions++; curVal = v; curLen = 1; }
  }
  if (curVal !== null) closeRun();
  const nRuns = trueRunCount + falseRunCount;

  let sawData = false, nullRunLen = 0, nullRunStart = -1, maxInteriorNullRun = 0, interiorNullRunCount = 0;
  for (const j of indices) {
    const v = days[j];
    if (v === -1) { if (sawData) { if (nullRunLen === 0) nullRunStart = j; nullRunLen++; } }
    else {
      if (nullRunLen > 0 && DATA.dates[nullRunStart] > INTERMITTENT_NULL_CUTOFF) {
        interiorNullRunCount++; if (nullRunLen > maxInteriorNullRun) maxInteriorNullRun = nullRunLen;
      }
      nullRunLen = 0; sawData = true;
    }
  }

  return {
    seqLen, transitions,
    avgRunLen: nRuns > 0 ? seqLen / nRuns : 0,
    avgTrueRun: trueRunCount ? trueRunSum / trueRunCount : 0,
    avgFalseRun: falseRunCount ? falseRunSum / falseRunCount : 0,
    maxTrueRun, maxFalseRun,
    maxInteriorNullRun, interiorNullRunCount,
  };
}

const INTERMITTENT_NULL_MIN_GAP = 3; // interior NULL gap of at least this many days = a deliberate dropout, not one missed reading
const INTERMITTENT_NULL_CUTOFF = new Date(2025, 11, 31); // only NULL gaps starting after Dec 31, 2025 count toward this rhythm

function classifyRhythm(t, f, rs) {
  const trueRate = t / (t + f);
  if (rs.interiorNullRunCount > 0 && rs.maxInteriorNullRun >= INTERMITTENT_NULL_MIN_GAP) return "Intermittent null";
  if (rs.avgRunLen <= ALT_RUN_THRESHOLD) return "Alternating";
  if (trueRate <= SPORADIC_RATE && rs.avgTrueRun <= BLIP_RUN_THRESHOLD) return "Sporadic";
  if (trueRate >= STABLE_RATE && rs.avgFalseRun <= BLIP_RUN_THRESHOLD) return "Stable (occasional drop)";
  if (rs.avgTrueRun >= BURST_RUN_THRESHOLD && rs.avgFalseRun >= BURST_RUN_THRESHOLD) return "Long burst";
  return "Irregular";
}

let RSTATS = null;
let rhythmCategory = "Alternating";
let rhTableSort = { key: "avgRunLen", dir: -1 };
let rhTableLimit = 25;
let rhTableSearch = "";

function computeRhythmStats(baseStats = STATS, runStatsOf = (days) => computeRunStats(days, baseStats.i0, baseStats.i1)) {
  const perMeter = [];
  const counts = {}; RHYTHM_ORDER.forEach((c) => (counts[c] = 0));
  const byCategoryRegion = new Map();
  for (const r of baseStats.perMeter) {
    if (r.t === 0 || r.f === 0) continue; // pure meters have no rhythm to classify
    const rs = runStatsOf(r.m.days);
    const category = classifyRhythm(r.t, r.f, rs);
    const rec = { m: r.m, t: r.t, f: r.f, uptime: r.uptime, rs, category };
    perMeter.push(rec);
    counts[category]++;
    let mm = byCategoryRegion.get(category);
    if (!mm) { mm = new Map(); byCategoryRegion.set(category, mm); }
    mm.set(r.m.region, (mm.get(r.m.region) || 0) + 1);
  }
  return { perMeter, counts, byCategoryRegion, totalMixed: perMeter.length };
}

function renderRhythm() {
  RSTATS = computeRhythmStats();
  renderRhythmKpis();
  renderRhythmDist();
  renderRhythmRegion();
  renderRhythmMeterTable();
}

function renderRhythmKpis() {
  const c = RSTATS.counts, total = RSTATS.totalMixed;
  const pctOf = (n) => (total > 0 ? pct(n / total) : "—");
  $("rhKpiRow").innerHTML = [
    tile("Alternating", fmt(c["Alternating"]), `${pctOf(c["Alternating"])} of mixed meters — flips on/off almost daily`, undefined, "Alternating"),
    tile("Long burst", fmt(c["Long burst"]), `${pctOf(c["Long burst"])} — communicates in extended multi-day blocks`, undefined, "Long burst"),
    tile("Sporadic", fmt(c["Sporadic"]), `${pctOf(c["Sporadic"])} — mostly silent, rare isolated TRUE blips`, c["Sporadic"] > 0 ? "bad" : "good", "Sporadic"),
    tile("Stable (occasional drop)", fmt(c["Stable (occasional drop)"]), `${pctOf(c["Stable (occasional drop)"])} — mostly on, rare isolated outages`, "good", "Stable (occasional drop)"),
    tile("Intermittent null", fmt(c["Intermittent null"]), `${pctOf(c["Intermittent null"])} — comm/non-comm days with a ${INTERMITTENT_NULL_MIN_GAP}+ day NULL gap in between, then data resumes`, c["Intermittent null"] > 0 ? "bad" : "good", "Intermittent null"),
    tile("Irregular", fmt(c["Irregular"]), `${pctOf(c["Irregular"])} — no clean rhythm detected`, undefined, "Irregular"),
  ].join("");
}
function selectRhythmCategory(category) {
  rhythmCategory = category;
  $("rhCategorySelect").value = rhythmCategory;
  rhTableLimit = 25; rhTableSearch = ""; $("rhMeterSearch").value = "";
  renderRhythmRegion();
  renderRhythmMeterTable();
  rhExportCsv();
}
wireKpiRow("rhKpiRow", (key) => {
  if (!RSTATS) return;
  selectRhythmCategory(key);
});

function renderRhythmDist() {
  const counts = RHYTHM_ORDER.map((c) => RSTATS.counts[c] || 0);
  if (tableModes["rhdist"]) {
    $("chartRhDist").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Rhythm</th><th class="num">Meters</th></tr></thead>
      <tbody>${RHYTHM_ORDER.map((c, i) => `<tr><td>${c}</td><td class="num">${fmt(counts[i])}</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  const W = Math.max(500, Math.min(900, $("chartRhDist").clientWidth || 700));
  const H = 260, mL = 52, mR = 14, mT = 12, mB = 40;
  const iw = W - mL - mR, ih = H - mT - mB;
  const max = Math.max(...counts, 1);
  const ticks = niceTicks(max);
  const yv = (v) => mT + ih - (v / ticks[ticks.length - 1]) * ih;
  const slotW = iw / counts.length, barW = Math.min(48, slotW * 0.55);

  let g = "";
  for (const t of ticks) {
    g += `<line x1="${mL}" y1="${yv(t)}" x2="${W - mR}" y2="${yv(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${yv(t) + 4}" text-anchor="end">${fmt(t)}</text>`;
  }
  let bars = "";
  counts.forEach((v, i) => {
    const cx = mL + slotW * i + slotW / 2;
    const yTop = yv(v), h = mT + ih - yTop;
    const r = Math.min(4, h);
    bars += `<path d="M${(cx - barW / 2).toFixed(1)} ${(mT + ih).toFixed(1)} v-${(h - r).toFixed(1)} a${r} ${r} 0 0 1 ${r} -${r} h${(barW - 2 * r).toFixed(1)} a${r} ${r} 0 0 1 ${r} ${r} v${(h - r).toFixed(1)} Z" fill="var(--series-1)"/>`;
    bars += `<text class="axis-text" x="${cx}" y="${H - 20}" text-anchor="middle">${RHYTHM_SHORT[RHYTHM_ORDER[i]]}</text>`;
    bars += `<rect x="${(cx - slotW / 2).toFixed(1)}" y="${mT}" width="${slotW.toFixed(1)}" height="${ih}" fill="transparent" class="rhcol" data-i="${i}"/>`;
  });
  bars += `<line x1="${mL}" y1="${mT + ih}" x2="${W - mR}" y2="${mT + ih}" stroke="var(--baseline)" stroke-width="1"/>`;

  $("chartRhDist").innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${g}${bars}</svg>`;
  $("chartRhDist").querySelectorAll(".rhcol").forEach((el) => {
    el.style.cursor = "pointer";
    const i = +el.dataset.i;
    el.addEventListener("mousemove", (e) => {
      showTip(`<div class="tt-title">${RHYTHM_ORDER[i]}</div><div class="tt-row"><span>Meters</span><b>${fmt(counts[i])}</b></div><div class="tt-row"><i>Click bar to export CSV</i></div>`, e.clientX, e.clientY);
    });
    el.addEventListener("mouseleave", hideTip);
    el.addEventListener("click", () => selectRhythmCategory(RHYTHM_ORDER[i]));
  });
}

function renderRhythmRegion() {
  const map = RSTATS.byCategoryRegion.get(rhythmCategory) || new Map();
  const items = [...map.entries()].map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count).slice(0, 10);
  $("cardRhRegionSub").textContent = `${fmt(RSTATS.counts[rhythmCategory] || 0)} meters classified as "${rhythmCategory}" — top regions`;
  if (tableModes["rhregion"]) {
    $("chartRhRegion").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Region</th><th class="num">Meters</th></tr></thead>
      <tbody>${items.map((it) => `<tr><td>${esc(it.label)}</td><td class="num">${fmt(it.count)}</td></tr>`).join("")}</tbody></table></div>`;
  } else if (items.length === 0) {
    $("chartRhRegion").innerHTML = `<p class="card-sub">No meters in this rhythm for the current filters.</p>`;
  } else {
    countBars("chartRhRegion", items, (it) => {
      const rows = RSTATS.perMeter.filter((r) => r.category === rhythmCategory && r.m.region === it.label);
      downloadCsv(`rhythm_${slug(rhythmCategory)}_${slug(it.label)}.csv`, RH_TABLE_COLS.map((c) => c.label).concat("Rhythm"), rhRowsToCells(rows));
    });
  }
}

$("rhCategorySelect").addEventListener("change", () => {
  rhythmCategory = $("rhCategorySelect").value;
  rhTableLimit = 25; rhTableSearch = ""; $("rhMeterSearch").value = "";
  renderRhythmRegion();
  renderRhythmMeterTable();
});

// ---------- rhythm meter table ----------
const RH_TABLE_COLS = [
  { key: "ip", label: "IP", num: false },
  { key: "tsp", label: "TSP", num: false },
  { key: "region", label: "Region", num: false },
  { key: "circle", label: "Circle", num: false },
  { key: "po", label: "PO", num: false },
  { key: "uptime", label: "Uptime %", num: true },
  { key: "avgRunLen", label: "Avg run (days)", num: true },
  { key: "transitions", label: "Transitions", num: true },
  { key: "maxTrueRun", label: "Longest TRUE run", num: true },
  { key: "maxFalseRun", label: "Longest FALSE run", num: true },
];

$("rhMeterSearch").addEventListener("input", (e) => { rhTableSearch = e.target.value.trim().toLowerCase(); rhTableLimit = 25; renderRhythmMeterTable(); });
$("rhLoadMore").addEventListener("click", () => { rhTableLimit += 50; renderRhythmMeterTable(); });
$("rhExportCsv").addEventListener("click", rhExportCsv);

function rhTableRows() {
  let rows = RSTATS.perMeter.filter((r) => r.category === rhythmCategory);
  if (rhTableSearch) rows = rows.filter((r) => r.m.ip.toLowerCase().includes(rhTableSearch));
  const k = rhTableSort.key, dir = rhTableSort.dir;
  const val = (r) => {
    if (k === "ip" || k === "tsp" || k === "region" || k === "circle" || k === "po") return r.m[k];
    if (k === "uptime") return r.uptime;
    return r.rs[k];
  };
  rows.sort((a, b) => {
    const va = val(a), vb = val(b);
    if (typeof va === "string") return dir * va.localeCompare(vb);
    return dir * (va - vb);
  });
  return rows;
}

function renderRhythmMeterTable() {
  $("rhMeterTableHead").innerHTML = RH_TABLE_COLS.map((c) =>
    `<th class="${c.num ? "num" : ""} ${rhTableSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${rhTableSort.key === c.key ? (rhTableSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  $("rhMeterTableHead").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (rhTableSort.key === key) rhTableSort.dir *= -1;
      else rhTableSort = { key, dir: -1 };
      renderRhythmMeterTable();
    });
  });
  const rows = rhTableRows();
  const shown = rows.slice(0, rhTableLimit);
  $("rhMeterTableBody").innerHTML = shown.map((r) => `<tr>
    <td>${esc(r.m.ip)}</td><td>${esc(r.m.tsp)}</td><td>${esc(r.m.region)}</td>
    <td>${esc(r.m.circle)}</td><td>${esc(r.m.po)}</td>
    <td class="num">${pct(r.uptime)}</td>
    <td class="num">${r.rs.avgRunLen.toFixed(1)}</td>
    <td class="num">${fmt(r.rs.transitions)}</td>
    <td class="num">${fmt(r.rs.maxTrueRun)}</td>
    <td class="num">${fmt(r.rs.maxFalseRun)}</td>
  </tr>`).join("");
  $("rhTableCount").textContent = `Showing ${fmt(shown.length)} of ${fmt(rows.length)} meters in "${rhythmCategory}"`;
  $("rhLoadMore").style.visibility = shown.length < rows.length ? "visible" : "hidden";
}

function rhRowsToCells(rows) {
  return rows.map((r) => [
    r.m.ip, r.m.tsp, r.m.region, r.m.circle, r.m.po,
    (r.uptime * 100).toFixed(2), r.rs.avgRunLen.toFixed(2), r.rs.transitions, r.rs.maxTrueRun, r.rs.maxFalseRun,
    r.category,
  ]);
}
function rhExportCsv() {
  downloadCsv(`rhythm_${slug(rhythmCategory)}.csv`, RH_TABLE_COLS.map((c) => c.label).concat("Rhythm"), rhRowsToCells(rhTableRows()));
}

function patRowsToCells(rows) {
  return rows.map((r) => [
    r.m.ip, r.m.tsp, r.m.region, r.m.circle, r.m.po,
    r.wdRate === null ? "" : (r.wdRate * 100).toFixed(2),
    r.weRate === null ? "" : (r.weRate * 100).toFixed(2),
    r.gap === null ? "" : (r.gap * 100).toFixed(2),
    r.category,
  ]);
}
function patExportCsv() {
  downloadCsv(`pattern_${slug(patternCategory)}.csv`, PAT_TABLE_COLS.map((c) => c.label).concat("Pattern"), patRowsToCells(patTableRows()));
}

// ==========================================================
// Date Range tab — arbitrary custom start/end window, all analytics rescoped to it
// ==========================================================
let DRSTATS = null;
let DR_IMPROVEMENT = null;

// split [i0,i1] at its midpoint and compare overall comm rate of each half —
// a simple, range-length-agnostic way to answer "has this improved over time?"
function computeImprovement(meters, i0, i1) {
  const mid = i0 + Math.floor((i1 - i0 + 1) / 2);
  if (mid <= i0 || mid > i1) return { firstRate: null, secondRate: null, deltaPp: null, firstStats: null, secondStats: null };
  const firstStats = computeStats(meters, i0, mid - 1);
  const secondStats = computeStats(meters, mid, i1);
  const firstRate = firstStats.totT + firstStats.totF > 0 ? firstStats.totT / (firstStats.totT + firstStats.totF) : null;
  const secondRate = secondStats.totT + secondStats.totF > 0 ? secondStats.totT / (secondStats.totT + secondStats.totF) : null;
  const deltaPp = firstRate !== null && secondRate !== null ? (secondRate - firstRate) * 100 : null;
  return { firstRate, secondRate, deltaPp, firstStats, secondStats };
}

function exportImprovementCsv() {
  if (!DR_IMPROVEMENT || !DR_IMPROVEMENT.firstStats) return;
  const { firstStats, secondStats } = DR_IMPROVEMENT;
  const rows = firstStats.perMeter.map((r1, idx) => {
    const r2 = secondStats.perMeter[idx];
    const delta = r1.uptime >= 0 && r2.uptime >= 0 ? r2.uptime - r1.uptime : null;
    return { m: r1.m, u1: r1.uptime, u2: r2.uptime, delta };
  });
  rows.sort((a, b) => (b.delta ?? -Infinity) - (a.delta ?? -Infinity));
  downloadCsv(
    "date_range_improvement.csv",
    ["IP", "TSP", "Region", "Circle", "PO", "First-half uptime %", "Second-half uptime %", "Change (pp)"],
    rows.map((r) => [
      r.m.ip, r.m.tsp, r.m.region, r.m.circle, r.m.po,
      r.u1 < 0 ? "" : (r.u1 * 100).toFixed(2),
      r.u2 < 0 ? "" : (r.u2 * 100).toFixed(2),
      r.delta === null ? "" : (r.delta * 100).toFixed(2),
    ])
  );
}

function computeWindowFromInputs(startId, endId) {
  const startVal = $(startId).value, endVal = $(endId).value;
  if (!startVal || !endVal) return null;
  const start = new Date(startVal + "T00:00:00");
  const end = new Date(endVal + "T00:00:00");
  if (isNaN(start) || isNaN(end) || start > end) return null;
  let i0 = DATA.dates.findIndex((d) => d >= start);
  if (i0 === -1) i0 = 0;
  let i1 = -1;
  for (let i = DATA.dates.length - 1; i >= 0; i--) { if (DATA.dates[i] <= end) { i1 = i; break; } }
  if (i1 === -1 || i1 < i0) return null;
  return [i0, i1];
}
function computeCustomWindow() {
  return computeWindowFromInputs("drStart", "drEnd");
}

["drStart", "drEnd"].forEach((id) => $(id).addEventListener("change", () => { if (currentTab === "daterange") renderDateRange(); }));

function renderDateRange() {
  const win = computeCustomWindow();
  const errEl = $("drRangeError");
  if (!win) {
    errEl.textContent = `No data in the selected range — pick a start/end date between ${dstr(DATA.dates[0])} and ${dstr(DATA.dates[DATA.dates.length - 1])}, with start on or before end.`;
    errEl.classList.remove("hidden");
    $("drContent").classList.add("hidden");
    return;
  }
  errEl.classList.add("hidden");
  $("drContent").classList.remove("hidden");
  const [i0, i1] = win;
  const meters = filteredMeters();
  DRSTATS = computeStats(meters, i0, i1);
  DR_IMPROVEMENT = computeImprovement(meters, i0, i1);
  $("drRangeMeta").textContent = `${fmt(meters.length)} meters in view · ${dstr(DATA.dates[i0])} → ${dstr(DATA.dates[i1])} (${i1 - i0 + 1} days)`;
  $("drRangeBanner").innerHTML = `Showing data for ${dstr(DATA.dates[i0])} → ${dstr(DATA.dates[i1])} <span class="rb-note">(${fmt(i1 - i0 + 1)} days · this range is independent of the other tabs)</span>`;
  renderKpisInto(DRSTATS, "drKpiRow", DR_IMPROVEMENT);
  renderTrendFor(DRSTATS, "chartDrTrend", "drtrend");
  renderGroupBar("tsp", "chartDrTsp", 99, DRSTATS, "drtsp");
  renderHealth(DRSTATS, "chartDrHealth", "drhealth");
  renderGroupBar("region", "chartDrRegion", 99, DRSTATS, "drregion");
  renderGroupBar("circle", "chartDrCircle", 15, DRSTATS, "drcircle");
  renderGroupBar("po", "chartDrPo", 99, DRSTATS, "drpo");
  drTableLimit = 25;
  drRenderMeterTable();
}

// simplified single-series trend line (no Overall/By-TSP toggle) for a custom window or index list
function renderTrendFor(s, container, cardKey, dateAt) {
  const at = dateAt || ((k) => DATA.dates[s.i0 + k]);
  const vals = [];
  for (let k = 0; k < s.nDays; k++) {
    const denom = s.dayTrue[k] + s.dayFalse[k];
    vals.push(denom > 0 ? s.dayTrue[k] / denom : null);
  }
  if (tableModes[cardKey]) {
    let rows = "";
    for (let k = 0; k < s.nDays; k++) {
      const denom = s.dayTrue[k] + s.dayFalse[k];
      rows += `<tr><td>${dstr(at(k))}</td><td class="num">${vals[k] === null ? "—" : pct(vals[k])}</td><td class="num">${fmt(denom)}</td></tr>`;
    }
    $(container).innerHTML = `<div class="table-scroll mini-table" style="max-height:320px;overflow-y:auto"><table>
      <thead><tr><th>Date</th><th class="num">Comm rate</th><th class="num">Meters with data</th></tr></thead><tbody>${rows}</tbody></table></div>`;
    return;
  }
  const W = Math.max(680, Math.min(1300, $(container).clientWidth || 900));
  const H = 260, mL = 46, mR = 16, mT = 12, mB = 34;
  const iw = W - mL - mR, ih = H - mT - mB;
  const n = s.nDays;
  const x = (k) => mL + (n === 1 ? iw / 2 : (k / (n - 1)) * iw);
  const y = (v) => mT + (1 - v) * ih;

  let g = "";
  for (const t of [0, 0.25, 0.5, 0.75, 1]) {
    g += `<line x1="${mL}" y1="${y(t)}" x2="${W - mR}" y2="${y(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${y(t) + 4}" text-anchor="end">${t * 100}%</text>`;
  }
  const stepX = Math.max(1, Math.round(n / 6));
  for (let k = 0; k < n; k += stepX) g += `<text class="axis-text" x="${x(k)}" y="${H - 10}" text-anchor="middle">${dstrShort(at(k))}</text>`;
  g += `<line x1="${mL}" y1="${y(0)}" x2="${W - mR}" y2="${y(0)}" stroke="var(--baseline)" stroke-width="1"/>`;

  let d = "", pen = false, first = null, last = null;
  for (let k = 0; k < n; k++) {
    if (vals[k] === null) { pen = false; continue; }
    if (first === null) first = k;
    last = k;
    d += (pen ? "L" : "M") + x(k).toFixed(1) + " " + y(vals[k]).toFixed(1);
    pen = true;
  }
  let area = "";
  if (first !== null) {
    let a = "M" + x(first).toFixed(1) + " " + y(0);
    for (let k = first; k <= last; k++) if (vals[k] !== null) a += "L" + x(k).toFixed(1) + " " + y(vals[k]).toFixed(1);
    a += "L" + x(last).toFixed(1) + " " + y(0) + "Z";
    area = `<path d="${a}" fill="var(--series-area)"/>`;
  }

  $(container).innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">
    ${g}${area}<path d="${d}" fill="none" stroke="var(--series-1)" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
    <line class="drCrosshair" x1="0" y1="${mT}" x2="0" y2="${mT + ih}" stroke="var(--baseline)" stroke-width="1" visibility="hidden"/>
    <g class="drHoverDots"></g>
    <rect x="${mL}" y="${mT}" width="${iw}" height="${ih}" fill="transparent" class="drTrendHover"/>
  </svg>`;
  const svg = $(container).querySelector("svg");
  const hover = svg.querySelector(".drTrendHover");
  const cross = svg.querySelector(".drCrosshair");
  const dots = svg.querySelector(".drHoverDots");
  hover.addEventListener("mousemove", (e) => {
    const rect = svg.getBoundingClientRect();
    const scale = W / rect.width;
    const px = (e.clientX - rect.left) * scale;
    const k = Math.max(0, Math.min(n - 1, Math.round(((px - mL) / iw) * (n - 1))));
    cross.setAttribute("x1", x(k)); cross.setAttribute("x2", x(k)); cross.setAttribute("visibility", "visible");
    dots.innerHTML = vals[k] === null ? "" : `<circle cx="${x(k)}" cy="${y(vals[k])}" r="4.5" fill="var(--series-1)" stroke="var(--surface-1)" stroke-width="2"/>`;
    const denom = s.dayTrue[k] + s.dayFalse[k];
    showTip(`<div class="tt-title">${dstr(at(k))}</div>
      <div class="tt-row"><span>Comm rate</span><b>${vals[k] === null ? "—" : pct(vals[k])}</b></div>
      <div class="tt-row"><span>Meters with data</span><b>${fmt(denom)}</b></div>
      <div class="tt-row"><span>Communicated</span><b>${fmt(s.dayTrue[k])}</b></div>`, e.clientX, e.clientY);
  });
  hover.addEventListener("mouseleave", () => { cross.setAttribute("visibility", "hidden"); dots.innerHTML = ""; hideTip(); });
}

// ---------- date-range meter table ----------
function drTableRows() {
  let rows = DRSTATS.perMeter.filter((r) => r.uptime >= 0);
  if (drTableSearch) rows = rows.filter((r) => r.m.ip.toLowerCase().includes(drTableSearch));
  const k = drTableSort.key, dir = drTableSort.dir;
  const val = (r) => {
    if (k === "ip" || k === "tsp" || k === "region" || k === "circle" || k === "po") return r.m[k];
    if (k === "lastSeen") return r.lastSeen;
    return r[k];
  };
  rows.sort((a, b) => {
    const va = val(a), vb = val(b);
    if (typeof va === "string") return dir * va.localeCompare(vb);
    return dir * (va - vb);
  });
  return rows;
}

function drRenderMeterTable() {
  $("drMeterTableHead").innerHTML = TABLE_COLS.map((c) =>
    `<th class="${c.num ? "num" : ""} ${drTableSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${drTableSort.key === c.key ? (drTableSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  $("drMeterTableHead").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (drTableSort.key === key) drTableSort.dir *= -1;
      else drTableSort = { key, dir: 1 };
      drRenderMeterTable();
    });
  });
  const rows = drTableRows();
  const shown = rows.slice(0, drTableLimit);
  $("drMeterTableBody").innerHTML = shown.map((r) => `<tr>
    <td>${esc(r.m.ip)}</td><td>${esc(r.m.tsp)}</td><td>${esc(r.m.region)}</td>
    <td>${esc(r.m.circle)}</td><td>${esc(r.m.po)}</td>
    <td class="num ${r.uptime < 0.5 ? "pct-bad" : ""}">${pct(r.uptime)}</td>
    <td class="num">${fmt(r.t)}</td><td class="num">${fmt(r.f)}</td><td class="num">${fmt(r.nl)}</td>
    <td class="num">${r.silentDays < 0 ? "—" : fmt(r.silentDays)}</td>
    <td>${r.lastSeen >= 0 ? dstr(DATA.dates[r.lastSeen]) : "never"}</td>
  </tr>`).join("");
  $("drTableCount").textContent = `Showing ${fmt(shown.length)} of ${fmt(rows.length)} meters`;
  $("drLoadMore").style.visibility = shown.length < rows.length ? "visible" : "hidden";
}

$("drMeterSearch").addEventListener("input", (e) => { drTableSearch = e.target.value.trim().toLowerCase(); drTableLimit = 25; drRenderMeterTable(); });
$("drLoadMore").addEventListener("click", () => { drTableLimit += 50; drRenderMeterTable(); });
$("drExportCsv").addEventListener("click", () => downloadMeterCsv("date_range_meters.csv", drTableRows()));

// ==========================================================
// Search IP tab — full single-meter drilldown
// ==========================================================
const srInput = $("srIpInput"), srSuggest = $("srSuggestions");
let srActiveIndex = -1;

function ipSuggestions(query) {
  if (!DATA || !query) return [];
  const q = query.toLowerCase();
  const out = [];
  for (const m of DATA.meters) {
    if (m.ip.toLowerCase().includes(q)) { out.push(m); if (out.length >= 20) break; }
  }
  return out;
}
function renderSuggestions(list) {
  srActiveIndex = -1;
  if (list.length === 0) { srSuggest.classList.add("hidden"); srSuggest.innerHTML = ""; return; }
  srSuggest.innerHTML = list.map((m) =>
    `<div class="suggest-item" data-ip="${esc(m.ip)}">${esc(m.ip)}<div class="s-sub">${esc(m.tsp)} · ${esc(m.region)}</div></div>`
  ).join("");
  srSuggest.classList.remove("hidden");
  srSuggest.querySelectorAll(".suggest-item").forEach((el) => {
    el.addEventListener("click", () => selectSearchIp(el.dataset.ip));
  });
}
function updateActiveSuggestion(items) {
  items.forEach((el, i) => el.classList.toggle("active", i === srActiveIndex));
}
srInput.addEventListener("input", () => renderSuggestions(ipSuggestions(srInput.value.trim())));
srInput.addEventListener("keydown", (e) => {
  const items = srSuggest.querySelectorAll(".suggest-item");
  if (e.key === "ArrowDown") { e.preventDefault(); srActiveIndex = Math.min(items.length - 1, srActiveIndex + 1); updateActiveSuggestion(items); }
  else if (e.key === "ArrowUp") { e.preventDefault(); srActiveIndex = Math.max(0, srActiveIndex - 1); updateActiveSuggestion(items); }
  else if (e.key === "Enter") {
    e.preventDefault();
    if (srActiveIndex >= 0 && items[srActiveIndex]) selectSearchIp(items[srActiveIndex].dataset.ip);
    else {
      const q = srInput.value.trim().toLowerCase();
      const exact = DATA && DATA.meters.find((m) => m.ip.toLowerCase() === q);
      if (exact) selectSearchIp(exact.ip);
      else if (items.length > 0) selectSearchIp(items[0].dataset.ip);
      else { $("srNotFound").classList.remove("hidden"); $("srResults").classList.add("hidden"); }
    }
  } else if (e.key === "Escape") srSuggest.classList.add("hidden");
});
document.addEventListener("click", (e) => {
  if (!e.target.closest("#srIpInput") && !e.target.closest("#srSuggestions")) srSuggest.classList.add("hidden");
});

function selectSearchIp(ip) {
  const m = DATA.meters.find((mm) => mm.ip === ip);
  srSuggest.classList.add("hidden");
  if (!m) { $("srNotFound").classList.remove("hidden"); $("srResults").classList.add("hidden"); return; }
  searchedMeter = m;
  srInput.value = m.ip;
  $("srNotFound").classList.add("hidden");
  $("srResults").classList.remove("hidden");
  renderSearchTab();
}

["srStart", "srEnd"].forEach((id) => $(id).addEventListener("change", () => { if (searchedMeter) renderSearchTab(); }));

function renderSearchTab() {
  const win = computeWindowFromInputs("srStart", "srEnd");
  if (!win) {
    $("srRangeMeta").textContent = `Invalid date range — pick a start/end date between ${dstr(DATA.dates[0])} and ${dstr(DATA.dates[DATA.dates.length - 1])}, with start on or before end.`;
    return;
  }
  const [i0, i1] = win;
  const m = searchedMeter;

  const sStats = computeStats([m], i0, i1);
  const sWeek = computeWeekendStats([m], i0, i1);
  const r = sStats.perMeter[0], wr = sWeek.perMeter[0];
  const category = classifyMeter(wr);
  const rhythmCat = r.t === 0 || r.f === 0 ? "—" : classifyRhythm(r.t, r.f, computeRunStats(m.days, i0, i1));
  const mLoss = computeMeterLoss(r);

  $("srRangeMeta").textContent = `${dstr(DATA.dates[i0])} → ${dstr(DATA.dates[i1])} (${i1 - i0 + 1} days)`;
  $("srRangeBanner").innerHTML = `Showing data for ${dstr(DATA.dates[i0])} → ${dstr(DATA.dates[i1])} <span class="rb-note">(${fmt(i1 - i0 + 1)} days · defaults to the full file — set the dates above to match another tab before comparing)</span>`;

  // duplicate info
  const rawRecords = (DATA.rawByIp && DATA.rawByIp.get(m.ip)) || [m];
  const dupCount = rawRecords.length;
  $("srDupNote").innerHTML = dupCount > 1
    ? `This IP appeared <b>${fmt(dupCount)}</b> times in the uploaded file. All rows were <b>merged day-by-day</b> into a single record ` +
      `(for each day: TRUE if any row was TRUE, else FALSE if any was FALSE, else no data) — see the raw rows below and the merged result.`
    : `This IP appeared <b>once</b> in the uploaded file — no duplicates to merge.`;
  const rawRows = rawRecords.map((rec, i) => `<tr>
    <td>${i + 1}</td><td>${esc(rec.tsp)}</td><td>${esc(rec.region)}</td><td>${esc(rec.circle)}</td><td>${esc(rec.po)}</td>
    <td class="num">${fmt(rec.t)}</td><td class="num">${fmt(rec.f)}</td><td class="num">${fmt(rec.days.length - rec.t - rec.f)}</td>
  </tr>`).join("");
  const mergedRow = dupCount > 1 ? `<tr style="border-top:2px solid var(--baseline);font-weight:600;">
    <td>Merged</td><td>${esc(m.tsp)}</td><td>${esc(m.region)}</td><td>${esc(m.circle)}</td><td>${esc(m.po)}</td>
    <td class="num">${fmt(m.t)}</td><td class="num">${fmt(m.f)}</td><td class="num">${fmt(m.days.length - m.t - m.f)}</td>
  </tr>` : "";
  $("srDupTableBody").innerHTML = rawRows + mergedRow;

  $("srKpiRow").innerHTML = [
    tile("Communication rate", r.uptime < 0 ? "—" : pct(r.uptime), "TRUE ÷ (TRUE+FALSE) over the picked range", true),
    tile("TRUE days", fmt(r.t), "communicated"),
    tile("FALSE days", fmt(r.f), "did not communicate", r.f > 0 ? "bad" : "good"),
    tile("No-data days", fmt(r.nl), "NULL / not reported"),
    tile("Silent for", r.silentDays < 0 ? "—" : `${fmt(r.silentDays)} day(s)`, "since last TRUE in this range", r.silentDays > 7 ? "bad" : "good"),
    tile("Last communicated", r.lastSeen >= 0 ? dstr(DATA.dates[r.lastSeen]) : "never", "within the picked range"),
    tile("Weekday/weekend pattern", category, "classification over this range"),
    tile("Rhythm", rhythmCat, "on/off pattern shape over this range"),
    tile("Revenue loss", `${inr(mLoss.lossAmount)} (${mLoss.lossPct.toFixed(1)}%)`, `SIM @ ₹${SIM_COST} · ${(mLoss.fraction * 100).toFixed(1)}% realized vs. ${inr(SIM_COST)} at 100% comm`, mLoss.lossPct > 0 ? "bad" : "good"),
  ].join("");

  renderSearchTrend(m, i0, i1);
  renderSearchMsnBarcode(m, i0, i1);
  renderSearchBarcode(m, i0, i1);
}

// For each day, decide which physical meter (MSN) "owned" this IP that day, from the
// raw duplicate rows for the IP. This is a LIFECYCLE model, not a per-day vote: a meter
// swap moves ownership forward and never flips back to a retired meter.
//
// Each record's "activation" is the first day it actually communicated (first TRUE);
// if it never has a TRUE, the first day it has any data (first non-NULL). On any given
// day, the owner is the MOST RECENTLY ACTIVATED meter among those already activated by
// that day. So once a newer MSN comes online, ties (e.g. both records FALSE, or the new
// one goes NULL for a bit) stay with the newer meter instead of alternating back.
// Days before any meter has activated have no owner (rendered grey).
function msnOwnershipByDay(m, i0, i1) {
  const records = (DATA.rawByIp && DATA.rawByIp.get(m.ip)) || [m];
  // activation index over the FULL timeline (so a meter that came online before the
  // visible window is still recognised as active within it)
  const withAct = records.map((rec) => {
    let firstTrue = -1, firstData = -1;
    for (let j = 0; j < rec.days.length; j++) {
      const v = rec.days[j];
      if (v !== -1 && firstData === -1) firstData = j;
      if (v === 1) { firstTrue = j; break; }
    }
    const activation = firstTrue !== -1 ? firstTrue : firstData; // -1 = never has any data
    return { rec, activation };
  });

  const owner = new Array(i1 - i0 + 1).fill(null); // { prefix, msn, status } per day
  for (let j = i0; j <= i1; j++) {
    let best = null, bestAct = -1;
    for (const w of withAct) {
      if (w.activation === -1 || w.activation > j) continue; // not activated yet by day j
      if (w.activation >= bestAct) { bestAct = w.activation; best = w.rec; } // latest activation wins
    }
    if (best) owner[j - i0] = { prefix: best.msnPrefix, msn: best.msn, status: best.days[j] };
  }
  return owner;
}

// per-record communication rate over a window [i0, i1]
function computeRecordStats(record, i0, i1) {
  let t = 0, f = 0;
  for (let j = i0; j <= i1; j++) {
    const v = record.days[j];
    if (v === 1) t++;
    else if (v === 0) f++;
  }
  const uptime = t + f > 0 ? t / (t + f) : -1;
  return { t, f, uptime };
}

// single-row day-by-day barcode coloured by the active meter (MSN prefix) each day
function renderSearchMsnBarcode(m, i0, i1) {
  const container = "chartSrMsn";
  const nDays = i1 - i0 + 1;
  const rowH = 64, mT = 4, mB = 4, mR = 10, mL = 4;
  const W = Math.max(600, Math.min(1300, $(container).clientWidth || 900));
  const iw = W - mL - mR;
  const cellW = iw / nDays;
  const gap = Math.min(1.4, cellW * 0.18);
  const barW = Math.max(0.6, cellW - gap);
  const H = mT + rowH + mB;

  const owner = msnOwnershipByDay(m, i0, i1);
  const records = (DATA.rawByIp && DATA.rawByIp.get(m.ip)) || [m];
  const colorMap = buildMsnColorMap(records); // full MSN -> colour (distinct even for same prefix)

  // legend: one entry per distinct MSN actually shown, in the order it first appears
  const seen = [];
  for (const o of owner) { if (o && !seen.some((e) => e.msn === o.msn)) seen.push(o); }
  const hasNoData = owner.some((o) => o === null);

  // per-MSN comm % for this window, for display below the chart
  const msnStats = new Map(); // msn -> { uptime, t, f }
  for (const rec of records) {
    const stats = computeRecordStats(rec, i0, i1);
    if (stats.t + stats.f > 0) msnStats.set(rec.msn, stats);
  }

  $("srMsnLegend").innerHTML =
    seen.map((o) => {
      const stats = msnStats.get(o.msn);
      const commStr = stats ? ` · ${pct(stats.uptime)}` : "";
      return `<span class="key"><span class="swatch" style="background:${colorMap.get(o.msn)}"></span>${esc(o.prefix === "Unknown" ? "Unknown" : o.prefix)} · ${esc(o.msn)}${commStr}</span>`;
    }).join("") +
    (hasNoData ? `<span class="key"><span class="swatch" style="background:var(--grid)"></span>No meter installed yet</span>` : "");

  let bars = "";
  for (let k = 0; k < nDays; k++) {
    const o = owner[k];
    const x = mL + k * cellW;
    const fill = o ? colorMap.get(o.msn) : "var(--grid)";
    bars += `<rect x="${x.toFixed(2)}" y="${mT}" width="${barW.toFixed(2)}" height="${rowH}" fill="${fill}"/>`;
  }

  $(container).innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${bars}</svg>`;
  const svg = $(container).querySelector("svg");
  svg.addEventListener("mousemove", (e) => {
    const rect = svg.getBoundingClientRect();
    const scale = W / rect.width;
    const px = (e.clientX - rect.left) * scale;
    const dk = Math.floor((px - mL) / cellW);
    if (dk < 0 || dk >= nDays) return hideTip();
    const o = owner[dk];
    const statusLabel = o && (o.status === 1 ? "Communicated (TRUE)" : o.status === 0 ? "Installed, did not communicate (FALSE)" : "Installed, no reading (NULL)");
    const body = o
      ? `<div class="tt-row"><span>Meter</span><b>${esc(o.prefix)}</b></div>` +
        `<div class="tt-row"><span>MSN</span><b>${esc(o.msn)}</b></div>` +
        `<div class="tt-row"><b>${statusLabel}</b></div>`
      : `<div class="tt-row"><b>No meter installed yet</b></div>`;
    showTip(`<div class="tt-title">${dstr(DATA.dates[i0 + dk])}</div>${body}`, e.clientX, e.clientY);
  });
  svg.addEventListener("mouseleave", hideTip);
}

// 7-day rolling communication rate for a single meter
function renderSearchTrend(m, i0, i1) {
  const nDays = i1 - i0 + 1, WIN = 7;
  const d = m.days;
  const vals = [];
  for (let k = 0; k < nDays; k++) {
    let t = 0, f = 0;
    for (let x = Math.max(0, k - WIN + 1); x <= k; x++) {
      const v = d[i0 + x];
      if (v === 1) t++; else if (v === 0) f++;
    }
    vals.push(t + f > 0 ? t / (t + f) : null);
  }

  const container = "chartSrTrend";
  const W = Math.max(680, Math.min(1300, $(container).clientWidth || 900));
  const H = 260, mL = 46, mR = 16, mT = 12, mB = 34;
  const iw = W - mL - mR, ih = H - mT - mB;
  const x = (k) => mL + (nDays === 1 ? iw / 2 : (k / (nDays - 1)) * iw);
  const y = (v) => mT + (1 - v) * ih;

  let g = "";
  for (const t of [0, 0.25, 0.5, 0.75, 1]) {
    g += `<line x1="${mL}" y1="${y(t)}" x2="${W - mR}" y2="${y(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${y(t) + 4}" text-anchor="end">${t * 100}%</text>`;
  }
  const stepX = Math.max(1, Math.round(nDays / 6));
  for (let k = 0; k < nDays; k += stepX) g += `<text class="axis-text" x="${x(k)}" y="${H - 10}" text-anchor="middle">${dstrShort(DATA.dates[i0 + k])}</text>`;
  g += `<line x1="${mL}" y1="${y(0)}" x2="${W - mR}" y2="${y(0)}" stroke="var(--baseline)" stroke-width="1"/>`;

  let dPath = "", pen = false;
  for (let k = 0; k < nDays; k++) {
    if (vals[k] === null) { pen = false; continue; }
    dPath += (pen ? "L" : "M") + x(k).toFixed(1) + " " + y(vals[k]).toFixed(1);
    pen = true;
  }

  $(container).innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">
    ${g}<path d="${dPath}" fill="none" stroke="var(--series-1)" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>
    <line class="srCrosshair" x1="0" y1="${mT}" x2="0" y2="${mT + ih}" stroke="var(--baseline)" stroke-width="1" visibility="hidden"/>
    <g class="srHoverDots"></g>
    <rect x="${mL}" y="${mT}" width="${iw}" height="${ih}" fill="transparent" class="srTrendHover"/>
  </svg>`;
  const svg = $(container).querySelector("svg");
  const hover = svg.querySelector(".srTrendHover");
  const cross = svg.querySelector(".srCrosshair");
  const dots = svg.querySelector(".srHoverDots");
  hover.addEventListener("mousemove", (e) => {
    const rect = svg.getBoundingClientRect();
    const scale = W / rect.width;
    const px = (e.clientX - rect.left) * scale;
    const k = Math.max(0, Math.min(nDays - 1, Math.round(((px - mL) / iw) * (nDays - 1))));
    cross.setAttribute("x1", x(k)); cross.setAttribute("x2", x(k)); cross.setAttribute("visibility", "visible");
    dots.innerHTML = vals[k] === null ? "" : `<circle cx="${x(k)}" cy="${y(vals[k])}" r="4.5" fill="var(--series-1)" stroke="var(--surface-1)" stroke-width="2"/>`;
    showTip(`<div class="tt-title">${dstr(DATA.dates[i0 + k])}</div><div class="tt-row"><span>7-day rate</span><b>${vals[k] === null ? "—" : pct(vals[k])}</b></div>`, e.clientX, e.clientY);
  });
  hover.addEventListener("mouseleave", () => { cross.setAttribute("visibility", "hidden"); dots.innerHTML = ""; hideTip(); });
}

// single-row day-by-day barcode for one meter
function renderSearchBarcode(m, i0, i1) {
  const container = "chartSrBarcode";
  const nDays = i1 - i0 + 1;
  const rowH = 64, mT = 4, mB = 4, mR = 10, mL = 4;
  const W = Math.max(600, Math.min(1300, $(container).clientWidth || 900));
  const iw = W - mL - mR;
  const cellW = iw / nDays;
  const gap = Math.min(1.4, cellW * 0.18); // thin gutter between bars — reads as a real barcode, not a solid strip
  const barW = Math.max(0.6, cellW - gap);
  const H = mT + rowH + mB;

  const COLOR = { 1: "var(--good-text)", 0: "var(--critical)", "-1": "var(--grid)" };
  const d = m.days;
  let bars = "";
  for (let j = i0; j <= i1; j++) {
    const v = d[j];
    const x = mL + (j - i0) * cellW;
    bars += `<rect x="${x.toFixed(2)}" y="${mT}" width="${barW.toFixed(2)}" height="${rowH}" fill="${COLOR[v]}"/>`;
  }

  $(container).innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${bars}</svg>`;
  const svg = $(container).querySelector("svg");
  svg.addEventListener("mousemove", (e) => {
    const rect = svg.getBoundingClientRect();
    const scale = W / rect.width;
    const px = (e.clientX - rect.left) * scale;
    const dj = Math.floor((px - mL) / cellW);
    if (dj < 0 || dj >= nDays) return hideTip();
    const day = i0 + dj, v = d[day];
    const label = v === 1 ? "Communicated (TRUE)" : v === 0 ? "Not communicated (FALSE)" : "No data (NULL)";
    showTip(`<div class="tt-title">${dstr(DATA.dates[day])}</div><div class="tt-row"><b>${label}</b></div>`, e.clientX, e.clientY);
  });
  svg.addEventListener("mouseleave", hideTip);
}

$("srExportCsv").addEventListener("click", () => {
  if (!searchedMeter) return;
  const win = computeWindowFromInputs("srStart", "srEnd");
  if (!win) return;
  const [i0, i1] = win;
  const m = searchedMeter;
  const rows = [];
  for (let j = i0; j <= i1; j++) {
    const v = m.days[j];
    rows.push([isoDate(DATA.dates[j]), v === 1 ? "TRUE" : v === 0 ? "FALSE" : "NULL"]);
  }
  downloadCsv(`ip_${slug(m.ip)}_daily.csv`, ["Date", "Status"], rows);
});

// ==========================================================
// Monthly Analytics tab — pick any set of calendar months (need not be contiguous)
// ==========================================================

// same shape as computeStats()'s return value, but driven by an arbitrary list of day
// indices instead of a contiguous [i0,i1] range — lets non-adjacent months (e.g. Jul + Sep) work
function computeStatsForIndices(meters, indices) {
  const nDays = indices.length;
  const dayTrue = new Float64Array(nDays), dayFalse = new Float64Array(nDays), dayNull = new Float64Array(nDays);
  const perMeter = [];
  const byGroup = { tsp: new Map(), region: new Map(), circle: new Map(), po: new Map() };
  let totT = 0, totF = 0, totN = 0, never = 0, silent7 = 0, silent30 = 0, full = 0;
  const buckets = [0, 0, 0, 0, 0, 0, 0, 0];

  for (const m of meters) {
    let t = 0, f = 0, nl = 0, lastSeenPos = -1; // position within `indices`, not an absolute day index
    const d = m.days;
    for (let k = 0; k < nDays; k++) {
      const v = d[indices[k]];
      if (v === 1) { t++; dayTrue[k]++; lastSeenPos = k; }
      else if (v === 0) { f++; dayFalse[k]++; }
      else { nl++; dayNull[k]++; }
    }
    totT += t; totF += f; totN += nl;
    const dataDays = t + f;
    const uptime = dataDays > 0 ? t / dataDays : -1;
    if (dataDays > 0 && t === 0) never++;
    if (uptime === 1) full++;
    let bucket = -1;
    const silentDays = dataDays > 0 ? (lastSeenPos === -1 ? nDays : (nDays - 1) - lastSeenPos) : -1;
    if (dataDays > 0) {
      if (silentDays >= 7 && nDays >= 7) silent7++;
      if (silentDays >= 30 && nDays >= 30) silent30++;
      if (uptime === 0) bucket = 0;
      else if (uptime === 1) bucket = 7;
      else if (uptime <= 0.25) bucket = 1;
      else if (uptime <= 0.5) bucket = 2;
      else if (uptime <= 0.75) bucket = 3;
      else if (uptime <= 0.85) bucket = 4;
      else if (uptime < 0.95) bucket = 5;
      else bucket = 6;
      buckets[bucket]++;
    }
    const lastSeen = lastSeenPos >= 0 ? indices[lastSeenPos] : -1; // absolute day index — DATA.dates[lastSeen] still works
    perMeter.push({ m, t, f, nl, uptime, lastSeen, bucket, silentDays });

    for (const g of ["tsp", "region", "circle", "po"]) {
      const key = m[g];
      let e = byGroup[g].get(key);
      if (!e) { e = { t: 0, f: 0, n: 0, meters: 0 }; byGroup[g].set(key, e); }
      e.t += t; e.f += f; e.n += nl; e.meters++;
    }
  }
  return { dayTrue, dayFalse, dayNull, perMeter, byGroup, totT, totF, totN, never, silent7, silent30, full, buckets, indices, nDays, meterCount: meters.length };
}

$("moSelectAll").addEventListener("click", () => {
  selectedMonths = new Set(DATA.monthList.map((mo) => mo.key));
  $("moMonthChips").querySelectorAll(".month-chip").forEach((btn) => btn.classList.add("active"));
  renderMonthly();
});
$("moClear").addEventListener("click", () => {
  selectedMonths = new Set();
  $("moMonthChips").querySelectorAll(".month-chip").forEach((btn) => btn.classList.remove("active"));
  renderMonthly();
});

function renderMonthly() {
  if (!DATA || selectedMonths.size === 0) {
    $("moNoneSelected").classList.remove("hidden");
    $("moContent").classList.add("hidden");
    return;
  }
  $("moNoneSelected").classList.add("hidden");
  $("moContent").classList.remove("hidden");

  const indices = [];
  const pickedLabels = [];
  for (const mo of DATA.monthList) {
    if (selectedMonths.has(mo.key)) { indices.push(...mo.indices); pickedLabels.push(mo.label); }
  }
  const meters = filteredMeters();
  MOSTATS = computeStatsForIndices(meters, indices);
  $("moRangeMeta").textContent = `${fmt(meters.length)} meters in view · ${pickedLabels.join(", ")} · ${fmt(indices.length)} days total`;
  $("moRangeBanner").innerHTML = `Showing data for ${esc(pickedLabels.join(", "))} <span class="rb-note">(${fmt(indices.length)} days · this selection is independent of the other tabs)</span>`;

  const dateAt = (k) => DATA.dates[MOSTATS.indices[k]];
  renderKpisInto(MOSTATS, "moKpiRow");
  renderTrendFor(MOSTATS, "chartMoTrend", "motrend", dateAt);
  renderGroupBar("tsp", "chartMoTsp", 99, MOSTATS, "motsp");
  renderHealth(MOSTATS, "chartMoHealth", "mohealth");
  renderGroupBar("region", "chartMoRegion", 99, MOSTATS, "moregion");
  renderGroupBar("circle", "chartMoCircle", 15, MOSTATS, "mocircle");
  renderGroupBar("po", "chartMoPo", 99, MOSTATS, "mopo");
  moTableLimit = 25;
  moRenderMeterTable();
  renderMoPatterns();
  renderMoRhythm();
}

// ---------- monthly-analytics meter table ----------
function moTableRows() {
  let rows = MOSTATS.perMeter.filter((r) => r.uptime >= 0);
  if (moTableSearch) rows = rows.filter((r) => r.m.ip.toLowerCase().includes(moTableSearch));
  const k = moTableSort.key, dir = moTableSort.dir;
  const val = (r) => {
    if (k === "ip" || k === "tsp" || k === "region" || k === "circle" || k === "po") return r.m[k];
    if (k === "lastSeen") return r.lastSeen;
    return r[k];
  };
  rows.sort((a, b) => {
    const va = val(a), vb = val(b);
    if (typeof va === "string") return dir * va.localeCompare(vb);
    return dir * (va - vb);
  });
  return rows;
}

function moRenderMeterTable() {
  $("moMeterTableHead").innerHTML = TABLE_COLS.map((c) =>
    `<th class="${c.num ? "num" : ""} ${moTableSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${moTableSort.key === c.key ? (moTableSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  $("moMeterTableHead").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (moTableSort.key === key) moTableSort.dir *= -1;
      else moTableSort = { key, dir: 1 };
      moRenderMeterTable();
    });
  });
  const rows = moTableRows();
  const shown = rows.slice(0, moTableLimit);
  $("moMeterTableBody").innerHTML = shown.map((r) => `<tr>
    <td>${esc(r.m.ip)}</td><td>${esc(r.m.tsp)}</td><td>${esc(r.m.region)}</td>
    <td>${esc(r.m.circle)}</td><td>${esc(r.m.po)}</td>
    <td class="num ${r.uptime < 0.5 ? "pct-bad" : ""}">${pct(r.uptime)}</td>
    <td class="num">${fmt(r.t)}</td><td class="num">${fmt(r.f)}</td><td class="num">${fmt(r.nl)}</td>
    <td class="num">${r.silentDays < 0 ? "—" : fmt(r.silentDays)}</td>
    <td>${r.lastSeen >= 0 ? dstr(DATA.dates[r.lastSeen]) : "never"}</td>
  </tr>`).join("");
  $("moTableCount").textContent = `Showing ${fmt(shown.length)} of ${fmt(rows.length)} meters`;
  $("moLoadMore").style.visibility = shown.length < rows.length ? "visible" : "hidden";
}

$("moMeterSearch").addEventListener("input", (e) => { moTableSearch = e.target.value.trim().toLowerCase(); moTableLimit = 25; moRenderMeterTable(); });
$("moLoadMore").addEventListener("click", () => { moTableLimit += 50; moRenderMeterTable(); });
$("moExportCsv").addEventListener("click", () => downloadMeterCsv("monthly_meters.csv", moTableRows()));

// ==========================================================
// Monthly Analytics — Patterns + Rhythm sections (mirrors the Patterns tab,
// scoped to the selected months' day indices instead of a contiguous window)
// ==========================================================
function renderMoPatterns() {
  const weekStats = computeWeekendStatsForIndices(filteredMeters(), MOSTATS.indices);
  MO_PSTATS = computePatternStats(weekStats);
  renderMoPatternKpis();
  renderMoPatternDist();
  renderMoPatternRegion();
  renderMoPatternMeterTable();
}

function renderMoPatternKpis() {
  const c = MO_PSTATS.counts;
  const classified = PATTERN_ORDER.reduce((a, k) => a + (k === "Insufficient data" ? 0 : c[k]), 0);
  const pctOf = (n) => (classified > 0 ? pct(n / classified) : "—");
  $("moPatKpiRow").innerHTML = [
    tile("Weekday-only meters", fmt(c["Weekday-only"]), `${pctOf(c["Weekday-only"])} of classified meters — ≥${PURE_SIDE_MIN_RATE * 100}% weekday comm rate, silent every weekend`, c["Weekday-only"] > 0 ? "bad" : "good", "Weekday-only"),
    tile("Weekend-only meters", fmt(c["Weekend-only"]), `${pctOf(c["Weekend-only"])} — ≥${PURE_SIDE_MIN_RATE * 100}% weekend comm rate, silent all week`, c["Weekend-only"] > 0 ? "bad" : "good", "Weekend-only"),
    tile("Always-on (7/7)", fmt(c["Always-on"]), `${pctOf(c["Always-on"])} — 100% uptime every weekday and weekend`, "good", "Always-on"),
    tile("Never active", fmt(c["Never active"]), `${pctOf(c["Never active"])} — zero TRUE days in the window`, c["Never active"] > 0 ? "bad" : "good", "Never active"),
    tile(`Weekend-drop (≥${GAP_THRESHOLD * 100}pp)`, fmt(c["Weekend-drop"]), `${pctOf(c["Weekend-drop"])} — meaningfully worse on weekends`, c["Weekend-drop"] > 0 ? "bad" : "good", "Weekend-drop"),
    tile(`Weekend-boost (≥${GAP_THRESHOLD * 100}pp)`, fmt(c["Weekend-boost"]), `${pctOf(c["Weekend-boost"])} — meaningfully better on weekends`, undefined, "Weekend-boost"),
  ].join("");
}
function selectMoPatternCategory(category) {
  moPatternCategory = category;
  $("moPatCategorySelect").value = moPatternCategory;
  moPatTableLimit = 25; moPatTableSearch = ""; $("moPatMeterSearch").value = "";
  renderMoPatternRegion();
  renderMoPatternMeterTable();
  moPatExportCsv();
}
wireKpiRow("moPatKpiRow", (key) => {
  if (!MO_PSTATS) return;
  selectMoPatternCategory(key);
});

function renderMoPatternDist() {
  const counts = PATTERN_ORDER.map((c) => MO_PSTATS.counts[c] || 0);
  if (tableModes["mopatdist"]) {
    $("chartMoPatDist").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Pattern</th><th class="num">Meters</th></tr></thead>
      <tbody>${PATTERN_ORDER.map((c, i) => `<tr><td>${c}</td><td class="num">${fmt(counts[i])}</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  const W = Math.max(560, Math.min(1000, $("chartMoPatDist").clientWidth || 800));
  const H = 260, mL = 52, mR = 14, mT = 12, mB = 46;
  const iw = W - mL - mR, ih = H - mT - mB;
  const max = Math.max(...counts, 1);
  const ticks = niceTicks(max);
  const yv = (v) => mT + ih - (v / ticks[ticks.length - 1]) * ih;
  const slotW = iw / counts.length, barW = Math.min(40, slotW * 0.55);

  let g = "";
  for (const t of ticks) {
    g += `<line x1="${mL}" y1="${yv(t)}" x2="${W - mR}" y2="${yv(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${yv(t) + 4}" text-anchor="end">${fmt(t)}</text>`;
  }
  let bars = "";
  counts.forEach((v, i) => {
    const cx = mL + slotW * i + slotW / 2;
    const yTop = yv(v), h = mT + ih - yTop;
    const r = Math.min(4, h);
    bars += `<path d="M${(cx - barW / 2).toFixed(1)} ${(mT + ih).toFixed(1)} v-${(h - r).toFixed(1)} a${r} ${r} 0 0 1 ${r} -${r} h${(barW - 2 * r).toFixed(1)} a${r} ${r} 0 0 1 ${r} ${r} v${(h - r).toFixed(1)} Z" fill="var(--series-1)"/>`;
    bars += `<text class="axis-text" x="${cx}" y="${H - 28}" text-anchor="middle">${PATTERN_SHORT[PATTERN_ORDER[i]]}</text>`;
    bars += `<rect x="${(cx - slotW / 2).toFixed(1)}" y="${mT}" width="${slotW.toFixed(1)}" height="${ih}" fill="transparent" class="mopatcol" data-i="${i}"/>`;
  });
  bars += `<line x1="${mL}" y1="${mT + ih}" x2="${W - mR}" y2="${mT + ih}" stroke="var(--baseline)" stroke-width="1"/>`;

  $("chartMoPatDist").innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${g}${bars}</svg>`;
  $("chartMoPatDist").querySelectorAll(".mopatcol").forEach((el) => {
    el.style.cursor = "pointer";
    const i = +el.dataset.i;
    el.addEventListener("mousemove", (e) => {
      showTip(`<div class="tt-title">${PATTERN_ORDER[i]}</div><div class="tt-row"><span>Meters</span><b>${fmt(counts[i])}</b></div><div class="tt-row"><i>Click bar to export CSV</i></div>`, e.clientX, e.clientY);
    });
    el.addEventListener("mouseleave", hideTip);
    el.addEventListener("click", () => selectMoPatternCategory(PATTERN_ORDER[i]));
  });
}

function renderMoPatternRegion() {
  const map = MO_PSTATS.byCategoryRegion.get(moPatternCategory) || new Map();
  const items = [...map.entries()].map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count).slice(0, 10);
  $("cardMoPatRegionSub").textContent = `${fmt(MO_PSTATS.counts[moPatternCategory] || 0)} meters classified as "${moPatternCategory}" — top regions`;
  if (tableModes["mopatregion"]) {
    $("chartMoPatRegion").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Region</th><th class="num">Meters</th></tr></thead>
      <tbody>${items.map((it) => `<tr><td>${esc(it.label)}</td><td class="num">${fmt(it.count)}</td></tr>`).join("")}</tbody></table></div>`;
  } else if (items.length === 0) {
    $("chartMoPatRegion").innerHTML = `<p class="card-sub">No meters in this pattern for the current filters.</p>`;
  } else {
    countBars("chartMoPatRegion", items, (it) => {
      const rows = MO_PSTATS.perMeter.filter((r) => r.category === moPatternCategory && r.m.region === it.label);
      downloadCsv(`monthly_pattern_${slug(moPatternCategory)}_${slug(it.label)}.csv`, PAT_TABLE_COLS.map((c) => c.label).concat("Pattern"), patRowsToCells(rows));
    });
  }
}

$("moPatCategorySelect").addEventListener("change", () => {
  moPatternCategory = $("moPatCategorySelect").value;
  moPatTableLimit = 25; moPatTableSearch = ""; $("moPatMeterSearch").value = "";
  renderMoPatternRegion();
  renderMoPatternMeterTable();
});

$("moPatMeterSearch").addEventListener("input", (e) => { moPatTableSearch = e.target.value.trim().toLowerCase(); moPatTableLimit = 25; renderMoPatternMeterTable(); });
$("moPatLoadMore").addEventListener("click", () => { moPatTableLimit += 50; renderMoPatternMeterTable(); });
$("moPatExportCsv").addEventListener("click", moPatExportCsv);

function moPatTableRows() {
  let rows = MO_PSTATS.perMeter.filter((r) => r.category === moPatternCategory);
  if (moPatTableSearch) rows = rows.filter((r) => r.m.ip.toLowerCase().includes(moPatTableSearch));
  const k = moPatTableSort.key, dir = moPatTableSort.dir;
  const val = (r) => {
    if (k === "ip" || k === "tsp" || k === "region" || k === "circle" || k === "po") return r.m[k];
    return r[k] === null ? -Infinity : r[k];
  };
  rows.sort((a, b) => {
    const va = val(a), vb = val(b);
    if (typeof va === "string") return dir * va.localeCompare(vb);
    return dir * (va - vb);
  });
  return rows;
}

function renderMoPatternMeterTable() {
  $("moPatMeterTableHead").innerHTML = PAT_TABLE_COLS.map((c) =>
    `<th class="${c.num ? "num" : ""} ${moPatTableSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${moPatTableSort.key === c.key ? (moPatTableSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  $("moPatMeterTableHead").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (moPatTableSort.key === key) moPatTableSort.dir *= -1;
      else moPatTableSort = { key, dir: -1 };
      renderMoPatternMeterTable();
    });
  });
  const rows = moPatTableRows();
  const shown = rows.slice(0, moPatTableLimit);
  $("moPatMeterTableBody").innerHTML = shown.map((r) => `<tr>
    <td>${esc(r.m.ip)}</td><td>${esc(r.m.tsp)}</td><td>${esc(r.m.region)}</td>
    <td>${esc(r.m.circle)}</td><td>${esc(r.m.po)}</td>
    <td class="num">${r.wdRate === null ? "—" : pct(r.wdRate)}</td>
    <td class="num">${r.weRate === null ? "—" : pct(r.weRate)}</td>
    <td class="num">${r.gap === null ? "—" : (r.gap * 100).toFixed(1)}</td>
  </tr>`).join("");
  $("moPatTableCount").textContent = `Showing ${fmt(shown.length)} of ${fmt(rows.length)} meters in "${moPatternCategory}"`;
  $("moPatLoadMore").style.visibility = shown.length < rows.length ? "visible" : "hidden";
}

function moPatExportCsv() {
  downloadCsv(`monthly_pattern_${slug(moPatternCategory)}.csv`, PAT_TABLE_COLS.map((c) => c.label).concat("Pattern"), patRowsToCells(moPatTableRows()));
}

// ---------- Monthly rhythm ----------
function renderMoRhythm() {
  MO_RSTATS = computeRhythmStats(MOSTATS, (days) => computeRunStatsForIndices(days, MOSTATS.indices));
  renderMoRhythmKpis();
  renderMoRhythmDist();
  renderMoRhythmRegion();
  renderMoRhythmMeterTable();
}

function renderMoRhythmKpis() {
  const c = MO_RSTATS.counts, total = MO_RSTATS.totalMixed;
  const pctOf = (n) => (total > 0 ? pct(n / total) : "—");
  $("moRhKpiRow").innerHTML = [
    tile("Alternating", fmt(c["Alternating"]), `${pctOf(c["Alternating"])} of mixed meters — flips on/off almost daily`, undefined, "Alternating"),
    tile("Long burst", fmt(c["Long burst"]), `${pctOf(c["Long burst"])} — communicates in extended multi-day blocks`, undefined, "Long burst"),
    tile("Sporadic", fmt(c["Sporadic"]), `${pctOf(c["Sporadic"])} — mostly silent, rare isolated TRUE blips`, c["Sporadic"] > 0 ? "bad" : "good", "Sporadic"),
    tile("Stable (occasional drop)", fmt(c["Stable (occasional drop)"]), `${pctOf(c["Stable (occasional drop)"])} — mostly on, rare isolated outages`, "good", "Stable (occasional drop)"),
    tile("Intermittent null", fmt(c["Intermittent null"]), `${pctOf(c["Intermittent null"])} — comm/non-comm days with a ${INTERMITTENT_NULL_MIN_GAP}+ day NULL gap in between, then data resumes`, c["Intermittent null"] > 0 ? "bad" : "good", "Intermittent null"),
    tile("Irregular", fmt(c["Irregular"]), `${pctOf(c["Irregular"])} — no clean rhythm detected`, undefined, "Irregular"),
  ].join("");
}
function selectMoRhythmCategory(category) {
  moRhythmCategory = category;
  $("moRhCategorySelect").value = moRhythmCategory;
  moRhTableLimit = 25; moRhTableSearch = ""; $("moRhMeterSearch").value = "";
  renderMoRhythmRegion();
  renderMoRhythmMeterTable();
  moRhExportCsv();
}
wireKpiRow("moRhKpiRow", (key) => {
  if (!MO_RSTATS) return;
  selectMoRhythmCategory(key);
});

function renderMoRhythmDist() {
  const counts = RHYTHM_ORDER.map((c) => MO_RSTATS.counts[c] || 0);
  if (tableModes["morhdist"]) {
    $("chartMoRhDist").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Rhythm</th><th class="num">Meters</th></tr></thead>
      <tbody>${RHYTHM_ORDER.map((c, i) => `<tr><td>${c}</td><td class="num">${fmt(counts[i])}</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  const W = Math.max(500, Math.min(900, $("chartMoRhDist").clientWidth || 700));
  const H = 260, mL = 52, mR = 14, mT = 12, mB = 40;
  const iw = W - mL - mR, ih = H - mT - mB;
  const max = Math.max(...counts, 1);
  const ticks = niceTicks(max);
  const yv = (v) => mT + ih - (v / ticks[ticks.length - 1]) * ih;
  const slotW = iw / counts.length, barW = Math.min(48, slotW * 0.55);

  let g = "";
  for (const t of ticks) {
    g += `<line x1="${mL}" y1="${yv(t)}" x2="${W - mR}" y2="${yv(t)}" stroke="var(--grid)" stroke-width="1"/>`;
    g += `<text class="axis-text" x="${mL - 8}" y="${yv(t) + 4}" text-anchor="end">${fmt(t)}</text>`;
  }
  let bars = "";
  counts.forEach((v, i) => {
    const cx = mL + slotW * i + slotW / 2;
    const yTop = yv(v), h = mT + ih - yTop;
    const r = Math.min(4, h);
    bars += `<path d="M${(cx - barW / 2).toFixed(1)} ${(mT + ih).toFixed(1)} v-${(h - r).toFixed(1)} a${r} ${r} 0 0 1 ${r} -${r} h${(barW - 2 * r).toFixed(1)} a${r} ${r} 0 0 1 ${r} ${r} v${(h - r).toFixed(1)} Z" fill="var(--series-1)"/>`;
    bars += `<text class="axis-text" x="${cx}" y="${H - 20}" text-anchor="middle">${RHYTHM_SHORT[RHYTHM_ORDER[i]]}</text>`;
    bars += `<rect x="${(cx - slotW / 2).toFixed(1)}" y="${mT}" width="${slotW.toFixed(1)}" height="${ih}" fill="transparent" class="morhcol" data-i="${i}"/>`;
  });
  bars += `<line x1="${mL}" y1="${mT + ih}" x2="${W - mR}" y2="${mT + ih}" stroke="var(--baseline)" stroke-width="1"/>`;

  $("chartMoRhDist").innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:${W}px">${g}${bars}</svg>`;
  $("chartMoRhDist").querySelectorAll(".morhcol").forEach((el) => {
    el.style.cursor = "pointer";
    const i = +el.dataset.i;
    el.addEventListener("mousemove", (e) => {
      showTip(`<div class="tt-title">${RHYTHM_ORDER[i]}</div><div class="tt-row"><span>Meters</span><b>${fmt(counts[i])}</b></div><div class="tt-row"><i>Click bar to export CSV</i></div>`, e.clientX, e.clientY);
    });
    el.addEventListener("mouseleave", hideTip);
    el.addEventListener("click", () => selectMoRhythmCategory(RHYTHM_ORDER[i]));
  });
}

function renderMoRhythmRegion() {
  const map = MO_RSTATS.byCategoryRegion.get(moRhythmCategory) || new Map();
  const items = [...map.entries()].map(([label, count]) => ({ label, count })).sort((a, b) => b.count - a.count).slice(0, 10);
  $("cardMoRhRegionSub").textContent = `${fmt(MO_RSTATS.counts[moRhythmCategory] || 0)} meters classified as "${moRhythmCategory}" — top regions`;
  if (tableModes["morhregion"]) {
    $("chartMoRhRegion").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Region</th><th class="num">Meters</th></tr></thead>
      <tbody>${items.map((it) => `<tr><td>${esc(it.label)}</td><td class="num">${fmt(it.count)}</td></tr>`).join("")}</tbody></table></div>`;
  } else if (items.length === 0) {
    $("chartMoRhRegion").innerHTML = `<p class="card-sub">No meters in this rhythm for the current filters.</p>`;
  } else {
    countBars("chartMoRhRegion", items, (it) => {
      const rows = MO_RSTATS.perMeter.filter((r) => r.category === moRhythmCategory && r.m.region === it.label);
      downloadCsv(`monthly_rhythm_${slug(moRhythmCategory)}_${slug(it.label)}.csv`, RH_TABLE_COLS.map((c) => c.label).concat("Rhythm"), rhRowsToCells(rows));
    });
  }
}

$("moRhCategorySelect").addEventListener("change", () => {
  moRhythmCategory = $("moRhCategorySelect").value;
  moRhTableLimit = 25; moRhTableSearch = ""; $("moRhMeterSearch").value = "";
  renderMoRhythmRegion();
  renderMoRhythmMeterTable();
});

$("moRhMeterSearch").addEventListener("input", (e) => { moRhTableSearch = e.target.value.trim().toLowerCase(); moRhTableLimit = 25; renderMoRhythmMeterTable(); });
$("moRhLoadMore").addEventListener("click", () => { moRhTableLimit += 50; renderMoRhythmMeterTable(); });
$("moRhExportCsv").addEventListener("click", moRhExportCsv);

function moRhTableRows() {
  let rows = MO_RSTATS.perMeter.filter((r) => r.category === moRhythmCategory);
  if (moRhTableSearch) rows = rows.filter((r) => r.m.ip.toLowerCase().includes(moRhTableSearch));
  const k = moRhTableSort.key, dir = moRhTableSort.dir;
  const val = (r) => {
    if (k === "ip" || k === "tsp" || k === "region" || k === "circle" || k === "po") return r.m[k];
    if (k === "uptime") return r.uptime;
    return r.rs[k];
  };
  rows.sort((a, b) => {
    const va = val(a), vb = val(b);
    if (typeof va === "string") return dir * va.localeCompare(vb);
    return dir * (va - vb);
  });
  return rows;
}

function renderMoRhythmMeterTable() {
  $("moRhMeterTableHead").innerHTML = RH_TABLE_COLS.map((c) =>
    `<th class="${c.num ? "num" : ""} ${moRhTableSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${moRhTableSort.key === c.key ? (moRhTableSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  $("moRhMeterTableHead").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (moRhTableSort.key === key) moRhTableSort.dir *= -1;
      else moRhTableSort = { key, dir: -1 };
      renderMoRhythmMeterTable();
    });
  });
  const rows = moRhTableRows();
  const shown = rows.slice(0, moRhTableLimit);
  $("moRhMeterTableBody").innerHTML = shown.map((r) => `<tr>
    <td>${esc(r.m.ip)}</td><td>${esc(r.m.tsp)}</td><td>${esc(r.m.region)}</td>
    <td>${esc(r.m.circle)}</td><td>${esc(r.m.po)}</td>
    <td class="num">${pct(r.uptime)}</td>
    <td class="num">${r.rs.avgRunLen.toFixed(1)}</td>
    <td class="num">${fmt(r.rs.transitions)}</td>
    <td class="num">${fmt(r.rs.maxTrueRun)}</td>
    <td class="num">${fmt(r.rs.maxFalseRun)}</td>
  </tr>`).join("");
  $("moRhTableCount").textContent = `Showing ${fmt(shown.length)} of ${fmt(rows.length)} meters in "${moRhythmCategory}"`;
  $("moRhLoadMore").style.visibility = shown.length < rows.length ? "visible" : "hidden";
}

function moRhExportCsv() {
  downloadCsv(`monthly_rhythm_${slug(moRhythmCategory)}.csv`, RH_TABLE_COLS.map((c) => c.label).concat("Rhythm"), rhRowsToCells(moRhTableRows()));
}

// ==========================================================
// "View all" modal — full sortable list for any group-by chart (circle, region, etc.)
// ==========================================================
const modalOverlay = $("modalOverlay");
let modalSort = { key: "value", dir: 1 };
let modalItemsGetter = null;
let modalGroupKey = null;
let modalStatsSrc = null;

const MODAL_COLS = [
  { key: "label", label: "Circle", num: false },
  { key: "value", label: "Comm rate", num: true },
  { key: "meters", label: "Meters", num: true },
  { key: "t", label: "TRUE", num: true },
  { key: "f", label: "FALSE", num: true },
  { key: "n", label: "No data", num: true },
];

function openGroupModal(title, g, statsSrc) {
  modalGroupKey = g;
  modalStatsSrc = statsSrc;
  modalSort = { key: "value", dir: 1 }; // default: worst-first, same as the chart
  $("modalTitle").textContent = title;
  modalOverlay.classList.remove("hidden");
  renderModalTable();
}
function closeGroupModal() {
  modalOverlay.classList.add("hidden");
}
$("modalClose").addEventListener("click", closeGroupModal);
modalOverlay.addEventListener("click", (e) => { if (e.target === modalOverlay) closeGroupModal(); });
document.addEventListener("keydown", (e) => { if (e.key === "Escape" && !modalOverlay.classList.contains("hidden")) closeGroupModal(); });

function renderModalTable() {
  const items = groupItems(modalGroupKey, Infinity, modalStatsSrc);
  const k = modalSort.key, dir = modalSort.dir;
  items.sort((a, b) => {
    const va = a[k], vb = b[k];
    if (typeof va === "string") return dir * va.localeCompare(vb);
    return dir * (va - vb);
  });
  const head = MODAL_COLS.map((c) =>
    `<th class="${c.num ? "num" : ""} ${modalSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${modalSort.key === c.key ? (modalSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  const rows = items.map((it) => `<tr>
    <td>${esc(it.label)}</td>
    <td class="num">${pct(it.value)}</td>
    <td class="num">${fmt(it.meters)}</td>
    <td class="num">${fmt(it.t)}</td>
    <td class="num">${fmt(it.f)}</td>
    <td class="num">${fmt(it.n)}</td>
  </tr>`).join("");
  $("modalBody").innerHTML = `<div class="table-scroll"><table><thead><tr>${head}</tr></thead><tbody>${rows}</tbody></table></div>
    <div class="table-foot"><span>${fmt(items.length)} total</span><button id="modalExportCsv" class="ghost-btn">Export CSV</button></div>`;
  $("modalBody").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (modalSort.key === key) modalSort.dir *= -1;
      else modalSort = { key, dir: key === "label" ? 1 : -1 };
      renderModalTable();
    });
  });
  $("modalExportCsv").addEventListener("click", () => {
    downloadCsv(`${modalGroupKey}_all.csv`,
      MODAL_COLS.map((c) => c.label),
      items.map((it) => [it.label, (it.value * 100).toFixed(2), it.meters, it.t, it.f, it.n]));
  });
}

$("circleViewAllBtn").addEventListener("click", () => openGroupModal("All circles", "circle", STATS));
$("drCircleViewAllBtn").addEventListener("click", () => { if (DRSTATS) openGroupModal("All circles", "circle", DRSTATS); });
$("moCircleViewAllBtn").addEventListener("click", () => { if (MOSTATS) openGroupModal("All circles", "circle", MOSTATS); });

// ==========================================================
// Smart Meters tab
// ==========================================================
["fSmRegion", "fSmCircle", "fSmMake"].forEach((id) => {
  $(id).addEventListener("change", () => {
    SM_FILTER = { region: $("fSmRegion").value, circle: $("fSmCircle").value, make: $("fSmMake").value };
    smTableLimit = 25;
    renderSmartMeters();
  });
});
$("smInsightStatus").addEventListener("change", () => {
  smInsightStatus = $("smInsightStatus").value;
  renderSmInsightRegion();
  renderSmInsightCircle();
});

function smFilteredRows() {
  if (!SMDATA) return [];
  return SMDATA.rows.filter((r) =>
    (!SM_FILTER.region || r.region === SM_FILTER.region) &&
    (!SM_FILTER.circle || r.circle === SM_FILTER.circle) &&
    (!SM_FILTER.make || r.make === SM_FILTER.make)
  );
}

// bucket rows by a key, tracking per-bucket status counts (and, for circles, the owning region)
function smBucket(rows, keyFn, withRegion) {
  const map = new Map();
  for (const r of rows) {
    const key = keyFn(r);
    let e = map.get(key);
    if (!e) { e = { total: 0, counts: new Map(), region: withRegion ? r.region : undefined }; map.set(key, e); }
    e.total++;
    e.counts.set(r.status, (e.counts.get(r.status) || 0) + 1);
  }
  return map;
}
function smCommRate(e) { return e.total > 0 ? (e.total - (e.counts.get("SIM Installation pending") || 0) - (e.counts.get("DC") || 0)) / e.total : 0; }

let SMSTATS = null;
function computeSmStats() {
  const rows = smFilteredRows();
  const statusCounts = new Map();
  for (const r of rows) statusCounts.set(r.status, (statusCounts.get(r.status) || 0) + 1);
  return {
    rows,
    total: rows.length,
    statusCounts,
    byRegion: smBucket(rows, (r) => r.region, false),
    byCircle: smBucket(rows, (r) => r.circle, true),
    byMake: smBucket(rows, (r) => r.make, false),
  };
}

function renderSmartMeters() {
  if (!SMDATA) return;
  SMSTATS = computeSmStats();
  $("smFilterMeta").textContent = `${fmt(SMSTATS.total)} smart meters in view`;
  renderSmKpis();
  renderSmRegionChart();
  renderSmCircleChart();
  renderSmStatusChart();
  renderSmMakeChart();
  renderSmInsightRegion();
  renderSmInsightCircle();
  renderSmMeterTable();
}

function renderSmKpis() {
  const s = SMSTATS;
  const dcN = s.statusCounts.get("DC") || 0;
  const simN = s.statusCounts.get("SIM Installation pending") || 0;
  const nonCommN = s.statusCounts.get(">1 Month Non Comm") || 0;
  const commN = s.statusCounts.get("Communicating") || 0;
  const simInstallPct = s.total > 0 ? (s.total - simN - dcN) / s.total : 0;
  $("smKpiRow").innerHTML =
    tile("Total smart meters", fmt(s.total), null, false) +
    tile("SIM Installation", pct(simInstallPct), `${fmt(s.total - simN - dcN)} meters`, simInstallPct >= 0.75 ? "good" : simInstallPct < 0.5 ? "bad" : null, "sm_sim") +
    tile("SIM pending", fmt(simN), s.total > 0 ? pct(simN / s.total) : "0%", simN > 0 ? "bad" : "good", "sm_simpending") +
    tile("DC (disconnected)", fmt(dcN), s.total > 0 ? pct(dcN / s.total) : "0%", dcN > 0 ? "bad" : "good", "sm_dc") +
    tile(">1 Month Non Comm", fmt(nonCommN), s.total > 0 ? pct(nonCommN / s.total) : "0%", nonCommN > 0 ? "bad" : "good", "sm_nc");
}
wireKpiRow("smKpiRow", (key) => {
  if (!SMSTATS) return;
  const statusOf = { sm_sim: "SIM Installation pending", sm_dc: "DC", sm_nc: ">1 Month Non Comm" };
  if (key === "sm_comm") return smExportRows(SMSTATS.rows.filter((r) => r.status === "Communicating"), "smartmeters_communicating.csv");
  const status = statusOf[key];
  if (status) smExportRows(SMSTATS.rows.filter((r) => r.status === status), `smartmeters_${slug(status)}.csv`);
});

function smMapToItems(map, limit) {
  const items = [];
  map.forEach((e, key) => { if (e.total > 0) items.push({ label: key, value: smCommRate(e), total: e.total, counts: e.counts }); });
  items.sort((a, b) => a.value - b.value); // worst first
  return items.slice(0, limit);
}
function smGroupTable(items) {
  return `<div class="table-scroll mini-table"><table>
    <thead><tr><th>Group</th><th class="num">Comm %</th><th class="num">Total</th><th class="num">Communicating</th><th class="num">DC</th><th class="num">SIM pending</th><th class="num">&gt;1M non-comm</th></tr></thead>
    <tbody>${items.map((it) => `<tr><td>${esc(it.label)}</td><td class="num">${pct(it.value)}</td><td class="num">${fmt(it.total)}</td>
      <td class="num">${fmt(it.counts.get("Communicating") || 0)}</td><td class="num">${fmt(it.counts.get("DC") || 0)}</td>
      <td class="num">${fmt(it.counts.get("SIM Installation pending") || 0)}</td><td class="num">${fmt(it.counts.get(">1 Month Non Comm") || 0)}</td></tr>`).join("")}</tbody>
  </table></div>`;
}
function smItemTip(it) {
  return `<div class="tt-title">${esc(it.label)}</div>
    <div class="tt-row"><span>Comm %</span><b>${pct(it.value)}</b></div>
    <div class="tt-row"><span>Total meters</span><b>${fmt(it.total)}</b></div>
    <div class="tt-row"><span>DC</span><b>${fmt(it.counts.get("DC") || 0)}</b></div>
    <div class="tt-row"><span>SIM pending</span><b>${fmt(it.counts.get("SIM Installation pending") || 0)}</b></div>
    <div class="tt-row"><span>&gt;1M non-comm</span><b>${fmt(it.counts.get(">1 Month Non Comm") || 0)}</b></div>`;
}

function renderSmRegionChart() {
  const items = smMapToItems(SMSTATS.byRegion, 99).map((it) => ({ ...it, sub: `${fmt(it.total)} meters`, tip: smItemTip(it) }));
  if (tableModes["smregion"]) { $("chartSmRegion").innerHTML = smGroupTable(items); return; }
  if (items.length === 0) { $("chartSmRegion").innerHTML = `<p class="card-sub">No data for this slice.</p>`; return; }
  hBars("chartSmRegion", items, { onClick: (it) => smExportRows(SMSTATS.rows.filter((r) => r.region === it.label), `smartmeters_region_${slug(it.label)}.csv`) });
}
function renderSmCircleChart() {
  const items = smMapToItems(SMSTATS.byCircle, 15).map((it) => ({ ...it, sub: `${fmt(it.total)} meters`, tip: smItemTip(it) }));
  if (tableModes["smcircle"]) { $("chartSmCircle").innerHTML = smGroupTable(items); return; }
  if (items.length === 0) { $("chartSmCircle").innerHTML = `<p class="card-sub">No data for this slice.</p>`; return; }
  hBars("chartSmCircle", items, { onClick: (it) => smExportRows(SMSTATS.rows.filter((r) => r.circle === it.label), `smartmeters_circle_${slug(it.label)}.csv`) });
}
function renderSmMakeChart() {
  const items = smMapToItems(SMSTATS.byMake, 99).map((it) => ({ ...it, sub: `${fmt(it.total)} meters`, tip: smItemTip(it) }));
  if (tableModes["smmake"]) { $("chartSmMake").innerHTML = smGroupTable(items); return; }
  if (items.length === 0) { $("chartSmMake").innerHTML = `<p class="card-sub">No data for this slice.</p>`; return; }
  hBars("chartSmMake", items, { onClick: (it) => smExportRows(SMSTATS.rows.filter((r) => r.make === it.label), `smartmeters_make_${slug(it.label)}.csv`) });
}

function renderSmStatusChart() {
  const s = SMSTATS;
  const items = (SMDATA.statuses || []).map((status) => ({ label: status, count: s.statusCounts.get(status) || 0 }));
  if (tableModes["smstatus"]) {
    $("chartSmStatus").innerHTML = `<div class="table-scroll mini-table"><table>
      <thead><tr><th>Status</th><th class="num">Meters</th><th class="num">Share</th></tr></thead>
      <tbody>${items.map((it) => `<tr><td>${esc(it.label)}</td><td class="num">${fmt(it.count)}</td><td class="num">${s.total > 0 ? pct(it.count / s.total) : "0%"}</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  countBars("chartSmStatus", items, (it) => smExportRows(s.rows.filter((r) => r.status === it.label), `smartmeters_status_${slug(it.label)}.csv`));
}

// ---------- insights: which region/circle carries the most of a given status ----------
function smInsightItems(map, limit) {
  const items = [];
  map.forEach((e, key) => items.push({ label: key, count: e.counts.get(smInsightStatus) || 0, total: e.total }));
  items.sort((a, b) => b.count - a.count);
  return items.filter((it) => it.count > 0).slice(0, limit);
}
function renderSmInsightRegion() {
  $("smInsightRegionSub").textContent = `Largest "${smInsightStatus}" count, by region`;
  const items = smInsightItems(SMSTATS.byRegion, 10);
  if (tableModes["sminsightregion"]) {
    $("chartSmInsightRegion").innerHTML = `<div class="table-scroll mini-table"><table><thead><tr><th>Region</th><th class="num">Count</th></tr></thead>
      <tbody>${items.map((it) => `<tr><td>${esc(it.label)}</td><td class="num">${fmt(it.count)}</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  if (items.length === 0) { $("chartSmInsightRegion").innerHTML = `<p class="card-sub">No meters with this status in view.</p>`; return; }
  countBars("chartSmInsightRegion", items, (it) => smExportRows(SMSTATS.rows.filter((r) => r.region === it.label && r.status === smInsightStatus), `smartmeters_${slug(smInsightStatus)}_region_${slug(it.label)}.csv`));
}
function renderSmInsightCircle() {
  $("smInsightCircleSub").textContent = `Largest "${smInsightStatus}" count, by circle`;
  const items = smInsightItems(SMSTATS.byCircle, 10);
  if (tableModes["sminsightcircle"]) {
    $("chartSmInsightCircle").innerHTML = `<div class="table-scroll mini-table"><table><thead><tr><th>Circle</th><th class="num">Count</th></tr></thead>
      <tbody>${items.map((it) => `<tr><td>${esc(it.label)}</td><td class="num">${fmt(it.count)}</td></tr>`).join("")}</tbody></table></div>`;
    return;
  }
  if (items.length === 0) { $("chartSmInsightCircle").innerHTML = `<p class="card-sub">No meters with this status in view.</p>`; return; }
  countBars("chartSmInsightCircle", items, (it) => smExportRows(SMSTATS.rows.filter((r) => r.circle === it.label && r.status === smInsightStatus), `smartmeters_${slug(smInsightStatus)}_circle_${slug(it.label)}.csv`));
}

// ---------- raw smart meter table ----------
const SM_TABLE_COLS = [
  { key: "kk", label: "MSN (kk)", num: false },
  { key: "mserial", label: "MSERIAL", num: false },
  { key: "circle", label: "Circle", num: false },
  { key: "region", label: "Region", num: false },
  { key: "status", label: "Status", num: false },
  { key: "make", label: "Make", num: false },
  { key: "cuscode", label: "CUSCODE", num: false },
  { key: "remarks", label: "Remarks", num: false },
];
$("smMeterSearch").addEventListener("input", (e) => { smTableSearch = e.target.value.trim().toLowerCase(); smTableLimit = 25; renderSmMeterTable(); });
$("smLoadMore").addEventListener("click", () => { smTableLimit += 50; renderSmMeterTable(); });
$("smExportCsv").addEventListener("click", () => smExportRows(smTableRows(), "smartmeters.csv"));

function smTableRows() {
  let rows = SMSTATS.rows;
  if (smTableSearch) {
    const q = smTableSearch;
    rows = rows.filter((r) => r.kk.toLowerCase().includes(q) || r.mserial.toLowerCase().includes(q) || r.circle.toLowerCase().includes(q) || r.region.toLowerCase().includes(q));
  }
  const k = smTableSort.key, dir = smTableSort.dir;
  rows = rows.slice().sort((a, b) => dir * String(a[k]).localeCompare(String(b[k])));
  return rows;
}
function renderSmMeterTable() {
  $("smMeterTableHead").innerHTML = SM_TABLE_COLS.map((c) =>
    `<th class="${c.num ? "num" : ""} ${smTableSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${smTableSort.key === c.key ? (smTableSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  $("smMeterTableHead").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (smTableSort.key === key) smTableSort.dir *= -1;
      else smTableSort = { key, dir: 1 };
      renderSmMeterTable();
    });
  });
  const rows = smTableRows();
  const shown = rows.slice(0, smTableLimit);
  $("smMeterTableBody").innerHTML = shown.map((r) => `<tr>
    <td>${esc(r.kk)}</td><td>${esc(r.mserial)}</td><td>${esc(r.circle)}</td><td>${esc(r.region)}</td>
    <td style="color:${smStatusColor(r.status)}">${esc(r.status)}</td><td>${esc(r.make)}</td><td>${esc(r.cuscode)}</td><td>${esc(r.remarks)}</td>
  </tr>`).join("");
  $("smTableCount").textContent = `Showing ${fmt(shown.length)} of ${fmt(rows.length)} smart meters`;
  $("smLoadMore").style.visibility = shown.length < rows.length ? "visible" : "hidden";
}
function smExportRows(rows, filename) {
  downloadCsv(filename, SM_TABLE_COLS.map((c) => c.label), rows.map((r) => SM_TABLE_COLS.map((c) => r[c.key])));
}

// ---------- "view all circles" modal for smart meters ----------
function renderSmCircleModalTable() {
  const items = smMapToItems(SMSTATS.byCircle, Infinity);
  const k = modalSort.key, dir = modalSort.dir;
  items.sort((a, b) => {
    if (k === "label") return dir * a.label.localeCompare(b.label);
    if (k === "value") return dir * (a.value - b.value);
    if (k === "total") return dir * (a.total - b.total);
    return dir * ((a.counts.get(k) || 0) - (b.counts.get(k) || 0));
  });
  const cols = [
    { key: "label", label: "Circle" }, { key: "value", label: "Comm %" }, { key: "total", label: "Total" },
    { key: "DC", label: "DC" }, { key: "SIM Installation pending", label: "SIM pending" }, { key: ">1 Month Non Comm", label: ">1M non-comm" },
  ];
  const head = cols.map((c) =>
    `<th class="${c.key === "label" ? "" : "num"} ${modalSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${c.label}${modalSort.key === c.key ? (modalSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  const rows = items.map((it) => `<tr>
    <td>${esc(it.label)}</td><td class="num">${pct(it.value)}</td><td class="num">${fmt(it.total)}</td>
    <td class="num">${fmt(it.counts.get("DC") || 0)}</td><td class="num">${fmt(it.counts.get("SIM Installation pending") || 0)}</td>
    <td class="num">${fmt(it.counts.get(">1 Month Non Comm") || 0)}</td>
  </tr>`).join("");
  $("modalBody").innerHTML = `<div class="table-scroll"><table><thead><tr>${head}</tr></thead><tbody>${rows}</tbody></table></div>
    <div class="table-foot"><span>${fmt(items.length)} total</span><button id="modalExportCsv" class="ghost-btn">Export CSV</button></div>`;
  $("modalBody").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (modalSort.key === key) modalSort.dir *= -1;
      else modalSort = { key, dir: key === "label" ? 1 : -1 };
      renderSmCircleModalTable();
    });
  });
  $("modalExportCsv").addEventListener("click", () => smExportRows(SMSTATS.rows.filter((r) => items.some((it) => it.label === r.circle)), "smartmeters_circles_all.csv"));
}
$("smCircleViewAllBtn").addEventListener("click", () => {
  if (!SMSTATS) return;
  modalSort = { key: "value", dir: 1 };
  $("modalTitle").textContent = "All circles — smart meters";
  modalOverlay.classList.remove("hidden");
  renderSmCircleModalTable();
});

// ---------- "view all circles" modal for the selected insight status ----------
function renderSmInsightCircleModalTable() {
  const items = [];
  SMSTATS.byCircle.forEach((e, circle) => {
    const count = e.counts.get(smInsightStatus) || 0;
    items.push({ label: circle, region: e.region, count, total: e.total, share: e.total > 0 ? count / e.total : 0 });
  });
  const k = modalSort.key, dir = modalSort.dir;
  items.sort((a, b) => {
    if (k === "label" || k === "region") return dir * a[k].localeCompare(b[k]);
    return dir * (a[k] - b[k]);
  });
  const cols = [
    { key: "label", label: "Circle" }, { key: "region", label: "Region" },
    { key: "count", label: `${smInsightStatus}` }, { key: "total", label: "Circle total" }, { key: "share", label: "Share" },
  ];
  const head = cols.map((c) =>
    `<th class="${c.key === "label" || c.key === "region" ? "" : "num"} ${modalSort.key === c.key ? "sorted" : ""}" data-key="${c.key}">${esc(c.label)}${modalSort.key === c.key ? (modalSort.dir === 1 ? " ↑" : " ↓") : ""}</th>`
  ).join("");
  const rows = items.map((it) => `<tr>
    <td>${esc(it.label)}</td><td>${esc(it.region)}</td>
    <td class="num">${fmt(it.count)}</td><td class="num">${fmt(it.total)}</td><td class="num">${pct(it.share)}</td>
  </tr>`).join("");
  $("modalBody").innerHTML = `<div class="table-scroll"><table><thead><tr>${head}</tr></thead><tbody>${rows}</tbody></table></div>
    <div class="table-foot"><span>${fmt(items.length)} total</span><button id="modalExportCsv" class="ghost-btn">Export CSV</button></div>`;
  $("modalBody").querySelectorAll("th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (modalSort.key === key) modalSort.dir *= -1;
      else modalSort = { key, dir: key === "label" || key === "region" ? 1 : -1 };
      renderSmInsightCircleModalTable();
    });
  });
  $("modalExportCsv").addEventListener("click", () => {
    downloadCsv(`smartmeters_${slug(smInsightStatus)}_all_circles.csv`,
      cols.map((c) => c.label),
      items.map((it) => [it.label, it.region, it.count, it.total, (it.share * 100).toFixed(2) + "%"]));
  });
}
$("smInsightCircleViewAllBtn").addEventListener("click", () => {
  if (!SMSTATS) return;
  modalSort = { key: "count", dir: -1 };
  $("modalTitle").textContent = `All circles — "${smInsightStatus}"`;
  modalOverlay.classList.remove("hidden");
  renderSmInsightCircleModalTable();
});
