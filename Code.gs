// ============================================================
// OIL ANALYSIS WEB APP — Code.gs
// Full app: Dashboard, List, Detail, Import, Add/Edit/Delete
// ============================================================

const SHEET_NAME = 'OilAnalysis';

const COLUMNS = [
  'Source','ID_Vendor','Rating','Health','Interp_Text',
  'Report_Date','Sampled_Date','Month','Section','Model',
  'Asset_ID','Component','Meter_On_Asset','Meter_on_Fluid',
  'PM_Type','Fluid_Changed','Fluid_Brand','Fluid_Type','Fluid_Weight','Foam_Test',
  'Iron','Chromium','Nickel','Aluminum','Copper','Tin','Lead',
  'Silicon','Sodium','Potasium','Natrium',
  'Moly','Magnesium','Calcium','Zinc','Phosphorus','Boron',
  'TBN','Soot','Oxidation','Sulfation','Nitration','Water','Fuel',
  'Visc_40C','Visc_100C','ISO4406','PQI'
];

function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  try {
    template.logoUrl = getLogoUrl();
  } catch(e) {
    template.logoUrl = '';
  }
  return template.evaluate()
    .setTitle('Oil Analysis System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── GET ALL DATA ─────────────────────────────────────────────
function getAllData() {
  const sheet = getSheet();
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 3) return [];
  const raw = sheet.getRange(3, 1, last - 2, COLUMNS.length).getValues();
  return raw.map(r => {
    const obj = {};
    COLUMNS.forEach((k, i) => {
    const val = r[i];
    if (val instanceof Date) {
      const yyyy = val.getFullYear();
      const mm   = String(val.getMonth() + 1).padStart(2, '0');
      const dd   = String(val.getDate()).padStart(2, '0');
      obj[k] = `${yyyy}-${mm}-${dd}`;
    } else {
      obj[k] = val !== undefined && val !== null ? String(val) : '';
    }
  });
    return obj;
  }).filter(r => r.ID_Vendor);
}

// ── GET DASHBOARD STATS ──────────────────────────────────────
// Reads sheet directly and returns only aggregated data (no raw rows sent to client).
// Much faster for large sheets because we never build or transfer the full row objects.
function getDashboardData() {
  const sheet = getSheet();
  if (!sheet) return { empty: true };
  const last = sheet.getLastRow();
  if (last < 3) return { empty: true };

  // Read only the columns we need for aggregation:
  // Col indices (1-based): Source=1, Rating=3, Sampled_Date=7, Month=8,
  //   Section=9, Model=10, Asset_ID=11, Component=12
  // We fetch them as a 2D array — much cheaper than getAllData() which
  // builds full objects for every column.
  const NEEDED_COLS = [1, 3, 7, 8, 9, 10, 11, 12]; // 1-based col numbers
  const numRows = last - 2;

  // Build a batch of getRangeValues per column group to minimise API calls.
  // Fastest: fetch the full width we need (col 1 → col 12 = 12 cols).
  const raw = sheet.getRange(3, 1, numRows, 12).getValues();

  const total_raw = raw.length;
  const byRating    = { A: 0, B: 0, C: 0, X: 0, other: 0 };
  const byComponent = {};
  const byVendor    = {};
  const byMonth     = {};
  const byAssetRating = {};
  let total = 0;

  for (let i = 0; i < total_raw; i++) {
    const r = raw[i];
    const idVendor = r[1]; // col B
    if (!idVendor) continue; // skip blank rows
    total++;

    const source  = String(r[0]  || 'Unknown');   // col A  Source
    const rating  = String(r[2]  || '').toUpperCase(); // col C  Rating
    const sampled = r[6];                          // col G  Sampled_Date
    const month   = r[7];                          // col H  Month
    const section = String(r[8]  || '');           // col I  Section
    const model   = String(r[9]  || '');           // col J  Model
    const assetId = String(r[10] || 'Unknown');    // col K  Asset_ID
    const comp    = String(r[11] || 'Unknown');    // col L  Component

    // rating counts
    if      (rating === 'A') byRating.A++;
    else if (rating === 'B') byRating.B++;
    else if (rating === 'C') byRating.C++;
    else if (rating === 'X') byRating.X++;
    else                     byRating.other++;

    // by component
    if (!byComponent[comp]) byComponent[comp] = { A: 0, B: 0, C: 0, X: 0 };
    byComponent[comp][rating] = (byComponent[comp][rating] || 0) + 1;

    // by vendor/source
    byVendor[source] = (byVendor[source] || 0) + 1;

    // by month — prefer Month col, fallback to Sampled_Date substring
    let m = '';
    if (month) {
      m = String(month).substring(0, 7);
    } else if (sampled) {
      if (sampled instanceof Date) {
        const yyyy = sampled.getFullYear();
        const mm   = String(sampled.getMonth() + 1).padStart(2, '0');
        m = yyyy + '-' + mm;
      } else {
        m = String(sampled).substring(0, 7);
      }
    }
    if (m) {
      if (!byMonth[m]) byMonth[m] = { A: 0, B: 0, C: 0, X: 0 };
      byMonth[m][rating] = (byMonth[m][rating] || 0) + 1;
    }

    // by asset
    if (!byAssetRating[assetId]) byAssetRating[assetId] = { A: 0, B: 0, C: 0, X: 0, model: model, section: section };
    byAssetRating[assetId][rating] = (byAssetRating[assetId][rating] || 0) + 1;
  }

  if (total === 0) return { empty: true };

  const months = Object.keys(byMonth).sort();
  const topBad = Object.entries(byAssetRating)
    .map(([id, vv]) => ({ id, A: vv.A||0, B: vv.B||0, C: vv.C||0, X: vv.X||0,
                          model: vv.model, section: vv.section,
                          score: (vv.X * 4) + (vv.C * 3) + (vv.B * 1) }))
    .sort((a, b) => b.score - a.score)
    .slice(0, 10);

  return { total, byRating, byComponent, byVendor, byMonth, months, topBad };
}

// ── GET TREND DATA ───────────────────────────────────────────
// Reads sheet directly — only fetches columns needed for trend chart,
// avoiding the full getAllData() round-trip.
function getTrendData(assetId, component) {
  const sheet = getSheet();
  if (!sheet) return { trend: {}, dates: [], records: [] };
  const last = sheet.getLastRow();
  if (last < 3) return { trend: {}, dates: [], records: [] };

  // Columns needed (1-based):
  // B=2 ID_Vendor, C=3 Rating, G=7 Sampled_Date, K=11 Asset_ID, L=12 Component
  // Wear/condition elements start at col U=21 (Iron) through AV=48 (PQI)
  // We fetch col 1-48 in one call then pick what we need.
  const numRows = last - 2;
  const raw = sheet.getRange(3, 1, numRows, 48).getValues();

  const elements = ['Iron','Chromium','Nickel','Aluminum','Copper','Silicon','TBN','Visc_100C','Oxidation','Soot'];
  // Element column index (0-based) matching COLUMNS array positions
  const elemIdx = { Iron:20, Chromium:21, Nickel:22, Aluminum:23, Copper:24,
                    Silicon:27, TBN:37, Visc_100C:45, Oxidation:39, Soot:38 };

  const trend = {};
  elements.forEach(el => { trend[el] = []; });

  const filtered = [];
  for (let i = 0; i < raw.length; i++) {
    const r = raw[i];
    if (!r[1]) continue; // skip blank rows (ID_Vendor col B)
    const rowAsset = String(r[10] || '');  // col K Asset_ID
    const rowComp  = String(r[11] || '');  // col L Component
    if (assetId    && rowAsset !== assetId)    continue;
    if (component  && rowComp  !== component)  continue;

    let sd = r[6]; // col G Sampled_Date
    if (sd instanceof Date) {
      const yyyy = sd.getFullYear();
      const mm   = String(sd.getMonth() + 1).padStart(2, '0');
      const dd   = String(sd.getDate()).padStart(2, '0');
      sd = yyyy + '-' + mm + '-' + dd;
    } else {
      sd = String(sd || '').substring(0, 10);
    }

    filtered.push({ sd, idVendor: String(r[1]), rating: String(r[2] || ''), row: r });
  }

  // sort by date
  filtered.sort((a, b) => a.sd.localeCompare(b.sd));

  filtered.forEach(item => {
    elements.forEach(el => {
      const vv = parseFloat(item.row[elemIdx[el]]);
      trend[el].push({ date: item.sd, value: isNaN(vv) ? null : vv, id: item.idVendor, rating: item.rating });
    });
  });

  return { trend, dates: filtered.map(x => x.sd), records: [] };
}

// ── GET FILTER OPTIONS (deduplicated & trimmed) ──────────────
function getFilterOptions() {
  const data = getAllData();
  const uniq = (arr) => [...new Set(arr.map(s => (s || '').trim()).filter(Boolean))].sort();
  return {
    vendors:    uniq(data.map(r => r.Source)),
    sections:   uniq(data.map(r => r.Section)),
    models:     uniq(data.map(r => r.Model)),
    assets:     uniq(data.map(r => r.Asset_ID)),
    components: uniq(data.map(r => r.Component))
  };
}

// ── SAVE RECORD (create or update) ──────────────────────────
function saveRecord(record) {
  try {
    const sheet = ensureSheet();
    const data = getAllData();
    const idVendor = (record.ID_Vendor || '').trim();
    if (!idVendor) return { success: false, error: 'ID_Vendor kosong.' };

    const row = COLUMNS.map(k => record[k] !== undefined ? record[k] : '');
    const existingIdx = data.findIndex(r => r.ID_Vendor === idVendor);

    if (existingIdx >= 0) {
      sheet.getRange(existingIdx + 3, 1, 1, COLUMNS.length).setValues([row]);
      return { success: true, action: 'updated' };
    } else {
      const last = Math.max(sheet.getLastRow(), 2);
      sheet.getRange(last + 1, 1, 1, COLUMNS.length).setValues([row]);
      return { success: true, action: 'created' };
    }
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ── DELETE RECORD ────────────────────────────────────────────
function deleteRecord(idVendor) {
  try {
    const sheet = getSheet();
    if (!sheet) return { success: false, error: 'Sheet tidak ditemukan.' };
    const data = getAllData();
    const idx = data.findIndex(r => r.ID_Vendor === idVendor);
    if (idx < 0) return { success: false, error: 'Record tidak ditemukan.' };
    sheet.deleteRow(idx + 3);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ── PREVIEW CSV (mapping only, no write) ─────────────────────
function previewCSV(csvContent, vendorType) {
  try {
    _assetDbCache = null; // always refresh Database sheet on each import
    const rows = parseCSV(csvContent);
    if (!rows || rows.length < 2) return { success: false, error: 'CSV kosong atau tidak valid.' };
    let mapped = [];
    switch (vendorType) {
      case 'PAP':        mapped = mapPAP(rows);        break;
      case 'TU':         mapped = mapTU(rows);         break;
      case 'Microlab':   mapped = mapMicrolab(rows);   break;
      case 'Tekonomiks': mapped = mapTekonomiks(rows); break;
      default: return { success: false, error: 'Vendor tidak dikenal: ' + vendorType };
    }
    if (!mapped.length) return { success: false, error: 'Tidak ada baris berhasil dipetakan dari CSV.' };
    return { success: true, rows: mapped, count: mapped.length };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ── WRITE ROWS TO SHEET (after preview/edit) ─────────────────
function writeRowsToSheet(rows) {
  try {
    if (!rows || !rows.length) return { success: false, error: 'Tidak ada data.' };
    const result = writeToSheet(rows);
    return { success: true, count: result.written, skipped: result.skipped, skippedIds: result.skippedIds };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ── PROCESS CSV IMPORT (legacy, still kept) ──────────────────
function processCSV(csvContent, vendorType) {
  try {
    _assetDbCache = null; // always refresh Database sheet on each import
    const rows = parseCSV(csvContent);
    if (!rows || rows.length < 2) return { success: false, error: 'CSV kosong atau tidak valid.' };
    let mapped = [];
    switch (vendorType) {
      case 'PAP':        mapped = mapPAP(rows);        break;
      case 'TU':         mapped = mapTU(rows);         break;
      case 'Microlab':   mapped = mapMicrolab(rows);   break;
      case 'Tekonomiks': mapped = mapTekonomiks(rows); break;
      default: return { success: false, error: 'Vendor tidak dikenal.' };
    }
    if (!mapped.length) return { success: false, error: 'Tidak ada baris berhasil dipetakan.' };
    const result = writeToSheet(mapped);
    return { success: true, count: result.written, skipped: result.skipped };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function clearAllData() {
  const sheet = getSheet();
  if (!sheet) return { success: false };
  const last = sheet.getLastRow();
  if (last > 2) sheet.deleteRows(3, last - 2);
  return { success: true };
}

// ── EXPOSE DATABASE SHEET INFO TO FRONTEND ───────────────────
// Returns a summary of the Database sheet for display in Import page.
function getAssetDatabaseInfo() {
  _assetDbCache = null; // force fresh read
  const db = getAssetDatabase();
  const keys = Object.keys(db);
  return {
    count:   keys.length,
    sample:  keys.slice(0, 5).map(k => db[k]),
    missing: [] // populated during import preview if needed
  };
}

// ── SHEET HELPERS ────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME);
}

function ensureSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const h1 = new Array(COLUMNS.length).fill('');
    h1[20] = 'WEAR ELEMENTS'; h1[27] = 'CONTAMINANT';
    h1[31] = 'ADDITIVES'; h1[37] = 'OIL CONDITION';
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([h1]);
    const h2 = ['Source','ID Vendor','Rating','Health','Interp. Text','Report Date','Sampled Date','Month','Section','Model','Asset ID','Component','Meter On Asset','Meter on Fluid','PM Type','Fluid Changed','Fluid Brand','Fluid Type','Fluid Weight','Foam Test','Iron','Chromium','Nickel','Aluminum','Copper','Tin','Lead','Silicon','Sodium','Potasium','Natrium','Moly','Magnesium','Calcium','Zinc','Phosphorus','Boron','TBN','Soot','Oxidation','Sulfation','Nitration','Water','Fuel','Visc 40°C','Visc 100°C','ISO4406','PQI'];
    sheet.getRange(2, 1, 1, COLUMNS.length).setValues([h2]);
    const hdr = sheet.getRange(1, 1, 2, COLUMNS.length);
    hdr.setBackground('#0f2744'); hdr.setFontColor('#ffffff'); hdr.setFontWeight('bold');
    sheet.setFrozenRows(2);
  }
  return sheet;
}

function writeToSheet(rows) {
  const sheet = ensureSheet();

  const existingLabNos = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    const existingIds = sheet.getRange(3, 2, lastRow - 2, 1).getValues();
    existingIds.forEach(r => {
      const id = (r[0] || '').toString().trim();
      if (id) existingLabNos.add(id);
    });
  }

  const toWrite   = [];
  const skipped   = [];
  rows.forEach(row => {
    const labNo = (row[1] || '').toString().trim();
    if (!labNo) { skipped.push('(kosong)'); return; }
    if (existingLabNos.has(labNo)) { skipped.push(labNo); return; }
    existingLabNos.add(labNo);
    toWrite.push(row);
  });

  if (toWrite.length > 0) {
    const startRow = Math.max(sheet.getLastRow(), 2) + 1;
    const vals = toWrite.map(r => COLUMNS.map((_, i) => r[i] !== undefined ? r[i] : ''));
    sheet.getRange(startRow, 1, vals.length, COLUMNS.length).setValues(vals);
  }

  return { written: toWrite.length, skipped: skipped.length, skippedIds: skipped };
}

// ── CSV PARSER ───────────────────────────────────────────────
function parseCSV(text) {
  const lines = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n');
  return lines.map(line => {
    const res = []; let cur = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') { if (inQ && line[i+1]==='"'){cur+='"';i++;} else inQ=!inQ; }
      else if (c === ',' && !inQ) { res.push(cur.trim()); cur=''; }
      else cur += c;
    }
    res.push(cur.trim()); return res;
  }).filter(r => r.some(c => c !== ''));
}

function hi(headers, names) {
  const lc = headers.map(h => (h||'').toLowerCase().trim());
  for (const n of names) { const i = lc.indexOf(n.toLowerCase().trim()); if(i>=0) return i; }
  return -1;
}
function v(row, idx) { return (idx>=0 && idx<row.length) ? (row[idx]||'').trim() : ''; }
function fmtDate(raw) {
  if (!raw) return '';
  const d = new Date(raw); if(isNaN(d)) return raw;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Normalises oil-change/fluid-changed values from any vendor format to 'Yes' or 'No'.
// Handles: true/false, 1/0, y/n, yes/no, changed/not changed, and empty values.
function normalizeOilChange(raw) {
  const s = (raw || '').toString().trim().toLowerCase();
  if (['yes', 'true', '1', 'y', 'changed', 'oil changed', 'oc'].includes(s)) return 'Yes';
  if (['no', 'false', '0', 'n', 'not changed', 'no change', 'nc'].includes(s)) return 'No';
  return s ? raw.toString().trim() : ''; // preserve unknown values as-is
}

function healthToRating(health) {
  const t = (health || '').toString().trim();
  const tl = t.toLowerCase();
  if (['no action required', 'acceptable', 'normal', 'good',
       'satisfactory', 'a', '1'].includes(tl)) return 'A';
  if (['monitor component', 'unacceptable', 'caution', 'watch',
       'monitor', 'b', '2'].includes(tl)) return 'B';
  if (['action required', 'critical', 'problem',
       'c', '3'].includes(tl)) return 'C';
  if (['urgent action required', 'severe', 'x'].includes(tl)) return 'X';
  return t ? t : '';
}

function normRating(raw) {
  return healthToRating(raw);
}

// Converts a final rating letter back to a standardised health label (Microlab-style).
function ratingToHealth(rating) {
  switch ((rating || '').toString().trim().toUpperCase()) {
    case 'A': return 'Acceptable';
    case 'B': return 'Monitor';
    case 'C': return 'Action Required';
    case 'X': return 'Urgent Action Required';
    default:  return rating || '';
  }
}

// PAP-specific eval code mapping (takes priority over healthToRating for PAP imports).
// N → A (Normal/No Action), B → B (Monitor), C → C (Critical), D → X (Urgent/Severe).
// Returns null if the eval code is not a recognised PAP single-letter code,
// so the caller can fall back to healthToRating on the Health column.
function papEvalCodeToRating(evalCode) {
  switch ((evalCode || '').toString().trim().toUpperCase()) {
    case 'N': return 'A';
    case 'B': return 'B';
    case 'C': return 'C';
    case 'D': return 'X';
    default:  return null; // not a PAP eval code — let healthToRating handle it
  }
}

// ── COMPONENT NORMALIZATION ──────────────────────────────────
// Maps vendor-specific component names to standardised EH names.
// Comparison is case-insensitive and trims whitespace.
const COMPONENT_MAP = {
  'BRAKE COOLING OIL':            'BRAKE COOLING',
  'CIRCLE':                       'CIRCLE DRIVE',
  'CIRCLE REVERSE GEAR':          'CIRCLE DRIVE',
  'CIRCLE DRIVE BOX':             'CIRCLE DRIVE',
  'DIFFERENTIAL REAR':            'DIFFERENTIAL',
  'ENGINE ASSY':                  'ENGINE',
  'ENGINE GP':                    'ENGINE',
  'FINAL DRIVE LEFT':             'FINAL DRIVE LH',
  'FINAL DRIVE RIGHT':            'FINAL DRIVE RH',
  'FINAL DRIVE REAR LEFT':        'FINAL DRIVE LH',
  'FINAL DRIVE REAR RIGHT':       'FINAL DRIVE RH',
  'LEFT FINAL DRIVE':             'FINAL DRIVE LH',
  'RIGHT FINAL DRIVE':            'FINAL DRIVE RH',
  'LEFT WHEEL HUB':               'FRONT WHEEL LH',
  'RIGHT WHEEL HUB':              'FRONT WHEEL RH',
  'FRONT LEFT WHEEL BEARING':     'FRONT WHEEL LH',
  'FRONT RIGHT WHEEL BEARING':    'FRONT WHEEL RH',
  'WHEEL BEARINGS FRONT LEFT':    'FRONT WHEEL LH',
  'WHEEL BEARINGS FRONT RIGHT':   'FRONT WHEEL RH',
  'SWING':                        'SWING DRIVE',
  'SWING BOX':                    'SWING DRIVE',
  'SWING MACHINERY':              'SWING DRIVE',
  'FRONT SWING MACHINERY':        'SWING DRIVE FRONT',
  'REAR SWING MACHINERY':         'SWING DRIVE REAR',
  'TANDEM LEFT':                  'TANDEM LH',
  'TANDEM RIGHT':                 'TANDEM RH',
  'LEFT TANDEM':                  'TANDEM LH',
  'RIGHT TANDEM':                 'TANDEM RH',
  'TRANSMISSION DIFF':            'TRANSMISSION',
  'TRANSMISSION POWER SHIFT':     'TRANSMISSION',
  'POWER TAKE OFF':               'TRANSMISSION',
  'STEERING SYSTEM':              'STEERING',
  'HYDRAULIC SYSTEM':             'HYDRAULIC',
  'HYDRAULIC FOAMING':            'HYDRAULIC',
};

function normalizeComponent(raw) {
  if (!raw) return '';
  const key = raw.toString().trim().toUpperCase();
  return COMPONENT_MAP[key] || raw.toString().trim();
}

// ── ASSET DATABASE LOOKUP ────────────────────────────────────
// Reads the "Database" sheet and builds a lookup map:
//   normalised-asset-id → { assetId: canonical, model: string, section: string }
// The lookup is cached per-execution to avoid repeated sheet reads.
let _assetDbCache = null;

function getAssetDatabase() {
  if (_assetDbCache) return _assetDbCache;

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Database');
  if (!sheet) { _assetDbCache = {}; return {}; }

  const last = sheet.getLastRow();
  if (last < 2) { _assetDbCache = {}; return {}; }

  // Read entire sheet as values; row 1 = header
  const raw  = sheet.getRange(1, 1, last, sheet.getLastColumn()).getValues();
  const hdr  = raw[0].map(h => (h || '').toString().trim().toLowerCase());

  // Locate relevant columns — accept common name variations
  const colAsset   = _findCol(hdr, ['asset-code','asset code','asset_code','asset id','asset_id','unit id','unit_id','plant id','plantid','machine id']);
  const colModel   = _findCol(hdr, ['model','unit model','unitmodel','machine name','machine']);
  const colSection = _findCol(hdr, ['category','section','site','location','department']);

  if (colAsset < 0 || colModel < 0) {
    Logger.log('Database sheet: could not find Asset ID or Model column. Headers: ' + hdr.join(', '));
    _assetDbCache = {};
    return {};
  }

  const db = {};
  for (let i = 1; i < raw.length; i++) {
    const r       = raw[i];
    const rawId   = (r[colAsset]  || '').toString().trim();
    const model   = (r[colModel]  || '').toString().trim();
    const section = colSection >= 0 ? (r[colSection] || '').toString().trim() : '';
    if (!rawId) continue;
    const key = _normaliseId(rawId);
    db[key] = { assetId: rawId, model: model, section: section };
  }

  _assetDbCache = db;
  return db;
}

// Look up canonical assetId + model from Database sheet.
// Returns { assetId, model, section } — falls back to rawId if not found.
function lookupAsset(rawId) {
  if (!rawId) return { assetId: '', model: '', section: '' };
  const db  = getAssetDatabase();
  const key = _normaliseId(rawId);
  if (db[key]) return db[key];
  // Try partial match: strip leading zeros, spaces, dashes
  for (const k of Object.keys(db)) {
    if (_normaliseId(k) === key) return db[k];
  }
  // Not found — return original id without model override
  return { assetId: rawId, model: '', section: '' };
}

function _findCol(headers, names) {
  for (const n of names) {
    const i = headers.indexOf(n.toLowerCase().trim());
    if (i >= 0) return i;
  }
  return -1;
}

// Normalise an asset ID for fuzzy matching:
// upper-case, strip leading zeros, collapse spaces/dashes/underscores
function _normaliseId(id) {
  return id.toString().toUpperCase()
    .replace(/[\s\-_]+/g, '')   // remove spaces, dashes, underscores
    .replace(/^0+/, '');         // strip leading zeros
}

// ── VENDOR MAPPINGS ──────────────────────────────────────────

function mapPAP(rows) {
  const h = rows[0];
  const cols = {
    labNo:    h.indexOf('Lab_No'),
    model:    h.indexOf('MODEL'),
    unit:     h.indexOf('UNIT_NO'),
    comp:     h.indexOf('COMPONENT'),
    samplDt:  h.indexOf('SAMPL_DT1'),
    rptDt:    h.indexOf('RPT_DT1'),
    hrsOH:    h.indexOf('HRS_KM_OH'),
    hrsOC:    h.indexOf('HRS_KM_OC'),
    hrsTot:   h.indexOf('HRS_KM_TOT'),
    oilChg:   h.indexOf('oil_change'),
    oilType:  h.indexOf('OIL_TYPE'),
    matrix:   h.indexOf('OIL_MATRIX'),
    origVisc: h.indexOf('ORIG_VISC'),
    recomm:   hi(h, ['Recomm1 ', 'Recomm1']),
    health:   hi(h, ['Health ', 'Health', 'HEALTH']),
    evalCode: hi(h, ['Eval Code ', 'Eval Code', 'EVAL_CODE', 'EvalCode']),
    v100:     hi(h, ['Visc@100C ', 'Visc@100C']),
    v40:      hi(h, ['Visc@40C ', 'Visc@40C']),
    tbn:      hi(h, ['T B N ', 'TBN']),
    si:       hi(h, ['Silicon (Si)  ', 'Silicon (Si)']),
    na:       hi(h, ['Natrium (Na) ', 'Natrium (Na)']),
    mg:       hi(h, ['Magnesium (Mg) ']),
    ca:       hi(h, ['Calcium (Ca) ']),
    zn:       hi(h, ['Zinc (Zn) ']),
    ni:       hi(h, ['Nickel (Ni) ']),
    fe:       hi(h, ['Iron (Fe) ']),
    cu:       hi(h, ['Copper (Cu) ']),
    al:       hi(h, ['Alumunium (Al) ']),
    cr:       hi(h, ['Chromium (Cr) ']),
    sn:       hi(h, ['Tin (Sn) ']),
    pb:       hi(h, ['Lead (Pb) ']),
    mo:       hi(h, ['Molybdenum (Mo) ']),
    b:        hi(h, ['Boron (B) ']),
    k:        hi(h, ['Potassium (K) ']),
    p:        hi(h, ['Phosphor (P) ']),
    fuel:     hi(h, ['Fuel Dilution ']),
    soot:     hi(h, ['Soot ']),
    oxi:      hi(h, ['Oxidation ']),
    nit:      hi(h, ['Nitration ']),
    water:    hi(h, ['Water Content ']),
    sulfat:   hi(h, ['Sulfation ']),
    iso:      hi(h, ['ISO CODE ']),
    pq:       hi(h, ['PQ Index '])
  };

  const res = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every(c => !c) || !v(r, cols.labNo)) continue;

    const hlth     = v(r, cols.health) || '';
    const evalCode = v(r, cols.evalCode) || '';
    // PAP eval code takes priority: N→A, B→B, C→C, D→X
    // If eval code is not a recognised PAP code, fall back to healthToRating on Health column
    const rat = papEvalCodeToRating(evalCode) || healthToRating(hlth) || healthToRating(evalCode);
    const sd   = fmtDate(v(r, cols.samplDt));
    const assetLookup = lookupAsset(v(r, cols.unit));

    const row = new Array(48).fill('');
    row[0]  = 'PAP';
    row[1]  = v(r, cols.labNo);
    row[2]  = rat;
    row[3]  = ratingToHealth(rat);  // distandarisasi sesuai Microlab
    row[4]  = v(r, cols.recomm);
    row[5]  = fmtDate(v(r, cols.rptDt));
    row[6]  = sd;
    row[7]  = sd.substring(0, 7);
    row[8]  = assetLookup.section || '';
    row[9]  = assetLookup.model   || v(r, cols.model);  // prefer Database model
    row[10] = assetLookup.assetId || v(r, cols.unit);   // canonical asset ID
    row[11] = normalizeComponent(v(r, cols.comp));
    row[12] = v(r, cols.hrsTot);
    row[13] = v(r, cols.hrsOH);
    row[14] = '';
    row[15] = normalizeOilChange(v(r, cols.oilChg));
    row[16] = v(r, cols.oilType);
    row[17] = v(r, cols.matrix);
    row[18] = v(r, cols.origVisc);
    row[19] = '';
    row[20] = v(r, cols.fe);
    row[21] = v(r, cols.cr);
    row[22] = v(r, cols.ni);
    row[23] = v(r, cols.al);
    row[24] = v(r, cols.cu);
    row[25] = v(r, cols.sn);
    row[26] = v(r, cols.pb);
    row[27] = v(r, cols.si);
    row[28] = v(r, cols.na);
    row[29] = v(r, cols.k);
    row[30] = v(r, cols.na);
    row[31] = v(r, cols.mo);
    row[32] = v(r, cols.mg);
    row[33] = v(r, cols.ca);
    row[34] = v(r, cols.zn);
    row[35] = v(r, cols.p);
    row[36] = v(r, cols.b);
    row[37] = v(r, cols.tbn);
    row[38] = v(r, cols.soot);
    row[39] = v(r, cols.oxi);
    row[40] = v(r, cols.sulfat);
    row[41] = v(r, cols.nit);
    row[42] = v(r, cols.water);
    row[43] = v(r, cols.fuel);
    row[44] = v(r, cols.v40);
    row[45] = v(r, cols.v100);
    row[46] = v(r, cols.iso);
    row[47] = v(r, cols.pq);
    res.push(row);
  }
  return res;
}

function mapTU(rows) {
  const h = rows[0];
  const c = {
    health:   hi(h, ['Health']),
    labNo:    hi(h, ['Lab No.']),
    labDate:  hi(h, ['Lab Date']),
    sampled:  hi(h, ['Sampled Date']),
    model:    hi(h, ['Model']),
    asset:    hi(h, ['Asset ID']),
    comp:     hi(h, ['Component']),
    meter:    hi(h, ['Meter']),
    mof:      hi(h, ['Meter on Fluid']),
    fluidChg: hi(h, ['Fluid Changed']),
    brand:    hi(h, ['Fluid Brand']),
    type:     hi(h, ['Fluid Type']),
    weight:   hi(h, ['Fluid Weight']),
    interp:   hi(h, ['Interp. Text']),
    distDate: hi(h, ['Distribution Date']),
    fe:       hi(h, ['Fe']),
    cr:       hi(h, ['Cr']),
    ni:       hi(h, ['Ni']),
    al:       hi(h, ['Al']),
    cu:       hi(h, ['Cu']),
    sn:       hi(h, ['Sn']),
    pb:       hi(h, ['Pb']),
    si:       hi(h, ['Si']),
    na:       hi(h, ['Na']),
    mo:       hi(h, ['Mo']),
    mg:       hi(h, ['Mg', 'Mg*']),
    ca:       hi(h, ['Ca']),
    zn:       hi(h, ['Zn', 'Zn*']),
    p:        hi(h, ['P']),
    b:        hi(h, ['B']),
    k:        hi(h, ['K*', 'K']),
    v100:     hi(h, ['V100']),
    v40:      hi(h, ['V40']),
    tbn:      hi(h, ['TBN']),
    iso:      hi(h, ['ISO']),
    pqi:      hi(h, ['PQI', 'PCV']),
    oxi:      hi(h, ['OXI']),
    nit:      hi(h, ['NIT']),
    sul:      hi(h, ['SUL']),
    st:       hi(h, ['ST']),
    water:    hi(h, ['W', 'Water %']),
    fuel:     hi(h, ['Fuel', 'F'])
  };

  const res = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every(x => !x) || !v(r, c.labNo)) continue;

    const hlth = v(r, c.health);
    const rat  = healthToRating(hlth);
    const sd   = fmtDate(v(r, c.sampled));
    const assetLookup = lookupAsset(v(r, c.asset));

    const row = new Array(48).fill('');
    row[0]  = 'TU SOS';
    row[1]  = v(r, c.labNo);
    row[2]  = rat;
    row[3]  = hlth;
    row[4]  = v(r, c.interp);
    row[5]  = fmtDate(v(r, c.distDate) || v(r, c.labDate));
    row[6]  = sd;
    row[7]  = sd.substring(0, 7);
    row[8]  = assetLookup.section || '';
    row[9]  = assetLookup.model   || v(r, c.model);   // prefer Database model
    row[10] = assetLookup.assetId || v(r, c.asset);   // canonical asset ID
    row[11] = normalizeComponent(v(r, c.comp));
    row[12] = v(r, c.meter);
    row[13] = v(r, c.mof);
    row[14] = '';
    row[15] = normalizeOilChange(v(r, c.fluidChg));
    row[16] = v(r, c.brand);
    row[17] = v(r, c.type);
    row[18] = v(r, c.weight);
    row[19] = '';
    row[20] = v(r, c.fe);
    row[21] = v(r, c.cr);
    row[22] = v(r, c.ni);
    row[23] = v(r, c.al);
    row[24] = v(r, c.cu);
    row[25] = v(r, c.sn);
    row[26] = v(r, c.pb);
    row[27] = v(r, c.si);
    row[28] = v(r, c.na);
    row[29] = v(r, c.k);
    row[30] = v(r, c.na);
    row[31] = v(r, c.mo);
    row[32] = v(r, c.mg);
    row[33] = v(r, c.ca);
    row[34] = v(r, c.zn);
    row[35] = v(r, c.p);
    row[36] = v(r, c.b);
    row[37] = v(r, c.tbn);
    row[38] = v(r, c.st);
    row[39] = v(r, c.oxi);
    row[40] = v(r, c.sul);
    row[41] = v(r, c.nit);
    row[42] = v(r, c.water);
    row[43] = v(r, c.fuel);
    row[44] = v(r, c.v40);
    row[45] = v(r, c.v100);
    row[46] = v(r, c.iso);
    row[47] = v(r, c.pqi);
    res.push(row);
  }
  return res;
}

function mapMicrolab(rows) {
  const h = rows[0];
  const c = {
    sampleId:   hi(h, ['SampleID']),
    unitId:     hi(h, ['UnitID']),
    unitModel:  hi(h, ['UnitModel']),
    compDesc:   hi(h, ['ComponentDescription']),
    compModel:  hi(h, ['ComponentModel']),
    sampleDate: hi(h, ['SampleTakenDate']),
    analDate:   hi(h, ['AnalysisDate']),
    unitHrs:    hi(h, ['UnitTimeHours']),
    oilTime:    hi(h, ['UnitTimeOnOil']),
    oilChg:     hi(h, ['OilChanged']),
    oilBrand:   hi(h, ['OilBrand']),
    oilType:    hi(h, ['OilType']),
    oilWt:      hi(h, ['OilWeight']),
    result:     hi(h, ['ResultStatement']),
    comment:    hi(h, ['Comment']),
    status:     hi(h, ['SampleStatus']),
    sampleDesc: hi(h, ['SampleDescription']),
    fe:         hi(h, ['Iron']),
    cr:         hi(h, ['Chromium']),
    ni:         hi(h, ['Nickel']),
    al:         hi(h, ['Aluminum']),
    cu:         hi(h, ['Copper']),
    sn:         hi(h, ['Tin']),
    pb:         hi(h, ['Lead']),
    si:         hi(h, ['Silicon']),
    na:         hi(h, ['Sodium']),
    k:          hi(h, ['Potassium']),
    mo:         hi(h, ['Molybdenum']),
    mg:         hi(h, ['Magnesium']),
    ca:         hi(h, ['Calcium']),
    zn:         hi(h, ['Zinc']),
    p:          hi(h, ['Phosphorus']),
    b:          hi(h, ['Boron']),
    tbn:        hi(h, ['TBN']),
    soot:       hi(h, ['Soot']),
    oxi:        hi(h, ['Oxidation']),
    sulf:       hi(h, ['Sulfation']),
    nit:        hi(h, ['Nitration']),
    water:      hi(h, ['Water']),
    fuel:       hi(h, ['Fuel']),
    v40:        hi(h, ['V40C']),
    v100:       hi(h, ['V100C']),
    iso:        hi(h, ['ISO Code'])
  };

  const res = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every(x => !x) || !v(r, c.sampleId)) continue;

    // Rating diambil dari kolom "Comment", kata sebelum simbol '|'.
    // Contoh: "Acceptable|detail..." → "Acceptable" → A
    //         "Unacceptable|detail..." → "Unacceptable" → B/C
    // Fallback ke ResultStatement jika Comment kosong.
    const commentRaw = v(r, c.comment);
    const resultRaw  = v(r, c.result);
    const hlthSrc    = commentRaw || resultRaw;
    const hlth       = hlthSrc.split('|')[0].trim();
    const rat        = healthToRating(hlth);
    const sd   = fmtDate(v(r, c.sampleDate));
    const assetLookup = lookupAsset(v(r, c.unitId));

    const row = new Array(48).fill('');
    row[0]  = 'Microlab';
    row[1]  = v(r, c.sampleId);
    row[2]  = rat;
    row[3]  = hlth;
    row[4]  = v(r, c.result);
    row[5]  = fmtDate(v(r, c.analDate));
    row[6]  = sd;
    row[7]  = sd.substring(0, 7);
    row[8]  = assetLookup.section || '';
    row[9]  = assetLookup.model   || v(r, c.unitModel);  // prefer Database model
    row[10] = assetLookup.assetId || v(r, c.unitId);     // canonical asset ID
    row[11] = normalizeComponent(v(r, c.compDesc) || v(r, c.compModel));
    row[12] = v(r, c.unitHrs);
    row[13] = v(r, c.oilTime);
    row[14] = v(r, c.sampleDesc);
    row[15] = normalizeOilChange(v(r, c.oilChg));
    row[16] = v(r, c.oilBrand);
    row[17] = v(r, c.oilType);
    row[18] = v(r, c.oilWt);
    row[19] = '';
    row[20] = v(r, c.fe);
    row[21] = v(r, c.cr);
    row[22] = v(r, c.ni);
    row[23] = v(r, c.al);
    row[24] = v(r, c.cu);
    row[25] = v(r, c.sn);
    row[26] = v(r, c.pb);
    row[27] = v(r, c.si);
    row[28] = v(r, c.na);
    row[29] = v(r, c.k);
    row[30] = v(r, c.na);
    row[31] = v(r, c.mo);
    row[32] = v(r, c.mg);
    row[33] = v(r, c.ca);
    row[34] = v(r, c.zn);
    row[35] = v(r, c.p);
    row[36] = v(r, c.b);
    row[37] = v(r, c.tbn);
    row[38] = v(r, c.soot);
    row[39] = v(r, c.oxi);
    row[40] = v(r, c.sulf);
    row[41] = v(r, c.nit);
    row[42] = v(r, c.water);
    row[43] = v(r, c.fuel);
    row[44] = v(r, c.v40);
    row[45] = v(r, c.v100);
    row[46] = v(r, c.iso);
    row[47] = '';
    res.push(row);
  }
  return res;
}

function mapTekonomiks(rows) {
  const h = rows[0];
  const c = {
    plantId:  hi(h, ['PlantId']),
    machine:  hi(h, ['Machine Name']),
    compName: hi(h, ['Component Name']),
    sampleNo: hi(h, ['SampleNumber']),
    health:   hi(h, ['Health']),
    status:   hi(h, ['OverallSampleStatus']),
    sampled:  hi(h, ['DateSampled']),
    reported: hi(h, ['DateReported']),
    smu:      hi(h, ['SMU']),
    oilHrs:   hi(h, ['OilHours']),
    oilChg:   hi(h, ['OilChanged']),
    oil:      hi(h, ['Oil']),
    diag:     hi(h, ['Diagnosis']),
    fe:       hi(h, ['Iron']),
    cr:       hi(h, ['Chromium']),
    ni:       hi(h, ['Nickel']),
    al:       hi(h, ['Aluminium']),
    cu:       hi(h, ['Copper']),
    sn:       hi(h, ['Tin']),
    pb:       hi(h, ['Lead']),
    si:       hi(h, ['Silicon']),
    na:       hi(h, ['Sodium']),
    mo:       hi(h, ['Molybdenum']),
    mg:       hi(h, ['Magnesium']),
    ca:       hi(h, ['Calcium']),
    zn:       hi(h, ['Zinc']),
    p:        hi(h, ['Phosphorous']),
    b:        hi(h, ['Boron']),
    tbn:      hi(h, ['TBN']),
    soot:     hi(h, ['Soot']),
    oxi:      hi(h, ['Oxidation']),
    sulf:     hi(h, ['Sulphation']),
    nit:      hi(h, ['Nitration']),
    water:    hi(h, ['water']),
    fuel:     hi(h, ['fuel dilution']),
    v40:      hi(h, ['ViscosityAt40C']),
    v100:     hi(h, ['ViscosityAt100C']),
    iso:      hi(h, ['ISOCode']),
    pq:       hi(h, ['PQIndex'])
  };

  const res = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every(x => !x) || !v(r, c.sampleNo)) continue;

    const hlth   = v(r, c.health);
    const statusVal = v(r, c.status).toString().trim();
    // OverallSampleStatus: 0→A, 1→B, 2→C, selain itu→X
    let rat;
    if      (statusVal === '0') rat = 'A';
    else if (statusVal === '1') rat = 'B';
    else if (statusVal === '2') rat = 'C';
    else if (statusVal !== '')  rat = 'X';
    else                        rat = healthToRating(hlth); // fallback jika kolom kosong
    const sd   = fmtDate(v(r, c.sampled));
    const assetLookup = lookupAsset(v(r, c.plantId));

    const row = new Array(48).fill('');
    row[0]  = 'Tekonomiks';
    row[1]  = v(r, c.sampleNo);
    row[2]  = rat;
    row[3]  = ratingToHealth(rat);  // distandarisasi sesuai Microlab
    row[4]  = v(r, c.diag);
    row[5]  = fmtDate(v(r, c.reported));
    row[6]  = sd;
    row[7]  = sd.substring(0, 7);
    row[8]  = assetLookup.section || '';
    row[9]  = assetLookup.model   || v(r, c.machine);  // prefer Database model
    row[10] = assetLookup.assetId || v(r, c.plantId);  // canonical asset ID
    row[11] = normalizeComponent(v(r, c.compName));
    row[12] = v(r, c.smu);
    row[13] = v(r, c.oilHrs);
    row[14] = '';
    row[15] = normalizeOilChange(v(r, c.oilChg));
    row[16] = v(r, c.oil);
    row[17] = '';
    row[18] = '';
    row[19] = '';
    row[20] = v(r, c.fe);
    row[21] = v(r, c.cr);
    row[22] = v(r, c.ni);
    row[23] = v(r, c.al);
    row[24] = v(r, c.cu);
    row[25] = v(r, c.sn);
    row[26] = v(r, c.pb);
    row[27] = v(r, c.si);
    row[28] = v(r, c.na);
    row[29] = '';
    row[30] = v(r, c.na);
    row[31] = v(r, c.mo);
    row[32] = v(r, c.mg);
    row[33] = v(r, c.ca);
    row[34] = v(r, c.zn);
    row[35] = v(r, c.p);
    row[36] = v(r, c.b);
    row[37] = v(r, c.tbn);
    row[38] = v(r, c.soot);
    row[39] = v(r, c.oxi);
    row[40] = v(r, c.sulf);
    row[41] = v(r, c.nit);
    row[42] = v(r, c.water);
    row[43] = v(r, c.fuel);
    row[44] = v(r, c.v40);
    row[45] = v(r, c.v100);
    row[46] = v(r, c.iso);
    row[47] = v(r, c.pq);
    res.push(row);
  }
  return res;
}

function getLogoUrl() {
  var fileId = '1gL_pqye7gg1HWPaqb93tvYk6Uik3egLr';
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var base64 = Utilities.base64Encode(blob.getBytes());
  var mimeType = blob.getContentType() || 'image/png';
  return 'data:' + mimeType + ';base64,' + base64;
}