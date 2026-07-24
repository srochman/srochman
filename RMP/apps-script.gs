// RMP — Activity Tracker · Google Apps Script Backend
// Paste deployed URL ke db.js setelah deploy sebagai New Deployment

const RMP_SS_ID    = 'YOUR_SPREADSHEET_ID_HERE';
const RMP_PIN      = 'YOUR_PIN_HERE';
const AE_SHEET     = 'AE_LIST';   // kolom: City | TSC | SPV | AE | Aktif
const ACT_SHEET    = 'Activity';

// ── Output helper ─────────────────────────────────────────────────
function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function checkPin(pin) { return String(pin).trim() === String(RMP_PIN).trim(); }

// ── doGet ─────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e.parameter.action || '').trim();
    const pin    = (e.parameter.pin    || '').trim();
    let result;

    if      (action === 'getConfig')     result = getConfig();
    else if (action === 'verifyPin')     result = { ok: checkPin(pin) };
    else if (action === 'getActivities') result = checkPin(pin) ? getActivities() : { error: 'Unauthorized' };
    else                                 result = { error: 'Unknown action: ' + action };

    return out(result);
  } catch (ex) { return out({ error: ex.toString() }); }
}

// ── doPost ────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = (payload.action || '').trim();
    const pin     = (payload.pin    || '').trim();
    let result;

    if (action === 'submitActivity') {
      if (!checkPin(pin)) return out({ error: 'Unauthorized' });
      result = submitActivity(payload.data);
    } else {
      result = { error: 'Unknown action: ' + action };
    }

    return out(result);
  } catch (ex) { return out({ error: ex.toString() }); }
}

// ── getConfig: return AE hierarchy list ──────────────────────────
function getConfig() {
  const ss    = SpreadsheetApp.openById(RMP_SS_ID);
  const sheet = ss.getSheetByName(AE_SHEET);
  if (!sheet) return { error: 'Sheet AE_LIST tidak ditemukan' };

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { ok: true, list: [] };

  const hdr    = rows[0].map(h => String(h).trim().toLowerCase());
  const iCity  = hdr.indexOf('city');
  const iTsc   = hdr.indexOf('tsc');
  const iSpv   = hdr.indexOf('spv');
  const iAe    = hdr.indexOf('ae');
  const iAktif = hdr.indexOf('aktif');

  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r     = rows[i];
    const ae    = String(r[iAe]    || '').trim();
    const aktif = r[iAktif];
    if (!ae) continue;
    if (aktif === false || String(aktif).toUpperCase() === 'FALSE') continue;
    list.push({
      city: String(r[iCity] || '').trim(),
      tsc:  String(r[iTsc]  || '').trim(),
      spv:  String(r[iSpv]  || '').trim(),
      ae,
    });
  }
  return { ok: true, list };
}

// ── getActivities ─────────────────────────────────────────────────
function getActivities() {
  const ss    = SpreadsheetApp.openById(RMP_SS_ID);
  const sheet = ss.getSheetByName(ACT_SHEET);
  if (!sheet) return [];

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const hdr  = rows[0].map(h => String(h).trim());
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const obj = {};
    hdr.forEach((h, j) => {
      const v = rows[i][j];
      obj[h] = v instanceof Date ? v.toISOString() : v;
    });
    data.push(obj);
  }
  return data;
}

// ── submitActivity ────────────────────────────────────────────────
function submitActivity(data) {
  const ss  = SpreadsheetApp.openById(RMP_SS_ID);
  let sheet = ss.getSheetByName(ACT_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(ACT_SHEET);
    sheet.appendRow([
      'ID','Timestamp','City','TSC','SPV','AE',
      'Nama Leads','Alamat','No Telp',
      'Current ISP','Contract Period','Family Needs','Budget',
      'Latitude','Longitude','Accuracy (m)'
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 16).setFontWeight('bold');
  }

  const id = 'RMP-' + new Date().getTime();
  const ts = new Date().toISOString();

  sheet.appendRow([
    id, ts,
    data.city           || '',
    data.tsc            || '',
    data.spv            || '',
    data.ae             || '',
    data.namaLeads      || '',
    data.alamat         || '',
    data.noTelp         || '',
    data.currentISP     || '',
    data.contractPeriod || '',
    data.familyNeeds    || '',
    data.budget         || '',
    data.lat            || '',
    data.lng            || '',
    data.accuracy       || '',
  ]);

  return { ok: true, id };
}
