// =============================================
// SURVIVAL MODE — Apps Script Backend
// =============================================
// SETUP:
// 1. Upload "survival mode.xlsx" ke Google Drive (auto-convert ke Google Sheet)
// 2. Buka spreadsheet > Extensions > Apps Script
// 3. Paste kode ini, Save
// 4. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy URL deployment, paste ke index.html (konstanta API_URL)
// 6. Setiap ada perubahan kode: Deploy > Manage Deployments > New Version
// =============================================

function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === 'getData')   return respond(getData());
    if (action === 'updateRow') return respond(updateRow(JSON.parse(e.parameter.data)));
    return respond({ error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function getData() {
  var ws     = getSheet();
  var values = ws.getDataRange().getValues();

  // Cari baris header (kolom A mengandung 'Customer Account')
  var hRow = 1;
  for (var i = 0; i < Math.min(5, values.length); i++) {
    if (String(values[i][0]).toLowerCase().includes('customer')) { hRow = i; break; }
  }

  var result = [];
  for (var i = hRow + 1; i < values.length; i++) {
    var row = values[i];
    var acc = String(row[0] || '').trim();
    if (!acc || acc === '0') continue;
    result.push({
      rowIndex       : i + 1,
      customerAccount: acc,
      saDate         : String(row[1] || '').trim(),
      cycleDate      : String(row[2] || '').trim(),
      salesName      : String(row[3] || '').trim(),
      spvName        : String(row[4] || '').trim(),
      m2             : String(row[5] || '').trim(),
      m2PaymentDate  : fmtDate(row[6]),
      m3             : String(row[7] || '').trim(),
      m3PaymentDate  : fmtDate(row[8]),
      notes          : String(row[9] || '').trim()
    });
  }
  return { data: result };
}

function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  if (!s || s === '0') return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  if (/^\d{8}$/.test(s)) return s.slice(0,4) + '-' + s.slice(4,6) + '-' + s.slice(6,8);
  return '';
}

function toSheetDate(isoStr) {
  if (!isoStr) return '';
  var m = String(isoStr).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : isoStr;
}

function updateRow(data) {
  var ws     = getSheet();
  var values = ws.getDataRange().getValues();
  var target = String(data.customerAccount || '').trim();

  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() !== target) continue;
    var r = i + 1;
    ws.getRange(r, 6).setValue(data.m2 || '');
    ws.getRange(r, 7).setValue(toSheetDate(data.m2PaymentDate));
    ws.getRange(r, 8).setValue(data.m3 || '');
    ws.getRange(r, 9).setValue(toSheetDate(data.m3PaymentDate));
    ws.getRange(r, 10).setValue(data.notes || '');
    SpreadsheetApp.flush();
    return { success: true, rowIndex: r };
  }
  return { error: 'Account tidak ditemukan: ' + target };
}
