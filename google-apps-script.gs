const SPREADSHEET_ID = '1-wwYal1BPl8CyhLwONfhAS-AKiW7u5RuEUl2ImbyhJg';
const SHEET_NAME = 'アプリ同期';

function sheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function setup() {
  const sh = sheet_();
  if (sh.getLastRow() === 0) sh.appendRow(['TYPE', 'JSON', 'UPDATED_AT']);
}

function doGet(e) {
  const sh = sheet_();
  const values = sh.getDataRange().getValues();
  const records = [];
  let config = null;

  for (let i = 1; i < values.length; i++) {
    const type = values[i][0];
    const json = values[i][1];
    if (type === 'RECORD') {
      try {
        const r = JSON.parse(json);
        if (r && typeof r.date === 'string' && r.date && Number.isFinite(Number(r.amount))) records.push(r);
      } catch (_) {}
    }
    if (type === 'CONFIG') {
      try { config = JSON.parse(json); } catch (_) {}
    }
  }

  const out = JSON.stringify({ ok: true, records, config });
  const rawCallback = e && e.parameter ? (e.parameter.callback || '') : '';
  const callback = /^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(rawCallback) ? rawCallback : '';

  if (callback) {
    return ContentService.createTextOutput(callback + '(' + out + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const p = JSON.parse(e.postData.contents || '{}');
  const sh = sheet_();
  const old = sh.getDataRange().getValues();
  const map = new Map();

  for (let i = 1; i < old.length; i++) {
    if (old[i][0] !== 'RECORD') continue;
    try {
      const r = JSON.parse(old[i][1]);
      if (r && r.id) map.set(r.id, { r, row: i + 1 });
    } catch (_) {}
  }

  for (const r of (p.records || [])) {
    if (!r || !r.id || typeof r.date !== 'string' || !r.date || !Number.isFinite(Number(r.amount))) continue;
    const o = map.get(r.id);
    if (!o) sh.appendRow(['RECORD', JSON.stringify(r), r.updatedAt || Date.now()]);
    else if ((r.updatedAt || 0) > (o.r.updatedAt || 0)) sh.getRange(o.row, 2, 1, 2).setValues([[JSON.stringify(r), r.updatedAt || Date.now()]]);
  }

  const c = p.config && typeof p.config === 'object' ? p.config : {};
  const data = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'CONFIG') { row = i + 1; break; }
  }

  if (row < 0) sh.appendRow(['CONFIG', JSON.stringify(c), c.updatedAt || Date.now()]);
  else {
    let oldConfig = {};
    try { oldConfig = JSON.parse(data[row - 1][1]); } catch (_) {}
    if ((c.updatedAt || 0) > (oldConfig.updatedAt || 0)) sh.getRange(row, 2, 1, 2).setValues([[JSON.stringify(c), c.updatedAt || Date.now()]]);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
