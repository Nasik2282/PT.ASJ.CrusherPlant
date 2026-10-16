/**
 * PT. Aneka Sumberbumi Jaya - Crusher Line
 * Backend penyimpanan laporan ke Google Sheet (Apps Script Web App).
 *
 * PASANG:
 * 1. Buka spreadsheet > Extensions > Apps Script, tempel kode ini.
 * 2. Ganti WRITE_TOKEN dengan kata sandi buatan sendiri (bebas, panjang lebih baik).
 * 3. Deploy > New deployment > Web app
 *      Execute as : Me
 *      Who has access : Anyone
 * 4. Salin URL yang berakhiran /exec ke SYNC_URL di index.html. Token TIDAK ditaruh di HTML; admin memasukkannya lewat tombol "Masuk admin".
 *
 * Membaca (GET) terbuka untuk semua yang punya URL.
 * Menulis / menghapus (POST) wajib token.
 */
const WRITE_TOKEN = 'GANTI_DENGAN_TOKEN_RAHASIA';
const SHEET_NAME = 'laporan';
const HEAD = ['id', 'type', 'date', 'start', 'end', 'saved', 'json'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() === 0) sh.appendRow(HEAD);
  return sh;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function readAll_() {
  const sh = sheet_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  const rows = sh.getRange(2, 1, n - 1, HEAD.length).getValues();
  const list = [];
  rows.forEach(r => {
    try { if (r[6]) list.push(JSON.parse(r[6])); } catch (e) {}
  });
  return list;
}

function doGet() {
  return out_({ ok: true, list: readAll_(), t: Date.now() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'JSON tidak valid' }); }
  if (!body || body.token !== WRITE_TOKEN) return out_({ ok: false, error: 'Token salah' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const idx = {};
    const n = sh.getLastRow();
    if (n > 1) sh.getRange(2, 1, n - 1, 1).getValues().forEach((r, i) => idx[r[0]] = i + 2);

    const upsert = r => {
      if (!r || !r.id) return;
      const row = [r.id, r.type || '', r.date || '', r.start || '', r.end || '', r.saved || Date.now(), JSON.stringify(r)];
      if (idx[r.id]) sh.getRange(idx[r.id], 1, 1, HEAD.length).setValues([row]);
      else { sh.appendRow(row); idx[r.id] = sh.getLastRow(); }
    };

    if (body.action === 'upsert') upsert(body.record);
    else if (body.action === 'bulk') (body.records || []).forEach(upsert);
    else if (body.action === 'delete') {
      if (idx[body.id]) sh.deleteRow(idx[body.id]);
    } else return out_({ ok: false, error: 'Aksi tidak dikenal' });

    SpreadsheetApp.flush();
    return out_({ ok: true, list: readAll_() });
  } finally {
    lock.releaseLock();
  }
}
