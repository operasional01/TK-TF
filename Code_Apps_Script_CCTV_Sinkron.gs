function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '');
  let result;
  try {
    verifyToken_(p.token);
    if (p.action !== 'loadAll') throw new Error('Aksi tidak dikenal.');
    result = { ok: true, payload: loadAll_() };
  } catch (err) {
    result = { ok: false, error: String(err && err.message || err) };
  }
  const json = JSON.stringify(result).replace(/</g, '\\u003c');
  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService.createTextOutput('Invalid callback').setMimeType(ContentService.MimeType.TEXT);
    }
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const body = e && e.postData && e.postData.contents;
    if (!body) throw new Error('Isi permintaan kosong.');
    const req = JSON.parse(body);
    verifyToken_(req.token);
    if (req.action !== 'saveAll') throw new Error('Aksi tidak dikenal.');
    saveAll_(req.payload || {});
    result = { ok: true, message: 'Data tersimpan.' };
  } catch (err) {
    result = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyToken_(token) {
  if (!ACCESS_TOKEN || ACCESS_TOKEN.indexOf('GANTI_DENGAN_') === 0) {
    throw new Error('ACCESS_TOKEN belum dikonfigurasi pada Apps Script.');
  }
  if (String(token || '') !== ACCESS_TOKEN) throw new Error('Token akses tidak cocok.');
}

function spreadsheet_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf('GANTI_DENGAN_') === 0) {
    throw new Error('SPREADSHEET_ID belum dikonfigurasi pada Apps Script.');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function saveAll_(payload) {
  const ss = spreadsheet_();
  writeRecords_(ss, 'TKTF_Utama', Array.isArray(payload.main) ? payload.main : []);
  writeRecords_(ss, 'Titik_Rawan', Array.isArray(payload.hotspots) ? payload.hotspots : []);
  writeRecords_(ss, 'Analisis_Volume', Array.isArray(payload.volume) ? payload.volume : []);
  // Data CCTV disimpan terpisah agar dapat dimuat lintas perangkat.
  writeRecords_(ss, 'CCTV_Lokasi', Array.isArray(payload.cctv) ? payload.cctv : []);
  writeRecords_(ss, 'Pengaturan', [payload.settings && typeof payload.settings === 'object' ? payload.settings : {}]);
  SpreadsheetApp.flush();
}

function loadAll_() {
  const ss = spreadsheet_();
  const settingsRows = readRecords_(ss, 'Pengaturan');
  return {
    main: readRecords_(ss, 'TKTF_Utama'),
    hotspots: readRecords_(ss, 'Titik_Rawan'),
    volume: readRecords_(ss, 'Analisis_Volume'),
    // Mengembalikan data CCTV ke aplikasi HTML saat memilih "Muat dari Sheets".
    cctv: readRecords_(ss, 'CCTV_Lokasi'),
    settings: settingsRows[0] || { benchmarkTf: 0.1211 }
  };
}

function writeRecords_(ss, name, records) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, 2).setValues([['No', 'Data_JSON']]);
  if (!records.length) return;
  const rows = records.map((r, i) => [i + 1, JSON.stringify(r)]);
  // A single record is kept per row; Google Sheets cells have a size limit.
  rows.forEach(row => {
    if (String(row[1]).length > 45000) {
      throw new Error('Satu record pada sheet ' + name + ' terlalu besar untuk satu sel.');
    }
  });
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

function readRecords_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  return values.map(row => {
    try { return JSON.parse(row[0]); } catch (err) { return null; }
  }).filter(v => v !== null);
}
