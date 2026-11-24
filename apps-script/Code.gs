/**
 * Backend Google Apps Script untuk Dashboard Monitoring Progress Project SDO (versi 2).
 *
 * Perbedaan dengan versi lama:
 *  - Web TIDAK lagi mengirim seluruh daftar project ("saveAll"). Web hanya mengirim
 *    kolom yang diubah untuk SATU project ("upsert") atau perintah hapus ("delete").
 *  - Setiap penulisan dikunci dengan LockService, jadi dua orang yang menyimpan
 *    bersamaan tidak saling menimpa.
 *  - Semua sel disimpan sebagai teks biasa, supaya "12/26" tidak berubah jadi tanggal
 *    dan "100%" tidak berubah jadi 1.
 *
 * CARA PASANG (URL web TIDAK berubah):
 *  1. Buka Google Sheets SDO -> Extensions -> Apps Script.
 *  2. BACKUP dulu: salin isi Code.gs lama ke tempat aman (dan buat salinan Sheet: File -> Make a copy).
 *  3. Ganti seluruh isi Code.gs dengan file ini, klik Save.
 *  4. Jika data project tidak ada di sheet pertama, isi SHEET_NAME di bawah dengan nama tab-nya.
 *  5. Deploy -> Manage deployments -> pilih deployment yang sekarang -> ikon pensil (Edit)
 *     -> Version: "New version" -> Deploy.
 *     (JANGAN pilih "New deployment", karena akan menghasilkan URL baru.)
 *  6. Buka web dashboard, tekan Ctrl+F5. Selesai.
 *
 * Syarat: baris 1 sheet berisi nama kolom persis seperti FIELDS di bawah (id, panel, name, ...).
 * Kolom yang belum ada akan ditambahkan otomatis di sebelah kanan.
 */

const SHEET_NAME = ""; // kosong = sheet pertama

const FIELDS = [
  "id", "panel", "name", "status", "score", "waBadge",
  "blueText", "blueStart", "blueEnd",
  "orangeText", "orangeStart", "orangeEnd",
  "greyText", "greyStart", "greyEnd",
  "redText", "redStart", "redEnd",
  "updatedBy", "updatedAt"
];

const API_VERSION = 2;

function doGet() {
  try {
    const sh = getSheet_();
    const headers = getHeaders_(sh, false);
    return json_({ status: "ok", version: API_VERSION, projects: readProjects_(sh, headers) });
  } catch (err) {
    return json_({ status: "error", version: API_VERSION, message: String(err && err.message || err), projects: [] });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    lock.waitLock(30000);

    const sh = getSheet_();
    const headers = getHeaders_(sh, true);

    if (body.action === "upsert") {
      upsert_(sh, headers, body);
    } else if (body.action === "delete") {
      delete_(sh, headers, body);
    } else if (body.action === "saveAll") {
      // Ditolak dengan sengaja: saveAll dari tab browser lama akan menimpa perubahan orang lain.
      throw new Error("Versi web lama terdeteksi. Silakan refresh halaman (Ctrl+F5) lalu ulangi perubahan.");
    } else {
      throw new Error("Aksi tidak dikenal: " + body.action);
    }

    SpreadsheetApp.flush();
    return json_({ status: "ok", version: API_VERSION, projects: readProjects_(sh, headers) });
  } catch (err) {
    return json_({ status: "error", version: API_VERSION, message: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// ---------------------------------------------------------------------------

function upsert_(sh, headers, body) {
  const id = String(body.id || "").trim();
  if (!id) throw new Error("id project kosong");
  const changes = body.changes || {};
  const rowNum = findRow_(sh, headers, id);

  if (rowNum > 0) {
    // Update hanya kolom yang dikirim
    const range = sh.getRange(rowNum, 1, 1, headers.length);
    const row = range.getDisplayValues()[0];
    Object.keys(changes).forEach(function (k) {
      if (k === "id") return;
      const col = headers.indexOf(k);
      if (col >= 0) row[col] = toText_(changes[k]);
    });
    range.setNumberFormat("@").setValues([row]);
  } else {
    // Project baru
    const full = Object.assign({}, body.project || {}, changes, { id: id });
    const row = headers.map(function (h) { return h ? toText_(full[h]) : ""; });
    const target = sh.getRange(sh.getLastRow() + 1, 1, 1, headers.length);
    target.setNumberFormat("@").setValues([row]);
  }
}

function delete_(sh, headers, body) {
  const id = String(body.id || "").trim();
  if (!id) throw new Error("id project kosong");
  const rowNum = findRow_(sh, headers, id);
  if (rowNum > 0) sh.deleteRow(rowNum);
}

function findRow_(sh, headers, id) {
  const idCol = headers.indexOf("id") + 1;
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, idCol, n, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) return i + 2;
  }
  return -1;
}

function readProjects_(sh, headers) {
  const n = sh.getLastRow() - 1;
  if (n < 1 || headers.indexOf("id") < 0) return [];
  const values = sh.getRange(2, 1, n, headers.length).getDisplayValues();
  const list = [];
  values.forEach(function (row) {
    const p = {};
    headers.forEach(function (h, j) { if (h) p[h] = row[j]; });
    if (String(p.id || "").trim()) list.push(p);
  });
  return list;
}

function getHeaders_(sh, createMissing) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  let headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) { return String(h).trim(); });
  if (headers.every(function (h) { return !h; })) headers = [];

  if (headers.length && headers.indexOf("id") < 0) {
    throw new Error("Baris 1 sheet '" + sh.getName() + "' tidak punya kolom 'id'. Periksa SHEET_NAME / nama kolom di Code.gs.");
  }

  if (createMissing) {
    const missing = FIELDS.filter(function (f) { return headers.indexOf(f) < 0; });
    if (missing.length) {
      sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
      headers = headers.concat(missing);
    }
  }
  return headers;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
  if (!sh) throw new Error("Sheet '" + SHEET_NAME + "' tidak ditemukan");
  return sh;
}

function toText_(v) {
  return v == null ? "" : String(v);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
