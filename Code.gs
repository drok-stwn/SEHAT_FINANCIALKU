/**
 * Backend Keuangan Keluarga - Hendrik Setiawan & Noor Loppa
 * Tempel di Apps Script yang terhubung ke Google Sheets (Extensions > Apps Script).
 * GANTI PIN di bawah sebelum deploy.
 * Foto struk disimpan di folder Google Drive "Struk Keluarga" dan hanya bisa dibuka lewat PIN.
 */
const USERS = {
  suami: { nama: 'Hendrik Setiawan', pin: '161223' }, // <-- ganti PIN suami
  istri: { nama: 'Noor Loppa',       pin: '161223' }  // <-- ganti PIN istri
};
const FOLDER_STRUK = 'Struk Keluarga';

const SHEETS = {
  pemasukan:   ['ID', 'Tanggal', 'Oleh', 'Keterangan', 'Jumlah'],
  pengeluaran: ['ID', 'Tanggal', 'Oleh', 'Pos', 'Keterangan', 'Jumlah', 'Struk']
};

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
  }
  // tanggal harus teks yyyy-mm-dd (jika tidak, Sheets mengubahnya jadi tanggal dan data tidak terbaca)
  if (sh.getRange('B2').getNumberFormat() !== '@') sh.getRange('B:B').setNumberFormat('@');
  // sheet lama: tambahkan kolom "Struk" jika belum ada
  if (name === 'pengeluaran' && sh.getRange(1, 7).getValue() !== 'Struk') {
    sh.getRange(1, 7).setValue('Struk');
  }
  return sh;
}

/** JALANKAN SEKALI dari editor Apps Script (pilih fungsi "authorize" lalu Run) untuk memberi izin Drive & Sheets. */
function authorize() {
  SpreadsheetApp.getActiveSpreadsheet().getName();
  const f = folder_();
  Logger.log('OK. Folder struk: ' + f.getUrl());
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_STRUK);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_STRUK);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function auth_(role, pin) {
  const u = USERS[role];
  return !!u && String(pin) === String(u.pin) ? u : null;
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const t = String(v).trim();
  const m = t.match(/^(\d{4})-(\d{2})-(\d{2})/);            // 2026-10-13
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  const d = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);        // 13/10/2026
  if (d) return d[3] + '-' + ('0' + d[2]).slice(-2) + '-' + ('0' + d[1]).slice(-2);
  return t.slice(0, 10);
}

function readAll_(name) {
  const sh = getSheet_(name);
  const values = sh.getDataRange().getValues();
  const head = values.shift();
  return values.filter(r => r[0] !== '').map(r => {
    const o = {};
    head.forEach((h, i) => o[h] = r[i]);
    o.Jumlah = Number(o.Jumlah) || 0;
    o.Tanggal = dateStr_(o.Tanggal);
    if (name === 'pengeluaran') o.Struk = o.Struk ? String(o.Struk) : '';
    return o;
  });
}

function doGet(e) {
  const p = e.parameter || {};
  const user = auth_(p.role, p.pin);
  if (!user) return json_({ ok: false, error: 'PIN salah' });

  // ambil foto struk (hanya file di folder struk)
  if (p.action === 'foto') {
    try {
      if (!p.id) return json_({ ok: false, error: 'ID foto kosong' });
      const file = DriveApp.getFileById(p.id);
      const folderId = folder_().getId();
      let inside = false;
      const parents = file.getParents();
      while (parents.hasNext()) { if (parents.next().getId() === folderId) { inside = true; break; } }
      if (!inside && String(file.getName()).indexOf('struk-') !== 0) {
        return json_({ ok: false, error: 'Bukan file struk' });
      }
      const blob = file.getBlob();
      return json_({ ok: true, data: 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes()) });
    } catch (err) {
      return json_({ ok: false, error: 'Gagal membuka foto: ' + err });
    }
  }

  return json_({
    ok: true,
    nama: user.nama,
    pemasukan: readAll_('pemasukan'),
    pengeluaran: readAll_('pengeluaran')
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    const user = auth_(b.role, b.pin);
    if (!user) return json_({ ok: false, error: 'PIN salah' });
    if (!SHEETS[b.sheet]) return json_({ ok: false, error: 'Sheet tidak dikenal' });

    const sh = getSheet_(b.sheet);

    if (b.action === 'add') {
      const jumlah = Number(b.jumlah);
      if (!(jumlah > 0)) return json_({ ok: false, error: 'Jumlah tidak valid' });
      const id = Utilities.getUuid().slice(0, 8);

      let struk = '', fotoError = '';
      if (b.sheet === 'pengeluaran' && b.foto) {
        try {
          const str = String(b.foto);
          const comma = str.indexOf(',');
          const mime = (str.substring(5, str.indexOf(';')) || 'image/jpeg');
          if (comma < 0 || str.indexOf('data:image/') !== 0) throw new Error('format foto tidak dikenal');
          const blob = Utilities.newBlob(Utilities.base64Decode(str.substring(comma + 1)), mime, 'struk-' + b.tanggal + '-' + id + '.jpg');
          struk = folder_().createFile(blob).getId();
        } catch (err) {
          fotoError = String(err);
        }
      }

      const row = b.sheet === 'pemasukan'
        ? [id, b.tanggal, user.nama, b.keterangan || '', jumlah]
        : [id, b.tanggal, user.nama, b.pos, b.keterangan || '', jumlah, struk];
      sh.appendRow(row);
      SpreadsheetApp.flush();
      return json_({ ok: true, id: id, struk: struk, fotoError: fotoError, ver: 3 });
    }

    if (b.action === 'delete') {
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat();
      const idx = ids.indexOf(b.id);
      if (idx < 1) return json_({ ok: false, error: 'Data tidak ditemukan' });
      if (b.sheet === 'pengeluaran') {
        const strukId = sh.getRange(idx + 1, 7).getValue();
        if (strukId) { try { DriveApp.getFileById(String(strukId)).setTrashed(true); } catch (err) {} }
      }
      sh.deleteRow(idx + 1);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'Aksi tidak dikenal' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
