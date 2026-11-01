/**
 * SISTEM KUIZ TINGKATAN 2 — Pangkalan Data Google Sheets + Google Drive
 * -------------------------------------------------------------------
 * Helaian yang dijana secara automatik:
 *   MURID           – profil murid, gambar, skor keseluruhan, bab paling lemah
 *   ANALISIS_BAB    – peratus ketepatan setiap murid bagi setiap bab (berwarna)
 *   RINGKASAN_KELAS – purata ketepatan kelas bagi setiap bab
 *   KEPUTUSAN_MODUL – rekod setiap kali murid tamat satu modul
 *   LOG_JAWAPAN     – setiap jawapan murid (secara live)
 *   _DATA           – data simpanan untuk sambung semula (jangan ubah)
 */
const SHEET_ID  = '1YojfjCGztOY-iZoRkHxfRDRcxfSz__jp6JFdaOmp86k';
const FOLDER_ID = '1lOkUBvaQzSN-iEzhBTa24RxeFk7xSe4-';
const TZ = 'Asia/Kuala_Lumpur';

const H_MURID = ['Nama Penuh','Kelas','Gambar','Pautan Gambar','Tarikh Daftar','Log Masuk Terakhir','Bil. Log Masuk',
  'Soalan Dijawab','Skor Keseluruhan','Peratus (%)','Gred','Tahap','Modul Selesai','Bab Paling Lemah','Bab Paling Kuat','Kemas Kini Terakhir','Mata Ranking','Jumlah Masa Jawab Betul (s)','Soalan Betul (Mata)'];
const H_KEP = ['Masa','Nama Penuh','Kelas','Bab / Modul','Skor (Cubaan Pertama)','Jumlah Soalan','Peratus (%)'];
const H_LOG = ['Masa','Nama Penuh','Kelas','Bab / Modul','Soalan','Jawapan Murid','Jawapan Betul','Keputusan','Cubaan'];
const H_DATA = ['Nama Penuh','Data (JSON)','Kemas Kini'];

/* ================= PINTU MASUK ================= */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistem Kuiz Tingkatan 2')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    return out_({ ok: true, data: api(b.action, b.payload) });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Dipanggil oleh google.script.run atau doPost.
 * PRESTASI (1000+ murid serentak):
 *  - 'save', 'log', 'result' TIDAK menulis terus ke Sheets. Data disimpan dalam baris gilir (CacheService)
 *    dalam ~50ms tanpa kunci global, kemudian flushData() (pencetus setiap 1 minit) menulis semuanya sekali gus.
 *  - 'rank' dibaca dari cache sahaja (dikira semula semasa flush).
 *  - Hanya 'login' (sekali bagi setiap murid) menggunakan kunci. */
function api(action, p) {
  switch (action) {
    case 'rank':   return rank_();
    case 'save':   return qSave_(p);
    case 'log':    return qPush_('L', p.nama, (p.rows || []).map(r => [r.masa || now_(), norm_(p.nama), norm_(p.kelas), r.bab, r.soalan, r.jawapan, r.betul, r.keputusan, r.cubaan]));
    case 'result': return qPush_('R', p.nama, [[now_(), norm_(p.nama), norm_(p.kelas), p.bab, p.skor, p.total, p.total ? Math.round(p.skor / p.total * 100) : 0]]);
    case 'login': {
      const foto = savePhoto_(norm_(p.nama), norm_(p.kelas), p.foto); // di luar kunci (perlahan)
      const lock = LockService.getScriptLock();
      lock.waitLock(28000);
      try { return login_(p, foto); } finally { lock.releaseLock(); }
    }
    default: throw new Error('Tindakan tidak sah: ' + action);
  }
}

/* ================= BARIS GILIR (CACHE) ================= */
const C_ = () => CacheService.getScriptCache();
const QTTL = 21600; // 6 jam
function names_() {
  const hit = C_().get('NAMES');
  if (hit) return JSON.parse(hit);
  const sh = sheet_('MURID', H_MURID), last = sh.getLastRow();
  const list = last < 2 ? [] : sh.getRange(2, 1, last - 1, 1).getValues().map(r => norm_(r[0])).filter(String);
  C_().put('NAMES', JSON.stringify(list), QTTL);
  return list;
}
function qSave_(p) {
  const nama = norm_(p.nama);
  const s = JSON.stringify({ t: Date.now(), state: p.state || {} });
  if (s.length < 95000) { C_().put('S_' + nama, s, QTTL); return { ok: true, q: true }; }
  // Terlalu besar untuk cache: tulis terus (jarang berlaku)
  const lock = LockService.getScriptLock(); lock.waitLock(28000);
  try { return save_(p); } finally { lock.releaseLock(); }
}
function qPush_(kind, nama, rows) {
  if (!rows.length) return { ok: true };
  // Kunci unik setiap permintaan: tiada perlumbaan antara murid
  const key = kind + '_' + norm_(nama) + '_' + Date.now() + '_' + Math.floor(Math.random() * 1e6);
  C_().put(key, JSON.stringify(rows), QTTL);
  const ik = 'IDX_' + kind + '_' + norm_(nama);
  const idx = JSON.parse(C_().get(ik) || '[]'); idx.push(key);
  C_().put(ik, JSON.stringify(idx), QTTL);
  return { ok: true, q: true };
}

/** Pencetus masa (setiap 1 minit): tulis semua data beratur ke Sheets secara pukal. */
function flushData() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // flush lain sedang berjalan
  try {
    const cache = C_(), names = names_();
    if (!names.length) return;
    // 1) Kemajuan murid
    const sKeys = names.map(n => 'S_' + n);
    const pending = {};
    for (let i = 0; i < sKeys.length; i += 500) Object.assign(pending, cache.getAll(sKeys.slice(i, i + 500)));
    const done = [];
    Object.keys(pending).forEach(k => {
      try { save_({ nama: k.slice(2), state: JSON.parse(pending[k]).state }, true); done.push(k); } catch (e) { console.error(k, e); }
    });
    if (done.length) { bulkWrite_(); cache.removeAll(done); }
    // 2) Log jawapan + keputusan modul
    [['L', 'LOG_JAWAPAN', H_LOG], ['R', 'KEPUTUSAN_MODUL', H_KEP]].forEach(([kind, sheetName, head]) => {
      const iks = names.map(n => 'IDX_' + kind + '_' + n);
      const idx = {};
      for (let i = 0; i < iks.length; i += 500) Object.assign(idx, cache.getAll(iks.slice(i, i + 500)));
      const keys = [].concat.apply([], Object.keys(idx).map(k => JSON.parse(idx[k])));
      if (!keys.length) return;
      cache.removeAll(Object.keys(idx));
      const vals = {};
      for (let i = 0; i < keys.length; i += 500) Object.assign(vals, cache.getAll(keys.slice(i, i + 500)));
      const rows = [].concat.apply([], Object.keys(vals).map(k => JSON.parse(vals[k])));
      if (rows.length) { const sh = sheet_(sheetName, head); sh.getRange(sh.getLastRow() + 1, 1, rows.length, head.length).setValues(rows); }
      cache.removeAll(keys);
    });
    rankBuild_();
  } finally { lock.releaseLock(); }
}

/* Penulisan pukal: save_() mengumpul perubahan dalam BUF_, bulkWrite_() menulis setiap helaian sekali sahaja */
let BUF_ = null;
function buf_() {
  if (BUF_) return BUF_;
  const read = (name, head, w) => { const sh = sheet_(name, head); const last = sh.getLastRow(); const width = w || Math.max(sh.getLastColumn(), head.length);
    return { sh, width, rows: last < 2 ? [] : sh.getRange(2, 1, last - 1, width).getValues(), n: Math.max(last - 1, 0), add: [], dirty: false }; };
  BUF_ = { M: read('MURID', H_MURID, H_MURID.length), D: read('_DATA', H_DATA, 3) };
  const saved = JSON.parse(PropertiesService.getScriptProperties().getProperty('MODULES') || '[]');
  if (saved.length) BUF_.A = read('ANALISIS_BAB', [], saved.length + 5);
  return BUF_;
}
function bufRow_(b, nama) { for (let i = 0; i < b.rows.length; i++) if (norm_(b.rows[i][0]) === nama) return i; return -1; }
function bulkWrite_() {
  if (!BUF_) return;
  ['M', 'D', 'A'].forEach(k => { const b = BUF_[k]; if (!b) return;
    if (b.dirty && b.n) b.sh.getRange(2, 1, b.n, b.width).setValues(b.rows.slice(0, b.n));
    if (b.add.length) b.sh.getRange(b.sh.getLastRow() + 1, 1, b.add.length, b.width).setValues(b.add.map(r => r.concat(Array(Math.max(0, b.width - r.length)).fill('')).slice(0, b.width)));
  });
  BUF_ = null;
}

/** Jalankan SEKALI secara manual: pasang pencetus flush setiap 1 minit. */
function pasangPencetus() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'flushData').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('flushData').timeBased().everyMinutes(1).create();
}

/* ================= UTILITI ================= */
const ss_ = () => SpreadsheetApp.openById(SHEET_ID);
const now_ = () => Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
function norm_(s) { return String(s || '').toUpperCase().replace(/\s+/g, ' ').trim(); }

function validName_(n) {
  if (n.length < 6 || n.length > 80) return false;
  if (!/^[A-Z@'.\-\/ ]+$/.test(n)) return false;
  return n.split(' ').filter(w => w.length >= 2).length >= 2;
}

function sheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#0b1f3a').setFontColor('#ffffff').setWrap(true);
    sh.setFrozenRows(1);
    if (name === '_DATA') sh.hideSheet();
  }
  return sh;
}

function findRow_(sh, nama) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (norm_(vals[i][0]) === nama) return i + 2;
  return -1;
}

function folderFor_(kelas) {
  const root = DriveApp.getFolderById(FOLDER_ID);
  const name = 'KELAS ' + kelas;
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function savePhoto_(nama, kelas, dataUrl) {
  if (!dataUrl || dataUrl.indexOf('base64,') < 0) return null;
  const bytes = Utilities.base64Decode(dataUrl.split('base64,')[1]);
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
  const blob = Utilities.newBlob(bytes, 'image/jpeg', nama + ' - ' + kelas + ' - ' + stamp + '.jpg');
  const file = folderFor_(kelas).createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { id: file.getId(), url: file.getUrl(), thumb: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400' };
}

/* ================= SUSUNAN BAB ================= */
function ensureAnalisis_(modules) {
  const head = ['Nama Penuh', 'Kelas'].concat(modules.map(m => m.nama + ' (%)'), ['Bab Paling Lemah', 'Bab Paling Kuat', 'Kemas Kini']);
  const sh = sheet_('ANALISIS_BAB', head);
  if (sh.getRange(1, 1, 1, head.length).getValues()[0].join('|') !== head.join('|')) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#0b1f3a').setFontColor('#ffffff').setWrap(true);
  }
  // Warna: merah (lemah) → kuning → hijau (kuat)
  const rng = sh.getRange(2, 3, Math.max(sh.getMaxRows() - 1, 1), modules.length);
  const rule = SpreadsheetApp.newConditionalFormatRule().setGradientMinpointWithValue('#f4a6a6', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue('#ffe599', SpreadsheetApp.InterpolationType.NUMBER, '60')
    .setGradientMaxpointWithValue('#9fe3b0', SpreadsheetApp.InterpolationType.NUMBER, '100').setRanges([rng]).build();
  sh.setConditionalFormatRules([rule]);

  // Ringkasan kelas (formula live)
  const rk = sheet_('RINGKASAN_KELAS', ['Bab / Modul', 'Purata Ketepatan (%)', 'Bil. Murid Mencuba', 'Bil. Murid < 50%']);
  const rows = modules.map((m, i) => {
    const col = colLetter_(3 + i);
    const r = `ANALISIS_BAB!${col}2:${col}`;
    return [m.nama, `=IFERROR(ROUND(AVERAGE(${r}),1),"-")`, `=COUNT(${r})`, `=COUNTIF(${r},"<50")`];
  });
  rk.getRange(2, 1, rows.length, 4).setValues(rows);
  PropertiesService.getScriptProperties().setProperty('MODULES', JSON.stringify(modules));
  return sh;
}
function colLetter_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

/* ================= LOG MASUK / DAFTAR ================= */
function login_(p, foto) {
  const nama = norm_(p.nama), kelas = norm_(p.kelas);
  if (!validName_(nama)) throw new Error('Sila masukkan NAMA PENUH yang sebenar (sekurang-kurangnya 2 perkataan, huruf sahaja).');
  if (!kelas) throw new Error('Sila masukkan kelas.');
  if (p.modules && p.modules.length && PropertiesService.getScriptProperties().getProperty('MODULES') !== JSON.stringify(p.modules)) ensureAnalisis_(p.modules);

  const sh = sheet_('MURID', H_MURID);
  let row = findRow_(sh, nama);
  const baru = row < 0;

  if (baru) {
    sh.appendRow([nama, kelas, '', foto ? foto.url : '', now_(), now_(), 1, 0, 0, 0, '-', '-', 0, '-', '-', now_(), 0, 0, 0]);
    row = sh.getLastRow();
    const nl = names_(); if (nl.indexOf(nama) < 0) { nl.push(nama); C_().put('NAMES', JSON.stringify(nl), QTTL); }
  } else {
    sh.getRange(row, 2).setValue(kelas);
    sh.getRange(row, 6, 1, 2).setValues([[now_(), (Number(sh.getRange(row, 7).getValue()) || 0) + 1]]);
    if (foto) sh.getRange(row, 4).setValue(foto.url);
  }
  if (foto) sh.getRange(row, 3).setFormula('=IMAGE("' + foto.thumb + '")');
  sh.setRowHeight(row, 70);

  const d = sheet_('_DATA', H_DATA);
  const dr = findRow_(d, nama);
  let state = null;
  if (dr > 0) { try { state = JSON.parse(d.getRange(dr, 2).getValue()); } catch (e) {} }
  const q = C_().get('S_' + nama); // kemajuan terbaru yang belum di-flush
  if (q) { try { state = JSON.parse(q).state.save || state; } catch (e) {} }
  return { baru: baru, nama: nama, kelas: kelas, state: state, fotoUrl: foto ? foto.thumb : null };
}

/* ================= SIMPAN KEMAJUAN =================
 * bulk=true (dari flushData): perubahan dikumpul dalam ingatan, ditulis sekali gus oleh bulkWrite_(). */
function save_(p, bulk) {
  const nama = norm_(p.nama);
  const B = buf_();
  const mi = bufRow_(B.M, nama);
  if (mi < 0) throw new Error('Murid belum berdaftar.');
  const M = B.M.rows[mi]; B.M.dirty = true;
  const kelas = M[1];
  const st = p.state || {};

  // 1) Data sambung semula
  const di = bufRow_(B.D, nama), dv = [nama, JSON.stringify(st.save || {}), now_()];
  if (di >= 0) { B.D.rows[di] = dv; B.D.dirty = true; } else { B.D.add.push(dv); B.D.rows.push(dv); }

  // 2) Ringkasan murid
  const s = st.summary || {};
  const mods = st.modStats || [];
  const tried = mods.filter(m => m.acc !== null && m.acc !== undefined);
  const weak = tried.length ? tried.reduce((a, b) => (b.acc < a.acc ? b : a)) : null;
  const strong = tried.length ? tried.reduce((a, b) => (b.acc > a.acc ? b : a)) : null;
  const save = st.save || {};
  const betul = Object.keys(save.scored || {}).reduce((n, k) => n + (save.scored[k] || []).length, 0);
  [s.answered || 0, s.skor || 0, s.pct || 0, s.gred || '-', s.tahap || '-', (s.done || 0) + '/' + (s.modTotal || mods.length),
    weak ? weak.nama + ' (' + weak.acc + '%)' : '-', strong ? strong.nama + ' (' + strong.acc + '%)' : '-', now_(),
    Number(s.mata) || 0, Number(s.masaJawab) || 0, betul].forEach((v, k) => M[7 + k] = v);

  // 3) Analisis bab
  if (mods.length) {
    if (!B.A) { ensureAnalisis_(mods.map(m => ({ id: m.id, nama: m.nama, total: m.total }))); BUF_ = null; return save_(p, bulk); }
    const saved = JSON.parse(PropertiesService.getScriptProperties().getProperty('MODULES') || '[]');
    const order = saved.map(m => m.id);
    const byId = {}; mods.forEach(m => byId[m.id] = m);
    const line = [nama, kelas].concat(order.map(id => (byId[id] && byId[id].acc !== null && byId[id].acc !== undefined) ? byId[id].acc : ''),
      [weak ? weak.nama : '-', strong ? strong.nama : '-', now_()]);
    const ai = bufRow_(B.A, nama);
    if (ai >= 0) { B.A.rows[ai] = line; B.A.dirty = true; } else B.A.add.push(line);
  }
  if (!bulk) bulkWrite_();
  return { ok: true };
}

/* ================= RANKING LIVE =================
 * Mata setiap soalan = 40 saat − masa menjawab betul (ketepatan milisaat).
 * Kedudukan: lebih banyak soalan dijawab betul = lebih tinggi (dikira LIVE walaupun belum habis).
 * Jika sama, mata kelajuan menentukan. Seri dipecahkan oleh jumlah masa terpendek. */
function rank_() {
  const hit = C_().get('RANK');
  if (hit) return JSON.parse(hit);
  const lock = LockService.getScriptLock(); // hanya seorang mengira semula jika cache kosong
  if (!lock.tryLock(8000)) return [];
  try { const again = C_().get('RANK'); return again ? JSON.parse(again) : rankBuild_(); } finally { lock.releaseLock(); }
}
function rankBuild_() {
  const sh = sheet_('MURID', H_MURID);
  if (sh.getRange(1, 17).getValue() !== H_MURID[16]) sh.getRange(1, 17, 1, 3).setValues([H_MURID.slice(16, 19)]).setFontWeight('bold').setBackground('#0b1f3a').setFontColor('#ffffff');
  const last = sh.getLastRow();
  const list = last < 2 ? [] : sh.getRange(2, 1, last - 1, 19).getValues()
    .filter(r => r[0] && (Number(r[18]) > 0 || Number(r[16]) > 0))
    .map(r => ({ nama: String(r[0]), kelas: String(r[1]), mata: Number(r[16]) || 0, masa: Number(r[17]) || 0, betul: Number(r[18]) || 0 }))
    .sort((a, b) => b.betul - a.betul || b.mata - a.mata || a.masa - b.masa || a.nama.localeCompare(b.nama))
    .slice(0, 50);
  C_().put('RANK', JSON.stringify(list), 120);
  return list;
}

/** Jalankan sekali secara manual untuk memberi kebenaran (Sheets + Drive). */
function setup() {
  sheet_('MURID', H_MURID); sheet_('KEPUTUSAN_MODUL', H_KEP); pasangPencetus(); sheet_('LOG_JAWAPAN', H_LOG); sheet_('_DATA', H_DATA);
  DriveApp.getFolderById(FOLDER_ID).getName();
  const s1 = ss_().getSheetByName('Sheet1') || ss_().getSheetByName('Helaian1');
  if (s1 && s1.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(s1);
}
