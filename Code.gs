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
  let html = HtmlService.createHtmlOutputFromFile('index').getContent();
  const raw = soalan_();
  if (raw.length) {
    const json = JSON.stringify(raw).replace(/</g, '\\u003c');
    html = html.replace(/\/\*SOALAN_MULA\*\/[\s\S]*?\/\*SOALAN_TAMAT\*\//, function(){ return 'const RAW=' + json + ';'; });
  }
  return HtmlService.createHtmlOutput(html)
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

/** Dipanggil oleh google.script.run atau doPost */
function api(action, p) {
  if (action === 'rank') return rank_();
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    switch (action) {
      case 'login':  return login_(p);
      case 'save':   return save_(p);
      case 'log':    return log_(p);
      case 'result': return result_(p);
      case 'rank':   return rank_();
      default: throw new Error('Tindakan tidak sah: ' + action);
    }
  } finally { lock.releaseLock(); }
}

/* ================= SOALAN (dari helaian SOALAN) =================
 * Lajur: ID Modul | Nama Modul | Keterangan Modul | Konteks / Rajah | Soalan |
 *        Jawapan Betul | Pilihan Salah 1 | Pilihan Salah 2 | Pilihan Salah 3 | Petunjuk | Nota Penerangan
 * Soalan kini DIAMBIL DARIPADA HELAIAN "SOALAN" SAHAJA (tiada soalan dalam index.html).
 */
function soalan_() {
  const sh = ss_().getSheetByName('SOALAN');
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getDisplayValues();
  const mods = [], map = {};
  rows.forEach(function (r) {
    const id = String(r[0]).trim(), q = String(r[4]).trim(), a = String(r[5]).trim();
    if (!id || !q || !a) return;
    if (!map[id]) { map[id] = { id: id, nama: String(r[1]).trim(), desc: String(r[2]).trim(), q: [] }; mods.push(map[id]); }
    const w = [r[6], r[7], r[8]].map(function (x) { return String(x).trim(); }).filter(String);
    map[id].q.push([q, a, w, String(r[9]), String(r[10]), String(r[3]).trim()]);
  });
  return mods;
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
function login_(p) {
  const nama = norm_(p.nama), kelas = norm_(p.kelas);
  if (!validName_(nama)) throw new Error('Sila masukkan NAMA PENUH yang sebenar (sekurang-kurangnya 2 perkataan, huruf sahaja).');
  if (!kelas) throw new Error('Sila masukkan kelas.');
  if (p.modules && p.modules.length) ensureAnalisis_(p.modules);

  const sh = sheet_('MURID', H_MURID);
  let row = findRow_(sh, nama);
  const baru = row < 0;
  const foto = savePhoto_(nama, kelas, p.foto);

  if (baru) {
    sh.appendRow([nama, kelas, '', foto ? foto.url : '', now_(), now_(), 1, 0, 0, 0, '-', '-', 0, '-', '-', now_(), 0, 0, 0]);
    row = sh.getLastRow();
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
  return { baru: baru, nama: nama, kelas: kelas, state: state, fotoUrl: foto ? foto.thumb : null };
}

/* ================= SIMPAN KEMAJUAN (LIVE) ================= */
function save_(p) {
  const nama = norm_(p.nama);
  const sh = sheet_('MURID', H_MURID);
  const row = findRow_(sh, nama);
  if (row < 0) throw new Error('Murid belum berdaftar.');
  const kelas = sh.getRange(row, 2).getValue();
  const st = p.state || {};

  // 1) Data sambung semula
  const d = sheet_('_DATA', H_DATA);
  const dr = findRow_(d, nama);
  const vals = [[nama, JSON.stringify(st.save || {}), now_()]];
  if (dr > 0) d.getRange(dr, 1, 1, 3).setValues(vals); else d.appendRow(vals[0]);

  // 2) Ringkasan murid
  const s = st.summary || {};
  const mods = st.modStats || [];
  const tried = mods.filter(m => m.acc !== null && m.acc !== undefined);
  const weak = tried.length ? tried.reduce((a, b) => (b.acc < a.acc ? b : a)) : null;
  const strong = tried.length ? tried.reduce((a, b) => (b.acc > a.acc ? b : a)) : null;
  sh.getRange(row, 8, 1, 9).setValues([[
    s.answered || 0, s.skor || 0, s.pct || 0, s.gred || '-', s.tahap || '-', (s.done || 0) + '/' + (s.modTotal || mods.length),
    weak ? weak.nama + ' (' + weak.acc + '%)' : '-', strong ? strong.nama + ' (' + strong.acc + '%)' : '-', now_()
  ]]);

  const save = st.save || {};
  const betul = Object.keys(save.scored || {}).reduce((n, k) => n + (save.scored[k] || []).length, 0);
  sh.getRange(row, 17, 1, 3).setValues([[Number(s.mata) || 0, Number(s.masaJawab) || 0, betul]]);
  CacheService.getScriptCache().remove('RANK');

  // 3) Analisis bab
  if (mods.length) {
    const saved = JSON.parse(PropertiesService.getScriptProperties().getProperty('MODULES') || '[]');
    const a = saved.length ? sheet_('ANALISIS_BAB', []) : ensureAnalisis_(mods.map(m => ({ id: m.id, nama: m.nama, total: m.total })));
    const order = (saved.length ? saved : mods).map(m => m.id);
    const byId = {}; mods.forEach(m => byId[m.id] = m);
    const line = [nama, kelas].concat(order.map(id => (byId[id] && byId[id].acc !== null && byId[id].acc !== undefined) ? byId[id].acc : ''),
      [weak ? weak.nama : '-', strong ? strong.nama : '-', now_()]);
    const ar = findRow_(a, nama);
    if (ar > 0) a.getRange(ar, 1, 1, line.length).setValues([line]); else a.appendRow(line);
  }
  return { ok: true };
}

/* ================= RANKING LIVE =================
 * Mata setiap soalan = 40 saat − masa menjawab betul (ketepatan milisaat).
 * Kedudukan: lebih banyak soalan dijawab betul = lebih tinggi (dikira LIVE walaupun belum habis).
 * Jika sama, mata kelajuan menentukan. Seri dipecahkan oleh jumlah masa terpendek. */
function rank_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('RANK');
  if (hit) return JSON.parse(hit);
  const sh = sheet_('MURID', H_MURID);
  if (sh.getRange(1, 17).getValue() !== H_MURID[16]) sh.getRange(1, 17, 1, 3).setValues([H_MURID.slice(16, 19)]).setFontWeight('bold').setBackground('#0b1f3a').setFontColor('#ffffff');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const list = sh.getRange(2, 1, last - 1, 19).getValues()
    .filter(r => r[0] && (Number(r[18]) > 0 || Number(r[16]) > 0))
    .map(r => ({ nama: String(r[0]), kelas: String(r[1]), mata: Number(r[16]) || 0, masa: Number(r[17]) || 0, betul: Number(r[18]) || 0 }))
    
    .sort((a, b) => b.betul - a.betul || b.mata - a.mata || a.masa - b.masa || a.nama.localeCompare(b.nama))
    .slice(0, 50);
  cache.put('RANK', JSON.stringify(list), 10);
  return list;
}

/* ================= LOG SETIAP JAWAPAN ================= */
function log_(p) {
  const rows = (p.rows || []).map(r => [r.masa || now_(), norm_(p.nama), norm_(p.kelas), r.bab, r.soalan, r.jawapan, r.betul, r.keputusan, r.cubaan]);
  if (!rows.length) return { ok: true };
  const sh = sheet_('LOG_JAWAPAN', H_LOG);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, H_LOG.length).setValues(rows);
  return { ok: true, n: rows.length };
}

/* ================= KEPUTUSAN MODUL ================= */
function result_(p) {
  const sh = sheet_('KEPUTUSAN_MODUL', H_KEP);
  const pct = p.total ? Math.round(p.skor / p.total * 100) : 0;
  sh.appendRow([now_(), norm_(p.nama), norm_(p.kelas), p.bab, p.skor, p.total, pct]);
  return { ok: true };
}

/** Jalankan sekali secara manual untuk memberi kebenaran (Sheets + Drive). */
function setup() {
  sheet_('MURID', H_MURID); sheet_('KEPUTUSAN_MODUL', H_KEP); sheet_('LOG_JAWAPAN', H_LOG); sheet_('_DATA', H_DATA);
  DriveApp.getFolderById(FOLDER_ID).getName();
  const s1 = ss_().getSheetByName('Sheet1') || ss_().getSheetByName('Helaian1');
  if (s1 && s1.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(s1);
}
