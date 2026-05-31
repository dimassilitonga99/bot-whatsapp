'use strict';

// ════════════════════════════════════════════════════════════════
//   BOT WHATSAPP - LAPORAN & CARI HARGA BARANG
//   Versi 3.6 - Auto Wizard + Auto Nama Kasir CP
// ════════════════════════════════════════════════════════════════

require('dotenv').config();
const express = require('express');
const axios   = require('axios');
const xlsx    = require('xlsx');
const path    = require('path');
const fs      = require('fs');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// ════════════════════════════════════════════════════════════════
//   1. KONFIGURASI
// ════════════════════════════════════════════════════════════════

const CONFIG = {
  port:         parseInt(process.env.PORT) || 3000,
  appName:      'Bot Toko Perabot',
  fonnteToken:  process.env.FONNTE_TOKEN,
  geminiKey:    process.env.GEMINI_KEY,
  geminiKey2:   process.env.GEMINI_KEY2 || '',
  geminiKey3:   process.env.GEMINI_KEY3 || '',
  adminNumber:  process.env.ADMIN_NUMBER || '6285829278962',
  fonnteUrl:    'https://api.fonnte.com/send',

  paths: {
    storage:  path.join(__dirname, 'storage'),
    members:  path.join(__dirname, 'storage', 'members.json'),
    kontak:   path.join(__dirname, 'storage', 'kontak.json'),
    sesi:     path.join(__dirname, 'storage', 'sesi.json'),
    disapa:   path.join(__dirname, 'storage', 'disapa.json'),
    excel:    path.join(__dirname, 'harga_barang_5toko.xlsx'),
    logs:     path.join(__dirname, 'logs', 'error.log'),
  },

  sesiTimeoutMenit: 30,
  maxRetry:         3,
  retryDelay:       2000,
  maxMember:        20,
  maxHasilCari:     20,
};

let currentGeminiKeyIndex = 0;
const geminiKeys = function() {
  const keys = [CONFIG.geminiKey];
  if (CONFIG.geminiKey2) keys.push(CONFIG.geminiKey2);
  if (CONFIG.geminiKey3) keys.push(CONFIG.geminiKey3);
  return keys;
};

function geminiUrl() {
  const keys = geminiKeys();
  const key = keys[currentGeminiKeyIndex % keys.length];
  return 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + key;
}

function rotateGeminiKey() {
  const keys = geminiKeys();
  if (keys.length > 1) {
    currentGeminiKeyIndex = (currentGeminiKeyIndex + 1) % keys.length;
    log.info('GEMINI', 'Rotate ke key #' + (currentGeminiKeyIndex + 1));
  }
}

if (!CONFIG.fonnteToken || !CONFIG.geminiKey) {
  console.error('\n❌ ERROR: FONNTE_TOKEN atau GEMINI_KEY belum diisi\n');
  process.exit(1);
}

[CONFIG.paths.storage, path.dirname(CONFIG.paths.logs)].forEach(function(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// ════════════════════════════════════════════════════════════════
//   2. EMOJI & DATA TOKO
// ════════════════════════════════════════════════════════════════

const EMOJI_NUM = ['0️⃣','1️⃣','2️⃣','3️⃣','4️⃣','5️⃣','6️⃣','7️⃣','8️⃣','9️⃣','🔟'];
function emojiNum(n) { return (n >= 0 && n <= 10) ? EMOJI_NUM[n] : String(n); }

const TOKO_LIST = [
  { kode: 'nk',     nama: 'Nasional Kitchen',          alias: ['nk', 'nasional', 'kitchen'] },
  { kode: 'tdm',    nama: 'Perabot Mama TDM',          alias: ['tdm', 'mama tdm'] },
  { kode: 'oesapa', nama: 'Perabot Mama Oesapa',       alias: ['oesapa', 'mama oesapa'] },
  { kode: 'kefa',   nama: 'Perabot Mamaku Kefamenanu', alias: ['kefa', 'mamaku', 'kefamenanu'] },
  { kode: 'cp',     nama: 'Central Perabot',           alias: ['cp', 'central'] },
];

const TOKO_COLS = {
  nk:     { ecer: 'Ecer NK',     ambil: 'Ambil NK',     stok: 'Stok NK'     },
  tdm:    { ecer: 'Ecer TDM',    ambil: 'Ambil TDM',    stok: 'Stok TDM'    },
  oesapa: { ecer: 'Ecer Oesapa', ambil: 'Ambil Oesapa', stok: 'Stok Oesapa' },
  kefa:   { ecer: 'Ecer Kefa',   ambil: 'Ambil Kefa',   stok: 'Stok Kefa'   },
  cp:     { ecer: 'Ecer CP',     ambil: 'Ambil CP',     stok: 'Stok CP'     },
};

const NAMA_TOKO = {};
TOKO_LIST.forEach(function(t) { NAMA_TOKO[t.kode] = t.nama; });

// ★★★ FIELD WIZARD per TOKO ★★★
const FIELD_LAPORAN = {
  nk: [
    { key: 'k1',     label: 'Kassa 1',               emoji: '💵' },
    { key: 'k2',     label: 'Kassa 2',               emoji: '💵' },
    { key: 'total',  label: 'Total Keseluruhan',     emoji: '📦', auto: true },
    { key: 'tunai',  label: 'Tunai',                 emoji: '💰' },
    { key: 'debit',  label: 'Debit',                 emoji: '💳' },
    { key: 'kredit', label: 'Credit',                emoji: '💳' },
    { key: 'ecer',   label: 'Ecer',                  emoji: '🛒' },
    { key: 'grosir', label: 'Grosir',                emoji: '📦' },
  ],
  tdm: [
    { key: 'k1',     label: 'Kassa 1',               emoji: '💵' },
    { key: 'k2',     label: 'Kassa 2',               emoji: '💵' },
    { key: 'total',  label: 'Total Keseluruhan',     emoji: '📦', auto: true },
    { key: 'tunai',  label: 'Tunai',                 emoji: '💰' },
    { key: 'debit',  label: 'Debit',                 emoji: '💳' },
    { key: 'kredit', label: 'Credit',                emoji: '💳' },
  ],
  oesapa: [
    { key: 'k1',     label: 'Kassa 1',               emoji: '💵' },
    { key: 'k2',     label: 'Kassa 2',               emoji: '💵' },
    { key: 'total',  label: 'Total Keseluruhan',     emoji: '📦', auto: true },
    { key: 'tunai',  label: 'Tunai',                 emoji: '💰' },
    { key: 'debit',  label: 'Debit',                 emoji: '💳' },
    { key: 'kredit', label: 'Credit',                emoji: '💳' },
  ],
  kefa: [
    { key: 'k1',     label: 'Kassa 1',               emoji: '💵' },
    { key: 'k2',     label: 'Kassa 2',               emoji: '💵' },
    { key: 'total',  label: 'Total Keseluruhan',     emoji: '📦', auto: true },
    { key: 'tunai',  label: 'Tunai',                 emoji: '💰' },
    { key: 'debit',  label: 'Debit',                 emoji: '💳' },
    { key: 'kredit', label: 'Credit',                emoji: '💳' },
  ],
  cp: [
    { key: 'k1', label: 'Kassa 1', emoji: '💵' },
    { key: 'k2', label: 'Kassa 2', emoji: '💵' },
    { key: 'k3', label: 'Kassa 3', emoji: '💵' },
    { key: 'k4', label: 'Kassa 4', emoji: '💵' },
    { key: 'total',  label: 'Total Keseluruhan',     emoji: '📦', auto: true },
    { key: 'tunai',  label: 'Tunai',                 emoji: '💰' },
    { key: 'debit',  label: 'Debit',                 emoji: '💳' },
    { key: 'kredit', label: 'Credit',                emoji: '💳' },
    { key: 'ecer',   label: 'Ecer',                  emoji: '🛒' },
    { key: 'grosir', label: 'Grosir',                emoji: '📦' },
    { key: 'promo',        label: 'Total Kasir Promo',  emoji: '🎁' },
    { key: 'promotunai',   label: 'Promo - Tunai',      emoji: '💰' },
    { key: 'promodebit',   label: 'Promo - Debit',      emoji: '💳' },
    { key: 'promokredit',  label: 'Promo - Credit',     emoji: '💳' },
    { key: 'parkirkomputer', label: 'Parkir di Komputer', emoji: '🅿️' },
    { key: 'parkirluar',     label: 'Parkir Stor Luar',   emoji: '🅿️' },
  ],
};

// ★★★ NAMA KASIR DEFAULT CENTRAL PERABOT ★★★
// Ganti nama-nama ini sesuai kasir asli kamu
const KASIR_CP_DEFAULT = {
  k1: 'Yuni-Salsa',
  k2: 'Nanda-Umi-Marselina',
  k3: 'Febri-Jien-Tika',
  k4: 'Delfi-Tirsa',
};

const DEFAULT_MEMBERS = [
  '6285253949803','6285737005301','6285211988252','6281383924057',
  '6282235572821','6287841617474','6281584937710'
];

const DEFAULT_KONTAK = {
  '6285253949803': 'Pak Security Marthen',
  '6285737005301': 'Kak Bagas Pacar Beda Agama',
  '6285211988252': 'Kak Admin Marketplace',
  '6281383924057': 'Kak Fajar (Buka Mas Fajar Kefa)',
  '6282235572821': 'Kak yang Saya Tidak Tau Namanya',
  '6287841617474': 'Mas Awin Gacor',
  '6281584937710': 'Kak Safira',
  '6282266026564': 'Mas Abi Mustafa',
  '6285829278962': 'Admin'
};

const ADMIN_COMMANDS = ['daftar','hapus','listmember','namakontak','hapuskontak','listkontak','reload','info'];

// ════════════════════════════════════════════════════════════════
//   3. LOGGER & UTILS
// ════════════════════════════════════════════════════════════════

function timestamp() {
  const d = new Date(Date.now() + 8 * 3600000);
  return d.toISOString().replace('T', ' ').substring(0, 19);
}

const log = {
  info:  function(ctx, msg) { console.log('[' + timestamp() + '] [INFO ] [' + ctx + '] ' + msg); },
  warn:  function(ctx, msg) { console.warn('[' + timestamp() + '] [WARN ] [' + ctx + '] ' + msg); },
  error: function(ctx, msg, data) {
    const fullMsg = '[' + timestamp() + '] [ERROR] [' + ctx + '] ' + msg;
    console.error(fullMsg, data || '');
    try { fs.appendFileSync(CONFIG.paths.logs, fullMsg + (data ? ' | ' + JSON.stringify(data) : '') + '\n'); } catch (e) {}
  },
};

function getWaktu() {
  const j = new Date(Date.now() + 8 * 3600000).getUTCHours();
  if (j >= 5  && j < 11) return 'Pagi';
  if (j >= 11 && j < 15) return 'Siang';
  if (j >= 15 && j < 19) return 'Sore';
  return 'Malam';
}

function getTanggal(kemarin) {
  const d = new Date(Date.now() + 8 * 3600000);
  if (kemarin) d.setDate(d.getDate() - 1);
  return d.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
}

function formatRp(n, prefix, fallback) {
  prefix = prefix || 'Rp';
  fallback = fallback || prefix + ' -';
  const v = parseFloat(n) || 0;
  if (v === 0) return fallback;
  return prefix + ' ' + v.toLocaleString('id-ID');
}

const fRp  = function(n) { return formatRp(n, 'Rp',  'Rp -'); };
const fRpP = function(n) { return formatRp(n, 'Rp',  'Rp 0'); };
const GARIS_TEBAL = '━━━━━━━━━━━━━━━━━━';
const GARIS_TIPIS = '──────────────────';

// ════════════════════════════════════════════════════════════════
//   4. STORAGE
// ════════════════════════════════════════════════════════════════

function loadJSON(filePath, defaultValue) {
  try { if (fs.existsSync(filePath)) return JSON.parse(fs.readFileSync(filePath, 'utf8')); }
  catch (e) { log.error('STORAGE', 'Gagal load ' + filePath, e.message); }
  return defaultValue;
}

function saveJSON(filePath, data) {
  try { fs.writeFileSync(filePath, JSON.stringify(data, null, 2)); return true; }
  catch (e) { log.error('STORAGE', 'Gagal save', e.message); return false; }
}

let MEMBERS      = loadJSON(CONFIG.paths.members, DEFAULT_MEMBERS.slice());
let KONTAK       = loadJSON(CONFIG.paths.kontak,  Object.assign({}, DEFAULT_KONTAK));
let SESI         = loadJSON(CONFIG.paths.sesi,    {});
let SUDAH_DISAPA = loadJSON(CONFIG.paths.disapa,  {});

log.info('INIT', 'Loaded: ' + MEMBERS.length + ' members, ' + Object.keys(KONTAK).length + ' kontak');

function isAdmin(n)  { return n === CONFIG.adminNumber; }
function isMember(n) { return isAdmin(n) || MEMBERS.indexOf(n) >= 0; }
function getNama(n)  { return KONTAK[n] || null; }

function tambahMember(nomor) {
  if (!nomor || !/^[0-9]+$/.test(nomor)) return { ok: false, alasan: 'Format nomor tidak valid' };
  if (isAdmin(nomor))                    return { ok: false, alasan: 'Nomor itu adalah admin' };
  if (MEMBERS.indexOf(nomor) >= 0)       return { ok: false, alasan: 'Nomor sudah terdaftar' };
  if (MEMBERS.length >= CONFIG.maxMember) return { ok: false, alasan: 'Slot penuh (' + CONFIG.maxMember + ')' };
  MEMBERS.push(nomor); saveJSON(CONFIG.paths.members, MEMBERS); return { ok: true };
}

function hapusMember(nomor) {
  const idx = MEMBERS.indexOf(nomor);
  if (idx === -1) return { ok: false, alasan: 'Nomor tidak ditemukan' };
  MEMBERS.splice(idx, 1); saveJSON(CONFIG.paths.members, MEMBERS); return { ok: true };
}

function setNama(nomor, nama) {
  if (!nomor || !nama) return { ok: false, alasan: 'Nomor & nama wajib' };
  KONTAK[nomor] = nama.trim(); saveJSON(CONFIG.paths.kontak, KONTAK); return { ok: true };
}

function hapusKontak(nomor) {
  if (!KONTAK[nomor]) return { ok: false, alasan: 'Kontak tidak ditemukan' };
  delete KONTAK[nomor]; saveJSON(CONFIG.paths.kontak, KONTAK); return { ok: true };
}

const TIMEOUT_MS = CONFIG.sesiTimeoutMenit * 60 * 1000;
function getSesi(nomor) { if (!SESI[nomor]) SESI[nomor] = {}; SESI[nomor]._lastActive = Date.now(); return SESI[nomor]; }
function resetSesi(nomor) { SESI[nomor] = { _lastActive: Date.now() }; saveJSON(CONFIG.paths.sesi, SESI); }
function updateSesi(nomor, data) {
  if (!SESI[nomor]) SESI[nomor] = {};
  Object.assign(SESI[nomor], data, { _lastActive: Date.now() });
  saveJSON(CONFIG.paths.sesi, SESI);
}

setInterval(function() {
  const now = Date.now();
  let buang = 0;
  Object.keys(SESI).forEach(function(n) {
    if (now - (SESI[n]._lastActive || 0) > TIMEOUT_MS) { delete SESI[n]; buang++; }
  });
  if (buang > 0) saveJSON(CONFIG.paths.sesi, SESI);
}, 5 * 60 * 1000);

// ════════════════════════════════════════════════════════════════
//   5. EXCEL & SMART SEARCH
// ════════════════════════════════════════════════════════════════

let DATA_BARANG = [];

function loadExcel() {
  if (!fs.existsSync(CONFIG.paths.excel)) return false;
  try {
    const wb = xlsx.readFile(CONFIG.paths.excel);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const allRows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: 0, blankrows: false });
    if (allRows.length < 3) return false;
    const headers = allRows[1].map(function(h) { return String(h || '').trim(); });
    function findCol(name) {
      return headers.findIndex(function(h) {
        return h.toLowerCase().replace(/\s+/g, ' ').trim() === name.toLowerCase();
      });
    }
    const colMap = {
      kode: findCol('Kode Item'), nama: findCol('Nama Item'),
      jenis: findCol('Jenis'), merek: findCol('Merek'), satuan: findCol('Satuan'),
    };
    const tokoColMap = {};
    Object.keys(TOKO_COLS).forEach(function(kode) {
      const c = TOKO_COLS[kode];
      tokoColMap[kode] = { ecer: findCol(c.ecer), ambil: findCol(c.ambil), stok: findCol(c.stok) };
    });
    if (colMap.kode === -1 || colMap.nama === -1) return false;
    DATA_BARANG = [];
    for (let i = 2; i < allRows.length; i++) {
      const row = allRows[i];
      if (!row || row.length === 0) continue;
      const kode = String(row[colMap.kode] || '').trim().toUpperCase();
      if (!kode || kode === 'UNDEFINED' || kode === '0') continue;
      const item = {
        kode, nama: String(row[colMap.nama] || '').trim().toUpperCase(),
        jenis: String(row[colMap.jenis] || '').trim(),
        merek: String(row[colMap.merek] || '').trim(),
        satuan: String(row[colMap.satuan] || '').trim(),
        harga: {},
      };
      Object.keys(TOKO_COLS).forEach(function(tk) {
        const tc = tokoColMap[tk];
        item.harga[tk] = {
          ecer:  tc.ecer  >= 0 ? (parseFloat(row[tc.ecer])  || 0) : 0,
          ambil: tc.ambil >= 0 ? (parseFloat(row[tc.ambil]) || 0) : 0,
          stok:  tc.stok  >= 0 ? (parseInt(row[tc.stok])    || 0) : 0,
        };
      });
      DATA_BARANG.push(item);
    }
    log.info('EXCEL', 'Loaded ' + DATA_BARANG.length + ' item');
    return true;
  } catch (e) { log.error('EXCEL', 'Gagal load', e.message); return false; }
}

function saveExcel() {
  try {
    const rows = DATA_BARANG.map(function(d) {
      const row = { 'Kode Item': d.kode, 'Nama Item': d.nama, 'Jenis': d.jenis, 'Merek': d.merek, 'Satuan': d.satuan };
      Object.keys(TOKO_COLS).forEach(function(k) {
        const c = TOKO_COLS[k];
        row[c.ecer] = d.harga[k].ecer; row[c.ambil] = d.harga[k].ambil; row[c.stok] = d.harga[k].stok;
      });
      return row;
    });
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, xlsx.utils.json_to_sheet(rows), 'Data Barang');
    xlsx.writeFile(wb, CONFIG.paths.excel);
    return true;
  } catch (e) { log.error('EXCEL', 'Gagal save', e.message); return false; }
}

function levenshtein(a, b) {
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  const matrix = [];
  for (let i = 0; i <= b.length; i++) matrix[i] = [i];
  for (let j = 0; j <= a.length; j++) matrix[0][j] = j;
  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) matrix[i][j] = matrix[i - 1][j - 1];
      else matrix[i][j] = Math.min(matrix[i - 1][j - 1] + 1, matrix[i][j - 1] + 1, matrix[i - 1][j] + 1);
    }
  }
  return matrix[b.length][a.length];
}

function bersihkanTeks(str) { return String(str).toUpperCase().replace(/[^A-Z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim(); }

function kataMirip(kata, target) {
  if (kata === target) return true;
  if (kata.length <= 2 || target.length <= 2) return kata === target;
  if (target.indexOf(kata) >= 0 || kata.indexOf(target) >= 0) return true;
  const maxJarak = kata.length <= 4 ? 1 : 2;
  return levenshtein(kata, target) <= maxJarak;
}

function hitungSkor(item, words) {
  const namaBersih = bersihkanTeks(item.nama);
  const kodeBersih = bersihkanTeks(item.kode);
  const namaWords  = namaBersih.split(/\s+/);
  let skor = 0, exactMatch = 0, fuzzyMatch = 0, partialMatch = 0;
  words.forEach(function(w) {
    let found = false;
    if (namaBersih.indexOf(w) >= 0) { exactMatch++; skor += 10; found = true; }
    if (kodeBersih.indexOf(w) >= 0) { exactMatch++; skor += 10; found = true; }
    if (!found) for (let i = 0; i < namaWords.length; i++) {
      if (kataMirip(w, namaWords[i])) { fuzzyMatch++; skor += 5; found = true; break; }
    }
    if (!found) for (let i = 0; i < namaWords.length; i++) {
      if (namaWords[i].indexOf(w) >= 0 || w.indexOf(namaWords[i]) >= 0) {
        partialMatch++; skor += 3; found = true; break;
      }
    }
  });
  if (exactMatch + fuzzyMatch + partialMatch >= words.length) skor += 20;
  return { skor, exactMatch, fuzzyMatch, partialMatch, totalMatch: exactMatch + fuzzyMatch + partialMatch };
}

function cariBarang(keyword) {
  const q = keyword.trim().toUpperCase();
  const qBersih = bersihkanTeks(q);
  const words = qBersih.split(/\s+/).filter(function(w) { return w.length > 0; });
  if (words.length === 0) return { hasil: [], saran: [], tipeHasil: 'kosong' };
  const byKode = DATA_BARANG.filter(function(d) { return d.kode === q; });
  if (byKode.length > 0) return { hasil: byKode, saran: [], tipeHasil: 'exact', totalDitemukan: byKode.length };
  const exactResults = DATA_BARANG.filter(function(d) {
    const namaBersih = bersihkanTeks(d.nama);
    return words.every(function(w) { return namaBersih.indexOf(w) >= 0 || d.kode.indexOf(w) >= 0; });
  });
  if (exactResults.length > 0) {
    return { hasil: exactResults.slice(0, CONFIG.maxHasilCari), saran: [], tipeHasil: 'exact', totalDitemukan: exactResults.length };
  }
  const skorItems = [];
  DATA_BARANG.forEach(function(item) {
    const info = hitungSkor(item, words);
    if (info.skor > 0 && info.totalMatch >= Math.ceil(words.length * 0.5)) {
      skorItems.push({ item, skor: info.skor, fuzzyMatch: info.fuzzyMatch });
    }
  });
  skorItems.sort(function(a, b) { return b.skor - a.skor; });
  if (skorItems.length > 0) {
    const batasSkor = skorItems[0].skor * 0.5;
    const hasilBagus = skorItems.filter(function(s) { return s.skor >= batasSkor; });
    const hasilTerbatas = hasilBagus.slice(0, CONFIG.maxHasilCari);
    const adaFuzzy = hasilTerbatas.some(function(s) { return s.fuzzyMatch > 0; });
    return {
      hasil: hasilTerbatas.map(function(s) { return s.item; }),
      saran: [], tipeHasil: adaFuzzy ? 'fuzzy' : 'exact', totalDitemukan: hasilBagus.length,
    };
  }
  const saranSet = {};
  words.forEach(function(w) {
    if (w.length < 2) return;
    DATA_BARANG.forEach(function(item) {
      const namaWords = bersihkanTeks(item.nama).split(/\s+/);
      namaWords.forEach(function(nw) {
        if (kataMirip(w, nw) || nw.indexOf(w) >= 0 || w.indexOf(nw) >= 0) {
          if (!saranSet[item.kode]) saranSet[item.kode] = { item, matchCount: 0 };
          saranSet[item.kode].matchCount++;
        }
      });
    });
  });
  const saranList = Object.values(saranSet)
    .sort(function(a, b) { return b.matchCount - a.matchCount; })
    .slice(0, 5).map(function(s) { return s.item; });
  return { hasil: [], saran: saranList, tipeHasil: saranList.length > 0 ? 'saran' : 'kosong' };
}

function updateStok(kode, tokoKode, jumlah) {
  const k = kode.trim().toUpperCase();
  const item = DATA_BARANG.find(function(d) { return d.kode === k; });
  if (!item) return null;
  item.harga[tokoKode].stok = jumlah;
  saveExcel();
  return item;
}

loadExcel();
// ════════════════════════════════════════════════════════════════
//   6. KIRIM WHATSAPP (FONNTE)
// ════════════════════════════════════════════════════════════════

function tunggu(ms) { return new Promise(function(r) { setTimeout(r, ms); }); }

async function kirimWA(target, message, retry) {
  retry = retry || 0;
  try {
    await axios.post(CONFIG.fonnteUrl, { target: target, message: message },
      { headers: { Authorization: CONFIG.fonnteToken }, timeout: 10000 });
    log.info('FONNTE', 'OK ke ' + target);
    return true;
  } catch (err) {
    log.warn('FONNTE', 'Gagal attempt ' + (retry + 1));
    if (retry < CONFIG.maxRetry - 1) { await tunggu(CONFIG.retryDelay); return kirimWA(target, message, retry + 1); }
    log.error('FONNTE', 'GAGAL TOTAL', err.message);
    return false;
  }
}

// ════════════════════════════════════════════════════════════════
//   7. GEMINI AI
// ════════════════════════════════════════════════════════════════

async function analisaGambar(imageUrl, prompt) {
  const imgResp   = await axios.get(imageUrl, { responseType: 'arraybuffer', timeout: 15000 });
  const mimeType  = imgResp.headers['content-type'] || 'image/jpeg';
  const imageData = Buffer.from(imgResp.data).toString('base64');
  const resp = await axios.post(geminiUrl(), {
    contents: [{ parts: [{ inline_data: { mime_type: mimeType, data: imageData } }, { text: prompt }] }],
  }, { timeout: 30000 });
  return resp.data.candidates[0].content.parts[0].text || '';
}

// ════════════════════════════════════════════════════════════════
//   ★★★ AI CHAT BARANG ★★★
// ════════════════════════════════════════════════════════════════

// Cari barang yang relevan dengan pertanyaan (untuk context AI)
function cariRelevan(pertanyaan, maxResults) {
  maxResults = maxResults || 50;
  const q = pertanyaan.toUpperCase();
  const qBersih = bersihkanTeks(q);
  const words = qBersih.split(/\s+/).filter(function(w) { return w.length >= 2; });
  
  if (words.length === 0) return DATA_BARANG.slice(0, maxResults);
  
  // Hitung skor relevan untuk tiap barang
  const scored = [];
  DATA_BARANG.forEach(function(item) {
    let score = 0;
    const namaBersih = bersihkanTeks(item.nama);
    const merekBersih = bersihkanTeks(item.merek);
    const jenisBersih = bersihkanTeks(item.jenis);
    const namaWords = namaBersih.split(/\s+/);
    
    words.forEach(function(w) {
      // Exact match
      if (namaBersih.indexOf(w) >= 0) score += 10;
      if (merekBersih.indexOf(w) >= 0) score += 8;
      if (jenisBersih.indexOf(w) >= 0) score += 6;
      if (item.kode.indexOf(w) >= 0) score += 15;
      
      // Fuzzy match per kata
      for (let i = 0; i < namaWords.length; i++) {
        if (kataMirip(w, namaWords[i])) {
          score += 5;
          break;
        }
      }
    });
    
    if (score > 0) scored.push({ item: item, score: score });
  });
  
  scored.sort(function(a, b) { return b.score - a.score; });
  
  // Kalau hasil sedikit, tambah hasil dari fuzzy yang lebih longgar
  if (scored.length < 5) {
    const tambahan = [];
    DATA_BARANG.forEach(function(item) {
      if (scored.find(function(s) { return s.item.kode === item.kode; })) return;
      const namaBersih = bersihkanTeks(item.nama);
      words.forEach(function(w) {
        if (w.length >= 3 && namaBersih.indexOf(w.substring(0, 3)) >= 0) {
          tambahan.push({ item: item, score: 2 });
        }
      });
    });
    scored.push(...tambahan.slice(0, 10));
  }
  
  return scored.slice(0, maxResults).map(function(s) { return s.item; });
}

// Format data barang jadi context untuk AI
function formatBarangUntukAI(items) {
  let context = '';
  items.forEach(function(d, i) {
    context += (i + 1) + '. Kode: ' + d.kode + ' | Nama: ' + d.nama;
    if (d.jenis) context += ' | Jenis: ' + d.jenis;
    if (d.merek) context += ' | Merek: ' + d.merek;
    context += ' | Satuan: ' + d.satuan + '\n';
    
    // Harga & stok semua toko
    context += '   Harga & Stok per Toko:\n';
    Object.keys(NAMA_TOKO).forEach(function(tk) {
      const h = d.harga[tk];
      const stok = h.stok > 0 ? h.stok : 'KOSONG';
      context += '   - ' + NAMA_TOKO[tk] + ': Ecer Rp' + h.ecer.toLocaleString('id-ID') + 
                 ', Ambil Rp' + h.ambil.toLocaleString('id-ID') + 
                 ', Stok: ' + stok + '\n';
    });
    context += '\n';
  });
  return context;
}

// AI Chat tentang barang
async function aiChatBarang(pertanyaan, sender) {
  try {
    // Cari barang relevan untuk context
    let relevantItems = cariRelevan(pertanyaan, 50);
    
    // Kalau gak ada hasil relevan, ambil 50 barang random sebagai context
    if (relevantItems.length === 0) {
      log.warn('AI_CHAT', 'Tidak ada barang relevan, pakai context kosong');
      relevantItems = [];
    }
    
    const context = relevantItems.length > 0 
      ? formatBarangUntukAI(relevantItems)
      : '(Tidak ada barang yang cocok dengan kata kunci di database)';
    
    const nama = getNama(sender);
    const sapaan = nama ? nama : 'kakak';
    const totalBarang = DATA_BARANG.length;
    
    const prompt = 'Kamu adalah asisten AI toko perabot bernama "Bot Perabot". Kamu ramah, helpful, dan jawab dengan bahasa Indonesia santai.\n\n' +
      'Saya punya 5 toko: Nasional Kitchen (NK), Perabot Mama TDM (TDM), Perabot Mama Oesapa (Oesapa), Perabot Mamaku Kefamenanu (Kefa), dan Central Perabot (CP).\n' +
      'Total database ada ' + totalBarang + ' barang.\n\n' +
      'DATA BARANG YANG RELEVAN DENGAN PERTANYAAN (' + relevantItems.length + ' item):\n' +
      context + '\n\n' +
      'PERTANYAAN USER (' + sapaan + '):\n' +
      '"' + pertanyaan + '"\n\n' +
      'ATURAN MENJAWAB:\n' +
      '1. Jawab ramah dan natural, panggil user dengan "' + sapaan + '"\n' +
      '2. Gunakan emoji yang sesuai (📦 🏷️ 💰 🏪 ✅ ⚠️ dll)\n' +
      '3. Format harga: Rp 1.000.000 (pakai titik)\n' +
      '4. Kalau bandingkan toko, gunakan format yang mudah dibaca\n' +
      '5. Kalau stok kosong, sebutkan dengan jelas ⚠️\n' +
      '6. Pakai *bold* untuk teks penting (WhatsApp format)\n' +
      '7. Maksimal 1500 karakter\n' +
      '8. PENTING: Kalau di data tidak ada barang yang cocok, JANGAN bilang "tidak menemukan informasi". Sebagai gantinya:\n' +
      '   - Tampilkan barang yang MIRIP dengan kata kunci dari data di atas\n' +
      '   - Beri saran kata kunci alternatif\n' +
      '   - Tawarkan untuk mencari dengan kata lain\n' +
      '9. Selalu berikan jawaban yang berguna, walaupun data terbatas\n' +
      '10. Akhiri dengan tawaran bantuan ("Ada yang ingin ditanyakan lagi?" atau "Mau dicarikan yang lain?")\n\n' +
      'Jawab pertanyaan user sekarang dengan ramah dan helpful:';

    log.info('AI_CHAT', 'Kirim ke Gemini, context: ' + relevantItems.length + ' items');
    
    const resp = await axios.post(geminiUrl(), {
      contents: [{ parts: [{ text: prompt }] }],
    }, { timeout: 30000 });
    
    const jawaban = resp.data.candidates[0].content.parts[0].text || '';
    log.info('AI_CHAT', 'Jawaban length: ' + jawaban.length);
    
    if (!jawaban || jawaban.trim().length < 10) {
      return null;
    }
    
    return jawaban.trim();
    
    } catch (err) {
    log.error('AI_CHAT', 'Gagal: ' + err.message);
    if (err.response) {
      log.error('AI_CHAT', 'Response: ' + JSON.stringify(err.response.data).substring(0, 200));
      // Kalau error 429 (quota), rotate ke key berikutnya
      if (err.response.status === 429) {
        rotateGeminiKey();
        log.warn('AI_CHAT', 'Quota habis, rotate API key & retry...');
        // Retry sekali dengan key baru
        try {
          const resp2 = await axios.post(geminiUrl(), {
            contents: [{ parts: [{ text: prompt }] }],
          }, { timeout: 30000 });
          const jawaban2 = resp2.data.candidates[0].content.parts[0].text || '';
          if (jawaban2 && jawaban2.trim().length >= 10) return jawaban2.trim();
        } catch (err2) {
          log.error('AI_CHAT', 'Retry juga gagal: ' + err2.message);
        }
      }
    }
    return null;
  }
}

// Cek apakah pesan adalah pertanyaan tentang barang
function isPertanyaanBarang(low) {
  // Kata kerja/tanya yang umum
  const KATA_TANYA = [
    'stok', 'stock', 'harga', 'price', 'berapa', 'ada gak', 'ada ga', 'ada kah',
    'apakah ada', 'masih ada', 'bandingkan', 'banding', 'compare', 'samakan',
    'rekomendasi', 'rekomen', 'sarankan', 'saran',
    'termurah', 'termahal', 'paling murah', 'paling mahal',
    'kosong', 'habis', 'tersedia', 'tersisa',
    'cek', 'check', 'lihat', 'tampilkan',
    'total', 'jumlah', 'berapa banyak',
    'cari', 'mencari', 'mau', 'butuh', 'perlu',
    'info', 'detail', 'data',
  ];
  
  // Kata kunci barang umum
  const KATA_BARANG_UMUM = [
    'panci', 'dandang', 'wajan', 'penggorengan', 'rice cooker', 'kompor', 'konpor',
    'gelas', 'piring', 'mangkok', 'sendok', 'garpu', 'pisau',
    'eagle', 'maxim', 'sunkist', 'golden', 'paramount', 'hock', 'sunlife', 'miyako',
    'aluminium', 'alm', 'stainless', 'plastik', 'kaca', 'keramik',
    'kursi', 'meja', 'lemari', 'rak',
    'sumbu', 'minyak', 'gas',
    'tl', 'serbaguna', 'susu',
    'jar', 'drink', 'keranjang',
    'periuk', 'panci tl',
  ];
  
  // Kata kunci toko
  const KATA_TOKO = ['nk', 'tdm', 'oesapa', 'kefa', 'cp', 'nasional', 'central', 'mama', 'mamaku', 'kefamenanu'];
  
  const adaKataTanya = KATA_TANYA.some(function(k) { return low.indexOf(k) >= 0; });
  const adaKataBarang = KATA_BARANG_UMUM.some(function(k) { return low.indexOf(k) >= 0; });
  const adaKataToko = KATA_TOKO.some(function(k) { return low.indexOf(k) >= 0; });
  
  // Anggap pertanyaan barang kalau:
  // - Ada kata tanya + ada kata barang/toko
  // - ATAU pesan punya kata barang + nama toko
  // - ATAU pesan panjang (>4 kata) + ada kata barang
  // - ATAU mengandung kode barang (NN diikuti angka)
  const adaKodeBarang = /nn\d{4,5}/i.test(low);
  const jumlahKata = low.split(/\s+/).length;
  
  return adaKodeBarang ||
         (adaKataTanya && adaKataBarang) ||
         (adaKataTanya && adaKataToko) ||
         (adaKataBarang && adaKataToko) ||
         (jumlahKata >= 4 && adaKataBarang);
}

function buatPromptAI(menuType, namaToko, tanggal, tokoKode) {
  const fmt = ' Format WhatsApp dengan emoji. Format rupiah Rp X.XXX.XXX.';
  if (menuType === 1) {
    if (tokoKode === 'cp') {
      return 'Baca data penjualan toko Central Perabot tanggal ' + tanggal + '. Format: 4 kassa (Kassa 1 Yuni-Salsa, Kassa 2 Nanda-Umi-Marselina, Kassa 3 Febri-Jien-Tika, Kassa 4 Delfi-Tirsa), total keseluruhan, metode bayar, jenis penjualan, kasir promo, parkir.' + fmt;
    }
    if (tokoKode === 'nk') {
      return 'Baca data penjualan toko Nasional Kitchen tanggal ' + tanggal + '. Format: 2 kassa, total, metode bayar, jenis penjualan. Rupiah: Rp. X.XXX.XXX.' + fmt;
    }
    if (tokoKode === 'tdm' || tokoKode === 'oesapa' || tokoKode === 'kefa') {
      const tNama = tokoKode === 'tdm' ? 'Perabot Mama TDM' : tokoKode === 'oesapa' ? 'Perabot Mama Oesapa' : 'Perabot Mamaku Kefamenanu';
      return 'Baca data penjualan toko ' + tNama + ' tanggal ' + tanggal + '. Format: 2-3 kassa, total, metode bayar TANPA Ecer/Grosir. Rupiah: Rp. X.XXX.XXX.' + fmt;
    }
    return 'Baca data penjualan toko "' + namaToko + '" tanggal ' + tanggal + '. Buat laporan: kassa, total, metode bayar, jenis penjualan.' + fmt;
  }
  if (menuType === 2) return 'Baca data harga barang toko "' + namaToko + '" tanggal ' + tanggal + '. Buat laporan: barang baru, naik harga, turun harga.' + fmt;
  if (menuType === 3) return 'Baca data marketplace tanggal ' + tanggal + '. Buat laporan: per toko, per channel, metode bayar.' + fmt;
  return 'Buat laporan rapi.' + fmt;
}

// ════════════════════════════════════════════════════════════════
//   8. SAPAAN PINTAR & MOTIVASI
// ════════════════════════════════════════════════════════════════

const SAPAAN_MAP = {
  'halo':     { kategori: 'kasual', balasan: ['Halo', 'Haloo', 'Hai juga', 'Haaai', 'Halooo'] },
  'hai':      { kategori: 'kasual', balasan: ['Hai', 'Haiii', 'Halo juga', 'Hai hai', 'Hai!'] },
  'hi':       { kategori: 'kasual', balasan: ['Hi', 'Hi there', 'Hiii', 'Hi juga', 'Haii'] },
  'hello':    { kategori: 'kasual', balasan: ['Hello', 'Hello juga', 'Helloo', 'Hi there'] },
  'hallo':    { kategori: 'kasual', balasan: ['Hallo', 'Hallooo', 'Halo juga', 'Hai!'] },
  'helo':     { kategori: 'kasual', balasan: ['Helo', 'Helooo', 'Halo juga', 'Hai!'] },
  'hey':      { kategori: 'kasual', balasan: ['Hey', 'Hey juga', 'Heyyy', 'Hai!'] },
  'hei':      { kategori: 'kasual', balasan: ['Hei', 'Hei juga', 'Heii', 'Halo!'] },
  'yo':       { kategori: 'kasual', balasan: ['Yo', 'Yo bro', 'Yoyoyo', 'Yo juga'] },
  'p':        { kategori: 'kasual', balasan: ['p juga', 'Halo', 'Hai', 'Iya?'] },
  'ping':     { kategori: 'kasual', balasan: ['Pong! 🏓', 'Pong 🏓', 'Pong juga'] },
  'pagi':     { kategori: 'waktu', waktu: 'Pagi',  balasan: ['Selamat pagi', 'Pagi juga', 'Pagiii', 'Pagi yang cerah'] },
  'siang':    { kategori: 'waktu', waktu: 'Siang', balasan: ['Selamat siang', 'Siang juga', 'Siaaang'] },
  'sore':     { kategori: 'waktu', waktu: 'Sore',  balasan: ['Selamat sore', 'Sore juga', 'Soreee'] },
  'malam':    { kategori: 'waktu', waktu: 'Malam', balasan: ['Selamat malam', 'Malam juga', 'Malaaam'] },
  'selamat pagi':  { kategori: 'waktu', waktu: 'Pagi',  balasan: ['Selamat pagi juga', 'Pagi yang indah', 'Selamat pagi'] },
  'selamat siang': { kategori: 'waktu', waktu: 'Siang', balasan: ['Selamat siang juga', 'Siang yang cerah', 'Selamat siang'] },
  'selamat sore':  { kategori: 'waktu', waktu: 'Sore',  balasan: ['Selamat sore juga', 'Sore yang menyenangkan', 'Selamat sore'] },
  'selamat malam': { kategori: 'waktu', waktu: 'Malam', balasan: ['Selamat malam juga', 'Malam yang tenang', 'Selamat malam'] },
  'met pagi':  { kategori: 'waktu', waktu: 'Pagi',  balasan: ['Met pagi juga', 'Selamat pagi', 'Pagi!'] },
  'met siang': { kategori: 'waktu', waktu: 'Siang', balasan: ['Met siang juga', 'Selamat siang', 'Siang!'] },
  'met sore':  { kategori: 'waktu', waktu: 'Sore',  balasan: ['Met sore juga', 'Selamat sore', 'Sore!'] },
  'met malam': { kategori: 'waktu', waktu: 'Malam', balasan: ['Met malam juga', 'Selamat malam', 'Malam!'] },
  'assalamualaikum':       { kategori: 'islami', balasan: ['Waalaikumsalam warahmatullahi wabarakatuh', 'Waalaikumsalam', 'Waalaikumsalam wr. wb.'] },
  'assalamu':              { kategori: 'islami', balasan: ['Waalaikumsalam warahmatullahi wabarakatuh', 'Waalaikumsalam'] },
  'assalamualaikum wr wb': { kategori: 'islami', balasan: ['Waalaikumsalam warahmatullahi wabarakatuh', 'Waalaikumsalam wr. wb.'] },
  'salam':                 { kategori: 'islami', balasan: ['Salam juga', 'Salam sejahtera', 'Waalaikumsalam'] },
  'good morning':   { kategori: 'english', waktu: 'Pagi',  balasan: ['Good morning', 'Morning!', 'Good morning to you'] },
  'good afternoon': { kategori: 'english', waktu: 'Siang', balasan: ['Good afternoon', 'Afternoon!', 'Good afternoon to you'] },
  'good evening':   { kategori: 'english', waktu: 'Sore',  balasan: ['Good evening', 'Evening!', 'Good evening to you'] },
  'good night':     { kategori: 'english', waktu: 'Malam', balasan: ['Good night', 'Have a nice dream', 'Sleep well'] },
  'permisi':     { kategori: 'sopan', balasan: ['Iya, silakan', 'Iya, ada yang bisa dibantu?', 'Ya, silakan'] },
  'maaf':        { kategori: 'sopan', balasan: ['Iya, tidak apa-apa', 'Iya, ada yang bisa dibantu?', 'Iya, santai saja'] },
  'maaf ganggu': { kategori: 'sopan', balasan: ['Tidak mengganggu kok', 'Santai saja', 'Tidak apa-apa, silakan'] },
};

const KATA_SAPAAN_LIST = Object.keys(SAPAAN_MAP);

const KALIMAT_MOTIVASI = [
  '✨ _Hari yang baru, semangat yang baru!_ 💪',
  '🌟 _Setiap hari adalah kesempatan untuk jadi lebih baik!_ 🚀',
  '💎 _Kerja keras hari ini, hasil manis besok!_ 🍯',
  '🌈 _Tetap semangat dan jangan menyerah!_ 🔥',
  '⭐ _Senyum dulu, rezeki menyusul!_ 😊',
  '🎯 _Fokus pada tujuan, abaikan keraguan!_ 💯',
  '🌸 _Hari ini akan jadi hari yang luar biasa!_ ✨',
  '🚀 _Sukses dimulai dari langkah kecil hari ini!_ 👣',
  '💪 _Kamu lebih kuat dari yang kamu kira!_ 💯',
  '🌻 _Jangan lupa bahagia hari ini ya!_ 😊',
  '🔥 _Semangat terus, kamu hebat!_ 👏',
  '✨ _Setiap detik adalah anugerah, manfaatkan dengan baik!_ 🙏',
  '🎁 _Hari ini adalah hadiah, makanya disebut present!_ 🎀',
  '🌟 _Percayalah pada diri sendiri, kamu bisa!_ 💪',
  '☀️ _Tetap positif, hal baik akan datang!_ 🌈',
  '💝 _Berkah selalu menyertai orang yang bersyukur_ 🙏',
  '🌺 _Mulailah hari dengan senyuman terbaikmu!_ 😊',
  '⚡ _Energi positif menarik hal-hal positif!_ ✨',
  '🎊 _Hari yang produktif menanti, ayo semangat!_ 🚀',
  '💫 _Mimpi besar dimulai dari hari ini!_ 🌟',
  '🌷 _Hidup itu indah, nikmati setiap momennya!_ 💖',
  '🦋 _Jadilah versi terbaik dari dirimu hari ini!_ ⭐',
  '🌊 _Mengalir seperti air, kuat seperti karang!_ 💪',
  '🌞 _Cerahkan hari orang lain dengan kebaikanmu!_ 😊',
  '🎵 _Hidup ini lagu, mainkan dengan indah!_ 🎶',
];

const TANYA_KABAR = [
  'Bagaimana kabarnya hari ini? 😊',
  'Apa kabar? Semoga sehat selalu ya 💪',
  'Gimana kabarnya? Semoga baik-baik saja 🌸',
  'Apa kabar hari ini? Semoga selalu dalam keadaan baik 🙏',
  'Bagaimana kabar? Semoga harimu menyenangkan 😊',
  'Kabar baik kah hari ini? 🌟',
  'Lagi sibuk apa nih hari ini? 💼',
  'Bagaimana hari ini? Lancar semua? ✨',
  'Apa kabar? Semoga selalu diberi kesehatan 🙏',
  'Gimana harinya? Semoga produktif 🚀',
];

function cocokKata(low, kata) {
  return low === kata || low.startsWith(kata + ' ') || low.startsWith(kata + ',') ||
    low.startsWith(kata + '!') || low.startsWith(kata + '.') || low.endsWith(' ' + kata);
}

function isSapaan(low) {
  const sorted = KATA_SAPAAN_LIST.slice().sort(function(a, b) { return b.length - a.length; });
  for (let i = 0; i < sorted.length; i++) {
    if (cocokKata(low, sorted[i])) return sorted[i];
  }
  return null;
}

const KATA_TERIMAKASIH = ['terima kasih','terimakasih','makasih','thanks','thank you','thx','tq','ty','tengkyu','mksh','trims','trimakasih'];
function isTerimakasih(low) { return KATA_TERIMAKASIH.some(function(k) { return cocokKata(low, k); }); }

function balasSapaanPintar(sender, kataSapaan) {
  const nama  = getNama(sender);
  const waktu = getWaktu();
  const dataSapa = SAPAAN_MAP[kataSapaan];
  if (!dataSapa) return (nama ? 'Halo, *' + nama + '*! 😊' : 'Halo! 😊');
  const opsiBalasan = dataSapa.balasan;
  const pilihan = opsiBalasan[Math.floor(Math.random() * opsiBalasan.length)];
  let respon = '';
  if (dataSapa.kategori === 'waktu') {
    if (dataSapa.waktu === waktu) respon = pilihan + (nama ? ', *' + nama + '*' : '') + '! 😊';
    else respon = pilihan + (nama ? ', *' + nama + '*' : '') + '! 😊\n_(Sekarang udah ' + waktu.toLowerCase() + ' loh)_';
  } else if (dataSapa.kategori === 'islami') respon = pilihan + (nama ? ', *' + nama + '*' : '') + ' 🤲';
  else if (dataSapa.kategori === 'english') respon = pilihan + (nama ? ', *' + nama + '*' : '') + '! 😊';
  else if (dataSapa.kategori === 'sopan')   respon = pilihan + (nama ? ', *' + nama + '*' : '') + ' 😊';
  else                                       respon = pilihan + (nama ? ', *' + nama + '*' : '') + '! 😊';
  return respon;
}

function sapaanPertama(sender) {
  const nama   = getNama(sender);
  const waktu  = getWaktu();
  const motiv  = KALIMAT_MOTIVASI[Math.floor(Math.random() * KALIMAT_MOTIVASI.length)];
  const kabar  = TANYA_KABAR[Math.floor(Math.random() * TANYA_KABAR.length)];
  const sambutan = ['Selamat ' + waktu, 'Halo, selamat ' + waktu, 'Hai, selamat ' + waktu, 'Halo', 'Hai'];
  const sapa = sambutan[Math.floor(Math.random() * sambutan.length)];
  return sapa + (nama ? ', *' + nama + '*' : '') + '! 👋\n\n' + motiv + '\n\n' + kabar;
}

function sapaanBerikutnya(sender, kataSapaan) {
  const balasan = balasSapaanPintar(sender, kataSapaan);
  const rand = Math.random();
  let tambahan = '';
  if (rand < 0.15) tambahan = '\n\n' + KALIMAT_MOTIVASI[Math.floor(Math.random() * KALIMAT_MOTIVASI.length)];
  else if (rand < 0.30) tambahan = '\n\n' + TANYA_KABAR[Math.floor(Math.random() * TANYA_KABAR.length)];
  return balasan + tambahan;
}

function balasTerimakasih(sender) {
  const nama = getNama(sender);
  const n = nama ? ', *' + nama + '*' : '';
  const opsi = [
    'Sama-sama' + n + '! 😊', 'Dengan senang hati' + n + '! 😊', 'Tentu' + n + '! 😊',
    'Iya, sama-sama' + n + '! 🙏', 'Senang bisa membantu' + n + '! ✨',
    'You\'re welcome' + n + '! 😊', 'Yuk, kalau ada yang lain ketik *menu* ya' + n + '! 💪',
  ];
  return opsi[Math.floor(Math.random() * opsi.length)];
}

function isAdminCommand(low) {
  if (['listmember','listkontak','reload','info'].indexOf(low) >= 0) return true;
  return ADMIN_COMMANDS.some(function(cmd) { return low.startsWith(cmd + ' '); });
}

// ════════════════════════════════════════════════════════════════
//   9. MENU FRIENDLY
// ════════════════════════════════════════════════════════════════

function getMenuUtama(nomor) {
  const nama = getNama(nomor);
  const salam = nama ? '*' + nama + '*' : 'Kamu';
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  🤖 *BOT TOKO PERABOT*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += 'Halo ' + salam + '! 👋\nSilakan pilih menu:\n\n';
  m += '┌─────────────────────\n';
  m += '│ ' + emojiNum(1) + ' 📊 Laporan Penjualan\n';
  m += '│ ' + emojiNum(2) + ' 🏷️ Laporan Harga Barang\n';
  m += '│ ' + emojiNum(3) + ' 🛒 Laporan Marketplace\n';
  if (isMember(nomor)) m += '│ ' + emojiNum(4) + ' 🔍 Cari Harga Barang\n';
  if (isAdmin(nomor))  m += '│ ' + emojiNum(9) + ' 👑 Menu Admin\n';
  m += '└─────────────────────\n\n';
  m += '💬 *Cara pilih:*\n   Ketik nomor (contoh: *1*)\n   atau ketik nama menunya\n\n';
  if (isMember(nomor)) {
    m += '🤖 *Atau tanya langsung ke AI:*\n';
    m += '   _"Stok dandang eagle 20 di NK?"_\n';
    m += '   _"Harga termurah panci?"_\n';
    m += '   _"Bandingkan rice cooker di toko"_';
  }
  return m;
}
function getMenuPilihToko(menuType) {
  const ic = menuType === 1 ? '📊' : menuType === 2 ? '🏷️' : menuType === 'cari' ? '🔍' : '🛒';
  const jd = menuType === 1 ? 'LAPORAN PENJUALAN'
           : menuType === 2 ? 'LAPORAN HARGA BARANG'
           : menuType === 'cari' ? 'CARI HARGA BARANG'
           : 'LAPORAN MARKETPLACE';
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  ' + ic + ' *' + jd + '*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🏦 *Pilih Toko:*\n\n┌─────────────────────\n';
  TOKO_LIST.forEach(function(t, i) { m += '│ ' + emojiNum(i + 1) + ' ' + t.nama + '\n'; });
  m += '└─────────────────────\n\n💬 Ketik nomor (1-5) atau nama toko\n   contoh: *1* atau *nk*\n\n🔙 Ketik *batal* untuk kembali';
  return m;
}

function getMenuPilihHari(namaToko) {
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  🏦 *' + namaToko + '*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '📅 *Laporan untuk hari:*\n\n┌─────────────────────\n';
  m += '│ ' + emojiNum(1) + ' 📅 *HARI INI*\n│    _(' + getTanggal(false) + ')_\n│\n';
  m += '│ ' + emojiNum(2) + ' 📅 *KEMARIN*\n│    _(' + getTanggal(true) + ')_\n';
  m += '└─────────────────────\n\n💬 Ketik *1* untuk hari ini\n   atau *2* untuk kemarin\n\n🔙 Ketik *batal* untuk kembali';
  return m;
}

function getMenuSiapInputMarket(namaToko, kemarin) {
  const t = getTanggal(kemarin);
  const k = kemarin ? ' _(kemarin)_' : '';
  const contoh = 'oesapa 0\ntdm 0\ncentral 21061000\nwa 21061000\nshopee 0\ntiktok 0\ntokopedia 0\ntunai 304000\ndebit 20757000';
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  ✅ *SIAP INPUT*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🏦 ' + namaToko + '\n📅 ' + t + k + '\n\n';
  m += '📸 *Kirim FOTO* atau *ketik manual:*\n\n```' + contoh + '```\n\n';
  m += '🔙 Ketik *batal* untuk membatalkan';
  return m;
}

function getMenuSiapInputHarga(namaToko, kemarin) {
  const t = getTanggal(kemarin);
  const k = kemarin ? ' _(kemarin)_' : '';
  const contoh = '---baru---\nNama barang baru\n---naik---\nNama barang naik\n---turun---\nNama barang turun';
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  ✅ *SIAP INPUT*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🏦 ' + namaToko + '\n📅 ' + t + k + '\n\n';
  m += '📸 *Kirim FOTO* atau *ketik manual:*\n\n```' + contoh + '```\n\n';
  m += '🔙 Ketik *batal* untuk membatalkan';
  return m;
}

function getMenuAdmin() {
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  👑 *MENU ADMIN*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🛠️ *Pilih aksi:*\n\n┌─────────────────────\n';
  m += '│ ' + emojiNum(1) + ' 👥 List Member\n';
  m += '│ ' + emojiNum(2) + ' 📒 List Kontak\n';
  m += '│ ' + emojiNum(3) + ' ➕ Tambah Member\n';
  m += '│ ' + emojiNum(4) + ' ➖ Hapus Member\n';
  m += '│ ' + emojiNum(5) + ' ✏️ Set Nama Kontak\n';
  m += '│ ' + emojiNum(6) + ' 🗑️ Hapus Kontak\n';
  m += '│ ' + emojiNum(7) + ' 🔄 Reload Excel\n';
  m += '│ ' + emojiNum(8) + ' ℹ️ Info Sistem\n';
  m += '└─────────────────────\n\n💬 Ketik nomor (1-8)\n\n🔙 Ketik *batal* untuk kembali';
  return m;
}

function getMenuSiapCari(namaToko) {
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  🔍 *CARI BARANG*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🏦 ' + namaToko + '\n\n━━━━━━━━━━━━━━━━━━\n⌨️ *Ketik nama atau kode:*\n━━━━━━━━━━━━━━━━━━\n\n';
  m += '💡 *Contoh:*\n   • _dandang eagle 20_\n   • _NN00001_\n   • _golden sunkist_\n\n';
  m += '✨ Bot otomatis koreksi typo!\n\n🔙 Ketik *batal* untuk kembali';
  return m;
}

function getMenuCariUlang(namaToko) {
  let m = '━━━━━━━━━━━━━━━━━━\n🔍 *Cari lagi di ' + namaToko + '?*\n━━━━━━━━━━━━━━━━━━\n\n';
  m += '⌨️ Ketik nama/kode barang lain\n\n┌─────────────────────\n';
  m += '│ ' + emojiNum(9) + ' 🔄 Ganti toko\n│ *batal* 🔙 Menu utama\n└─────────────────────';
  return m;
}

// ════════════════════════════════════════════════════════════════
//   10. ★★★ WIZARD MODE ★★★
// ════════════════════════════════════════════════════════════════

function wizardGetNextField(tokoKode, dataWizard) {
  const fields = FIELD_LAPORAN[tokoKode];
  if (!fields) return null;
  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (dataWizard[f.key] === undefined) return { field: f, index: i, total: fields.length };
  }
  return null;
}

function wizardTanyaField(tokoKode, fieldInfo, namaToko, dataWizard) {
  const f = fieldInfo.field;
  const no = fieldInfo.index + 1;
  const total = fieldInfo.total;
  
  let m = '🤖 *INPUT LAPORAN*\n';
  m += '🏦 ' + namaToko + '\n';
  m += '━━━━━━━━━━━━━━━━━━\n';
  m += '📊 *Step ' + no + ' dari ' + total + '*\n';
  
  const progress = Math.floor((no / total) * 10);
  m += '[';
  for (let i = 0; i < 10; i++) m += i < progress ? '█' : '░';
  m += '] ' + Math.round((no / total) * 100) + '%\n';
  m += '━━━━━━━━━━━━━━━━━━\n\n';
  
  // Tampilkan label dengan nama kasir otomatis (untuk CP)
  let labelLengkap = f.label;
  if (tokoKode === 'cp' && KASIR_CP_DEFAULT[f.key]) {
    labelLengkap = f.label + ' (' + KASIR_CP_DEFAULT[f.key] + ')';
  }
  
  // Field "total" auto
  if (f.key === 'total' && f.auto) {
    let totalCalc = 0;
    Object.keys(dataWizard).forEach(function(k) {
      if (k.startsWith('k') && k.length <= 2 && !k.startsWith('kr')) {
        totalCalc += (parseInt(dataWizard[k]) || 0);
      }
    });
    const formatTotal = totalCalc === 0 ? 'Rp. -' : 'Rp. ' + totalCalc.toLocaleString('id-ID');
    m += f.emoji + ' *' + labelLengkap + '*\n\n';
    m += '📋 Total dari kassa: *' + formatTotal + '*\n\n';
    m += '💰 *Masukkan total penjualan keseluruhan:*\n\n';
    m += '💡 Ketik angka atau tekan:\n';
    m += '   • _angka_ jika total berbeda\n';
    m += '   • *ok* untuk pakai total kassa otomatis\n\n';
  }
  // Tanya nominal biasa
  else {
    m += f.emoji + ' *' + labelLengkap + '*\n\n';
    m += '💰 *Berapa nominalnya?*\n\n';
    m += '💡 Ketik angka saja:\n';
    m += '   • _29812000_\n';
    m += '   • _29.812.000_\n';
    m += '   • _-_ atau *0* jika kosong\n\n';
  }
  
  m += '━━━━━━━━━━━━━━━━━━\n';
  m += '🔙 *batal* | ⏭️ *skip* | 👁️ *review*';
  return m;
}

function wizardReview(tokoKode, dataWizard, namaToko) {
  const fields = FIELD_LAPORAN[tokoKode];
  let m = '👁️ *REVIEW DATA SEMENTARA*\n';
  m += '🏦 ' + namaToko + '\n━━━━━━━━━━━━━━━━━━\n\n';
  
  fields.forEach(function(f) {
    const v = dataWizard[f.key];
    let labelLengkap = f.label;
    if (tokoKode === 'cp' && KASIR_CP_DEFAULT[f.key]) {
      labelLengkap = f.label + ' (' + KASIR_CP_DEFAULT[f.key] + ')';
    }
    if (v === undefined) {
      m += '⬜ ' + labelLengkap + ': _(belum diisi)_\n';
    } else if (v === 0) {
      m += '✅ ' + labelLengkap + ': Rp. -\n';
    } else {
      m += '✅ ' + labelLengkap + ': Rp. ' + parseInt(v).toLocaleString('id-ID') + '\n';
    }
  });
  
  m += '\n━━━━━━━━━━━━━━━━━━\n';
  m += '💬 *lanjut* — lanjut input\n';
  m += '💬 *selesai* — generate laporan\n';
  m += '🔙 *batal* — batalkan';
  return m;
}

function wizardToText(dataWizard, tokoKode) {
  let text = '';
  Object.keys(dataWizard).forEach(function(key) {
    if (key.startsWith('nama_')) return;
    if (key.startsWith('_')) return;
    if (key === 'total') return; // skip "total" karena dihitung otomatis di generator
    const v = dataWizard[key];
    
    // Untuk CP: auto-tambahkan nama kasir
    if (tokoKode === 'cp' && KASIR_CP_DEFAULT[key]) {
      text += key + ' ' + KASIR_CP_DEFAULT[key] + ' ' + (v || 0) + '\n';
    } else {
      text += key + ' ' + (v || 0) + '\n';
    }
  });
  return text;
}

// ════════════════════════════════════════════════════════════════
//   11. PARSER PINTAR
// ════════════════════════════════════════════════════════════════

function parsePilihanMenu(low) {
  if (low === '1' || low === 'satu') return 1;
  if (low === '2' || low === 'dua')  return 2;
  if (low === '3' || low === 'tiga') return 3;
  if (low === '4' || low === 'empat') return 4;
  if (low === '9' || low === 'admin') return 9;
  if (low.indexOf('penjualan') >= 0) return 1;
  if (low.indexOf('harga') >= 0 && low.indexOf('barang') >= 0) return 2;
  if (low.indexOf('marketplace') >= 0 || low.indexOf('market') >= 0) return 3;
  if (low.indexOf('cari') >= 0) return 4;
  if (low.indexOf('admin') >= 0) return 9;
  return null;
}

function parsePilihanToko(low) {
  const num = parseInt(low);
  if (num >= 1 && num <= TOKO_LIST.length) return TOKO_LIST[num - 1];
  for (let i = 0; i < TOKO_LIST.length; i++) {
    const t = TOKO_LIST[i];
    if (low === t.kode) return t;
    if (low === t.nama.toLowerCase()) return t;
    for (let j = 0; j < t.alias.length; j++) {
      if (low === t.alias[j] || low.indexOf(t.alias[j]) >= 0) return t;
    }
  }
  return null;
}

function parsePilihanHari(low) {
  if (low === '1' || low.indexOf('hari ini') >= 0 || low === 'sekarang' || low === 'today') return false;
  if (low === '2' || low.indexOf('kemarin') >= 0 || low === 'yesterday') return true;
  return null;
}

function parsePilihanAdmin(low) {
  const map = { '1': 'listmember', '2': 'listkontak', '3': 'daftar', '4': 'hapus', '5': 'namakontak', '6': 'hapuskontak', '7': 'reload', '8': 'info' };
  if (map[low]) return map[low];
  if (low.indexOf('list member') >= 0 || low === 'listmember') return 'listmember';
  if (low.indexOf('list kontak') >= 0 || low === 'listkontak') return 'listkontak';
  if (low.indexOf('tambah') >= 0 || low === 'daftar') return 'daftar';
  if (low === 'hapus member' || low === 'hapus') return 'hapus';
  if (low.indexOf('set nama') >= 0 || low === 'namakontak') return 'namakontak';
  if (low.indexOf('hapus kontak') >= 0) return 'hapuskontak';
  if (low === 'reload') return 'reload';
  if (low === 'info') return 'info';
  return null;
}
// ════════════════════════════════════════════════════════════════
//   12. FORMAT HASIL CARI
// ════════════════════════════════════════════════════════════════

function formatHasil(searchResult, tokoKode, sender) {
  const namaToko = NAMA_TOKO[tokoKode] || '-';
  const nama     = getNama(sender);
  const suffix   = '\n\n' + (nama ? 'Semoga membantu, *' + nama + '*! 😊' : 'Semoga membantu! 😊');
  const items          = searchResult.hasil     || [];
  const saran          = searchResult.saran     || [];
  const tipeHasil      = searchResult.tipeHasil || 'kosong';
  const totalDitemukan = searchResult.totalDitemukan || items.length;

  if (tipeHasil === 'kosong') {
    return '❌ *Barang Tidak Ditemukan*\n' + GARIS_TEBAL +
      '\n\n💡 *Tips:*\n• Gunakan kata kunci lebih singkat\n• Cari by kode (contoh: _NN00001_)\n• Periksa ejaan kata kunci';
  }
  if (tipeHasil === 'saran' && items.length === 0 && saran.length > 0) {
    let msg = '🤔 *Tidak Ditemukan Persis*\n' + GARIS_TEBAL + '\n🏦 *' + namaToko + '*\n\n💡 Mungkin yang kamu cari:\n\n';
    saran.forEach(function(d, i) {
      const h = d.harga[tokoKode];
      msg += '*' + (i + 1) + '.* ' + d.nama + '\n   🔖 _' + d.kode + '_\n   💰 ' + fRp(h.ecer) + '\n';
      if (i < saran.length - 1) msg += GARIS_TIPIS + '\n';
    });
    msg += '\n💡 Ketik *kode* untuk detail. Contoh: _' + saran[0].kode + '_';
    return msg + suffix;
  }
  if (items.length === 1) {
    const d = items[0]; const h = d.harga[tokoKode];
    let msg = '🏷️ *Detail Barang*\n🏦 *' + namaToko + '*\n' + GARIS_TEBAL + '\n' +
      '🔖 *Kode*   : ' + d.kode + '\n📦 *Nama*   : ' + d.nama + '\n' +
      '🏷️ *Jenis*  : ' + (d.jenis || '-') + '\n🏗️ *Merek*  : ' + (d.merek || '-') + '\n' +
      '📏 *Satuan* : ' + d.satuan + '\n' + GARIS_TEBAL + '\n' +
      '💰 *Harga Ecer*  : ' + fRp(h.ecer) + '\n💰 *Harga Ambil* : ' + fRp(h.ambil) + '\n' +
      '📊 *Stok*        : ' + (h.stok > 0 ? h.stok + ' ' + d.satuan : '⚠️ Kosong') + '\n' + GARIS_TEBAL;
    if (tipeHasil === 'fuzzy') msg += '\n\n💡 _Hasil koreksi otomatis_';
    return msg + suffix;
  }
  let header = totalDitemukan > CONFIG.maxHasilCari
    ? '🔍 *' + items.length + ' dari ' + totalDitemukan + ' Barang*\n'
    : '🔍 *Ditemukan ' + items.length + ' Barang*\n';
  let msg = header + '🏦 *' + namaToko + '*\n' + GARIS_TEBAL + '\n';
  if (tipeHasil === 'fuzzy') msg += '💡 _Hasil koreksi otomatis_\n\n';
  items.forEach(function(d, i) {
    const h = d.harga[tokoKode];
    msg += '*' + (i + 1) + '.* ' + d.nama + '\n   🔖 ' + d.kode + ' | ' + d.satuan + '\n';
    msg += '   💰 Ecer: ' + fRp(h.ecer) + ' | Ambil: ' + fRp(h.ambil) + '\n';
    msg += '   📊 Stok: ' + (h.stok > 0 ? h.stok + ' ' + d.satuan : '⚠️ Kosong') + '\n';
    if (i < items.length - 1) msg += GARIS_TIPIS + '\n';
  });
  msg += GARIS_TEBAL;
  if (totalDitemukan > CONFIG.maxHasilCari) msg += '\n\n⚠️ Masih ada *' + (totalDitemukan - CONFIG.maxHasilCari) + '* barang lagi.';
  return msg + suffix;
}

// ════════════════════════════════════════════════════════════════
//   13. GENERATOR LAPORAN
// ════════════════════════════════════════════════════════════════

function genLapPenjualan(text, namaToko, kemarin, tokoKode) {
  if (tokoKode === 'cp') return genLapPenjualanCP(text, kemarin);
  if (tokoKode === 'nk') return genLapPenjualanStandar(text, kemarin, 'Nasional Kitchen', 'fullEcer');
  if (tokoKode === 'tdm') return genLapPenjualanStandar(text, kemarin, 'Perabot Mama TDM', 'noEcer');
  if (tokoKode === 'oesapa') return genLapPenjualanStandar(text, kemarin, 'Perabot Mama Oesapa', 'noEcer');
  if (tokoKode === 'kefa') return genLapPenjualanStandar(text, kemarin, 'Perabot Mamaku Kefamenanu', 'noEcer');
  
  const t = getTanggal(kemarin);
  const k = kemarin ? ' _(kemarin)_' : '';
  const d = {};
  text.trim().toLowerCase().split('\n').forEach(function(l) {
    const p = l.trim().split(/\s+/);
    if (p.length >= 2) {
      const raw = p.slice(1).join(' ').trim();
      if (raw === '-' || raw === 'kosong' || raw === 'null') { d[p[0]] = 0; return; }
      const v = p[1].replace(/[^0-9]/g, '');
      if (v) d[p[0]] = parseFloat(v);
    }
  });
  const k1 = d.k1 || 0, k2 = d.k2 || 0, k3 = d.k3 || 0;
  const tot = k1 + k2 + k3;
  let ks = '';
  if (k1) ks += '• Kassa 1 : ' + fRpP(k1) + '\n';
  if (k2) ks += '• Kassa 2 : ' + fRpP(k2) + '\n';
  if (k3) ks += '• Kassa 3 : ' + fRpP(k3) + '\n';
  if (!ks) ks = '• -\n';
  return GARIS_TEBAL + '\n📊 *LAPORAN PENJUALAN*\n🏦 *Toko ' + namaToko + '*\n' + GARIS_TEBAL + '\n' +
    '📅 *' + t + '*' + k + '\n\n💵 *PENJUALAN PER KASSA*\n' + ks + '\n' +
    '📦 *TOTAL KESELURUHAN*\n' + fRpP(tot) + '\n\n💳 *METODE PEMBAYARAN*\n' +
    '• Tunai  : ' + fRpP(d.tunai || 0) + '\n• Debit  : ' + fRpP(d.debit || 0) + '\n• Kredit : ' + fRpP(d.kredit || 0) + '\n\n' +
    '🛒 *JENIS PENJUALAN*\n• Ecer   : ' + fRpP(d.ecer || 0) + '\n• Grosir : ' + fRpP(d.grosir || 0) + '\n' +
    GARIS_TEBAL + '\n_Laporan otomatis_';
}

function genLapPenjualanStandar(text, kemarin, namaTokoFull, mode) {
  const t = getTanggal(kemarin);
  const d = { k1: 0, k2: 0, k3: 0, tunai: 0, debit: 0, kredit: 0, ecer: 0, grosir: 0 };
  
  text.trim().split('\n').forEach(function(line) {
    const tr = line.trim();
    if (!tr) return;
    const parts = tr.split(/\s+/);
    if (parts.length < 2) return;
    const key = parts[0].toLowerCase();
    const rawValue = parts.slice(1).join(' ').trim();
    let value = 0;
    if (rawValue !== '-' && rawValue !== 'kosong' && rawValue !== 'null') {
      const valueStr = rawValue.replace(/[^0-9]/g, '');
      value = parseFloat(valueStr) || 0;
    }
    if      (key === 'k1') d.k1 = value;
    else if (key === 'k2') d.k2 = value;
    else if (key === 'k3') d.k3 = value;
    else if (key === 'tunai') d.tunai = value;
    else if (key === 'debit') d.debit = value;
    else if (key === 'kredit' || key === 'credit') d.kredit = value;
    else if (key === 'ecer') d.ecer = value;
    else if (key === 'grosir') d.grosir = value;
  });
  
  function fr(n) {
    const v = parseFloat(n) || 0;
    if (v === 0) return 'Rp. -';
    return 'Rp. ' + v.toLocaleString('id-ID');
  }
  
  const totalUtama = d.k1 + d.k2 + d.k3;
  
  let msg = 'Laporan Penjualan\n';
  msg += 'Toko ' + namaTokoFull + '\n';
  msg += 'Periode ' + t + '\n\n';
  msg += 'Kassa 1 ' + fr(d.k1) + '\n';
  msg += 'Kassa 2 ' + fr(d.k2) + '\n';
  if (d.k3 > 0) msg += 'Kassa 3 ' + fr(d.k3) + '\n';
  msg += '\n';
  msg += 'Total Penjualan Keseluruhan\n';
  msg += fr(totalUtama) + '\n';
  msg += '---------------------------------------------\n\n';
  msg += 'Tunai  ' + fr(d.tunai) + '\n';
  msg += 'Debit  ' + fr(d.debit) + '\n';
  msg += 'Credit ' + fr(d.kredit) + '\n';
  msg += '---------------------------------------------\n';
  if (mode === 'fullEcer' || d.ecer > 0 || d.grosir > 0) {
    msg += 'Ecer : ' + fr(d.ecer) + '\n';
    msg += 'Grosir : ' + fr(d.grosir) + '\n';
    msg += '--------------------------------------------';
  }
  return msg;
}

function genLapPenjualanCP(text, kemarin) {
  const t = getTanggal(kemarin);
  const d = {
    k1: 0, k2: 0, k3: 0, k4: 0,
    nk1: '', nk2: '', nk3: '', nk4: '',
    tunai: 0, debit: 0, kredit: 0,
    ecer: 0, grosir: 0,
    promo: 0, promoTunai: 0, promoDebit: 0, promoKredit: 0,
    parkirKomputer: 0, parkirLuar: 0,
  };
  
  text.trim().split('\n').forEach(function(line) {
    const tr = line.trim();
    if (!tr) return;
    const parts = tr.split(/\s+/);
    if (parts.length < 2) return;
    const key = parts[0].toLowerCase();
    let valueStart = 1;
    let kasirNama = '';
    if (['k1','k2','k3','k4'].indexOf(key) >= 0 && parts.length >= 3) {
      const possibleNum = parts[1].replace(/[^0-9]/g, '');
      if (!possibleNum || possibleNum.length < 4) {
        for (let i = 1; i < parts.length; i++) {
          const num = parts[i].replace(/[^0-9]/g, '');
          if (num && num.length >= 4) {
            valueStart = i;
            kasirNama = parts.slice(1, i).join(' ');
            break;
          }
        }
      }
    }
    const rawValue = parts.slice(valueStart).join(' ').trim();
    let value = 0;
    if (rawValue !== '-' && rawValue !== 'kosong' && rawValue !== 'null') {
      const valueStr = rawValue.replace(/[^0-9]/g, '');
      value = parseFloat(valueStr) || 0;
    }
    if      (key === 'k1') { d.k1 = value; if (kasirNama) d.nk1 = kasirNama; }
    else if (key === 'k2') { d.k2 = value; if (kasirNama) d.nk2 = kasirNama; }
    else if (key === 'k3') { d.k3 = value; if (kasirNama) d.nk3 = kasirNama; }
    else if (key === 'k4') { d.k4 = value; if (kasirNama) d.nk4 = kasirNama; }
    else if (key === 'tunai') d.tunai = value;
    else if (key === 'debit') d.debit = value;
    else if (key === 'kredit' || key === 'credit') d.kredit = value;
    else if (key === 'ecer') d.ecer = value;
    else if (key === 'grosir') d.grosir = value;
    else if (key === 'promo') d.promo = value;
    else if (key === 'promotunai' || key === 'ptunai') d.promoTunai = value;
    else if (key === 'promodebit' || key === 'pdebit') d.promoDebit = value;
    else if (key === 'promokredit' || key === 'pkredit') d.promoKredit = value;
    else if (key === 'parkirkomputer' || key === 'pkomp') d.parkirKomputer = value;
    else if (key === 'parkirluar' || key === 'pluar') d.parkirLuar = value;
  });
  
  function fr(n) {
    const v = parseFloat(n) || 0;
    if (v === 0) return 'Rp. -';
    return 'Rp. ' + v.toLocaleString('id-ID');
  }
  
  // Pastikan nama kasir terisi (fallback ke default)
  if (!d.nk1) d.nk1 = KASIR_CP_DEFAULT.k1;
  if (!d.nk2) d.nk2 = KASIR_CP_DEFAULT.k2;
  if (!d.nk3) d.nk3 = KASIR_CP_DEFAULT.k3;
  if (!d.nk4) d.nk4 = KASIR_CP_DEFAULT.k4;
  
  const totalUtama = d.k1 + d.k2 + d.k3 + d.k4;
  const totalParkir = d.parkirKomputer + d.parkirLuar;
  
  let msg = 'Laporan Penjualan Toko Central Perabot Periode ' + t + '\n\n';
  msg += 'Kassa 1 (' + d.nk1 + ') ' + fr(d.k1) + '\n';
  msg += 'Kassa 2 (' + d.nk2 + ') ' + fr(d.k2) + '\n';
  msg += 'Kassa 3 (' + d.nk3 + ') ' + fr(d.k3) + '\n';
  msg += 'Kassa 4 (' + d.nk4 + ') ' + fr(d.k4) + '\n\n';
  msg += 'Total Penjualan Keseluruhan: ' + fr(totalUtama) + '\n';
  msg += '---------------------------------------------\n\n';
  msg += 'Tunai  ' + fr(d.tunai) + '\n';
  msg += 'Debit  ' + fr(d.debit) + '\n';
  msg += 'Credit ' + fr(d.kredit) + '\n';
  msg += '---------------------------------------------\n';
  msg += '---------------------------------------------\n\n';
  msg += 'Ecer: ' + fr(d.ecer) + '\n';
  msg += 'Grosir : ' + fr(d.grosir) + '\n\n';
  msg += '---------------------------------------------\n\n';
  msg += 'Laporan Penjualan Kasir Promo\n';
  msg += 'Periode ' + t + '\n\n';
  msg += 'Total Penjualan Keseluruhan: ' + fr(d.promo) + '\n';
  msg += '---------------------------------------------\n';
  msg += 'Tunai  ' + fr(d.promoTunai) + '\n';
  msg += 'Debit  ' + fr(d.promoDebit) + '\n';
  msg += 'Credit ' + fr(d.promoKredit) + '\n';
  msg += '---------------------------------------------\n\n';
  msg += 'Laporan Parkir \n';
  msg += 'Periode ' + t + '\n\n';
  msg += 'Parkir di Komputer : ' + fr(d.parkirKomputer) + '\n';
  msg += 'Parkir Stor Luar : ' + fr(d.parkirLuar) + '\n';
  msg += '---------------------------------------------\n';
  msg += 'Total Parkir  ' + fr(totalParkir) + '\n';
  msg += '---------------------------------------------';
  return msg;
}

function genLapHarga(text, namaToko, kemarin) {
  const t = getTanggal(kemarin);
  const h = kemarin ? 'Kemarin' : 'Ini';
  const k = kemarin ? ' _(kemarin)_' : '';
  const s = 'Selamat ' + getWaktu() + ' Team ' + namaToko;
  const d = { baru: [], naik: [], turun: [], note: [] };
  let mode = null;
  text.trim().split('\n').forEach(function(line) {
    const tr = line.trim(); if (!tr) return; const lo = tr.toLowerCase();
    if (lo.indexOf('---baru---') >= 0 || lo === 'baru')  { mode = 'baru';  return; }
    if (lo.indexOf('---naik---') >= 0 || lo === 'naik')  { mode = 'naik';  return; }
    if (lo.indexOf('---turun---') >= 0 || lo === 'turun') { mode = 'turun'; return; }
    if (lo.indexOf('---note---') >= 0 || lo === 'note')   { mode = 'note';  return; }
    if (mode) d[mode].push(tr);
  });
  const cat = 'Nota Semuanya Sudah Diinput Di Sistem, Bisa Langsung Di Print Barcodenya Ya.\n\nMohon Dicek Kembali Fisik Barang Dengan Yang Di Input Disistem, Jika Ada Yang Tidak Sesuai Mohon Di Konfirmasi Lagi. Terima Kasih🙏🏼';
  let msg = s + '\n\nHarga Barang Untuk Hari ' + h + ' *' + t + '*' + k + '\n';
  if (d.baru.length > 0)  { msg += '\n🆕 *Barang Yang Baru:*\n';        d.baru.forEach(function(b)  { msg += '• ' + b + '\n'; }); }
  if (d.naik.length > 0)  { msg += '\n📈 *Barang Yang Naik Harga:*\n';  d.naik.forEach(function(b)  { msg += '• ' + b + '\n'; }); }
  if (d.turun.length > 0) { msg += '\n📉 *Barang Yang Turun Harga:*\n'; d.turun.forEach(function(b) { msg += '• ' + b + '\n'; }); }
  if (d.note.length > 0)  { msg += '\n📝 *Catatan:*\n';                 d.note.forEach(function(b)  { msg += b + '\n'; }); }
  return msg + '\n' + cat;
}

function genLapMarket(text, kemarin) {
  const t = getTanggal(kemarin);
  const k = kemarin ? ' _(kemarin)_' : '';
  const d = { oesapa:0, tdm:0, central:0, wa:0, shopee:0, tiktok:0, tokopedia:0, tunai:0, debit:0, kredit:0, nota:[] };
  text.trim().toLowerCase().split('\n').forEach(function(line) {
    const tr = line.trim(); if (!tr) return;
    if (tr.indexOf('nota ') === 0) { d.nota.push(line.trim().substring(5)); return; }
    const p = tr.split(/\s+/);
    if (p.length >= 2 && p[0] in d) d[p[0]] = parseFloat(p.slice(1).join('').replace(/[^0-9]/g, '')) || 0;
  });
  const tT = d.oesapa + d.tdm + d.central;
  const tC = d.wa + d.shopee + d.tiktok + d.tokopedia;
  let nt = '';
  if (d.nota.length > 0) { nt = '\n'; d.nota.forEach(function(n) { nt += '- Nomor Nota ' + n + '\n'; }); }
  return GARIS_TEBAL + '\n🛒 *Total Penjualan Marketplace*\n*Perabot Mama*\n📅 Periode ' + t + k +
    '\n' + GARIS_TEBAL + '\n🏦 *Per Toko*\n' +
    '• Toko Perabot Mama Oesapa : ' + fRp(d.oesapa)  + '\n' +
    '• Toko Perabot Mama TDM    : ' + fRp(d.tdm)     + '\n' +
    '• Toko Central Perabot     : ' + fRp(d.central) + '\n' +
    GARIS_TIPIS + '\n💰 *Total* : ' + fRp(tT) + '\n\n📱 *Per Channel*\n' +
    '• WA        : ' + fRp(d.wa) + '\n• Shopee    : ' + fRp(d.shopee) + '\n' +
    '• Tiktok    : ' + fRp(d.tiktok) + '\n• Tokopedia : ' + fRp(d.tokopedia) + '\n' +
    GARIS_TIPIS + '\n💰 *Total Penjualan* : ' + fRp(tC) + '\n\n💳 *Metode Bayar*\n' +
    '• Tunai/CASH : ' + fRp(d.tunai) + '\n• Debit/TF   : ' + fRp(d.debit) + '\n• Credit     : ' + fRp(d.kredit) + '\n' +
    GARIS_TEBAL + '\n' + nt + '_Laporan otomatis_';
}

// ════════════════════════════════════════════════════════════════
//   14. HANDLER ADMIN
// ════════════════════════════════════════════════════════════════

async function handleAdmin(sender, msg, low) {
  log.info('ADMIN', 'Cmd: ' + low);
  if (low.startsWith('daftar ')) {
    const nomor = msg.substring(7).trim().replace(/[^0-9]/g, '');
    if (!nomor) { await kirimWA(sender, '⚠️ Format: daftar 6281234567890'); return true; }
    const r = tambahMember(nomor);
    if (!r.ok) { await kirimWA(sender, '⚠️ ' + r.alasan); return true; }
    await kirimWA(sender, '✅ *Member Terdaftar!*\n📱 ' + nomor + '\n👤 ' + (getNama(nomor) || '(belum ada)') +
      '\n📊 Total: ' + MEMBERS.length + '/' + CONFIG.maxMember);
    return true;
  }
  if (low.startsWith('hapus ')) {
    const nomor = msg.substring(6).trim().replace(/[^0-9]/g, '');
    const r = hapusMember(nomor);
    if (!r.ok) { await kirimWA(sender, '❌ ' + r.alasan); return true; }
    await kirimWA(sender, '✅ Member ' + nomor + ' dihapus!\n📊 Total: ' + MEMBERS.length + '/' + CONFIG.maxMember);
    return true;
  }
  if (low === 'listmember') {
    try {
      let m = 'Daftar Member (' + MEMBERS.length + '/' + CONFIG.maxMember + ')\n------------------\n';
      if (MEMBERS.length === 0) m += '(belum ada member)\n';
      else for (let i = 0; i < MEMBERS.length; i++) {
        m += (i + 1) + '. ' + MEMBERS[i] + '\n   ' + (KONTAK[MEMBERS[i]] || '(belum ada nama)') + '\n';
      }
      m += '------------------\nSlot tersisa: ' + (CONFIG.maxMember - MEMBERS.length);
      await kirimWA(sender, m);
    } catch (e) { await kirimWA(sender, 'Error: ' + e.message); }
    return true;
  }
  if (low.startsWith('namakontak ')) {
    const p = msg.substring(11).trim().split(/\s+/);
    if (p.length < 2) { await kirimWA(sender, '⚠️ Format: namakontak 628xxx Nama'); return true; }
    const nomor = p[0].replace(/[^0-9]/g, '');
    const nama  = p.slice(1).join(' ');
    const r = setNama(nomor, nama);
    if (!r.ok) { await kirimWA(sender, '❌ ' + r.alasan); return true; }
    await kirimWA(sender, '✅ Nama disimpan!\n📱 ' + nomor + ' → 👤 ' + nama);
    return true;
  }
  if (low.startsWith('hapuskontak ')) {
    const nomor = msg.substring(12).trim().replace(/[^0-9]/g, '');
    const r = hapusKontak(nomor);
    if (!r.ok) { await kirimWA(sender, '❌ ' + r.alasan); return true; }
    await kirimWA(sender, '✅ Kontak ' + nomor + ' dihapus!');
    return true;
  }
  if (low === 'listkontak') {
    const keys = Object.keys(KONTAK);
    let m = 'Daftar Kontak (' + keys.length + ')\n------------------\n';
    keys.forEach(function(k, i) { m += (i + 1) + '. ' + k + '\n   ' + KONTAK[k] + '\n'; });
    await kirimWA(sender, m);
    return true;
  }
  if (low === 'reload') {
    const ok = loadExcel();
    await kirimWA(sender, ok ? '✅ Excel reloaded! Total: ' + DATA_BARANG.length + ' item' : '❌ Gagal reload');
    return true;
  }
  if (low === 'info') {
    const up = Math.floor(process.uptime());
    const jam = Math.floor(up / 3600), mnt = Math.floor((up % 3600) / 60);
    await kirimWA(sender, 'ℹ️ Info Sistem\n------------------' +
      '\nUptime: ' + jam + 'j ' + mnt + 'm\nMember: ' + MEMBERS.length + '/' + CONFIG.maxMember +
      '\nData: ' + DATA_BARANG.length + ' item\nKontak: ' + Object.keys(KONTAK).length +
      '\nSesi: ' + Object.keys(SESI).length + '\nNode: ' + process.version);
    return true;
  }
  return false;
}

// ════════════════════════════════════════════════════════════════
//   15. ROUTES
// ════════════════════════════════════════════════════════════════

app.get('/', function(req, res) {
  res.json({ status: 'ok', app: CONFIG.appName + ' v3.6', items: DATA_BARANG.length, members: MEMBERS.length + '/' + CONFIG.maxMember });
});
app.get('/reload', function(req, res) {
  const ok = loadExcel();
  res.json({ success: ok, total: DATA_BARANG.length });
});
app.get('/resetsesi/:nomor', function(req, res) {
  resetSesi(req.params.nomor);
  res.json({ ok: true });
});

// ════════════════════════════════════════════════════════════════
//   16. WEBHOOK UTAMA
// ════════════════════════════════════════════════════════════════

const KATA_RESET = ['batal', 'menu', 'mulai', 'start', 'kembali', 'home'];

app.post('/webhook', async function(req, res) {
  res.sendStatus(200);

  try {
    const body    = req.body || {};
    const sender  = body.sender || body.from || body.phone || null;
    const message = (body.message || body.text || body.msg || '').trim();
    const image   = body.image || body.file || body.media || '';

    if (!sender) return;
    const msg = message;
    const low = msg.toLowerCase();

    log.info('WEBHOOK', sender + ': ' + msg.substring(0, 50));

    // ── SAPAAN PERTAMA ──
    if (!SUDAH_DISAPA[sender]) {
      SUDAH_DISAPA[sender] = true;
      saveJSON(CONFIG.paths.disapa, SUDAH_DISAPA);
      await kirimWA(sender, sapaanPertama(sender));
      await tunggu(1000);
      await kirimWA(sender, getMenuUtama(sender));
      return;
    }

    // ── ADMIN COMMAND ──
    if (isAdmin(sender) && isAdminCommand(low)) {
      if (SESI[sender] && (SESI[sender].mode || SESI[sender].menu || SESI[sender].adminAksi)) resetSesi(sender);
      const handled = await handleAdmin(sender, msg, low);
      if (handled) return;
    }

    // Cek state - apakah lagi di tahap input data laporan atau wizard
    const _s = SESI[sender] || {};
    const _lagiInput = (_s.menu && (_s.kemarin !== undefined && _s.kemarin !== null)) || _s.wizardActive;

    // ── RESET ──
    if (KATA_RESET.indexOf(low) >= 0 || (low === '0' && !_lagiInput)) {
      resetSesi(sender);
      await kirimWA(sender, getMenuUtama(sender));
      return;
    }

    // ── SAPAAN PINTAR ──
    const kataSapaan = isSapaan(low);
    if (kataSapaan && !_lagiInput) {
      await kirimWA(sender, sapaanBerikutnya(sender, kataSapaan));
      await tunggu(800);
      await kirimWA(sender, getMenuUtama(sender));
      return;
    }
    if (isTerimakasih(low) && !_lagiInput) {
      await kirimWA(sender, balasTerimakasih(sender));
      return;
    }

    if (isTerimakasih(low) && !_lagiInput) {
      await kirimWA(sender, balasTerimakasih(sender));
      return;
    }

        // ★★★ AI CHAT BARANG ★★★
    if (!_lagiInput && isMember(sender) && isPertanyaanBarang(low) && msg.length >= 5) {
      log.info('AI_CHAT', sender + ' tanya: ' + msg.substring(0, 80));
      await kirimWA(sender, '🤖 _Sedang berpikir..._');
      
      try {
        const jawaban = await aiChatBarang(msg, sender);
        if (jawaban) {
          await kirimWA(sender, '🤖 ' + jawaban);
        } else {
          // Fallback: pakai cari biasa
          log.warn('AI_CHAT', 'AI return null, fallback ke cari biasa');
          const hasil = cariBarang(msg);
          if (hasil.hasil && hasil.hasil.length > 0) {
            await kirimWA(sender, '🤖 Saya cari barang yang mirip ya:\n\n' + 
              'Coba pakai menu *4* (Cari Barang) untuk hasil lebih lengkap, atau lihat di bawah:');
            await tunggu(500);
            
            let resp = '📦 *Hasil Pencarian:*\n' + GARIS_TEBAL + '\n\n';
            hasil.hasil.slice(0, 5).forEach(function(d, i) {
              resp += '*' + (i + 1) + '.* ' + d.nama + '\n';
              resp += '   🔖 ' + d.kode + '\n';
              resp += '   💰 NK: ' + fRp(d.harga.nk.ecer) + ' | TDM: ' + fRp(d.harga.tdm.ecer) + '\n';
              resp += '   💰 Oesapa: ' + fRp(d.harga.oesapa.ecer) + ' | Kefa: ' + fRp(d.harga.kefa.ecer) + '\n';
              resp += '   💰 CP: ' + fRp(d.harga.cp.ecer) + '\n\n';
            });
            await kirimWA(sender, resp);
          } else if (hasil.saran && hasil.saran.length > 0) {
            let resp = '🤔 Maaf, tidak nemu yang persis. Mungkin yang dimaksud:\n\n';
            hasil.saran.slice(0, 5).forEach(function(d, i) {
              resp += '*' + (i + 1) + '.* ' + d.nama + ' (' + d.kode + ')\n';
            });
            resp += '\n💡 Coba ketik kata kunci lain atau kode barangnya.';
            await kirimWA(sender, resp);
          } else {
            await kirimWA(sender, '🤔 Maaf, saya tidak menemukan barang dengan kata kunci tersebut.\n\n' +
              '💡 *Coba:*\n' +
              '• Pakai kata kunci lebih singkat (contoh: "kompor" bukan "konpor hock 22")\n' +
              '• Cek ejaan barang\n' +
              '• Atau ketik *menu* → pilih *4* untuk cari manual');
          }
        }
      } catch (e) {
        log.error('AI_CHAT', 'Error: ' + e.message);
        await kirimWA(sender, '⚠️ Sistem AI sedang sibuk. Coba lagi sebentar atau ketik *menu*.');
      }
      return;
    }
    
    const s = getSesi(sender);
    
    // ── ADMIN MODE INPUT ──
    if (isAdmin(sender) && s.adminAksi) {
      const aksi = s.adminAksi;
      updateSesi(sender, { adminAksi: null });
      if (aksi === 'daftar')      return await handleAdmin(sender, 'daftar ' + msg, 'daftar ' + low);
      if (aksi === 'hapus')       return await handleAdmin(sender, 'hapus ' + msg, 'hapus ' + low);
      if (aksi === 'namakontak')  return await handleAdmin(sender, 'namakontak ' + msg, 'namakontak ' + low);
      if (aksi === 'hapuskontak') return await handleAdmin(sender, 'hapuskontak ' + msg, 'hapuskontak ' + low);
    }

    // ── MENU ADMIN ──
    if (s.mode === 'admin_menu') {
      const aksi = parsePilihanAdmin(low);
      if (aksi === 'listmember' || aksi === 'listkontak' || aksi === 'reload' || aksi === 'info') {
        resetSesi(sender);
        return await handleAdmin(sender, aksi, aksi);
      }
      if (aksi === 'daftar')      { updateSesi(sender, { adminAksi: 'daftar', mode: null });      await kirimWA(sender, '➕ *Tambah Member*\n' + GARIS_TEBAL + '\n\nKetik nomor HP:\n_6281234567890_\n\n🔙 Ketik *batal* untuk membatalkan'); return; }
      if (aksi === 'hapus')       { updateSesi(sender, { adminAksi: 'hapus', mode: null });       await kirimWA(sender, '➖ *Hapus Member*\n' + GARIS_TEBAL + '\n\nKetik nomor HP:\n_6281234567890_\n\n🔙 Ketik *batal* untuk membatalkan'); return; }
      if (aksi === 'namakontak')  { updateSesi(sender, { adminAksi: 'namakontak', mode: null });  await kirimWA(sender, '✏️ *Set Nama*\n' + GARIS_TEBAL + '\n\nFormat: nomor nama\n_6281234567890 Pak Budi_\n\n🔙 Ketik *batal* untuk membatalkan'); return; }
      if (aksi === 'hapuskontak') { updateSesi(sender, { adminAksi: 'hapuskontak', mode: null }); await kirimWA(sender, '🗑️ *Hapus Kontak*\n' + GARIS_TEBAL + '\n\nKetik nomor HP\n\n🔙 Ketik *batal* untuk membatalkan'); return; }
      await kirimWA(sender, getMenuAdmin());
      return;
    }

    // ── MODE CARI ──
    if (s.mode === 'cari') {
      if (!s.tokoKode) {
        if (low === '9') { updateSesi(sender, { tokoKode: null }); await kirimWA(sender, getMenuPilihToko('cari')); return; }
        const toko = parsePilihanToko(low);
        if (toko) {
          updateSesi(sender, { tokoKode: toko.kode });
          if (s.pendingKw) {
            const kw = s.pendingKw;
            updateSesi(sender, { pendingKw: null });
            await kirimWA(sender, formatHasil(cariBarang(kw), toko.kode, sender));
            await tunggu(800);
            await kirimWA(sender, getMenuCariUlang(toko.nama));
          } else {
            await kirimWA(sender, getMenuSiapCari(toko.nama));
          }
          return;
        }
        await kirimWA(sender, getMenuPilihToko('cari'));
        return;
      }
      if (low === '9' || low === 'ganti toko' || low === 'ganti') {
        updateSesi(sender, { tokoKode: null });
        await kirimWA(sender, getMenuPilihToko('cari'));
        return;
      }
      if (!msg) return;
      await kirimWA(sender, formatHasil(cariBarang(msg), s.tokoKode, sender));
      await tunggu(800);
      await kirimWA(sender, getMenuCariUlang(NAMA_TOKO[s.tokoKode]));
      return;
    }

    // ── LAPORAN: Pilih menu ──
    if (!s.menu && !s.mode) {
      const pilihan = parsePilihanMenu(low);
      if (pilihan === 1 || pilihan === 2) { updateSesi(sender, { menu: pilihan }); await kirimWA(sender, getMenuPilihToko(pilihan)); return; }
      if (pilihan === 3) { updateSesi(sender, { menu: 3 }); await kirimWA(sender, getMenuPilihHari('Marketplace Perabot Mama')); return; }
      if (pilihan === 4) {
        if (!isMember(sender)) { await kirimWA(sender, '🚫 *Akses Ditolak*\n\nFitur ini hanya untuk member.\nHubungi admin untuk mendaftar.'); return; }
        resetSesi(sender);
        updateSesi(sender, { mode: 'cari' });
        await kirimWA(sender, getMenuPilihToko('cari'));
        return;
      }
      if (pilihan === 9) {
        if (!isAdmin(sender)) { await kirimWA(sender, '🚫 Menu khusus admin.'); return; }
        updateSesi(sender, { mode: 'admin_menu' });
        await kirimWA(sender, getMenuAdmin());
        return;
      }
      await kirimWA(sender, '🤔 Maaf, saya tidak mengerti.\n\nKetik *menu* untuk lihat pilihan.');
      return;
    }

    // ── LAPORAN: Pilih toko ──
    if (s.menu !== 3 && !s.toko) {
      const toko = parsePilihanToko(low);
      if (toko) { updateSesi(sender, { toko: toko.kode }); await kirimWA(sender, getMenuPilihHari(toko.nama)); return; }
      await kirimWA(sender, getMenuPilihToko(s.menu));
      return;
    }

    // ── LAPORAN: Pilih hari ──
    if (s.kemarin === undefined || s.kemarin === null) {
      const kem = parsePilihanHari(low);
      if (kem === null) {
        const nm = s.menu === 3 ? 'Marketplace Perabot Mama' : NAMA_TOKO[s.toko];
        await kirimWA(sender, getMenuPilihHari(nm));
        return;
      }
      updateSesi(sender, { kemarin: kem });
      const nm = s.menu === 3 ? 'Marketplace Perabot Mama' : NAMA_TOKO[s.toko];
      
      // ★★★ MENU 1 (PENJUALAN) → LANGSUNG MASUK WIZARD ★★★
      if (s.menu === 1) {
        const t = getTanggal(kem);
        const kk = kem ? ' _(kemarin)_' : '';
        let intro = '╭━━━━━━━━━━━━━━━━━╮\n│  ✅ *MULAI INPUT*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
        intro += '🏦 ' + nm + '\n📅 ' + t + kk + '\n\n';
        intro += '📸 *Kirim FOTO* untuk dibaca AI\n';
        intro += '   _atau jawab pertanyaan di bawah_\n\n';
        intro += '━━━━━━━━━━━━━━━━━━';
        await kirimWA(sender, intro);
        await tunggu(800);
        
        updateSesi(sender, { wizardActive: true, wizardData: {} });
        const nextField = wizardGetNextField(s.toko, {});
        if (nextField) await kirimWA(sender, wizardTanyaField(s.toko, nextField, nm, {}));
        return;
      }
      
      // Menu 2: laporan harga
      if (s.menu === 2) {
        await kirimWA(sender, getMenuSiapInputHarga(nm, kem));
        return;
      }
      // Menu 3: laporan marketplace
      if (s.menu === 3) {
        await kirimWA(sender, getMenuSiapInputMarket(nm, kem));
        return;
      }
      return;
    }

    // ── LAPORAN: Input data ──
    const namaToko = s.menu === 3 ? 'Marketplace Perabot Mama' : NAMA_TOKO[s.toko];
    let laporan = '';

    // ★★★ WIZARD MODE - PROCESS INPUT ★★★
    if (s.menu === 1 && s.wizardActive) {
      const dataWizard = s.wizardData || {};
      
      // Review data
      if (low === 'review') {
        await kirimWA(sender, wizardReview(s.toko, dataWizard, namaToko));
        return;
      }
      
      // Selesai paksa
      if (low === 'selesai' || low === 'finish' || low === 'done') {
        const fields = FIELD_LAPORAN[s.toko];
        fields.forEach(function(f) { if (dataWizard[f.key] === undefined) dataWizard[f.key] = 0; });
        const text = wizardToText(dataWizard, s.toko);
        laporan = genLapPenjualan(text, namaToko, s.kemarin, s.toko);
        if (laporan) {
          await kirimWA(sender, laporan);
          resetSesi(sender);
          await tunggu(1500);
          await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
        }
        return;
      }
      
      // Lanjut setelah review
      if (low === 'lanjut' || low === 'next' || low === 'continue') {
        const nextField = wizardGetNextField(s.toko, dataWizard);
        if (nextField) await kirimWA(sender, wizardTanyaField(s.toko, nextField, namaToko, dataWizard));
        else {
          const text = wizardToText(dataWizard, s.toko);
          laporan = genLapPenjualan(text, namaToko, s.kemarin, s.toko);
          if (laporan) {
            await kirimWA(sender, laporan);
            resetSesi(sender);
            await tunggu(1500);
            await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
          }
        }
        return;
      }
      
      const currentField = wizardGetNextField(s.toko, dataWizard);
      if (!currentField) {
        const text = wizardToText(dataWizard, s.toko);
        laporan = genLapPenjualan(text, namaToko, s.kemarin, s.toko);
        if (laporan) {
          await kirimWA(sender, laporan);
          resetSesi(sender);
          await tunggu(1500);
          await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
        }
        return;
      }
      
      const f = currentField.field;
      
      // Skip field
      if (low === 'skip' || low === 'lewati') {
        dataWizard[f.key] = 0;
        updateSesi(sender, { wizardData: dataWizard });
        const next = wizardGetNextField(s.toko, dataWizard);
        if (next) {
          await tunggu(300);
          await kirimWA(sender, '⏭️ _Dilewati (diisi 0)_\n\n' + wizardTanyaField(s.toko, next, namaToko, dataWizard));
        } else {
          const text = wizardToText(dataWizard, s.toko);
          laporan = genLapPenjualan(text, namaToko, s.kemarin, s.toko);
          if (laporan) {
            await kirimWA(sender, laporan);
            resetSesi(sender);
            await tunggu(1500);
            await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
          }
        }
        return;
      }
      
      // Parse nominal
      let nominal = 0;
      
      // Untuk field "total" dengan auto → user bisa ketik "ok" untuk pakai total kassa
      if (f.key === 'total' && f.auto && (low === 'ok' || low === 'ya' || low === 'yes' || low === 'y' || low === 'otomatis' || low === 'auto')) {
        let totalCalc = 0;
        Object.keys(dataWizard).forEach(function(k) {
          if (k.startsWith('k') && k.length <= 2 && !k.startsWith('kr')) {
            totalCalc += (parseInt(dataWizard[k]) || 0);
          }
        });
        nominal = totalCalc;
      } else if (msg === '-' || low === 'kosong' || low === 'null') {
        nominal = 0;
      } else {
        const angka = msg.replace(/[^0-9]/g, '');
        if (!angka) {
          await kirimWA(sender, '⚠️ Format salah!\n\nKetik angka saja (contoh: _29812000_)\natau *-* untuk kosong.');
          return;
        }
        nominal = parseInt(angka);
      }
      
      dataWizard[f.key] = nominal;
      updateSesi(sender, { wizardData: dataWizard });
      
      // Format konfirmasi dengan nama kasir (CP)
      let labelKonfirmasi = f.label;
      if (s.toko === 'cp' && KASIR_CP_DEFAULT[f.key]) {
        labelKonfirmasi = f.label + ' (' + KASIR_CP_DEFAULT[f.key] + ')';
      }
      
      const formatNominal = nominal === 0 ? 'Rp. -' : 'Rp. ' + nominal.toLocaleString('id-ID');
      const next = wizardGetNextField(s.toko, dataWizard);
      
      if (next) {
        await tunggu(300);
        await kirimWA(sender, '✅ _' + labelKonfirmasi + ': ' + formatNominal + ' disimpan_\n\n' + wizardTanyaField(s.toko, next, namaToko, dataWizard));
      } else {
        await kirimWA(sender, '✅ _' + labelKonfirmasi + ': ' + formatNominal + ' disimpan_\n\n🎉 *Semua data sudah diisi!*\n⏳ _Sedang generate laporan..._');
        await tunggu(800);
        const text = wizardToText(dataWizard, s.toko);
        laporan = genLapPenjualan(text, namaToko, s.kemarin, s.toko);
        if (laporan) {
          await kirimWA(sender, laporan);
          resetSesi(sender);
          await tunggu(1500);
          await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
        }
      }
      return;
    }

    // ── MODE NORMAL (foto atau ketik manual untuk menu 2 & 3) ──
    if (image && image.length > 0) {
      await kirimWA(sender, '📸 Foto diterima, sedang dianalisa AI...');
      try {
        const prompt = buatPromptAI(s.menu, namaToko, getTanggal(s.kemarin), s.toko);
        laporan = await analisaGambar(image, prompt);
      } catch (e) {
        log.error('GEMINI', 'Gagal', e.message);
        await kirimWA(sender, '❌ Gagal baca foto. Coba kirim ulang atau ketik manual.');
        return;
      }
    } else if (msg) {
      if (s.menu === 1) laporan = genLapPenjualan(msg, namaToko, s.kemarin, s.toko);
      if (s.menu === 2) laporan = genLapHarga(msg, namaToko, s.kemarin);
      if (s.menu === 3) laporan = genLapMarket(msg, s.kemarin);
    } else return;

    if (laporan) {
      await kirimWA(sender, laporan);
      resetSesi(sender);
      await tunggu(1500);
      await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
    }

  } catch (err) {
    log.error('WEBHOOK', 'Unhandled', err.message);
    try {
      const sender = req.body ? (req.body.sender || req.body.from || '') : '';
      if (sender) await kirimWA(sender, '⚠️ Terjadi kesalahan. Ketik *menu* untuk mulai ulang.');
    } catch (e) {}
  }
});

// ════════════════════════════════════════════════════════════════
//   17. START SERVER
// ════════════════════════════════════════════════════════════════

app.listen(CONFIG.port, function() {
  console.log('\n=====================================');
  console.log('  ' + CONFIG.appName + ' v3.6');
  console.log('  (Auto Wizard + Auto Nama Kasir CP)');
  console.log('=====================================');
  console.log('  Port      : ' + CONFIG.port);
  console.log('  Admin     : ' + CONFIG.adminNumber);
  console.log('  Items     : ' + DATA_BARANG.length);
  console.log('  Members   : ' + MEMBERS.length + '/' + CONFIG.maxMember);
  console.log('=====================================\n');
});

process.on('uncaughtException',  function(e) { log.error('SYSTEM', 'Uncaught',  e.message); });
process.on('unhandledRejection', function(r) { log.error('SYSTEM', 'Unhandled', String(r));  });
