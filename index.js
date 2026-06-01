'use strict';

// ════════════════════════════════════════════════════════════════
//   BOT WHATSAPP - LAPORAN & CARI HARGA BARANG
//   Versi 3.13 - Scan Foto Multi-Step + All Features
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
  geminiKey:     process.env.GEMINI_KEY,
  geminiKey2:    process.env.GEMINI_KEY2 || '',
  geminiKey3:    process.env.GEMINI_KEY3 || '',
  groqKey:       process.env.GROQ_API_KEY || '',
  openrouterKey: process.env.OPENROUTER_API_KEY || '',
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

if (!CONFIG.fonnteToken) {
  console.error('\n❌ ERROR: FONNTE_TOKEN belum diisi\n');
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
  { kode: 'cp',     nama: 'Central Perabot',           alias: ['cp', 'central', 'alak'] },
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

const FIELD_LAPORAN = {
  nk: [
    { key: 'k1', label: 'Kassa 1', emoji: '💵' },
    { key: 'k2', label: 'Kassa 2', emoji: '💵' },
    { key: 'total',  label: 'Total Keseluruhan', emoji: '📦', auto: true },
    { key: 'tunai',  label: 'Tunai',  emoji: '💰' },
    { key: 'debit',  label: 'Debit',  emoji: '💳' },
    { key: 'kredit', label: 'Credit', emoji: '💳' },
    { key: 'ecer',   label: 'Ecer',   emoji: '🛒' },
    { key: 'grosir', label: 'Grosir', emoji: '📦' },
  ],
  tdm: [
    { key: 'k1', label: 'Kassa 1', emoji: '💵' },
    { key: 'k2', label: 'Kassa 2', emoji: '💵' },
    { key: 'total',  label: 'Total Keseluruhan', emoji: '📦', auto: true },
    { key: 'tunai',  label: 'Tunai',  emoji: '💰' },
    { key: 'debit',  label: 'Debit',  emoji: '💳' },
    { key: 'kredit', label: 'Credit', emoji: '💳' },
  ],
  oesapa: [
    { key: 'k1', label: 'Kassa 1', emoji: '💵' },
    { key: 'k2', label: 'Kassa 2', emoji: '💵' },
    { key: 'total',  label: 'Total Keseluruhan', emoji: '📦', auto: true },
    { key: 'tunai',  label: 'Tunai',  emoji: '💰' },
    { key: 'debit',  label: 'Debit',  emoji: '💳' },
    { key: 'kredit', label: 'Credit', emoji: '💳' },
  ],
  kefa: [
    { key: 'k1', label: 'Kassa 1', emoji: '💵' },
    { key: 'k2', label: 'Kassa 2', emoji: '💵' },
    { key: 'total',  label: 'Total Keseluruhan', emoji: '📦', auto: true },
    { key: 'tunai',  label: 'Tunai',  emoji: '💰' },
    { key: 'debit',  label: 'Debit',  emoji: '💳' },
    { key: 'kredit', label: 'Credit', emoji: '💳' },
  ],
  cp: [
    { key: 'k1', label: 'Kassa 1', emoji: '💵' },
    { key: 'k2', label: 'Kassa 2', emoji: '💵' },
    { key: 'k3', label: 'Kassa 3', emoji: '💵' },
    { key: 'k4', label: 'Kassa 4', emoji: '💵' },
    { key: 'total',  label: 'Total Keseluruhan', emoji: '📦', auto: true },
    { key: 'tunai',  label: 'Tunai',  emoji: '💰' },
    { key: 'debit',  label: 'Debit',  emoji: '💳' },
    { key: 'kredit', label: 'Credit', emoji: '💳' },
    { key: 'ecer',   label: 'Ecer',   emoji: '🛒' },
    { key: 'grosir', label: 'Grosir', emoji: '📦' },
    { key: 'promo',        label: 'Total Kasir Promo', emoji: '🎁' },
    { key: 'promotunai',   label: 'Promo - Tunai',     emoji: '💰' },
    { key: 'promodebit',   label: 'Promo - Debit',     emoji: '💳' },
    { key: 'promokredit',  label: 'Promo - Credit',    emoji: '💳' },
    { key: 'parkirkomputer', label: 'Parkir di Komputer', emoji: '🅿️' },
    { key: 'parkirluar',     label: 'Parkir Stor Luar',   emoji: '🅿️' },
  ],
};

const KASIR_CP_DEFAULT = {
  k1: 'Yuni-Salsa',
  k2: 'Nanda-Umi-Marselina',
  k3: 'Febri-Jien-Tika',
  k4: 'Delfi-Tirsa',
};

// ★★★ SCAN FOTO STEPS PER TOKO ★★★
const SCAN_STEPS = {
  nk: [
    { step: 1, label: 'Kassa 1', fields: ['k1'], scanField: 'total_transaksi' },
    { step: 2, label: 'Kassa 2', fields: ['k2'], scanField: 'total_transaksi' },
    { step: 3, label: 'Total + Tunai + Debit + Credit', fields: ['total', 'tunai', 'debit', 'kredit'], scanField: 'multi' },
    { step: 4, label: 'Ecer', fields: ['ecer'], scanField: 'total_transaksi' },
    { step: 5, label: 'Grosir', fields: ['grosir'], scanField: 'total_transaksi' },
  ],
  tdm: [
    { step: 1, label: 'Kassa 1', fields: ['k1'], scanField: 'total_transaksi' },
    { step: 2, label: 'Kassa 2', fields: ['k2'], scanField: 'total_transaksi' },
    { step: 3, label: 'Total + Tunai + Debit + Credit', fields: ['total', 'tunai', 'debit', 'kredit'], scanField: 'multi' },
  ],
  oesapa: [
    { step: 1, label: 'Kassa 1', fields: ['k1'], scanField: 'total_transaksi' },
    { step: 2, label: 'Kassa 2', fields: ['k2'], scanField: 'total_transaksi' },
    { step: 3, label: 'Total + Tunai + Debit + Credit', fields: ['total', 'tunai', 'debit', 'kredit'], scanField: 'multi' },
  ],
  kefa: [
    { step: 1, label: 'Kassa 1', fields: ['k1'], scanField: 'total_transaksi' },
    { step: 2, label: 'Kassa 2', fields: ['k2'], scanField: 'total_transaksi' },
    { step: 3, label: 'Total + Tunai + Debit + Credit', fields: ['total', 'tunai', 'debit', 'kredit'], scanField: 'multi' },
  ],
  cp: [
    { step: 1, label: 'Kassa 1 (Yuni-Salsa)', fields: ['k1'], scanField: 'total_transaksi' },
    { step: 2, label: 'Kassa 2 (Nanda-Umi-Marselina)', fields: ['k2'], scanField: 'total_transaksi' },
    { step: 3, label: 'Kassa 3 (Febri-Jien-Tika)', fields: ['k3'], scanField: 'total_transaksi' },
    { step: 4, label: 'Kassa 4 (Delfi-Tirsa)', fields: ['k4'], scanField: 'total_transaksi' },
    { step: 5, label: 'Total + Tunai + Debit + Credit', fields: ['total', 'tunai', 'debit', 'kredit'], scanField: 'multi' },
    { step: 6, label: 'Ecer', fields: ['ecer'], scanField: 'total_transaksi' },
    { step: 7, label: 'Grosir', fields: ['grosir'], scanField: 'total_transaksi' },
    { step: 8, label: 'Kasir Promo (Total + Tunai + Debit + Credit)', fields: ['promo', 'promotunai', 'promodebit', 'promokredit'], scanField: 'multi_promo' },
    { step: 9, label: 'Parkir di Komputer', fields: ['parkirkomputer'], scanField: 'manual' },
    { step: 10, label: 'Parkir Stor Luar', fields: ['parkirluar'], scanField: 'manual' },
  ],
};

// Prompt scan per type
const SCAN_PROMPTS = {
  total_transaksi: 'Baca gambar tabel laporan penjualan iPos. Cari baris TOTAL (paling bawah). Ambil angka dari kolom "Total Transaksi" pada baris TOTAL.\n\nJawab HANYA dengan angka saja (tanpa Rp, tanpa titik, tanpa koma, tanpa teks lain).\n\nContoh jawaban yang BENAR:\n15741500\n\nContoh jawaban yang SALAH:\nRp 15.741.500\nTotal: 15.741.500\nTotal transaksi adalah 15741500\n\nSekarang baca dan jawab dengan angka saja:',
  
  multi: 'Baca gambar tabel laporan penjualan iPos. Cari baris TOTAL (paling bawah tabel).\n\nAmbil 4 angka berikut dari baris TOTAL:\n1. "Total Transaksi" → total\n2. "Jml Bayar Tunai" → tunai\n3. "Jml Bayar K.Debit" → debit\n4. "Jml Bayar Kredit" atau "Jml Bayar K.Kredit" → kredit\n\nJawab dengan format PERSIS seperti ini (hanya angka, tanpa Rp/titik/koma):\n\ntotal: 40899000\ntunai: 26326500\ndebit: 14254500\nkredit: 318000\n\nKalau ada nilai 0 atau kosong, tulis: 0\n\nSekarang baca gambar dan jawab:',
  
  multi_promo: 'Baca gambar tabel laporan penjualan KASIR PROMO. Cari baris TOTAL (paling bawah).\n\nAmbil 4 angka berikut dari baris TOTAL:\n1. "Total Transaksi" → promo\n2. "Jml Bayar Tunai" → promotunai\n3. "Jml Bayar K.Debit" → promodebit\n4. "Jml Bayar Kredit" atau "Jml Bayar K.Kredit" → promokredit\n\nJawab dengan format PERSIS seperti ini (hanya angka):\n\npromo: 1675000\npromotunai: 1675000\npromodebit: 0\npromokredit: 0\n\nSekarang baca gambar dan jawab:',
};

function getScanStep(tokoKode, currentStep) {
  const steps = SCAN_STEPS[tokoKode];
  if (!steps) return null;
  if (currentStep >= steps.length) return null;
  return steps[currentStep];
}

function parseScanSingle(aiResponse) {
  if (!aiResponse) return 0;
  const angka = aiResponse.replace(/[^0-9]/g, '');
  if (!angka || angka.length < 1) return 0;
  return parseInt(angka) || 0;
}

function parseScanMulti(aiResponse, fields) {
  const hasil = {};
  fields.forEach(function(f) { hasil[f] = 0; });
  if (!aiResponse) return hasil;
  
  const lines = aiResponse.split('\n');
  lines.forEach(function(line) {
    const lower = line.toLowerCase().trim();
    fields.forEach(function(f) {
      const fLow = f.toLowerCase();
      if (lower.startsWith(fLow + ':') || lower.startsWith(fLow + ' :') || lower.startsWith('- ' + fLow)) {
        const angkaPart = line.substring(line.indexOf(':') + 1).trim();
        const angka = angkaPart.replace(/[^0-9]/g, '');
        if (angka) hasil[f] = parseInt(angka) || 0;
      }
    });
  });
  return hasil;
}

const DEFAULT_MEMBERS=['6285253949803','6285737005301','6285211988252','6281383924057','6282235572821','6287841617474','6281584937710','6281238774152','6282266026564','6281238643890','6281353888652'];

const DEFAULT_KONTAK={
  '6285253949803':'Pak Security Marthen',
  '6285737005301':'Kak Bagas Pacar Beda Agama',
  '6285211988252':'Kak Admin Marketplace',
  '6281383924057':'Kak Fajar (Bukan Mas Fajar Kefa)',
  '6282235572821':'Kak yang Saya Tidak Tau Namanya',
  '6287841617474':'Mas Awin Gacor',
  '6281584937710':'Kak Safira',
  '6281238774152':'Ibu Risti HRD',
  '6282266026564':'Mas Abi Mustafa',
  '6285829278962':'Admin'
  '6281238643890':'Cowok YANG SERING DI PHP IN'
  '6281353888652':'MY BINI'
};

// ★★★ ROLE-BASED ACCESS ★★★
// Nomor yang bisa akses menu Laporan (1-3)
const ROLE_LAPORAN = [
  '6281584937710',  // Kak Safira
  '6285211988252',  // Kak Admin Marketplace
  '6287841617474',  // Mas Awin Gacor
  '6281238774152',  // Ibu Risti HRD
  '6281353888652',  // MY BINI
];

function bisaAksesLaporan(nomor) {
  return isAdmin(nomor) || ROLE_LAPORAN.indexOf(nomor) >= 0;
}

const ADMIN_COMMANDS = ['daftar','hapus','listmember','namakontak','hapuskontak','listkontak','reload','info','resetall'];

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

function getSesi(nomor) { 
  if (!SESI[nomor]) SESI[nomor] = {}; 
  SESI[nomor]._lastActive = Date.now(); 
  return SESI[nomor]; 
}

function resetSesi(nomor) { 
  delete SESI[nomor];
  SESI[nomor] = { _lastActive: Date.now() };
  saveJSON(CONFIG.paths.sesi, SESI);
  log.info('SESI', 'Reset TOTAL untuk ' + nomor);
}

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
  
  // 1. Exact match by kode
  const byKode = DATA_BARANG.filter(function(d) { return d.kode === q; });
  if (byKode.length > 0) return { hasil: byKode, saran: [], tipeHasil: 'exact', totalDitemukan: byKode.length };
  
  // 2. Exact match (SEMUA kata harus ada di nama)
  const exactResults = DATA_BARANG.filter(function(d) {
    const namaBersih = bersihkanTeks(d.nama);
    return words.every(function(w) { return namaBersih.indexOf(w) >= 0 || d.kode.indexOf(w) >= 0; });
  });
  
  if (exactResults.length > 0) {
    // ★ SORT ABJAD ★
    exactResults.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
    return { 
      hasil: exactResults.slice(0, CONFIG.maxHasilCari), 
      saran: [], 
      tipeHasil: 'exact', 
      totalDitemukan: exactResults.length 
    };
  }
  
  // 3. Fuzzy search — TAPI lebih ketat
  // Minimal 70% kata harus cocok (exact atau fuzzy per kata)
  const skorItems = [];
  const minMatchRatio = 0.7; // Minimal 70% kata harus match
  
  DATA_BARANG.forEach(function(item) {
    const namaBersih = bersihkanTeks(item.nama);
    const namaWords = namaBersih.split(/\s+/);
    
    let matchCount = 0;
    let totalSkor = 0;
    
    words.forEach(function(w) {
      let bestMatch = 0;
      
      // Exact contain
      if (namaBersih.indexOf(w) >= 0) {
        bestMatch = 10;
      }
      // Kode contain
      else if (item.kode.indexOf(w) >= 0) {
        bestMatch = 10;
      }
      // Fuzzy per kata (toleransi typo)
      else {
        for (let i = 0; i < namaWords.length; i++) {
          if (kataMirip(w, namaWords[i])) {
            bestMatch = 5;
            break;
          }
        }
      }
      
      if (bestMatch > 0) {
        matchCount++;
        totalSkor += bestMatch;
      }
    });
    
    // ★ FILTER KETAT: minimal 70% kata harus match ★
    const matchRatio = matchCount / words.length;
    
    if (matchRatio >= minMatchRatio && totalSkor > 0) {
      // Bonus skor kalau semua kata match
      if (matchCount === words.length) totalSkor += 20;
      // Bonus kalau nama pendek (lebih spesifik)
      if (namaBersih.split(/\s+/).length <= words.length + 3) totalSkor += 5;
      
      skorItems.push({ 
        item: item, 
        skor: totalSkor, 
        matchRatio: matchRatio,
        fuzzy: matchCount < words.length 
      });
    }
  });
  
  skorItems.sort(function(a, b) { 
    // Sort: skor tertinggi dulu, kalau sama → abjad
    if (b.skor !== a.skor) return b.skor - a.skor;
    return a.item.nama.localeCompare(b.item.nama);
  });
  
  if (skorItems.length > 0) {
    const batasSkor = skorItems[0].skor * 0.5;
    const hasilBagus = skorItems.filter(function(s) { return s.skor >= batasSkor; });
    
    // ★ SORT ABJAD untuk hasil akhir ★
    const hasilTerbatas = hasilBagus.slice(0, CONFIG.maxHasilCari);
    hasilTerbatas.sort(function(a, b) { return a.item.nama.localeCompare(b.item.nama); });
    
    const adaFuzzy = hasilTerbatas.some(function(s) { return s.fuzzy; });
    return {
      hasil: hasilTerbatas.map(function(s) { return s.item; }),
      saran: [],
      tipeHasil: adaFuzzy ? 'fuzzy' : 'exact',
      totalDitemukan: hasilBagus.length,
    };
  }
  
  // 4. Saran — cari per kata yang paling cocok
  const saranSet = {};
  words.forEach(function(w) {
    if (w.length < 3) return;
    DATA_BARANG.forEach(function(item) {
      const namaBersih = bersihkanTeks(item.nama);
      const namaWords = namaBersih.split(/\s+/);
      
      // Hanya tambah saran kalau kata ADA di nama (exact atau mirip)
      let cocok = false;
      if (namaBersih.indexOf(w) >= 0) cocok = true;
      else {
        for (let i = 0; i < namaWords.length; i++) {
          if (kataMirip(w, namaWords[i])) { cocok = true; break; }
        }
      }
      
      if (cocok) {
        if (!saranSet[item.kode]) saranSet[item.kode] = { item: item, matchCount: 0 };
        saranSet[item.kode].matchCount++;
      }
    });
  });
  
  const saranList = Object.values(saranSet)
    .sort(function(a, b) { 
      // Sort: paling banyak match dulu, lalu abjad
      if (b.matchCount !== a.matchCount) return b.matchCount - a.matchCount;
      return a.item.nama.localeCompare(b.item.nama);
    })
    .slice(0, 5)
    .map(function(s) { return s.item; });
  
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
//   6. ANALISA BANDING HARGA
// ════════════════════════════════════════════════════════════════

function analisaPerbandinganHarga(item) { return analisaPerbandinganHargaToko(item, null); }

function analisaPerbandinganHargaToko(item, tokoFilter) {
  const hasil = {
    item: item, analisa: { ecer: null, ambil: null },
    adaToko: [], kosongToko: [], stokAda: [], stokKosong: [], tokoFilter: tokoFilter,
  };
  const dataEcer = [];
  const dataAmbil = [];
  const tokoUntukDicek = tokoFilter && tokoFilter.length > 0 ? tokoFilter : TOKO_LIST;
  tokoUntukDicek.forEach(function(t) {
    const h = item.harga[t.kode];
    if (h.ecer > 0) { dataEcer.push({ toko: t, harga: h.ecer }); hasil.adaToko.push(t); }
    else hasil.kosongToko.push(t);
    if (h.ambil > 0) dataAmbil.push({ toko: t, harga: h.ambil });
    if (h.stok > 0) hasil.stokAda.push({ toko: t, stok: h.stok });
    else hasil.stokKosong.push(t);
  });
  if (dataEcer.length > 0) {
    const sortedEcer = dataEcer.sort(function(a, b) { return a.harga - b.harga; });
    const termurah = sortedEcer[0]; const termahal = sortedEcer[sortedEcer.length - 1];
    const avgEcer = Math.round(sortedEcer.reduce(function(sum, x) { return sum + x.harga; }, 0) / sortedEcer.length);
    const selisih = termahal.harga - termurah.harga;
    const persenSelisih = termurah.harga > 0 ? Math.round((selisih / termurah.harga) * 100) : 0;
    hasil.analisa.ecer = {
      termurah, termahal, rataRata: avgEcer, selisih, persenSelisih,
      semuaSama: sortedEcer.every(function(x) { return x.harga === sortedEcer[0].harga; }),
      sorted: sortedEcer,
    };
  }
  return hasil;
}

function formatAnalisaBandingHargaToko(item, tokoFilter, sender) {
  const nama = getNama(sender);
  const suffix = '\n\n' + (nama ? '_Semoga membantu, *' + nama + '*!_ 😊' : '_Semoga membantu!_ 😊');
  const a = analisaPerbandinganHargaToko(item, tokoFilter);
  let msg = '📊 *ANALISA PERBANDINGAN HARGA*\n' + GARIS_TEBAL + '\n';
  msg += '🔖 *Kode:* ' + item.kode + '\n📦 *Nama:* ' + item.nama + '\n';
  msg += '🏷️ *Jenis:* ' + (item.jenis || '-') + '\n🏗️ *Merek:* ' + (item.merek || '-') + '\n';
  msg += '📏 *Satuan:* ' + item.satuan + '\n';
  if (tokoFilter && tokoFilter.length > 0 && tokoFilter.length < TOKO_LIST.length) {
    msg += '🎯 *Toko yang dibandingkan:* ' + tokoFilter.length + ' toko\n';
    tokoFilter.forEach(function(t) { msg += '   • ' + t.nama + '\n'; });
  }
  msg += GARIS_TEBAL + '\n\n💰 *HARGA ECER:*\n' + GARIS_TIPIS + '\n';
  const tokoUntukTampil = tokoFilter && tokoFilter.length > 0 ? tokoFilter : TOKO_LIST;
  tokoUntukTampil.forEach(function(t) {
    const h = item.harga[t.kode];
    const hargaStr = h.ecer > 0 ? fRp(h.ecer) : '⚠️ Tidak ada harga';
    msg += '🏪 *' + t.nama + '*\n   ' + hargaStr;
    if (a.analisa.ecer) {
      if (h.ecer === a.analisa.ecer.termurah.harga && h.ecer > 0) msg += ' 🟢 *TERMURAH*';
      else if (h.ecer === a.analisa.ecer.termahal.harga && h.ecer > 0 && !a.analisa.ecer.semuaSama) msg += ' 🔴 *TERMAHAL*';
    }
    msg += (h.stok > 0 ? ' _(stok: ' + h.stok + ')_' : ' _(⚠️ kosong)_') + '\n';
  });
  msg += '\n';
  if (a.analisa.ecer) {
    msg += '📈 *KESIMPULAN:*\n' + GARIS_TIPIS + '\n';
    if (a.analisa.ecer.semuaSama) {
      msg += '✅ *Harga SAMA*: ' + fRp(a.analisa.ecer.termurah.harga) + '\n';
    } else {
      msg += '🟢 *Termurah:* ' + a.analisa.ecer.termurah.toko.nama + '\n   ' + fRp(a.analisa.ecer.termurah.harga) + '\n\n';
      msg += '🔴 *Termahal:* ' + a.analisa.ecer.termahal.toko.nama + '\n   ' + fRp(a.analisa.ecer.termahal.harga) + '\n\n';
      if (tokoUntukTampil.length > 2) msg += '📊 *Rata-rata:* ' + fRp(a.analisa.ecer.rataRata) + '\n';
      msg += '💸 *Selisih:* ' + fRp(a.analisa.ecer.selisih) + ' _(' + a.analisa.ecer.persenSelisih + '% lebih mahal)_\n';
      msg += '\n💡 *Rekomendasi:*\nHemat *' + fRp(a.analisa.ecer.selisih) + '* (' + a.analisa.ecer.persenSelisih + '%) kalau beli di *' + a.analisa.ecer.termurah.toko.nama + '*!';
    }
  }
  msg += '\n\n' + GARIS_TEBAL;
  return msg + suffix;
}

function isPertanyaanBanding(low) {
  const KATA_BANDING = ['banding','bandingkan','compare','komparasi','termurah','termahal','paling murah','paling mahal','lebih murah','lebih mahal','mana yang murah','mana yang mahal','dimana murah','dimana mahal','selisih','beda harga','perbedaan harga','analisa harga','analisis harga'];
  return KATA_BANDING.some(function(k) { return low.indexOf(k) >= 0; });
}

function cariBarangPrioritas(keyword) {
  const q = keyword.trim().toUpperCase();
  const byKode = DATA_BARANG.filter(function(d) { return d.kode === q; });
  if (byKode.length > 0) return { hasil: byKode, exact: true };
  const byNamaFull = DATA_BARANG.filter(function(d) { return d.nama === q; });
  if (byNamaFull.length > 0) return { hasil: byNamaFull, exact: true };
  const namaMengandung = DATA_BARANG.filter(function(d) { return d.nama.indexOf(q) >= 0; });
  if (namaMengandung.length > 0) return { hasil: namaMengandung, exact: true };
  const words = q.split(/\s+/).filter(function(w) { return w.length > 0; });
  const semuaKataAda = DATA_BARANG.filter(function(d) {
    return words.every(function(w) { return d.nama.indexOf(w) >= 0 || d.kode.indexOf(w) >= 0; });
  });
  if (semuaKataAda.length > 0) return { hasil: semuaKataAda, exact: false };
  return { hasil: cariBarang(keyword).hasil || [], exact: false };
}

// ════════════════════════════════════════════════════════════════
//   7. HELPER DETEKSI TOKO & KONFIRMASI
// ════════════════════════════════════════════════════════════════

function deteksiTokoDariTeks(low) {
  const tokoDitemukan = [];
  const sudahAda = {};
  TOKO_LIST.forEach(function(t) {
    let ketemu = false;
    const regexKode = new RegExp('\\b' + t.kode + '\\b', 'i');
    if (regexKode.test(low)) ketemu = true;
    if (!ketemu) {
      for (let i = 0; i < t.alias.length; i++) {
        const regexAlias = new RegExp('\\b' + t.alias[i] + '\\b', 'i');
        if (regexAlias.test(low)) { ketemu = true; break; }
      }
    }
    if (!ketemu) {
      const namaLow = t.nama.toLowerCase();
      const kataNama = namaLow.split(/\s+/);
      for (let i = 0; i < kataNama.length; i++) {
        if (kataNama[i].length >= 4 && low.indexOf(kataNama[i]) >= 0) { ketemu = true; break; }
      }
    }
    if (ketemu && !sudahAda[t.kode]) { tokoDitemukan.push(t); sudahAda[t.kode] = true; }
  });
  return tokoDitemukan;
}

function isKonfirmasiYa(low) {
  const KATA_YA = ['iya','ya','yes','yup','yep','yoi','oke','ok','okey','okay','betul','bener','benar','sip','siap','silakan','silahkan','lanjut','lanjutkan','gas','mau','boleh','setuju','tentu','pasti','cocok','mantap','pindah'];
  const cleaned = low.trim().toLowerCase();
  if (KATA_YA.indexOf(cleaned) >= 0) return true;
  for (let i = 0; i < KATA_YA.length; i++) {
    if (cleaned.startsWith(KATA_YA[i] + ' ') || cleaned.startsWith(KATA_YA[i] + ',') ||
        cleaned.startsWith(KATA_YA[i] + '!') || cleaned.startsWith(KATA_YA[i] + '.')) return true;
  }
  return false;
}

function isKonfirmasiTidak(low) {
  const KATA_TIDAK = ['tidak','tdk','tdak','no','nope','engga','enggak','ga','gak','jangan','jgn','salah','bukan','tetap','tetap di sini'];
  const cleaned = low.trim().toLowerCase();
  if (KATA_TIDAK.indexOf(cleaned) >= 0) return true;
  for (let i = 0; i < KATA_TIDAK.length; i++) {
    if (cleaned.startsWith(KATA_TIDAK[i] + ' ') || cleaned.startsWith(KATA_TIDAK[i] + ',')) return true;
  }
  return false;
}

function bersihkanKeywordDariToko(pesan) {
  let cleaned = pesan;
  TOKO_LIST.forEach(function(t) {
    const regexKode = new RegExp('\\b' + t.kode + '\\b', 'gi');
    cleaned = cleaned.replace(regexKode, '');
    t.alias.forEach(function(a) {
      cleaned = cleaned.replace(new RegExp('\\b' + a + '\\b', 'gi'), '');
    });
    const kataNama = t.nama.toLowerCase().split(/\s+/);
    kataNama.forEach(function(kn) {
      if (kn.length >= 4) cleaned = cleaned.replace(new RegExp('\\b' + kn + '\\b', 'gi'), '');
    });
  });
  cleaned = cleaned.replace(/\b(di|ke|untuk|toko|cek|cari|harga|stok)\b/gi, ' ');
  return cleaned.trim().replace(/\s+/g, ' ');
}

// ════════════════════════════════════════════════════════════════
//   8. KIRIM WHATSAPP (FONNTE)
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
//   9. MULTI-AI PROVIDER
// ════════════════════════════════════════════════════════════════

let currentGeminiKeyIndex = 0;
const geminiKeys = function() {
  const keys = [];
  if (CONFIG.geminiKey)  keys.push(CONFIG.geminiKey);
  if (CONFIG.geminiKey2) keys.push(CONFIG.geminiKey2);
  if (CONFIG.geminiKey3) keys.push(CONFIG.geminiKey3);
  return keys;
};

function geminiUrl() {
  const keys = geminiKeys();
  const key = keys[currentGeminiKeyIndex % keys.length] || CONFIG.geminiKey;
  return 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + key;
}

function rotateGeminiKey() {
  const keys = geminiKeys();
  if (keys.length > 1) {
    currentGeminiKeyIndex = (currentGeminiKeyIndex + 1) % keys.length;
    log.info('GEMINI', 'Rotate ke key #' + (currentGeminiKeyIndex + 1));
  }
}

async function chatGroq(prompt) {
  if (!CONFIG.groqKey) return null;
  try {
    const resp = await axios.post('https://api.groq.com/openai/v1/chat/completions',
      { model: 'llama-3.3-70b-versatile', messages: [{ role: 'user', content: prompt }], temperature: 0.7, max_tokens: 1500 },
      { headers: { 'Authorization': 'Bearer ' + CONFIG.groqKey, 'Content-Type': 'application/json' }, timeout: 30000 });
    return (resp.data.choices[0].message.content || '').trim();
  } catch (err) { log.warn('GROQ', 'Gagal: ' + err.message); return null; }
}

async function chatGemini(prompt) {
  if (!CONFIG.geminiKey) return null;
  const keys = geminiKeys();
  for (let i = 0; i < keys.length; i++) {
    try {
      const resp = await axios.post(geminiUrl(), { contents: [{ parts: [{ text: prompt }] }] }, { timeout: 30000 });
      return (resp.data.candidates[0].content.parts[0].text || '').trim();
    } catch (err) {
      const status = err.response ? err.response.status : 'NETWORK';
      log.warn('GEMINI', 'Gagal key #' + (currentGeminiKeyIndex + 1) + ' (' + status + ')');
      if (status === 429 && keys.length > 1) { rotateGeminiKey(); continue; }
      return null;
    }
  }
  return null;
}

async function chatOpenRouter(prompt) {
  if (!CONFIG.openrouterKey) return null;
  try {
    const resp = await axios.post('https://openrouter.ai/api/v1/chat/completions',
      { model: 'meta-llama/llama-3.2-3b-instruct:free', messages: [{ role: 'user', content: prompt }], temperature: 0.7, max_tokens: 1500 },
      { headers: { 'Authorization': 'Bearer ' + CONFIG.openrouterKey, 'Content-Type': 'application/json', 'HTTP-Referer': 'https://github.com/dimassilitonga99/bot-whatsapp', 'X-Title': 'Bot Toko Perabot' }, timeout: 30000 });
    return (resp.data.choices[0].message.content || '').trim();
  } catch (err) { log.warn('OPENROUTER', 'Gagal: ' + err.message); return null; }
}

async function chatAI(prompt) {
  if (CONFIG.groqKey) {
    const r = await chatGroq(prompt);
    if (r && r.length >= 10) return { jawaban: r, provider: 'GROQ' };
  }
  if (CONFIG.geminiKey) {
    const r = await chatGemini(prompt);
    if (r && r.length >= 10) return { jawaban: r, provider: 'GEMINI' };
  }
  if (CONFIG.openrouterKey) {
    const r = await chatOpenRouter(prompt);
    if (r && r.length >= 10) return { jawaban: r, provider: 'OPENROUTER' };
  }
  log.error('AI', 'Semua provider gagal!');
  return null;
}

// ════════════════════════════════════════════════════════════════
//   10. ★★★ SCAN GAMBAR (untuk laporan & search) ★★★
// ════════════════════════════════════════════════════════════════

async function analisaGambar(imageUrl, prompt) {
  const imgResp   = await axios.get(imageUrl, { responseType: 'arraybuffer', timeout: 15000 });
  const mimeType  = imgResp.headers['content-type'] || 'image/jpeg';
  const imageData = Buffer.from(imgResp.data).toString('base64');
  const keys = geminiKeys();
  for (let i = 0; i < keys.length; i++) {
    try {
      const resp = await axios.post(geminiUrl(), {
        contents: [{ parts: [{ inline_data: { mime_type: mimeType, data: imageData } }, { text: prompt }] }],
      }, { timeout: 30000 });
      return resp.data.candidates[0].content.parts[0].text || '';
    } catch (err) {
      const status = err.response ? err.response.status : 'NETWORK';
      log.warn('GEMINI_IMG', 'Gagal key #' + (currentGeminiKeyIndex + 1) + ' (' + status + ')');
      if (status === 429 && keys.length > 1) { rotateGeminiKey(); continue; }
      throw err;
    }
  }
  throw new Error('Semua Gemini key gagal');
}
// ════════════════════════════════════════════════════════════════
//   11. AI CHAT BARANG
// ════════════════════════════════════════════════════════════════

function cariRelevan(pertanyaan, maxResults) {
  maxResults = maxResults || 30;
  const q = pertanyaan.toUpperCase();
  const qBersih = bersihkanTeks(q);
  const words = qBersih.split(/\s+/).filter(function(w) { return w.length >= 2; });
  if (words.length === 0) return DATA_BARANG.slice(0, maxResults);
  const scored = [];
  DATA_BARANG.forEach(function(item) {
    let score = 0;
    const namaBersih = bersihkanTeks(item.nama);
    const merekBersih = bersihkanTeks(item.merek);
    const namaWords = namaBersih.split(/\s+/);
    words.forEach(function(w) {
      if (namaBersih.indexOf(w) >= 0) score += 10;
      if (merekBersih.indexOf(w) >= 0) score += 8;
      if (item.kode.indexOf(w) >= 0) score += 15;
      for (let i = 0; i < namaWords.length; i++) {
        if (kataMirip(w, namaWords[i])) { score += 5; break; }
      }
    });
    if (score > 0) scored.push({ item, score });
  });
  scored.sort(function(a, b) { return b.score - a.score; });
  return scored.slice(0, maxResults).map(function(s) { return s.item; });
}

async function aiChatBarang(pertanyaan, sender, tokoAktif) {
  let relevantItems = cariRelevan(pertanyaan, 30);
  let tokoFilter = [];
  if (tokoAktif) {
    const tokoObj = TOKO_LIST.find(function(t) { return t.kode === tokoAktif; });
    if (tokoObj) tokoFilter = [tokoObj];
  } else {
    tokoFilter = deteksiTokoDariTeks(pertanyaan.toLowerCase());
  }
  let context = '';
  if (relevantItems.length > 0) {
    relevantItems.forEach(function(d, i) {
      context += (i + 1) + '. Kode: ' + d.kode + ' | Nama: ' + d.nama + ' | Satuan: ' + d.satuan + '\n';
      const tokoShow = tokoFilter.length > 0 ? tokoFilter : TOKO_LIST;
      tokoShow.forEach(function(t) {
        const h = d.harga[t.kode];
        context += '   - ' + t.nama + ': Ecer Rp' + h.ecer.toLocaleString('id-ID') + ', Ambil Rp' + h.ambil.toLocaleString('id-ID') + ', Stok: ' + (h.stok > 0 ? h.stok : 'KOSONG') + '\n';
      });
      context += '\n';
    });
  } else context = '(Tidak ada barang cocok)';
  
  const nama = getNama(sender);
  const sapaan = nama ? nama : 'kakak';
  let filterInfo = '';
  if (tokoAktif) {
    const tokoObj = TOKO_LIST.find(function(t) { return t.kode === tokoAktif; });
    if (tokoObj) filterInfo = '\n⚠️ KONTEKS: User di MODE CARI di *' + tokoObj.nama + '*. JAWAB HANYA UNTUK TOKO INI!\n';
  } else if (tokoFilter.length > 0 && tokoFilter.length < TOKO_LIST.length) {
    filterInfo = '\nUSER MENANYAKAN TOKO: ' + tokoFilter.map(function(t) { return t.nama; }).join(', ') + '. JAWAB HANYA TOKO ITU!\n';
  }
  
  const prompt = 'Kamu asisten AI toko perabot "Bot Perabot". Ramah, bahasa Indonesia santai.\n5 toko: NK, TDM, Oesapa, Kefa, CP.\n' +
    filterInfo + '\nDATA (' + relevantItems.length + ' item):\n' + context + '\n' +
    'PERTANYAAN (' + sapaan + '): "' + pertanyaan + '"\n\n' +
    'ATURAN:\n1. Panggil "' + sapaan + '"\n2. Emoji sesuai\n3. Harga: Rp 1.000.000\n4. *bold* penting\n5. Max 1500 char\n' +
    '6. Patuhi konteks toko!\n7. ⚠️ SELALU tampilkan harga Ecer + Ambil walau stok KOSONG!\n8. Akhiri tawaran bantuan\nJawab:';

  const result = await chatAI(prompt);
  if (result && result.jawaban) return result.jawaban;
  return null;
}

function isPertanyaanBarang(low) {
  const KATA_TANYA = ['stok','stock','harga','price','berapa','ada gak','ada ga','ada kah','apakah ada','masih ada','rekomendasi','rekomen','sarankan','saran','kosong','habis','tersedia','cek','check','lihat','tampilkan','total','jumlah','cari','mencari','mau','butuh','info','detail','data'];
  const KATA_BARANG = ['panci','dandang','wajan','penggorengan','rice cooker','kompor','konpor','gelas','piring','mangkok','sendok','garpu','pisau','eagle','maxim','sunkist','golden','paramount','hock','sunlife','miyako','aluminium','alm','stainless','plastik','kaca','keramik','kursi','meja','lemari','rak','sumbu','minyak','gas','tl','serbaguna','susu','jar','drink','keranjang','periuk','magic','com','mcm'];
  const KATA_TOKO = ['nk','tdm','oesapa','kefa','cp','nasional','central','mama','mamaku','kefamenanu'];
  const adaTanya = KATA_TANYA.some(function(k) { return low.indexOf(k) >= 0; });
  const adaBarang = KATA_BARANG.some(function(k) { return low.indexOf(k) >= 0; });
  const adaToko = KATA_TOKO.some(function(k) { return low.indexOf(k) >= 0; });
  const adaKode = /nn\d{4,5}/i.test(low);
  const jumlahKata = low.split(/\s+/).length;
  return adaKode || (adaTanya && adaBarang) || (adaTanya && adaToko) || (adaBarang && adaToko) || (jumlahKata >= 4 && adaBarang);
}

function isPertanyaanUmum(low) {
  const TENTANG_BOT = ['apa kabar','gimana kabar','bisa apa','bisa apa saja','kamu bisa apa','siapa kamu','kamu siapa','cara pakai','tolong','bantu','help','bantuan','jelaskan'];
  const RANDOM = ['cerita','lucu','joke','lagi apa','sedang apa','ngapain','bingung','pusing'];
  const TANYA = ['apa','siapa','dimana','kapan','bagaimana','kenapa','mengapa'];
  return TENTANG_BOT.some(function(k) { return low.indexOf(k) >= 0; }) || 
         RANDOM.some(function(k) { return low.indexOf(k) >= 0; }) || 
         (TANYA.some(function(k) { return low.indexOf(k) >= 0; }) && low.split(/\s+/).length >= 3);
}

async function aiChatUmum(pertanyaan, sender) {
  const nama = getNama(sender);
  const sapaan = nama ? nama : 'kakak';
  const prompt = 'Kamu "Bot Perabot" asisten WhatsApp 5 toko perabot. Database 9.982 barang.\nFitur: Cari harga/stok, banding harga, rekomendasi, laporan penjualan (Wizard + Scan Foto), ngobrol.\nPerintah: *menu*, *1*-*4*, *9* (admin)\n\nPERTANYAAN (' + sapaan + '): "' + pertanyaan + '"\n\nJawab ramah, emoji, *bold*, max 1000 char, akhiri tawaran bantuan:';
  const result = await chatAI(prompt);
  if (result && result.jawaban) return result.jawaban;
  return null;
}

function buatPromptAI(menuType, namaToko, tanggal, tokoKode) {
  const fmt = ' Format WhatsApp dengan emoji. Format rupiah Rp X.XXX.XXX.';
  if (menuType === 1) {
    if (tokoKode === 'cp') return 'Baca data penjualan toko Central Perabot Alak tanggal ' + tanggal + '. Format: 4 kassa (Kassa 1 Yuni-Salsa, Kassa 2 Nanda-Umi-Marselina, Kassa 3 Febri-Jien-Tika, Kassa 4 Delfi-Tirsa), total keseluruhan, metode bayar, jenis penjualan, kasir promo, parkir.' + fmt;
    if (tokoKode === 'nk') return 'Baca data penjualan toko Nasional Kitchen tanggal ' + tanggal + '. Format: 2 kassa, total, metode bayar, jenis penjualan. Rupiah: Rp. X.XXX.XXX.' + fmt;
    if (tokoKode === 'tdm' || tokoKode === 'oesapa' || tokoKode === 'kefa') {
      const tNama = tokoKode === 'tdm' ? 'Perabot Mama TDM' : tokoKode === 'oesapa' ? 'Perabot Mama Oesapa' : 'Perabot Mamaku Kefamenanu';
      return 'Baca data penjualan toko ' + tNama + ' tanggal ' + tanggal + '. Format: 2 kassa, total, metode bayar TANPA Ecer/Grosir. Rupiah: Rp. X.XXX.XXX.' + fmt;
    }
  }
  if (menuType === 2) return 'Baca data harga barang toko "' + namaToko + '" tanggal ' + tanggal + '. Buat laporan: barang baru, naik harga, turun harga.' + fmt;
  if (menuType === 3) return 'Baca data marketplace tanggal ' + tanggal + '. Buat laporan: per toko, per channel, metode bayar.' + fmt;
  return 'Buat laporan rapi.' + fmt;
}

// ════════════════════════════════════════════════════════════════
//   12. SAPAAN PINTAR
// ════════════════════════════════════════════════════════════════

const SAPAAN_MAP = {
  'halo':{ kategori:'kasual', balasan:['Halo','Haloo','Hai juga','Halooo'] },
  'hai':{ kategori:'kasual', balasan:['Hai','Haiii','Halo juga','Hai!'] },
  'hi':{ kategori:'kasual', balasan:['Hi','Hi there','Hiii'] },
  'hello':{ kategori:'kasual', balasan:['Hello','Hello juga','Helloo'] },
  'hallo':{ kategori:'kasual', balasan:['Hallo','Hallooo','Halo juga'] },
  'helo':{ kategori:'kasual', balasan:['Helo','Helooo','Hai!'] },
  'hey':{ kategori:'kasual', balasan:['Hey','Hey juga','Heyyy'] },
  'hei':{ kategori:'kasual', balasan:['Hei','Hei juga','Halo!'] },
  'yo':{ kategori:'kasual', balasan:['Yo','Yo bro','Yoyoyo'] },
  'p':{ kategori:'kasual', balasan:['p juga','Halo','Hai'] },
  'pagi':{ kategori:'waktu', waktu:'Pagi', balasan:['Selamat pagi','Pagi juga','Pagiii'] },
  'siang':{ kategori:'waktu', waktu:'Siang', balasan:['Selamat siang','Siang juga'] },
  'sore':{ kategori:'waktu', waktu:'Sore', balasan:['Selamat sore','Sore juga'] },
  'malam':{ kategori:'waktu', waktu:'Malam', balasan:['Selamat malam','Malam juga'] },
  'selamat pagi':{ kategori:'waktu', waktu:'Pagi', balasan:['Selamat pagi juga','Pagi yang indah'] },
  'selamat siang':{ kategori:'waktu', waktu:'Siang', balasan:['Selamat siang juga'] },
  'selamat sore':{ kategori:'waktu', waktu:'Sore', balasan:['Selamat sore juga'] },
  'selamat malam':{ kategori:'waktu', waktu:'Malam', balasan:['Selamat malam juga'] },
  'assalamualaikum':{ kategori:'islami', balasan:['Waalaikumsalam warahmatullahi wabarakatuh','Waalaikumsalam'] },
  'assalamu':{ kategori:'islami', balasan:['Waalaikumsalam warahmatullahi wabarakatuh'] },
  'salam':{ kategori:'islami', balasan:['Salam juga','Waalaikumsalam'] },
  'permisi':{ kategori:'sopan', balasan:['Iya, silakan','Ada yang bisa dibantu?'] },
  'maaf':{ kategori:'sopan', balasan:['Iya, tidak apa-apa','Santai saja'] },
  'maaf ganggu':{ kategori:'sopan', balasan:['Tidak mengganggu kok','Santai saja'] },
  'good morning':{ kategori:'english', waktu:'Pagi', balasan:['Good morning','Morning!'] },
  'good afternoon':{ kategori:'english', waktu:'Siang', balasan:['Good afternoon'] },
  'good evening':{ kategori:'english', waktu:'Sore', balasan:['Good evening'] },
  'good night':{ kategori:'english', waktu:'Malam', balasan:['Good night','Sleep well'] },
};

const KATA_SAPAAN_LIST = Object.keys(SAPAAN_MAP);

const KALIMAT_MOTIVASI = [
  '✨ _Hari yang baru, semangat yang baru!_ 💪',
  '🌟 _Setiap hari adalah kesempatan untuk jadi lebih baik!_ 🚀',
  '💎 _Kerja keras hari ini, hasil manis besok!_ 🍯',
  '🌈 _Tetap semangat dan jangan menyerah!_ 🔥',
  '⭐ _Senyum dulu, rezeki menyusul!_ 😊',
  '🎯 _Fokus pada tujuan, abaikan keraguan!_ 💯',
  '🚀 _Sukses dimulai dari langkah kecil hari ini!_ 👣',
  '💪 _Kamu lebih kuat dari yang kamu kira!_ 💯',
  '🌻 _Jangan lupa bahagia hari ini ya!_ 😊',
  '🔥 _Semangat terus, kamu hebat!_ 👏',
  '🌟 _Percayalah pada diri sendiri, kamu bisa!_ 💪',
  '☀️ _Tetap positif, hal baik akan datang!_ 🌈',
  '💝 _Berkah selalu menyertai orang yang bersyukur_ 🙏',
  '🌺 _Mulailah hari dengan senyuman terbaikmu!_ 😊',
  '💫 _Mimpi besar dimulai dari hari ini!_ 🌟',
];

const TANYA_KABAR = [
  'Bagaimana kabarnya hari ini? 😊',
  'Apa kabar? Semoga sehat selalu ya 💪',
  'Gimana kabarnya? Semoga baik-baik saja 🌸',
  'Bagaimana kabar? Semoga harimu menyenangkan 😊',
  'Kabar baik kah hari ini? 🌟',
  'Bagaimana hari ini? Lancar semua? ✨',
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

const KATA_TERIMAKASIH = ['terima kasih','terimakasih','makasih','thanks','thank you','thx','tq','ty','tengkyu','mksh','trims'];
function isTerimakasih(low) { return KATA_TERIMAKASIH.some(function(k) { return cocokKata(low, k); }); }

function balasSapaanPintar(sender, kataSapaan) {
  const nama = getNama(sender);
  const waktu = getWaktu();
  const dataSapa = SAPAAN_MAP[kataSapaan];
  if (!dataSapa) return (nama ? 'Halo, *' + nama + '*! 😊' : 'Halo! 😊');
  const pilihan = dataSapa.balasan[Math.floor(Math.random() * dataSapa.balasan.length)];
  let respon = '';
  if (dataSapa.kategori === 'waktu') {
    if (dataSapa.waktu === waktu) respon = pilihan + (nama ? ', *' + nama + '*' : '') + '! 😊';
    else respon = pilihan + (nama ? ', *' + nama + '*' : '') + '! 😊\n_(Sekarang udah ' + waktu.toLowerCase() + ' loh)_';
  } else if (dataSapa.kategori === 'islami') respon = pilihan + (nama ? ', *' + nama + '*' : '') + ' 🤲';
  else if (dataSapa.kategori === 'sopan')   respon = pilihan + (nama ? ', *' + nama + '*' : '') + ' 😊';
  else                                       respon = pilihan + (nama ? ', *' + nama + '*' : '') + '! 😊';
  return respon;
}

function sapaanPertama(sender) {
  const nama = getNama(sender);
  const waktu = getWaktu();
  const motiv = KALIMAT_MOTIVASI[Math.floor(Math.random() * KALIMAT_MOTIVASI.length)];
  const kabar = TANYA_KABAR[Math.floor(Math.random() * TANYA_KABAR.length)];
  const sambutan = ['Selamat ' + waktu, 'Halo, selamat ' + waktu, 'Hai'];
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
  const opsi = ['Sama-sama' + n + '! 😊', 'Dengan senang hati' + n + '! 😊', 'Tentu' + n + '! 😊', 'Senang bisa membantu' + n + '! ✨'];
  return opsi[Math.floor(Math.random() * opsi.length)];
}

function isAdminCommand(low) {
  if (['listmember','listkontak','reload','info','resetall'].indexOf(low) >= 0) return true;
  return ADMIN_COMMANDS.some(function(cmd) { return low.startsWith(cmd + ' '); });
}

// ════════════════════════════════════════════════════════════════
//   13. ★★★ SCAN FOTO HELPER ★★★
// ════════════════════════════════════════════════════════════════

function msgMintaFoto(tokoKode, stepIdx, namaToko, dataWizard) {
  const steps = SCAN_STEPS[tokoKode];
  if (!steps || stepIdx >= steps.length) return null;
  const stepInfo = steps[stepIdx];
  const totalSteps = steps.length;
  const no = stepIdx + 1;
  
  let m = '📸 *SCAN FOTO LAPORAN*\n';
  m += '🏦 ' + namaToko + '\n';
  m += '━━━━━━━━━━━━━━━━━━\n';
  m += '📊 *Foto ' + no + ' dari ' + totalSteps + '*\n';
  const progress = Math.floor((no / totalSteps) * 10);
  m += '[';
  for (let i = 0; i < 10; i++) m += i < progress ? '█' : '░';
  m += '] ' + Math.round((no / totalSteps) * 100) + '%\n';
  m += '━━━━━━━━━━━━━━━━━━\n\n';
  
  if (stepInfo.scanField === 'manual') {
    m += '🅿️ *' + stepInfo.label + '*\n\n';
    m += '⌨️ Ketik nominal ' + stepInfo.label + ':\n';
    m += '   (ketik angka atau *-* jika kosong)\n\n';
  } else {
    m += '📸 *Kirim foto untuk:*\n';
    m += '🏷️ *' + stepInfo.label + '*\n\n';
    
    if (stepInfo.scanField === 'total_transaksi') {
      m += '💡 Saya akan scan angka *Total Transaksi*\n   dari tabel iPos yang kamu kirim\n\n';
    } else if (stepInfo.scanField === 'multi' || stepInfo.scanField === 'multi_promo') {
      m += '💡 Saya akan scan beberapa data sekaligus:\n';
      stepInfo.fields.forEach(function(f) {
        m += '   • ' + f + '\n';
      });
      m += '\n';
    }
    
    m += '📸 *Kirim foto sekarang*\n';
    m += '⌨️ Atau ketik angka manual\n';
  }
  
  // Tampilkan progress data
  const filled = Object.keys(dataWizard).filter(function(k) { return !k.startsWith('_') && dataWizard[k] !== undefined; });
  if (filled.length > 0) {
    m += '\n✅ *Data terisi: ' + filled.length + ' field*\n';
  }
  
  m += '\n━━━━━━━━━━━━━━━━━━\n';
  m += '🔙 *batal* | ⏭️ *skip* | 👁️ *review* | 💬 *selesai*';
  return m;
}

function wizardScanReview(tokoKode, dataWizard, namaToko) {
  const steps = SCAN_STEPS[tokoKode];
  let m = '👁️ *REVIEW DATA SCAN*\n🏦 ' + namaToko + '\n' + GARIS_TEBAL + '\n\n';
  
  // Tampilkan semua field berdasarkan FIELD_LAPORAN
  const fields = FIELD_LAPORAN[tokoKode];
  fields.forEach(function(f) {
    const v = dataWizard[f.key];
    let label = f.label;
    if (tokoKode === 'cp' && KASIR_CP_DEFAULT[f.key]) label = f.label + ' (' + KASIR_CP_DEFAULT[f.key] + ')';
    if (v === undefined) m += '⬜ ' + label + ': _(belum diisi)_\n';
    else if (v === 0) m += '✅ ' + label + ': Rp. -\n';
    else m += '✅ ' + label + ': Rp. ' + parseInt(v).toLocaleString('id-ID') + '\n';
  });
  
  // Khusus CP: tampilkan total parkir
  if (tokoKode === 'cp' && (dataWizard.parkirkomputer !== undefined || dataWizard.parkirluar !== undefined)) {
    const totalParkir = (dataWizard.parkirkomputer || 0) + (dataWizard.parkirluar || 0);
    m += '\n🅿️ Total Parkir: Rp. ' + totalParkir.toLocaleString('id-ID') + ' (otomatis)\n';
  }
  
  m += '\n' + GARIS_TEBAL + '\n';
  m += '💬 *lanjut* — lanjut scan\n';
  m += '💬 *selesai* — generate laporan\n';
  m += '🔙 *batal* — batalkan';
  return m;
}
// ════════════════════════════════════════════════════════════════
//   14. MENU FRIENDLY
// ════════════════════════════════════════════════════════════════

function getMenuUtama(nomor) {
  const nama = getNama(nomor);
  const salam = nama ? '*' + nama + '*' : 'Kamu';
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  🤖 *BOT TOKO PERABOT*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += 'Halo ' + salam + '! 👋\nSilakan pilih menu:\n\n';
  m += '┌─────────────────────\n';
  
  // Menu Laporan (hanya untuk role laporan & admin)
  if (bisaAksesLaporan(nomor)) {
    m += '│ ' + emojiNum(1) + ' 📊 Laporan Penjualan\n';
    m += '│ ' + emojiNum(2) + ' 🏷️ Laporan Harga Barang\n';
    m += '│ ' + emojiNum(3) + ' 🛒 Laporan Marketplace\n';
  }
  
  // Menu Cari (untuk semua member)
  if (isMember(nomor)) m += '│ ' + emojiNum(4) + ' 🔍 Cari Harga Barang\n';
  
  // Menu Admin
  if (isAdmin(nomor)) m += '│ ' + emojiNum(9) + ' 👑 Menu Admin\n';
  
  m += '└─────────────────────\n\n';
  m += '💬 *Cara pilih:*\n   Ketik nomor (contoh: *' + (bisaAksesLaporan(nomor) ? '1' : '4') + '*)\n\n';
  
  if (isMember(nomor)) {
    m += '🤖 *Tanya AI:*\n   _"Stok dandang eagle 20 di NK?"_\n\n';
    m += '📊 *Banding Harga:*\n   _"Bandingkan harga NN00001"_';
  }
  return m;
}
function getMenuPilihToko(menuType) {
  const ic = menuType === 1 ? '📊' : menuType === 2 ? '🏷️' : menuType === 'cari' ? '🔍' : '🛒';
  const jd = menuType === 1 ? 'LAPORAN PENJUALAN' : menuType === 2 ? 'LAPORAN HARGA BARANG' : menuType === 'cari' ? 'CARI HARGA BARANG' : 'LAPORAN MARKETPLACE';
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  ' + ic + ' *' + jd + '*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🏦 *Pilih Toko:*\n\n┌─────────────────────\n';
  TOKO_LIST.forEach(function(t, i) { m += '│ ' + emojiNum(i + 1) + ' ' + t.nama + '\n'; });
  m += '└─────────────────────\n\n💬 Ketik nomor (1-5)\n\n🔙 Ketik *batal* untuk kembali';
  return m;
}

function getMenuPilihHari(namaToko) {
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  🏦 *' + namaToko + '*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '📅 *Laporan untuk hari:*\n\n┌─────────────────────\n';
  m += '│ ' + emojiNum(1) + ' 📅 *HARI INI* _(' + getTanggal(false) + ')_\n│\n';
  m += '│ ' + emojiNum(2) + ' 📅 *KEMARIN* _(' + getTanggal(true) + ')_\n';
  m += '└─────────────────────\n\n💬 Ketik *1* atau *2*\n\n🔙 Ketik *batal* untuk kembali';
  return m;
}

function getMenuSiapInputMarket(namaToko, kemarin) {
  const t = getTanggal(kemarin); const k = kemarin ? ' _(kemarin)_' : '';
  const contoh = 'oesapa 0\ntdm 0\ncentral 21061000\nwa 21061000\nshopee 0\ntiktok 0\ntokopedia 0\ntunai 304000\ndebit 20757000';
  return '╭━━━━━━━━━━━━━━━━━╮\n│  ✅ *SIAP INPUT*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n🏦 ' + namaToko + '\n📅 ' + t + k + '\n\n📸 *Kirim FOTO* atau *ketik manual:*\n\n```' + contoh + '```\n\n🔙 Ketik *batal* untuk membatalkan';
}

function getMenuSiapInputHarga(namaToko, kemarin) {
  const t = getTanggal(kemarin); const k = kemarin ? ' _(kemarin)_' : '';
  const contoh = '---baru---\nNama barang baru\n---naik---\nNama barang naik\n---turun---\nNama barang turun';
  return '╭━━━━━━━━━━━━━━━━━╮\n│  ✅ *SIAP INPUT*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n🏦 ' + namaToko + '\n📅 ' + t + k + '\n\n📸 *Kirim FOTO* atau *ketik manual:*\n\n```' + contoh + '```\n\n🔙 Ketik *batal* untuk membatalkan';
}

function getMenuAdmin() {
  let m = '╭━━━━━━━━━━━━━━━━━╮\n│  👑 *MENU ADMIN*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n🛠️ *Pilih aksi:*\n\n┌─────────────────────\n';
  m += '│ ' + emojiNum(1) + ' 👥 List Member\n│ ' + emojiNum(2) + ' 📒 List Kontak\n';
  m += '│ ' + emojiNum(3) + ' ➕ Tambah Member\n│ ' + emojiNum(4) + ' ➖ Hapus Member\n';
  m += '│ ' + emojiNum(5) + ' ✏️ Set Nama Kontak\n│ ' + emojiNum(6) + ' 🗑️ Hapus Kontak\n';
  m += '│ ' + emojiNum(7) + ' 🔄 Reload Excel\n│ ' + emojiNum(8) + ' ℹ️ Info Sistem\n';
  m += '└─────────────────────\n\n💬 Ketik nomor (1-8)\n\n🔙 Ketik *batal* untuk kembali';
  return m;
}

function getMenuSiapCari(namaToko) {
  return '╭━━━━━━━━━━━━━━━━━╮\n│  🔍 *CARI BARANG*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n🏦 ' + namaToko + '\n\n━━━━━━━━━━━━━━━━━━\n⌨️ *Ketik nama atau kode:*\n━━━━━━━━━━━━━━━━━━\n\n💡 *Contoh:*\n   • _dandang eagle 20_\n   • _NN00001_\n\n✨ Bot otomatis koreksi typo!\n\n🔙 Ketik *batal* untuk kembali';
}

function getMenuCariUlang(namaToko) {
  return '━━━━━━━━━━━━━━━━━━\n🔍 *Cari lagi di ' + namaToko + '?*\n━━━━━━━━━━━━━━━━━━\n\n⌨️ Ketik nama/kode barang lain\n\n┌─────────────────────\n│ ' + emojiNum(9) + ' 🔄 Ganti toko\n│ *batal* 🔙 Menu utama\n└─────────────────────';
}

// ════════════════════════════════════════════════════════════════
//   15. WIZARD TEXT HELPERS
// ════════════════════════════════════════════════════════════════

function wizardToText(dataWizard, tokoKode) {
  let text = '';
  Object.keys(dataWizard).forEach(function(key) {
    if (key.startsWith('_')) return;
    if (key === 'total') return;
    const v = dataWizard[key];
    if (tokoKode === 'cp' && KASIR_CP_DEFAULT[key]) {
      text += key + ' ' + KASIR_CP_DEFAULT[key] + ' ' + (v || 0) + '\n';
    } else {
      text += key + ' ' + (v || 0) + '\n';
    }
  });
  return text;
}

// ════════════════════════════════════════════════════════════════
//   16. PARSER PINTAR
// ════════════════════════════════════════════════════════════════

function parsePilihanMenu(low) {
  if (low === '1' || low === 'satu') return 1;
  if (low === '2' || low === 'dua')  return 2;
  if (low === '3' || low === 'tiga') return 3;
  if (low === '4' || low === 'empat') return 4;
  if (low === '9' || low === 'admin') return 9;
  if (low.indexOf('penjualan') >= 0) return 1;
  if (low.indexOf('harga') >= 0 && low.indexOf('barang') >= 0) return 2;
  if (low.indexOf('marketplace') >= 0) return 3;
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
    for (let j = 0; j < t.alias.length; j++) {
      if (low === t.alias[j] || low.indexOf(t.alias[j]) >= 0) return t;
    }
  }
  return null;
}

function parsePilihanHari(low) {
  if (low === '1' || low.indexOf('hari ini') >= 0) return false;
  if (low === '2' || low.indexOf('kemarin') >= 0) return true;
  return null;
}

function parsePilihanAdmin(low) {
  const map = { '1':'listmember','2':'listkontak','3':'daftar','4':'hapus','5':'namakontak','6':'hapuskontak','7':'reload','8':'info' };
  if (map[low]) return map[low];
  if (low === 'listmember' || low.indexOf('list member') >= 0) return 'listmember';
  if (low === 'listkontak' || low.indexOf('list kontak') >= 0) return 'listkontak';
  if (low.indexOf('tambah') >= 0 || low === 'daftar') return 'daftar';
  if (low === 'hapus') return 'hapus';
  if (low === 'namakontak' || low.indexOf('set nama') >= 0) return 'namakontak';
  if (low.indexOf('hapus kontak') >= 0) return 'hapuskontak';
  if (low === 'reload') return 'reload';
  if (low === 'info') return 'info';
  return null;
}

// ════════════════════════════════════════════════════════════════
//   17. FORMAT HASIL CARI
// ════════════════════════════════════════════════════════════════

function formatHasil(searchResult, tokoKode, sender) {
  const namaToko = NAMA_TOKO[tokoKode] || '-';
  const nama = getNama(sender);
  const suffix = '\n\n' + (nama ? 'Semoga membantu, *' + nama + '*! 😊' : 'Semoga membantu! 😊');
  const items = searchResult.hasil || [];
  const saran = searchResult.saran || [];
  const tipeHasil = searchResult.tipeHasil || 'kosong';
  const totalDitemukan = searchResult.totalDitemukan || items.length;

  if (tipeHasil === 'kosong') return '❌ *Barang Tidak Ditemukan*\n' + GARIS_TEBAL + '\n\n💡 *Tips:*\n• Kata kunci lebih singkat\n• Cari by kode (_NN00001_)';
  if (tipeHasil === 'saran' && items.length === 0 && saran.length > 0) {
    let msg = '🤔 *Tidak Ditemukan Persis*\n' + GARIS_TEBAL + '\n🏦 *' + namaToko + '*\n\n💡 Mungkin:\n\n';
    saran.forEach(function(d, i) {
      msg += '*' + (i + 1) + '.* ' + d.nama + '\n   🔖 _' + d.kode + '_ | 💰 ' + fRp(d.harga[tokoKode].ecer) + '\n';
      if (i < saran.length - 1) msg += GARIS_TIPIS + '\n';
    });
    return msg + suffix;
  }
  if (items.length === 1) {
    const d = items[0]; const h = d.harga[tokoKode];
    let msg = '🏷️ *Detail Barang*\n🏦 *' + namaToko + '*\n' + GARIS_TEBAL + '\n';
    msg += '🔖 *Kode*: ' + d.kode + '\n📦 *Nama*: ' + d.nama + '\n';
    msg += '🏷️ *Jenis*: ' + (d.jenis || '-') + '\n🏗️ *Merek*: ' + (d.merek || '-') + '\n📏 *Satuan*: ' + d.satuan + '\n' + GARIS_TEBAL + '\n';
    msg += '💰 *Ecer (1-5 Pcs)*: ' + fRp(h.ecer) + '\n💰 *Ambil (6 Pcs Keatas)*: ' + fRp(h.ambil) + '\n📊 *Stok*: ' + (h.stok > 0 ? h.stok + ' ' + d.satuan : '⚠️ Kosong') + '\n' + GARIS_TEBAL;
    if (tipeHasil === 'fuzzy') msg += '\n\n💡 _Hasil koreksi otomatis_';
    return msg + suffix;
  }
  let header = totalDitemukan > CONFIG.maxHasilCari ? '🔍 *' + items.length + ' dari ' + totalDitemukan + ' Barang*\n' : '🔍 *Ditemukan ' + items.length + ' Barang*\n';
  let msg = header + '🏦 *' + namaToko + '*\n' + GARIS_TEBAL + '\n';
  items.forEach(function(d, i) {
    const h = d.harga[tokoKode];
    msg += '*' + (i + 1) + '.* ' + d.nama + '\n';
    msg += '   🔖 ' + d.kode + ' | ' + d.satuan + '\n';
    msg += '   💰 Ecer(1-5 Pcs): ' + fRp(h.ecer) + '\n';
    msg += '   💰 Ambil(6 Pcs+): ' + fRp(h.ambil) + '\n';
    msg += '   📊 Stok: ' + (h.stok > 0 ? h.stok + ' ' + d.satuan : '⚠️ Kosong') + '\n';
    if (i < items.length - 1) msg += GARIS_TIPIS + '\n';
  });
  msg += GARIS_TEBAL;
  if (totalDitemukan > CONFIG.maxHasilCari) msg += '\n\n⚠️ +' + (totalDitemukan - CONFIG.maxHasilCari) + ' barang lagi';
  return msg + suffix;
}
// ════════════════════════════════════════════════════════════════
//   18. GENERATOR LAPORAN
// ════════════════════════════════════════════════════════════════

function genLapPenjualan(text, namaToko, kemarin, tokoKode) {
  if (tokoKode === 'cp') return genLapPenjualanCP(text, kemarin);
  if (tokoKode === 'nk') return genLapPenjualanStandar(text, kemarin, 'Nasional Kitchen', 'fullEcer');
  if (tokoKode === 'tdm') return genLapPenjualanStandar(text, kemarin, 'Perabot Mama TDM', 'noEcer');
  if (tokoKode === 'oesapa') return genLapPenjualanStandar(text, kemarin, 'Perabot Mama Oesapa', 'noEcer');
  if (tokoKode === 'kefa') return genLapPenjualanStandar(text, kemarin, 'Perabot Mamaku Kefamenanu', 'noEcer');
  return 'Format tidak dikenali';
}

function genLapPenjualanStandar(text, kemarin, namaTokoFull, mode) {
  const t = getTanggal(kemarin);
  const d = { k1:0, k2:0, k3:0, tunai:0, debit:0, kredit:0, ecer:0, grosir:0 };
  text.trim().split('\n').forEach(function(line) {
    const tr = line.trim(); if (!tr) return;
    const parts = tr.split(/\s+/); if (parts.length < 2) return;
    const key = parts[0].toLowerCase();
    const rawValue = parts.slice(1).join(' ').trim();
    let value = 0;
    if (rawValue !== '-' && rawValue !== 'kosong' && rawValue !== 'null') {
      value = parseFloat(rawValue.replace(/[^0-9]/g, '')) || 0;
    }
    if (key === 'k1') d.k1 = value;
    else if (key === 'k2') d.k2 = value;
    else if (key === 'k3') d.k3 = value;
    else if (key === 'tunai') d.tunai = value;
    else if (key === 'debit') d.debit = value;
    else if (key === 'kredit' || key === 'credit') d.kredit = value;
    else if (key === 'ecer') d.ecer = value;
    else if (key === 'grosir') d.grosir = value;
  });
  function fr(n) { const v = parseFloat(n) || 0; return v === 0 ? 'Rp. -' : 'Rp. ' + v.toLocaleString('id-ID'); }
  const totalUtama = d.k1 + d.k2 + d.k3;
  let msg = 'Laporan Penjualan\nToko ' + namaTokoFull + '\nPeriode ' + t + '\n\n';
  msg += 'Kassa 1 ' + fr(d.k1) + '\nKassa 2 ' + fr(d.k2) + '\n';
  if (d.k3 > 0) msg += 'Kassa 3 ' + fr(d.k3) + '\n';
  msg += '\nTotal Penjualan Keseluruhan\n' + fr(totalUtama) + '\n---------------------------------------------\n\n';
  msg += 'Tunai  ' + fr(d.tunai) + '\nDebit  ' + fr(d.debit) + '\nCredit ' + fr(d.kredit) + '\n---------------------------------------------\n';
  if (mode === 'fullEcer' || d.ecer > 0 || d.grosir > 0) {
    msg += 'Ecer : ' + fr(d.ecer) + '\nGrosir : ' + fr(d.grosir) + '\n--------------------------------------------';
  }
  return msg;
}

function genLapPenjualanCP(text, kemarin) {
  const t = getTanggal(kemarin);
  const d = { k1:0, k2:0, k3:0, k4:0, nk1:'', nk2:'', nk3:'', nk4:'', tunai:0, debit:0, kredit:0, ecer:0, grosir:0, promo:0, promoTunai:0, promoDebit:0, promoKredit:0, parkirKomputer:0, parkirLuar:0 };
  text.trim().split('\n').forEach(function(line) {
    const tr = line.trim(); if (!tr) return;
    const parts = tr.split(/\s+/); if (parts.length < 2) return;
    const key = parts[0].toLowerCase();
    let valueStart = 1; let kasirNama = '';
    if (['k1','k2','k3','k4'].indexOf(key) >= 0 && parts.length >= 3) {
      const possibleNum = parts[1].replace(/[^0-9]/g, '');
      if (!possibleNum || possibleNum.length < 4) {
        for (let i = 1; i < parts.length; i++) {
          const num = parts[i].replace(/[^0-9]/g, '');
          if (num && num.length >= 4) { valueStart = i; kasirNama = parts.slice(1, i).join(' '); break; }
        }
      }
    }
    const rawValue = parts.slice(valueStart).join(' ').trim();
    let value = 0;
    if (rawValue !== '-' && rawValue !== 'kosong' && rawValue !== 'null') {
      value = parseFloat(rawValue.replace(/[^0-9]/g, '')) || 0;
    }
    if (key === 'k1') { d.k1 = value; if (kasirNama) d.nk1 = kasirNama; }
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
  function fr(n) { const v = parseFloat(n) || 0; return v === 0 ? 'Rp. -' : 'Rp. ' + v.toLocaleString('id-ID'); }
  if (!d.nk1) d.nk1 = KASIR_CP_DEFAULT.k1; if (!d.nk2) d.nk2 = KASIR_CP_DEFAULT.k2;
  if (!d.nk3) d.nk3 = KASIR_CP_DEFAULT.k3; if (!d.nk4) d.nk4 = KASIR_CP_DEFAULT.k4;
  const totalUtama = d.k1 + d.k2 + d.k3 + d.k4;
  const totalParkir = d.parkirKomputer + d.parkirLuar;
  let msg = 'Laporan Penjualan Toko Central Perabot Alak Periode ' + t + '\n\n';
  msg += 'Kassa 1 (' + d.nk1 + ') ' + fr(d.k1) + '\nKassa 2 (' + d.nk2 + ') ' + fr(d.k2) + '\n';
  msg += 'Kassa 3 (' + d.nk3 + ') ' + fr(d.k3) + '\nKassa 4 (' + d.nk4 + ') ' + fr(d.k4) + '\n\n';
  msg += 'Total Penjualan Keseluruhan: ' + fr(totalUtama) + '\n---------------------------------------------\n\n';
  msg += 'Tunai  ' + fr(d.tunai) + '\nDebit  ' + fr(d.debit) + '\nCredit ' + fr(d.kredit) + '\n';
  msg += '---------------------------------------------\n---------------------------------------------\n\n';
  msg += 'Ecer: ' + fr(d.ecer) + '\nGrosir : ' + fr(d.grosir) + '\n\n---------------------------------------------\n\n';
  msg += 'Laporan Penjualan Kasir Promo\nPeriode ' + t + '\n\n';
  msg += 'Total Penjualan Keseluruhan: ' + fr(d.promo) + '\n---------------------------------------------\n';
  msg += 'Tunai  ' + fr(d.promoTunai) + '\nDebit  ' + fr(d.promoDebit) + '\nCredit ' + fr(d.promoKredit) + '\n---------------------------------------------\n\n';
  msg += 'Laporan Parkir \nPeriode ' + t + '\n\n';
  msg += 'Parkir di Komputer : ' + fr(d.parkirKomputer) + '\nParkir Stor Luar : ' + fr(d.parkirLuar) + '\n---------------------------------------------\n';
  msg += 'Total Parkir  ' + fr(totalParkir) + '\n---------------------------------------------';
  return msg;
}

function genLapHarga(text, namaToko, kemarin) {
  const t = getTanggal(kemarin); const h = kemarin ? 'Kemarin' : 'Ini'; const k = kemarin ? ' _(kemarin)_' : '';
  const s = 'Selamat ' + getWaktu() + ' Team ' + namaToko;
  const d = { baru:[], naik:[], turun:[], note:[] }; let mode = null;
  text.trim().split('\n').forEach(function(line) {
    const tr = line.trim(); if (!tr) return; const lo = tr.toLowerCase();
    if (lo.indexOf('---baru---') >= 0 || lo === 'baru') { mode = 'baru'; return; }
    if (lo.indexOf('---naik---') >= 0 || lo === 'naik') { mode = 'naik'; return; }
    if (lo.indexOf('---turun---') >= 0 || lo === 'turun') { mode = 'turun'; return; }
    if (lo.indexOf('---note---') >= 0 || lo === 'note') { mode = 'note'; return; }
    if (mode) d[mode].push(tr);
  });
  const cat = 'Nota Semuanya Sudah Diinput Di Sistem, Bisa Langsung Di Print Barcodenya Ya.\n\nMohon Dicek Kembali Fisik Barang Dengan Yang Di Input Disistem, Jika Ada Yang Tidak Sesuai Mohon Di Konfirmasi Lagi. Terima Kasih🙏🏼';
  let msg = s + '\n\nHarga Barang Untuk Hari ' + h + ' *' + t + '*' + k + '\n';
  if (d.baru.length > 0) { msg += '\n🆕 *Barang Yang Baru:*\n'; d.baru.forEach(function(b) { msg += '• ' + b + '\n'; }); }
  if (d.naik.length > 0) { msg += '\n📈 *Barang Yang Naik Harga:*\n'; d.naik.forEach(function(b) { msg += '• ' + b + '\n'; }); }
  if (d.turun.length > 0) { msg += '\n📉 *Barang Yang Turun Harga:*\n'; d.turun.forEach(function(b) { msg += '• ' + b + '\n'; }); }
  if (d.note.length > 0) { msg += '\n📝 *Catatan:*\n'; d.note.forEach(function(b) { msg += b + '\n'; }); }
  return msg + '\n' + cat;
}

function genLapMarket(text, kemarin) {
  const t = getTanggal(kemarin); const k = kemarin ? ' _(kemarin)_' : '';
  const d = { oesapa:0, tdm:0, central:0, wa:0, shopee:0, tiktok:0, tokopedia:0, tunai:0, debit:0, kredit:0, nota:[] };
  text.trim().toLowerCase().split('\n').forEach(function(line) {
    const tr = line.trim(); if (!tr) return;
    if (tr.indexOf('nota ') === 0) { d.nota.push(line.trim().substring(5)); return; }
    const p = tr.split(/\s+/);
    if (p.length >= 2 && p[0] in d) d[p[0]] = parseFloat(p.slice(1).join('').replace(/[^0-9]/g, '')) || 0;
  });
  const tT = d.oesapa + d.tdm + d.central;
  const tC = d.wa + d.shopee + d.tiktok + d.tokopedia;
  let nt = ''; if (d.nota.length > 0) { nt = '\n'; d.nota.forEach(function(n) { nt += '- Nomor Nota ' + n + '\n'; }); }
  return GARIS_TEBAL + '\n🛒 *Total Penjualan Marketplace*\n*Perabot Mama*\n📅 Periode ' + t + k +
    '\n' + GARIS_TEBAL + '\n🏦 *Per Toko*\n• Oesapa: ' + fRp(d.oesapa) + '\n• TDM: ' + fRp(d.tdm) + '\n• Central: ' + fRp(d.central) +
    '\n' + GARIS_TIPIS + '\n💰 *Total*: ' + fRp(tT) + '\n\n📱 *Per Channel*\n• WA: ' + fRp(d.wa) + '\n• Shopee: ' + fRp(d.shopee) +
    '\n• Tiktok: ' + fRp(d.tiktok) + '\n• Tokopedia: ' + fRp(d.tokopedia) +
    '\n' + GARIS_TIPIS + '\n💰 *Total*: ' + fRp(tC) + '\n\n💳 *Metode Bayar*\n• Tunai: ' + fRp(d.tunai) + '\n• Debit: ' + fRp(d.debit) + '\n• Credit: ' + fRp(d.kredit) +
    '\n' + GARIS_TEBAL + '\n' + nt + '_Laporan otomatis_';
}

// ════════════════════════════════════════════════════════════════
//   19. HANDLER ADMIN
// ════════════════════════════════════════════════════════════════

async function handleAdmin(sender, msg, low) {
  log.info('ADMIN', 'Cmd: ' + low);
  if (low.startsWith('daftar ')) { const nomor = msg.substring(7).trim().replace(/[^0-9]/g, ''); if (!nomor) { await kirimWA(sender, '⚠️ Format: daftar 628xxx'); return true; } const r = tambahMember(nomor); if (!r.ok) { await kirimWA(sender, '⚠️ ' + r.alasan); return true; } await kirimWA(sender, '✅ Member terdaftar! ' + nomor + '\n📊 Total: ' + MEMBERS.length + '/' + CONFIG.maxMember); return true; }
  if (low.startsWith('hapus ')) { const nomor = msg.substring(6).trim().replace(/[^0-9]/g, ''); const r = hapusMember(nomor); if (!r.ok) { await kirimWA(sender, '❌ ' + r.alasan); return true; } await kirimWA(sender, '✅ Member ' + nomor + ' dihapus!'); return true; }
  if (low === 'listmember') { try { let m = 'Daftar Member (' + MEMBERS.length + '/' + CONFIG.maxMember + ')\n------------------\n'; if (MEMBERS.length === 0) m += '(kosong)\n'; else for (let i = 0; i < MEMBERS.length; i++) { m += (i+1) + '. ' + MEMBERS[i] + '\n   ' + (KONTAK[MEMBERS[i]] || '(belum ada nama)') + '\n'; } m += '------------------\nSlot tersisa: ' + (CONFIG.maxMember - MEMBERS.length); await kirimWA(sender, m); } catch (e) { await kirimWA(sender, 'Error: ' + e.message); } return true; }
  if (low.startsWith('namakontak ')) { const p = msg.substring(11).trim().split(/\s+/); if (p.length < 2) { await kirimWA(sender, '⚠️ Format: namakontak 628xxx Nama'); return true; } const nomor = p[0].replace(/[^0-9]/g, ''); const nama = p.slice(1).join(' '); const r = setNama(nomor, nama); if (!r.ok) { await kirimWA(sender, '❌ ' + r.alasan); return true; } await kirimWA(sender, '✅ ' + nomor + ' → ' + nama); return true; }
  if (low.startsWith('hapuskontak ')) { const nomor = msg.substring(12).trim().replace(/[^0-9]/g, ''); const r = hapusKontak(nomor); if (!r.ok) { await kirimWA(sender, '❌ ' + r.alasan); return true; } await kirimWA(sender, '✅ Kontak ' + nomor + ' dihapus!'); return true; }
  if (low === 'listkontak') { const keys = Object.keys(KONTAK); let m = 'Daftar Kontak (' + keys.length + ')\n------------------\n'; keys.forEach(function(k, i) { m += (i+1) + '. ' + k + '\n   ' + KONTAK[k] + '\n'; }); await kirimWA(sender, m); return true; }
  if (low === 'reload') { const ok = loadExcel(); await kirimWA(sender, ok ? '✅ Reloaded! ' + DATA_BARANG.length + ' item' : '❌ Gagal'); return true; }
  if (low === 'resetall') { SESI = {}; saveJSON(CONFIG.paths.sesi, SESI); await kirimWA(sender, '✅ Semua sesi direset!'); return true; }
  if (low === 'info') { const up = Math.floor(process.uptime()); const jam = Math.floor(up/3600), mnt = Math.floor((up%3600)/60); let ai = ''; if (CONFIG.groqKey) ai += '\n  ✅ Groq'; if (CONFIG.geminiKey) ai += '\n  ✅ Gemini (' + geminiKeys().length + ' keys)'; if (CONFIG.openrouterKey) ai += '\n  ✅ OpenRouter'; if (!ai) ai = '\n  ⚠️ Tidak ada'; await kirimWA(sender, 'ℹ️ Info Sistem v3.13\n------------------\nUptime: ' + jam + 'j ' + mnt + 'm\nMember: ' + MEMBERS.length + '/' + CONFIG.maxMember + '\nData: ' + DATA_BARANG.length + ' item\nSesi: ' + Object.keys(SESI).length + '\n\n🤖 AI:' + ai); return true; }
  return false;
}

// ════════════════════════════════════════════════════════════════
//   20. ROUTES
// ════════════════════════════════════════════════════════════════

app.get('/', function(req, res) { res.json({ status: 'ok', app: CONFIG.appName + ' v3.13', items: DATA_BARANG.length, members: MEMBERS.length + '/' + CONFIG.maxMember }); });
app.get('/reload', function(req, res) { const ok = loadExcel(); res.json({ success: ok, total: DATA_BARANG.length }); });
app.get('/resetsesi/:nomor', function(req, res) { resetSesi(req.params.nomor); res.json({ ok: true }); });
app.get('/resetall', function(req, res) { SESI = {}; saveJSON(CONFIG.paths.sesi, SESI); res.json({ ok: true, message: 'Semua sesi direset' }); });

// ════════════════════════════════════════════════════════════════
//   21. ★★★ WEBHOOK UTAMA (dengan SCAN FOTO MULTI-STEP) ★★★
// ════════════════════════════════════════════════════════════════

const KATA_RESET = ['batal','menu','mulai','start','kembali','home','keluar','exit','stop','reset'];

app.post('/webhook', async function(req, res) {
  res.sendStatus(200);

  try {
    const body = req.body || {};
    const sender = body.sender || body.from || body.phone || null;
    const message = (body.message || body.text || body.msg || '').trim();
        const image = body.image || body.file || body.media || body.url || body.attachment || body.fileUrl || body.mediaUrl || '';
    if (!sender) return;
    const msg = message;
    const low = msg.toLowerCase();

        log.info('WEBHOOK', sender + ': ' + (image ? '[FOTO: ' + image.substring(0, 50) + '] ' : '') + msg.substring(0, 50));
    
    // DEBUG: log semua field body untuk cek field foto
    if (!image && body) {
      const bodyKeys = Object.keys(body).filter(function(k) { return body[k] && typeof body[k] === 'string' && (body[k].startsWith('http') || body[k].indexOf('.jpg') >= 0 || body[k].indexOf('.png') >= 0 || body[k].indexOf('.jpeg') >= 0); });
      if (bodyKeys.length > 0) {
        log.warn('WEBHOOK', 'POSSIBLE IMAGE FIELDS: ' + bodyKeys.map(function(k) { return k + '=' + String(body[k]).substring(0, 50); }).join(' | '));
      }
    }

    // ── SAPAAN PERTAMA ──
    if (!SUDAH_DISAPA[sender]) {
      SUDAH_DISAPA[sender] = true;
      saveJSON(CONFIG.paths.disapa, SUDAH_DISAPA);
      await kirimWA(sender, sapaanPertama(sender));
      await tunggu(1000);
      await kirimWA(sender, getMenuUtama(sender));
      return;
    }

    // ── ADMIN ──
    if (isAdmin(sender) && isAdminCommand(low)) {
      if (SESI[sender] && (SESI[sender].mode || SESI[sender].menu || SESI[sender].adminAksi)) resetSesi(sender);
      const handled = await handleAdmin(sender, msg, low);
      if (handled) return;
    }

    const _s = SESI[sender] || {};
    const _lagiInput = (_s.menu && (_s.kemarin !== undefined && _s.kemarin !== null)) || _s.wizardActive || _s.scanActive;

    // ── RESET ──
    if (KATA_RESET.indexOf(low) >= 0 || (low === '0' && !_lagiInput)) {
      resetSesi(sender);
      await kirimWA(sender, getMenuUtama(sender));
      return;
    }

    // ── SAPAAN ──
    const kataSapaan = isSapaan(low);
    if (kataSapaan && !_lagiInput && !_s.pendingPindahToko) {
      await kirimWA(sender, sapaanBerikutnya(sender, kataSapaan));
      await tunggu(800);
      await kirimWA(sender, getMenuUtama(sender));
      return;
    }
    if (isTerimakasih(low) && !_lagiInput && !_s.pendingPindahToko) {
      await kirimWA(sender, balasTerimakasih(sender));
      return;
    }

    const s = getSesi(sender);

    // ════════════════════════════════════════════════════════════
    //   ★★★ SCAN FOTO MODE ★★★
    // ════════════════════════════════════════════════════════════
    if (s.scanActive) {
      const scanData = s.scanData || {};
      const currentStepIdx = s.scanStepIdx || 0;
      const tokoKode = s.toko;
      const namaToko = NAMA_TOKO[tokoKode];
      const steps = SCAN_STEPS[tokoKode];
      
      if (!steps || currentStepIdx >= steps.length) {
        // Semua step selesai → generate laporan
        const text = wizardToText(scanData, tokoKode);
        const laporan = genLapPenjualan(text, namaToko, s.kemarin, tokoKode);
        if (laporan) {
          await kirimWA(sender, laporan);
          resetSesi(sender);
          await tunggu(1500);
          await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
        }
        return;
      }
      
      const stepInfo = steps[currentStepIdx];
      
      // Handle review
      if (low === 'review') {
        await kirimWA(sender, wizardScanReview(tokoKode, scanData, namaToko));
        return;
      }
      
      // Handle selesai paksa
      if (low === 'selesai' || low === 'finish' || low === 'done') {
        // Isi field yang belum ada dengan 0
        const allFields = FIELD_LAPORAN[tokoKode];
        allFields.forEach(function(f) { if (scanData[f.key] === undefined) scanData[f.key] = 0; });
        const text = wizardToText(scanData, tokoKode);
        const laporan = genLapPenjualan(text, namaToko, s.kemarin, tokoKode);
        if (laporan) {
          await kirimWA(sender, laporan);
          resetSesi(sender);
          await tunggu(1500);
          await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
        }
        return;
      }
      
      // Handle lanjut (setelah review)
      if (low === 'lanjut' || low === 'next' || low === 'continue') {
        await kirimWA(sender, msgMintaFoto(tokoKode, currentStepIdx, namaToko, scanData));
        return;
      }
      
      // Handle skip
      if (low === 'skip' || low === 'lewati') {
        stepInfo.fields.forEach(function(f) { scanData[f] = 0; });
        const nextIdx = currentStepIdx + 1;
        updateSesi(sender, { scanData: scanData, scanStepIdx: nextIdx });
        
        if (nextIdx >= steps.length) {
          // Semua selesai
          const text = wizardToText(scanData, tokoKode);
          const laporan = genLapPenjualan(text, namaToko, s.kemarin, tokoKode);
          if (laporan) {
            await kirimWA(sender, laporan);
            resetSesi(sender);
            await tunggu(1500);
            await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
          }
        } else {
          await tunggu(300);
          await kirimWA(sender, '⏭️ _Dilewati_\n\n' + msgMintaFoto(tokoKode, nextIdx, namaToko, scanData));
        }
        return;
      }
      
      // ★★★ HANDLE MANUAL INPUT (untuk step parkir) ★★★
      if (stepInfo.scanField === 'manual') {
        let nominal = 0;
        if (msg === '-' || low === 'kosong' || low === 'null' || low === '0') nominal = 0;
        else {
          const angka = msg.replace(/[^0-9]/g, '');
          if (!angka) { await kirimWA(sender, '⚠️ Ketik angka saja atau *-* untuk kosong'); return; }
          nominal = parseInt(angka);
        }
        
        scanData[stepInfo.fields[0]] = nominal;
        const nextIdx = currentStepIdx + 1;
        updateSesi(sender, { scanData: scanData, scanStepIdx: nextIdx });
        
        const fr = nominal === 0 ? 'Rp. -' : 'Rp. ' + nominal.toLocaleString('id-ID');
        
        if (nextIdx >= steps.length) {
          // Semua selesai
          await kirimWA(sender, '✅ _' + stepInfo.label + ': ' + fr + ' disimpan_\n\n🎉 *Semua data lengkap!*\n⏳ _Generate laporan..._');
          await tunggu(800);
          const text = wizardToText(scanData, tokoKode);
          const laporan = genLapPenjualan(text, namaToko, s.kemarin, tokoKode);
          if (laporan) {
            await kirimWA(sender, laporan);
            resetSesi(sender);
            await tunggu(1500);
            await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
          }
        } else {
          await tunggu(300);
          await kirimWA(sender, '✅ _' + stepInfo.label + ': ' + fr + ' disimpan_\n\n' + msgMintaFoto(tokoKode, nextIdx, namaToko, scanData));
        }
        return;
      }
      
      // ★★★ HANDLE FOTO ATAU ANGKA MANUAL ★★★
      let hasilScan = null;
      
      if (image && image.length > 0) {
        // SCAN FOTO
        await kirimWA(sender, '📸 _Scanning foto ' + stepInfo.label + '..._');
        try {
          const prompt = SCAN_PROMPTS[stepInfo.scanField] || SCAN_PROMPTS.total_transaksi;
          const aiResult = await analisaGambar(image, prompt);
          log.info('SCAN', 'AI result: ' + aiResult.substring(0, 100));
          
          if (stepInfo.scanField === 'total_transaksi') {
            const angka = parseScanSingle(aiResult);
            hasilScan = {};
            hasilScan[stepInfo.fields[0]] = angka;
          } else if (stepInfo.scanField === 'multi' || stepInfo.scanField === 'multi_promo') {
            hasilScan = parseScanMulti(aiResult, stepInfo.fields);
          }
        } catch (e) {
          log.error('SCAN', 'Gagal: ' + e.message);
          await kirimWA(sender, '❌ Gagal scan foto. Coba kirim ulang atau ketik angka manual.');
          return;
        }
      } else if (msg) {
        // MANUAL ANGKA
        if (stepInfo.fields.length === 1) {
          let nominal = 0;
          if (msg === '-' || low === 'kosong') nominal = 0;
          else {
            const angka = msg.replace(/[^0-9]/g, '');
            if (!angka) { await kirimWA(sender, '⚠️ Ketik angka saja atau kirim foto'); return; }
            nominal = parseInt(angka);
          }
          hasilScan = {};
          hasilScan[stepInfo.fields[0]] = nominal;
        } else {
          // Multi-field manual → parse format "field: angka"
          hasilScan = parseScanMulti(msg, stepInfo.fields);
        }
      } else return;
      
      if (hasilScan) {
        // Simpan hasil
        Object.keys(hasilScan).forEach(function(k) { scanData[k] = hasilScan[k]; });
        const nextIdx = currentStepIdx + 1;
        updateSesi(sender, { scanData: scanData, scanStepIdx: nextIdx });
        
        // Format konfirmasi
        let konfirmasi = '✅ *Data disimpan:*\n';
        Object.keys(hasilScan).forEach(function(k) {
          const v = hasilScan[k];
          const fr = v === 0 ? 'Rp. -' : 'Rp. ' + v.toLocaleString('id-ID');
          konfirmasi += '   • ' + k + ': ' + fr + '\n';
        });
        
        if (nextIdx >= steps.length) {
          await kirimWA(sender, konfirmasi + '\n🎉 *Semua data lengkap!*\n⏳ _Generate laporan..._');
          await tunggu(800);
          const text = wizardToText(scanData, tokoKode);
          const laporan = genLapPenjualan(text, namaToko, s.kemarin, tokoKode);
          if (laporan) {
            await kirimWA(sender, laporan);
            resetSesi(sender);
            await tunggu(1500);
            await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
          }
        } else {
          await tunggu(300);
          await kirimWA(sender, konfirmasi + '\n' + msgMintaFoto(tokoKode, nextIdx, namaToko, scanData));
        }
      }
      return;
    }

    // ════════════════════════════════════════════════════════════
    //   MODE CARI (dengan konfirmasi pindah toko)
    // ════════════════════════════════════════════════════════════
    if (s.mode === 'cari') {
      if (s.pendingPindahToko) {
        if (isKonfirmasiYa(low)) {
          const tokoBaru = s.pendingPindahToko.toko;
          const keyword = s.pendingPindahToko.keyword;
          updateSesi(sender, { tokoKode: tokoBaru.kode, pendingPindahToko: null });
          await kirimWA(sender, '✅ *Pindah ke ' + tokoBaru.nama + '*');
          await tunggu(500);
          if (keyword && keyword.length >= 2) {
            await kirimWA(sender, formatHasil(cariBarang(keyword), tokoBaru.kode, sender));
            await tunggu(800);
            await kirimWA(sender, getMenuCariUlang(tokoBaru.nama));
          } else await kirimWA(sender, getMenuSiapCari(tokoBaru.nama));
          return;
        }
        if (isKonfirmasiTidak(low)) { updateSesi(sender, { pendingPindahToko: null }); await kirimWA(sender, '👍 Tetap di *' + NAMA_TOKO[s.tokoKode] + '*'); return; }
        await kirimWA(sender, '🤔 Jawab *iya* / *tidak* / *batal*'); return;
      }
      if (!s.tokoKode) {
        const toko = parsePilihanToko(low);
        if (toko) {
          updateSesi(sender, { tokoKode: toko.kode });
          if (s.pendingKw) { const kw = s.pendingKw; updateSesi(sender, { pendingKw: null }); await kirimWA(sender, formatHasil(cariBarang(kw), toko.kode, sender)); await tunggu(800); await kirimWA(sender, getMenuCariUlang(toko.nama)); }
          else await kirimWA(sender, getMenuSiapCari(toko.nama));
          return;
        }
        await kirimWA(sender, getMenuPilihToko('cari')); return;
      }
      if (low === '9' || low === 'ganti toko' || low === 'ganti') { updateSesi(sender, { tokoKode: null }); await kirimWA(sender, getMenuPilihToko('cari')); return; }
      if (!msg) return;
      
      const tokoLain = deteksiTokoDariTeks(low);
      const tokoBeda = tokoLain.filter(function(t) { return t.kode !== s.tokoKode; });
      if (tokoBeda.length > 0) {
        const tokoTarget = tokoBeda[0];
        const kwBersih = bersihkanKeywordDariToko(msg);
        updateSesi(sender, { pendingPindahToko: { toko: tokoTarget, keyword: kwBersih } });
        await kirimWA(sender, '🤔 *Hmm...*\n' + GARIS_TEBAL + '\n📍 Kamu di: *' + NAMA_TOKO[s.tokoKode] + '*\n🎯 Tapi sebut: *' + tokoTarget.nama + '*\n' + (kwBersih ? '🔍 Keyword: _' + kwBersih + '_\n' : '') + GARIS_TEBAL + '\n❓ *Pindah ke ' + tokoTarget.nama + '?*\n\n✅ *iya* | ❌ *tidak* | 🔙 *batal*');
        return;
      }
      
      await kirimWA(sender, formatHasil(cariBarang(msg), s.tokoKode, sender));
      await tunggu(800);
      await kirimWA(sender, getMenuCariUlang(NAMA_TOKO[s.tokoKode]));
      return;
    }

    // ── AI CHAT BARANG ──
    if (!_lagiInput && isMember(sender) && isPertanyaanBarang(low) && msg.length >= 5) {
      const tokoDisebutAI = deteksiTokoDariTeks(low);
      await kirimWA(sender, '🤖 _Sedang berpikir..._');
      let aiJawaban = null;
      try { aiJawaban = await aiChatBarang(msg, sender, null); } catch (e) { log.error('AI_CHAT', e.message); }
      if (aiJawaban && aiJawaban.length > 10) { await kirimWA(sender, '🤖 ' + aiJawaban); return; }
      // Fallback
      const hasil = cariBarang(msg);
      if (hasil.hasil && hasil.hasil.length > 0) {
        let resp = '🤖 *Hasil Pencarian:*\n' + GARIS_TEBAL + '\n\n';
        const tokoShow = tokoDisebutAI.length > 0 ? tokoDisebutAI : TOKO_LIST;
        hasil.hasil.slice(0, 5).forEach(function(d, i) {
          resp += '*' + (i+1) + '. ' + d.nama + '* (' + d.kode + ')\n';
          tokoShow.forEach(function(t) {
            const h = d.harga[t.kode];
            resp += '🏪 ' + t.nama + ': ' + fRp(h.ecer) + ' | Stok: ' + (h.stok > 0 ? h.stok : '⚠️ Kosong') + '\n';
          });
          if (i < 4) resp += '\n';
        });
        await kirimWA(sender, resp);
        return;
      }
      await kirimWA(sender, '🤔 Tidak ditemukan. Ketik *menu* untuk cari manual.'); return;
    }

    // ── BANDING HARGA ──
    if (!_lagiInput && isMember(sender) && isPertanyaanBanding(low)) {
      await kirimWA(sender, '📊 _Menganalisa..._');
      try {
        const tokoDisebut = deteksiTokoDariTeks(low);
        let keyword = msg.replace(/banding(kan)?|termurah|termahal|paling murah|paling mahal|lebih murah|lebih mahal|dimana murah|dimana mahal|selisih|beda harga|analisa harga|compare|di toko|harga|harganya/gi, '').replace(/\bdan\b|\bdengan\b|\bvs\b|\batau\b|\bdi\b|\bke\b/gi, ' ');
        TOKO_LIST.forEach(function(t) {
          keyword = keyword.replace(new RegExp('\\b' + t.kode + '\\b', 'gi'), '');
          t.alias.forEach(function(a) { keyword = keyword.replace(new RegExp('\\b' + a + '\\b', 'gi'), ''); });
        });
        keyword = keyword.trim().replace(/\s+/g, ' ');
        if (!keyword || keyword.length < 2) { await kirimWA(sender, '⚠️ Contoh: _bandingkan harga NN00001_'); return; }
        const hasilCari = cariBarangPrioritas(keyword);
        if (!hasilCari.hasil || hasilCari.hasil.length === 0) { await kirimWA(sender, '❌ Tidak ditemukan: _' + keyword + '_'); return; }
        if (hasilCari.hasil.length === 1) { await kirimWA(sender, formatAnalisaBandingHargaToko(hasilCari.hasil[0], tokoDisebut, sender)); return; }
        // Multiple
        let resp = '📊 *PERBANDINGAN HARGA*\n' + GARIS_TEBAL + '\n🔍 _' + keyword + '_\n📦 ' + hasilCari.hasil.length + ' barang\n' + GARIS_TEBAL + '\n\n';
        hasilCari.hasil.slice(0, 5).forEach(function(item, i) {
          const a = analisaPerbandinganHargaToko(item, tokoDisebut);
          resp += '*' + (i+1) + '. ' + item.nama + '* (' + item.kode + ')\n';
          if (a.analisa.ecer) {
            if (a.analisa.ecer.semuaSama) resp += '   💰 ' + fRp(a.analisa.ecer.termurah.harga) + ' _(sama)_\n';
            else resp += '   🟢 ' + fRp(a.analisa.ecer.termurah.harga) + ' (' + a.analisa.ecer.termurah.toko.nama + ')\n   🔴 ' + fRp(a.analisa.ecer.termahal.harga) + ' (' + a.analisa.ecer.termahal.toko.nama + ')\n   💸 Selisih: ' + fRp(a.analisa.ecer.selisih) + '\n';
          }
          resp += '\n';
        });
        await kirimWA(sender, resp);
        return;
      } catch (e) { log.error('BANDING', e.message); await kirimWA(sender, '⚠️ Error. Ketik *menu*.'); return; }
    }

    // ── AI CHAT UMUM ──
    if (!_lagiInput && isMember(sender) && isPertanyaanUmum(low) && !isPertanyaanBarang(low) && !isPertanyaanBanding(low) && msg.length >= 5) {
      await kirimWA(sender, '🤖 _Menjawab..._');
      try { const j = await aiChatUmum(msg, sender); if (j && j.length > 10) { await kirimWA(sender, '🤖 ' + j); return; } } catch (e) {}
      await kirimWA(sender, '🤖 Hai! Saya Bot Perabot 5 toko.\n\n💡 Ketik *menu* untuk fitur lengkap!');
      return;
    }

    // ── ADMIN MODE INPUT ──
    if (isAdmin(sender) && s.adminAksi) {
      const aksi = s.adminAksi; updateSesi(sender, { adminAksi: null });
      if (aksi === 'daftar') return await handleAdmin(sender, 'daftar ' + msg, 'daftar ' + low);
      if (aksi === 'hapus') return await handleAdmin(sender, 'hapus ' + msg, 'hapus ' + low);
      if (aksi === 'namakontak') return await handleAdmin(sender, 'namakontak ' + msg, 'namakontak ' + low);
      if (aksi === 'hapuskontak') return await handleAdmin(sender, 'hapuskontak ' + msg, 'hapuskontak ' + low);
    }

    // ── MENU ADMIN ──
    if (s.mode === 'admin_menu') {
      const aksi = parsePilihanAdmin(low);
      if (aksi === 'listmember' || aksi === 'listkontak' || aksi === 'reload' || aksi === 'info') { resetSesi(sender); return await handleAdmin(sender, aksi, aksi); }
      if (aksi === 'daftar') { updateSesi(sender, { adminAksi:'daftar', mode:null }); await kirimWA(sender, '➕ Ketik nomor HP:\n_6281234567890_\n\n🔙 *batal*'); return; }
      if (aksi === 'hapus') { updateSesi(sender, { adminAksi:'hapus', mode:null }); await kirimWA(sender, '➖ Ketik nomor HP:\n\n🔙 *batal*'); return; }
      if (aksi === 'namakontak') { updateSesi(sender, { adminAksi:'namakontak', mode:null }); await kirimWA(sender, '✏️ Format: nomor nama\n_628xxx Pak Budi_\n\n🔙 *batal*'); return; }
      if (aksi === 'hapuskontak') { updateSesi(sender, { adminAksi:'hapuskontak', mode:null }); await kirimWA(sender, '🗑️ Ketik nomor HP\n\n🔙 *batal*'); return; }
      await kirimWA(sender, getMenuAdmin()); return;
    }

    // ── PILIH MENU ──
               if(!s.menu&&!s.mode){
      const pilihan=parsePilihanMenu(low);
      if(pilihan===1||pilihan===2){if(!bisaAksesLaporan(sender)){await kirimWA(sender,'🚫 Menu Laporan hanya untuk staff ditunjuk.\nGunakan menu *4* atau tanya AI.');return;}resetSesi(sender);updateSesi(sender,{menu:pilihan});await kirimWA(sender,getMenuPilihToko(pilihan));return;}
      if(pilihan===3){if(!bisaAksesLaporan(sender)){await kirimWA(sender,'🚫 Menu Laporan hanya untuk staff ditunjuk.\nGunakan menu *4* atau tanya AI.');return;}resetSesi(sender);updateSesi(sender,{menu:3});await kirimWA(sender,getMenuPilihHari('Marketplace Perabot Mama'));return;}
      if(pilihan===4){if(!isMember(sender)){await kirimWA(sender,'🚫 Hanya untuk member.');return;}resetSesi(sender);updateSesi(sender,{mode:'cari',tokoKode:null});await kirimWA(sender,getMenuPilihToko('cari'));return;}
      if(pilihan===9){if(!isAdmin(sender)){await kirimWA(sender,'🚫 Khusus admin.');return;}resetSesi(sender);updateSesi(sender,{mode:'admin_menu'});await kirimWA(sender,getMenuAdmin());return;}
      await kirimWA(sender,'🤔 Maaf, tidak mengerti.\n\nKetik *menu*.');return;
    }

    // ── PILIH TOKO ──
    if (s.menu !== 3 && !s.toko) {
      const toko = parsePilihanToko(low);
      if (toko) { updateSesi(sender, { toko: toko.kode }); await kirimWA(sender, getMenuPilihHari(toko.nama)); return; }
      await kirimWA(sender, getMenuPilihToko(s.menu)); return;
    }

    // ── PILIH HARI ──
    if (s.kemarin === undefined || s.kemarin === null) {
      const kem = parsePilihanHari(low);
      if (kem === null) { const nm = s.menu === 3 ? 'Marketplace Perabot Mama' : NAMA_TOKO[s.toko]; await kirimWA(sender, getMenuPilihHari(nm)); return; }
      updateSesi(sender, { kemarin: kem });
      const nm = s.menu === 3 ? 'Marketplace Perabot Mama' : NAMA_TOKO[s.toko];
      
      // ★★★ MENU 1 → LANGSUNG MASUK SCAN FOTO MODE ★★★
      if (s.menu === 1) {
        const t = getTanggal(kem);
        const kk = kem ? ' _(kemarin)_' : '';
        let intro = '╭━━━━━━━━━━━━━━━━━╮\n│  ✅ *MULAI INPUT*  │\n╰━━━━━━━━━━━━━━━━━╯\n\n';
        intro += '🏦 ' + nm + '\n📅 ' + t + kk + '\n\n';
        intro += '📸 *Kirim foto satu-per-satu*\n   Bot akan scan otomatis\n\n';
        intro += '⌨️ Atau ketik angka manual\n\n';
        intro += '━━━━━━━━━━━━━━━━━━';
        await kirimWA(sender, intro);
        await tunggu(800);
        
        // Mulai scan mode
        updateSesi(sender, { scanActive: true, scanData: {}, scanStepIdx: 0 });
        await kirimWA(sender, msgMintaFoto(s.toko, 0, nm, {}));
        return;
      }
      
      if (s.menu === 2) { await kirimWA(sender, getMenuSiapInputHarga(nm, kem)); return; }
      if (s.menu === 3) { await kirimWA(sender, getMenuSiapInputMarket(nm, kem)); return; }
      return;
    }

    // ── INPUT DATA (menu 2 & 3 — foto/teks biasa) ──
    const namaToko = s.menu === 3 ? 'Marketplace Perabot Mama' : NAMA_TOKO[s.toko];
    let laporan = '';

    if (image && image.length > 0) {
      await kirimWA(sender, '📸 _Dianalisa AI..._');
      try {
        const prompt = buatPromptAI(s.menu, namaToko, getTanggal(s.kemarin), s.toko);
        laporan = await analisaGambar(image, prompt);
      } catch (e) { log.error('GEMINI', e.message); await kirimWA(sender, '❌ Gagal baca foto.'); return; }
    } else if (msg) {
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
    try { const sender = req.body ? (req.body.sender || req.body.from || '') : ''; if (sender) await kirimWA(sender, '⚠️ Error. Ketik *menu*.'); } catch (e) {}
  }
});

// ════════════════════════════════════════════════════════════════
//   22. START
// ════════════════════════════════════════════════════════════════

app.listen(CONFIG.port, function() {
  console.log('\n=====================================');
  console.log('  ' + CONFIG.appName + ' v3.13');
  console.log('  (Scan Foto Multi-Step + All Features)');
  console.log('=====================================');
  console.log('  Port      : ' + CONFIG.port);
  console.log('  Admin     : ' + CONFIG.adminNumber);
  console.log('  Items     : ' + DATA_BARANG.length);
  console.log('  Members   : ' + MEMBERS.length + '/' + CONFIG.maxMember);
  console.log('  AI Stack  :');
  if (CONFIG.groqKey)       console.log('    ✅ GROQ');
  if (CONFIG.geminiKey)     console.log('    ✅ GEMINI (' + geminiKeys().length + ' keys)');
  if (CONFIG.openrouterKey) console.log('    ✅ OPENROUTER');
  console.log('=====================================\n');
});

process.on('uncaughtException',  function(e) { log.error('SYSTEM', 'Uncaught',  e.message); });
process.on('unhandledRejection', function(r) { log.error('SYSTEM', 'Unhandled', String(r));  });
