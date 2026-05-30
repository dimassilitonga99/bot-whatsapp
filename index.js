'use strict';

// ════════════════════════════════════════════════════════════════
//   BOT WHATSAPP - LAPORAN & CARI HARGA BARANG
//   Versi 3.2 - FREE Friendly Menu + Format Khusus CP
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

function geminiUrl() {
  return 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + CONFIG.geminiKey;
}

if (!CONFIG.fonnteToken || !CONFIG.geminiKey) {
  console.error('\n❌ ERROR: FONNTE_TOKEN atau GEMINI_KEY belum diisi\n');
  process.exit(1);
}

[CONFIG.paths.storage, path.dirname(CONFIG.paths.logs)].forEach(function(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// ════════════════════════════════════════════════════════════════
//   2. EMOJI ANGKA
// ════════════════════════════════════════════════════════════════

const EMOJI_NUM = ['0️⃣','1️⃣','2️⃣','3️⃣','4️⃣','5️⃣','6️⃣','7️⃣','8️⃣','9️⃣','🔟'];
function emojiNum(n) {
  if (n >= 0 && n <= 10) return EMOJI_NUM[n];
  return String(n);
}

// ════════════════════════════════════════════════════════════════
//   3. DATA TOKO
// ════════════════════════════════════════════════════════════════

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

const ADMIN_COMMANDS = [
  'daftar', 'hapus', 'listmember', 'namakontak', 'hapuskontak',
  'listkontak', 'reload', 'info'
];

// ★★★ KONFIGURASI INPUT BERTAHAP CENTRAL PERABOT ★★★
const CP_FIELDS = [
  { key: 'k1',             label: 'Kassa 1 (Yuni-Salsa)',                emoji: '💵' },
  { key: 'k2',             label: 'Kassa 2 (Nanda-Umi-Marselina)',       emoji: '💵' },
  { key: 'k3',             label: 'Kassa 3 (Febri-Jien-Tika)',           emoji: '💵' },
  { key: 'k4',             label: 'Kassa 4 (Delfi-Tirsa)',               emoji: '💵' },
  { key: 'tunai',          label: 'Total Tunai',                          emoji: '💰' },
  { key: 'debit',          label: 'Total Debit',                          emoji: '💳' },
  { key: 'kredit',         label: 'Total Credit',                         emoji: '💳' },
  { key: 'ecer',           label: 'Total Penjualan Ecer',                 emoji: '🛒' },
  { key: 'grosir',         label: 'Total Penjualan Grosir',               emoji: '📦' },
  { key: 'promo',          label: 'Total Kasir Promo',                    emoji: '🎁' },
  { key: 'promoTunai',     label: 'Promo - Tunai',                        emoji: '🎁' },
  { key: 'promoDebit',     label: 'Promo - Debit',                        emoji: '🎁' },
  { key: 'promoKredit',    label: 'Promo - Credit',                       emoji: '🎁' },
  { key: 'parkirKomputer', label: 'Parkir di Komputer',                   emoji: '🅿️' },
  { key: 'parkirLuar',     label: 'Parkir Stor Luar',                     emoji: '🅿️' },
];

// ════════════════════════════════════════════════════════════════
//   4. LOGGER
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
    try {
      fs.appendFileSync(CONFIG.paths.logs, fullMsg + (data ? ' | ' + JSON.stringify(data) : '') + '\n');
    } catch (e) {}
  },
};

// ════════════════════════════════════════════════════════════════
//   5. UTILS
// ════════════════════════════════════════════════════════════════

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
  prefix   = prefix   || 'Rp';
  fallback = fallback || prefix + ' -';
  const v  = parseFloat(n) || 0;
  if (v === 0) return fallback;
  return prefix + ' ' + v.toLocaleString('id-ID');
}

const fRp  = function(n) { return formatRp(n, 'Rp',  'Rp -');  };
const fRpP = function(n) { return formatRp(n, 'Rp',  'Rp 0');  };
const fRpD = function(n) {
  const v = parseFloat(n) || 0;
  if (v === 0) return 'Rp. -';
  return 'Rp. ' + v.toLocaleString('id-ID');
};

const GARIS_TEBAL = '━━━━━━━━━━━━━━━━━━';
const GARIS_TIPIS = '──────────────────';

// ════════════════════════════════════════════════════════════════
//   6. STORAGE
// ════════════════════════════════════════════════════════════════

function loadJSON(filePath, defaultValue) {
  try {
    if (fs.existsSync(filePath)) {
      return JSON.parse(fs.readFileSync(filePath, 'utf8'));
    }
  } catch (e) {
    log.error('STORAGE', 'Gagal load ' + filePath, e.message);
  }
  return defaultValue;
}

function saveJSON(filePath, data) {
  try {
    fs.writeFileSync(filePath, JSON.stringify(data, null, 2));
    return true;
  } catch (e) {
    log.error('STORAGE', 'Gagal save ' + filePath, e.message);
    return false;
  }
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
  MEMBERS.push(nomor);
  saveJSON(CONFIG.paths.members, MEMBERS);
  return { ok: true };
}

function hapusMember(nomor) {
  const idx = MEMBERS.indexOf(nomor);
  if (idx === -1) return { ok: false, alasan: 'Nomor tidak ditemukan' };
  MEMBERS.splice(idx, 1);
  saveJSON(CONFIG.paths.members, MEMBERS);
  return { ok: true };
}

function setNama(nomor, nama) {
  if (!nomor || !nama) return { ok: false, alasan: 'Nomor & nama wajib' };
  KONTAK[nomor] = nama.trim();
  saveJSON(CONFIG.paths.kontak, KONTAK);
  return { ok: true };
}

function hapusKontak(nomor) {
  if (!KONTAK[nomor]) return { ok: false, alasan: 'Kontak tidak ditemukan' };
  delete KONTAK[nomor];
  saveJSON(CONFIG.paths.kontak, KONTAK);
  return { ok: true };
}

const TIMEOUT_MS = CONFIG.sesiTimeoutMenit * 60 * 1000;

function getSesi(nomor) {
  if (!SESI[nomor]) SESI[nomor] = {};
  SESI[nomor]._lastActive = Date.now();
  return SESI[nomor];
}

function resetSesi(nomor) {
  SESI[nomor] = { _lastActive: Date.now() };
  saveJSON(CONFIG.paths.sesi, SESI);
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
    if (now - (SESI[n]._lastActive || 0) > TIMEOUT_MS) {
      delete SESI[n];
      buang++;
    }
  });
  if (buang > 0) saveJSON(CONFIG.paths.sesi, SESI);
}, 5 * 60 * 1000);

// ════════════════════════════════════════════════════════════════
//   7. EXCEL
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

// ════════════════════════════════════════════════════════════════
//   8. SMART SEARCH
// ════════════════════════════════════════════════════════════════

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

function bersihkanTeks(str) {
  return String(str).toUpperCase().replace(/[^A-Z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
}

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
//   9. KIRIM WHATSAPP (FONNTE)
// ════════════════════════════════════════════════════════════════

function tunggu(ms) { return new Promise(function(r) { setTimeout(r, ms); }); }

async function kirimWA(target, message, retry) {
  retry = retry || 0;
  try {
    await axios.post(
      CONFIG.fonnteUrl,
      { target: target, message: message },
      { headers: { Authorization: CONFIG.fonnteToken }, timeout: 10000 }
    );
    log.info('FONNTE', 'OK ke ' + target);
    return true;
  } catch (err) {
    log.warn('FONNTE', 'Gagal attempt ' + (retry + 1));
    if (retry < CONFIG.maxRetry - 1) {
      await tunggu(CONFIG.retryDelay);
      return kirimWA(target, message, retry + 1);
    }
    log.error('FONNTE', 'GAGAL TOTAL', err.message);
    return false;
  }
}

// ════════════════════════════════════════════════════════════════
//   10. GEMINI AI
// ════════════════════════════════════════════════════════════════

async function analisaGambar(imageUrl, prompt) {
  const imgResp   = await axios.get(imageUrl, { responseType: 'arraybuffer', timeout: 15000 });
  const mimeType  = imgResp.headers['content-type'] || 'image/jpeg';
  const imageData = Buffer.from(imgResp.data).toString('base64');
  const resp = await axios.post(geminiUrl(), {
    contents: [{
      parts: [
        { inline_data: { mime_type: mimeType, data: imageData } },
        { text: prompt },
      ],
    }],
  }, { timeout: 30000 });
  return resp.data.candidates[0].content.parts[0].text || '';
}

function buatPromptAI(menuType, namaToko, tanggal, tokoKode) {
  const fmt = ' Format WhatsApp dengan emoji. Format rupiah Rp X.XXX.XXX.';
  if (menuType === 1) {
    if (tokoKode === 'cp') {
      return 'Baca data penjualan toko Central Perabot tanggal ' + tanggal + '. ' +
        'Buat laporan dengan format: 4 kassa (Kassa 1-4 dengan nama kasir), total keseluruhan, ' +
        'metode pembayaran (Tunai/Debit/Credit), jenis penjualan (Ecer/Grosir), ' +
        'laporan kasir promo (total + metode bayar), laporan parkir (komputer + stor luar + total).' + fmt;
    }
    return 'Baca data penjualan toko "' + namaToko + '" tanggal ' + tanggal + '. Buat laporan: kassa, total, metode bayar, jenis penjualan.' + fmt;
  }
  if (menuType === 2) return 'Baca data harga barang toko "' + namaToko + '" tanggal ' + tanggal + '. Buat laporan: barang baru, naik harga, turun harga.' + fmt;
  if (menuType === 3) return 'Baca data marketplace tanggal ' + tanggal + '. Buat laporan: per toko, per channel, metode bayar.' + fmt;
  return 'Buat laporan rapi.' + fmt;
}

// ════════════════════════════════════════════════════════════════
//   11. SAPAAN
// ════════════════════════════════════════════════════════════════

const KATA_SAPAAN = [
  'halo','hai','hi','hello','hey','pagi','siang','sore','malam','selamat',
  'assalamualaikum','salam','permisi',
];

const KATA_TERIMAKASIH = [
  'terima kasih','terimakasih','makasih','thanks','thank you',
  'thx','tq','ty','tengkyu','mksh','trims',
];

function cocokKata(low, kata) {
  return low === kata || low.startsWith(kata + ' ') || low.startsWith(kata + ',') ||
    low.startsWith(kata + '!') || low.startsWith(kata + '.') || low.endsWith(' ' + kata);
}

function isSapaan(low)      { return KATA_SAPAAN.some(function(k) { return cocokKata(low, k); }); }
function isTerimakasih(low) { return KATA_TERIMAKASIH.some(function(k) { return cocokKata(low, k); }); }

function sapaanPertama(sender) {
  const nama = getNama(sender);
  return (nama ? 'Selamat ' + getWaktu() + ', *' + nama + '*! 😊' : 'Selamat ' + getWaktu() + '! 😊');
}

function balasTerimakasih(sender) {
  const nama = getNama(sender);
  const n = nama ? ', *' + nama + '*' : '';
  return ['Sama-sama' + n + '! 😊', 'Dengan senang hati' + n + '! 😊', 'Tentu' + n + '! 😊'][Math.floor(Math.random() * 3)];
}

function isAdminCommand(low) {
  if (['listmember','listkontak','reload','info'].indexOf(low) >= 0) return true;
  return ADMIN_COMMANDS.some(function(cmd) { return low.startsWith(cmd + ' '); });
}

// ════════════════════════════════════════════════════════════════
//   12. MENU FRIENDLY (Tampilan Mirip Tombol)
// ════════════════════════════════════════════════════════════════

function getMenuUtama(nomor) {
  const nama = getNama(nomor);
  const salam = nama ? '*' + nama + '*' : 'Kamu';
  let m = '╭━━━━━━━━━━━━━━━━━╮\n';
  m += '│  🤖 *BOT TOKO PERABOT*  │\n';
  m += '╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += 'Halo ' + salam + '! 👋\n';
  m += 'Silakan pilih menu:\n\n';
  m += '┌─────────────────────\n';
  m += '│ ' + emojiNum(1) + ' 📊 Laporan Penjualan\n';
  m += '│ ' + emojiNum(2) + ' 🏷️ Laporan Harga Barang\n';
  m += '│ ' + emojiNum(3) + ' 🛒 Laporan Marketplace\n';
  if (isMember(nomor)) m += '│ ' + emojiNum(4) + ' 🔍 Cari Harga Barang\n';
  if (isAdmin(nomor))  m += '│ ' + emojiNum(9) + ' 👑 Menu Admin\n';
  m += '└─────────────────────\n\n';
  m += '💬 *Cara pilih:*\n';
  m += '   Ketik nomor (contoh: *1*)\n';
  m += '   atau ketik nama menunya';
  return m;
}

function getMenuPilihToko(menuType) {
  const ic = menuType === 1 ? '📊' : menuType === 2 ? '🏷️' : menuType === 'cari' ? '🔍' : '🛒';
  const jd = menuType === 1 ? 'LAPORAN PENJUALAN'
           : menuType === 2 ? 'LAPORAN HARGA BARANG'
           : menuType === 'cari' ? 'CARI HARGA BARANG'
           : 'LAPORAN MARKETPLACE';
  let m = '╭━━━━━━━━━━━━━━━━━╮\n';
  m += '│  ' + ic + ' *' + jd + '*  │\n';
  m += '╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🏦 *Pilih Toko:*\n\n';
  m += '┌─────────────────────\n';
  TOKO_LIST.forEach(function(t, i) { m += '│ ' + emojiNum(i + 1) + ' ' + t.nama + '\n'; });
  m += '└─────────────────────\n\n';
  m += '💬 Ketik nomor (1-5) atau nama toko\n';
  m += '   contoh: *1* atau *nk*\n\n';
  m += '🔙 Ketik *0* untuk kembali';
  return m;
}

function getMenuPilihHari(namaToko) {
  let m = '╭━━━━━━━━━━━━━━━━━╮\n';
  m += '│  🏦 *' + namaToko + '*  │\n';
  m += '╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '📅 *Laporan untuk hari:*\n\n';
  m += '┌─────────────────────\n';
  m += '│ ' + emojiNum(1) + ' 📅 *HARI INI*\n';
  m += '│    _(' + getTanggal(false) + ')_\n│\n';
  m += '│ ' + emojiNum(2) + ' 📅 *KEMARIN*\n';
  m += '│    _(' + getTanggal(true) + ')_\n';
  m += '└─────────────────────\n\n';
  m += '💬 Ketik *1* untuk hari ini\n';
  m += '   atau *2* untuk kemarin\n\n';
  m += '🔙 Ketik *0* untuk kembali';
  return m;
}

function getMenuSiapInput(namaToko, kemarin, menuType, tokoKode) {
  const t = getTanggal(kemarin);
  const k = kemarin ? ' _(kemarin)_' : '';
  let contoh = '';
  
  if (menuType === 1) {
    // Khusus Central Perabot - format 4 kassa + promo + parkir
    if (tokoKode === 'cp') {
      contoh = 'k1 Yuni-Salsa 42345000\n' +
               'k2 Nanda-Umi-Marselina 45213500\n' +
               'k3 Febri-Jien-Tika 1979500\n' +
               'k4 Delfi-Tirsa 51025500\n' +
               'tunai 66619000\n' +
               'debit 73944500\n' +
               'kredit 0\n' +
               'ecer 105346500\n' +
               'grosir 35217000\n' +
               'promo 1675000\n' +
               'promotunai 1675000\n' +
               'promodebit 0\n' +
               'promokredit 0\n' +
               'parkirkomputer 0\n' +
               'parkirluar 369500';
    } else {
      contoh = 'k1 29000000\nk2 11000000\ntunai 26000000\ndebit 14000000\necer 23000000\ngrosir 17000000';
    }
  } else if (menuType === 2) {
    contoh = '---baru---\nNama barang baru\n---naik---\nNama barang naik\n---turun---\nNama barang turun';
  } else {
    contoh = 'oesapa 0\ntdm 0\ncentral 21061000\nwa 21061000\nshopee 0\ntiktok 0\ntokopedia 0\ntunai 304000\ndebit 20757000';
  }
  
  let m = '╭━━━━━━━━━━━━━━━━━╮\n';
  m += '│  ✅ *SIAP INPUT*  │\n';
  m += '╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🏦 ' + namaToko + '\n';
  m += '📅 ' + t + k + '\n\n';
  m += '━━━━━━━━━━━━━━━━━━\n';
  m += '📤 *Pilih cara input:*\n';
  m += '━━━━━━━━━━━━━━━━━━\n\n';
  m += '📸 *OPSI 1: Kirim FOTO*\n';
  m += '   Foto akan dibaca otomatis\n   oleh AI 🤖\n\n';
  m += '⌨️ *OPSI 2: Ketik manual*\n';
  m += '   Contoh format:\n\n';
  m += '```' + contoh + '```\n\n';
  m += '🔙 Ketik *0* untuk batal';
  return m;
}

function getMenuAdmin() {
  let m = '╭━━━━━━━━━━━━━━━━━╮\n';
  m += '│  👑 *MENU ADMIN*  │\n';
  m += '╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🛠️ *Pilih aksi:*\n\n';
  m += '┌─────────────────────\n';
  m += '│ ' + emojiNum(1) + ' 👥 List Member\n';
  m += '│ ' + emojiNum(2) + ' 📒 List Kontak\n';
  m += '│ ' + emojiNum(3) + ' ➕ Tambah Member\n';
  m += '│ ' + emojiNum(4) + ' ➖ Hapus Member\n';
  m += '│ ' + emojiNum(5) + ' ✏️ Set Nama Kontak\n';
  m += '│ ' + emojiNum(6) + ' 🗑️ Hapus Kontak\n';
  m += '│ ' + emojiNum(7) + ' 🔄 Reload Excel\n';
  m += '│ ' + emojiNum(8) + ' ℹ️ Info Sistem\n';
  m += '└─────────────────────\n\n';
  m += '💬 Ketik nomor (1-8)\n\n';
  m += '🔙 Ketik *0* untuk kembali';
  return m;
}

function getMenuSiapCari(namaToko) {
  let m = '╭━━━━━━━━━━━━━━━━━╮\n';
  m += '│  🔍 *CARI BARANG*  │\n';
  m += '╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🏦 ' + namaToko + '\n\n';
  m += '━━━━━━━━━━━━━━━━━━\n';
  m += '⌨️ *Ketik nama atau kode:*\n';
  m += '━━━━━━━━━━━━━━━━━━\n\n';
  m += '💡 *Contoh:*\n';
  m += '   • _dandang eagle 20_\n';
  m += '   • _NN00001_\n';
  m += '   • _golden sunkist_\n\n';
  m += '✨ Bot otomatis koreksi typo!\n\n';
  m += '🔙 Ketik *0* untuk kembali';
  return m;
}

function getMenuCariUlang(namaToko) {
  let m = '━━━━━━━━━━━━━━━━━━\n';
  m += '🔍 *Cari lagi di ' + namaToko + '?*\n';
  m += '━━━━━━━━━━━━━━━━━━\n\n';
  m += '⌨️ Ketik nama/kode barang lain\n\n';
  m += '┌─────────────────────\n';
  m += '│ ' + emojiNum(9) + ' 🔄 Ganti toko\n';
  m += '│ ' + emojiNum(0) + ' 🔙 Menu utama\n';
  m += '└─────────────────────';
  return m;
}

// ════════════════════════════════════════════════════════════════
//   13. PARSER PINTAR
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
  const map = {
    '1': 'listmember', '2': 'listkontak',
    '3': 'daftar', '4': 'hapus',
    '5': 'namakontak', '6': 'hapuskontak',
    '7': 'reload', '8': 'info',
  };
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
//   14. FORMAT HASIL CARI
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
//   15. GENERATOR LAPORAN
// ════════════════════════════════════════════════════════════════

function genLapPenjualan(text, namaToko, kemarin, tokoKode) {
  // Khusus Central Perabot pakai format khusus
  if (tokoKode === 'cp') return genLapPenjualanCP(text, kemarin);
  
  // Format umum untuk toko lain
  const t = getTanggal(kemarin);
  const k = kemarin ? ' _(kemarin)_' : '';
  const d = {};
  text.trim().toLowerCase().split('\n').forEach(function(l) {
    const p = l.trim().split(/\s+/);
    if (p.length >= 2) {
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

// ★★★ FUNGSI MENU INPUT BERTAHAP CP ★★★

function getMenuTanyaField(fieldIndex, dataSudahIsi) {
  const field = CP_FIELDS[fieldIndex];
  const total = CP_FIELDS.length;
  const sisa = total - fieldIndex - 1;
  
  let m = '╭━━━━━━━━━━━━━━━━━╮\n';
  m += '│  📝 *INPUT LAPORAN CP*  │\n';
  m += '╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🏦 Central Perabot\n';
  m += '📊 Progress: ' + (fieldIndex + 1) + '/' + total + '\n';
  m += '━━━━━━━━━━━━━━━━━━\n\n';
  m += field.emoji + ' *Masukkan ' + field.label + ':*\n\n';
  m += '💬 *Ketik angka saja*\n';
  m += '   Contoh: _42345000_\n\n';
  m += 'ℹ️ Kalau kosong/tidak ada,\n';
  m += '   ketik *0* atau *skip*\n\n';
  
  // Tampilkan ringkasan data yang sudah diisi
  if (fieldIndex > 0) {
    m += '━━━━━━━━━━━━━━━━━━\n';
    m += '✅ *Sudah diisi:*\n';
    for (let i = 0; i < fieldIndex; i++) {
      const f = CP_FIELDS[i];
      const v = dataSudahIsi[f.key] || 0;
      m += '   ' + (i + 1) + '. ' + f.label + ': Rp ' + v.toLocaleString('id-ID') + '\n';
    }
    m += '━━━━━━━━━━━━━━━━━━\n\n';
  }
  
  m += '🔙 Ketik *batal* untuk membatalkan';
  return m;
}

function getMenuKonfirmasiCP(data) {
  function fr(n) {
    const v = parseFloat(n) || 0;
    if (v === 0) return 'Rp. -';
    return 'Rp. ' + v.toLocaleString('id-ID');
  }
  
  let m = '╭━━━━━━━━━━━━━━━━━╮\n';
  m += '│  ✅ *KONFIRMASI DATA*  │\n';
  m += '╰━━━━━━━━━━━━━━━━━╯\n\n';
  m += '🏦 *Central Perabot*\n';
  m += '📅 ' + getTanggal(data._kemarin) + '\n';
  m += '━━━━━━━━━━━━━━━━━━\n\n';
  
  m += '*KASSA:*\n';
  m += '• Kassa 1: ' + fr(data.k1) + '\n';
  m += '• Kassa 2: ' + fr(data.k2) + '\n';
  m += '• Kassa 3: ' + fr(data.k3) + '\n';
  m += '• Kassa 4: ' + fr(data.k4) + '\n\n';
  
  m += '*METODE BAYAR:*\n';
  m += '• Tunai: ' + fr(data.tunai) + '\n';
  m += '• Debit: ' + fr(data.debit) + '\n';
  m += '• Credit: ' + fr(data.kredit) + '\n\n';
  
  m += '*JENIS PENJUALAN:*\n';
  m += '• Ecer: ' + fr(data.ecer) + '\n';
  m += '• Grosir: ' + fr(data.grosir) + '\n\n';
  
  m += '*KASIR PROMO:*\n';
  m += '• Total: ' + fr(data.promo) + '\n';
  m += '• Tunai: ' + fr(data.promoTunai) + '\n';
  m += '• Debit: ' + fr(data.promoDebit) + '\n';
  m += '• Credit: ' + fr(data.promoKredit) + '\n\n';
  
  m += '*PARKIR:*\n';
  m += '• Komputer: ' + fr(data.parkirKomputer) + '\n';
  m += '• Stor Luar: ' + fr(data.parkirLuar) + '\n';
  m += '━━━━━━━━━━━━━━━━━━\n\n';
  
  m += '📤 *Lanjutkan kirim laporan?*\n\n';
  m += '┌─────────────────────\n';
  m += '│ ' + emojiNum(1) + ' ✅ Ya, kirim laporan\n';
  m += '│ ' + emojiNum(2) + ' ✏️ Edit ulang dari awal\n';
  m += '│ ' + emojiNum(0) + ' ❌ Batal\n';
  m += '└─────────────────────';
  return m;
}

// ★★★ Generate laporan CP dari data object (mode bertahap) ★★★
function genLapPenjualanCP_FromData(d, kemarin) {
  const t = getTanggal(kemarin);
  
  function fr(n) {
    const v = parseFloat(n) || 0;
    if (v === 0) return 'Rp. -';
    return 'Rp. ' + v.toLocaleString('id-ID');
  }
  
  const totalUtama = (d.k1 || 0) + (d.k2 || 0) + (d.k3 || 0) + (d.k4 || 0);
  const totalParkir = (d.parkirKomputer || 0) + (d.parkirLuar || 0);
  
  let msg = 'Laporan Penjualan Toko Central Perabot Periode ' + t + '\n\n';
  msg += 'Kassa 1 (Yuni-Salsa) ' + fr(d.k1) + '\n';
  msg += 'Kassa 2 (Nanda-Umi-Marselina) ' + fr(d.k2) + '\n';
  msg += 'Kassa 3 (Febri-Jien-Tika) ' + fr(d.k3) + '\n';
  msg += 'Kassa 4 (Delfi-Tirsa) ' + fr(d.k4) + '\n\n';
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

// ★★★ FORMAT KHUSUS CENTRAL PERABOT ★★★
function genLapPenjualanCP(text, kemarin) {
  const t = getTanggal(kemarin);
  const d = {
    k1: 0, k2: 0, k3: 0, k4: 0,
    nk1: '', nk2: '', nk3: '', nk4: '',
    tunai: 0, debit: 0, kredit: 0,
    ecer: 0, grosir: 0,
    promo: 0,
    promoTunai: 0, promoDebit: 0, promoKredit: 0,
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
    
    // Deteksi pola: k1 [nama-kasir] [angka]
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
    
    const valueStr = parts.slice(valueStart).join('').replace(/[^0-9]/g, '');
    const value = parseFloat(valueStr) || 0;
    
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
  
  const totalUtama = d.k1 + d.k2 + d.k3 + d.k4;
  const totalParkir = d.parkirKomputer + d.parkirLuar;
  
  let msg = 'Laporan Penjualan Toko Central Perabot Periode ' + t + '\n\n';
  msg += 'Kassa 1' + (d.nk1 ? ' (' + d.nk1 + ')' : '') + ' ' + fr(d.k1) + '\n';
  msg += 'Kassa 2' + (d.nk2 ? ' (' + d.nk2 + ')' : '') + ' ' + fr(d.k2) + '\n';
  msg += 'Kassa 3' + (d.nk3 ? ' (' + d.nk3 + ')' : '') + ' ' + fr(d.k3) + '\n';
  msg += 'Kassa 4' + (d.nk4 ? ' (' + d.nk4 + ')' : '') + ' ' + fr(d.k4) + '\n\n';
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
    if (p.length >= 2 && p[0] in d) {
      d[p[0]] = parseFloat(p.slice(1).join('').replace(/[^0-9]/g, '')) || 0;
    }
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
//   16. HANDLER ADMIN
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
      let m = 'Daftar Member (' + MEMBERS.length + '/' + CONFIG.maxMember + ')\n';
      m += '------------------\n';
      if (MEMBERS.length === 0) m += '(belum ada member)\n';
      else for (let i = 0; i < MEMBERS.length; i++) {
        m += (i + 1) + '. ' + MEMBERS[i] + '\n   ' + (KONTAK[MEMBERS[i]] || '(belum ada nama)') + '\n';
      }
      m += '------------------\n';
      m += 'Slot tersisa: ' + (CONFIG.maxMember - MEMBERS.length);
      await kirimWA(sender, m);
    } catch (e) {
      await kirimWA(sender, 'Error: ' + e.message);
    }
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
    let m = 'Daftar Kontak (' + keys.length + ')\n';
    m += '------------------\n';
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
//   17. ROUTES
// ════════════════════════════════════════════════════════════════

app.get('/', function(req, res) {
  res.json({ status: 'ok', app: CONFIG.appName + ' v3.2', items: DATA_BARANG.length, members: MEMBERS.length + '/' + CONFIG.maxMember });
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
//   18. WEBHOOK UTAMA
// ════════════════════════════════════════════════════════════════

const KATA_RESET = ['0', 'batal', 'menu', 'mulai', 'start', 'kembali', 'home'];

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
      await tunggu(800);
      await kirimWA(sender, getMenuUtama(sender));
      return;
    }

    // ── ADMIN COMMAND ──
    if (isAdmin(sender) && isAdminCommand(low)) {
      if (SESI[sender] && (SESI[sender].mode || SESI[sender].menu || SESI[sender].adminAksi)) {
        resetSesi(sender);
      }
      const handled = await handleAdmin(sender, msg, low);
      if (handled) return;
    }

    // ── RESET ──
    if (KATA_RESET.indexOf(low) >= 0) {
      resetSesi(sender);
      await kirimWA(sender, getMenuUtama(sender));
      return;
    }

    // ── SAPAAN ──
    if (isSapaan(low)) {
      const nama = getNama(sender);
      const sl = nama ? 'Selamat ' + getWaktu() + ' juga, *' + nama + '*! 😊' : 'Selamat ' + getWaktu() + ' juga! 😊';
      await kirimWA(sender, sl);
      await tunggu(500);
      await kirimWA(sender, getMenuUtama(sender));
      return;
    }
    if (isTerimakasih(low)) {
      await kirimWA(sender, balasTerimakasih(sender));
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
      if (aksi === 'daftar') {
        updateSesi(sender, { adminAksi: 'daftar', mode: null });
        await kirimWA(sender, '➕ *Tambah Member*\n' + GARIS_TEBAL + '\n\nKetik nomor HP:\n_6281234567890_\n\n🔙 Ketik *0* untuk batal');
        return;
      }
      if (aksi === 'hapus') {
        updateSesi(sender, { adminAksi: 'hapus', mode: null });
        await kirimWA(sender, '➖ *Hapus Member*\n' + GARIS_TEBAL + '\n\nKetik nomor HP:\n_6281234567890_\n\n🔙 Ketik *0* untuk batal');
        return;
      }
      if (aksi === 'namakontak') {
        updateSesi(sender, { adminAksi: 'namakontak', mode: null });
        await kirimWA(sender, '✏️ *Set Nama*\n' + GARIS_TEBAL + '\n\nFormat: nomor nama\n_6281234567890 Pak Budi_\n\n🔙 Ketik *0* untuk batal');
        return;
      }
      if (aksi === 'hapuskontak') {
        updateSesi(sender, { adminAksi: 'hapuskontak', mode: null });
        await kirimWA(sender, '🗑️ *Hapus Kontak*\n' + GARIS_TEBAL + '\n\nKetik nomor HP\n\n🔙 Ketik *0* untuk batal');
        return;
      }
      await kirimWA(sender, getMenuAdmin());
      return;
    }

    // ── MODE CARI ──
    if (s.mode === 'cari') {
      if (!s.tokoKode) {
        if (low === '9') {
          updateSesi(sender, { tokoKode: null });
          await kirimWA(sender, getMenuPilihToko('cari'));
          return;
        }
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
      if (pilihan === 1 || pilihan === 2) {
        updateSesi(sender, { menu: pilihan });
        await kirimWA(sender, getMenuPilihToko(pilihan));
        return;
      }
      if (pilihan === 3) {
        updateSesi(sender, { menu: 3 });
        await kirimWA(sender, getMenuPilihHari('Marketplace Perabot Mama'));
        return;
      }
      if (pilihan === 4) {
        if (!isMember(sender)) {
          await kirimWA(sender, '🚫 *Akses Ditolak*\n\nFitur ini hanya untuk member.\nHubungi admin untuk mendaftar.');
          return;
        }
        resetSesi(sender);
        updateSesi(sender, { mode: 'cari' });
        await kirimWA(sender, getMenuPilihToko('cari'));
        return;
      }
      if (pilihan === 9) {
        if (!isAdmin(sender)) {
          await kirimWA(sender, '🚫 Menu khusus admin.');
          return;
        }
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
      if (toko) {
        updateSesi(sender, { toko: toko.kode });
        await kirimWA(sender, getMenuPilihHari(toko.nama));
        return;
      }
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
      
      // ★★★ KHUSUS CP — mode input bertahap ★★★
      if (s.menu === 1 && s.toko === 'cp') {
        updateSesi(sender, { 
          cpMode: true, 
          cpFieldIndex: 0, 
          cpData: { _kemarin: kem } 
        });
        await kirimWA(sender, getMenuTanyaField(0, {}));
        return;
      }
      
      await kirimWA(sender, getMenuSiapInput(nm, kem, s.menu, s.toko));
      return;
    }

      // ★★★ MODE INPUT BERTAHAP CENTRAL PERABOT ★★★
    if (s.cpMode === true) {
      const idx = s.cpFieldIndex || 0;
      
      // Cek kalau lagi di tahap konfirmasi
      if (s.cpKonfirmasi === true) {
        if (low === '1' || low === 'ya' || low === 'kirim') {
          // Generate laporan
          const data = s.cpData || {};
          const laporanCP = genLapPenjualanCP_FromData(data, data._kemarin);
          await kirimWA(sender, laporanCP);
          resetSesi(sender);
          await tunggu(1500);
          await kirimWA(sender, '✅ *Laporan selesai!* 😊\n\n' + getMenuUtama(sender));
          return;
        }
        if (low === '2' || low === 'edit' || low === 'ulang') {
          // Edit ulang dari awal
          updateSesi(sender, { 
            cpFieldIndex: 0, 
            cpData: { _kemarin: s.cpData._kemarin },
            cpKonfirmasi: false 
          });
          await kirimWA(sender, '🔄 *Mulai input ulang dari awal*\n\n');
          await tunggu(500);
          await kirimWA(sender, getMenuTanyaField(0, {}));
          return;
        }
        if (low === '0' || low === 'batal' || low === 'cancel') {
          resetSesi(sender);
          await kirimWA(sender, '❌ Input dibatalkan.\n\n' + getMenuUtama(sender));
          return;
        }
        // Pilihan tidak dikenali
        await kirimWA(sender, '🤔 Pilih *1* untuk kirim, *2* untuk edit ulang, atau *0* untuk batal');
        return;
      }
      
      // Cek perintah batal
      if (low === 'batal' || low === 'cancel') {
        resetSesi(sender);
        await kirimWA(sender, '❌ Input dibatalkan.\n\n' + getMenuUtama(sender));
        return;
      }
      
      // Parse nilai (terima angka, "0", "skip", atau dengan format)
      let value = 0;
      if (low === 'skip' || low === '-') {
        value = 0;
      } else {
        const cleaned = msg.replace(/[^0-9]/g, '');
        if (cleaned) value = parseFloat(cleaned);
      }
      
      // Simpan data
      const field = CP_FIELDS[idx];
      const newData = Object.assign({}, s.cpData || {}, { [field.key]: value });
      updateSesi(sender, { cpData: newData });
      
      // Cek apakah masih ada field berikutnya
      const nextIdx = idx + 1;
      if (nextIdx < CP_FIELDS.length) {
        updateSesi(sender, { cpFieldIndex: nextIdx });
        await tunggu(300);
        await kirimWA(sender, '✅ Tersimpan: *Rp ' + value.toLocaleString('id-ID') + '*');
        await tunggu(500);
        await kirimWA(sender, getMenuTanyaField(nextIdx, newData));
      } else {
        // Semua field sudah diisi → tampil konfirmasi
        updateSesi(sender, { cpKonfirmasi: true });
        await tunggu(300);
        await kirimWA(sender, '✅ *Semua data sudah lengkap!*');
        await tunggu(500);
        await kirimWA(sender, getMenuKonfirmasiCP(newData));
      }
      return;
    }
    
    // ── LAPORAN: Input data ──
    const namaToko = s.menu === 3 ? 'Marketplace Perabot Mama' : NAMA_TOKO[s.toko];
    let laporan = '';

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
//   19. START SERVER
// ════════════════════════════════════════════════════════════════

app.listen(CONFIG.port, function() {
  console.log('\n=====================================');
  console.log('  ' + CONFIG.appName + ' v3.2');
  console.log('  (FREE Friendly + Format CP)');
  console.log('=====================================');
  console.log('  Port      : ' + CONFIG.port);
  console.log('  Admin     : ' + CONFIG.adminNumber);
  console.log('  Items     : ' + DATA_BARANG.length);
  console.log('  Members   : ' + MEMBERS.length + '/' + CONFIG.maxMember);
  console.log('=====================================\n');
});

process.on('uncaughtException',  function(e) { log.error('SYSTEM', 'Uncaught',  e.message); });
process.on('unhandledRejection', function(r) { log.error('SYSTEM', 'Unhandled', String(r));  });
