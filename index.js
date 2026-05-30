'use strict';

// ════════════════════════════════════════════════════════════════
//   BOT WHATSAPP - LAPORAN & CARI HARGA BARANG
//   Versi 2.0 - Single File
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

  // Path file
  paths: {
    storage:  path.join(__dirname, 'storage'),
    members:  path.join(__dirname, 'storage', 'members.json'),
    kontak:   path.join(__dirname, 'storage', 'kontak.json'),
    sesi:     path.join(__dirname, 'storage', 'sesi.json'),
    disapa:   path.join(__dirname, 'storage', 'disapa.json'),
    excel:    path.join(__dirname, 'harga_barang_5toko.xlsx'),
    logs:     path.join(__dirname, 'logs', 'error.log'),
  },

  // Timeout sesi (menit)
  sesiTimeoutMenit: 30,

  // Retry kirim WA
  maxRetry: 3,
  retryDelay: 2000,
};

// URL Gemini (function agar selalu fresh)
function geminiUrl() {
  return 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + CONFIG.geminiKey;
}

// Validasi env wajib
if (!CONFIG.fonnteToken || !CONFIG.geminiKey) {
  console.error('\n❌ ERROR: FONNTE_TOKEN atau GEMINI_KEY belum diisi di file .env');
  console.error('   Buat file .env di folder yang sama dengan index.js, isi:');
  console.error('   FONNTE_TOKEN=xxx');
  console.error('   GEMINI_KEY=xxx');
  console.error('   ADMIN_NUMBER=628xxx\n');
  process.exit(1);
}

// Buat folder storage & logs
[CONFIG.paths.storage, path.dirname(CONFIG.paths.logs)].forEach(function(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// ════════════════════════════════════════════════════════════════
//   2. DATA TOKO
// ════════════════════════════════════════════════════════════════

const TOKO_LIST = [
  { kode: 'nk',     nama: 'Nasional Kitchen'          },
  { kode: 'tdm',    nama: 'Perabot Mama TDM'          },
  { kode: 'oesapa', nama: 'Perabot Mama Oesapa'       },
  { kode: 'kefa',   nama: 'Perabot Mamaku Kefamenanu' },
  { kode: 'cp',     nama: 'Central Perabot'           },
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

// Default member & kontak (fallback)
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

// ════════════════════════════════════════════════════════════════
//   3. LOGGER
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
//   4. UTILS - WAKTU & FORMAT
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

const GARIS_TEBAL = '━━━━━━━━━━━━━━━━━━';
const GARIS_TIPIS = '──────────────────';

// ════════════════════════════════════════════════════════════════
//   5. STORAGE - MEMBER, KONTAK, SESI
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

log.info('INIT', 'Loaded: ' + MEMBERS.length + ' members, ' +
  Object.keys(KONTAK).length + ' kontak, ' +
  Object.keys(SESI).length + ' sesi');

// ── Access Control ─────────────────────────────────────
function isAdmin(n)  { return n === CONFIG.adminNumber; }
function isMember(n) { return isAdmin(n) || MEMBERS.indexOf(n) >= 0; }
function getNama(n)  { return KONTAK[n] || null; }

// ── Member CRUD ────────────────────────────────────────
function tambahMember(nomor) {
  if (!nomor || !/^[0-9]+$/.test(nomor)) return { ok: false, alasan: 'Format nomor tidak valid' };
  if (isAdmin(nomor))                    return { ok: false, alasan: 'Nomor itu adalah admin' };
  if (MEMBERS.indexOf(nomor) >= 0)       return { ok: false, alasan: 'Nomor sudah terdaftar' };
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

// ── Kontak CRUD ────────────────────────────────────────
function setNama(nomor, nama) {
  if (!nomor || !nama) return { ok: false, alasan: 'Nomor dan nama wajib diisi' };
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

// ── Sesi (auto-save & auto-cleanup) ────────────────────
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

// Auto-cleanup sesi expired tiap 5 menit
setInterval(function() {
  const now = Date.now();
  let buang = 0;
  Object.keys(SESI).forEach(function(n) {
    if (now - (SESI[n]._lastActive || 0) > TIMEOUT_MS) {
      delete SESI[n];
      buang++;
    }
  });
  if (buang > 0) {
    log.info('SESI', 'Cleanup ' + buang + ' sesi expired');
    saveJSON(CONFIG.paths.sesi, SESI);
  }
}, 5 * 60 * 1000);

// ════════════════════════════════════════════════════════════════
//   6. DATA BARANG (EXCEL)
// ════════════════════════════════════════════════════════════════

let DATA_BARANG = [];

function loadExcel() {
  if (!fs.existsSync(CONFIG.paths.excel)) {
    log.error('EXCEL', 'File tidak ditemukan: ' + CONFIG.paths.excel);
    return false;
  }
  try {
    const wb = xlsx.readFile(CONFIG.paths.excel);
    const ws = wb.Sheets[wb.SheetNames[0]];
    
    // Baca semua data sebagai array of array
    const allRows = xlsx.utils.sheet_to_json(ws, { 
      header: 1, 
      defval: 0,
      blankrows: false 
    });
    
    if (allRows.length < 3) {
      log.error('EXCEL', 'Data terlalu sedikit, minimal 3 baris');
      return false;
    }
    
    // Baris ke-2 (index 1) adalah header asli
    const headers = allRows[1].map(function(h) { 
      return String(h || '').trim(); 
    });
    
    log.info('EXCEL', 'Headers terdeteksi: ' + JSON.stringify(headers));
    
    // Cari index kolom berdasarkan nama
    function findCol(name) {
      const idx = headers.findIndex(function(h) {
        return h.toLowerCase().replace(/\s+/g, ' ').trim() === name.toLowerCase();
      });
      return idx;
    }
    
    const colMap = {
      kode:   findCol('Kode Item'),
      nama:   findCol('Nama Item'),
      jenis:  findCol('Jenis'),
      merek:  findCol('Merek'),
      satuan: findCol('Satuan'),
    };
    
    // Map kolom per toko
    const tokoColMap = {};
    Object.keys(TOKO_COLS).forEach(function(kode) {
      const c = TOKO_COLS[kode];
      tokoColMap[kode] = {
        ecer:  findCol(c.ecer),
        ambil: findCol(c.ambil),
        stok:  findCol(c.stok),
      };
    });
    
    log.info('EXCEL', 'Kolom info: ' + JSON.stringify(colMap));
    log.info('EXCEL', 'Kolom toko: ' + JSON.stringify(tokoColMap));
    
    // Validasi kolom wajib
    if (colMap.kode === -1 || colMap.nama === -1) {
      log.error('EXCEL', 'Kolom "Kode Item" atau "Nama Item" tidak ditemukan!');
      return false;
    }
    
    // Parse data mulai baris ke-3 (index 2)
    DATA_BARANG = [];
    for (let i = 2; i < allRows.length; i++) {
      const row = allRows[i];
      if (!row || row.length === 0) continue;
      
      const kode = String(row[colMap.kode] || '').trim().toUpperCase();
      if (!kode || kode === 'UNDEFINED' || kode === '0') continue;
      
      const item = {
        kode:   kode,
        nama:   String(row[colMap.nama]   || '').trim().toUpperCase(),
        jenis:  String(row[colMap.jenis]  || '').trim(),
        merek:  String(row[colMap.merek]  || '').trim(),
        satuan: String(row[colMap.satuan] || '').trim(),
        harga:  {},
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
    
    // Test: cari NN00001 dan tampilkan harganya
    const test = DATA_BARANG.find(function(d) { return d.kode === 'NN00001'; });
    if (test) {
      log.info('EXCEL', 'Test NN00001 - Ecer NK: ' + test.harga.nk.ecer + ', Ambil NK: ' + test.harga.nk.ambil);
    }
    
    return true;
  } catch (e) {
    log.error('EXCEL', 'Gagal load', e.message);
    return false;
  }
}

function saveExcel() {
  try {
    const rows = DATA_BARANG.map(function(d) {
      const row = {
        'Kode Item': d.kode,
        'Nama Item': d.nama,
        'Jenis':     d.jenis,
        'Merek':     d.merek,
        'Satuan':    d.satuan,
      };
      Object.keys(TOKO_COLS).forEach(function(k) {
        const c = TOKO_COLS[k];
        row[c.ecer]  = d.harga[k].ecer;
        row[c.ambil] = d.harga[k].ambil;
        row[c.stok]  = d.harga[k].stok;
      });
      return row;
    });

    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, xlsx.utils.json_to_sheet(rows), 'Data Barang');
    xlsx.writeFile(wb, CONFIG.paths.excel);
    log.info('EXCEL', 'Saved ' + DATA_BARANG.length + ' item');
    return true;
  } catch (e) {
    log.error('EXCEL', 'Gagal save', e.message);
    return false;
  }
}

function cariBarang(keyword) {
  const q = keyword.trim().toUpperCase();
  const byKode = DATA_BARANG.filter(function(d) { return d.kode === q; });
  if (byKode.length > 0) return byKode;
  const words = q.split(/\s+/).filter(function(w) { return w.length > 0; });
  return DATA_BARANG.filter(function(d) {
    return words.every(function(w) {
      return d.nama.indexOf(w) >= 0 || d.kode.indexOf(w) >= 0;
    });
  });
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
//   7. KIRIM WHATSAPP (FONNTE) - dengan retry
// ════════════════════════════════════════════════════════════════

function tunggu(ms) {
  return new Promise(function(r) { setTimeout(r, ms); });
}

async function kirimWA(target, message, retry) {
  retry = retry || 0;
  try {
    await axios.post(
      CONFIG.fonnteUrl,
      { target: target, message: message },
      { headers: { Authorization: CONFIG.fonnteToken }, timeout: 10000 }
    );
    return true;
  } catch (err) {
    log.warn('FONNTE', 'Gagal kirim ke ' + target + ' (attempt ' + (retry + 1) + ')');
    if (retry < CONFIG.maxRetry - 1) {
      await tunggu(CONFIG.retryDelay);
      return kirimWA(target, message, retry + 1);
    }
    log.error('FONNTE', 'Gagal kirim ke ' + target + ' setelah ' + CONFIG.maxRetry + 'x', err.message);
    return false;
  }
}

// ════════════════════════════════════════════════════════════════
//   8. GEMINI AI
// ════════════════════════════════════════════════════════════════

async function analisaGambar(imageUrl, prompt) {
  const imgResp = await axios.get(imageUrl, { responseType: 'arraybuffer', timeout: 15000 });
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

function buatPromptAI(menuType, namaToko, tanggal) {
  const fmt = ' Format WhatsApp dengan emoji. Format rupiah Rp X.XXX.XXX.';
  if (menuType === 1) return 'Baca data penjualan toko "' + namaToko + '" tanggal ' + tanggal + '. Buat laporan lengkap: total per kassa, total keseluruhan, metode pembayaran, jenis penjualan.' + fmt;
  if (menuType === 2) return 'Baca data harga barang toko "' + namaToko + '" tanggal ' + tanggal + '. Buat laporan: barang baru, naik harga, turun harga.' + fmt;
  if (menuType === 3) return 'Baca data marketplace tanggal ' + tanggal + '. Buat laporan: per toko, per channel, metode bayar.' + fmt;
  return 'Buat laporan rapi.' + fmt;
}

// ════════════════════════════════════════════════════════════════
//   9. DETEKSI SAPAAN & TERIMA KASIH
// ════════════════════════════════════════════════════════════════

const KATA_SAPAAN = [
  'halo','hai','hi','hello','hey','hei','holla','ola','yo','sup','howdy',
  'pagi','siang','sore','malam','selamat',
  'met pagi','met siang','met sore','met malam',
  'assalamualaikum','assalamu','waalaikumsalam','waalaikum','salam',
  'permisi','maaf ganggu','excuse me',
  'good morning','good afternoon','good evening','good night',
];

const KATA_TERIMAKASIH = [
  'terima kasih','terimakasih','makasih','thanks','thank you',
  'thx','tq','ty','tengkyu','tengkyuu','mksh','trims','trimakasih',
];

function cocokKata(low, kata) {
  return low === kata
    || low.startsWith(kata + ' ')
    || low.startsWith(kata + ',')
    || low.startsWith(kata + '!')
    || low.startsWith(kata + '.')
    || low.endsWith(' ' + kata);
}

function isSapaan(low)      { return KATA_SAPAAN.some(function(k)      { return cocokKata(low, k); }); }
function isTerimakasih(low) { return KATA_TERIMAKASIH.some(function(k) { return cocokKata(low, k); }); }

function sapaanPertama(sender) {
  const nama  = getNama(sender);
  const waktu = getWaktu();
  return (nama
    ? 'Selamat ' + waktu + ', *' + nama + '*! 😊'
    : 'Selamat ' + waktu + '! 😊')
    + '\n\nKirim *menu* untuk melihat pilihan yang tersedia.';
}

function balasSapaan(sender) {
  const nama  = getNama(sender);
  const waktu = getWaktu();
  return (nama
    ? 'Selamat ' + waktu + ' juga, *' + nama + '*! 😊'
    : 'Selamat ' + waktu + ' juga! 😊')
    + '\n\nAda yang bisa saya bantu?\nKirim *menu* untuk melihat pilihan yang tersedia.';
}

function balasTerimakasih(sender) {
  const nama = getNama(sender);
  const n    = nama ? ', *' + nama + '*' : '';
  const opsi = [
    'Sama-sama' + n + '! 😊\n\nJika ada yang ingin dicari lagi, ketik *menu* ya.',
    'Dengan senang hati' + n + '! 😊\n\nSemoga bermanfaat! Ketik *menu* jika butuh bantuan lagi.',
    'Tentu' + n + '! Senang bisa membantu 😊\n\nAda yang lain? Ketik *menu*.',
  ];
  return opsi[Math.floor(Math.random() * opsi.length)];
}

// ════════════════════════════════════════════════════════════════
//   10. MENU & PROMPT
// ════════════════════════════════════════════════════════════════

function getMenu(nomor) {
  let m = '🤖 *' + CONFIG.appName + '*\n' + GARIS_TEBAL + '\nPilih menu:\n\n' +
    '*1.* 📊 Laporan Penjualan\n' +
    '*2.* 🏷️ Laporan Harga Barang\n' +
    '*3.* 🛒 Laporan Marketplace';

  if (isMember(nomor)) {
    m += '\n*4.* 🔍 Cari Harga Barang\n\n' + GARIS_TEBAL + '\n' +
      'Atau langsung ketik:\n' +
      '• _cari dandang eagle 20_\n' +
      '• _cari NN00001_\n' +
      '• _stok nk NN00001 10_';
  } else {
    m += '\n\n' + GARIS_TEBAL;
  }

  if (isAdmin(nomor)) {
    m += '\n\n👑 *Perintah Admin:*\n' +
      '• _daftar 628xxx_\n' +
      '• _hapus 628xxx_\n' +
      '• _listmember_\n' +
      '• _namakontak 628xxx Nama_\n' +
      '• _hapuskontak 628xxx_\n' +
      '• _listkontak_\n' +
      '• _reload_ — reload Excel\n' +
      '• _info_ — info sistem';
  }
  return m;
}

const MSG_PILIH_TOKO_CARI = (function() {
  let m = '🔍 *Cari Harga Barang*\n' + GARIS_TEBAL + '\nPilih toko:\n\n';
  TOKO_LIST.forEach(function(t, i) { m += '*' + (i + 1) + '.* ' + t.nama + '\n'; });
  return m + '\nBalas *0* untuk kembali';
})();

function msgSiapCari(nm) {
  return '🔍 *Cari Harga Barang*\n🏦 *' + nm + '*\n' + GARIS_TEBAL +
    '\nKetik nama atau kode barang:\n\n' +
    '• _dandang eagle 20_\n• _NN00001_\n• _golden sunkist_\n\n' +
    'Balas *0* untuk kembali.';
}

function msgPilihToko(menuType) {
  const ic = menuType === 1 ? '📊' : menuType === 2 ? '🏷️' : '🛒';
  const jd = menuType === 1 ? 'Laporan Penjualan' : menuType === 2 ? 'Laporan Harga Barang' : 'Laporan Marketplace';
  let m = ic + ' *' + jd + '*\n' + GARIS_TEBAL + '\nPilih toko:\n\n';
  TOKO_LIST.forEach(function(t, i) { m += '*' + (i + 1) + '.* ' + t.nama + '\n'; });
  return m + '\nBalas *0* untuk kembali';
}

function msgPilihHari(nm) {
  return '🏦 *' + nm + '*\n' + GARIS_TEBAL +
    '\nLaporan untuk:\n\n*1.* 📅 Hari ini\n*2.* 📅 Kemarin\n\nBalas *0* untuk kembali';
}

function msgSiapInput(nm, kemarin, menuType) {
  const t = getTanggal(kemarin);
  const k = kemarin ? ' _(kemarin)_' : '';
  let contoh = '';
  if (menuType === 1) {
    contoh = 'k1 29000000\nk2 11000000\ntunai 26000000\ndebit 14000000\nkredit 0\necer 23000000\ngrosir 17000000';
  } else if (menuType === 2) {
    contoh = '---baru---\nNama barang baru\n---naik---\nNama barang naik\n---turun---\nNama barang turun';
  } else {
    contoh = 'oesapa 0\ntdm 0\ncentral 21061000\nwa 21061000\nshopee 0\ntiktok 0\ntokopedia 0\ntunai 304000\ndebit 20757000\nkredit 0\nnota 019';
  }
  return '✅ *Siap Input!*\n🏦 ' + nm + '\n📅 ' + t + k + '\n' + GARIS_TEBAL +
    '\nKirim *FOTO* atau ketik data:\n\n' + contoh + '\n\nBalas *0* untuk batal.';
}

// ════════════════════════════════════════════════════════════════
//   11. FORMAT HASIL CARI
// ════════════════════════════════════════════════════════════════

function formatHasil(items, tokoKode, sender) {
  const namaToko = NAMA_TOKO[tokoKode] || '-';
  const nama     = getNama(sender);
  const suffix   = '\n\n' + (nama ? 'Semoga membantu, *' + nama + '*! 😊' : 'Semoga membantu! 😊');

  if (items.length === 0) {
    return '❌ *Barang Tidak Ditemukan*\n' + GARIS_TEBAL +
      '\nCoba kata kunci berbeda.\n\n💡 *Tips:*\n' +
      '• Gunakan kata kunci lebih singkat\n' +
      '• Cari by kode (contoh: _NN00001_)\n' +
      '• Periksa ejaan';
  }

  if (items.length > 10) {
    return '⚠️ Ditemukan *' + items.length + ' barang*.\n\n' +
      'Terlalu banyak, gunakan kata kunci lebih spesifik.';
  }

  if (items.length === 1) {
    const d = items[0];
    const h = d.harga[tokoKode];
    return '🏷️ *Detail Barang*\n🏦 *' + namaToko + '*\n' + GARIS_TEBAL + '\n' +
      '🔖 *Kode*   : ' + d.kode + '\n' +
      '📦 *Nama*   : ' + d.nama + '\n' +
      '🏷️ *Jenis*  : ' + (d.jenis || '-') + '\n' +
      '🏗️ *Merek*  : ' + (d.merek || '-') + '\n' +
      '📏 *Satuan* : ' + d.satuan + '\n' + GARIS_TEBAL + '\n' +
      '💰 *Harga Ecer*  : ' + fRp(h.ecer)  + '\n' +
      '💰 *Harga Ambil* : ' + fRp(h.ambil) + '\n' +
      '📊 *Stok*        : ' + (h.stok > 0 ? h.stok + ' ' + d.satuan : '⚠️ Kosong') + '\n' +
      GARIS_TEBAL + suffix;
  }

  let msg = '🔍 *Ditemukan ' + items.length + ' Barang*\n🏦 *' + namaToko + '*\n' + GARIS_TEBAL + '\n';
  items.forEach(function(d, i) {
    const h = d.harga[tokoKode];
    msg += (i + 1) + '. *' + d.nama + '*\n';
    msg += '   🔖 ' + d.kode + ' | ' + d.satuan + '\n';
    msg += '   💰 Ecer: ' + fRp(h.ecer) + ' | Ambil: ' + fRp(h.ambil) + '\n';
    msg += '   📊 Stok: ' + (h.stok > 0 ? h.stok + ' ' + d.satuan : '⚠️ Kosong') + '\n';
    if (i < items.length - 1) msg += GARIS_TIPIS + '\n';
  });
  return msg + GARIS_TEBAL + suffix;
}

// ════════════════════════════════════════════════════════════════
//   12. GENERATOR LAPORAN
// ════════════════════════════════════════════════════════════════

function genLapPenjualan(text, namaToko, kemarin) {
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
    '📅 *' + t + '*' + k + '\n\n' +
    '💵 *PENJUALAN PER KASSA*\n' + ks + '\n' +
    '📦 *TOTAL KESELURUHAN*\n' + fRpP(tot) + '\n\n' +
    '💳 *METODE PEMBAYARAN*\n' +
    '• Tunai  : ' + fRpP(d.tunai  || 0) + '\n' +
    '• Debit  : ' + fRpP(d.debit  || 0) + '\n' +
    '• Kredit : ' + fRpP(d.kredit || 0) + '\n\n' +
    '🛒 *JENIS PENJUALAN*\n' +
    '• Ecer   : ' + fRpP(d.ecer   || 0) + '\n' +
    '• Grosir : ' + fRpP(d.grosir || 0) + '\n' +
    GARIS_TEBAL + '\n_Laporan otomatis_';
}

function genLapHarga(text, namaToko, kemarin) {
  const t = getTanggal(kemarin);
  const h = kemarin ? 'Kemarin' : 'Ini';
  const k = kemarin ? ' _(kemarin)_' : '';
  const s = 'Selamat ' + getWaktu() + ' Team ' + namaToko;
  const d = { baru: [], naik: [], turun: [], note: [] };
  let mode = null;

  text.trim().split('\n').forEach(function(line) {
    const tr = line.trim();
    if (!tr) return;
    const lo = tr.toLowerCase();
    if (lo.indexOf('---baru---')  >= 0 || lo === 'baru')  { mode = 'baru';  return; }
    if (lo.indexOf('---naik---')  >= 0 || lo === 'naik')  { mode = 'naik';  return; }
    if (lo.indexOf('---turun---') >= 0 || lo === 'turun') { mode = 'turun'; return; }
    if (lo.indexOf('---note---')  >= 0 || lo === 'note')  { mode = 'note';  return; }
    if (mode) d[mode].push(tr);
  });

  const cat = 'Nota Semuanya Sudah Diinput Di Sistem, Bisa Langsung Di Print Barcodenya Ya.\n\n' +
    'Mohon Dicek Kembali Fisik Barang Dengan Yang Di Input Disistem, Jika Ada Yang Tidak Sesuai Mohon Di Konfirmasi Lagi. Terima Kasih🙏🏼';

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
    const tr = line.trim();
    if (!tr) return;
    if (tr.indexOf('nota ') === 0) { d.nota.push(line.trim().substring(5)); return; }
    const p = tr.split(/\s+/);
    if (p.length >= 2 && p[0] in d) {
      d[p[0]] = parseFloat(p.slice(1).join('').replace(/[^0-9]/g, '')) || 0;
    }
  });

  const tT = d.oesapa + d.tdm + d.central;
  const tC = d.wa + d.shopee + d.tiktok + d.tokopedia;
  let nt = '';
  if (d.nota.length > 0) {
    nt = '\n';
    d.nota.forEach(function(n) { nt += '- Nomor Nota ' + n + '\n'; });
  }

  return '━━━━━━━━━━━━━━━━━━━━━━\n🛒 *Total Penjualan Marketplace*\n*Perabot Mama*\n📅 Periode ' + t + k +
    '\n━━━━━━━━━━━━━━━━━━━━━━\n🏦 *Per Toko*\n' +
    '• Toko Perabot Mama Oesapa : ' + fRp(d.oesapa)  + '\n' +
    '• Toko Perabot Mama TDM    : ' + fRp(d.tdm)     + '\n' +
    '• Toko Central Perabot     : ' + fRp(d.central) + '\n' +
    GARIS_TIPIS + '\n💰 *Total* : ' + fRp(tT) + '\n\n' +
    '📱 *Per Channel*\n' +
    '• WA        : ' + fRp(d.wa)        + '\n' +
    '• Shopee    : ' + fRp(d.shopee)    + '\n' +
    '• Tiktok    : ' + fRp(d.tiktok)    + '\n' +
    '• Tokopedia : ' + fRp(d.tokopedia) + '\n' +
    GARIS_TIPIS + '\n💰 *Total Penjualan* : ' + fRp(tC) + '\n\n' +
    '💳 *Metode Bayar*\n' +
    '• Tunai/CASH : ' + fRp(d.tunai)  + '\n' +
    '• Debit/TF   : ' + fRp(d.debit)  + '\n' +
    '• Credit     : ' + fRp(d.kredit) + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━\n' + nt + '_Laporan otomatis_';
}

// ════════════════════════════════════════════════════════════════
//   13. HANDLER ADMIN
// ════════════════════════════════════════════════════════════════

async function handleAdmin(sender, msg, low) {
  // Daftar member
  if (low.startsWith('daftar ')) {
    const nomor = msg.substring(7).trim().replace(/[^0-9]/g, '');
    if (!nomor) { await kirimWA(sender, '⚠️ Format: _daftar 6281234567890_'); return true; }
    const r = tambahMember(nomor);
    if (!r.ok) { await kirimWA(sender, '⚠️ Gagal: ' + r.alasan); return true; }
    await kirimWA(sender, '✅ *Member Terdaftar!*\n' + GARIS_TEBAL +
      '\n📱 Nomor: *' + nomor + '*\n👤 Nama: ' + (getNama(nomor) || '(belum ada)') +
      '\n👥 Total: ' + MEMBERS.length);
    return true;
  }

  // Hapus member
  if (low.startsWith('hapus ')) {
    const nomor = msg.substring(6).trim().replace(/[^0-9]/g, '');
    const r = hapusMember(nomor);
    if (!r.ok) { await kirimWA(sender, '❌ ' + r.alasan); return true; }
    await kirimWA(sender, '✅ Member *' + nomor + '* dihapus!\n👥 Total: ' + MEMBERS.length);
    return true;
  }

  // List member
  if (low === 'listmember') {
    let m = '👥 *Daftar Member (' + MEMBERS.length + '):*\n' + GARIS_TEBAL + '\n';
    MEMBERS.forEach(function(n, i) {
      m += (i + 1) + '. ' + n + '\n   👤 ' + (getNama(n) || '(belum ada nama)') + '\n';
    });
    m += '\n👑 Admin: ' + CONFIG.adminNumber;
    await kirimWA(sender, m);
    return true;
  }

  // Set nama kontak
  if (low.startsWith('namakontak ')) {
    const p = msg.substring(11).trim().split(/\s+/);
    if (p.length < 2) { await kirimWA(sender, '⚠️ Format: _namakontak 628xxx Nama_'); return true; }
    const nomor = p[0].replace(/[^0-9]/g, '');
    const nama  = p.slice(1).join(' ');
    const r = setNama(nomor, nama);
    if (!r.ok) { await kirimWA(sender, '❌ ' + r.alasan); return true; }
    await kirimWA(sender, '✅ *Nama Disimpan!*\n📱 ' + nomor + '\n👤 ' + nama);
    return true;
  }

  // Hapus kontak
  if (low.startsWith('hapuskontak ')) {
    const nomor = msg.substring(12).trim().replace(/[^0-9]/g, '');
    const r = hapusKontak(nomor);
    if (!r.ok) { await kirimWA(sender, '❌ ' + r.alasan); return true; }
    await kirimWA(sender, '✅ Kontak *' + nomor + '* dihapus!');
    return true;
  }

  // List kontak
  if (low === 'listkontak') {
    const keys = Object.keys(KONTAK);
    let m = '📒 *Daftar Kontak (' + keys.length + '):*\n' + GARIS_TEBAL + '\n';
    keys.forEach(function(k, i) { m += (i + 1) + '. ' + k + '\n   👤 ' + KONTAK[k] + '\n'; });
    await kirimWA(sender, m);
    return true;
  }

  // Reload Excel
  if (low === 'reload') {
    const ok = loadExcel();
    await kirimWA(sender, ok
      ? '✅ Excel di-reload!\n📦 Total: *' + DATA_BARANG.length + ' item*'
      : '❌ Gagal reload, cek log.');
    return true;
  }

  // Info sistem
  if (low === 'info') {
    const up = Math.floor(process.uptime());
    const jam = Math.floor(up / 3600), mnt = Math.floor((up % 3600) / 60), dtk = up % 60;
    await kirimWA(sender, '🤖 *Info Sistem*\n' + GARIS_TEBAL +
      '\n⏱️ Uptime: ' + jam + 'j ' + mnt + 'm ' + dtk + 'd' +
      '\n👥 Member: ' + MEMBERS.length +
      '\n📦 Data: ' + DATA_BARANG.length + ' item' +
      '\n📒 Kontak: ' + Object.keys(KONTAK).length +
      '\n💬 Sesi aktif: ' + Object.keys(SESI).length +
      '\n💾 Node: ' + process.version);
    return true;
  }

  return false;
}

// ════════════════════════════════════════════════════════════════
//   14. ROUTES
// ════════════════════════════════════════════════════════════════

app.get('/', function(req, res) {
  res.json({
    status: 'ok',
    app: CONFIG.appName,
    uptime: Math.floor(process.uptime()) + 's',
    items: DATA_BARANG.length,
    waktu: new Date().toLocaleString('id-ID'),
  });
});

app.get('/reload', function(req, res) {
  const ok = loadExcel();
  res.json({ success: ok, total: DATA_BARANG.length });
});

// ── DEBUG ENDPOINTS ──────────────────────────────────
app.get('/debug', function(req, res) {
  res.json({
    total: DATA_BARANG.length,
    contoh_5_pertama: DATA_BARANG.slice(0, 5),
  });
});

app.get('/debug/:kode', function(req, res) {
  const kode = req.params.kode.toUpperCase();
  const item = DATA_BARANG.find(function(d) { return d.kode === kode; });
  if (!item) return res.status(404).json({ error: 'Tidak ditemukan', kode: kode });
  res.json(item);
});

// ════════════════════════════════════════════════════════════════
//   15. WEBHOOK UTAMA
// ════════════════════════════════════════════════════════════════

const KATA_RESET = ['0', 'batal', 'menu', 'mulai', 'start', 'kembali'];

app.post('/webhook', async function(req, res) {
  res.sendStatus(200);

  try {
    const body    = req.body || {};
    const sender  = body.sender || body.from || body.phone || null;
    const message = (body.message || body.text || body.msg || '').trim();
    const image   = body.image || body.file || body.media || '';

    if (!sender) { log.warn('WEBHOOK', 'Request tanpa sender'); return; }

    const msg = message;
    const low = msg.toLowerCase();

    log.info('WEBHOOK', 'Pesan dari ' + sender + ': ' + msg.substring(0, 50));

    // ── SAPAAN PERTAMA ──
    if (!SUDAH_DISAPA[sender]) {
      SUDAH_DISAPA[sender] = true;
      saveJSON(CONFIG.paths.disapa, SUDAH_DISAPA);
      await kirimWA(sender, sapaanPertama(sender));
      const trigger = ['menu','halo','hai','hi','mulai','start','hello'];
      if (trigger.indexOf(low) === -1) return;
    }

    // ── ADMIN COMMANDS ──
    if (isAdmin(sender)) {
      const handled = await handleAdmin(sender, msg, low);
      if (handled) return;
    }

    // ── SAPAAN ──
    if (isSapaan(low))      { await kirimWA(sender, balasSapaan(sender));      return; }
    if (isTerimakasih(low)) { await kirimWA(sender, balasTerimakasih(sender)); return; }

    // ── RESET / MENU ──
    if (KATA_RESET.indexOf(low) >= 0) {
      resetSesi(sender);
      await kirimWA(sender, getMenu(sender));
      return;
    }

    // ── CEK AKSES ──
    if ((low.startsWith('cari ') || low.startsWith('stok ')) && !isMember(sender)) {
      await kirimWA(sender, '🚫 *Akses Ditolak*\n\nFitur ini hanya untuk member terdaftar.\n\nHubungi admin untuk mendaftar.');
      return;
    }

    // ── UPDATE STOK ──
    if (low.startsWith('stok ') && isMember(sender)) {
      const p = msg.substring(5).trim().split(/\s+/);
      if (p.length < 3) {
        await kirimWA(sender, '⚠️ Format: _stok [toko] [kode] [jumlah]_\nContoh: _stok nk NN00001 10_\n\nKode: nk / tdm / oesapa / kefa / cp');
        return;
      }
      const tk = p[0].toLowerCase(), kd = p[1], jm = parseInt(p[2]);
      if (!TOKO_COLS[tk])         { await kirimWA(sender, '❌ Kode toko tidak valid (nk/tdm/oesapa/kefa/cp)'); return; }
      if (isNaN(jm) || jm < 0)    { await kirimWA(sender, '❌ Jumlah harus angka positif'); return; }
      const item = updateStok(kd, tk, jm);
      if (!item)                  { await kirimWA(sender, '❌ Kode *' + kd + '* tidak ditemukan'); return; }
      await kirimWA(sender, '✅ *Stok Diperbarui!*\n' + GARIS_TEBAL +
        '\n🏦 ' + NAMA_TOKO[tk] + '\n🔖 ' + item.kode + '\n📦 ' + item.nama +
        '\n📊 Stok baru: *' + jm + ' ' + item.satuan + '*');
      return;
    }

    // ── CARI LANGSUNG ──
    if (low.startsWith('cari ')) {
      const kw = msg.substring(5).trim();
      if (!kw) { await kirimWA(sender, '⚠️ Contoh: _cari dandang eagle 20_'); return; }
      const s = getSesi(sender);
      if (s.mode === 'cari' && s.tokoKode) {
        await kirimWA(sender, formatHasil(cariBarang(kw), s.tokoKode, sender));
      } else {
        updateSesi(sender, { mode: 'cari', pendingKw: kw, tokoKode: null });
        await kirimWA(sender, MSG_PILIH_TOKO_CARI);
      }
      return;
    }

    const s = getSesi(sender);

    // ── MODE CARI: pilih toko ──
    if (s.mode === 'cari' && !s.tokoKode) {
      const idx = parseInt(msg) - 1;
      if (idx >= 0 && idx < TOKO_LIST.length && !isNaN(idx)) {
        const toko = TOKO_LIST[idx];
        updateSesi(sender, { tokoKode: toko.kode });
        if (s.pendingKw) {
          const kw = s.pendingKw;
          updateSesi(sender, { pendingKw: null });
          await kirimWA(sender, formatHasil(cariBarang(kw), toko.kode, sender));
          setTimeout(async function() {
            await kirimWA(sender, '🔍 Cari lagi di *' + toko.nama + '*?\nKetik nama/kode atau *0* kembali ke menu.');
          }, 800);
        } else {
          await kirimWA(sender, msgSiapCari(toko.nama));
        }
      } else {
        await kirimWA(sender, MSG_PILIH_TOKO_CARI);
      }
      return;
    }

    // ── MODE CARI: terima keyword ──
    if (s.mode === 'cari' && s.tokoKode) {
      if (!msg) return;
      await kirimWA(sender, formatHasil(cariBarang(msg), s.tokoKode, sender));
      setTimeout(async function() {
        await kirimWA(sender, '🔍 Cari lagi di *' + NAMA_TOKO[s.tokoKode] + '*?\nKetik nama/kode atau *0* kembali ke menu.');
      }, 800);
      return;
    }

    // ── LEVEL 1: Pilih menu ──
    if (!s.menu) {
      if (msg === '1' || msg === '2') {
        updateSesi(sender, { menu: parseInt(msg) });
        await kirimWA(sender, msgPilihToko(parseInt(msg)));
        return;
      }
      if (msg === '3') {
        updateSesi(sender, { menu: 3 });
        await kirimWA(sender, msgPilihHari('Marketplace Perabot Mama'));
        return;
      }
      if (msg === '4') {
        if (!isMember(sender)) {
          await kirimWA(sender, '🚫 *Akses Ditolak*\n\nFitur ini hanya untuk member.');
          return;
        }
        resetSesi(sender);
        updateSesi(sender, { mode: 'cari' });
        await kirimWA(sender, MSG_PILIH_TOKO_CARI);
        return;
      }
      await kirimWA(sender, '🤔 Maaf, saya tidak mengerti.\n\nKetik *menu* untuk melihat pilihan.');
      return;
    }

    // ── LEVEL 2: Pilih toko (menu 1 & 2) ──
    if (s.menu !== 3 && !s.toko) {
      const idx = parseInt(msg) - 1;
      if (idx >= 0 && idx < TOKO_LIST.length && !isNaN(idx)) {
        updateSesi(sender, { toko: TOKO_LIST[idx].kode });
        await kirimWA(sender, msgPilihHari(TOKO_LIST[idx].nama));
      } else {
        await kirimWA(sender, msgPilihToko(s.menu));
      }
      return;
    }

    // ── LEVEL 3: Pilih hari ──
    if (s.kemarin === undefined || s.kemarin === null) {
      let kem;
      if (msg === '1') kem = false;
      else if (msg === '2') kem = true;
      else {
        const nm = s.menu === 3 ? 'Marketplace Perabot Mama' : NAMA_TOKO[s.toko];
        await kirimWA(sender, msgPilihHari(nm));
        return;
      }
      updateSesi(sender, { kemarin: kem });
      const nm = s.menu === 3 ? 'Marketplace Perabot Mama' : NAMA_TOKO[s.toko];
      await kirimWA(sender, msgSiapInput(nm, kem, s.menu));
      return;
    }

    // ── LEVEL 4: Terima data / foto ──
    const namaToko = s.menu === 3 ? 'Marketplace Perabot Mama' : NAMA_TOKO[s.toko];
    let laporan = '';

    if (image && image.length > 0) {
      await kirimWA(sender, '📸 Foto diterima, sedang dianalisa AI...');
      try {
        const prompt = buatPromptAI(s.menu, namaToko, getTanggal(s.kemarin));
        laporan = await analisaGambar(image, prompt);
      } catch (e) {
        log.error('GEMINI', 'Gagal', e.message);
        await kirimWA(sender, '❌ Gagal baca foto. Coba kirim ulang atau ketik manual.');
        return;
      }
    } else if (msg) {
      if (s.menu === 1) laporan = genLapPenjualan(msg, namaToko, s.kemarin);
      if (s.menu === 2) laporan = genLapHarga(msg, namaToko, s.kemarin);
      if (s.menu === 3) laporan = genLapMarket(msg, s.kemarin);
    } else {
      return;
    }

    if (laporan) {
      await kirimWA(sender, laporan);
      resetSesi(sender);
      const nama = getNama(sender);
      setTimeout(async function() {
        await kirimWA(sender, '✅ Laporan selesai! 😊\n\n' +
          (nama ? 'Terima kasih, *' + nama + '*!\n' : '') +
          'Kirim *menu* untuk laporan berikutnya.');
      }, 1000);
    }

  } catch (err) {
    log.error('WEBHOOK', 'Unhandled error', err.message);
    try {
      const sender = req.body ? (req.body.sender || req.body.from || '') : '';
      if (sender) {
        await kirimWA(sender, '⚠️ Terjadi kesalahan sistem. Coba lagi atau ketik *menu*.');
      }
      await kirimWA(CONFIG.adminNumber, '🚨 *Bot Error*\n' + GARIS_TEBAL +
        '\n📍 WEBHOOK\n❌ ' + err.message + '\n🕐 ' + new Date().toLocaleString('id-ID'));
    } catch (e) {}
  }
});

// ════════════════════════════════════════════════════════════════
//   16. START SERVER
// ════════════════════════════════════════════════════════════════

app.listen(CONFIG.port, function() {
  console.log('\n═══════════════════════════════════════');
  console.log('  ' + CONFIG.appName);
  console.log('═══════════════════════════════════════');
  console.log('  ✅ Status  : AKTIF');
  console.log('  🌐 Port    : ' + CONFIG.port);
  console.log('  👑 Admin   : ' + CONFIG.adminNumber);
  console.log('  📦 Items   : ' + DATA_BARANG.length);
  console.log('  👥 Members : ' + MEMBERS.length);
  console.log('═══════════════════════════════════════\n');
});

process.on('uncaughtException',  function(e) { log.error('SYSTEM', 'Uncaught',   e.message); });
process.on('unhandledRejection', function(r) { log.error('SYSTEM', 'Unhandled', String(r));  });