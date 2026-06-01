'use strict';

// ════════════════════════════════════════════════════════════════
//   CONVERTER: 5 Excel Toko → 1 File harga_barang_5toko.xlsx
//   Jalankan: node convert-excel.js
// ════════════════════════════════════════════════════════════════

const xlsx = require('xlsx');
const path = require('path');
const fs   = require('fs');

// ── KONFIGURASI ──
const TOKO_FILES = [
  { kode: 'cp',     file: 'data-cp.xlsx',     nama: 'Central Perabot' },
  { kode: 'nk',     file: 'data-nk.xlsx',     nama: 'Nasional Kitchen' },
  { kode: 'oesapa', file: 'data-oesapa.xlsx', nama: 'Perabot Mama Oesapa' },
  { kode: 'tdm',    file: 'data-tdm.xlsx',    nama: 'Perabot Mama TDM' },
  { kode: 'kefa',   file: 'data-kefa.xlsx',   nama: 'Perabot Mamaku Kefamenanu' },
];

const HEADER_ROW   = 12; // Baris 13 (index 12) = header kolom
const DATA_START   = 13; // Baris 14 (index 13) = mulai data
const OUTPUT_FILE  = 'harga_barang_5toko.xlsx';

// ── KOLOM MAPPING (berdasarkan screenshot) ──
// A=No, B=Kode Item, C=Nama Item, D=Jenis, E=Stok, F=Satuan, G=Qty/Paket, H=Harga Price(Ecer), I=Harga Jual(Ambil)
const COL = {
  kode:   1,  // B
  nama:   2,  // C
  jenis:  3,  // D
  stok:   4,  // E
  satuan: 5,  // F
  ecer:   7,  // H (Harga Price = Ecer 1-5 pcs)
  ambil:  8,  // I (Harga Jual = Ambil 6+ pcs)
};

// ── MEREK KEYWORDS (untuk extract merek dari nama item) ──
const MEREK_LIST = [
  'EAGLE', 'GLOBAL EAGLE', 'GOLDEN', 'GOLDEN SUNKIST', 'SUNKIST',
  'PARAMOUNT', 'HOCK', 'MIYAKO', 'MAXIM', 'SUNLIFE', 'SUN LIFE',
  'MASPION', 'COSMOS', 'NATIONAL', 'PHILIPS', 'SHARP',
  'KIRIN', 'ADVANCE', 'RINNAI', 'QUANTUM', 'OXONE',
  'BOLDE', 'SIGNORA', 'ELECTROLUX', 'PANASONIC',
  'LOCK&LOCK', 'LION STAR', 'TUPPERWARE',
  'ROYAL', 'NAGAKO', 'AKEBONO', 'ZEBRA',
  'BUTTERFLY', 'BIMA', 'CLARIS', 'GREENLEAF',
  'MYLAND', 'SEAGULL', 'NIKITA', 'AKSARA',
  'DAIKIN', 'SANKEN', 'YONG MA', 'CUCKOO',
];

// Sort dari panjang terpanjang (agar "GOLDEN SUNKIST" match sebelum "GOLDEN")
MEREK_LIST.sort(function(a, b) { return b.length - a.length; });

/**
 * Extract merek dari nama item
 * "DANDANG ALM EAGLE 20 CM" → "EAGLE"
 * "KOMPOR HOCK 10 SUMBU" → "HOCK"
 */
function extractMerek(namaItem) {
  if (!namaItem) return '';
  const upper = namaItem.toUpperCase();
  for (let i = 0; i < MEREK_LIST.length; i++) {
    if (upper.indexOf(MEREK_LIST[i]) >= 0) {
      return MEREK_LIST[i];
    }
  }
  return '';
}

/**
 * Baca data dari 1 file Excel toko
 * Return: array of { kode, nama, jenis, merek, satuan, ecer, ambil, stok }
 */
function bacaFileToko(filePath) {
  if (!fs.existsSync(filePath)) {
    console.error('❌ File tidak ditemukan: ' + filePath);
    return [];
  }
  
  try {
    const wb = xlsx.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const allRows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });
    
    console.log('  📄 Total baris: ' + allRows.length);
    
    // Cek header
    if (allRows.length <= HEADER_ROW) {
      console.error('  ❌ File terlalu pendek, tidak ada data');
      return [];
    }
    
    const headerRow = allRows[HEADER_ROW];
    console.log('  📋 Header: ' + (headerRow || []).slice(0, 10).join(' | '));
    
    const items = [];
    
    for (let i = DATA_START; i < allRows.length; i++) {
      const row = allRows[i];
      if (!row || row.length < 3) continue;
      
      const kode = String(row[COL.kode] || '').trim().toUpperCase();
      const nama = String(row[COL.nama] || '').trim().toUpperCase();
      
      // Skip baris kosong atau tanpa kode
      if (!kode || kode.length < 2 || !nama) continue;
      // Skip kalau kode = "Kode Item" (header duplikat)
      if (kode === 'KODE ITEM' || kode === 'KODE' || kode === 'NO') continue;
      
      const jenis  = String(row[COL.jenis]  || '').trim();
      const satuan = String(row[COL.satuan] || '').trim();
      const stok   = parseInt(row[COL.stok])   || 0;
      const ecer   = parseFloat(row[COL.ecer]) || 0;
      const ambil  = parseFloat(row[COL.ambil])|| 0;
      const merek  = extractMerek(nama);
      
      items.push({ kode, nama, jenis, merek, satuan, stok, ecer, ambil });
    }
    
    console.log('  ✅ Data valid: ' + items.length + ' item');
    return items;
    
  } catch (e) {
    console.error('  ❌ Error baca file: ' + e.message);
    return [];
  }
}

/**
 * Gabungkan data dari 5 toko jadi 1 master
 * - Kalau kode sama → gabungkan harga/stok per toko
 * - Kalau kode beda → tambah sebagai item baru
 */
function gabungkanData(dataToko) {
  const master = {}; // kode → { kode, nama, jenis, merek, satuan, harga: { nk: {ecer, ambil, stok}, ... } }
  
  TOKO_FILES.forEach(function(toko) {
    const items = dataToko[toko.kode] || [];
    
    items.forEach(function(item) {
      if (!master[item.kode]) {
        // Item baru
        master[item.kode] = {
          kode:   item.kode,
          nama:   item.nama,
          jenis:  item.jenis,
          merek:  item.merek,
          satuan: item.satuan,
          harga:  {},
        };
        
        // Inisialisasi semua toko dengan 0
        TOKO_FILES.forEach(function(t) {
          master[item.kode].harga[t.kode] = { ecer: 0, ambil: 0, stok: 0 };
        });
      }
      
      // Update nama/jenis/merek kalau lebih lengkap
      if (item.nama.length > master[item.kode].nama.length) {
        master[item.kode].nama = item.nama;
      }
      if (item.jenis && !master[item.kode].jenis) {
        master[item.kode].jenis = item.jenis;
      }
      if (item.merek && !master[item.kode].merek) {
        master[item.kode].merek = item.merek;
      }
      if (item.satuan && !master[item.kode].satuan) {
        master[item.kode].satuan = item.satuan;
      }
      
      // Set harga & stok untuk toko ini
      master[item.kode].harga[toko.kode] = {
        ecer:  item.ecer,
        ambil: item.ambil,
        stok:  item.stok,
      };
    });
  });
  
  return Object.values(master);
}

/**
 * Tulis ke file output format bot
 */
function tulisOutput(masterData) {
  // Baris 1: Header grup (INFO BARANG | NASIONAL KITCHEN | TDM | ...)
  const header1 = ['INFO BARANG', '', '', '', '',
    'NASIONAL KITCHEN', '', '',
    'PERABOT MAMA TDM', '', '',
    'PERABOT MAMA OESAPA', '', '',
    'PERABOT MAMAKU KEFAMENANU', '', '',
    'CENTRAL PERABOT', '', ''
  ];
  
  // Baris 2: Header kolom detail
  const header2 = [
    'Kode Item', 'Nama Item', 'Jenis', 'Merek', 'Satuan',
    'Ecer NK', 'Ambil NK', 'Stok NK',
    'Ecer TDM', 'Ambil TDM', 'Stok TDM',
    'Ecer Oesapa', 'Ambil Oesapa', 'Stok Oesapa',
    'Ecer Kefa', 'Ambil Kefa', 'Stok Kefa',
    'Ecer CP', 'Ambil CP', 'Stok CP',
  ];
  
  // Data rows
  const rows = [header1, header2];
  
  masterData.forEach(function(d) {
    rows.push([
      d.kode, d.nama, d.jenis, d.merek, d.satuan,
      d.harga.nk.ecer,     d.harga.nk.ambil,     d.harga.nk.stok,
      d.harga.tdm.ecer,    d.harga.tdm.ambil,    d.harga.tdm.stok,
      d.harga.oesapa.ecer, d.harga.oesapa.ambil,  d.harga.oesapa.stok,
      d.harga.kefa.ecer,   d.harga.kefa.ambil,    d.harga.kefa.stok,
      d.harga.cp.ecer,     d.harga.cp.ambil,      d.harga.cp.stok,
    ]);
  });
  
  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.aoa_to_sheet(rows);
  xlsx.utils.book_append_sheet(wb, ws, 'Data Barang');
  xlsx.writeFile(wb, OUTPUT_FILE);
  
  console.log('\n✅ File output: ' + OUTPUT_FILE);
  console.log('📊 Total item: ' + masterData.length);
}

// ════════════════════════════════════════════════════════════════
//   MAIN PROGRAM
// ════════════════════════════════════════════════════════════════

console.log('\n═══════════════════════════════════════');
console.log('  CONVERTER: 5 Excel → 1 File Bot');
console.log('═══════════════════════════════════════\n');

// Step 1: Baca semua file
const dataToko = {};
let totalItems = 0;

TOKO_FILES.forEach(function(toko) {
  console.log('📂 Membaca: ' + toko.file + ' (' + toko.nama + ')');
  const items = bacaFileToko(path.join(__dirname, toko.file));
  dataToko[toko.kode] = items;
  totalItems += items.length;
  console.log('');
});

console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
console.log('📊 Total item dari semua toko: ' + totalItems);

// Step 2: Gabungkan
console.log('\n🔄 Menggabungkan data...');
const masterData = gabungkanData(dataToko);
console.log('✅ Item unik (setelah gabung): ' + masterData.length);

// Step 3: Tulis output
console.log('\n📝 Menulis file output...');
tulisOutput(masterData);

console.log('\n═══════════════════════════════════════');
console.log('  ✅ SELESAI!');
console.log('═══════════════════════════════════════');
console.log('\n📋 Langkah selanjutnya:');
console.log('   1. Upload ' + OUTPUT_FILE + ' ke GitHub');
console.log('   2. Kirim "reload" ke bot WhatsApp');
console.log('   3. Bot pakai data terbaru!\n');
