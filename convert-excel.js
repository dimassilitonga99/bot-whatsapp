'use strict';

const xlsx = require('xlsx');
const path = require('path');
const fs   = require('fs');

const TOKO_FILES = [
  { kode: 'cp',     file: 'data-cp.xlsx',     nama: 'Central Perabot' },
  { kode: 'nk',     file: 'data-nk.xlsx',     nama: 'Nasional Kitchen' },
  { kode: 'oesapa', file: 'data-oesapa.xlsx', nama: 'Perabot Mama Oesapa' },
  { kode: 'tdm',    file: 'data-tdm.xlsx',    nama: 'Perabot Mama TDM' },
  { kode: 'kefa',   file: 'data-kefa.xlsx',   nama: 'Perabot Mamaku Kefamenanu' },
];

const OUTPUT_FILE = 'harga_barang_5toko.xlsx';

const MEREK_LIST = [
  'GLOBAL EAGLE','GOLDEN SUNKIST','SUN LIFE','LION STAR',
  'EAGLE','GOLDEN','SUNKIST','PARAMOUNT','HOCK','MIYAKO','MAXIM','SUNLIFE',
  'MASPION','COSMOS','NATIONAL','PHILIPS','SHARP','KIRIN','ADVANCE',
  'RINNAI','QUANTUM','OXONE','BOLDE','SIGNORA','PANASONIC',
  'ROYAL','NAGAKO','AKEBONO','ZEBRA','BUTTERFLY','BIMA','CLARIS',
  'GREENLEAF','MYLAND','SEAGULL','NIKITA','AKSARA',
  'SANKEN','YONG MA','CUCKOO',
];
MEREK_LIST.sort(function(a, b) { return b.length - a.length; });

function extractMerek(namaItem) {
  if (!namaItem) return '';
  const upper = namaItem.toUpperCase();
  for (let i = 0; i < MEREK_LIST.length; i++) {
    if (upper.indexOf(MEREK_LIST[i]) >= 0) return MEREK_LIST[i];
  }
  return '';
}

/**
 * Auto-detect header row & kolom mapping
 */
function findHeaderAndCols(allRows) {
  // Cari baris yang mengandung "Kode" dan "Nama" (case insensitive)
  for (let i = 0; i < Math.min(allRows.length, 30); i++) {
    const row = allRows[i];
    if (!row) continue;
    
    const rowStr = row.map(function(c) { return String(c || '').toLowerCase(); });
    
    // Cek apakah baris ini mengandung header kolom
    let kodeIdx = -1, namaIdx = -1, jenisIdx = -1, stokIdx = -1, satuanIdx = -1, ecerIdx = -1, ambilIdx = -1;
    
    for (let j = 0; j < rowStr.length; j++) {
      const val = rowStr[j].trim();
      
      if (val.indexOf('kode') >= 0 && val.indexOf('item') >= 0) kodeIdx = j;
      else if (val === 'kode' || val === 'kode item' || val === 'kode barang') kodeIdx = j;
      
      if (val.indexOf('nama') >= 0 && (val.indexOf('item') >= 0 || val.indexOf('barang') >= 0)) namaIdx = j;
      else if (val === 'nama item' || val === 'nama barang' || val === 'description') namaIdx = j;
      
      if (val === 'jenis' || val === 'kategori' || val === 'category') jenisIdx = j;
      if (val === 'stok' || val === 'stock' || val === 'qty' || val === 'jumlah') stokIdx = j;
      if (val === 'satuan' || val === 'unit' || val === 'uom') satuanIdx = j;
      
      if (val.indexOf('harga') >= 0 && val.indexOf('price') >= 0) ecerIdx = j;
      else if (val === 'harga price' || val === 'harga ecer' || val === 'selling price') ecerIdx = j;
      
      if (val.indexOf('harga') >= 0 && val.indexOf('jual') >= 0) ambilIdx = j;
      else if (val === 'harga jual' || val === 'harga ambil' || val === 'cost price') ambilIdx = j;
    }
    
    // Minimal harus ada kode + nama
    if (kodeIdx >= 0 && namaIdx >= 0) {
      console.log('  📋 Header ditemukan di baris ' + (i + 1));
      console.log('     Kode=' + kodeIdx + ' Nama=' + namaIdx + ' Jenis=' + jenisIdx + 
                  ' Stok=' + stokIdx + ' Satuan=' + satuanIdx + ' Ecer=' + ecerIdx + ' Ambil=' + ambilIdx);
      
      return {
        headerRow: i,
        dataStart: i + 1,
        cols: {
          kode:   kodeIdx,
          nama:   namaIdx,
          jenis:  jenisIdx >= 0 ? jenisIdx : -1,
          stok:   stokIdx >= 0 ? stokIdx : -1,
          satuan: satuanIdx >= 0 ? satuanIdx : -1,
          ecer:   ecerIdx >= 0 ? ecerIdx : -1,
          ambil:  ambilIdx >= 0 ? ambilIdx : -1,
        }
      };
    }
  }
  
  // Fallback: coba cari baris dengan "No" di kolom pertama
  for (let i = 0; i < Math.min(allRows.length, 30); i++) {
    const row = allRows[i];
    if (!row) continue;
    const firstCol = String(row[0] || '').toLowerCase().trim();
    if (firstCol === 'no' || firstCol === 'no.') {
      console.log('  📋 Header (fallback "No") di baris ' + (i + 1));
      // Asumsi format: No | Kode | Nama | Jenis | Stok | Satuan | Qty | Ecer | Ambil
      return {
        headerRow: i,
        dataStart: i + 1,
        cols: { kode: 1, nama: 2, jenis: 3, stok: 4, satuan: 5, ecer: 7, ambil: 8 }
      };
    }
  }
  
  console.log('  ❌ Header tidak ditemukan!');
  return null;
}

function bacaFileToko(filePath) {
  if (!fs.existsSync(filePath)) {
    console.error('  ❌ File tidak ditemukan: ' + filePath);
    return [];
  }
  
  try {
    const wb = xlsx.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const allRows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });
    
    console.log('  📄 Total baris: ' + allRows.length);
    
    // Auto-detect header
    const detected = findHeaderAndCols(allRows);
    if (!detected) {
      console.error('  ❌ Tidak bisa detect header!');
      // Print baris 1-15 untuk debug
      for (let i = 0; i < Math.min(15, allRows.length); i++) {
        console.log('  Baris ' + (i+1) + ': ' + (allRows[i] || []).slice(0, 10).join(' | '));
      }
      return [];
    }
    
    const COL = detected.cols;
    const items = [];
    
    for (let i = detected.dataStart; i < allRows.length; i++) {
      const row = allRows[i];
      if (!row || row.length < 3) continue;
      
      const kode = String(row[COL.kode] || '').trim().toUpperCase();
      const nama = String(row[COL.nama] || '').trim().toUpperCase();
      
      if (!kode || kode.length < 2 || !nama) continue;
      if (kode === 'KODE ITEM' || kode === 'KODE' || kode === 'NO' || kode === 'NO.') continue;
      
      const jenis  = COL.jenis >= 0 ? String(row[COL.jenis] || '').trim() : '';
      const satuan = COL.satuan >= 0 ? String(row[COL.satuan] || '').trim() : '';
      const stok   = COL.stok >= 0 ? (parseInt(row[COL.stok]) || 0) : 0;
      const ecer   = COL.ecer >= 0 ? (parseFloat(row[COL.ecer]) || 0) : 0;
      const ambil  = COL.ambil >= 0 ? (parseFloat(row[COL.ambil]) || 0) : 0;
      const merek  = extractMerek(nama);
      
      items.push({ kode, nama, jenis, merek, satuan, stok, ecer, ambil });
    }
    
    console.log('  ✅ Data valid: ' + items.length + ' item');
    return items;
    
  } catch (e) {
    console.error('  ❌ Error: ' + e.message);
    return [];
  }
}

function gabungkanData(dataToko) {
  const master = {};
  
  TOKO_FILES.forEach(function(toko) {
    const items = dataToko[toko.kode] || [];
    items.forEach(function(item) {
      if (!master[item.kode]) {
        master[item.kode] = {
          kode: item.kode, nama: item.nama, jenis: item.jenis,
          merek: item.merek, satuan: item.satuan, harga: {},
        };
        TOKO_FILES.forEach(function(t) {
          master[item.kode].harga[t.kode] = { ecer: 0, ambil: 0, stok: 0 };
        });
      }
      if (item.nama.length > master[item.kode].nama.length) master[item.kode].nama = item.nama;
      if (item.jenis && !master[item.kode].jenis) master[item.kode].jenis = item.jenis;
      if (item.merek && !master[item.kode].merek) master[item.kode].merek = item.merek;
      if (item.satuan && !master[item.kode].satuan) master[item.kode].satuan = item.satuan;
      
      master[item.kode].harga[toko.kode] = { ecer: item.ecer, ambil: item.ambil, stok: item.stok };
    });
  });
  
  return Object.values(master);
}

function tulisOutput(masterData) {
  const header1 = ['INFO BARANG','','','','',
    'NASIONAL KITCHEN','','',
    'PERABOT MAMA TDM','','',
    'PERABOT MAMA OESAPA','','',
    'PERABOT MAMAKU KEFAMENANU','','',
    'CENTRAL PERABOT','',''];
  
  const header2 = [
    'Kode Item','Nama Item','Jenis','Merek','Satuan',
    'Ecer NK','Ambil NK','Stok NK',
    'Ecer TDM','Ambil TDM','Stok TDM',
    'Ecer Oesapa','Ambil Oesapa','Stok Oesapa',
    'Ecer Kefa','Ambil Kefa','Stok Kefa',
    'Ecer CP','Ambil CP','Stok CP'];
  
  const rows = [header1, header2];
  masterData.forEach(function(d) {
    rows.push([
      d.kode, d.nama, d.jenis, d.merek, d.satuan,
      d.harga.nk.ecer, d.harga.nk.ambil, d.harga.nk.stok,
      d.harga.tdm.ecer, d.harga.tdm.ambil, d.harga.tdm.stok,
      d.harga.oesapa.ecer, d.harga.oesapa.ambil, d.harga.oesapa.stok,
      d.harga.kefa.ecer, d.harga.kefa.ambil, d.harga.kefa.stok,
      d.harga.cp.ecer, d.harga.cp.ambil, d.harga.cp.stok,
    ]);
  });
  
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(rows), 'Data Barang');
  xlsx.writeFile(wb, OUTPUT_FILE);
  console.log('\n✅ File output: ' + OUTPUT_FILE);
  console.log('📊 Total item: ' + masterData.length);
}

// ═══ MAIN ═══
console.log('\n═══════════════════════════════════════');
console.log('  CONVERTER: 5 Excel → 1 File Bot');
console.log('═══════════════════════════════════════\n');

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
console.log('📊 Total item semua toko: ' + totalItems);

console.log('\n🔄 Menggabungkan...');
const masterData = gabungkanData(dataToko);
console.log('✅ Item unik: ' + masterData.length);

console.log('\n📝 Menulis output...');
tulisOutput(masterData);

console.log('\n═══════════════════════════════════════');
console.log('  ✅ SELESAI!');
console.log('═══════════════════════════════════════\n');
