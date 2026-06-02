// index.js - Bot WhatsApp Multi-AI dengan Fonnte Gateway
// Platform: Node.js + Express | Deploy: Railway
// LENGKAP UTUH - SATU FILE

const express = require('express');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const axios = require('axios');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ============ ENV VARIABLES ============
const FONNTE_TOKEN = process.env.FONNTE_TOKEN || '';
const GEMINI_KEYS = [
  process.env.GEMINI_KEY || '',
  process.env.GEMINI_KEY2 || '',
  process.env.GEMINI_KEY3 || ''
].filter(k => k);
const GROQ_API_KEY = process.env.GROQ_API_KEY || '';
const OPENROUTER_API_KEY = process.env.OPENROUTER_API_KEY || '';
const ADMIN_NUMBER = process.env.ADMIN_NUMBER || '6285829278962';
const PORT = process.env.PORT || 3000;

// ============ TOKO CONFIG ============
const TOKO_CONFIG = {
  nk: { nama: 'Nasional Kitchen', alias: ['nk', 'nasional kitchen', 'nasional', 'kitchen'] },
  tdm: { nama: 'Perabot Mama TDM', alias: ['tdm', 'perabot mama tdm', 'mama tdm'] },
  oesapa: { nama: 'Perabot Mama Oesapa', alias: ['oesapa', 'perabot mama oesapa', 'mama oesapa'] },
  kefa: { nama: 'Perabot Mamaku Kefamenanu', alias: ['kefa', 'kefamenanu', 'perabot mamaku kefamenanu', 'mamaku kefamenanu', 'mamaku kefa'] },
  cp: { nama: 'Central Perabot (Alak)', alias: ['cp', 'central perabot', 'central', 'alak', 'central perabot alak'] }
};

const CP_KASIR_NAMES = {
  k1: 'Yuni-Salsa',
  k2: 'Nanda-Umi-Marselina',
  k3: 'Febri-Jien-Tika',
  k4: 'Delfi-Tirsa'
};

// Steps laporan per toko
const LAPORAN_STEPS = {
  nk: [
    { key: 'k1', label: 'Kassa 1' },
    { key: 'k2', label: 'Kassa 2' },
    { key: 'total', label: 'Total Penjualan Keseluruhan' },
    { key: 'tunai', label: 'Tunai' },
    { key: 'debit', label: 'Debit' },
    { key: 'kredit', label: 'Credit' },
    { key: 'ecer', label: 'Ecer' },
    { key: 'grosir', label: 'Grosir' }
  ],
  tdm: [
    { key: 'k1', label: 'Kassa 1' },
    { key: 'k2', label: 'Kassa 2' },
    { key: 'total', label: 'Total Penjualan Keseluruhan' },
    { key: 'tunai', label: 'Tunai' },
    { key: 'debit', label: 'Debit' },
    { key: 'kredit', label: 'Credit' }
  ],
  oesapa: [
    { key: 'k1', label: 'Kassa 1' },
    { key: 'k2', label: 'Kassa 2' },
    { key: 'total', label: 'Total Penjualan Keseluruhan' },
    { key: 'tunai', label: 'Tunai' },
    { key: 'debit', label: 'Debit' },
    { key: 'kredit', label: 'Credit' }
  ],
  kefa: [
    { key: 'k1', label: 'Kassa 1' },
    { key: 'k2', label: 'Kassa 2' },
    { key: 'total', label: 'Total Penjualan Keseluruhan' },
    { key: 'tunai', label: 'Tunai' },
    { key: 'debit', label: 'Debit' },
    { key: 'kredit', label: 'Credit' }
  ],
  cp: [
    { key: 'k1', label: 'Kassa 1 (' + CP_KASIR_NAMES.k1 + ')' },
    { key: 'k2', label: 'Kassa 2 (' + CP_KASIR_NAMES.k2 + ')' },
    { key: 'k3', label: 'Kassa 3 (' + CP_KASIR_NAMES.k3 + ')' },
    { key: 'k4', label: 'Kassa 4 (' + CP_KASIR_NAMES.k4 + ')' },
    { key: 'total', label: 'Total Penjualan Keseluruhan' },
    { key: 'tunai', label: 'Tunai' },
    { key: 'debit', label: 'Debit' },
    { key: 'kredit', label: 'Credit' },
    { key: 'ecer', label: 'Ecer' },
    { key: 'grosir', label: 'Grosir' },
    { key: 'promo', label: 'Kasir Promo - Total' },
    { key: 'promotunai', label: 'Kasir Promo - Tunai' },
    { key: 'promodebit', label: 'Kasir Promo - Debit' },
    { key: 'promokredit', label: 'Kasir Promo - Credit' },
    { key: 'parkirkomputer', label: 'Parkir Komputer' },
    { key: 'parkirluar', label: 'Parkir Luar' }
  ]
};

// ============ ROLE ACCESS ============
const ROLE_LAPORAN = [
  '6281584937710',  // Kak Safira
  '6285211988252',  // Kak Admin Marketplace
  '6287841617474',  // Mas Awin
  '6281238774152'   // Ibu Risti
];

const ROLE_LAPORAN_NAMES = {
  '6281584937710': 'Kak Safira',
  '6285211988252': 'Kak Admin Marketplace',
  '6287841617474': 'Mas Awin',
  '6281238774152': 'Ibu Risti'
};

function isAdmin(number) {
  return number === ADMIN_NUMBER;
}

function isRoleLaporan(number) {
  return ROLE_LAPORAN.includes(number) || isAdmin(number);
}

// ============ DATA FILES ============
const SESI_FILE = path.join(__dirname, 'sesi.json');
const MEMBERS_FILE = path.join(__dirname, 'members.json');
const CONTACTS_FILE = path.join(__dirname, 'contacts.json');
const LOG_FILE = path.join(__dirname, 'bot.log');
const EXCEL_FILE = path.join(__dirname, 'harga_barang_5toko.xlsx');

// ============ LOGGING ============
function logToFile(msg) {
  const timestamp = new Date().toISOString();
  const line = `[${timestamp}] ${msg}\n`;
  try {
    fs.appendFileSync(LOG_FILE, line);
  } catch (e) { /* ignore */ }
  console.log(line.trim());
}

// ============ SESSION MANAGEMENT ============
let sessions = {};

function loadSessions() {
  try {
    if (fs.existsSync(SESI_FILE)) {
      const raw = fs.readFileSync(SESI_FILE, 'utf-8');
      sessions = JSON.parse(raw);
      logToFile(`Sessions loaded: ${Object.keys(sessions).length} sesi`);
    }
  } catch (e) {
    logToFile(`Error load sessions: ${e.message}`);
    sessions = {};
  }
}

function saveSessions() {
  try {
    fs.writeFileSync(SESI_FILE, JSON.stringify(sessions, null, 2));
  } catch (e) {
    logToFile(`Error save sessions: ${e.message}`);
  }
}

function getSession(number) {
  if (!sessions[number]) {
    sessions[number] = {
      state: 'idle',
      data: {},
      lastActive: Date.now(),
      lang: 'id',
      greeted: false
    };
    saveSessions();
  }
  sessions[number].lastActive = Date.now();
  return sessions[number];
}

function resetSession(number) {
  sessions[number] = {
    state: 'idle',
    data: {},
    lastActive: Date.now(),
    lang: sessions[number]?.lang || 'id',
    greeted: sessions[number]?.greeted || false
  };
  saveSessions();
}

function deleteSession(number) {
  delete sessions[number];
  saveSessions();
}

function resetAllSessions() {
  sessions = {};
  saveSessions();
  logToFile('All sessions reset');
}

// Auto cleanup 30 menit
setInterval(() => {
  const now = Date.now();
  const threshold = 30 * 60 * 1000;
  let cleaned = 0;
  for (const num of Object.keys(sessions)) {
    if (now - (sessions[num].lastActive || 0) > threshold) {
      delete sessions[num];
      cleaned++;
    }
  }
  if (cleaned > 0) {
    saveSessions();
    logToFile(`Auto-cleanup: ${cleaned} sesi dihapus`);
  }
}, 5 * 60 * 1000);

// ============ MEMBERS MANAGEMENT ============
let members = {};

function loadMembers() {
  try {
    if (fs.existsSync(MEMBERS_FILE)) {
      members = JSON.parse(fs.readFileSync(MEMBERS_FILE, 'utf-8'));
      logToFile(`Members loaded: ${Object.keys(members).length}`);
    }
  } catch (e) {
    logToFile(`Error load members: ${e.message}`);
    members = {};
  }
}

function saveMembers() {
  try {
    fs.writeFileSync(MEMBERS_FILE, JSON.stringify(members, null, 2));
  } catch (e) {
    logToFile(`Error save members: ${e.message}`);
  }
}

// ============ CONTACTS MANAGEMENT ============
let contacts = {};

function loadContacts() {
  try {
    if (fs.existsSync(CONTACTS_FILE)) {
      contacts = JSON.parse(fs.readFileSync(CONTACTS_FILE, 'utf-8'));
      logToFile(`Contacts loaded: ${Object.keys(contacts).length}`);
    }
  } catch (e) {
    logToFile(`Error load contacts: ${e.message}`);
    contacts = {};
  }
}

function saveContacts() {
  try {
    fs.writeFileSync(CONTACTS_FILE, JSON.stringify(contacts, null, 2));
  } catch (e) {
    logToFile(`Error save contacts: ${e.message}`);
  }
}

// ============ EXCEL DATA ============
let excelData = {};

function loadExcel() {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      logToFile('Excel file not found: ' + EXCEL_FILE);
      return;
    }
    const workbook = XLSX.readFile(EXCEL_FILE);
    excelData = {};
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      const tokoKey = sheetName.toLowerCase().trim();
      excelData[tokoKey] = rows;
    }
    logToFile(`Excel loaded. Sheets: ${Object.keys(excelData).join(', ')}`);
    for (const key of Object.keys(excelData)) {
      logToFile(`  Sheet "${key}": ${excelData[key].length} rows, columns: ${excelData[key].length > 0 ? Object.keys(excelData[key][0]).join(', ') : 'empty'}`);
    }
  } catch (e) {
    logToFile(`Error load Excel: ${e.message}`);
  }
}

function getTokoData(tokoKey) {
  // Coba match langsung
  if (excelData[tokoKey]) return excelData[tokoKey];
  // Coba match partial
  for (const key of Object.keys(excelData)) {
    if (key.includes(tokoKey) || tokoKey.includes(key)) return excelData[key];
  }
  return null;
}

// ============ FUZZY SEARCH ============
function levenshtein(a, b) {
  const matrix = [];
  for (let i = 0; i <= b.length; i++) matrix[i] = [i];
  for (let j = 0; j <= a.length; j++) matrix[0][j] = j;
  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b[i - 1] === a[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

function fuzzyMatch(query, text) {
  const q = query.toLowerCase().trim();
  const t = text.toLowerCase().trim();
  if (t.includes(q)) return { match: true, score: 0 };
  // Check each word
  const qWords = q.split(/\s+/);
  let allFound = true;
  for (const w of qWords) {
    if (!t.includes(w)) {
      allFound = false;
      break;
    }
  }
  if (allFound) return { match: true, score: 1 };
  // Levenshtein on words
  const tWords = t.split(/\s+/);
  let totalScore = 0;
  let matchedWords = 0;
  for (const qw of qWords) {
    let bestDist = Infinity;
    for (const tw of tWords) {
      const dist = levenshtein(qw, tw);
      const threshold = Math.max(1, Math.floor(qw.length * 0.4));
      if (dist <= threshold) {
        bestDist = Math.min(bestDist, dist);
      }
    }
    if (bestDist < Infinity) {
      matchedWords++;
      totalScore += bestDist;
    }
  }
  if (matchedWords >= Math.ceil(qWords.length * 0.6)) {
    return { match: true, score: 2 + totalScore };
  }
  return { match: false, score: Infinity };
}

function searchBarang(tokoKey, query) {
  const data = getTokoData(tokoKey);
  if (!data || data.length === 0) return [];

  const results = [];
  const q = query.toLowerCase().trim();

  for (const row of data) {
    // Cari di semua kolom yang mungkin berisi nama/kode
    let namaBarang = '';
    let kodeBarang = '';
    let hargaEcer = '';
    let hargaAmbil = '';

    for (const col of Object.keys(row)) {
      const colLower = col.toLowerCase();
      if (colLower.includes('nama') || colLower.includes('barang') || colLower.includes('item') || colLower.includes('produk') || colLower.includes('product')) {
        namaBarang = String(row[col]);
      }
      if (colLower.includes('kode') || colLower.includes('code') || colLower.includes('sku')) {
        kodeBarang = String(row[col]);
      }
      if (colLower.includes('ecer') || colLower.includes('retail') || (colLower.includes('harga') && colLower.includes('1'))) {
        hargaEcer = row[col];
      }
      if (colLower.includes('ambil') || colLower.includes('grosir') || colLower.includes('wholesale') || (colLower.includes('harga') && colLower.includes('2'))) {
        hargaAmbil = row[col];
      }
    }

    // Fallback: gunakan kolom pertama sebagai kode, kedua sebagai nama
    const cols = Object.keys(row);
    if (!kodeBarang && cols.length > 0) kodeBarang = String(row[cols[0]]);
    if (!namaBarang && cols.length > 1) namaBarang = String(row[cols[1]]);
    if (!hargaEcer && cols.length > 2) hargaEcer = row[cols[2]];
    if (!hargaAmbil && cols.length > 3) hargaAmbil = row[cols[3]];

    const searchText = `${kodeBarang} ${namaBarang}`;
    const fm = fuzzyMatch(q, searchText);

    if (fm.match) {
      results.push({
        kode: kodeBarang,
        nama: namaBarang,
        ecer: hargaEcer,
        ambil: hargaAmbil,
        score: fm.score,
        raw: row
      });
    }
  }

  // Sort by score then abjad
  results.sort((a, b) => {
    if (a.score !== b.score) return a.score - b.score;
    return a.nama.localeCompare(b.nama);
  });

  return results.slice(0, 20);
}

function searchAllToko(query) {
  const allResults = {};
  for (const tokoKey of Object.keys(TOKO_CONFIG)) {
    const results = searchBarang(tokoKey, query);
    if (results.length > 0) {
      allResults[tokoKey] = results;
    }
  }
  return allResults;
}

// ============ FORMAT CURRENCY ============
function formatRp(value) {
  if (value === undefined || value === null || value === '') return 'Rp. 0';
  let num = typeof value === 'string' ? parseInt(value.replace(/[^\d]/g, '')) : Number(value);
  if (isNaN(num)) return 'Rp. 0';
  return 'Rp. ' + num.toLocaleString('id-ID');
}

function parseAngka(text) {
  if (!text) return NaN;
  const cleaned = text.toString().replace(/[^\d]/g, '');
  return cleaned ? parseInt(cleaned) : NaN;
}

// ============ FORMAT TANGGAL ============
function formatTanggal(dayOffset = 0) {
  const d = new Date();
  d.setDate(d.getDate() + dayOffset);
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return `${days[d.getDay()]}, ${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

// ============ SAPAAN & BAHASA ============
const GREETING_PATTERNS = [
  /^(halo|hai|hey|hello|hi|helo|p|pagi|siang|sore|malam|selamat|assalamualaikum|assalamu|salam|apa kabar|waalaikumsalam|wa'alaikumsalam)/i,
  /^(good morning|good afternoon|good evening|good night|howdy|greetings)/i,
  /^(beta mau|beta mo|kaka|kakak|bos|boss|bang|om|mas|mba|mbak|kak)/i
];

const MOTIVASI = [
  "Semangat kerja hari ini ya! 💪",
  "Hari baru, rezeki baru! 🌟",
  "Semoga harimu menyenangkan! 😊",
  "Tetap semangat dan produktif! 🔥",
  "Sukses selalu untuk hari ini! ⭐",
  "Jangan lupa minum air putih ya! 💧",
  "Keep positive and stay awesome! ✨",
  "Bismillah, semoga lancar semua! 🤲",
  "Hari ini pasti lebih baik dari kemarin! 🌈",
  "Mari raih sukses hari ini! 🎯"
];

function detectLanguage(text) {
  const kupangWords = ['beta', 'su', 'sa', 'ko', 'kaka', 'dong', 'bos', 'sonde', 'tra', 'sampe', 'bilang', 'ju', 'ose', 'lu', 'mau', 'karmana', 'bagaimana', 'seng', 'ta', 'lai', 'katong', 'ambe'];
  const englishWords = ['hello', 'hi', 'how', 'are', 'you', 'what', 'where', 'when', 'why', 'which', 'please', 'thank', 'thanks', 'good', 'morning', 'price', 'find', 'search', 'compare'];
  
  const lower = text.toLowerCase();
  const words = lower.split(/\s+/);
  
  let kupangCount = 0;
  let englishCount = 0;
  
  for (const w of words) {
    if (kupangWords.includes(w)) kupangCount++;
    if (englishWords.includes(w)) englishCount++;
  }
  
  if (kupangCount >= 2 || (kupangCount >= 1 && words.length <= 5)) return 'kupang';
  if (englishCount >= 2 || (englishCount >= 1 && words.length <= 3)) return 'en';
  return 'id';
}

function isGreeting(text) {
  for (const pattern of GREETING_PATTERNS) {
    if (pattern.test(text.trim())) return true;
  }
  return false;
}

function getGreetingTime() {
  const hour = new Date().getHours() + 8; // WITA (UTC+8)
  const h = hour >= 24 ? hour - 24 : hour;
  if (h >= 4 && h < 11) return { id: 'Selamat Pagi', en: 'Good Morning', kupang: 'Selamat Pagi' };
  if (h >= 11 && h < 15) return { id: 'Selamat Siang', en: 'Good Afternoon', kupang: 'Selamat Siang' };
  if (h >= 15 && h < 18) return { id: 'Selamat Sore', en: 'Good Afternoon', kupang: 'Selamat Sore' };
  return { id: 'Selamat Malam', en: 'Good Evening', kupang: 'Selamat Malam' };
}

function buildGreetingResponse(lang, contactName) {
  const greet = getGreetingTime();
  const motivasi = MOTIVASI[Math.floor(Math.random() * MOTIVASI.length)];
  const name = contactName || '';

  if (lang === 'kupang') {
    return `${greet.kupang}${name ? ' ' + name : ''} 👋\n\n${motivasi}\n\nBeta bisa bantu apa hari ini?\n\nKetik *menu* untuk lihat menu ya! 😊`;
  }
  if (lang === 'en') {
    return `${greet.en}${name ? ' ' + name : ''} 👋\n\n${motivasi}\n\nHow can I help you today?\n\nType *menu* to see the menu! 😊`;
  }
  return `${greet.id}${name ? ' ' + name : ''} 👋\n\n${motivasi}\n\nAda yang bisa saya bantu hari ini?\n\nKetik *menu* untuk melihat menu ya! 😊`;
}

// ============ AI INTEGRATION ============
let geminiKeyIndex = 0;

async function callGroq(messages) {
  try {
    const resp = await axios.post('https://api.groq.com/openai/v1/chat/completions', {
      model: 'llama-3.1-70b-versatile',
      messages: messages,
      max_tokens: 2048,
      temperature: 0.7
    }, {
      headers: {
        'Authorization': `Bearer ${GROQ_API_KEY}`,
        'Content-Type': 'application/json'
      },
      timeout: 30000
    });
    return resp.data.choices[0].message.content;
  } catch (e) {
    logToFile(`Groq error: ${e.message}`);
    throw e;
  }
}

async function callGemini(prompt) {
  const errors = [];
  for (let i = 0; i < GEMINI_KEYS.length; i++) {
    const keyIdx = (geminiKeyIndex + i) % GEMINI_KEYS.length;
    const key = GEMINI_KEYS[keyIdx];
    try {
      const resp = await axios.post(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${key}`,
        {
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { maxOutputTokens: 2048, temperature: 0.7 }
        },
        { headers: { 'Content-Type': 'application/json' }, timeout: 30000 }
      );
      geminiKeyIndex = (keyIdx + 1) % GEMINI_KEYS.length;
      return resp.data.candidates[0].content.parts[0].text;
    } catch (e) {
      errors.push(`Key${keyIdx + 1}: ${e.message}`);
      logToFile(`Gemini key ${keyIdx + 1} error: ${e.message}`);
    }
  }
  throw new Error(`All Gemini keys failed: ${errors.join('; ')}`);
}

async function callOpenRouter(messages) {
  try {
    const resp = await axios.post('https://openrouter.ai/api/v1/chat/completions', {
      model: 'meta-llama/llama-3.1-70b-instruct:free',
      messages: messages,
      max_tokens: 2048,
      temperature: 0.7
    }, {
      headers: {
        'Authorization': `Bearer ${OPENROUTER_API_KEY}`,
        'Content-Type': 'application/json'
      },
      timeout: 30000
    });
    return resp.data.choices[0].message.content;
  } catch (e) {
    logToFile(`OpenRouter error: ${e.message}`);
    throw e;
  }
}

async function callAI(userMessage, context = '') {
  const systemPrompt = `Kamu adalah asisten bot WhatsApp toko perabot dan peralatan rumah tangga. Kamu membantu menjawab pertanyaan tentang barang, harga, stok, dan rekomendasi.

PENTING:
- Jawab sesuai bahasa user (Indonesia/English/Kupang-NTT)
- SELALU tampilkan harga walau stok kosong (tulis "stok kosong" tapi tetap tulis harganya)
- Gunakan format Rupiah (Rp.)
- Label harga: Ecer(1-5 Pcs) dan Ambil(6 Pcs Keatas)
- Jangan terlalu panjang, ringkas dan informatif
- Jika ada data barang di context, gunakan itu sebagai acuan

${context}`;

  const messages = [
    { role: 'system', content: systemPrompt },
    { role: 'user', content: userMessage }
  ];

  // Try Groq first
  if (GROQ_API_KEY) {
    try {
      return await callGroq(messages);
    } catch (e) {
      logToFile(`Groq failed, trying Gemini...`);
    }
  }

  // Try Gemini
  if (GEMINI_KEYS.length > 0) {
    try {
      return await callGemini(`${systemPrompt}\n\nUser: ${userMessage}`);
    } catch (e) {
      logToFile(`Gemini failed, trying OpenRouter...`);
    }
  }

  // Try OpenRouter
  if (OPENROUTER_API_KEY) {
    try {
      return await callOpenRouter(messages);
    } catch (e) {
      logToFile(`OpenRouter failed too.`);
    }
  }

  return 'Maaf, semua layanan AI sedang tidak tersedia. Silakan coba lagi nanti. 🙏';
}

// ============ SEND WHATSAPP MESSAGE ============
async function sendMessage(to, text, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      const resp = await axios.post('https://api.fonnte.com/send', {
        target: to,
        message: text,
        typing: false
      }, {
        headers: {
          'Authorization': FONNTE_TOKEN
        },
        timeout: 15000
      });
      logToFile(`Sent to ${to}: ${text.substring(0, 100)}...`);
      return resp.data;
    } catch (e) {
      logToFile(`Send failed attempt ${i + 1} to ${to}: ${e.message}`);
      if (i < retries - 1) await new Promise(r => setTimeout(r, 2000));
    }
  }
  logToFile(`FAILED to send to ${to} after ${retries} retries`);
  return null;
}

// ============ MENU TEXTS ============
function getMenuText(number) {
  let menu = '📋 *MENU BOT PERABOT*\n\n';

  if (isRoleLaporan(number)) {
    menu += '1️⃣ Laporan Penjualan\n';
    menu += '2️⃣ Laporan Harga Barang\n';
    menu += '3️⃣ Laporan Marketplace\n';
  }

  menu += '4️⃣ Cari Harga Barang\n';
  menu += '\n💬 Tanya bebas tentang barang (AI)\n';
  menu += '📊 Ketik "bandingkan harga [KODE]" untuk banding harga\n';

  if (isAdmin(number)) {
    menu += '\n9️⃣ Menu Admin\n';
  }

  menu += '\n🌐 Ganti bahasa: ketik *indonesia* / *english* / *kupang*\n';
  menu += '❌ Ketik *batal* untuk kembali ke menu utama';

  return menu;
}

function getAdminMenuText() {
  return `🔧 *MENU ADMIN*\n
*listmember* - Lihat semua member
*listkontak* - Lihat semua kontak
*daftar [nomor] [nama]* - Daftar member baru
*hapus [nomor]* - Hapus member
*namakontak [nomor] [nama]* - Set nama kontak
*hapuskontak [nomor]* - Hapus kontak
*reload* - Reload data Excel
*info* - Info sistem
*resetall* - Reset semua sesi

Ketik *batal* untuk kembali.`;
}

function getTokoListText() {
  let text = '🏪 *PILIH TOKO:*\n\n';
  text += '1. Nasional Kitchen (NK)\n';
  text += '2. Perabot Mama TDM\n';
  text += '3. Perabot Mama Oesapa\n';
  text += '4. Perabot Mamaku Kefamenanu\n';
  text += '5. Central Perabot/Alak (CP)\n';
  text += '\nKetik angka atau nama toko:';
  return text;
}

function parseTokoInput(input) {
  const lower = input.toLowerCase().trim();
  if (lower === '1' || lower === 'nk' || lower.includes('nasional')) return 'nk';
  if (lower === '2' || lower === 'tdm' || lower.includes('tdm')) return 'tdm';
  if (lower === '3' || lower === 'oesapa' || lower.includes('oesapa')) return 'oesapa';
  if (lower === '4' || lower === 'kefa' || lower.includes('kefa')) return 'kefa';
  if (lower === '5' || lower === 'cp' || lower.includes('central') || lower.includes('alak')) return 'cp';
  return null;
}

function parseDayInput(input) {
  const lower = input.toLowerCase().trim();
  if (lower === '1' || lower === 'hari ini' || lower === 'today') return 0;
  if (lower === '2' || lower === 'kemarin' || lower === 'yesterday') return -1;
  if (lower === '3' || lower === 'lusa' || lower === 'kemarin lusa') return -2;
  return null;
}

// ============ LAPORAN FORMAT ============
function buildLaporanPenjualan(tokoKey, data, tanggal) {
  const tokoNama = TOKO_CONFIG[tokoKey].nama;
  let text = `📊 *Laporan Penjualan*\nToko ${tokoNama}\nPeriode ${tanggal}\n\n`;

  if (tokoKey === 'nk') {
    text += `Kassa 1 ${formatRp(data.k1)}\n`;
    text += `Kassa 2 ${formatRp(data.k2)}\n\n`;
    text += `Total Penjualan Keseluruhan\n${formatRp(data.total)}\n`;
    text += `-----\n`;
    text += `Tunai ${formatRp(data.tunai)}\n`;
    text += `Debit ${formatRp(data.debit)}\n`;
    text += `Credit ${formatRp(data.kredit)}\n`;
    text += `-----\n`;
    text += `Ecer : ${formatRp(data.ecer)}\n`;
    text += `Grosir : ${formatRp(data.grosir)}`;
  } else if (tokoKey === 'tdm' || tokoKey === 'oesapa' || tokoKey === 'kefa') {
    text += `Kassa 1 ${formatRp(data.k1)}\n`;
    text += `Kassa 2 ${formatRp(data.k2)}\n\n`;
    text += `Total Penjualan Keseluruhan\n${formatRp(data.total)}\n`;
    text += `-----\n`;
    text += `Tunai ${formatRp(data.tunai)}\n`;
    text += `Debit ${formatRp(data.debit)}\n`;
    text += `Credit ${formatRp(data.kredit)}`;
  } else if (tokoKey === 'cp') {
    text += `Kassa 1 (${CP_KASIR_NAMES.k1}) ${formatRp(data.k1)}\n`;
    text += `Kassa 2 (${CP_KASIR_NAMES.k2}) ${formatRp(data.k2)}\n`;
    text += `Kassa 3 (${CP_KASIR_NAMES.k3}) ${formatRp(data.k3)}\n`;
    text += `Kassa 4 (${CP_KASIR_NAMES.k4}) ${formatRp(data.k4)}\n\n`;
    text += `Total Penjualan Keseluruhan\n${formatRp(data.total)}\n`;
    text += `-----\n`;
    text += `Tunai ${formatRp(data.tunai)}\n`;
    text += `Debit ${formatRp(data.debit)}\n`;
    text += `Credit ${formatRp(data.kredit)}\n`;
    text += `-----\n`;
    text += `Ecer : ${formatRp(data.ecer)}\n`;
    text += `Grosir : ${formatRp(data.grosir)}\n`;
    text += `-----\n`;
    text += `Kasir Promo\n`;
    text += `Total : ${formatRp(data.promo)}\n`;
    text += `Tunai : ${formatRp(data.promotunai)}\n`;
    text += `Debit : ${formatRp(data.promodebit)}\n`;
    text += `Credit : ${formatRp(data.promokredit)}\n`;
    text += `-----\n`;
    text += `Parkir\n`;
    text += `Komputer : ${formatRp(data.parkirkomputer)}\n`;
    text += `Luar : ${formatRp(data.parkirluar)}\n`;
    const parkirTotal = (parseAngka(data.parkirkomputer) || 0) + (parseAngka(data.parkirluar) || 0);
    text += `Total Parkir : ${formatRp(parkirTotal)}`;
  }

  return text;
}

// ============ BANDING HARGA ============
function bandingHarga(query, tokoFilter = null) {
  const allResults = {};
  const tokos = tokoFilter ? [tokoFilter] : Object.keys(TOKO_CONFIG);

  for (const tokoKey of tokos) {
    const results = searchBarang(tokoKey, query);
    if (results.length > 0) {
      allResults[tokoKey] = results[0]; // ambil yang paling relevan
    }
  }

  if (Object.keys(allResults).length === 0) {
    return 'Barang tidak ditemukan di toko manapun. 🔍';
  }

  if (Object.keys(allResults).length === 1) {
    const tk = Object.keys(allResults)[0];
    const item = allResults[tk];
    return `📊 *Hasil Pencarian Harga*\n\nBarang: ${item.nama}\nKode: ${item.kode}\n\nHanya ditemukan di *${TOKO_CONFIG[tk].nama}*:\n- Ecer(1-5 Pcs): ${formatRp(item.ecer)}\n- Ambil(6 Pcs Keatas): ${formatRp(item.ambil)}`;
  }

  let text = `📊 *PERBANDINGAN HARGA*\n\n`;
  text += `Barang: ${Object.values(allResults)[0].nama}\n`;
  text += `Kode: ${Object.values(allResults)[0].kode}\n\n`;

  let minEcer = Infinity, maxEcer = 0, minTokoE = '', maxTokoE = '';
  let minAmbil = Infinity, maxAmbil = 0, minTokoA = '', maxTokoA = '';

  for (const [tk, item] of Object.entries(allResults)) {
    const ecer = parseAngka(item.ecer) || 0;
    const ambil = parseAngka(item.ambil) || 0;
    text += `🏪 *${TOKO_CONFIG[tk].nama}*\n`;
    text += `   Ecer(1-5 Pcs): ${formatRp(item.ecer)}\n`;
    text += `   Ambil(6 Pcs Keatas): ${formatRp(item.ambil)}\n\n`;

    if (ecer > 0 && ecer < minEcer) { minEcer = ecer; minTokoE = TOKO_CONFIG[tk].nama; }
    if (ecer > maxEcer) { maxEcer = ecer; maxTokoE = TOKO_CONFIG[tk].nama; }
    if (ambil > 0 && ambil < minAmbil) { minAmbil = ambil; minTokoA = TOKO_CONFIG[tk].nama; }
    if (ambil > maxAmbil) { maxAmbil = ambil; maxTokoA = TOKO_CONFIG[tk].nama; }
  }

  text += `-----\n📈 *ANALISA:*\n`;
  if (minEcer < Infinity) {
    text += `\nHarga Ecer Termurah: *${minTokoE}* (${formatRp(minEcer)})`;
    text += `\nHarga Ecer Termahal: *${maxTokoE}* (${formatRp(maxEcer)})`;
    if (maxEcer > minEcer) text += `\nSelisih Ecer: ${formatRp(maxEcer - minEcer)}`;
  }
  if (minAmbil < Infinity) {
    text += `\n\nHarga Ambil Termurah: *${minTokoA}* (${formatRp(minAmbil)})`;
    text += `\nHarga Ambil Termahal: *${maxTokoA}* (${formatRp(maxAmbil)})`;
    if (maxAmbil > minAmbil) text += `\nSelisih Ambil: ${formatRp(maxAmbil - minAmbil)}`;
  }

  return text;
}

// ============ FORMAT SEARCH RESULTS ============
function formatSearchResults(results, tokoKey) {
  if (results.length === 0) return 'Barang tidak ditemukan. 🔍\n\nCoba kata kunci lain atau periksa ejaan.';

  let text = `🔍 *Hasil Pencarian di ${TOKO_CONFIG[tokoKey].nama}*\n`;
  text += `Ditemukan ${results.length} barang:\n\n`;

  for (let i = 0; i < results.length; i++) {
    const item = results[i];
    text += `${i + 1}. *${item.nama}*\n`;
    text += `   Kode: ${item.kode}\n`;
    text += `   Ecer(1-5 Pcs): ${formatRp(item.ecer)}\n`;
    text += `   Ambil(6 Pcs Keatas): ${formatRp(item.ambil)}\n\n`;
  }

  text += `\n💡 Ketik nama/kode barang lain untuk mencari lagi`;
  text += `\nKetik *ganti toko* untuk pindah toko`;
  text += `\nKetik *batal* untuk kembali ke menu`;

  return text;
}

// ============ DETECT TOKO MENTION ============
function detectTokoMention(text) {
  const lower = text.toLowerCase();
  for (const [key, config] of Object.entries(TOKO_CONFIG)) {
    for (const alias of config.alias) {
      if (lower.includes(alias)) return key;
    }
  }
  return null;
}

// ============ DETECT BANDING COMMAND ============
function detectBandingCommand(text) {
  const lower = text.toLowerCase().trim();
  const patterns = [
    /^bandingkan?\s+harga\s+(.+)/i,
    /^banding\s+(.+)/i,
    /^compare\s+price\s+(.+)/i,
    /^compare\s+(.+)/i
  ];
  for (const p of patterns) {
    const m = lower.match(p);
    if (m) return m[1].trim();
  }
  return null;
}

// ============ MAIN MESSAGE HANDLER ============
async function handleMessage(sender, message) {
  const text = message.trim();
  const lower = text.toLowerCase().trim();
  const session = getSession(sender);

  logToFile(`[${sender}] State: ${session.state} | Msg: ${text}`);

  // ---- LANGUAGE SWITCH ----
  if (lower === 'indonesia' || lower === 'bahasa indonesia') {
    session.lang = 'id';
    saveSessions();
    return sendMessage(sender, '🇮🇩 Bahasa diubah ke Indonesia.\n\nKetik *menu* untuk melihat menu.');
  }
  if (lower === 'english' || lower === 'inggris') {
    session.lang = 'en';
    saveSessions();
    return sendMessage(sender, '🇬🇧 Language set to English.\n\nType *menu* to see the menu.');
  }
  if (lower === 'kupang' || lower === 'ntt' || lower === 'bahasa kupang') {
    session.lang = 'kupang';
    saveSessions();
    return sendMessage(sender, '🏝️ Bahasa su ganti ke Kupang.\n\nKetik *menu* untuk liat menu ya!');
  }

  // ---- BATAL / CANCEL ----
  if (lower === 'batal' || lower === 'cancel' || lower === 'keluar' || lower === 'exit' || lower === 'back' || lower === 'kembali') {
    resetSession(sender);
    return sendMessage(sender, '✅ Kembali ke menu utama.\n\nKetik *menu* untuk melihat menu.');
  }

  // ---- GREETING DETECTION ----
  if (session.state === 'idle' && isGreeting(text)) {
    const lang = detectLanguage(text);
    session.lang = lang;
    session.greeted = true;
    saveSessions();
    const contactName = contacts[sender] || ROLE_LAPORAN_NAMES[sender] || (isAdmin(sender) ? 'Admin' : '');
    return sendMessage(sender, buildGreetingResponse(lang, contactName));
  }

  // ---- MENU ----
  if (lower === 'menu' || lower === 'help' || lower === 'bantuan' || lower === 'tolong') {
    resetSession(sender);
    return sendMessage(sender, getMenuText(sender));
  }

  // ---- ADMIN COMMANDS ----
  if (lower === '9' && isAdmin(sender)) {
    session.state = 'admin_menu';
    saveSessions();
    return sendMessage(sender, getAdminMenuText());
  }

  if (session.state === 'admin_menu' || (isAdmin(sender) && (
    lower.startsWith('listmember') || lower.startsWith('listkontak') ||
    lower.startsWith('daftar ') || lower.startsWith('hapus ') ||
    lower.startsWith('namakontak ') || lower.startsWith('hapuskontak ') ||
    lower === 'reload' || lower === 'info' || lower === 'resetall'
  ))) {
    return handleAdminCommand(sender, text);
  }

  // ---- MENU 1: LAPORAN PENJUALAN ----
  if (lower === '1') {
    if (!isRoleLaporan(sender)) {
      return sendMessage(sender, '⛔ Maaf, Anda tidak memiliki akses ke menu ini.');
    }
    session.state = 'laporan_pilih_toko';
    session.data = { menu: 1 };
    saveSessions();
    return sendMessage(sender, '📊 *LAPORAN PENJUALAN*\n\n' + getTokoListText());
  }

  // ---- MENU 2: LAPORAN HARGA BARANG ----
  if (lower === '2') {
    if (!isRoleLaporan(sender)) {
      return sendMessage(sender, '⛔ Maaf, Anda tidak memiliki akses ke menu ini.');
    }
    session.state = 'laporan_harga_pilih_toko';
    session.data = { menu: 2 };
    saveSessions();
    return sendMessage(sender, '📋 *LAPORAN HARGA BARANG*\n\n' + getTokoListText());
  }

  // ---- MENU 3: LAPORAN MARKETPLACE ----
  if (lower === '3') {
    if (!isRoleLaporan(sender)) {
      return sendMessage(sender, '⛔ Maaf, Anda tidak memiliki akses ke menu ini.');
    }
    session.state = 'laporan_marketplace';
    session.data = { menu: 3 };
    saveSessions();
    return sendMessage(sender, '🛒 *LAPORAN MARKETPLACE*\n\nFitur ini sedang dalam pengembangan.\n\nKetik *batal* untuk kembali.');
  }

  // ---- MENU 4: CARI HARGA BARANG ----
  if (lower === '4') {
    session.state = 'cari_pilih_toko';
    session.data = { menu: 4 };
    saveSessions();
    return sendMessage(sender, '🔍 *CARI HARGA BARANG*\n\n' + getTokoListText());
  }

  // ---- STATE HANDLERS ----
  
  // == LAPORAN PENJUALAN FLOW ==
  if (session.state === 'laporan_pilih_toko') {
    const toko = parseTokoInput(text);
    if (!toko) {
      return sendMessage(sender, '❌ Toko tidak valid. Silakan pilih:\n\n' + getTokoListText());
    }
    session.data.toko = toko;
    session.state = 'laporan_pilih_hari';
    saveSessions();
    return sendMessage(sender, `Toko: *${TOKO_CONFIG[toko].nama}*\n\n📅 Pilih hari:\n1. Hari ini\n2. Kemarin\n3. Kemarin lusa\n\nKetik angka:`);
  }

  if (session.state === 'laporan_pilih_hari') {
    const dayOffset = parseDayInput(text);
    if (dayOffset === null) {
      return sendMessage(sender, '❌ Pilihan tidak valid.\n\n1. Hari ini\n2. Kemarin\n3. Kemarin lusa');
    }
    session.data.dayOffset = dayOffset;
    session.data.tanggal = formatTanggal(dayOffset);
    
    // Masuk wizard step-by-step
    const toko = session.data.toko;
    const steps = LAPORAN_STEPS[toko];
    session.data.stepIndex = 0;
    session.data.values = {};
    session.state = 'laporan_scan_step';
    saveSessions();

    const step = steps[0];
    return sendMessage(sender, `📊 *Input Laporan ${TOKO_CONFIG[toko].nama}*\nPeriode: ${session.data.tanggal}\n\n📝 Masukkan angka untuk *${step.label}*:\n(Step 1/${steps.length})`);
  }

  if (session.state === 'laporan_scan_step') {
    const toko = session.data.toko;
    const steps = LAPORAN_STEPS[toko];
    const stepIndex = session.data.stepIndex;
    const step = steps[stepIndex];

    // Parse angka
    const angka = parseAngka(text);
    if (isNaN(angka)) {
      return sendMessage(sender, `❌ Input harus berupa angka.\n\nMasukkan angka untuk *${step.label}*:`);
    }

    // Simpan value
    session.data.values[step.key] = angka;

    // Next step
    const nextIndex = stepIndex + 1;
    if (nextIndex >= steps.length) {
      // Selesai - generate laporan
      const laporan = buildLaporanPenjualan(toko, session.data.values, session.data.tanggal);
      resetSession(sender);
      return sendMessage(sender, `✅ Laporan selesai!\n\n${laporan}`);
    }

    session.data.stepIndex = nextIndex;
    saveSessions();

    const nextStep = steps[nextIndex];
    return sendMessage(sender, `✅ ${step.label}: ${formatRp(angka)}\n\n📝 Masukkan angka untuk *${nextStep.label}*:\n(Step ${nextIndex + 1}/${steps.length})`);
  }

  // == LAPORAN HARGA BARANG FLOW ==
  if (session.state === 'laporan_harga_pilih_toko') {
    const toko = parseTokoInput(text);
    if (!toko) {
      return sendMessage(sender, '❌ Toko tidak valid.\n\n' + getTokoListText());
    }
    session.data.toko = toko;
    session.state = 'laporan_harga_cari';
    saveSessions();
    return sendMessage(sender, `📋 *Laporan Harga - ${TOKO_CONFIG[toko].nama}*\n\nKetik nama/kode barang untuk melihat harga:\n\nKetik *batal* untuk kembali.`);
  }

  if (session.state === 'laporan_harga_cari') {
    const toko = session.data.toko;
    const results = searchBarang(toko, text);
    const formatted = formatSearchResults(results, toko);
    return sendMessage(sender, formatted);
  }

  // == CARI BARANG FLOW ==
  if (session.state === 'cari_pilih_toko') {
    const toko = parseTokoInput(text);
    if (!toko) {
      return sendMessage(sender, '❌ Toko tidak valid.\n\n' + getTokoListText());
    }
    session.data.toko = toko;
    session.state = 'cari_barang';
    saveSessions();
    return sendMessage(sender, `🔍 *Cari Barang - ${TOKO_CONFIG[toko].nama}*\n\nKetik nama atau kode barang:\n\nKetik *ganti toko* untuk pindah toko\nKetik *batal* untuk kembali`);
  }

  if (session.state === 'cari_barang') {
    // Ganti toko
    if (lower === 'ganti toko' || lower === 'pindah toko') {
      session.state = 'cari_pilih_toko';
      saveSessions();
      return sendMessage(sender, getTokoListText());
    }

    // Deteksi sebut toko lain
    const mentionedToko = detectTokoMention(text);
    if (mentionedToko && mentionedToko !== session.data.toko) {
      session.state = 'cari_konfirmasi_pindah';
      session.data.pindahToko = mentionedToko;
      saveSessions();
      return sendMessage(sender, `🔄 Anda menyebut toko *${TOKO_CONFIG[mentionedToko].nama}*.\n\nMau pindah ke toko ini?\n1. Ya, pindah\n2. Tidak, tetap di ${TOKO_CONFIG[session.data.toko].nama}`);
    }

    const toko = session.data.toko;
    const results = searchBarang(toko, text);
    const formatted = formatSearchResults(results, toko);
    return sendMessage(sender, formatted);
  }

  if (session.state === 'cari_konfirmasi_pindah') {
    if (lower === '1' || lower === 'ya' || lower === 'yes' || lower === 'iya') {
      session.data.toko = session.data.pindahToko;
      delete session.data.pindahToko;
      session.state = 'cari_barang';
      saveSessions();
      return sendMessage(sender, `✅ Pindah ke *${TOKO_CONFIG[session.data.toko].nama}*\n\nKetik nama/kode barang:`);
    } else {
      delete session.data.pindahToko;
      session.state = 'cari_barang';
      saveSessions();
      return sendMessage(sender, `✅ Tetap di *${TOKO_CONFIG[session.data.toko].nama}*\n\nKetik nama/kode barang:`);
    }
  }

  // ---- BANDING HARGA (any state) ----
  const bandingQuery = detectBandingCommand(text);
  if (bandingQuery) {
    const tokoFilter = detectTokoMention(bandingQuery);
    const cleanQuery = bandingQuery.replace(/(nk|tdm|oesapa|kefa|cp|nasional|central|alak|kitchen|kefamenanu|mama)/gi, '').trim();
    const result = bandingHarga(cleanQuery || bandingQuery, tokoFilter);
    return sendMessage(sender, result);
  }

  // ---- AI CHAT (default) ----
  if (session.state === 'idle' || session.state === 'ai_chat') {
    session.state = 'ai_chat';
    saveSessions();

    // Deteksi bahasa
    const lang = detectLanguage(text);
    session.lang = lang;

    // Build context dari data Excel
    let context = '';
    const mentionedToko = detectTokoMention(text);
    
    // Cek apakah user bertanya tentang barang
    const isBarangQuestion = /harga|stok|stock|barang|produk|item|price|berapa|jual|ada|cari|find|search|recommend|rekomendasi|saran|suggest/i.test(text);
    
    if (isBarangQuestion) {
      // Clean query for search
      let searchQuery = text.replace(/(harga|stok|stock|barang|produk|berapa|jual|ada|cari|tolong|carikan|yang|di|toko|dari|untuk|bisa|kasih|lihat|find|search|price|show)/gi, '').trim();
      
      if (searchQuery.length >= 2) {
        if (mentionedToko) {
          const results = searchBarang(mentionedToko, searchQuery);
          if (results.length > 0) {
            context = `\nData barang ditemukan di ${TOKO_CONFIG[mentionedToko].nama}:\n`;
            for (const r of results.slice(0, 10)) {
              context += `- ${r.nama} (${r.kode}): Ecer ${formatRp(r.ecer)}, Ambil ${formatRp(r.ambil)}\n`;
            }
          }
        } else {
          const allResults = searchAllToko(searchQuery);
          if (Object.keys(allResults).length > 0) {
            context = '\nData barang ditemukan:\n';
            for (const [tk, results] of Object.entries(allResults)) {
              context += `\n${TOKO_CONFIG[tk].nama}:\n`;
              for (const r of results.slice(0, 5)) {
                context += `- ${r.nama} (${r.kode}): Ecer ${formatRp(r.ecer)}, Ambil ${formatRp(r.ambil)}\n`;
              }
            }
          }
        }
      }
    }

    if (lang === 'kupang') context += '\nUser menggunakan bahasa Kupang/NTT. Jawab dalam bahasa Kupang.';
    if (lang === 'en') context += '\nUser is using English. Reply in English.';

    const aiReply = await callAI(text, context);
    return sendMessage(sender, aiReply);
  }

  // Fallback
  return sendMessage(sender, 'Ketik *menu* untuk melihat menu atau tanya langsung tentang barang. 😊');
}

// ============ ADMIN COMMAND HANDLER ============
async function handleAdminCommand(sender, text) {
  const lower = text.toLowerCase().trim();
  const session = getSession(sender);

  if (!isAdmin(sender)) {
    return sendMessage(sender, '⛔ Akses ditolak.');
  }

  // List member
  if (lower === 'listmember') {
    const list = Object.entries(members);
    if (list.length === 0) return sendMessage(sender, '📋 Belum ada member terdaftar.');
    let msg = '📋 *DAFTAR MEMBER*\n\n';
    for (const [num, data] of list) {
      msg += `- ${num}: ${data.nama || 'No Name'}\n`;
    }
    msg += `\nTotal: ${list.length} member`;
    return sendMessage(sender, msg);
  }

  // List kontak
  if (lower === 'listkontak') {
    const list = Object.entries(contacts);
    if (list.length === 0) return sendMessage(sender, '📋 Belum ada kontak.');
    let msg = '📋 *DAFTAR KONTAK*\n\n';
    for (const [num, nama] of list) {
      msg += `- ${num}: ${nama}\n`;
    }
    msg += `\nTotal: ${list.length} kontak`;
    return sendMessage(sender, msg);
  }

  // Daftar member
  if (lower.startsWith('daftar ')) {
    const parts = text.substring(7).trim().split(/\s+/);
    if (parts.length < 2) {
      return sendMessage(sender, 'Format: *daftar [nomor] [nama]*\nContoh: daftar 6281234567890 Budi');
    }
    let nomor = parts[0].replace(/[^\d]/g, '');
    if (!nomor.startsWith('62')) nomor = '62' + nomor.replace(/^0/, '');
    const nama = parts.slice(1).join(' ');
    members[nomor] = { nama, registeredAt: new Date().toISOString() };
    saveMembers();
    return sendMessage(sender, `✅ Member terdaftar!\nNomor: ${nomor}\nNama: ${nama}`);
  }

  // Hapus member
  if (lower.startsWith('hapus ') && !lower.startsWith('hapuskontak')) {
    let nomor = text.substring(6).trim().replace(/[^\d]/g, '');
    if (!nomor.startsWith('62')) nomor = '62' + nomor.replace(/^0/, '');
    if (members[nomor]) {
      delete members[nomor];
      saveMembers();
      return sendMessage(sender, `✅ Member ${nomor} dihapus.`);
    }
    return sendMessage(sender, `❌ Member ${nomor} tidak ditemukan.`);
  }

  // Nama kontak
  if (lower.startsWith('namakontak ')) {
    const parts = text.substring(11).trim().split(/\s+/);
    if (parts.length < 2) {
      return sendMessage(sender, 'Format: *namakontak [nomor] [nama]*');
    }
    let nomor = parts[0].replace(/[^\d]/g, '');
    if (!nomor.startsWith('62')) nomor = '62' + nomor.replace(/^0/, '');
    const nama = parts.slice(1).join(' ');
    contacts[nomor] = nama;
    saveContacts();
    return sendMessage(sender, `✅ Kontak disimpan!\nNomor: ${nomor}\nNama: ${nama}`);
  }

  // Hapus kontak
  if (lower.startsWith('hapuskontak ')) {
    let nomor = text.substring(12).trim().replace(/[^\d]/g, '');
    if (!nomor.startsWith('62')) nomor = '62' + nomor.replace(/^0/, '');
    if (contacts[nomor]) {
      delete contacts[nomor];
      saveContacts();
      return sendMessage(sender, `✅ Kontak ${nomor} dihapus.`);
    }
    return sendMessage(sender, `❌ Kontak ${nomor} tidak ditemukan.`);
  }

  // Reload
  if (lower === 'reload') {
    loadExcel();
    return sendMessage(sender, '✅ Data Excel berhasil di-reload!\n\nSheets: ' + Object.keys(excelData).join(', '));
  }

  // Info
  if (lower === 'info') {
    const uptime = process.uptime();
    const hours = Math.floor(uptime / 3600);
    const mins = Math.floor((uptime % 3600) / 60);
    let msg = '🔧 *INFO SISTEM*\n\n';
    msg += `⏱️ Uptime: ${hours}h ${mins}m\n`;
    msg += `👥 Member: ${Object.keys(members).length}\n`;
    msg += `📇 Kontak: ${Object.keys(contacts).length}\n`;
    msg += `📊 Sesi Aktif: ${Object.keys(sessions).length}\n`;
    msg += `📁 Excel Sheets: ${Object.keys(excelData).length}\n`;
    for (const key of Object.keys(excelData)) {
      msg += `   - ${key}: ${excelData[key].length} items\n`;
    }
    msg += `\n🤖 AI Status:\n`;
    msg += `   Groq: ${GROQ_API_KEY ? '✅' : '❌'}\n`;
    msg += `   Gemini Keys: ${GEMINI_KEYS.length}\n`;
    msg += `   OpenRouter: ${OPENROUTER_API_KEY ? '✅' : '❌'}\n`;
    msg += `\n📡 Fonnte: ${FONNTE_TOKEN ? '✅' : '❌'}`;
    return sendMessage(sender, msg);
  }

  // Reset all
  if (lower === 'resetall') {
    resetAllSessions();
    return sendMessage(sender, '✅ Semua sesi telah direset.');
  }

  // Default admin menu
  return sendMessage(sender, getAdminMenuText());
}

// ============ WEBHOOK ENDPOINT ============
app.post('/webhook', async (req, res) => {
  try {
    const { sender, message } = req.body;

    if (!sender || !message) {
      return res.status(200).json({ status: 'no data' });
    }

    // Normalize sender number
    let senderNum = sender.replace(/[^\d]/g, '');
    if (!senderNum.startsWith('62')) {
      senderNum = '62' + senderNum.replace(/^0/, '');
    }

    // Skip group messages
    if (sender.includes('-') || sender.includes('@g.us')) {
      return res.status(200).json({ status: 'group ignored' });
    }

    logToFile(`Webhook received from ${senderNum}: ${message}`);

    // Process async
    handleMessage(senderNum, message).catch(err => {
      logToFile(`Error handling message: ${err.message}`);
    });

    return res.status(200).json({ status: 'ok' });
  } catch (err) {
    logToFile(`Webhook error: ${err.message}`);
    return res.status(200).json({ status: 'error' });
  }
});

// Fonnte sometimes sends to /
app.post('/', async (req, res) => {
  try {
    const { sender, message } = req.body;

    if (!sender || !message) {
      return res.status(200).json({ status: 'alive', message: 'Bot WhatsApp Perabot aktif! 🤖' });
    }

    let senderNum = sender.replace(/[^\d]/g, '');
    if (!senderNum.startsWith('62')) {
      senderNum = '62' + senderNum.replace(/^0/, '');
    }

    if (sender.includes('-') || sender.includes('@g.us')) {
      return res.status(200).json({ status: 'group ignored' });
    }

    logToFile(`Root webhook received from ${senderNum}: ${message}`);

    handleMessage(senderNum, message).catch(err => {
      logToFile(`Error handling message: ${err.message}`);
    });

    return res.status(200).json({ status: 'ok' });
  } catch (err) {
    logToFile(`Root webhook error: ${err.message}`);
    return res.status(200).json({ status: 'error' });
  }
});

// ============ API ENDPOINTS ============

app.get('/', (req, res) => {
  const uptime = process.uptime();
  const hours = Math.floor(uptime / 3600);
  const mins = Math.floor((uptime % 3600) / 60);
  res.json({
    status: 'alive',
    message: 'Bot WhatsApp Perabot aktif! 🤖',
    uptime: `${hours}h ${mins}m`,
    sessions: Object.keys(sessions).length,
    members: Object.keys(members).length,
    excelSheets: Object.keys(excelData).length
  });
});

app.get('/reload', (req, res) => {
  loadExcel();
  res.json({
    status: 'ok',
    message: 'Excel reloaded',
    sheets: Object.keys(excelData),
    counts: Object.fromEntries(Object.entries(excelData).map(([k, v]) => [k, v.length]))
  });
});

app.get('/resetsesi/:nomor', (req, res) => {
  let nomor = req.params.nomor.replace(/[^\d]/g, '');
  if (!nomor.startsWith('62')) nomor = '62' + nomor.replace(/^0/, '');
  if (sessions[nomor]) {
    deleteSession(nomor);
    res.json({ status: 'ok', message: `Sesi ${nomor} direset.` });
  } else {
    res.json({ status: 'not_found', message: `Sesi ${nomor} tidak ditemukan.` });
  }
});

app.get('/resetall', (req, res) => {
  resetAllSessions();
  res.json({ status: 'ok', message: 'Semua sesi direset.' });
});

app.get('/sessions', (req, res) => {
  res.json({
    count: Object.keys(sessions).length,
    sessions: Object.fromEntries(
      Object.entries(sessions).map(([k, v]) => [k, { state: v.state, lang: v.lang, lastActive: new Date(v.lastActive).toISOString() }])
    )
  });
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ============ STARTUP ============
function startup() {
  logToFile('========================================');
  logToFile('Bot WhatsApp Perabot starting...');
  logToFile('========================================');

  loadSessions();
  loadMembers();
  loadContacts();
  loadExcel();

  app.listen(PORT, '0.0.0.0', () => {
    logToFile(`Server running on port ${PORT}`);
    logToFile(`Admin: ${ADMIN_NUMBER}`);
    logToFile(`Role Laporan: ${ROLE_LAPORAN.join(', ')}`);
    logToFile(`Fonnte: ${FONNTE_TOKEN ? 'Configured' : 'NOT SET'}`);
    logToFile(`Groq: ${GROQ_API_KEY ? 'Configured' : 'NOT SET'}`);
    logToFile(`Gemini Keys: ${GEMINI_KEYS.length}`);
    logToFile(`OpenRouter: ${OPENROUTER_API_KEY ? 'Configured' : 'NOT SET'}`);
    logToFile('Bot ready!');
  });
}

startup();
