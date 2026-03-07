require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');

// === ENV VARIABLES ===
const TOKEN = process.env.TELEGRAM_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
let GOOGLE_SERVICE_ACCOUNT_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

// Validasi
if (!TOKEN) {
  console.error('❌ TELEGRAM_TOKEN not set');
  process.exit(1);
}
if (!SHEET_ID) {
  console.error('❌ SHEET_ID not set');
  process.exit(1);
}
if (!GOOGLE_SERVICE_ACCOUNT_JSON) {
  console.error('❌ GOOGLE_SERVICE_ACCOUNT_JSON not set');
  process.exit(1);
}

// === PARSE GOOGLE SERVICE ACCOUNT ===
let serviceAccount;
try {
  let keyData = GOOGLE_SERVICE_ACCOUNT_JSON.trim();
  
  // Jika base64, decode dulu
  if (!keyData.startsWith('{')) {
    try {
      keyData = Buffer.from(keyData, 'base64').toString('utf-8');
    } catch (e) {
      // bukan base64
    }
  }
  
  serviceAccount = JSON.parse(keyData);
  console.log('✅ Google Service Account parsed');
} catch (e) {
  console.error('❌ Failed to parse JSON:', e.message);
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

const ASSURANCE_SHEET = 'PROGRES ASSURANCE';
const ORDER_ASSURANCE_SHEET = 'ORDER ASSURANCE';
const MASTER_SHEET = 'MASTER';

// === CACHING LAYER ===
const cache = {
  masterData: null,
  masterDataTime: 0,
  assuranceData: null,
  assuranceDataTime: 0,
  orderAssuranceData: null,
  orderAssuranceDataTime: 0,
  cacheExpiry: 5 * 60 * 1000, // 5 menit
};

// === HELPER: Get sheet data with caching ===
async function getSheetData(sheetName, useCache = true) {
  try {
    // Cek cache untuk MASTER_SHEET
    if (useCache && sheetName === MASTER_SHEET && cache.masterData) {
      if (Date.now() - cache.masterDataTime < cache.cacheExpiry) {
        console.log('📦 Using cached MASTER_SHEET');
        return cache.masterData;
      }
    }

    // Cek cache untuk ASSURANCE_SHEET
    if (useCache && sheetName === ASSURANCE_SHEET && cache.assuranceData) {
      if (Date.now() - cache.assuranceDataTime < cache.cacheExpiry) {
        console.log('📦 Using cached ASSURANCE_SHEET');
        return cache.assuranceData;
      }
    }

    // Cek cache untuk ORDER_ASSURANCE_SHEET
    if (useCache && sheetName === ORDER_ASSURANCE_SHEET && cache.orderAssuranceData) {
      if (Date.now() - cache.orderAssuranceDataTime < cache.cacheExpiry) {
        console.log('📦 Using cached ORDER_ASSURANCE_SHEET');
        return cache.orderAssuranceData;
      }
    }

    // Fetch dari API
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 15000); // 15 detik timeout

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: sheetName,
    });

    clearTimeout(timeout);
    const data = res.data.values || [];

    // Cache hasil
    if (sheetName === MASTER_SHEET) {
      cache.masterData = data;
      cache.masterDataTime = Date.now();
    } else if (sheetName === ASSURANCE_SHEET) {
      cache.assuranceData = data;
      cache.assuranceDataTime = Date.now();
    } else if (sheetName === ORDER_ASSURANCE_SHEET) {
      cache.orderAssuranceData = data;
      cache.orderAssuranceDataTime = Date.now();
    }

    return data;
  } catch (error) {
    console.error(`Error reading ${sheetName}:`, error.message);
    // Return cache meskipun expired jika API error
    if (sheetName === MASTER_SHEET && cache.masterData) {
      console.log('⚠️ API error, fallback ke cached MASTER_SHEET');
      return cache.masterData;
    }
    if (sheetName === ASSURANCE_SHEET && cache.assuranceData) {
      console.log('⚠️ API error, fallback ke cached ASSURANCE_SHEET');
      return cache.assuranceData;
    }
    if (sheetName === ORDER_ASSURANCE_SHEET && cache.orderAssuranceData) {
      console.log('⚠️ API error, fallback ke cached ORDER_ASSURANCE_SHEET');
      return cache.orderAssuranceData;
    }
    throw error;
  }
}

// === HELPER: Append to sheet ===
async function appendSheetData(sheetName, values) {
  try {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: sheetName,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [values] },
    });
  } catch (error) {
    console.error(`Error writing to ${sheetName}:`, error.message);
    throw error;
  }
}

// === HELPER: Send Telegram ===
async function sendTelegram(chatId, text, options = {}) {
  const maxLength = 4000;
  try {
    if (text.length <= maxLength) {
      return await bot.sendMessage(chatId, text, { parse_mode: 'HTML', ...options });
    } else {
      const lines = text.split('\n');
      let chunk = '';
      for (let i = 0; i < lines.length; i++) {
        if ((chunk + lines[i] + '\n').length > maxLength) {
          await bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options });
          chunk = '';
        }
        chunk += lines[i] + '\n';
      }
      if (chunk.trim()) {
        await bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options });
      }
    }
  } catch (error) {
    console.error('Error sending message:', error.message);
  }
}

// === HELPER: Wrap with timeout ===
function withTimeout(promise, ms = 10000) {
  return Promise.race([
    promise,
    new Promise((_, reject) => 
      setTimeout(() => reject(new Error('Timeout - Google API response too slow')), ms)
    )
  ]);
}

// === HELPER: Get user role ===
async function getUserRole(username) {
  try {
    const data = await getSheetData(MASTER_SHEET);
    for (let i = 1; i < data.length; i++) {
      const sheetUser = (data[i][8] || '').replace('@', '').toLowerCase().trim();
      const inputUser = (username || '').replace('@', '').toLowerCase().trim();
      const status = (data[i][10] || '').toUpperCase().trim();
      const role = (data[i][9] || '').toUpperCase().trim();
      
      if (sheetUser === inputUser && status === 'AKTIF') {
        return role;
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting user role:', error.message);
    return null;
  }
}

// === HELPER: Check authorization ===
async function checkAuthorization(username, requiredRoles = []) {
  try {
    const userRole = await withTimeout(getUserRole(username), 8000);
    if (!userRole) {
      return { authorized: false, role: null, message: '❌ Anda tidak terdaftar di sistem.' };
    }
    
    if (requiredRoles.length > 0 && !requiredRoles.includes(userRole)) {
      return { authorized: false, role: userRole, message: `❌ Akses ditolak. Role ${userRole} tidak memiliki izin untuk command ini.` };
    }
    
    return { authorized: true, role: userRole };
  } catch (error) {
    console.error('Authorization error:', error.message);
    return { authorized: false, role: null, message: '❌ Terjadi kesalahan saat verifikasi. Server sedang sibuk.' };
  }
}

// === HELPER: Parse assurance data ===
function parseAssurance(text, username) {
  let data = {
    incidentNo: '',
    closeDesc: '',
    dropcore: '',
    patchcord: '',
    soc: '',
    pslave: '',
    passive1_8: '',
    passive1_4: '',
    pigtail: '',
    adaptor: '',
    roset: '',
    rj45: '',
    lan: '',
    dateCreated: new Date().toLocaleDateString('id-ID', {
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric',
      timeZone: 'Asia/Jakarta',
    }),
    teknisi: (username || '').replace('@', ''),
  };

  // Parse INCIDENT (pattern: INC47052822)
  const incidentMatch = text.match(/INC[0-9]+/i);
  if (incidentMatch) {
    data.incidentNo = incidentMatch[0].trim().toUpperCase();
  }

  // Parse CLOSE (dari pattern CLOSE: deskripsi)
  const closeMatch = text.match(/CLOSE\s*:\s*(.+?)(?=\n|MATERIAL|$)/i);
  if (closeMatch && closeMatch[1]) {
    data.closeDesc = closeMatch[1].trim();
  }

  // Parse MATERIAL (key: value)
  const materialPatterns = {
    dropcore: /DROPCORE\s*:\s*([0-9\.]+)/i,
    patchcord: /PATCHCORD\s*:\s*([0-9\.]+)/i,
    soc: /SOC\s*:\s*([0-9\.]+)/i,
    pslave: /PSLAVE\s*:\s*([0-9\.]+)/i,
    passive1_8: /PASSIVE\s*1\/8\s*:\s*([0-9\.]+)/i,
    passive1_4: /PASSIVE\s*1\/4\s*:\s*([0-9\.]+)/i,
    pigtail: /PIGTAIL\s*:\s*([0-9\.]+)/i,
    adaptor: /ADAPTOR\s*:\s*([0-9\.]+)/i,
    roset: /ROSET\s*:\s*([0-9\.]+)/i,
    rj45: /RJ\s*45\s*:\s*([0-9\.]+)/i,
    lan: /LAN\s*:\s*([0-9\.]+)/i,
  };

  for (const [key, pattern] of Object.entries(materialPatterns)) {
    const match = text.match(pattern);
    if (match && match[1]) {
      data[key] = match[1].trim();
    }
  }

  return data;
}

// === BOT SETUP ===
const PORT = process.env.PORT || 3002;
const RAILWAY_STATIC_URL = process.env.RAILWAY_STATIC_URL;
const USE_WEBHOOK = !!RAILWAY_STATIC_URL;

let bot;

if (USE_WEBHOOK) {
  const express = require('express');
  const app = express();
  app.use(express.json());

  bot = new TelegramBot(TOKEN);
  const webhookUrl = `https://${RAILWAY_STATIC_URL}/assurance${TOKEN}`;

  bot.setWebHook(webhookUrl).then(() => {
    console.log(`✅ Webhook set: ${webhookUrl}`);
  }).catch(err => {
    console.error('❌ Webhook error:', err.message);
  });

  app.post(`/assurance${TOKEN}`, (req, res) => {
    bot.processUpdate(req.body);
    res.sendStatus(200);
  });

  app.get('/', (req, res) => {
    res.send('Bot Assurance is running!');
  });

  app.listen(PORT, () => {
    console.log(`✅ Server running on port ${PORT}`);
  });
} else {
  // Polling mode
  bot = new TelegramBot(TOKEN, { 
    polling: {
      interval: 300,
      autoStart: true,
      params: {
        timeout: 10,
        allowed_updates: ['message']
      }
    }
  });
  console.log('✅ Bot running in polling mode');
  
  bot.on('polling_error', (error) => {
    if (error.code === 'EFATAL') {
      console.error('❌ Polling fatal error:', error.message);
    } else {
      console.error('⚠️ Polling error:', error.message);
    }
  });
}

// === MESSAGE HANDLER ===
bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const msgId = msg.message_id;
  const text = (msg.text || '').trim();
  const username = msg.from.username || '';
  const groupName = msg.chat.title || msg.chat.first_name || `ID:${chatId}`;
  const groupType = msg.chat.type;

  // Early return untuk pesan kosong atau non-text
  if (!text) {
    return;
  }

  // Early return untuk pesan yang bukan command
  if (!text.startsWith('/')) {
    return;
  }

  console.log(`📨 [${groupType}] ${groupName} | [@${username}] ${text.substring(0, 60)}`);

  try {
    // === /INPUT ===
    if (/^\/INPUT\b/i.test(text)) {
      try {
        const auth = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!auth.authorized) {
          return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
        }

        const inputText = text.replace(/^\/INPUT\s*/i, '').trim();
        if (!inputText) {
          return sendTelegram(chatId, '❌ Silakan kirim data assurance setelah /INPUT.', { reply_to_message_id: msgId });
        }

        const parsed = parseAssurance(inputText, username);

        // Validasi required fields
        const required = ['incidentNo', 'closeDesc'];
        const missing = required.filter(f => !parsed[f]);

        if (missing.length > 0) {
          return sendTelegram(chatId, `❌ Field wajib: ${missing.join(', ')}`, { reply_to_message_id: msgId });
        }

        // Row untuk sheet PROGRES ASSURANCE
        const row = [
          parsed.dateCreated,      // A: TANGGAL
          parsed.incidentNo,       // B: INCIDENT
          parsed.dropcore,         // C: DROPCORE
          parsed.patchcord,        // D: PATCHCORD
          parsed.soc,              // E: SOC
          parsed.pslave,           // F: PSLAVE
          parsed.passive1_8,       // G: PASSIVE 1/8
          parsed.passive1_4,       // H: PASSIVE 1/4
          parsed.pigtail,          // I: PIGTAIL
          parsed.adaptor,          // J: ADAPTOR
          parsed.roset,            // K: ROSET
          parsed.rj45,             // L: RJ 45
          parsed.lan,              // M: LAN
          parsed.closeDesc,        // N: CLOSE
          parsed.teknisi,          // O: TEKNISI
        ];

        await withTimeout(appendSheetData(ASSURANCE_SHEET, row), 10000);

        let confirmMsg = `✅ Data Assurance berhasil disimpan!\n\n`;
        confirmMsg += `<b>Incident:</b> ${parsed.incidentNo}\n`;
        confirmMsg += `<b>Close:</b> ${parsed.closeDesc}\n`;
        confirmMsg += `<b>Material:</b>\n`;
        confirmMsg += `  • Dropcore: ${parsed.dropcore || '-'}\n`;
        confirmMsg += `  • Patchcord: ${parsed.patchcord || '-'}\n`;
        confirmMsg += `  • SOC: ${parsed.soc || '-'}\n`;
        confirmMsg += `  • PSLAVE: ${parsed.pslave || '-'}\n`;
        confirmMsg += `  • PASSIVE 1/8: ${parsed.passive1_8 || '-'}\n`;
        confirmMsg += `  • PASSIVE 1/4: ${parsed.passive1_4 || '-'}\n`;
        confirmMsg += `  • Pigtail: ${parsed.pigtail || '-'}\n`;
        confirmMsg += `  • Adaptor: ${parsed.adaptor || '-'}\n`;
        confirmMsg += `  • Roset: ${parsed.roset || '-'}\n`;
        confirmMsg += `  • RJ 45: ${parsed.rj45 || '-'}\n`;
        confirmMsg += `  • LAN: ${parsed.lan || '-'}`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (inputErr) {
        console.error('❌ /INPUT Error:', inputErr.message);
        return sendTelegram(chatId, `❌ Error: ${inputErr.message}. Silakan coba lagi.`, { reply_to_message_id: msgId });
      }
    }

    // === /hari_ini [TEKNISI] ===
    else if (/^\/hari_ini\b/i.test(text)) {
      try {
        const auth = await checkAuthorization(username, ['ADMIN']);
        if (!auth.authorized) {
          return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
        }

        const args = text.replace(/^\/hari_ini\s*/i, '').trim();
        if (!args) {
          return sendTelegram(chatId, '❌ Format: /hari_ini TEKNISI_USERNAME', { reply_to_message_id: msgId });
        }

        const today = new Date().toLocaleDateString('id-ID', {
          day: 'numeric',
          month: 'long',
          year: 'numeric',
          timeZone: 'Asia/Jakarta',
        });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        let incidents = [];

        for (let i = 1; i < data.length; i++) {
          const dateCreated = (data[i][0] || '').trim();
          const teknisi = (data[i][14] || '-').trim();
          
          if (dateCreated === today && teknisi === args) {
            incidents.push({
              incident: data[i][1] || '-',
              closeDesc: data[i][13] || '-',
            });
          }
        }

        let msg = `📋 <b>INCIDENTS HARI INI - ${args}</b>\n<b>${today}</b>\n\n`;
        msg += `<b>Total: ${incidents.length} Tickets</b>\n\n`;
        
        if (incidents.length === 0) {
          msg += '<i>Belum ada incident hari ini</i>';
        } else {
          incidents.forEach((inc, idx) => {
            msg += `${idx + 1}. <b>${inc.incident}</b>\n   ${inc.closeDesc}\n\n`;
          });
        }

        return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /hari_ini Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // === /summary (HARI INI) ===
    else if (/^\/summary\b/i.test(text)) {
      try {
        const auth = await checkAuthorization(username, ['ADMIN']);
        if (!auth.authorized) {
          return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
        }

        const today = new Date().toLocaleDateString('id-ID', {
          day: 'numeric',
          month: 'long',
          year: 'numeric',
          timeZone: 'Asia/Jakarta',
        });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        let map = {};

        for (let i = 1; i < data.length; i++) {
          const dateCreated = (data[i][0] || '').trim();
          if (dateCreated !== today) continue;
          
          const teknisi = (data[i][14] || '-').trim();
          
          if (!map[teknisi]) {
            map[teknisi] = 0;
          }
          map[teknisi]++;
        }

        const entries = Object.entries(map)
          .sort((a, b) => b[1] - a[1]);

        let totalTickets = Object.values(map).reduce((sum, count) => sum + count, 0);
        let msg = `📊 <b>SUMMARY TEAM - ${today}</b>\n\n`;
        msg += `<b>Total Tickets: ${totalTickets}</b>\n\n`;
        
        if (entries.length === 0) {
          msg += '<i>Belum ada data hari ini</i>';
        } else {
          entries.forEach((entry) => {
            const [teknisi, count] = entry;
            msg += `🔸 <b>${teknisi}</b>: ${count} tickets\n`;
          });
        }

        return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /summary Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // === /all_summary (KESELURUHAN) ===
    else if (/^\/all_summary\b/i.test(text)) {
      try {
        const auth = await checkAuthorization(username, ['ADMIN']);
        if (!auth.authorized) {
          return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
        }

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        let map = {};

        for (let i = 1; i < data.length; i++) {
          const teknisi = (data[i][14] || '-').trim();
          
          if (!map[teknisi]) {
            map[teknisi] = 0;
          }
          map[teknisi]++;
        }

        const entries = Object.entries(map)
          .sort((a, b) => b[1] - a[1]);

        let totalTickets = Object.values(map).reduce((sum, count) => sum + count, 0);
        let msg = `📊 <b>SUMMARY TEAM - KESELURUHAN</b>\n\n`;
        msg += `<b>Total Tickets: ${totalTickets}</b>\n\n`;
        
        if (entries.length === 0) {
          msg += '<i>Belum ada data</i>';
        } else {
          entries.forEach((entry) => {
            const [teknisi, count] = entry;
            msg += `🔸 <b>${teknisi}</b>: ${count} tickets\n`;
          });
        }

        return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /all_summary Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // === /material_used ===
    else if (/^\/material_used\b/i.test(text)) {
      try {
        const auth = await checkAuthorization(username, ['ADMIN']);
        if (!auth.authorized) {
          return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
        }

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        let materialMap = {
          'DROPCORE': 0,
          'PATCHCORD': 0,
          'SOC': 0,
          'PSLAVE': 0,
          'PASSIVE 1/8': 0,
          'PASSIVE 1/4': 0,
          'PIGTAIL': 0,
          'ADAPTOR': 0,
          'ROSET': 0,
          'RJ 45': 0,
          'LAN': 0,
        };

        // Column indices: C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11, M=12
        const materialColumns = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
        const materialNames = Object.keys(materialMap);

        for (let i = 1; i < data.length; i++) {
          materialColumns.forEach((colIdx, idx) => {
            const value = parseInt((data[i][colIdx] || '0').trim()) || 0;
            materialMap[materialNames[idx]] += value;
          });
        }

        const entries = Object.entries(materialMap)
          .filter(([_, count]) => count > 0)
          .sort((a, b) => b[1] - a[1]);

        let msg = `📦 <b>PENGGUNAAN MATERIAL - KESELURUHAN</b>\n\n`;
        
        if (entries.length === 0) {
          msg += '<i>Belum ada material yang dipakai</i>';
        } else {
          entries.forEach((entry) => {
            const [material, count] = entry;
            msg += `• <b>${material}</b>: ${count} unit\n`;
          });
        }

        return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /material_used Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // === /laporan_detail [TEKNISI] ===
    else if (/^\/laporan_detail\b/i.test(text)) {
      try {
        const auth = await checkAuthorization(username, ['ADMIN']);
        if (!auth.authorized) {
          return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
        }

        const args = text.replace(/^\/laporan_detail\s*/i, '').trim();
        if (!args) {
          return sendTelegram(chatId, '❌ Format: /laporan_detail TEKNISI_USERNAME', { reply_to_message_id: msgId });
        }

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        let incidents = [];
        let totalMaterial = {
          'DROPCORE': 0,
          'PATCHCORD': 0,
          'SOC': 0,
          'PSLAVE': 0,
          'PASSIVE 1/8': 0,
          'PASSIVE 1/4': 0,
          'PIGTAIL': 0,
          'ADAPTOR': 0,
          'ROSET': 0,
          'RJ 45': 0,
          'LAN': 0,
        };

        for (let i = 1; i < data.length; i++) {
          const teknisi = (data[i][14] || '-').trim();
          
          if (teknisi === args) {
            incidents.push({
              tanggal: data[i][0] || '-',
              incident: data[i][1] || '-',
              close: data[i][13] || '-',
            });

            // Hitung material
            const materialColumns = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
            const materialNames = Object.keys(totalMaterial);
            materialColumns.forEach((colIdx, idx) => {
              const value = parseInt((data[i][colIdx] || '0').trim()) || 0;
              totalMaterial[materialNames[idx]] += value;
            });
          }
        }

        let msg = `📋 <b>LAPORAN DETAIL - ${args}</b>\n\n`;
        msg += `<b>Total Incidents: ${incidents.length}</b>\n`;
        msg += `<b>Material Yang Dipakai:</b>\n`;
        
        Object.entries(totalMaterial)
          .filter(([_, count]) => count > 0)
          .forEach((entry) => {
            const [material, count] = entry;
            msg += `  • ${material}: ${count}\n`;
          });

        msg += `\n<b>Daftar Incidents:</b>\n`;
        
        if (incidents.length === 0) {
          msg += '<i>Belum ada incidents</i>';
        } else {
          incidents.forEach((inc, idx) => {
            msg += `\n${idx + 1}. <b>${inc.incident}</b>\n`;
            msg += `   Tanggal: ${inc.tanggal}\n`;
            msg += `   Close: ${inc.close}\n`;
          });
        }

        return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /laporan_detail Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // === /sisa_ticket ===
    else if (/^\/sisa_ticket\b/i.test(text)) {
      try {
        const auth = await checkAuthorization(username, ['ADMIN']);
        if (!auth.authorized) {
          return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
        }

        const data = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET), 10000);
        
        // Get current date/time
        const now = new Date();
        const dayNameID = ['MINGGU', 'SENIN', 'SELASA', 'RABU', 'KAMIS', 'JUMAT', 'SABTU'][now.getDay()];
        const dateStr = now.toLocaleDateString('id-ID', { 
          day: '2-digit', 
          month: 'long', 
          year: 'numeric',
          timeZone: 'Asia/Jakarta'
        }).toUpperCase();
        const timeStr = now.toLocaleTimeString('id-ID', { 
          hour: '2-digit', 
          minute: '2-digit',
          timeZone: 'Asia/Jakarta'
        });

        // Group tickets by TEKNISI
        let ticketsByTeknisi = {};
        let totalOpenTickets = 0;

        for (let i = 1; i < data.length; i++) {
          const incident = (data[i][1] || '-').trim();              // Column B: INCIDENT
          const teknisi = (data[i][2] || '-').trim();               // Column C: TEKNISI
          const ttrCustomer = (data[i][3] || '-').trim();           // Column D: TTR CUSTOMER
          const status = (data[i][9] || 'OPEN').trim().toUpperCase();  // Column J: STATUS

          // Only include tickets that are NOT CLOSED
          if (status !== 'CLOSED') {
            if (!ticketsByTeknisi[teknisi]) {
              ticketsByTeknisi[teknisi] = [];
            }
            ticketsByTeknisi[teknisi].push({
              incident,
              ttr: ttrCustomer,
            });
            totalOpenTickets++;
          }
        }

        // Sort teknisi alphabetically
        const sortedTeknisi = Object.keys(ticketsByTeknisi).sort();

        let msg = `🔴 <b>SISA TICKET OPEN ${dayNameID}, ${dateStr} ${timeStr}</b>\n\n`;
        msg += `<b>Total Open: ${totalOpenTickets} tickets</b>\n\n`;
        
        if (sortedTeknisi.length === 0) {
          msg += '<i>Tidak ada ticket yang masih open</i>';
        } else {
          sortedTeknisi.forEach((teknisiName, idx) => {
            const tickets = ticketsByTeknisi[teknisiName];
            msg += `${idx + 1}. <b>@${teknisiName}</b>\n`;
            
            // Display incidents with TTR
            tickets.forEach((ticket) => {
              msg += `   ${ticket.incident}   ${ticket.ttr}\n`;
            });
            msg += '\n';
          });
        }

        return sendTelegram(chatId, msg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /sisa_ticket Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // === /help ===
    else if (/^\/(help|start)\b/i.test(text)) {
      try {
        const auth = await checkAuthorization(username);
        if (!auth.authorized) {
          return sendTelegram(chatId, auth.message, { reply_to_message_id: msgId });
        }

        const helpMsg = `🤖 Bot Assurance

<b>📝 INPUT COMMAND:</b>
/INPUT - Input data assurance

<b>📊 MONITORING COMMANDS (ADMIN):</b>
/hari_ini TEKNISI - Incidents teknisi hari ini
/summary - Total team HARI INI
/all_summary - Total team KESELURUHAN
/laporan_detail TEKNISI - Laporan detail teknisi
/material_used - Material yang dipakai
/sisa_ticket - Ticket yang masih OPEN

<b>📋 FORMAT /INPUT:</b>
/INPUT: INC47052822
CLOSE: deskripsi perbaikan
MATERIAL:
DROPCORE: 0
PATCHCORD: 0
SOC: 0
PSLAVE: 2
PASSIVE 1/8: 0
PASSIVE 1/4: 0
PIGTAIL: 0
ADAPTOR: 0
ROSET: 0
RJ 45: 0
LAN: 0`;

        return sendTelegram(chatId, helpMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /help Error:', err.message);
        return sendTelegram(chatId, '❌ Terjadi kesalahan.', { reply_to_message_id: msgId });
      }
    }

  } catch (err) {
    console.error('Error:', err.message);
    sendTelegram(chatId, '❌ Terjadi kesalahan sistem.', { reply_to_message_id: msgId });
  }
});

process.on('unhandledRejection', (reason) => {
  console.error('Error:', reason);
});

console.log('\n🚀 Bot Assurance started!');
console.log(`Mode: ${USE_WEBHOOK ? 'Webhook' : 'Polling'}`);
console.log('═'.repeat(50));
console.log('✅ Auto-Cache Enabled (5 min expiry)');
console.log('✅ Timeout Protection Enabled');
console.log('✅ Error Fallback Enabled');
console.log('═'.repeat(50));
