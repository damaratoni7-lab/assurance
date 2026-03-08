require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');

// === ENV VARIABLES ===
const TOKEN = process.env.TELEGRAM_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const GROUP_CHAT_ID = process.env.GROUP_CHAT_ID;
let GOOGLE_SERVICE_ACCOUNT_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

if (!TOKEN) { console.error('❌ TELEGRAM_TOKEN not set'); process.exit(1); }
if (!SHEET_ID) { console.error('❌ SHEET_ID not set'); process.exit(1); }
if (!GOOGLE_SERVICE_ACCOUNT_JSON) { console.error('❌ GOOGLE_SERVICE_ACCOUNT_JSON not set'); process.exit(1); }
if (!GROUP_CHAT_ID) { console.warn('⚠️ GROUP_CHAT_ID not set - TTR alerts disabled'); }

// === PARSE GOOGLE SERVICE ACCOUNT ===
let serviceAccount;
try {
  let keyData = GOOGLE_SERVICE_ACCOUNT_JSON.trim();
  if (!keyData.startsWith('{')) {
    try { keyData = Buffer.from(keyData, 'base64').toString('utf-8'); } catch (e) { }
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

// === CONSTANTS ===
const ASSURANCE_SHEET = 'PROGRES ASSURANCE';
const ORDER_ASSURANCE_SHEET = 'ORDER ASSURANCE';
const MASTER_SHEET = 'MASTER';

const TTR_TABLE = {
  'HVC_DIAMOND': 3,
  'FFG': 3,
  'BGES HSI': 4,
  'DATIN K2': 3.6,
  'DATIN K3': 7.2,
  'HVC_PLATINUM': 6,
  'HVC_GOLD': 12,
  'REGULER': 36,
};

const BULAN_ID = {
  'januari': 1, 'februari': 2, 'maret': 3, 'april': 4,
  'mei': 5, 'juni': 6, 'juli': 7, 'agustus': 8,
  'september': 9, 'oktober': 10, 'november': 11, 'desember': 12,
};

// === CACHING ===
const cache = {
  masterData: null, masterDataTime: 0,
  assuranceData: null, assuranceDataTime: 0,
  orderAssuranceData: null, orderAssuranceDataTime: 0,
  cacheExpiry: 5 * 60 * 1000,
};

// === STATE ===
const alertState = { warned: new Set(), expired: new Set() };
const userChatIds = {};

// === HELPER: Get sheet data with caching ===
async function getSheetData(sheetName, useCache = true) {
  try {
    if (useCache) {
      if (sheetName === MASTER_SHEET && cache.masterData && Date.now() - cache.masterDataTime < cache.cacheExpiry) {
        return cache.masterData;
      }
      if (sheetName === ASSURANCE_SHEET && cache.assuranceData && Date.now() - cache.assuranceDataTime < cache.cacheExpiry) {
        return cache.assuranceData;
      }
      if (sheetName === ORDER_ASSURANCE_SHEET && cache.orderAssuranceData && Date.now() - cache.orderAssuranceDataTime < cache.cacheExpiry) {
        return cache.orderAssuranceData;
      }
    }

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: sheetName,
    });
    const data = res.data.values || [];

    if (sheetName === MASTER_SHEET) { cache.masterData = data; cache.masterDataTime = Date.now(); }
    else if (sheetName === ASSURANCE_SHEET) { cache.assuranceData = data; cache.assuranceDataTime = Date.now(); }
    else if (sheetName === ORDER_ASSURANCE_SHEET) { cache.orderAssuranceData = data; cache.orderAssuranceDataTime = Date.now(); }

    return data;
  } catch (error) {
    console.error(`Error reading ${sheetName}:`, error.message);
    if (sheetName === MASTER_SHEET && cache.masterData) return cache.masterData;
    if (sheetName === ASSURANCE_SHEET && cache.assuranceData) return cache.assuranceData;
    if (sheetName === ORDER_ASSURANCE_SHEET && cache.orderAssuranceData) return cache.orderAssuranceData;
    throw error;
  }
}

// === HELPER: Append to sheet ===
async function appendSheetData(sheetName, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: sheetName,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [values] },
  });
}

// === HELPER: Update a single cell ===
async function updateSheetCell(sheetName, cell, value) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `'${sheetName}'!${cell}`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [[value]] },
  });
}

// === HELPER: Send Telegram (with chunking) ===
async function sendTelegram(chatId, text, options = {}) {
  const maxLength = 4000;
  try {
    if (text.length <= maxLength) {
      return await bot.sendMessage(chatId, text, { parse_mode: 'HTML', ...options });
    }
    const lines = text.split('\n');
    let chunk = '';
    for (const line of lines) {
      if ((chunk + line + '\n').length > maxLength) {
        await bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options });
        chunk = '';
      }
      chunk += line + '\n';
    }
    if (chunk.trim()) {
      await bot.sendMessage(chatId, chunk, { parse_mode: 'HTML', ...options });
    }
  } catch (error) {
    console.error('Error sending message:', error.message);
  }
}

// === HELPER: Timeout wrapper ===
function withTimeout(promise, ms = 10000) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout - Google API too slow')), ms)),
  ]);
}

// === HELPER: Get user role from MASTER ===
async function getUserRole(username) {
  try {
    const data = await getSheetData(MASTER_SHEET);
    for (let i = 1; i < data.length; i++) {
      const sheetUser = (data[i][8] || '').replace('@', '').toLowerCase().trim();
      const inputUser = (username || '').replace('@', '').toLowerCase().trim();
      const status = (data[i][10] || '').toUpperCase().trim();
      const role = (data[i][9] || '').toUpperCase().trim();
      if (sheetUser === inputUser && status === 'AKTIF') return role;
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
    if (!userRole) return { authorized: false, role: null, message: '❌ Anda tidak terdaftar di sistem.' };
    if (requiredRoles.length > 0 && !requiredRoles.includes(userRole))
      return { authorized: false, role: userRole, message: `❌ Akses ditolak. Role ${userRole} tidak memiliki izin.` };
    return { authorized: true, role: userRole };
  } catch (error) {
    return { authorized: false, role: null, message: '❌ Terjadi kesalahan saat verifikasi.' };
  }
}

// === HELPER: Get active admins from MASTER ===
async function getActiveAdmins() {
  try {
    const data = await getSheetData(MASTER_SHEET);
    const admins = [];
    for (let i = 1; i < data.length; i++) {
      const uname = (data[i][8] || '').replace('@', '').trim();
      const role = (data[i][9] || '').toUpperCase().trim();
      const status = (data[i][10] || '').toUpperCase().trim();
      if (role === 'ADMIN' && status === 'AKTIF' && uname) admins.push(uname);
    }
    return admins;
  } catch (error) {
    console.error('Error getting admins:', error.message);
    return [];
  }
}

// === HELPER: Get workzone mappings from ORDER ASSURANCE cols Q & R ===
function getWorkzoneMappings(data) {
  const mappings = [];
  for (let i = 1; i < data.length; i++) {
    const team = (data[i][16] || '').trim();
    const wz = (data[i][17] || '').trim();
    if (team && wz) mappings.push({ team, workzone: wz });
  }
  // Remove duplicates
  const seen = new Set();
  return mappings.filter(m => {
    const key = `${m.team}|${m.workzone}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// === HELPER: Find mapping team for a workzone ===
function findMappingTeam(workzone, status, mappings) {
  if (!workzone && status !== 'GAMAS') return null;
  if (status === 'GAMAS') {
    const allMap = mappings.find(m => m.workzone.toUpperCase() === 'ALL');
    return allMap ? allMap.team : null;
  }
  for (const mapping of mappings) {
    if (mapping.workzone.toUpperCase() === 'ALL') continue;
    const zones = mapping.workzone.split(/[&,]/).map(z => z.trim().toUpperCase());
    if (zones.includes(workzone.toUpperCase())) return mapping.team;
  }
  return null;
}

// === HELPER: Parse TTR duration string to hours ===
function parseTTRHours(ttrStr) {
  if (!ttrStr || ttrStr === '-' || ttrStr === '') return null;
  const parts = ttrStr.toString().trim().split(':');
  if (parts.length >= 2) {
    const h = parseFloat(parts[0]) || 0;
    const m = parseFloat(parts[1]) || 0;
    const s = parts.length > 2 ? (parseFloat(parts[2]) || 0) : 0;
    return h + m / 60 + s / 3600;
  }
  const num = parseFloat(ttrStr);
  return isNaN(num) ? null : num;
}

// === HELPER: Format hours to HH:MM:SS ===
function formatHours(hours) {
  const h = Math.floor(hours);
  const m = Math.floor((hours - h) * 60);
  const s = Math.floor(((hours - h) * 60 - m) * 60);
  return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}

// === HELPER: Parse Indonesian date ===
function parseIndonesianDate(dateStr) {
  if (!dateStr) return null;
  const cleaned = dateStr.replace(/^[^,]*,\s*/, '').trim();
  const parts = cleaned.split(/\s+/);
  if (parts.length < 3) return null;
  const day = parseInt(parts[0]);
  const month = BULAN_ID[parts[1].toLowerCase()];
  const year = parseInt(parts[2]);
  if (!day || !month || !year) return null;
  return { day, month, year };
}

// === HELPER: Get today in Jakarta timezone ===
function getTodayJakarta() {
  const now = new Date();
  return {
    day: parseInt(now.toLocaleDateString('id-ID', { day: 'numeric', timeZone: 'Asia/Jakarta' })),
    month: parseInt(now.toLocaleDateString('id-ID', { month: 'numeric', timeZone: 'Asia/Jakarta' })),
    year: parseInt(now.toLocaleDateString('id-ID', { year: 'numeric', timeZone: 'Asia/Jakarta' })),
  };
}

// === HELPER: Parse assurance input text ===
function parseAssurance(text, username) {
  let data = {
    incidentNo: '', closeDesc: '',
    dropcore: '', patchcord: '', soc: '', pslave: '',
    passive1_8: '', passive1_4: '', pigtail: '', adaptor: '',
    roset: '', rj45: '', lan: '',
    dateCreated: new Date().toLocaleDateString('id-ID', {
      weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta',
    }),
    teknisi: (username || '').replace('@', ''),
  };

  const incidentMatch = text.match(/INC[0-9]+/i);
  if (incidentMatch) data.incidentNo = incidentMatch[0].trim().toUpperCase();

  const closeMatch = text.match(/CLOSE\s*:\s*(.+?)(?=\n|MATERIAL|$)/i);
  if (closeMatch && closeMatch[1]) data.closeDesc = closeMatch[1].trim();

  const patterns = {
    dropcore: /DROPCORE\s*:\s*([0-9\.]+)/i, patchcord: /PATCHCORD\s*:\s*([0-9\.]+)/i,
    soc: /SOC\s*:\s*([0-9\.]+)/i, pslave: /PSLAVE\s*:\s*([0-9\.]+)/i,
    passive1_8: /PASSIVE\s*1\/8\s*:\s*([0-9\.]+)/i, passive1_4: /PASSIVE\s*1\/4\s*:\s*([0-9\.]+)/i,
    pigtail: /PIGTAIL\s*:\s*([0-9\.]+)/i, adaptor: /ADAPTOR\s*:\s*([0-9\.]+)/i,
    roset: /ROSET\s*:\s*([0-9\.]+)/i, rj45: /RJ\s*45\s*:\s*([0-9\.]+)/i,
    lan: /LAN\s*:\s*([0-9\.]+)/i,
  };
  for (const [key, pattern] of Object.entries(patterns)) {
    const match = text.match(pattern);
    if (match && match[1]) data[key] = match[1].trim();
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
  bot.setWebHook(webhookUrl).then(() => console.log(`✅ Webhook set: ${webhookUrl}`)).catch(err => console.error('❌ Webhook error:', err.message));
  app.post(`/assurance${TOKEN}`, (req, res) => { bot.processUpdate(req.body); res.sendStatus(200); });
  app.get('/', (req, res) => res.send('Bot Assurance is running!'));
  app.listen(PORT, () => console.log(`✅ Server running on port ${PORT}`));
} else {
  bot = new TelegramBot(TOKEN, { polling: { interval: 300, autoStart: true, params: { timeout: 10, allowed_updates: ['message'] } } });
  console.log('✅ Bot running in polling mode');
  bot.on('polling_error', (error) => console.error(error.code === 'EFATAL' ? '❌ Polling fatal:' : '⚠️ Polling error:', error.message));
}

// ============================================================
// MONITORING FUNCTIONS (TTR Alerts + Auto-Fill Teknisi)
// ============================================================

async function autoFillTeknisi() {
  try {
    const data = await getSheetData(ORDER_ASSURANCE_SHEET, false);
    if (!data || data.length < 2) return;

    const mappings = getWorkzoneMappings(data);
    let filled = 0;

    for (let i = 1; i < data.length; i++) {
      const teknisi = (data[i][2] || '').trim();
      const status = (data[i][9] || '').toUpperCase().trim();
      const workzone = (data[i][4] || '').trim();

      if (teknisi || !status || status === 'CLOSE') continue;

      const team = findMappingTeam(workzone, status, mappings);
      if (team) {
        const rowNum = i + 1;
        await updateSheetCell(ORDER_ASSURANCE_SHEET, `C${rowNum}`, team);
        filled++;
        console.log(`🔧 Auto-fill teknisi row ${rowNum}: ${team}`);
      }
    }

    if (filled > 0) {
      cache.orderAssuranceData = null; // Invalidate cache
      console.log(`✅ Auto-fill: ${filled} teknisi filled`);
    }
  } catch (error) {
    console.error('❌ Auto-fill error:', error.message);
  }
}

async function checkTTRAlerts() {
  if (!GROUP_CHAT_ID) return;

  try {
    const data = await getSheetData(ORDER_ASSURANCE_SHEET, false);
    if (!data || data.length < 2) return;

    const mappings = getWorkzoneMappings(data);
    const admins = await getActiveAdmins();
    const adminTags = admins.map(a => `@${a}`).join(' ');

    for (let i = 1; i < data.length; i++) {
      const incident = (data[i][1] || '').trim();
      const ttrStr = (data[i][3] || '').trim();
      const workzone = (data[i][4] || '').trim();
      const custType = (data[i][5] || '').trim().toUpperCase();
      const status = (data[i][9] || '').toUpperCase().trim();

      if (status !== 'OPEN' || !incident || !ttrStr) continue;

      const elapsed = parseTTRHours(ttrStr);
      if (elapsed === null) continue;

      // Find matching customer type in TTR_TABLE (case-insensitive)
      let maxTTR = null;
      for (const [type, hours] of Object.entries(TTR_TABLE)) {
        if (type.toUpperCase() === custType) { maxTTR = hours; break; }
      }
      if (maxTTR === null) continue;

      const team = findMappingTeam(workzone, status, mappings) || (data[i][2] || '-').trim();
      // Clean team from double @
      const cleanTeam = team.replace(/@@/g, '@');

      // === EXPIRED ===
      if (elapsed >= maxTTR && !alertState.expired.has(incident)) {
        alertState.expired.add(incident);
        alertState.warned.delete(incident);

        const overtime = elapsed - maxTTR;
        let msg = `🔴 <b>TTR EXPIRED!</b>\n\n`;
        msg += `<b>Incident:</b> ${incident}\n`;
        msg += `<b>Customer Type:</b> ${custType} (Max: ${maxTTR} Jam)\n`;
        msg += `<b>Elapsed:</b> ${ttrStr}\n`;
        msg += `<b>Overtime:</b> ${formatHours(overtime)}\n\n`;
        msg += `<b>Teknisi:</b> ${cleanTeam}\n`;
        if (adminTags) msg += `<b>Admin:</b> ${adminTags}`;

        await sendTelegram(GROUP_CHAT_ID, msg);
        console.log(`🔴 TTR EXPIRED: ${incident} (${elapsed.toFixed(1)}h / ${maxTTR}h)`);
      }
      // === WARNING (1 jam sebelum expired) ===
      else if (elapsed >= maxTTR - 1 && elapsed < maxTTR && !alertState.warned.has(incident)) {
        alertState.warned.add(incident);

        const sisa = maxTTR - elapsed;
        let msg = `⚠️ <b>TTR WARNING - MENDEKATI EXPIRED!</b>\n\n`;
        msg += `<b>Incident:</b> ${incident}\n`;
        msg += `<b>Customer Type:</b> ${custType} (Max: ${maxTTR} Jam)\n`;
        msg += `<b>Elapsed:</b> ${ttrStr}\n`;
        msg += `<b>Sisa:</b> ${formatHours(sisa)}\n\n`;
        msg += `<b>Teknisi:</b> ${cleanTeam}\n`;
        if (adminTags) msg += `<b>Admin:</b> ${adminTags}`;

        await sendTelegram(GROUP_CHAT_ID, msg);
        console.log(`⚠️ TTR WARNING: ${incident} (${elapsed.toFixed(1)}h / ${maxTTR}h)`);
      }
    }

    // Cleanup: remove closed incidents from alert state
    const openIncidents = new Set();
    for (let i = 1; i < data.length; i++) {
      const status = (data[i][9] || '').toUpperCase().trim();
      if (status === 'OPEN') openIncidents.add((data[i][1] || '').trim());
    }
    for (const inc of alertState.warned) { if (!openIncidents.has(inc)) alertState.warned.delete(inc); }
    for (const inc of alertState.expired) { if (!openIncidents.has(inc)) alertState.expired.delete(inc); }

  } catch (error) {
    console.error('❌ TTR check error:', error.message);
  }
}

// ============================================================
// MESSAGE HANDLER
// ============================================================
bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const msgId = msg.message_id;
  const text = (msg.text || '').trim();
  const username = msg.from.username || '';
  const groupType = msg.chat.type;

  if (!text || !text.startsWith('/')) return;

  // Store chat ID for potential DM
  if (username) userChatIds[username.replace('@', '').toLowerCase()] = chatId;

  console.log(`📨 [${groupType}] [@${username}] ${text.substring(0, 60)}`);

  try {
    // ============================================================
    // /INPUT - Input data assurance + auto-close di ORDER ASSURANCE
    // ============================================================
    if (/^\/INPUT\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['USER', 'ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const inputText = text.replace(/^\/INPUT\s*/i, '').trim();
        if (!inputText) return sendTelegram(chatId, '❌ Silakan kirim data assurance setelah /INPUT.', { reply_to_message_id: msgId });

        const parsed = parseAssurance(inputText, username);
        const missing = ['incidentNo', 'closeDesc'].filter(f => !parsed[f]);
        if (missing.length > 0) return sendTelegram(chatId, `❌ Field wajib: ${missing.join(', ')}`, { reply_to_message_id: msgId });

        // Simpan ke PROGRES ASSURANCE
        const row = [
          parsed.dateCreated, parsed.incidentNo,
          parsed.dropcore, parsed.patchcord, parsed.soc, parsed.pslave,
          parsed.passive1_8, parsed.passive1_4, parsed.pigtail, parsed.adaptor,
          parsed.roset, parsed.rj45, parsed.lan, parsed.closeDesc, parsed.teknisi,
        ];
        await withTimeout(appendSheetData(ASSURANCE_SHEET, row), 10000);

        // Auto-close di ORDER ASSURANCE
        let orderClosed = false;
        try {
          const orderData = await getSheetData(ORDER_ASSURANCE_SHEET, false);
          for (let i = 1; i < orderData.length; i++) {
            const incInOrder = (orderData[i][1] || '').trim().toUpperCase();
            if (incInOrder === parsed.incidentNo) {
              await updateSheetCell(ORDER_ASSURANCE_SHEET, `J${i + 1}`, 'CLOSE');
              cache.orderAssuranceData = null;
              alertState.warned.delete(parsed.incidentNo);
              alertState.expired.delete(parsed.incidentNo);
              orderClosed = true;
              console.log(`✅ Auto-close ORDER: ${parsed.incidentNo} row ${i + 1}`);
              break;
            }
          }
        } catch (closeErr) {
          console.error('⚠️ Auto-close error:', closeErr.message);
        }

        let confirmMsg = `✅ Data Assurance berhasil disimpan!\n\n`;

        return sendTelegram(chatId, confirmMsg, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /INPUT Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /sisa_ticket - Ticket OPEN grouped by mapping team
    // ============================================================
    else if (/^\/sisa_ticket\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ORDER_ASSURANCE_SHEET), 10000);
        const mappings = getWorkzoneMappings(data);

        const now = new Date();
        const dayNameID = ['MINGGU', 'SENIN', 'SELASA', 'RABU', 'KAMIS', 'JUMAT', 'SABTU'][now.getDay()];
        const dateStr = now.toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' }).toUpperCase();
        const timeStr = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Jakarta' });

        // Group OPEN tickets by mapping team
        let ticketsByTeam = {};
        let totalOpen = 0;

        for (let i = 1; i < data.length; i++) {
          const incident = (data[i][1] || '-').trim();
          const ttrCustomer = (data[i][3] || '-').trim();
          const workzone = (data[i][4] || '').trim();
          const status = (data[i][9] || '').toUpperCase().trim();

          if (status !== 'OPEN') continue;

          const team = findMappingTeam(workzone, status, mappings) || workzone || '-';
          const cleanTeam = team.replace(/@@/g, '@');

          if (!ticketsByTeam[cleanTeam]) ticketsByTeam[cleanTeam] = [];
          ticketsByTeam[cleanTeam].push({ incident, ttr: ttrCustomer });
          totalOpen++;
        }

        const sortedTeams = Object.keys(ticketsByTeam).sort();

        let response = `🔴 <b>SISA TICKET OPEN ${dayNameID}, ${dateStr} ${timeStr}</b>\n\n`;
        response += `<b>Total OPEN : ${totalOpen} tickets</b>\n\n`;

        if (sortedTeams.length === 0) {
          response += '<i>Tidak ada ticket yang masih OPEN</i>';
        } else {
          sortedTeams.forEach((teamName, idx) => {
            const tickets = ticketsByTeam[teamName];
            response += `${idx + 1}. <b>${teamName}</b>\n`;
            tickets.forEach(t => {
              response += `   ${t.incident}   ${t.ttr}\n`;
            });
            response += '\n';
          });
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /sisa_ticket Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /material_used - Total material keseluruhan
    // ============================================================
    else if (/^\/material_used\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        let materialMap = {
          'DROPCORE': 0, 'PATCHCORD': 0, 'SOC': 0, 'PSLAVE': 0,
          'PASSIVE 1/8': 0, 'PASSIVE 1/4': 0, 'PIGTAIL': 0, 'ADAPTOR': 0,
          'ROSET': 0, 'RJ 45': 0, 'LAN': 0,
        };
        const materialColumns = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
        const materialNames = Object.keys(materialMap);

        for (let i = 1; i < data.length; i++) {
          materialColumns.forEach((colIdx, idx) => {
            materialMap[materialNames[idx]] += parseInt((data[i][colIdx] || '0').trim()) || 0;
          });
        }

        const entries = Object.entries(materialMap).filter(([_, c]) => c > 0).sort((a, b) => b[1] - a[1]);
        let response = `📦 <b>PENGGUNAAN MATERIAL - KESELURUHAN</b>\n\n`;
        if (entries.length === 0) {
          response += '<i>Belum ada material yang dipakai</i>';
        } else {
          entries.forEach(([mat, count]) => { response += `• <b>${mat}</b>: ${count} unit\n`; });
        }

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /material_used Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /rekap_close - Rekap total close per teknisi (hari/bulan/tahun)
    // ============================================================
    else if (/^\/rekap_close\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username, ['ADMIN']);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const data = await withTimeout(getSheetData(ASSURANCE_SHEET), 10000);
        const today = getTodayJakarta();

        let hariIni = {}, bulanIni = {}, tahunIni = {};

        for (let i = 1; i < data.length; i++) {
          const tanggal = (data[i][0] || '').trim();
          const teknisi = (data[i][14] || '-').trim();
          const d = parseIndonesianDate(tanggal);
          if (!d) continue;

          // Tahun ini
          if (d.year === today.year) {
            tahunIni[teknisi] = (tahunIni[teknisi] || 0) + 1;

            // Bulan ini
            if (d.month === today.month) {
              bulanIni[teknisi] = (bulanIni[teknisi] || 0) + 1;

              // Hari ini
              if (d.day === today.day) {
                hariIni[teknisi] = (hariIni[teknisi] || 0) + 1;
              }
            }
          }
        }

        const now = new Date();
        const todayStr = now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
        const bulanStr = now.toLocaleDateString('id-ID', { month: 'long', year: 'numeric', timeZone: 'Asia/Jakarta' });
        const tahunStr = now.toLocaleDateString('id-ID', { year: 'numeric', timeZone: 'Asia/Jakarta' });

        function buildSection(title, map) {
          const entries = Object.entries(map).sort((a, b) => b[1] - a[1]);
          const total = entries.reduce((sum, [_, c]) => sum + c, 0);
          let s = `📅 <b>${title}:</b>\n`;
          if (entries.length === 0) {
            s += '<i>Belum ada data</i>\n';
          } else {
            entries.forEach(([tek, c]) => { s += `🔸 <b>${tek}</b>: ${c} tickets\n`; });
            s += `<b>Total: ${total} tickets</b>\n`;
          }
          return s;
        }

        let response = `📊 <b>REKAP CLOSE TEKNISI</b>\n\n`;
        response += buildSection(`HARI INI (${todayStr})`, hariIni) + '\n';
        response += buildSection(`BULAN INI (${bulanStr})`, bulanIni) + '\n';
        response += buildSection(`TAHUN INI (${tahunStr})`, tahunIni);

        return sendTelegram(chatId, response, { reply_to_message_id: msgId });
      } catch (err) {
        console.error('❌ /rekap_close Error:', err.message);
        return sendTelegram(chatId, `❌ Error: ${err.message}`, { reply_to_message_id: msgId });
      }
    }

    // ============================================================
    // /help or /start
    // ============================================================
    else if (/^\/(help|start)\b/i.test(text)) {
      try {
        const authResult = await checkAuthorization(username);
        if (!authResult.authorized) return sendTelegram(chatId, authResult.message, { reply_to_message_id: msgId });

        const helpMsg = `🤖 <b>Bot Assurance</b>

<b>📝 INPUT COMMAND:</b>
/INPUT - Input data assurance (auto-close ORDER)

<b>📊 MONITORING (ADMIN):</b>
/sisa_ticket - Ticket yang masih OPEN
/material_used - Total material yang dipakai
/rekap_close - Rekap close per teknisi (hari/bulan/tahun)

<b>📋 FORMAT /INPUT:</b>
/INPUT INC47052822
CLOSE: deskripsi perbaikan
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
LAN: 0

<b>⚙️ FITUR OTOMATIS:</b>
• Auto-fill teknisi berdasarkan workzone
• TTR monitoring & alert ke group
• Auto-close status saat /INPUT`;

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

// === ERROR HANDLER ===
process.on('unhandledRejection', (reason) => console.error('Error:', reason));

// === STARTUP ===
console.log('\n🚀 Bot Assurance started!');
console.log(`Mode: ${USE_WEBHOOK ? 'Webhook' : 'Polling'}`);
console.log('═'.repeat(50));
console.log('✅ Auto-Cache Enabled (5 min expiry)');
console.log('✅ TTR Monitoring Enabled (5 min interval)');
console.log('✅ Auto-Fill Teknisi Enabled');
console.log('✅ Timeout Protection Enabled');
console.log('═'.repeat(50));

// Start monitoring after 10 seconds delay
setTimeout(() => {
  console.log('🔄 Starting initial monitoring check...');
  autoFillTeknisi();
  checkTTRAlerts();

  // Then every 5 minutes
  setInterval(() => {
    autoFillTeknisi();
    checkTTRAlerts();
  }, 5 * 60 * 1000);
}, 10000);
