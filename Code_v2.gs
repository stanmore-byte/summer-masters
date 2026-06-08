// ═══════════════════════════════════════════════════════════════
//  AZZURRI STORM — Google Apps Script Backend v2
//  Paste this entire file into your Apps Script editor.
//  SPREADSHEET_ID is already set below — do not change it.
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1RHGMB9n8Vet-0eT8Mq4NlbcQSl3Ds_lPSD89odd0HuI';

// ── ENTRY POINTS ────────────────────────────────────────────────

function doGet(e) {
  const cb     = e.parameter.callback;
  const team   = (e.parameter.team   || 'u11').toLowerCase();
  const player = e.parameter.player  || '';

  try {
    const data = getAllPlayerData(team, player);
    const json = JSON.stringify({ ok: true, data: data });
    if (cb) return out(cb + '(' + json + ')', 'JAVASCRIPT');
    return out(json, 'JSON');
  } catch (err) {
    const json = JSON.stringify({ ok: false, error: err.message });
    if (cb) return out(cb + '(' + json + ')', 'JAVASCRIPT');
    return out(json, 'JSON');
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    savePlayerData(payload);
    return out(JSON.stringify({ ok: true }), 'JSON');
  } catch (err) {
    return out(JSON.stringify({ ok: false, error: err.message }), 'JSON');
  }
}

function out(text, mime) {
  return ContentService.createTextOutput(text)
    .setMimeType(ContentService.MimeType[mime]);
}

// ── DATE HELPER ──────────────────────────────────────────────────
// Google Sheets auto-converts date-like strings into Date objects.
// This always returns YYYY-MM-DD regardless of how it was stored.
function fmtDate(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).trim();
}

// ── READ ─────────────────────────────────────────────────────────

function getAllPlayerData(team, player) {
  const pfx = team === 'u13' ? 'U13' : 'U11';
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    daily   : getDailyData   (ss, pfx, player),
    juggling: getJugglingData(ss, pfx, player),
    cde     : getCDEData     (ss, pfx, player),
    profile : getProfileData (ss, pfx, player),
    plans   : getPlansData   (ss, pfx, player)
  };
}

function getDailyData(ss, pfx, player) {
  const sheet = ss.getSheetByName(pfx + '_Daily');
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== player) continue;
    const date = fmtDate(rows[i][1]);
    if (!date) continue;
    result[date] = {
      bs: rows[i][2] ? String(rows[i][2]).split(',').filter(Boolean) : [],
      st: rows[i][3] ? String(rows[i][3]).split(',').filter(Boolean) : []
    };
  }
  return result;
}

function getJugglingData(ss, pfx, player) {
  const sheet = ss.getSheetByName(pfx + '_Juggling');
  if (!sheet) return { pr: 0, log: [] };
  const rows = sheet.getDataRange().getValues();
  const log  = [];
  let pr     = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== player) continue;
    const count = parseInt(rows[i][2]) || 0;
    if (count > pr) pr = count;
    const date = fmtDate(rows[i][1]);
    if (date) log.push({ date: date, count: count });
  }
  log.sort((a, b) => b.date.localeCompare(a.date));
  return { pr: pr, log: log };
}

function getCDEData(ss, pfx, player) {
  const sheet = ss.getSheetByName(pfx + '_CDE');
  if (!sheet) return [];
  const rows    = sheet.getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== player) continue;
    entries.push({
      id  : rows[i][2],
      date: fmtDate(rows[i][1]),
      text: String(rows[i][3] || ''),
      photo: null
    });
  }
  return entries.sort((a, b) => Number(b.id) - Number(a.id));
}

function getProfileData(ss, pfx, player) {
  const sheet = ss.getSheetByName(pfx + '_Profiles');
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== player) continue;
    return { nickname: rows[i][1] || '', number: rows[i][2] || '', color: rows[i][3] || '' };
  }
  return {};
}

function getPlansData(ss, pfx, player) {
  const sheet = ss.getSheetByName(pfx + '_Plans');
  if (!sheet) return null;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== player) continue;
    try { return JSON.parse(rows[i][1]); } catch(e) { return null; }
  }
  return null;
}

// ── WRITE ────────────────────────────────────────────────────────

function savePlayerData(payload) {
  const pfx    = payload.team === 'u13' ? 'U13' : 'U11';
  const player = payload.player;
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const d      = payload.data || {};
  const t      = payload.type;

  if      (t === 'daily')          saveDailyRow       (ss, pfx, player, d);
  else if (t === 'juggling')       saveJugglingRow    (ss, pfx, player, d);
  else if (t === 'juggling_delete')deleteJugglingRow  (ss, pfx, player, d.date);
  else if (t === 'cde')            saveCDERow         (ss, pfx, player, d);
  else if (t === 'cde_delete')     deleteCDERow       (ss, pfx, player, d.id);
  else if (t === 'profile')        saveProfileRow     (ss, pfx, player, d);
  else if (t === 'plans')          savePlansRow       (ss, pfx, player, d);
}

function saveDailyRow(ss, pfx, player, dayObj) {
  const sheet = getOrCreate(ss, pfx + '_Daily',
    ['Player','Date','Ball Skills Checked (IDs)','Strength Checked (IDs)','Last Updated']);
  const rows = sheet.getDataRange().getValues();
  const now  = new Date().toISOString();

  for (const [date, vals] of Object.entries(dayObj)) {
    const bs = (vals.bs || []).join(',');
    const st = (vals.st || []).join(',');
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === player && fmtDate(rows[i][1]) === date) {
        sheet.getRange(i+1, 3, 1, 3).setValues([[bs, st, now]]);
        found = true; break;
      }
    }
    if (!found) {
      const newRow = sheet.getLastRow() + 1;
      sheet.appendRow([player, date, bs, st, now]);
      // Force date column to plain text so it reads back as a string
      sheet.getRange(newRow, 2).setNumberFormat('@STRING@');
    }
  }
}

function saveJugglingRow(ss, pfx, player, entry) {
  const sheet = getOrCreate(ss, pfx + '_Juggling',
    ['Player','Date','Juggle Count','Last Updated']);
  const rows = sheet.getDataRange().getValues();
  const now  = new Date().toISOString();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === player && fmtDate(rows[i][1]) === String(entry.date)) {
      sheet.getRange(i+1, 3, 1, 2).setValues([[entry.count, now]]);
      return;
    }
  }
  const newRow = sheet.getLastRow() + 1;
  sheet.appendRow([player, entry.date, entry.count, now]);
  sheet.getRange(newRow, 2).setNumberFormat('@STRING@');
}

function deleteJugglingRow(ss, pfx, player, date) {
  const sheet = ss.getSheetByName(pfx + '_Juggling');
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() === player && fmtDate(rows[i][1]) === String(date)) {
      sheet.deleteRow(i+1); break;
    }
  }
}

function saveCDERow(ss, pfx, player, entry) {
  const sheet = getOrCreate(ss, pfx + '_CDE',
    ['Player','Date','Entry ID','Activity Text','Updated']);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === player && String(rows[i][2]) === String(entry.id)) return;
  }
  const newRow = sheet.getLastRow() + 1;
  sheet.appendRow([player, entry.date, entry.id, entry.text || '', new Date().toISOString()]);
  sheet.getRange(newRow, 2).setNumberFormat('@STRING@');
}

function deleteCDERow(ss, pfx, player, entryId) {
  const sheet = ss.getSheetByName(pfx + '_CDE');
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() === player && String(rows[i][2]) === String(entryId)) {
      sheet.deleteRow(i+1); break;
    }
  }
}

function saveProfileRow(ss, pfx, player, profile) {
  const sheet = getOrCreate(ss, pfx + '_Profiles',
    ['Player','Nickname','Jersey #','Avatar Color','Last Updated']);
  const rows = sheet.getDataRange().getValues();
  const now  = new Date().toISOString();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === player) {
      sheet.getRange(i+1, 2, 1, 4).setValues(
        [[profile.nickname||'', profile.number||'', profile.color||'', now]]);
      return;
    }
  }
  sheet.appendRow([player, profile.nickname||'', profile.number||'', profile.color||'', now]);
}

function savePlansRow(ss, pfx, player, plans) {
  const sheet = getOrCreate(ss, pfx + '_Plans',
    ['Player','Custom Plan (JSON)','Last Updated']);
  const rows = sheet.getDataRange().getValues();
  const now  = new Date().toISOString();
  const json = JSON.stringify(plans);

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === player) {
      sheet.getRange(i+1, 2, 1, 2).setValues([[json, now]]);
      return;
    }
  }
  sheet.appendRow([player, json, now]);
}

// ── HELPER ───────────────────────────────────────────────────────

function getOrCreate(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    const hdrRange = sheet.getRange(1, 1, 1, headers.length);
    hdrRange.setFontWeight('bold');
    hdrRange.setBackground('#263238');
    hdrRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
