// ═══════════════════════════════════════════════════════════════
//  AZZURRI STORM — Google Apps Script Backend v3
//  Paste this entire file into Apps Script, then deploy new version.
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1RHGMB9n8Vet-0eT8Mq4NlbcQSl3Ds_lPSD89odd0HuI';

// ── ENTRY POINTS ─────────────────────────────────────────────────

function doGet(e) {
  const cb     = e.parameter.callback;
  const team   = (e.parameter.team || 'u11').toLowerCase();
  const player = e.parameter.player || '';
  const all    = e.parameter.all === '1';   // batch mode: all players at once

  try {
    const data = all ? getAllTeamData(team) : getAllPlayerData(team, player);
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

// ── DATE HELPER ───────────────────────────────────────────────────
// Google Sheets silently converts date-like strings to Date objects.
// This always returns YYYY-MM-DD regardless.
function fmtDate(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).trim();
}

// ── SINGLE PLAYER READ ───────────────────────────────────────────

function getAllPlayerData(team, player) {
  const pfx = team === 'u13' ? 'U13' : (team === 'frn' ? 'FRN' : 'U11');
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    daily   : getDailyData   (ss, pfx, player),
    juggling: getJugglingData(ss, pfx, player),
    cde     : getCDEData     (ss, pfx, player),
    profile : getProfileData (ss, pfx, player),
    plans   : getPlansData   (ss, pfx, player)
  };
}

// ── BATCH READ (all players, used by Leaderboard) ────────────────

// ── FRN BATCH: finds friends in FRN_, U11_, U13_ tabs (no migration needed) ──
function getAllFRNData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const allTeam = new Set([...U11_TEAM, ...U13_TEAM]);
  const result = {};

  function scanPrefix(pfx) {
    const dSheet = ss.getSheetByName(pfx + '_Daily');
    if (dSheet) {
      const rows = dSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const p = String(rows[i][0]).trim();
        if (!p || (pfx !== 'FRN' && allTeam.has(p))) continue;
        const d = fmtDate(rows[i][1]);
        if (!d) continue;
        if (!result[p]) result[p] = {daily:{},juggling:{pr:0,log:[]},cde:[],profile:{},plans:null};
        if (!result[p].daily[d]) result[p].daily[d] = {bs:new Set(),st:new Set()};
        (rows[i][2]?String(rows[i][2]).split(',').filter(Boolean):[]).forEach(id=>result[p].daily[d].bs.add(id));
        (rows[i][3]?String(rows[i][3]).split(',').filter(Boolean):[]).forEach(id=>result[p].daily[d].st.add(id));
      }
    }
    const jSheet = ss.getSheetByName(pfx + '_Juggling');
    if (jSheet) {
      const rows = jSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const p = String(rows[i][0]).trim();
        if (!p || (pfx !== 'FRN' && allTeam.has(p))) continue;
        const count = parseInt(rows[i][2]) || 0;
        const d = fmtDate(rows[i][1]);
        if (!d) continue;
        if (!result[p]) result[p] = {daily:{},juggling:{pr:0,log:[]},cde:[],profile:{},plans:null};
        if (count > result[p].juggling.pr) result[p].juggling.pr = count;
        result[p].juggling.log.push({date:d,count:count});
      }
    }
    const cSheet = ss.getSheetByName(pfx + '_CDE');
    if (cSheet) {
      const rows = cSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const p = String(rows[i][0]).trim();
        if (!p || (pfx !== 'FRN' && allTeam.has(p))) continue;
        if (!result[p]) result[p] = {daily:{},juggling:{pr:0,log:[]},cde:[],profile:{},plans:null};
        result[p].cde.push({id:rows[i][2],date:fmtDate(rows[i][1]),text:String(rows[i][3]||''),photo:null});
      }
    }
    const prSheet = ss.getSheetByName(pfx + '_Profiles');
    if (prSheet) {
      const rows = prSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const p = String(rows[i][0]).trim();
        if (!p || (pfx !== 'FRN' && allTeam.has(p))) continue;
        if (!result[p]) result[p] = {daily:{},juggling:{pr:0,log:[]},cde:[],profile:{},plans:null};
        result[p].profile = {nickname:rows[i][1]||'',number:String(rows[i][2]||''),color:rows[i][3]||''};
      }
    }
  }

  scanPrefix('FRN');
  scanPrefix('U11');
  scanPrefix('U13');

  for (const p in result) {
    for (const d in result[p].daily) {
      result[p].daily[d] = {bs:Array.from(result[p].daily[d].bs),st:Array.from(result[p].daily[d].st)};
    }
    result[p].juggling.log.sort((a,b)=>b.date.localeCompare(a.date));
    result[p].cde.sort((a,b)=>Number(b.id)-Number(a.id));
  }
  return result;
}

function getAllTeamData(team) {
  if (team === 'frn') return getAllFRNData();
  const pfx = team === 'u13' ? 'U13' : 'U11';
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Read each sheet once and group by player
  const daily    = {};
  const juggling = {};
  const cde      = {};
  const profiles = {};

  // Daily — union all rows per player+date (handles historical duplicates)
  const dSheet = ss.getSheetByName(pfx + '_Daily');
  if (dSheet) {
    const rows = dSheet.getDataRange().getValues();
    const dailyRaw = {};
    for (let i = 1; i < rows.length; i++) {
      const p = String(rows[i][0]).trim();
      const d = fmtDate(rows[i][1]);
      if (!p || !d) continue;
      if (!dailyRaw[p]) dailyRaw[p] = {};
      if (!dailyRaw[p][d]) dailyRaw[p][d] = { bs: new Set(), st: new Set() };
      (rows[i][2] ? String(rows[i][2]).split(',').filter(Boolean) : []).forEach(id => dailyRaw[p][d].bs.add(id));
      (rows[i][3] ? String(rows[i][3]).split(',').filter(Boolean) : []).forEach(id => dailyRaw[p][d].st.add(id));
    }
    for (const p in dailyRaw) {
      daily[p] = {};
      for (const d in dailyRaw[p]) {
        daily[p][d] = { bs: Array.from(dailyRaw[p][d].bs), st: Array.from(dailyRaw[p][d].st) };
      }
    }
  }

  // Juggling
  const jSheet = ss.getSheetByName(pfx + '_Juggling');
  if (jSheet) {
    const rows = jSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const p     = String(rows[i][0]).trim();
      const d     = fmtDate(rows[i][1]);
      const count = parseInt(rows[i][2]) || 0;
      if (!p || !d) continue;
      if (!juggling[p]) juggling[p] = { pr: 0, log: [] };
      if (count > juggling[p].pr) juggling[p].pr = count;
      juggling[p].log.push({ date: d, count: count });
    }
    for (const p in juggling) {
      juggling[p].log.sort((a, b) => b.date.localeCompare(a.date));
    }
  }

  // CDE
  const cSheet = ss.getSheetByName(pfx + '_CDE');
  if (cSheet) {
    const rows = cSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const p = String(rows[i][0]).trim();
      if (!p) continue;
      if (!cde[p]) cde[p] = [];
      cde[p].push({
        id   : rows[i][2],
        date : fmtDate(rows[i][1]),
        text : String(rows[i][3] || ''),
        photo: null
      });
    }
    for (const p in cde) {
      cde[p].sort((a, b) => Number(b.id) - Number(a.id));
    }
  }

  // Profiles
  const prSheet = ss.getSheetByName(pfx + '_Profiles');
  if (prSheet) {
    const rows = prSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const p = String(rows[i][0]).trim();
      if (!p) continue;
      profiles[p] = {
        nickname: rows[i][1] || '',
        number  : String(rows[i][2] || ''),
        color   : rows[i][3] || ''
      };
    }
  }

  // Combine into one object per player
  const allPlayers = new Set([
    ...Object.keys(daily),
    ...Object.keys(juggling),
    ...Object.keys(cde),
    ...Object.keys(profiles)
  ]);

  const result = {};
  allPlayers.forEach(p => {
    result[p] = {
      daily   : daily[p]    || {},
      juggling: juggling[p] || { pr: 0, log: [] },
      cde     : cde[p]      || [],
      profile : profiles[p] || {},
      plans   : null
    };
  });

  return result;
}

// ── INDIVIDUAL SHEET READERS ─────────────────────────────────────

function getDailyData(ss, pfx, player) {
  const sheet = ss.getSheetByName(pfx + '_Daily');
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues();
  // Union all rows per date — recovers full state from historical duplicate rows
  const raw = {};
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== player) continue;
    const date = fmtDate(rows[i][1]);
    if (!date) continue;
    if (!raw[date]) raw[date] = { bs: new Set(), st: new Set() };
    (rows[i][2] ? String(rows[i][2]).split(',').filter(Boolean) : []).forEach(id => raw[date].bs.add(id));
    (rows[i][3] ? String(rows[i][3]).split(',').filter(Boolean) : []).forEach(id => raw[date].st.add(id));
  }
  const result = {};
  for (const date in raw) {
    result[date] = { bs: Array.from(raw[date].bs), st: Array.from(raw[date].st) };
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
      id   : rows[i][2],
      date : fmtDate(rows[i][1]),
      text : String(rows[i][3] || ''),
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
    return {
      nickname: rows[i][1] || '',
      number  : String(rows[i][2] || ''),
      color   : rows[i][3] || ''
    };
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

// ── WRITE ─────────────────────────────────────────────────────────

function savePlayerData(payload) {
  const pfx    = payload.team === 'u13' ? 'U13' : (payload.team === 'frn' ? 'FRN' : 'U11');
  const player = payload.player;
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const d      = payload.data || {};
  const t      = payload.type;

  if      (t === 'daily')           saveDailyRow      (ss, pfx, player, d);
  else if (t === 'juggling')        saveJugglingRow   (ss, pfx, player, d);
  else if (t === 'juggling_delete') deleteJugglingRow (ss, pfx, player, d.date);
  else if (t === 'cde')             saveCDERow        (ss, pfx, player, d);
  else if (t === 'cde_delete')      deleteCDERow      (ss, pfx, player, d.id);
  else if (t === 'profile')         saveProfileRow    (ss, pfx, player, d);
  else if (t === 'plans')           savePlansRow      (ss, pfx, player, d);
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

// ── HELPER ────────────────────────────────────────────────────────

function getOrCreate(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    const hdr = sheet.getRange(1, 1, 1, headers.length);
    hdr.setFontWeight('bold');
    hdr.setBackground('#263238');
    hdr.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── MIGRATION: run once from Apps Script editor ──────────────────
// Open Apps Script → select migrateFriendsToFRN → click Run
function migrateFriendsToFRN() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const allTeam = new Set([...U11_TEAM, ...U13_TEAM]);
  const tabs = ['Daily','Juggling','CDE','Profiles'];
  let moved = 0;

  tabs.forEach(tab => {
    const u11Sheet = ss.getSheetByName('U11_' + tab);
    if (!u11Sheet) return;
    const rows = u11Sheet.getDataRange().getValues();
    const headers = rows[0];
    const toMove = [];
    const toKeep = [headers];

    for (let i = 1; i < rows.length; i++) {
      const player = String(rows[i][0]).trim();
      if (player && !allTeam.has(player)) {
        toMove.push(rows[i]);
      } else {
        toKeep.push(rows[i]);
      }
    }

    if (toMove.length === 0) return;

    const frnSheet = getOrCreate(ss, 'FRN_' + tab, headers);
    toMove.forEach(row => {
      const existing = frnSheet.getDataRange().getValues();
      const isDupe = existing.some(r =>
        String(r[0]).trim() === String(row[0]).trim() &&
        String(r[1]) === String(row[1]) &&
        String(r[2]) === String(row[2])
      );
      if (!isDupe) { frnSheet.appendRow(row); moved++; }
    });

    u11Sheet.clearContents();
    toKeep.forEach((row, i) => {
      u11Sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
    });
  });

  Logger.log('Migration complete. Rows moved: ' + moved);
  SpreadsheetApp.getUi().alert('Done! Moved ' + moved + ' rows to FRN_ tabs.');
}
