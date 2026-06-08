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

function getAllTeamData(team) {
  const pfx = team === 'u13' ? 'U13' : (team === 'frn' ? 'FRN' : 'U11');
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

// ── MIGRATION: run once from Apps Script editor ──────────────────
// Open Apps Script → select migrateFriendsToFRN → click Run
// This moves friends' data from U11_ tabs to FRN_ tabs automatically.
const U11_TEAM = ['Louise','Mia','Audrey','Charlotte','Avi','Annika','Mikayla','Aria','Aubree','Gracie','Amelia','Shailynn','Ellie'];
const U13_TEAM = ['Madi','Marie','Susie','Hailey','Reese','Noelle','Charlotte','Kaia','Melissa','Baylor','Caroline','Hannah','Camryn','London','Ellie','Myra'];

function migrateFriendsToFRN() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tabs = ['Daily','Juggling','CDE','Profiles'];
  const allTeam = new Set([...U11_TEAM, ...U13_TEAM]);
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

    // Write to FRN_ tab
    const frnSheet = getOrCreate(ss, 'FRN_' + tab, headers);
    toMove.forEach(row => {
      // Check not already there (avoid dupes)
      const existing = frnSheet.getDataRange().getValues();
      const isDupe = existing.some(r =>
        String(r[0]).trim() === String(row[0]).trim() &&
        String(r[1]) === String(row[1]) &&
        String(r[2]) === String(row[2])
      );
      if (!isDupe) { frnSheet.appendRow(row); moved++; }
    });

    // Rewrite U11_ without the moved rows
    u11Sheet.clearContents();
    toKeep.forEach((row, i) => {
      u11Sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
    });

    Logger.log('Migrated ' + toMove.length + ' rows from U11_' + tab + ' to FRN_' + tab);
  });

  Logger.log('Migration complete. Total rows moved: ' + moved);
  SpreadsheetApp.getUi().alert('Done! Moved ' + moved + ' rows to FRN_ tabs. Check the Execution log for details.');
}
// ─────────────────────────────────────────────────────────────────

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


// ── ONE-TIME DATA FIX ─────────────────────────────────────────────
// Run this ONCE from Apps Script to fix Myra's June 4 data
// Select fixMyraJune4 in the function dropdown → click Run
function fixMyraJune4() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const date = '2026-06-04';
  const now  = new Date().toISOString();
  const bs   = 'bs1,bs2,bs3,bs4,bs5,bs6';
  const st   = 'st1,st2,st3,st4,st5,st6';

  // ── U13_Daily: add/update Myra for June 4 ──
  const dailySheet = getOrCreate(ss, 'U13_Daily',
    ['Player','Date','Ball Skills Checked (IDs)','Strength Checked (IDs)','Last Updated']);
  const dRows = dailySheet.getDataRange().getValues();
  let myraFound = false;
  for (let i = 1; i < dRows.length; i++) {
    if (String(dRows[i][0]).trim() === 'Myra' && fmtDate(dRows[i][1]) === date) {
      dailySheet.getRange(i+1, 3, 1, 3).setValues([[bs, st, now]]);
      myraFound = true;
      Logger.log('Updated existing Myra row at row ' + (i+1));
      break;
    }
  }
  if (!myraFound) {
    const newRow = dailySheet.getLastRow() + 1;
    dailySheet.appendRow(['Myra', date, bs, st, now]);
    dailySheet.getRange(newRow, 2).setNumberFormat('@STRING@');
    Logger.log('Added new Myra row for ' + date);
  }

  // ── U13_Juggling: add Myra juggling=30 for June 4 ──
  const jugSheet = getOrCreate(ss, 'U13_Juggling',
    ['Player','Date','Juggle Count','Last Updated']);
  const jRows = jugSheet.getDataRange().getValues();
  let myraJugFound = false;
  for (let i = 1; i < jRows.length; i++) {
    if (String(jRows[i][0]).trim() === 'Myra' && fmtDate(jRows[i][1]) === date) {
      jugSheet.getRange(i+1, 3, 1, 2).setValues([[30, now]]);
      myraJugFound = true;
      Logger.log('Updated Myra juggling row');
      break;
    }
  }
  if (!myraJugFound) {
    const newRow = jugSheet.getLastRow() + 1;
    jugSheet.appendRow(['Myra', date, 30, now]);
    jugSheet.getRange(newRow, 2).setNumberFormat('@STRING@');
    Logger.log('Added Myra juggling=30 for ' + date);
  }

  Logger.log('Fix complete. Myra now has 6/6 BS, 6/6 ST, juggling=30 for 2026-06-04.');
  SpreadsheetApp.getUi().alert(
    '✅ Done!\n\nMyra → June 4, 2026:\n• Ball Skills: 6/6 ✓\n• Strength: 6/6 ✓\n• Juggling: 30\n\nLondon\'s data is unchanged (she\'s assumed to have done the work too).'
  );
}

// ── ONE-TIME FIX: Reese June 4 ─────────────────────────────────
// Adds Wall-offs (bs1, bs2) to existing checkmarks + sets juggling = 89
// Select fixReeseJune4 in dropdown → click Run
function fixReeseJune4() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const date = '2026-06-04';
  const now  = new Date().toISOString();

  const dailySheet = getOrCreate(ss, 'U13_Daily',
    ['Player','Date','Ball Skills Checked (IDs)','Strength Checked (IDs)','Last Updated']);
  const dRows = dailySheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < dRows.length; i++) {
    if (String(dRows[i][0]).trim() === 'Reese' && fmtDate(dRows[i][1]) === date) {
      const existingBS = dRows[i][2] ? String(dRows[i][2]).split(',').filter(Boolean) : [];
      const bsSet = new Set(existingBS);
      bsSet.add('bs1'); bsSet.add('bs2');
      const existingST = dRows[i][3] ? String(dRows[i][3]) : '';
      dailySheet.getRange(i+1,3,1,3).setValues([[Array.from(bsSet).join(','),existingST,now]]);
      found = true; break;
    }
  }
  if (!found) {
    const newRow = dailySheet.getLastRow() + 1;
    dailySheet.appendRow(['Reese', date, 'bs1,bs2', '', now]);
    dailySheet.getRange(newRow, 2).setNumberFormat('@STRING@');
  }

  const jugSheet = getOrCreate(ss, 'U13_Juggling',
    ['Player','Date','Juggle Count','Last Updated']);
  const jRows = jugSheet.getDataRange().getValues();
  let jugFound = false;
  for (let i = 1; i < jRows.length; i++) {
    if (String(jRows[i][0]).trim() === 'Reese' && fmtDate(jRows[i][1]) === date) {
      jugSheet.getRange(i+1,3,1,2).setValues([[89,now]]);
      jugFound = true; break;
    }
  }
  if (!jugFound) {
    const newRow = jugSheet.getLastRow() + 1;
    jugSheet.appendRow(['Reese', date, 89, now]);
    jugSheet.getRange(newRow, 2).setNumberFormat('@STRING@');
  }

  SpreadsheetApp.getUi().alert(
    '\u2705 Done!\n\nReese \u2192 June 4, 2026:\n\u2022 Wall-offs Right & Left (bs1, bs2) added \u2713\n\u2022 Juggling: 89 \u2713\n\nAll other existing checkmarks unchanged.'
  );
}

// ── ONE-TIME FIX: Baylor June 2 & 3 ────────────────────────────
// Sets all Ball Skills + Strength Training complete for both dates
// Select fixBaylorJune2And3 in dropdown → click Run
function fixBaylorJune2And3() {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const now = new Date().toISOString();
  const bs  = 'bs1,bs2,bs3,bs4,bs5,bs6';
  const st  = 'st1,st2,st3,st4,st5,st6';
  const dates = ['2026-06-02', '2026-06-03'];

  const dailySheet = getOrCreate(ss, 'U13_Daily',
    ['Player','Date','Ball Skills Checked (IDs)','Strength Checked (IDs)','Last Updated']);

  dates.forEach(function(date) {
    const dRows = dailySheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < dRows.length; i++) {
      if (String(dRows[i][0]).trim() === 'Baylor' && fmtDate(dRows[i][1]) === date) {
        // Union existing with full sets
        const existingBS = dRows[i][2] ? String(dRows[i][2]).split(',').filter(Boolean) : [];
        const existingST = dRows[i][3] ? String(dRows[i][3]).split(',').filter(Boolean) : [];
        const bsSet = new Set([...existingBS, 'bs1','bs2','bs3','bs4','bs5','bs6']);
        const stSet = new Set([...existingST, 'st1','st2','st3','st4','st5','st6']);
        dailySheet.getRange(i+1,3,1,3).setValues([[
          Array.from(bsSet).join(','),
          Array.from(stSet).join(','),
          now
        ]]);
        found = true;
        Logger.log('Updated Baylor row for ' + date);
        break;
      }
    }
    if (!found) {
      const newRow = dailySheet.getLastRow() + 1;
      dailySheet.appendRow(['Baylor', date, bs, st, now]);
      dailySheet.getRange(newRow, 2).setNumberFormat('@STRING@');
      Logger.log('Created Baylor row for ' + date);
    }
  });

  SpreadsheetApp.getUi().alert(
    '\u2705 Done!\n\nBaylor:\n\u2022 Jun 2, 2026: Ball Skills 6/6 \u2713  Strength 6/6 \u2713\n\u2022 Jun 3, 2026: Ball Skills 6/6 \u2713  Strength 6/6 \u2713\n\nExisting data merged (nothing overwritten).'
  );
}
