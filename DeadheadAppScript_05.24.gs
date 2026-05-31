// ─────────────────────────────────────────────────────────────────
// Ron's Deadhead Listener — Apps Script
// Deploy as: Web App › Execute as Me › Anyone can access
// ─────────────────────────────────────────────────────────────────

// ── CONFIG ────────────────────────────────────────────────────────
const SHEET_ID      = '1Hl1iSUkBQbfjbzXpEl5DrpRZQhu_xnOfqgvgmoTMURk';
const USER_LIMIT    = 100;   // max registered users before waitlist kicks in
const RESERVED_NAMES = /^ron(\s+h)?$/i;  // block 'Ron' and 'Ron H' only

// Sheet names
const SHEET_MAIN      = 'Sheet1';
const SHEET_VISITORS  = 'Visitors';
const SHEET_WAITLIST  = 'Waitlist';
const SHEET_FAVORITES = 'Favorites';
const BASE_META_HEADERS = ['Date', 'Venue', 'City,State/Country'];
const ARCHIVE_ID_HEADER = 'Archive ID';
const SOURCE_LOCK_HEADER = 'Source Lock';

// ── TIMEZONE HELPER ───────────────────────────────────────────────
function nowPacific() {
  return Utilities.formatDate(
    new Date(),
    'America/Los_Angeles',
    'yyyy-MM-dd HH:mm:ss'
  );
}

// ── CORS HEADERS ──────────────────────────────────────────────────
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
}

function makeResponse(data, callback) {
  const json = JSON.stringify(data);
  const cb = (callback || '').toString().trim();
  const output = ContentService.createTextOutput(
    cb ? (cb + '(' + json + ');') : json
  ).setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  return output;
}

// ── ENTRY POINTS ──────────────────────────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  const callback = (e.parameter && e.parameter.callback) || '';

  if (action === 'getFavorites') {
    return makeResponse(getFavorites(), callback);
  }
  if (action === 'getModes') {
    return makeResponse(getModes(), callback);
  }
  if (action === 'getCatalog') {
    return makeResponse(getCatalog(e.parameter || {}), callback);
  }
  if (action === 'getVisitorStats') {
    return makeResponse(getVisitorStats(), callback);
  }
  if (action === 'getUserCount') {
    return makeResponse(getUserCount(), callback);
  }
  if (action === 'getMessage') {
    return makeResponse(getMessage(), callback);
  }

  return makeResponse({ status: 'ok' }, callback);
}

function doPost(e) {
  let data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return makeResponse({ error: 'bad json' });
  }

  const action = data.action || '';

  if (action === 'addUser')       return makeResponse(addUser(data));
  if (action === 'markAttended')  return makeResponse(markAttended(data));
  if (action === 'trackVisit')    return makeResponse(trackVisit(data));
  if (action === 'addFavorite')   return makeResponse(addFavorite(data));
  if (action === 'removeFavorite')return makeResponse(removeFavorite(data));
  if (action === 'touchFavoriteShow') return makeResponse(touchFavoriteShow(data));
  if (action === 'joinWaitlist')  return makeResponse(joinWaitlist(data));
  if (action === 'saveArchiveId') return makeResponse(saveArchiveId(data));

  return makeResponse({ error: 'unknown action: ' + action });
}

// ── MAIN SHEET METADATA HELPERS ──────────────────────────────────
function headerKey_(v) {
  return (v || '').toString().trim().toLowerCase().replace(/\s+/g, '');
}

function isArchiveIdHeader_(v) {
  const k = headerKey_(v);
  return k === 'archiveid' || k === 'archiveurl';
}

function isSourceLockHeader_(v) {
  const k = headerKey_(v);
  return k === 'sourcelock' || k === 'sourcelocks';
}

function isMetadataHeader_(v) {
  return isArchiveIdHeader_(v) || isSourceLockHeader_(v);
}

function firstUserColumn_(headerRow) {
  let c = 3;
  while (c < headerRow.length && isMetadataHeader_(headerRow[c])) c++;
  return c;
}

function modeKey_(label) {
  return (label || '').toString().trim().toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'mode';
}

function normalizeMode_(mode) {
  const label = (mode || '').toString().trim();
  return label || 'Grateful Dead';
}

function isDefaultMode_(mode) {
  return normalizeMode_(mode).toLowerCase() === 'grateful dead';
}

function sheetForMode_(ss, data) {
  const mode = normalizeMode_(data && data.mode);
  const requestedSheet = (data && data.sheetName || '').toString().trim();
  if (requestedSheet && (requestedSheet === SHEET_MAIN || requestedSheet.indexOf('MODE ') === 0)) {
    const sheet = ss.getSheetByName(requestedSheet);
    if (sheet) return sheet;
  }
  if (isDefaultMode_(mode)) return ss.getSheetByName(SHEET_MAIN);
  return ss.getSheetByName('MODE ' + mode);
}

function userStartColumn_(headerRow, nameRow) {
  for (let c = 3; c < headerRow.length; c++) {
    const h = (headerRow[c] || '').toString().trim();
    const n = (nameRow && nameRow[c] || '').toString().trim();
    if (/^Attended/i.test(h) || (n && !isMetadataHeader_(h) && h.toLowerCase() !== 'notes')) return c;
  }
  return firstUserColumn_(headerRow);
}

function ensureUserColumn_(sheet, user, withPin) {
  const lastCol = Math.max(4, sheet.getLastColumn());
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const nameRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const userStart = userStartColumn_(headerRow, nameRow);
  for (let c = userStart; c < nameRow.length; c++) {
    if (isMetadataHeader_(headerRow[c])) continue;
    if ((nameRow[c] || '').toString().trim() === user) return c;
  }
  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, newCol).setValue('Attended (Y)');
  sheet.getRange(2, newCol).setValue(user);
  if (withPin) sheet.getRange(3, newCol).setValue(withPin);
  return newCol - 1;
}

function getModes() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const main = ss.getSheetByName(SHEET_MAIN);
    const modes = [modeSummary_(main, 'Grateful Dead', SHEET_MAIN, true)];
    ss.getSheets().forEach(function(sheet) {
      const name = sheet.getName();
      if (name.indexOf('MODE ') === 0) {
        const label = name.replace(/^MODE\s+/, '').trim();
        if (label) modes.push(modeSummary_(sheet, label, name, false));
      }
    });
    const totalShows = modes.reduce(function(sum, m) { return sum + (m.showCount || 0); }, 0);
    return { modes: modes, totalShows: totalShows };
  } catch (e) {
    return { modes: [{ key: 'gd', label: 'Grateful Dead', sheetName: SHEET_MAIN, isDefault: true }], error: e.message };
  }
}

function sheetDateText_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'America/Los_Angeles', 'yyyy-MM-dd');
  }
  return (value || '').toString().trim();
}

function modeSummary_(sheet, label, sheetName, isDefault) {
  const summary = { key: isDefault ? 'gd' : modeKey_(label), label: label, sheetName: sheetName, isDefault: !!isDefault, showCount: 0, yearStart: '', yearEnd: '' };
  if (!sheet) return summary;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return summary;
  const startRow = isDefault ? 4 : 3;
  if (lastRow < startRow) return summary;
  const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).getValues();
  const seen = {};
  rows.forEach(function(row) {
    const date = sheetDateText_(row[0]);
    if (!/^\d{4}-\d{2}-\d{2}/.test(date)) return;
    const venue = (row[1] || '').toString().trim();
    const city = (row[2] || '').toString().trim();
    const key = showGroupKey_(date, venue, city);
    if (seen[key]) return;
    seen[key] = true;
    summary.showCount++;
    const y = date.substring(0, 4);
    if (!summary.yearStart || y < summary.yearStart) summary.yearStart = y;
    if (!summary.yearEnd || y > summary.yearEnd) summary.yearEnd = y;
  });
  return summary;
}

function getCatalog(params) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = sheetForMode_(ss, params || {});
    if (!sheet) return { error: 'mode_not_found' };
    const sheetName = sheet.getName();
    const isDefault = sheetName === SHEET_MAIN || isDefaultMode_(params && params.mode);
    const label = isDefault ? 'Grateful Dead' : sheetName.replace(/^MODE\s+/, '').trim();
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(4, sheet.getLastColumn());
    const values = lastRow ? sheet.getRange(1, 1, lastRow, lastCol).getValues() : [];
    if (values.length < 2) return { error: 'sheet_empty', mode: label, sheetName: sheetName };

    const headerRow = values[0] || [];
    const nameRow = values[1] || [];
    const pinRow = values[2] || [];
    let archiveCol = -1;
    let sourceLockCol = -1;
    for (let c = 3; c < headerRow.length; c++) {
      if (archiveCol < 0 && isArchiveIdHeader_(headerRow[c])) archiveCol = c;
      if (sourceLockCol < 0 && isSourceLockHeader_(headerRow[c])) sourceLockCol = c;
    }
    const userStart = userStartColumn_(headerRow, nameRow);
    const users = [];
    const pins = {};
    const userColumns = [];
    for (let c = userStart; c < nameRow.length; c++) {
      if (c === archiveCol || c === sourceLockCol || isMetadataHeader_(headerRow[c])) continue;
      const user = (nameRow[c] || '').toString().trim();
      if (!user || user === 'Name' || user === 'Notes') continue;
      users.push(user);
      userColumns.push(c);
      const pin = (pinRow[c] || '').toString().trim();
      if (pin) pins[user] = pin;
    }

    const startIdx = isDefault ? 3 : 2;
    const shows = [];
    const seen = {};
    for (let r = startIdx; r < values.length; r++) {
      const row = values[r] || [];
      const date = sheetDateText_(row[0]);
      if (!/^\d{4}-\d{2}-\d{2}/.test(date)) continue;
      const venue = (row[1] || '').toString().trim() || 'Unknown Venue';
      const city = (row[2] || '').toString().trim();
      const archiveId = archiveCol >= 0 ? (row[archiveCol] || '').toString().trim() : '';
      const sourceLock = sourceLockCol >= 0 ? (row[sourceLockCol] || '').toString().trim() : '';
      const attended = {};
      users.forEach(function(user, idx) {
        const v = row[userColumns[idx]];
        const s = (v || '').toString().trim().toLowerCase();
        attended[user] = s === 'y' || s === '1' || s === 'yes' || s === 'true';
      });
      const key = showGroupKey_(date, venue, city);
      if (!isDefault && seen[key]) {
        seen[key].altArchiveIds.push(archiveId);
        users.forEach(function(user) {
          if (attended[user]) seen[key].attended[user] = true;
        });
        continue;
      }
      const item = {
        row: r + 1,
        date: date,
        venue: venue,
        city: city,
        archiveId: archiveId,
        sourceLock: sourceLock,
        attended: attended,
        altArchiveIds: []
      };
      shows.push(item);
      if (!isDefault) seen[key] = item;
    }
    return {
      status: 'ok',
      mode: label,
      key: isDefault ? 'gd' : modeKey_(label),
      sheetName: sheetName,
      isDefault: isDefault,
      users: users,
      pins: pins,
      userLimit: USER_LIMIT,
      shows: shows
    };
  } catch (e) {
    return { error: e.message };
  }
}

function ensureMetadataColumn_(sheet, headerRow, wantedHeader) {
  const isWanted = wantedHeader === ARCHIVE_ID_HEADER ? isArchiveIdHeader_ : isSourceLockHeader_;
  for (let c = 0; c < headerRow.length; c++) {
    if (isWanted(headerRow[c])) return c;
  }
  const insertAt = firstUserColumn_(headerRow) + 1;
  sheet.insertColumnBefore(insertAt);
  sheet.getRange(1, insertAt).setValue(wantedHeader);
  headerRow.splice(insertAt - 1, 0, wantedHeader);
  return insertAt - 1;
}

function normalizeArchiveId_(v) {
  const s = (v || '').toString().trim();
  if (!s) return '';
  const m = s.match(/archive\.org\/(?:details|download)\/([^/?#]+)/i);
  return (m ? m[1] : s)
    .replace(/^https?:\/\/archive\.org\/details\//i, '')
    .replace(/[?#].*$/,'')
    .trim();
}

function plausibleArchiveId_(v) {
  const id = normalizeArchiveId_(v);
  return !!id && /^[A-Za-z0-9][A-Za-z0-9._-]{2,120}$/.test(id) && id.toUpperCase() !== 'NO_AUDIO';
}

function sameSheetDate_(cellValue, wantedDate) {
  if (!wantedDate) return false;
  if (cellValue instanceof Date) {
    return Utilities.formatDate(cellValue, 'America/Los_Angeles', 'yyyy-MM-dd') === wantedDate;
  }
  return (cellValue || '').toString().trim() === wantedDate;
}

function rowVenueMatches_(rowVenue, wantedVenue) {
  if (!wantedVenue) return true;
  return (rowVenue || '').toString().trim().toLowerCase() === wantedVenue.toString().trim().toLowerCase();
}

function showGroupPart_(value) {
  return (value || '').toString()
    .replace(/\s+/g, ' ')
    .replace(/\s*,\s*/g, ',')
    .trim()
    .toLowerCase();
}

function showGroupKey_(date, venue, city) {
  return [sheetDateText_(date), showGroupPart_(venue), showGroupPart_(city)].join('|');
}

// ── ADD USER ──────────────────────────────────────────────────────
function addUser(data) {
  const name = (data.name || '').trim();
  const pin  = (data.pin  || '').toString().trim();

  if (!name || name.length < 2) return { error: 'Name too short' };
  if (!pin  || pin.length !== 4) return { error: 'PIN must be 4 digits' };

  // Block Ron/Ronald
  if (RESERVED_NAMES.test(name)) {
    return { error: 'That name is reserved' };
  }

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_MAIN);
  // Sheet layout: row 1 = column headers, row 2 = names, row 3 = PINs, row 4+ = shows
  const rows = sheet.getRange(2, 1, 2, sheet.getLastColumn()).getValues();
  const nameRow = rows[0];  // row 2 (names)
  const pinRow  = rows[1];  // row 3 (PINs)
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const userStart = userStartColumn_(headerRow, nameRow);

  // Check duplicate
  for (let c = userStart; c < nameRow.length; c++) {
    if (isMetadataHeader_(headerRow[c])) continue;
    if ((nameRow[c] || '').toString().trim().toLowerCase() === name.toLowerCase()) {
      return { error: 'Name already taken' };
    }
  }

  // Count existing users after base + optional metadata columns.
  const userCount = Math.max(0, nameRow.filter((n, i) => i >= userStart && n && !isMetadataHeader_(headerRow[i])).length);

  if (userCount >= USER_LIMIT) {
    // Auto-add to waitlist instead
    joinWaitlist({ name: name, note: 'auto from addUser — limit reached' });
    return {
      error: 'limit_reached',
      message: 'The bus is full right now (' + USER_LIMIT + ' riders max). ' +
               'You\'ve been added to the waitlist — Ron will be in touch!'
    };
  }

  // Add new column — name in row 2, PIN in row 3
  const newCol = sheet.getLastColumn() + 1;
  sheet.getRange(2, newCol).setValue(name);
  sheet.getRange(3, newCol).setValue(pin);

  // Log in Visitors sheet
  trackVisit({ user: name, deviceId: data.deviceId || 'new-reg', isNew: true });

  return { status: 'ok', name: name };
}

// ── MARK ATTENDED ─────────────────────────────────────────────────
function markAttended(data) {
  const date     = data.date     || '';
  const venue    = data.venue    || '';
  const user     = data.user     || '';
  const attended = data.attended === true || data.attended === 'true';

  if (!date || !user) return { error: 'Missing date or user' };

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = sheetForMode_(ss, data);
  if (!sheet) return { error: 'Mode sheet not found' };
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 3 || lastCol < 1) return { error: 'Sheet empty' };

  const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headerRow = allData[0]; // row 1 = headers
  const nameRow = allData[1];  // row 2 = names (index 1)
  const userStart = userStartColumn_(headerRow, nameRow);

  // Find user column
  let userCol = -1;
  for (let c = userStart; c < nameRow.length; c++) {
    if (isMetadataHeader_(headerRow[c])) continue;
    if ((nameRow[c] || '').toString().trim() === user) {
      userCol = c;
      break;
    }
  }
  if (userCol < 0) {
    userCol = ensureUserColumn_(sheet, user, '');
    const fresh = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    allData.length = 0;
    fresh.forEach(function(row) { allData.push(row); });
  }

  // Find show row (prefer date + venue, fall back to date for older callers)
  let showRow = -1;
  let dateOnlyRow = -1;
  const startRow = isDefaultMode_(data.mode) ? 3 : 2;
  for (let r = startRow; r < allData.length; r++) {
    if (sameSheetDate_(allData[r][0], date)) {
      if (dateOnlyRow < 0) dateOnlyRow = r;
      if (rowVenueMatches_(allData[r][1], venue)) {
        showRow = r;
        break;
      }
    }
  }
  if (showRow < 0) showRow = dateOnlyRow;
  if (showRow < 0) return { error: 'Show not found: ' + date };

  sheet.getRange(showRow + 1, userCol + 1).setValue(attended ? 1 : '');
  return { status: 'ok' };
}

// ── SAVE ARCHIVE ID HINT ─────────────────────────────────────────
function saveArchiveId(data) {
  try {
    const date = (data.date || '').toString().trim();
    const venue = (data.venue || '').toString().trim();
    const archiveId = normalizeArchiveId_(data.archiveId || '');
    if (!date || !venue || !archiveId) return { status: 'missing_data' };
    if (!plausibleArchiveId_(archiveId)) return { status: 'bad_archive_id' };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = sheetForMode_(ss, data);
    if (!sheet) return { status: 'mode_not_found' };
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(3, sheet.getLastColumn());
    if (lastRow < 4) return { status: 'not_found' };

    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    ensureMetadataColumn_(sheet, headerRow, ARCHIVE_ID_HEADER);
    ensureMetadataColumn_(sheet, headerRow, SOURCE_LOCK_HEADER);
    const freshHeader = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let archiveCol = -1;
    let lockCol = -1;
    for (let c = 0; c < freshHeader.length; c++) {
      if (archiveCol < 0 && isArchiveIdHeader_(freshHeader[c])) archiveCol = c;
      if (lockCol < 0 && isSourceLockHeader_(freshHeader[c])) lockCol = c;
    }
    const rows = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();

    let rowIdx = -1;
    for (let r = 3; r < rows.length; r++) {
      if (sameSheetDate_(rows[r][0], date) && rowVenueMatches_(rows[r][1], venue)) {
        rowIdx = r;
        break;
      }
    }
    if (rowIdx < 0) return { status: 'not_found' };

    const existing = (rows[rowIdx][archiveCol] || '').toString().trim();
    const lock = (rows[rowIdx][lockCol] || '').toString().trim().toUpperCase();
    if (lock === 'LOCK') return { status: 'locked' };
    if (existing) return { status: 'already_set' };

    sheet.getRange(rowIdx + 1, archiveCol + 1).setValue(archiveId);
    return {
      status: 'ok',
      archiveId: archiveId,
      trackCount: data.trackCount || '',
      source: data.source || ''
    };
  } catch (e) {
    return { status: 'error', error: e.message };
  }
}

// ── TRACK VISIT ───────────────────────────────────────────────────
function ensureVisitorsSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_VISITORS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_VISITORS);
    sheet.appendRow(['DeviceID', 'User', 'FirstSeen (PT)', 'LastSeen (PT)', 'Visits', 'IsNew', 'LastDeviceID']);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const headers = sheet.getRange(1, 1, 1, Math.max(7, sheet.getLastColumn())).getValues()[0];
  const desired = ['DeviceID', 'User', 'FirstSeen (PT)', 'LastSeen (PT)', 'Visits', 'IsNew', 'LastDeviceID'];
  for (let i = 0; i < desired.length; i++) {
    if (!headers[i]) sheet.getRange(1, i + 1).setValue(desired[i]);
  }
  return sheet;
}

function visitorKey_(user, deviceId) {
  const u = (user || 'guest').toString().trim();
  if (u && u.toLowerCase() !== 'guest') {
    return 'user:' + u.toLowerCase().replace(/\s+/g, ' ');
  }
  return deviceId || 'guest:unknown';
}

function trackVisit(data) {
  const user     = (data.user     || 'guest').toString().trim() || 'guest';
  const deviceId = (data.deviceId || data.guestId || '').toString().trim();
  const browserId = (data.browserId || data.guestId || data.deviceId || '').toString().trim();
  const isNew    = data.isNew === true;
  const key      = deviceId || visitorKey_(user, browserId);

  try {
    const sheet = ensureVisitorsSheet_();
    const now = nowPacific();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      sheet.appendRow([key, user, now, now, 1, isNew ? 'new' : '', browserId]);
      return { status: 'ok', action: 'created', key: key };
    }

    const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(7, sheet.getLastColumn())).getValues();
    for (let r = 0; r < rows.length; r++) {
      if ((rows[r][0] || '').toString() === key) {
        const sheetRow = r + 2;
        const visits = parseInt(rows[r][4] || '0', 10) + 1;
        sheet.getRange(sheetRow, 2).setValue(user);
        sheet.getRange(sheetRow, 4).setValue(now);
        sheet.getRange(sheetRow, 5).setValue(visits);
        if (isNew && !rows[r][5]) sheet.getRange(sheetRow, 6).setValue('new');
        sheet.getRange(sheetRow, 7).setValue(browserId);
        return { status: 'ok', action: 'updated', visits: visits, key: key };
      }
    }

    sheet.appendRow([key, user, now, now, 1, isNew ? 'new' : '', browserId]);
    return { status: 'ok', action: 'created', key: key };

  } catch (e) {
    return { error: e.message };
  }
}

// ── VISITOR STATS ─────────────────────────────────────────────────
function getVisitorStats() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_VISITORS);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { total: 0, today: 0, registered: 0 };
    }

    const todayPT = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

    let total = 0, todayCount = 0, registered = 0;
    data.forEach(function(row) {
      if (!row[0]) return;
      total++;
      const lastSeen = (row[3] || '').toString();
      if (lastSeen.startsWith(todayPT)) todayCount++;
      const user = (row[1] || '').toString().trim();
      if (user && user !== 'guest') registered++;
    });

    return { total: total, today: todayCount, registered: registered };
  } catch (e) {
    return { error: e.message };
  }
}

// ── GET USER COUNT ────────────────────────────────────────────────
function getUserCount() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_MAIN);
    const nameRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];  // row 2 = names
    const count = nameRow.filter(function(n, i) { return i >= 3 && n; }).length;
    return { count: count, limit: USER_LIMIT, spotsLeft: Math.max(0, USER_LIMIT - count) };
  } catch (e) {
    return { error: e.message };
  }
}

// ── ADMIN MESSAGE ────────────────────────────────────────────────
// Edit the message in the sheet: put message ID in Sheet1 cell B3,
// and message text in cell C3. Leave B3 empty to show no message.
// ID must change each time you want users to see a new message.
function getMessage() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_MAIN);
    // Row 3 (index 2), col B (index 1) = message ID
    // Row 3 (index 2), col C (index 2) = message text
    const id   = (sheet.getRange(3, 2).getValue() || '').toString().trim();
    const text = (sheet.getRange(3, 3).getValue() || '').toString().trim();
    if (!id || !text) return { id: '', text: '' };
    return { id: id, text: text };
  } catch (e) {
    return { id: '', text: '' };
  }
}

// ── WAITLIST ──────────────────────────────────────────────────────
function joinWaitlist(data) {
  const name  = (data.name  || '').toString().trim();
  const email = (data.email || '').toString().trim();
  const note  = (data.note  || '').toString().trim();

  if (!name) return { error: 'Name required' };

  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let sheet   = ss.getSheetByName(SHEET_WAITLIST);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_WAITLIST);
      sheet.appendRow(['Name', 'Email', 'Requested (PT)', 'Note', 'Status']);
      sheet.setFrozenRows(1);
    }

    // Check if already on waitlist
    if (sheet.getLastRow() > 1) {
      const existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let r = 0; r < existing.length; r++) {
        if ((existing[r][0] || '').toString().toLowerCase() === name.toLowerCase()) {
          return { status: 'already_listed', message: 'Already on the waitlist!' };
        }
      }
    }

    sheet.appendRow([name, email, nowPacific(), note, 'pending']);
    return { status: 'ok', message: 'Added to waitlist' };

  } catch (e) {
    return { error: e.message };
  }
}

// ── FAVORITES ─────────────────────────────────────────────────────
function ensureFavoritesSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_FAVORITES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FAVORITES);
    sheet.appendRow(['User', 'ShowID', 'Date', 'Venue', 'AddedAt', 'LastPlayedAt', 'Mode']);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const headers = sheet.getRange(1, 1, 1, Math.max(7, sheet.getLastColumn())).getValues()[0];
  const desired = ['User', 'ShowID', 'Date', 'Venue', 'AddedAt', 'LastPlayedAt', 'Mode'];
  for (let i = 0; i < desired.length; i++) {
    if (!headers[i]) sheet.getRange(1, i + 1).setValue(desired[i]);
  }
  return sheet;
}

function favoriteDateString_(value, fallback) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'America/Los_Angeles', 'yyyy-MM-dd HH:mm:ss');
  }
  const s = (value || '').toString().trim();
  return s || fallback || '';
}

function favoriteDateKey_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'America/Los_Angeles', 'yyyy-MM-dd');
  }
  const s = (value || '').toString().trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + String(Number(m[2])).padStart(2, '0') + '-' + String(Number(m[3])).padStart(2, '0');
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + String(Number(m[1])).padStart(2, '0') + '-' + String(Number(m[2])).padStart(2, '0');
  return s.slice(0, 10);
}

function getFavorites() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FAVORITES);
    if (!sheet || sheet.getLastRow() <= 1) return { favorites: [] };

    const colCount = Math.max(7, sheet.getLastColumn());
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, colCount).getValues();
    const favorites = rows
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        const addedAt = favoriteDateString_(r[4], '');
        const lastPlayedAt = favoriteDateString_(r[5], addedAt);
        return {
          user: r[0],
          showId: r[1],
          date: favoriteDateString_(r[2], ''),
          venue: r[3],
          addedAt: addedAt,
          lastPlayedAt: lastPlayedAt,
          mode: normalizeMode_(r[6])
        };
      });
    return { favorites: favorites };
  } catch (e) {
    return { error: e.message };
  }
}

function addFavorite(data) {
  const user   = (data.user   || '').trim();
  const showId = (data.showId || '').trim();
  const date   = (data.date   || '').trim();
  const venue  = (data.venue  || '').trim();
  const mode   = normalizeMode_(data.mode);
  if (!user || !showId) return { error: 'Missing user or showId' };

  try {
    const sheet = ensureFavoritesSheet_();
    const now = nowPacific();

    // Prevent duplicates, and backfill dates for older rows when found.
    if (sheet.getLastRow() > 1) {
      const existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(7, sheet.getLastColumn())).getValues();
      for (let r = 0; r < existing.length; r++) {
        if (existing[r][0] === user && existing[r][1] === showId && normalizeMode_(existing[r][6]) === mode) {
          const rowNum = r + 2;
          if (!existing[r][4]) sheet.getRange(rowNum, 5).setValue(now);
          if (!existing[r][5]) sheet.getRange(rowNum, 6).setValue(now);
          if (!existing[r][6]) sheet.getRange(rowNum, 7).setValue(mode);
          return { status: 'already_favorited' };
        }
      }
    }

    sheet.appendRow([user, showId, date, venue, now, now, mode]);
    return { status: 'ok', addedAt: now, lastPlayedAt: now };
  } catch (e) {
    return { error: e.message };
  }
}

function removeFavorite(data) {
  const user   = (data.user   || '').trim();
  const showId = (data.showId || '').trim();
  const mode   = normalizeMode_(data.mode);
  if (!user || !showId) return { error: 'Missing user or showId' };

  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_FAVORITES);
    if (!sheet || sheet.getLastRow() <= 1) return { status: 'not_found' };

    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(7, sheet.getLastColumn())).getValues();
    for (let r = rows.length - 1; r >= 0; r--) {
      if (rows[r][0] === user && rows[r][1] === showId && normalizeMode_(rows[r][6]) === mode) {
        sheet.deleteRow(r + 2);
        return { status: 'ok' };
      }
    }
    return { status: 'not_found' };
  } catch (e) {
    return { error: e.message };
  }
}

function touchFavoriteShow(data) {
  const showId = (data.showId || '').trim();
  const date = (data.date || '').trim();
  const venue = (data.venue || '').trim();
  const mode = normalizeMode_(data.mode);
  if (!showId && (!date || !venue)) return { error: 'Missing showId or date/venue' };

  try {
    const sheet = ensureFavoritesSheet_();
    if (sheet.getLastRow() <= 1) return { status: 'not_found', updated: 0 };

    const now = nowPacific();
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(7, sheet.getLastColumn())).getValues();
    let updated = 0;
    for (let r = 0; r < rows.length; r++) {
      if (normalizeMode_(rows[r][6]) === mode && (rows[r][1] === showId || (date && venue && favoriteDateKey_(rows[r][2]) === favoriteDateKey_(date) && String(rows[r][3] || '').trim() === venue))) {
        const rowNum = r + 2;
        if (!rows[r][4]) sheet.getRange(rowNum, 5).setValue(now);
        sheet.getRange(rowNum, 6).setValue(now);
        if (!rows[r][6]) sheet.getRange(rowNum, 7).setValue(mode);
        updated++;
      }
    }
    return { status: updated ? 'ok' : 'not_found', updated: updated, lastPlayedAt: now };
  } catch (e) {
    return { error: e.message };
  }
}
