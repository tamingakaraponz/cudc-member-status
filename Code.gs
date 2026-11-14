// Backend for Index.html. Apps Script runs this file on Google's servers.
// Only permitted display fields are returned to the browser; sheet rows stay here.
const CUDC_CONFIG = Object.freeze({
  spreadsheetId: '1quLuXIcS5aIvmXqn5U4KEx5s61kqrFz0DMgDFzF7BnA',
  sheetId: 0,
  sheetName: 'Attendance',
  maxQueryLength: 30
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CUDC • Member Status')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Return only the displayed update time, without exposing other cells.
function getLastUpdated() {
  try {
    const sheet = SpreadsheetApp.openById(CUDC_CONFIG.spreadsheetId).getSheetByName('ชีต4');
    if (!sheet) return { ok: false };
    const value = sheet.getRange('B10').getDisplayValue().trim();
    return { ok: true, value: value || null };
  } catch (error) { return { ok: false }; }
}

// Run manually in the Apps Script editor to check authorization and headers.
// Private: unavailable through google.script.run. No personal data is logged.
function testConnection_() {
  readAttendance_();
  console.log('Sheet connection and headers verified.');
}

// Only matching names are returned for ambiguous searches.
function searchMembers(query) {
  return safeResponse_(function () {
    const key = queryKey_(query);
    const data = readAttendance_(); // Fresh read on every search; no cache.
    const found = matchingRows_(data, key);
    if (!found.length) return { kind: 'empty' };
    if (found.length === 1) return { kind: 'single', member: memberView_(data, found[0]) };
    return { kind: 'matches', matches: found.map(function (index) {
      const row = data.rows[index];
      return { row: index + 2, name: row[data.columns.name].trim(), nickname: row[data.columns.nickname].trim() };
    }) };
  });
}

// Selection is checked against another fresh read, including its original query.
function getMember(query, selection) {
  return safeResponse_(function () {
    const key = queryKey_(query);
    if (!selection || !Number.isInteger(selection.row) || selection.row < 2 ||
        typeof selection.name !== 'string' || typeof selection.nickname !== 'string') {
      throw new Error('INVALID_SELECTION');
    }
    const data = readAttendance_();
    const index = selection.row - 2;
    const row = data.rows[index];
    if (!row || row[data.columns.name].trim() !== selection.name || row[data.columns.nickname].trim() !== selection.nickname ||
        !matchingRows_(data, key).includes(index)) throw new Error('STALE_SELECTION');
    return { kind: 'single', member: memberView_(data, index) };
  });
}

// Helpers ending in _ are not callable through google.script.run.
function safeResponse_(action) {
  try { return { ok: true, result: action() }; }
  catch (error) {
    const allowed = ['INVALID_QUERY', 'INVALID_SELECTION', 'STALE_SELECTION', 'SHEET_STRUCTURE', 'CHOREO_LOOKUP'];
    return { ok: false, code: allowed.includes(error.message) ? error.message : 'SERVICE_UNAVAILABLE' };
  }
}

function normalizeName_(value) {
  return String(value).normalize('NFC').toLowerCase()
    .replace(/[\s\u200B-\u200D\uFEFF*]+/g, '');
}

function queryKey_(query) {
  if (typeof query !== 'string' || query.length > CUDC_CONFIG.maxQueryLength) throw new Error('INVALID_QUERY');
  const key = identifierDigits_(query);
  if (!/^\d{10}$/.test(key) && !/^66\d{9}$/.test(key)) throw new Error('INVALID_QUERY');
  return key;
}

function identifierDigits_(value) {
  const text = String(value).replace(/[๐-๙]/g, function (digit) { return String(digit.charCodeAt(0) - 0x0E50); });
  if (!/^\+?[\d\s()\-\u200B-\u200D\uFEFF]+$/.test(text.trim())) return '';
  return text.replace(/[\s()\-\u200B-\u200D\uFEFF+]/g, '');
}

function phoneKey_(value) {
  let digits = identifierDigits_(value);
  if (/^66\d{9}$/.test(digits)) digits = '0' + digits.slice(2);
  return /^0\d{9}$/.test(digits) ? digits : '';
}

function readAttendance_() {
  const book = SpreadsheetApp.openById(CUDC_CONFIG.spreadsheetId);
  const sheet = book.getSheets().find(function (item) { return item.getSheetId() === CUDC_CONFIG.sheetId; });
  if (!sheet || sheet.getName() !== CUDC_CONFIG.sheetName || sheet.getLastColumn() < 9 || sheet.getLastRow() < 1) {
    throw new Error('SHEET_STRUCTURE');
  }
  // One rectangular read minimizes service calls and keeps each response together.
  // E/F can exist in server memory but are NEVER copied into a response or log.
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
  const headers = values.shift();
  // Locate fields by their header: columns can move as the sheet evolves.
  const expected = { status: 'Criteria', name: 'ชื่อ - สกุล', nickname: 'ชื่อเล่น',
    studentId: 'รหัสนิสิต', phone: 'เบอร์โทร', dancer: 'Dancer', staff: 'Staff' };
  const columns = {};
  Object.keys(expected).forEach(function (key) {
    const indexes = [];
    headers.forEach(function (header, index) {
      if (normalizeName_(header) === normalizeName_(expected[key])) indexes.push(index);
    });
    if (indexes.length !== 1) throw new Error('SHEET_STRUCTURE');
    columns[key] = indexes[0];
  });
  const workshopColumns = headers.map(function (header, index) { return normalizeName_(header) === 'workshop' ? index : -1; }).filter(function (index) { return index >= 0; });
  if (workshopColumns.length > 1) throw new Error('SHEET_STRUCTURE');
  columns.workshop = workshopColumns.length ? workshopColumns[0] : -1;
  if (columns.staff <= columns.dancer) throw new Error('SHEET_STRUCTURE');
  return { headers: headers, rows: values, columns: columns };
}

function matchingRows_(data, key) {
  const found = [];
  const isStudentId = /^6\d{9}$/.test(key);
  const phone = isStudentId ? '' : phoneKey_(key);
  data.rows.forEach(function (row, index) {
    if (!row[data.columns.name].trim()) return; // Skip blank and footer/note rows.
    const studentId = identifierDigits_(row[data.columns.studentId]);
    if ((isStudentId && studentId === key) || (phone && phoneKey_(row[data.columns.phone]) === phone)) found.push(index);
  });
  return found;
}

function countValue_(value) {
  const text = value.trim();
  if (!/^\d+$/.test(text)) return null;
  const number = Number(text);
  return Number.isSafeInteger(number) ? number : null;
}

function activityTitle_(header, column) {
  // English interface; original member names remain in their source language.
  const aliases = { 'งานบอล?': 'Football event (?)' };
  const title = header.trim();
  return Object.prototype.hasOwnProperty.call(aliases, title) ? aliases[title] :
    (/[\u0E00-\u0E7F]/.test(title) ? 'Activity ' + column : title);
}

function choreoView_(studentId) {
  const id = identifierDigits_(studentId);
  if (!/^\d{10}$/.test(id)) throw new Error('CHOREO_LOOKUP');
  const sheet = SpreadsheetApp.openById(CUDC_CONFIG.spreadsheetId).getSheets()
    .find(function (item) { return item.getName() === 'Choreo'; });
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 5) throw new Error('CHOREO_LOOKUP');
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
  const headers = values.shift();
  const idColumns = headers.map(function (h, i) { return normalizeName_(h) === normalizeName_('รหัสนิสิต') ? i : -1; }).filter(function (i) { return i >= 0; });
  if (normalizeName_(headers[0]) !== 'criteria' || idColumns.length !== 1) throw new Error('CHOREO_LOOKUP');
  const matches = values.filter(function (row) { return identifierDigits_(row[idColumns[0]]) === id; });
  if (matches.length !== 1) throw new Error('CHOREO_LOOKUP');
  const status = matches[0][0].trim();
  if (!/^(qualified|unqualified)$/i.test(status)) throw new Error('CHOREO_LOOKUP');
  const countColumns = headers.map(function (h, i) { return normalizeName_(h) === normalizeName_('จำนวนโชว์') ? i : -1; }).filter(function (i) { return i >= 0; });
  if (countColumns.length !== 1) throw new Error('CHOREO_LOOKUP');
  const row = matches[0], activities = [];
  let hasUnrecognizedRole = false;
  for (let column = countColumns[0] + 1; column < headers.length; column++) {
    const value = row[column].trim().toLowerCase();
    if (!value || value === 'false') continue;
    if (value !== 'true') { hasUnrecognizedRole = true; continue; }
    if (!headers[column].trim()) throw new Error('CHOREO_LOOKUP');
    activities.push({ name: activityTitle_(headers[column], column + 1), role: 'Choreo' });
  }
  return { status: status, choreo: countValue_(row[countColumns[0]]), activities: activities, hasUnrecognizedRole: hasUnrecognizedRole };
}

function memberView_(data, index) {
  const row = data.rows[index];
  const criteria = row[data.columns.status].trim();
  if (criteria.toLowerCase() === 'choreo') {
    const choreo = choreoView_(row[data.columns.studentId]);
    return { name: row[data.columns.name].trim(), nickname: row[data.columns.nickname].trim(),
      status: choreo.status, choreo: choreo.choreo, activities: choreo.activities, hasUnrecognizedRole: choreo.hasUnrecognizedRole, isChoreo: true };
  }
  const activities = [];
  let hasUnrecognizedRole = false;
  for (let column = data.columns.staff + 1; column < data.headers.length; column++) {
    const value = row[column].trim();
    if (!value) continue; // Missing record does not establish absence.
    const roleKey = value.toLowerCase();
    if (roleKey !== 'dancer' && roleKey !== 'staff' && roleKey !== 'workshop') { hasUnrecognizedRole = true; continue; }
    if (!data.headers[column].trim()) throw new Error('SHEET_STRUCTURE');
    activities.push({ name: activityTitle_(data.headers[column], column + 1), role: {dancer:'Dancer',staff:'Staff',workshop:'Workshop'}[roleKey] });
  }
  const status = criteria || 'Not recorded';
  // Explicit allowlist: no source row, title, student ID, phone, or other columns.
  return { name: row[data.columns.name].trim(), nickname: row[data.columns.nickname].trim(), status: status,
    dancer: countValue_(row[data.columns.dancer]), staff: countValue_(row[data.columns.staff]), workshop: data.columns.workshop < 0 ? null : countValue_(row[data.columns.workshop]),
    activities: activities, hasUnrecognizedRole: hasUnrecognizedRole };
}
