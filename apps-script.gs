/**
 * On The Table - Lead API
 * קובץ זה יושב באותו פרויקט Apps Script של השיטס (לצד otb-onedit).
 * הוא משרת את דף ניהול הלידים של אורטל: קריאת לידים, עדכון שורה, הוספת ליד.
 *
 * התקנה (חד-פעמית):
 * 1. בשיטס: Extensions ← Apps Script ← קובץ חדש בשם lead-api ← להדביק את כל הקוד הזה
 * 2. לשנות את PIN לקוד שסוכם עם אורטל
 * 3. Deploy ← New deployment ← Web app ← Execute as: Me ← Who has access: Anyone
 * 4. להעתיק את כתובת ה-Web app ולהדביק אותה ב-index.html (API_URL)
 */

var PIN = 'CHANGE_ME'; // הקוד שאורטל תקליד בכניסה לדף
var SHEET_NAME = 'גיליון1';
var TZ = 'Asia/Jerusalem';

// עמודות (1-based): A תאריך פנייה, B שם, C טלפון, D קמפיין, E מודעה,
// F תאריך אירוע, G סטטוס, H סיבה, I הערות, J פולואפ, K נסגרה, L שווי
var COL = { created: 1, name: 2, phone: 3, campaign: 4, ad: 5, eventDate: 6,
            status: 7, reason: 8, notes: 9, followup: 10, closed: 11, value: 12 };
var LIST_COL_STATUSES = 14; // N
var LIST_COL_REASONS = 15;  // O

function doGet() {
  return ContentService.createTextOutput('OTB Lead API פעיל');
}

function doPost(e) {
  var res;
  try {
    var req = JSON.parse(e.postData.contents);
    if (String(req.pin || '') !== PIN) {
      res = { ok: false, error: 'bad_pin' };
    } else if (req.action === 'list') {
      res = listLeads();
    } else if (req.action === 'update') {
      res = withLock(function () { return updateLead(req); });
    } else if (req.action === 'add') {
      res = withLock(function () { return addLead(req); });
    } else {
      res = { ok: false, error: 'unknown_action' };
    }
  } catch (err) {
    res = { ok: false, error: 'server', message: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function getSheet() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('לא נמצא גיליון בשם ' + SHEET_NAME);
  return sheet;
}

// שורת הכותרות מזוהה לפי "שם מלא" בעמודה B - עמיד גם אם תתווסף שורה מעל
function headerRow(sheet) {
  var top = sheet.getRange(1, COL.name, Math.min(5, sheet.getLastRow()), 1).getDisplayValues();
  for (var i = 0; i < top.length; i++) {
    if (top[i][0].trim() === 'שם מלא') return i + 1;
  }
  throw new Error('לא נמצאה שורת כותרות (שם מלא בעמודה B)');
}

function listLeads() {
  var sheet = getSheet();
  var hr = headerRow(sheet);
  var lastRow = sheet.getLastRow();
  var leads = [];
  if (lastRow > hr) {
    var vals = sheet.getRange(hr + 1, 1, lastRow - hr, 12).getDisplayValues();
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i];
      // שורה נחשבת ליד רק אם יש בה שם, טלפון או תאריך פנייה
      // (עמודת "נסגרה עסקה" מלאה FALSE מראש בהרבה שורות ריקות)
      if (!v[0].trim() && !v[1].trim() && !v[2].trim()) continue;
      leads.push({
        row: hr + 1 + i,
        created: v[0], name: v[1], phone: v[2], campaign: v[3], ad: v[4],
        eventDate: v[5], status: v[6], reason: v[7], notes: v[8],
        followup: v[9], closed: v[10].toUpperCase() === 'TRUE', value: v[11]
      });
    }
  }
  return {
    ok: true,
    leads: leads,
    statuses: readList(sheet, hr, LIST_COL_STATUSES),
    reasons: readList(sheet, hr, LIST_COL_REASONS)
  };
}

function readList(sheet, hr, col) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= hr) return [];
  var vals = sheet.getRange(hr + 1, col, lastRow - hr, 1).getDisplayValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i][0].trim();
    if (!v) break; // הרשימה נגמרת בתא הריק הראשון
    out.push(v);
  }
  return out;
}

// עדכון ליד: לפני כתיבה מוודאים שהשורה עדיין מכילה את אותו ליד
// (הגנה מפני מיון/מחיקה בשיטס בין רענון הדף לבין השמירה)
function updateLead(req) {
  var sheet = getSheet();
  var row = Number(req.row);
  var hr = headerRow(sheet);
  if (!row || row <= hr || row > sheet.getLastRow()) return { ok: false, error: 'row_mismatch' };

  var current = sheet.getRange(row, COL.name, 1, 2).getDisplayValues()[0];
  var expect = req.expect || {};
  if (norm(current[0]) !== norm(expect.name) || norm(current[1]) !== norm(expect.phone)) {
    return { ok: false, error: 'row_mismatch' };
  }

  var fields = req.fields || {};
  var writable = ['eventDate', 'status', 'reason', 'notes', 'followup', 'closed', 'value'];
  for (var i = 0; i < writable.length; i++) {
    var key = writable[i];
    if (!(key in fields)) continue;
    sheet.getRange(row, COL[key]).setValue(toCellValue(key, fields[key]));
  }
  return { ok: true };
}

function addLead(req) {
  var sheet = getSheet();
  var hr = headerRow(sheet);
  var f = req.fields || {};
  if (!norm(f.name) && !norm(f.phone)) return { ok: false, error: 'empty_lead' };

  // השורה הריקה הראשונה אחרי הכותרות לפי עמודות A-C
  // (getLastRow מחזיר גם שורות שמכילות רק תיבות סימון ריקות ב-K)
  var lastRow = sheet.getLastRow();
  var row = hr + 1;
  if (lastRow > hr) {
    var vals = sheet.getRange(hr + 1, 1, lastRow - hr, 3).getDisplayValues();
    var i = 0;
    while (i < vals.length && (vals[i][0].trim() || vals[i][1].trim() || vals[i][2].trim())) i++;
    row = hr + 1 + i;
  }

  sheet.getRange(row, COL.created).setValue(Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm'));
  sheet.getRange(row, COL.name).setValue(String(f.name || ''));
  sheet.getRange(row, COL.phone).setValue(String(f.phone || ''));
  sheet.getRange(row, COL.campaign).setValue(String(f.campaign || 'הוזן ידנית'));
  var opt = ['eventDate', 'status', 'reason', 'notes', 'followup', 'closed', 'value'];
  for (var j = 0; j < opt.length; j++) {
    var key = opt[j];
    if (key in f && f[key] !== '' && f[key] !== null) {
      sheet.getRange(row, COL[key]).setValue(toCellValue(key, f[key]));
    }
  }
  return { ok: true, row: row };
}

// המרת ערכים מהדף לערכי תא: תאריכים כ-Date כדי שאימות הנתונים בשיטס יעבוד
function toCellValue(key, val) {
  if (key === 'closed') return val === true || val === 'true';
  if (key === 'value') {
    var n = Number(String(val).replace(/[^\d.\-]/g, ''));
    return isNaN(n) || String(val).trim() === '' ? '' : n;
  }
  if (key === 'eventDate' || key === 'followup') {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(val).trim()); // מהדף מגיע yyyy-mm-dd
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return String(val);
  }
  return String(val);
}

function norm(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
}
