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
    var followRaw = sheet.getRange(hr + 1, COL.followup, lastRow - hr, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i];
      // שורה נחשבת ליד רק אם יש בה שם, טלפון או תאריך פנייה
      // (עמודת "נסגרה עסקה" מלאה FALSE מראש בהרבה שורות ריקות)
      if (!v[0].trim() && !v[1].trim() && !v[2].trim()) continue;
      leads.push({
        row: hr + 1 + i,
        created: v[0], name: v[1], phone: v[2], campaign: v[3], ad: v[4],
        eventDate: v[5], status: v[6], reason: v[7], notes: v[8],
        followup: followupText(followRaw[i][0], v[9]), closed: v[10].toUpperCase() === 'TRUE', value: v[11]
      });
    }
  }
  var statuses = dropdownValues(sheet, hr, COL.status, LIST_COL_STATUSES);
  var reasons = dropdownValues(sheet, hr, COL.reason, LIST_COL_REASONS);
  return {
    ok: true,
    leads: leads,
    statuses: statuses,
    reasons: reasons,
    statusColors: formatColors(sheet, COL.status, statuses),
    reasonColors: formatColors(sheet, COL.reason, reasons)
  };
}

// צבע לכל ערך, כפי שהוא נראה בשיטס: נקרא מכללי העיצוב המותנה שחלים על העמודה.
// כך שינוי צבע בשיטס מתעדכן בדף בלי לגעת בקוד. הכלל הראשון שמתאים לערך קובע (כמו בשיטס).
// צבעים שהוגדרו בתוך התפריט הנפתח עצמו (צ'יפים) לא נחשפים לקוד - ערך כזה פשוט לא יוחזר.
function formatColors(sheet, col, names) {
  var out = {};
  var rules;
  try { rules = sheet.getConditionalFormatRules(); } catch (err) { return out; }
  for (var r = 0; r < rules.length; r++) {
    var cond = rules[r].getBooleanCondition();
    if (!cond || !coversColumn(rules[r].getRanges(), col)) continue;
    var bg = cond.getBackground() || '';
    var fg = cond.getFontColor() || '';
    if (!bg && !fg) continue;
    var type = cond.getCriteriaType();
    var vals = cond.getCriteriaValues();
    var needle = String(vals && vals.length ? vals[0] : '');
    for (var n = 0; n < names.length; n++) {
      var name = names[n];
      if (out[name]) continue;
      if (ruleMatches(type, needle, name)) out[name] = { bg: bg, fg: fg };
    }
  }
  return out;
}

function coversColumn(ranges, col) {
  for (var i = 0; i < ranges.length; i++) {
    if (ranges[i].getColumn() <= col && col <= ranges[i].getLastColumn()) return true;
  }
  return false;
}

function ruleMatches(type, needle, name) {
  var T = SpreadsheetApp.BooleanCriteria;
  var a = norm(needle), b = norm(name);
  if (!a) return false;
  if (type === T.TEXT_EQUAL_TO) return a === b;
  if (type === T.TEXT_CONTAINS) return b.indexOf(a) !== -1;
  if (type === T.TEXT_STARTS_WITH) return b.indexOf(a) === 0;
  if (type === T.TEXT_ENDS_WITH) return b.length >= a.length && b.lastIndexOf(a) === b.length - a.length;
  if (type === T.CUSTOM_FORMULA) return needle.indexOf('"' + name + '"') !== -1;
  return false;
}

// הרשימה נלקחת מהתפריט הנפתח עצמו (אימות הנתונים על התא הראשון בעמודה),
// כי זה מה שאורטל רואה בשיטס - בין אם הערכים הוקלדו בתוך התפריט ובין אם הוא מפנה לטווח.
// עמודות N/O נשארות רק כגיבוי למקרה שאין אימות נתונים על העמודה.
function dropdownValues(sheet, hr, dataCol, listCol) {
  var rule = sheet.getRange(hr + 1, dataCol).getDataValidation();
  var out = [];
  if (rule) {
    var type = rule.getCriteriaType();
    var args = rule.getCriteriaValues();
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      out = cleanList(args[0]);
    } else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      var rows = args[0].getDisplayValues();
      var flat = [];
      for (var i = 0; i < rows.length; i++) flat = flat.concat(rows[i]);
      out = cleanList(flat);
    }
  }
  return out.length ? out : readList(sheet, hr, listCol);
}

function cleanList(values) {
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i] == null ? '' : values[i]).trim();
    if (v && out.indexOf(v) === -1) out.push(v);
  }
  return out;
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
    writeCell(sheet, row, key, fields[key]);
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
      writeCell(sheet, row, key, f[key]);
    }
  }
  return { ok: true, row: row };
}

// תאריך המעקב נקרא מהערך האמיתי של התא ולא מהתצוגה שלו, כדי שהשעה לא תלך לאיבוד
// כשהעמודה מעוצבת להציג תאריך בלבד. בלי שעה (חצות) מוחזר תאריך בלבד.
function followupText(raw, display) {
  if (!(raw instanceof Date) || isNaN(raw.getTime())) return display;
  var tz = Session.getScriptTimeZone();
  var hasTime = Utilities.formatDate(raw, tz, 'HH:mm') !== '00:00';
  return Utilities.formatDate(raw, tz, hasTime ? 'dd-MM-yyyy HH:mm' : 'dd-MM-yyyy');
}

// כותב ערך לתא. לתאריך מעקב עם שעה - מעצב את התא כך שהשעה תיראה גם בשיטס;
// תאריך בלי שעה מחזיר תא שהוצגה בו שעה לתצוגת תאריך בלבד.
function writeCell(sheet, row, key, val) {
  var range = sheet.getRange(row, COL[key]);
  var value = toCellValue(key, val);
  range.setValue(value);
  if (key !== 'followup' || !(value instanceof Date)) return;
  var hasTime = value.getHours() !== 0 || value.getMinutes() !== 0;
  if (hasTime) range.setNumberFormat('dd/MM/yyyy HH:mm');
  else if (/h/i.test(range.getNumberFormat())) range.setNumberFormat('dd/MM/yyyy');
}

// המרת ערכים מהדף לערכי תא: תאריכים כ-Date כדי שאימות הנתונים בשיטס יעבוד
function toCellValue(key, val) {
  if (key === 'closed') return val === true || val === 'true';
  if (key === 'value') {
    var n = Number(String(val).replace(/[^\d.\-]/g, ''));
    return isNaN(n) || String(val).trim() === '' ? '' : n;
  }
  if (key === 'eventDate' || key === 'followup') {
    // מהדף מגיע yyyy-mm-dd, ולתאריך מעקב עם שעה: yyyy-mm-ddTHH:MM
    var m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?$/.exec(String(val).trim());
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0));
    return String(val);
  }
  return String(val);
}

function norm(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
}
