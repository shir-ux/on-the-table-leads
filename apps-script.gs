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
var TASKS_SHEET_NAME = 'משימות'; // לשונית נפרדת, שורה לכל משימה. נוצרת אוטומטית במשימה הראשונה
var TZ = 'Asia/Jerusalem';

// העמודות מזוהות לפי שם הכותרת שלהן ולא לפי מיקום, כדי שאותו קוד יעבוד גם בגיליון
// שבו העמודות מסודרות אחרת. לכל שדה - שמות הכותרת המקובלים. חובה: שם, טלפון, סטטוס.
// שדה שהכותרת שלו לא קיימת בגיליון פשוט לא נקרא ולא נכתב.
var HEADERS = {
  created: ['תאריך פנייה', 'תאריך פניה'],
  name: ['שם מלא', 'שם'],
  phone: ['טלפון'],
  campaign: ['קמפיין'],
  ad: ['מודעה'],
  audience: ['שם הקהל', 'קהל'],
  status: ['סטטוס'],
  reason: ['סיבת אי-סגירה', 'סיבת אי סגירה'],
  notes: ['דיווח והערות', 'הערות'],
  followup: ['תאריך פולואפ', 'תאריך מעקב'],
  closed: ['נסגרה עסקה'],
  value: ['שווי הזמנה'],
  updated: ['עדכון אחרון', 'תאריך עדכון אחרון']
};
var REQUIRED_FIELDS = ['name', 'phone', 'status'];
var COL = {};        // שדה -> מספר עמודה (1-based). מתמלא ב-readLayout בכל בקשה
var HEADER_ROW = 0;
var HISTORY_SHEET_NAME = 'היסטוריה'; // שורה לכל אירוע בליד. נוצרת אוטומטית באירוע הראשון
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
    } else if (req.action === 'history') {
      res = leadHistory(req);
    } else if (req.action === 'tasksDone') {
      res = doneTasks(req);
    } else if (req.action === 'taskAdd') {
      res = withLock(function () { return addTask(req); });
    } else if (req.action === 'taskDone') {
      res = withLock(function () { return setTaskDone(req); });
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
  readLayout(sheet);
  return sheet;
}

// שורת הכותרות מזוהה לפי "שם מלא" בעמודה B - עמיד גם אם תתווסף שורה מעל
function headerRow(sheet) {
  return HEADER_ROW;
}

// שורת הכותרות היא הראשונה (מתוך 5 העליונות) שיש בה גם "שם" וגם "טלפון".
function readLayout(sheet) {
  var rows = Math.min(5, sheet.getLastRow());
  var width = sheet.getLastColumn();
  if (!rows || !width) throw new Error('הגיליון ריק');
  var top = sheet.getRange(1, 1, rows, width).getDisplayValues();
  for (var r = 0; r < top.length; r++) {
    var found = {};
    for (var c = 0; c < width; c++) {
      var key = fieldForHeader(top[r][c]);
      if (key && !found[key]) found[key] = c + 1;
    }
    if (found.name && found.phone) {
      var missing = [];
      for (var i = 0; i < REQUIRED_FIELDS.length; i++) {
        if (!found[REQUIRED_FIELDS[i]]) missing.push(HEADERS[REQUIRED_FIELDS[i]][0]);
      }
      if (missing.length) throw new Error('חסרה בגיליון עמודה בשם: ' + missing.join(', '));
      COL = found; HEADER_ROW = r + 1;
      return;
    }
  }
  throw new Error('לא נמצאה שורת כותרות (עמודות "שם מלא" ו"טלפון")');
}

function fieldForHeader(text) {
  var t = norm(text).replace(/[₪\s]+$/, '');
  if (!t) return '';
  for (var key in HEADERS) {
    if (HEADERS[key].indexOf(t) !== -1) return key;
  }
  if (t.indexOf('שווי הזמנה') === 0) return 'value';
  return '';
}

function cell(v, key) { return COL[key] ? v[COL[key] - 1] : ''; }

function listLeads() {
  var sheet = getSheet();
  var hr = headerRow(sheet);
  var lastRow = sheet.getLastRow();
  var leads = [];
  if (lastRow > hr) {
    var vals = sheet.getRange(hr + 1, 1, lastRow - hr, sheet.getLastColumn()).getDisplayValues();
    var followRaw = COL.followup ? sheet.getRange(hr + 1, COL.followup, lastRow - hr, 1).getValues() : null;
    var updatedRaw = COL.updated ? sheet.getRange(hr + 1, COL.updated, lastRow - hr, 1).getValues() : null;
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i];
      // שורה נחשבת ליד רק אם יש בה שם, טלפון או תאריך פנייה
      // (עמודת "נסגרה עסקה" מלאה FALSE מראש בהרבה שורות ריקות)
      if (!cell(v, 'created').trim() && !cell(v, 'name').trim() && !cell(v, 'phone').trim()) continue;
      leads.push({
        row: hr + 1 + i,
        created: cell(v, 'created'), name: cell(v, 'name'), phone: cell(v, 'phone'),
        campaign: cell(v, 'campaign'), ad: cell(v, 'ad'), audience: cell(v, 'audience'),
        status: cell(v, 'status'), reason: cell(v, 'reason'), notes: cell(v, 'notes'),
        followup: followRaw ? followupText(followRaw[i][0], cell(v, 'followup')) : '',
        closed: cell(v, 'closed').toUpperCase() === 'TRUE', value: cell(v, 'value'),
        updated: updatedRaw ? followupText(updatedRaw[i][0], cell(v, 'updated')) : ''
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
    tasks: listTasks(),
    statusColors: formatColors(sheet, COL.status, statuses),
    reasonColors: formatColors(sheet, COL.reason, reasons)
  };
}

// צבע לכל ערך, כפי שהוא נראה בשיטס: נקרא מכללי העיצוב המותנה שחלים על העמודה.
// כך שינוי צבע בשיטס מתעדכן בדף בלי לגעת בקוד. הכלל הראשון שמתאים לערך קובע (כמו בשיטס).
// צבעים שהוגדרו בתוך התפריט הנפתח עצמו (צ'יפים) לא נחשפים לקוד - ערך כזה פשוט לא יוחזר.
function formatColors(sheet, col, names) {
  var out = {};
  if (!col) return out;
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
  if (!dataCol) return [];
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

  var before = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var expect = req.expect || {};
  if (norm(cell(before, 'name')) !== norm(expect.name) || norm(cell(before, 'phone')) !== norm(expect.phone)) {
    return { ok: false, error: 'row_mismatch' };
  }

  var fields = req.fields || {};
  var writable = ['status', 'reason', 'notes', 'followup', 'closed', 'value'];
  for (var i = 0; i < writable.length; i++) {
    var key = writable[i];
    if (!(key in fields)) continue;
    writeCell(sheet, row, key, fields[key]);
  }
  logChanges(cell(before, 'name'), cell(before, 'phone'), before, fields);
  stampUpdated(sheet, row);
  return { ok: true };
}

// חותמת "עדכון אחרון": תאריך ושעה של השינוי האחרון בליד, בעמודה משלה בגיליון הלידים.
// כך אפשר למיין ולסנן לפיה גם בשיטס עצמו. אם אין עמודה בשם הזה - לא נכתב כלום.
function stampUpdated(sheet, row) {
  if (!COL.updated) return;
  sheet.getRange(row, COL.updated).setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm');
}

// משימה שנוספה או בוצעה נחשבת גם היא עדכון של הליד. הליד מאותר לפי טלפון.
function stampUpdatedByPhone(phone) {
  try {
    var key = phoneKey(phone);
    if (!key) return;
    var sheet = getSheet();
    if (!COL.updated || sheet.getLastRow() <= HEADER_ROW) return;
    var phones = sheet.getRange(HEADER_ROW + 1, COL.phone, sheet.getLastRow() - HEADER_ROW, 1).getDisplayValues();
    for (var i = 0; i < phones.length; i++) {
      if (phoneKey(phones[i][0]) === key) { stampUpdated(sheet, HEADER_ROW + 1 + i); return; }
    }
  } catch (err) {}
}

// רושם בהיסטוריה את מה שהשתנה. כישלון ברישום לא מכשיל את השמירה עצמה.
function logChanges(name, phone, before, fields) {
  try {
    if ('status' in fields && norm(fields.status) !== norm(cell(before, 'status'))) {
      var from = norm(cell(before, 'status'));
      logEvent(name, phone, 'סטטוס', (from ? 'מ"' + from + '" ל' : '') + '"' + norm(fields.status) + '"');
    }
    if ('reason' in fields && norm(fields.reason)) logEvent(name, phone, 'סיבת אי-סגירה', norm(fields.reason));
    if ('value' in fields && norm(fields.value)) logEvent(name, phone, 'שווי הזמנה', norm(fields.value) + ' ₪');
    if ('followup' in fields) {
      var due = toCellValue('followup', fields.followup);
      logEvent(name, phone, 'מעקב', due instanceof Date ? 'נקבע ל-' + followupText(due, '') : 'תאריך המעקב נוקה');
    }
    if ('notes' in fields && norm(fields.notes) && norm(fields.notes) !== norm(cell(before, 'notes'))) {
      logEvent(name, phone, 'הערה', String(fields.notes).trim());
    }
  } catch (err) {}
}

function addLead(req) {
  var sheet = getSheet();
  var hr = headerRow(sheet);
  var f = req.fields || {};
  if (!norm(f.name) && !norm(f.phone)) return { ok: false, error: 'empty_lead' };

  // השורה הריקה הראשונה אחרי הכותרות לפי תאריך פנייה / שם / טלפון
  // (getLastRow מחזיר גם שורות שמכילות רק תיבות סימון ריקות)
  var lastRow = sheet.getLastRow();
  var row = hr + 1;
  if (lastRow > hr) {
    var vals = sheet.getRange(hr + 1, 1, lastRow - hr, sheet.getLastColumn()).getDisplayValues();
    var i = 0;
    while (i < vals.length && (cell(vals[i], 'created').trim() || cell(vals[i], 'name').trim() || cell(vals[i], 'phone').trim())) i++;
    row = hr + 1 + i;
  }

  if (COL.created) sheet.getRange(row, COL.created).setValue(Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm'));
  sheet.getRange(row, COL.name).setValue(String(f.name || ''));
  sheet.getRange(row, COL.phone).setValue(String(f.phone || ''));
  if (COL.campaign) sheet.getRange(row, COL.campaign).setValue(String(f.campaign || 'הוזן ידנית'));
  var opt = ['status', 'reason', 'notes', 'followup', 'closed', 'value'];
  for (var j = 0; j < opt.length; j++) {
    var key = opt[j];
    if (key in f && f[key] !== '' && f[key] !== null) {
      writeCell(sheet, row, key, f[key]);
    }
  }
  try { logEvent(String(f.name || ''), String(f.phone || ''), 'ליד חדש', 'הוזן ידנית דרך הדף'); } catch (err) {}
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
  if (!COL[key]) return; // אין עמודה כזו בגיליון הזה
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
  if (key === 'followup') {
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

// ---------- משימות ----------
// לשונית נפרדת ולא עמודה בגיליון הלידים: לליד אחד יכולות להיות כמה משימות,
// ומשימה שבוצעה נשארת בשיטס כהיסטוריה (הדף מציג רק פתוחות).
// עמודות: A מזהה, B נוצר, C שם הליד, D טלפון, E משימה, F תאריך יעד, G בוצע, H בוצע בתאריך
var TASK_COL = { id: 1, created: 2, name: 3, phone: 4, text: 5, due: 6, done: 7, doneAt: 8 };

function getTasksSheet(create) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(TASKS_SHEET_NAME);
  if (sheet || !create) return sheet;
  sheet = ss.insertSheet(TASKS_SHEET_NAME);
  sheet.setRightToLeft(true);
  sheet.getRange(1, 1, 1, 8).setValues([['מזהה', 'נוצר', 'שם הליד', 'טלפון', 'משימה', 'תאריך יעד', 'בוצע', 'בוצע בתאריך']])
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  // טקסט פשוט, כדי שגוגל לא ימחק את האפס המוביל בטלפון
  sheet.getRange(1, TASK_COL.id, sheet.getMaxRows(), 1).setNumberFormat('@');
  sheet.getRange(1, TASK_COL.phone, sheet.getMaxRows(), 1).setNumberFormat('@');
  sheet.hideColumns(TASK_COL.id);
  sheet.setColumnWidth(TASK_COL.text, 320);
  return sheet;
}

// רק משימות פתוחות. אם הלשונית עוד לא קיימת - אין משימות.
function listTasks() {
  var sheet = getTasksSheet(false);
  var out = [];
  if (!sheet || sheet.getLastRow() < 2) return out;
  var n = sheet.getLastRow() - 1;
  var vals = sheet.getRange(2, 1, n, 8).getValues();
  var shown = sheet.getRange(2, 1, n, 8).getDisplayValues();
  for (var i = 0; i < n; i++) {
    var v = vals[i];
    if (!String(v[0]).trim() || v[6] === true) continue;
    out.push({
      id: String(v[0]), leadName: shown[i][2], leadPhone: shown[i][3],
      text: shown[i][4], due: followupText(v[5], shown[i][5])
    });
  }
  return out;
}

// המשימות שבוצעו של ליד אחד (לפי טלפון), מהחדשה לישנה. נטען רק כשפותחים את הליד.
function doneTasks(req) {
  var out = [];
  var key = phoneKey(req.phone);
  var sheet = getTasksSheet(false);
  if (!key || !sheet || sheet.getLastRow() < 2) return { ok: true, tasks: out };
  var n = sheet.getLastRow() - 1;
  var vals = sheet.getRange(2, 1, n, 8).getValues();
  var shown = sheet.getRange(2, 1, n, 8).getDisplayValues();
  for (var i = n - 1; i >= 0 && out.length < 50; i--) {
    if (vals[i][6] !== true || !String(vals[i][0]).trim() || phoneKey(shown[i][3]) !== key) continue;
    out.push({
      id: String(vals[i][0]), leadName: shown[i][2], leadPhone: shown[i][3],
      text: shown[i][4], due: followupText(vals[i][5], shown[i][5]), doneAt: shown[i][7]
    });
  }
  return { ok: true, tasks: out };
}

function addTask(req) {
  var f = req.fields || {};
  var text = norm(f.text);
  if (!text) return { ok: false, error: 'empty_task' };
  var sheet = getTasksSheet(true);
  var id = Utilities.getUuid();
  var due = f.due ? toCellValue('followup', f.due) : '';
  var row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 5).setValues([[
    id, Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm'),
    String(f.leadName || ''), String(f.leadPhone || ''), text
  ]]);
  var dueCell = sheet.getRange(row, TASK_COL.due);
  dueCell.setValue(due);
  if (due instanceof Date) {
    var hasTime = due.getHours() !== 0 || due.getMinutes() !== 0;
    dueCell.setNumberFormat(hasTime ? 'dd/MM/yyyy HH:mm' : 'dd/MM/yyyy');
  }
  sheet.getRange(row, TASK_COL.done).insertCheckboxes().setValue(false);
  try {
    logEvent(String(f.leadName || ''), String(f.leadPhone || ''), 'משימה',
             text + (due instanceof Date ? ' · עד ' + followupText(due, '') : ''));
  } catch (err) {}
  stampUpdatedByPhone(f.leadPhone);
  return { ok: true, task: { id: id, leadName: String(f.leadName || ''), leadPhone: String(f.leadPhone || ''),
                             text: text, due: followupText(due, '') } };
}

function setTaskDone(req) {
  var sheet = getTasksSheet(false);
  var id = String(req.id || '');
  if (!sheet || !id || sheet.getLastRow() < 2) return { ok: false, error: 'task_not_found' };
  var ids = sheet.getRange(2, TASK_COL.id, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) !== id) continue;
    var done = req.done !== false;
    sheet.getRange(i + 2, TASK_COL.done).setValue(done);
    sheet.getRange(i + 2, TASK_COL.doneAt).setValue(done ? Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm') : '');
    if (done) {
      try {
        var t = sheet.getRange(i + 2, TASK_COL.name, 1, 3).getDisplayValues()[0]; // שם, טלפון, משימה
        logEvent(t[0], t[1], 'משימה בוצעה', t[2]);
        stampUpdatedByPhone(t[1]);
      } catch (err) {}
    } else {
      try {
        var back = sheet.getRange(i + 2, TASK_COL.name, 1, 3).getDisplayValues()[0];
        logEvent(back[0], back[1], 'משימה נפתחה מחדש', back[2]);
        stampUpdatedByPhone(back[1]);
      } catch (err) {}
    }
    return { ok: true };
  }
  return { ok: false, error: 'task_not_found' };
}

// ---------- היסטוריה ----------
// לשונית נפרדת, שורה לכל אירוע: שינוי סטטוס, הערה, מעקב, שווי, משימה.
// נרשם רק מה שנעשה דרך הדף - שינוי ידני בגיליון לא עובר כאן.
// הקישור לליד הוא לפי טלפון (9 הספרות האחרונות), כי מספרי שורות משתנים במיון ובמחיקה.
// עמודות: A תאריך, B שם הליד, C טלפון, D סוג, E פירוט
function getHistorySheet(create) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (sheet || !create) return sheet;
  sheet = ss.insertSheet(HISTORY_SHEET_NAME);
  sheet.setRightToLeft(true);
  sheet.getRange(1, 1, 1, 5).setValues([['תאריך', 'שם הליד', 'טלפון', 'סוג', 'פירוט']]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange(1, 3, sheet.getMaxRows(), 1).setNumberFormat('@'); // טלפון כטקסט, בלי לאבד אפס מוביל
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(5, 420);
  return sheet;
}

function logEvent(name, phone, type, text) {
  var sheet = getHistorySheet(true);
  var row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 5).setValues([[new Date(), String(name || ''), String(phone || ''), type, String(text || '')]]);
  sheet.getRange(row, 1).setNumberFormat('dd/MM/yyyy HH:mm');
}

function phoneKey(p) {
  var d = String(p == null ? '' : p).replace(/\D/g, '');
  return d.length > 9 ? d.slice(-9) : d;
}

// האירועים של ליד אחד, מהחדש לישן
function leadHistory(req) {
  var events = [];
  var key = phoneKey(req.phone);
  var sheet = getHistorySheet(false);
  if (!key || !sheet || sheet.getLastRow() < 2) return { ok: true, events: events };
  var n = sheet.getLastRow() - 1;
  var vals = sheet.getRange(2, 1, n, 5).getValues();
  var shown = sheet.getRange(2, 1, n, 5).getDisplayValues();
  for (var i = n - 1; i >= 0 && events.length < 200; i--) {
    if (phoneKey(shown[i][2]) !== key) continue;
    var at = vals[i][0] instanceof Date ? Utilities.formatDate(vals[i][0], Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') : shown[i][0];
    events.push({ at: at, type: shown[i][3], text: shown[i][4] });
  }
  return { ok: true, events: events };
}
