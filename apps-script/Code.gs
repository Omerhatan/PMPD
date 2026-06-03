// ════════════════════════════════════════════════════════
//  PMPD — Google Apps Script  |  הדבק ב-Extensions → Apps Script
// ════════════════════════════════════════════════════════

var DASHBOARD_URL = 'https://omerhatan.github.io/PMPD/';
var SHEET_INDEX   = 0;   // הגיליון הראשון (0-based)

// עמודות קבועות — אל תשנה!
var COL_EMPTY    = 1;   // A — ריק (מרווח)
var COL_ID       = 2;   // B — מספר פרויקט
var COL_NAME     = 3;   // C — שם פרויקט
var COL_ARCH     = 4;   // D — אדריכל
var COL_UNITS    = 5;   // E — יח"ד
var COL_TYPE     = 6;   // F — סוג פרויקט
// G = התנעה, H = ריק, I-M = תכנון (5), N = ריק, O-V = רישוי (8), W = ריק, X-AA = שיווק (4)
var TOTAL_COLS   = 27;

// ════════════════════════════════════════════════════════
//  תפריט — נפתח אוטומטית בכל פתיחת הגיליון
// ════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗️ ניהול פרויקטים')
    .addItem('➕  פרויקט חדש',        'showNewProjectDialog')
    .addSeparator()
    .addItem('🔒  הגן על המבנה',       'protectStructure')
    .addItem('✅  בדוק תקינות נתונים', 'validateAllData')
    .addSeparator()
    .addItem('🚀  פתח דשבורד',         'openDashboard')
    .addToUi();
}

// ════════════════════════════════════════════════════════
//  פתיחת דשבורד
// ════════════════════════════════════════════════════════
function openDashboard() {
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + DASHBOARD_URL + '","_blank");google.script.host.close();</script>'
  ).setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, 'פותח דשבורד...');
}

// ════════════════════════════════════════════════════════
//  הגנה על המבנה
// ════════════════════════════════════════════════════════
function protectStructure() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[SHEET_INDEX];
  var me    = Session.getEffectiveUser();
  var ui    = SpreadsheetApp.getUi();

  // הסר הגנות קיימות
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { p.remove(); });

  // הגן על שורות הכותרת (1-2)
  var hProt = sheet.getRange('1:2').protect();
  hProt.setDescription('כותרות — אל תשנה');
  _lockToMe(hProt, me);

  // הגן על עמודות A-B (מרווח + מספר)
  var cProt = sheet.getRange('A:B').protect();
  cProt.setDescription('עמודות מבנה — אוטומטי');
  _lockToMe(cProt, me);

  ui.alert('✅ המבנה מוגן!\n\nשורות הכותרת ועמודות A-B נעולות.\nרק אתה יכול לשנות אותן.');
}

function _lockToMe(protection, me) {
  var others = protection.getEditors().filter(function(e) {
    return e.getEmail() !== me.getEmail();
  });
  if (others.length > 0) protection.removeEditors(others);
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

// ════════════════════════════════════════════════════════
//  בדיקת תקינות
// ════════════════════════════════════════════════════════
var VALID_KEYWORDS = ['בוצע','הושלם','התקבל','לא נדרש','לא יבוצע','טרם התקבל','תתקבל החלטה','-'];
var DATE_PATTERN   = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
var MONTH_PATTERN  = /^[\u05D0-\u05EA\u05F3']{2,6}-\d{2}$/;

function validateAllData() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheets()[SHEET_INDEX];
  var data    = sheet.getDataRange().getValues();
  var errors  = [];

  for (var r = 2; r < data.length; r++) {
    var row = data[r];
    var id  = row[1];
    if (!id || isNaN(parseInt(id))) continue;
    var name = row[2];
    if (!name || String(name).trim() === '') {
      errors.push('שורה ' + (r+1) + ': שם פרויקט ריק');
    }
    var units = row[4];
    if (units !== '' && isNaN(parseInt(units))) {
      errors.push('שורה ' + (r+1) + ' (' + name + '): יח"ד לא תקין — "' + units + '"');
    }
    // בדוק ערכי אבני דרך
    var milestoneCols = [8,9,10,11,12,14,15,16,17,18,19,20,21,23,24,25,26];
    milestoneCols.forEach(function(c) {
      var val = String(row[c] || '').trim();
      if (val === '') return;
      if (VALID_KEYWORDS.indexOf(val) !== -1) return;
      if (DATE_PATTERN.test(val)) return;
      if (MONTH_PATTERN.test(val)) return;
      errors.push('שורה ' + (r+1) + ' (' + name + '), עמודה ' + (c+1) + ': ערך לא מוכר — "' + val + '"');
    });
  }

  if (errors.length === 0) {
    SpreadsheetApp.getUi().alert('✅ כל הנתונים תקינים!');
  } else {
    SpreadsheetApp.getUi().alert(
      '⚠️ נמצאו ' + errors.length + ' בעיות:\n\n' + errors.slice(0,15).join('\n') +
      (errors.length > 15 ? '\n...ועוד ' + (errors.length-15) + ' בעיות' : '')
    );
  }
}

// ════════════════════════════════════════════════════════
//  דיאלוג פרויקט חדש
// ════════════════════════════════════════════════════════
function showNewProjectDialog() {
  var html = HtmlService.createHtmlOutputFromFile('NewProject')
    .setWidth(580)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ פרויקט חדש');
}

// נקרא מה-HTML
function addNewProject(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[SHEET_INDEX];

  // מצא את הפרויקט הבא
  var values  = sheet.getDataRange().getValues();
  var maxId   = 0;
  var lastDataRow = 2; // שורת נתונים ראשונה (אחרי 2 שורות כותרת)
  for (var r = 2; r < values.length; r++) {
    var id = parseInt(values[r][1]);
    if (!isNaN(id)) {
      maxId = Math.max(maxId, id);
      lastDataRow = r + 1; // 1-based
    }
  }
  var newId  = maxId + 1;
  var newRow = lastDataRow + 1;

  // בנה שורה ריקה בגודל הנכון
  var row = [];
  for (var i = 0; i < TOTAL_COLS; i++) row.push('');
  row[COL_ID    - 1] = newId;
  row[COL_NAME  - 1] = data.name;
  row[COL_ARCH  - 1] = data.architect || '';
  row[COL_UNITS - 1] = data.units     || '';
  row[COL_TYPE  - 1] = data.type      || '';

  sheet.getRange(newRow, 1, 1, TOTAL_COLS).setValues([row]);

  // עצב את השורה החדשה כמו שאר הנתונים
  var range = sheet.getRange(newRow, 1, 1, TOTAL_COLS);
  range.setFontFamily('Arial');
  range.setFontSize(11);

  return { success: true, id: newId, row: newRow };
}

// מחזיר רשימת הפרויקטים לדיאלוג
function getProjectTypes() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheets()[SHEET_INDEX];
  var values = sheet.getDataRange().getValues();
  var types  = {};
  for (var r = 2; r < values.length; r++) {
    var t = String(values[r][5] || '').trim();
    if (t) types[t] = true;
  }
  return Object.keys(types);
}
