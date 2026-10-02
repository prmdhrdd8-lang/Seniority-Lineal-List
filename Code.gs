/** Personnel form for Google Sheets. Paste into Extensions > Apps Script as Code.gs */
const SHEET_NAME = 'PLANTILLA';
const HEADERS = ['UNIT ASSIGNMENT','ITEM NUMBER (2)','POSITION TITLE (3)','SALARY GRADE (4)','AUTHORIZED ANNUAL SALARY(5)','ACTUAL ANNUAL SALARY (6)','STEP (7)','LAST NAME (11)','FIRST NAME (12)','MIDDLE NAME (13)','SEX (14)','DATE OF BIRTH (15)','AGE','DATE OF COMPULSORY RETIREMENT','TIN (16)','DATE OF ORIGINAL APPOINTMENT (17)','DATE OF LAST PROMOTION/APPOINTMENT (18)','STATUS (19)','CIVIL SERVICE ELIGIBILITY (20)','COMMENT/ANNOTATION (21)','PWD (Y) (22)','INDIGENOUS PEOPLE (Y) (23)','SOLO PARENT (ID_NUMBER) (24)','GSIS BP NUMBER','POSITION CLASSIFICATION','DEGREE COURSE/ UNITS EARNED','LATEST CAREER COURSE TAKEN','DESIGNATION HELD FOR THE LAST 6 MONTHS','REMARKS'];
const DATE_COLS = [11, 13, 15, 16];   // 0-based: DOB, retirement, original appt, last promotion
const TEXT_COLS = [1, 14, 22, 23];    // item no., TIN, solo parent ID, GSIS (keep leading zeros/dashes)
const AGE_COL = 12, DOB_COL = 11, LAST = 7, FIRST = 8;

/** Web app entry point (Deploy > New deployment > Web app). */
function doGet() {
  getSheet_();
  return HtmlService.createHtmlOutputFromFile('Form')
    .setTitle('Personnel Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Personnel Form')
    .addItem('Open form', 'openForm')
    .addItem('Set up sheet (first time)', 'setupSheet')
    .addToUi();
}

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET_NAME) || setupSheet();
}

function setupSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#1f5c46').setFontColor('#ffffff').setWrap(true);
  sh.setFrozenRows(1);
  TEXT_COLS.forEach(c => sh.getRange(2, c + 1, sh.getMaxRows() - 1).setNumberFormat('@'));
  DATE_COLS.forEach(c => sh.getRange(2, c + 1, sh.getMaxRows() - 1).setNumberFormat('yyyy-mm-dd'));
  return sh;
}

function openForm() {
  getSheet_();
  const html = HtmlService.createHtmlOutputFromFile('Form').setWidth(980).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Personnel Form');
}

function getConfig() { return { headers: HEADERS, dateCols: DATE_COLS, ageCol: AGE_COL, dobCol: DOB_COL }; }

/** Search by last name, first name, item number or TIN. */
function searchRecords(q) {
  const sh = getSheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const data = sh.getRange(2, 1, n, HEADERS.length).getDisplayValues();
  q = String(q || '').toLowerCase().trim();
  const out = [];
  for (let i = 0; i < data.length && out.length < 60; i++) {
    const r = data[i];
    const label = [r[LAST], r[FIRST]].filter(String).join(', ');
    if (!label && !r[1]) continue;
    if (!q || (label + ' ' + r[1] + ' ' + r[14]).toLowerCase().indexOf(q) > -1)
      out.push({ row: i + 2, label: label || '(no name)', sub: r[2] + (r[0] ? ' · ' + r[0] : '') });
  }
  return out;
}

function getRecord(row) {
  const sh = getSheet_();
  const vals = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  return vals.map((v, c) => v instanceof Date
    ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : (v === null ? '' : String(v)));
}

/** row = 0 or null appends a new record; otherwise updates that row. */
function saveRecord(values, row) {
  const sh = getSheet_();
  const out = values.map((v, c) => {
    if (DATE_COLS.indexOf(c) > -1) {
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
      return m ? new Date(+m[1], +m[2] - 1, +m[3]) : '';
    }
    if (c === AGE_COL || [3, 6].indexOf(c) > -1) return v === '' ? '' : Number(v);
    return v;
  });
  const target = row || Math.max(sh.getLastRow() + 1, 2);
  sh.getRange(target, 1, 1, HEADERS.length).setValues([out]);
  return { row: target, saved: new Date().toLocaleTimeString() };
}

function deleteRecord(row) {
  if (row > 1) getSheet_().deleteRow(row);
  return true;
}
