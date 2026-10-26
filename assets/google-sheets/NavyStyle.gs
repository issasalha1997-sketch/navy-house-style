/**
 * Navy house style for Google Sheets (Apps Script).
 *
 * Set-up: open the spreadsheet, Extensions > Apps Script, paste this file as NavyStyle.gs and Dialog.html as an
 * HTML file named Dialog, save, then reload the spreadsheet. Run buildStarter() once on a new, empty spreadsheet
 * to get a Dashboard, a Tracker and a Lists tab in the house look; change the names and columns to suit.
 *
 * Menus, dialogs and drawing buttons work in a web browser only. On phones people type in the sheets, which is
 * why nothing here is needed for the data to stay correct.
 */

var NAVY = '#243168', INK = '#1B1F2A', SOFT = '#2C3A78', WHITE = '#FFFFFF';
var LINE = '#D3D8EA', RULE = '#C9CEDF', PICKER = '#EEF0F7', BAR = '#5B6AA6';
var REF_TAB = '#8F99C2', FADED = '#7A82AE', LINK = '#0563C1';
var STATUS = {done: '#1E7B3C', waiting: '#8A5300', withOthers: '#00618A', stopped: '#A4262C'};
var LOOK_FILL = '#FFE699', BANNER_FILL = '#FFF4CE', BANNER_EDGE = '#C27C0E';
var FONT = 'Verdana';            // Tahoma's wider sibling: same x-height, about 14% wider (see google-sheets.md)
var DATE_FORMAT = 'd mmm yyyy';
var MEDIUM = SpreadsheetApp.BorderStyle.SOLID_MEDIUM, THIN = SpreadsheetApp.BorderStyle.SOLID;

/** The custom menu, added each time the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Actions')
    .addItem('Add a request', 'showAddDialog')
    .addSeparator()
    .addItem('Apply the navy theme', 'applyNavyTheme')
    .addToUi();
}

/** Theme: the font and colours used by every cell that has no format of its own, and the link colour. */
function applyNavyTheme() {
  var ss = SpreadsheetApp.getActive();
  var theme = ss.getSpreadsheetTheme() || ss.resetSpreadsheetTheme();
  theme.setFontFamily(FONT);
  theme.setConcreteColor(SpreadsheetApp.ThemeColorType.TEXT, colour_(INK));
  theme.setConcreteColor(SpreadsheetApp.ThemeColorType.ACCENT1, colour_(NAVY));
  theme.setConcreteColor(SpreadsheetApp.ThemeColorType.HYPERLINK, colour_(LINK));
  ss.setSpreadsheetTheme(theme);
}

/** Builds Dashboard | Tracker | Lists on a new spreadsheet. Refuses to touch sheets that already exist. */
function buildStarter() {
  var ss = SpreadsheetApp.getActive();
  ['Dashboard', 'Tracker', 'Lists'].forEach(function (name) {
    if (ss.getSheetByName(name)) throw new Error('A sheet called ' + name + ' already exists. Run this on a new, empty spreadsheet.');
  });
  applyNavyTheme();
  var statuses = [['Waiting', STATUS.waiting, false, false], ['On hold', STATUS.waiting, true, false],
                  ['With supplier', STATUS.withOthers, false, false], ['Active', STATUS.done, true, false],
                  ['Stopped', STATUS.stopped, false, false]];
  var teams = ['Platform', 'Data', 'Sales'];

  // Lists: the drop-down values (a reference tab)
  var ls = ss.insertSheet('Lists');
  ls.setTabColor(REF_TAB).setHiddenGridlines(true);
  ls.setColumnWidth(1, 19).setColumnWidth(2, 180).setColumnWidth(3, 19).setColumnWidth(4, 180);
  title_(ls, 'B1', 'Lists', 14, 37);
  hint_(ls.getRange('B2'), 'Values for the drop-downs. New teams go at the end of the Team list.');
  ls.getRange('B3').setValue('Fixed: the Dashboard counts each status.').setFontFamily(FONT).setFontSize(8.5)
    .setFontWeight('bold').setFontStyle('italic').setFontColor(STATUS.stopped).setWrap(true).setVerticalAlignment('bottom');
  ls.setRowHeight(3, 40);
  header_(ls.getRange('B4'), ['Status']);
  header_(ls.getRange('D4'), ['Team']);
  body_(ls.getRange(5, 2, statuses.length, 1)).setValues(statuses.map(function (s) { return [s[0]]; }));
  body_(ls.getRange(5, 4, teams.length, 1)).setValues(teams.map(function (t) { return [t]; }));

  // Tracker: plain typeable values only, in the order of the form people already use
  var ws = ss.insertSheet('Tracker');
  ws.setTabColor(NAVY);
  var cols = [['Name', 22], ['Email', 32], ['Team', 24], ['Status', 18], ['Requested', 13], ['Done on', 13], ['Notes', 60]];
  cols.forEach(function (c, i) { ws.setColumnWidth(i + 1, px_(c[1])); });
  groupLabels_(ws, 1, [['Who', 3], ['Tracking', 3], ['Notes', 1]]);
  header_(ws.getRange(2, 1), cols.map(function (c) { return c[0]; }));
  ws.setRowHeight(2, 35);
  var data = ws.getRange('A3:G1000');
  body_(data).setVerticalAlignment('top');
  ws.getRange('E3:F1000').setNumberFormat(DATE_FORMAT).setHorizontalAlignment('right');
  ws.getRange('G3:G1000').setWrap(true);
  ws.setFrozenRows(2);
  ws.setFrozenColumns(2);
  ws.getRange('A2:G').createFilter();                      // filter buttons on the header row
  ws.getRange('D3:D1000').setDataValidation(listRule_(ls.getRange('B5:B'), true, 'Pick a status from the list (the Dashboard counts use it).'));
  ws.getRange('C3:C1000').setDataValidation(listRule_(ls.getRange('D5:D'), false, 'Not on the Lists tab yet. Keep it, or add it there first.'));
  ws.getRange('E3:F1000').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=AND(ISNUMBER(E3),E3>=DATE(2020,1,1),E3<=TODAY())')
    .setAllowInvalid(false)
    .setHelpText('Type a date, for example 16 Mar 2026, or press Ctrl+; for today.')
    .build());
  // Google Sheets applies only the first rule that is true for a cell, so the most specific rules come first.
  var rules = [formatRule_('=AND($B3<>"",ISERROR(SEARCH("@",$B3)))', ws.getRange('B3:B1000'), {background: LOOK_FILL})];
  statuses.forEach(function (s) {
    rules.push(formatRule_('=$D3="' + s[0] + '"', ws.getRange('D3:D1000'), {font: s[1], bold: s[2], italic: s[3]}));
  });
  rules.push(formatRule_('=$D3="Stopped"', ws.getRange('A3:G1000'), {font: FADED}));
  ws.setConditionalFormatRules(rules);

  // Dashboard: first sheet, reads like a page, formulas only
  var ds = ss.insertSheet('Dashboard');
  ds.setTabColor(NAVY).setHiddenGridlines(true);
  [19, px_(30), px_(14), px_(14), px_(14), px_(14), 19].forEach(function (w, i) { ds.setColumnWidth(i + 1, w); });
  banner_(ds.getRange('B1:F1'), 'Use the Actions menu or the buttons in a web browser.', 'On a phone, type in the Tracker as usual.');
  ds.setRowHeight(2, 13);
  title_(ds, 'B3', 'Requests at a glance', 20, 43);
  ds.getRange('B4').setValue('Counts work themselves out from the Tracker. Nothing here needs typing.')
    .setFontFamily(FONT).setFontSize(10).setFontColor(SOFT);
  ds.setRowHeight(5, 61);                                  // the button row: drawings go here (google-sheets.md)
  ds.getRange('B6:F6').setBorder(null, null, true, null, null, null, NAVY, MEDIUM);
  ds.setRowHeight(6, 8);
  ds.setRowHeight(7, 16);
  section_(ds.getRange('B8'), 'By status');
  boxHeader_(ds.getRange('B9:E9'), ['Status', 'Lines', 'Oldest request', '']);
  statuses.forEach(function (s, i) {
    var r = 10 + i;
    ds.getRange(r, 2).setValue(s[0]);
    ds.getRange(r, 3).setFormula('=COUNTIF(Tracker!$D$3:$D,B' + r + ')');
    ds.getRange(r, 4).setFormula('=IF(C' + r + '=0,"",MINIFS(Tracker!$E$3:$E,Tracker!$D$3:$D,B' + r + '))');
    ds.getRange(r, 5).setFormula('=SPARKLINE(C' + r + ',{"charttype","bar";"max",MAX($C$10:$C$14);"color1","' + BAR + '"})');
    boxRow_(ds.getRange(r, 2, 1, 3));
    ds.getRange(r, 2, 1, 4).setBorder(null, null, true, null, null, null, LINE, THIN);
  });
  ds.getRange('D10:D14').setNumberFormat(DATE_FORMAT);
  var total = ds.getRange('B15:D15');
  total.setValues([['All lines', '=SUM(C10:C14)', '']]).setFontFamily(FONT).setFontSize(11).setFontWeight('bold')
    .setFontColor(NAVY).setVerticalAlignment('middle');
  ds.getRange('C15:D15').setHorizontalAlignment('right');
  total.setBorder(true, null, true, null, null, null, NAVY, THIN);
  ds.setRowHeight(15, 27);
  ds.setRowHeight(16, 21);
  section_(ds.getRange('B17'), 'Needs a look');
  boxHeader_(ds.getRange('B18:C18'), ['What', 'Lines']);
  ds.getRange('B19:C19').setValues([['Email without an @', '=COUNTIFS(Tracker!$B$3:$B,"<>",Tracker!$B$3:$B,"<>*@*")']]);
  boxRow_(ds.getRange('B19:C19'), 10.5, false);
  ds.setConditionalFormatRules([formatRule_('=N($C19)>0', ds.getRange('C19'), {font: STATUS.stopped, bold: true})]);
  ds.setRowHeight(20, 21);
  title_(ds, 'B21', 'This month', 13, 53);
  tile_(ds, 'C', 21, '=COUNTIF(Tracker!$D$3:$D,"Active")', 'Active now');
  tile_(ds, 'D', 21, '=COUNTIF(Tracker!$D$3:$D,"Waiting")+COUNTIF(Tracker!$D$3:$D,"On hold")', 'Waiting on us');
  tile_(ds, 'E', 21, '=COUNTIF(Tracker!$D$3:$D,"With supplier")', 'With the supplier');
  ds.setRowHeight(22, 40);
  var noteCell = ds.getRange('B24:F24').merge();
  noteCell.setValue('Notes: a line counts once per status. Stopped lines stay in the Tracker with their date and reason.')
    .setFontFamily(FONT).setFontSize(8.5).setFontStyle('italic').setFontColor(SOFT).setWrap(true).setVerticalAlignment('top');
  ds.setRowHeight(24, 40);
  ds.protect().setWarningOnly(true);                       // read-only counts; the Tracker stays open

  // order the tabs Dashboard | Tracker | Lists, and remove the empty first sheet a new spreadsheet comes with
  [ds, ws, ls].forEach(function (sheet, i) { ss.setActiveSheet(sheet); ss.moveActiveSheet(i + 1); });
  var blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && blank.getLastColumn() === 0) ss.deleteSheet(blank);
  ss.setActiveSheet(ds);
}

/** Opens the "Add a request" screen (Dialog.html). */
function showAddDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog').setWidth(560).setHeight(470);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add a request');
}

/** Values under a heading on the Lists tab, for the screen's drop-downs. */
function listValues(heading) {
  var ls = SpreadsheetApp.getActive().getSheetByName('Lists');
  var headings = ls.getRange(4, 1, 1, ls.getLastColumn()).getValues()[0];
  var col = headings.indexOf(heading) + 1;
  if (col < 1 || ls.getLastRow() < 5) return [];
  return ls.getRange(5, col, ls.getLastRow() - 4, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (v) { return v !== ''; });
}

/** Adds one line to the Tracker with plain values. Notes are stamped with today's date. */
function addRequest(data) {
  var ss = SpreadsheetApp.getActive(), tz = ss.getSpreadsheetTimeZone();
  var requested = readDate_(data.requested, tz);
  var note = data.notes ? Utilities.formatDate(new Date(), tz, 'd MMM yyyy') + ': ' + data.notes : '';
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);                                    // two people adding at once never share a row
  try {
    var ws = ss.getSheetByName('Tracker');
    var row = Math.max(ws.getLastRow() + 1, 3);
    ws.getRange(row, 1, 1, 7).setValues([[asText_(data.name), asText_(data.email), asText_(data.team), 'Waiting', requested, '', asText_(note)]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

/** A date typed like 16 Mar 2026 (or 16 March 2026), as a date with no time in the spreadsheet's time zone.
 *  Empty means today. Anything else, including 16/03/2026 or 31 Feb 2026, is refused with an example. */
function readDate_(text, tz) {
  var t = String(text || '').trim().replace(/\s+/g, ' ') || Utilities.formatDate(new Date(), tz, 'd MMM yyyy');
  var m = t.match(/^(\d{1,2}) ([A-Za-z]{3})[A-Za-z]* (\d{4})$/);
  if (m) {
    var tidy = Number(m[1]) + ' ' + m[2].charAt(0).toUpperCase() + m[2].slice(1).toLowerCase() + ' ' + m[3];
    var d = Utilities.parseDate(tidy, tz, 'd MMM yyyy');
    if (Utilities.formatDate(d, tz, 'd MMM yyyy') === tidy) return d;
  }
  throw new Error('Type the date like 16 Mar 2026.');
}

// ---------------------------------------------------------------- helpers

function colour_(hex) { return SpreadsheetApp.newColor().setRgbColor(hex).build(); }

/** Excel column-width units to pixels (7 per unit + 5), widened 15% because Verdana sets wider than Tahoma. */
function px_(excelWidth) { return Math.round((excelWidth * 7 + 5) * 1.15); }

/** Text that Sheets would read as a formula or a number is kept exactly as typed. */
function asText_(s) { return /^[=+\-@]/.test(s || '') ? "'" + s : (s || ''); }

function body_(range) { return range.setFontFamily(FONT).setFontSize(10).setFontColor(INK); }

function hint_(range, text) {
  return range.setValue(text).setFontFamily(FONT).setFontSize(9).setFontStyle('italic').setFontColor(SOFT);
}

function title_(sheet, a1, text, size, heightPx) {
  sheet.getRange(a1).setValue(text).setFontFamily(FONT).setFontSize(size).setFontWeight('bold')
    .setFontColor(NAVY).setVerticalAlignment('middle');
  sheet.setRowHeight(sheet.getRange(a1).getRow(), heightPx);
}

/** The white-on-navy header row of a table, starting at the given cell. */
function header_(firstCell, labels) {
  var r = firstCell.offset(0, 0, 1, labels.length);
  return r.setValues([labels]).setBackground(NAVY).setFontColor(WHITE).setFontWeight('bold')
    .setFontFamily(FONT).setFontSize(10).setVerticalAlignment('middle').setWrap(true);
}

/** Labels over groups of columns: in each group's first cell, bold navy, with a medium navy rule under the row. */
function groupLabels_(sheet, row, groups) {
  var col = 1;
  groups.forEach(function (g) {
    sheet.getRange(row, col).setValue(g[0]);
    col += g[1];
  });
  sheet.getRange(row, 1, 1, col - 1).setFontFamily(FONT).setFontSize(10).setFontWeight('bold').setFontColor(NAVY)
    .setVerticalAlignment('middle').setBorder(null, null, true, null, null, null, NAVY, MEDIUM);
  sheet.setRowHeight(row, 27);
}

function section_(range, text) {
  range.setValue(text).setFontFamily(FONT).setFontSize(13).setFontWeight('bold').setFontColor(NAVY).setVerticalAlignment('bottom');
  range.getSheet().setRowHeight(range.getRow(), 32);
}

/** A dashboard box header: 9.5 bold navy, wrapped, thin navy underline; first column left, numbers right. */
function boxHeader_(range, labels) {
  range.setValues([labels]).setFontFamily(FONT).setFontSize(9.5).setFontWeight('bold').setFontColor(NAVY)
    .setVerticalAlignment('bottom').setWrap(true).setHorizontalAlignment('right')
    .setBorder(null, null, true, null, null, null, NAVY, THIN);
  range.offset(0, 0, 1, 1).setHorizontalAlignment('left');
  range.getSheet().setRowHeight(range.getRow(), 37);
}

function boxRow_(range, size, boldFirst) {
  range.setFontFamily(FONT).setFontSize(size || 11).setFontColor(INK).setVerticalAlignment('middle').setHorizontalAlignment('right');
  range.offset(0, 0, 1, 1).setHorizontalAlignment('left').setFontWeight(boldFirst === false ? 'normal' : 'bold');
  range.getSheet().setRowHeight(range.getRow(), 27);
}

/** A big number over a small caption, inside one thin box. */
function tile_(sheet, col, row, formula, caption) {
  var n = sheet.getRange(col + row), t = sheet.getRange(col + (row + 1));
  n.setFormula(formula).setFontFamily(FONT).setFontSize(22).setFontWeight('bold').setFontColor(NAVY)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, null, true, null, null, RULE, THIN);
  t.setValue(caption).setFontFamily(FONT).setFontSize(8.5).setFontColor(SOFT).setHorizontalAlignment('center')
    .setVerticalAlignment('top').setWrap(true).setBorder(null, true, true, true, null, null, RULE, THIN);
}

/** A notice across the top: amber fill, medium amber left edge, first sentence bold. */
function banner_(range, boldPart, rest) {
  var text = boldPart + ' ' + rest;
  var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
  range.merge().setBackground(BANNER_FILL).setFontFamily(FONT).setFontSize(9.5).setFontColor(INK)
    .setVerticalAlignment('middle').setWrap(true)
    .setBorder(null, true, null, null, null, null, BANNER_EDGE, MEDIUM);
  range.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text).setTextStyle(0, boldPart.length, bold).build());
  range.getSheet().setRowHeight(range.getRow(), 48);
}

/** A drop-down from a range on the Lists tab: strict where counts depend on the value, a warning elsewhere. */
function listRule_(source, strict, helpText) {
  return SpreadsheetApp.newDataValidation().requireValueInRange(source, true).setAllowInvalid(!strict)
    .setHelpText(helpText).build();
}

/** A conditional format from a custom formula. fmt: {background, font, bold, italic, strikethrough} */
function formatRule_(formula, range, fmt) {
  var b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setRanges([range]);
  if (fmt.background) b.setBackground(fmt.background);
  if (fmt.font) b.setFontColor(fmt.font);
  if (fmt.bold) b.setBold(true);
  if (fmt.italic) b.setItalic(true);
  if (fmt.strikethrough) b.setStrikethrough(true);
  return b.build();
}
