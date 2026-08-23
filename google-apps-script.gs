/**
 * PicaPool intent-form logger + live dashboard.
 *
 * This file does NOT run on Vercel - it runs inside a Google Sheet.
 * Setup (see README.md for the full walkthrough):
 *   1. Create a blank Google Sheet.
 *   2. Extensions > Apps Script, delete the placeholder code, paste this whole file.
 *   3. Run `setupDashboard` once from the editor toolbar (grant the permissions it asks for).
 *   4. Deploy > New deployment > type "Web app". Execute as "Me", access "Anyone".
 *   5. Copy the /exec URL it gives you into APPS_SCRIPT_URL in api/track.js.
 *
 * Re-run `setupDashboard` any time from the sheet's "PicaPool" menu to rebuild
 * the Dashboard, Source Sheet, and Read Me tabs (e.g. after you change the
 * layout below). It never touches Submissions or Events - your data is never
 * at risk.
 */

var SHEET_SUBMISSIONS = 'Submissions';
var SHEET_EVENTS = 'Events';
var SHEET_DASHBOARD = 'Dashboard';
var SHEET_SOURCE = 'Source Sheet';
var SHEET_README = 'Read Me';
var BRAND = '#F03506';
var BRAND_WASH = '#FFE5DA';

var SUBMISSION_HEADERS = [
  'Timestamp', 'Visitor ID', 'Session ID', 'Is Returning Visitor', 'Source', 'Landing Path',
  'Full Name', 'College', 'Phone', 'Email', 'Travel Mode', 'Daily Spend', 'Commute Time',
  'Nearest Metro Station', 'Would Prefer Cab Share', 'Commute Effort', 'Time To Complete (s)',
  'Referrer', 'UTM Source', 'UTM Medium', 'UTM Campaign',
  'Device Type', 'Browser', 'OS', 'Language', 'Screen',
  'IP Address', 'City', 'Region', 'Country', 'ISP / Org', 'Status'
]; // A..AF (32 cols) -- Status is "Complete" (finished the form) or "Partial" (answered something, left)

var EVENT_HEADERS = [
  'Timestamp', 'Event Type', 'Visitor ID', 'Session ID', 'Is New Visitor', 'Source', 'Landing Path',
  'Furthest Question Reached', 'Percent Fields Filled', 'Time On Page (s)',
  'Referrer', 'UTM Source', 'UTM Medium', 'UTM Campaign',
  'Device Type', 'Browser', 'OS', 'Language', 'Screen',
  'IP Address', 'City', 'Region', 'Country', 'ISP / Org'
]; // A..X (24 cols)

// A "Partial" row is only written once someone has actually answered
// something (percentFilled above the 6% do-nothing floor) -- a bare
// bounce with zero interaction doesn't clutter Submissions.
var PARTIAL_MIN_PERCENT = 6;

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.kind === 'submission') {
      appendSubmissionRow_(p, 'Complete');
    } else if (p.kind === 'exit') {
      appendEvent(p);
      if ((p.percentFilled || 0) > PARTIAL_MIN_PERCENT) {
        appendSubmissionRow_(p, 'Partial');
      }
    } else if (p.kind === 'view') {
      appendEvent(p);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput('PicaPool intent-form logger is running.');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PicaPool')
    .addItem('Rebuild dashboard', 'setupDashboard')
    .addToUi();
}

function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground(BRAND).setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Ensures a Submissions sheet created before the Status column existed gets
// it added retroactively, so old sheets migrate cleanly on the next setup run.
// Every row written before this column existed was, by definition, a
// completed submission -- "Partial" rows didn't exist yet -- so those rows
// get backfilled as "Complete" rather than left blank/uncolored.
function ensureStatusColumn_(sheet) {
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var lastHeader = lastCol > 0 ? sheet.getRange(1, lastCol).getValue() : '';
  if (lastHeader !== 'Status') {
    var statusCol = lastCol + 1;
    sheet.getRange(1, statusCol).setValue('Status').setFontWeight('bold').setBackground(BRAND).setFontColor('#FFFFFF');
    if (lastRow > 1) {
      var backfill = [];
      for (var i = 0; i < lastRow - 1; i++) backfill.push(['Complete']);
      sheet.getRange(2, statusCol, backfill.length, 1).setValues(backfill);
    }
  }
}

// Colors whole rows by their Status: green for Complete, red for Partial.
// Status is always the last column (AF) per SUBMISSION_HEADERS.
function applySubmissionFormatting_(sheet) {
  var range = sheet.getRange('A2:AF5000');
  var completeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$AF2="Complete"')
    .setBackground('#E3F5E9')
    .setRanges([range])
    .build();
  var partialRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$AF2="Partial"')
    .setBackground('#FDE8E4')
    .setRanges([range])
    .build();
  sheet.setConditionalFormatRules([completeRule, partialRule]);
}

// Writes one Submissions row -- status is 'Complete' (finished and hit
// submit) or 'Partial' (answered at least one question, then left).
function appendSubmissionRow_(p, status) {
  var sheet = getOrCreateSheet_(SHEET_SUBMISSIONS, SUBMISSION_HEADERS);
  ensureStatusColumn_(sheet);
  sheet.appendRow([
    new Date(), p.visitorId || '', p.sessionId || '', p.isNewVisitor ? 'No' : 'Yes', p.source || '(direct)', p.path || '',
    p.fullName || '', p.college || '', p.phone || '', p.email || '', p.travelMode || '', p.spend || '', p.commuteTime || '',
    p.startArea || '', p.wouldPrefer || '', p.effort || '', p.timeToCompleteSec || p.timeOnPageSec || '',
    p.referrer || '', p.utmSource || '', p.utmMedium || '', p.utmCampaign || '',
    p.deviceType || '', p.browser || '', p.os || '', p.language || '', p.screen || '',
    p.ip || '', p.city || '', p.region || '', p.country || '', p.isp || '',
    status
  ]);
}

function appendEvent(p) {
  var sheet = getOrCreateSheet_(SHEET_EVENTS, EVENT_HEADERS);
  sheet.appendRow([
    new Date(), p.kind, p.visitorId || '', p.sessionId || '', p.isNewVisitor ? 'Yes' : 'No', p.source || '(direct)', p.path || '',
    p.furthestQuestion || 0, p.percentFilled || 0, p.timeOnPageSec || '',
    p.referrer || '', p.utmSource || '', p.utmMedium || '', p.utmCampaign || '',
    p.deviceType || '', p.browser || '', p.os || '', p.language || '', p.screen || '',
    p.ip || '', p.city || '', p.region || '', p.country || '', p.isp || ''
  ]);
}

/**
 * Builds (or rebuilds) the Dashboard + Read Me tabs, and ensures Submissions
 * and Events exist. Run this once after first pasting the script, and any
 * time after from the sheet's "PicaPool" menu. Never touches your data rows.
 */
function setupDashboard() {
  var submissionsSheet = getOrCreateSheet_(SHEET_SUBMISSIONS, SUBMISSION_HEADERS);
  ensureStatusColumn_(submissionsSheet);
  applySubmissionFormatting_(submissionsSheet);
  getOrCreateSheet_(SHEET_EVENTS, EVENT_HEADERS);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName(SHEET_DASHBOARD);
  if (old) ss.deleteSheet(old);
  var d = ss.insertSheet(SHEET_DASHBOARD, 0);

  function label(a1, text, opts) {
    var r = d.getRange(a1);
    r.setValue(text);
    if (opts && opts.bold) r.setFontWeight('bold');
    if (opts && opts.size) r.setFontSize(opts.size);
    if (opts && opts.bg) r.setBackground(opts.bg);
    if (opts && opts.color) r.setFontColor(opts.color);
    if (opts && opts.italic) r.setFontStyle('italic');
  }
  function formula(a1, f) { d.getRange(a1).setFormula(f); }

  // ---- Title ----
  d.getRange('A1:H1').merge();
  label('A1', '🚕 PicaPool - Live Dashboard', { bold: true, size: 18, bg: BRAND, color: '#FFFFFF' });
  d.getRange('A1:H1').setVerticalAlignment('middle');
  d.setRowHeight(1, 40);
  label('A2', 'Recalculates automatically whenever this sheet is open. Trackable links: share cab.picapool.tech/<anything> and it shows up below as a Source. See the "Read Me" tab for the full guide.', { italic: true, color: '#6C6560' });
  d.getRange('A2:H2').merge();

  // ---- Headline metrics ----
  d.getRange('A4:H4').merge();
  label('A4', 'HEADLINE METRICS', { bold: true, bg: BRAND_WASH });
  var metricHeaders = ['Total Views', 'Unique Visitors', 'Total Visits (Sessions)', 'Total Submissions', 'Conversion Rate', 'Avg. Time on Page - Exits (s)', 'Avg. Time to Complete (s)', 'Approx. Drop-offs'];
  d.getRange('A5:H5').setValues([metricHeaders]).setFontWeight('bold').setFontSize(9).setFontColor('#6C6560');
  formula('A6', '=COUNTIF(Events!B:B,"view")');
  formula('B6', '=IFERROR(COUNTA(UNIQUE(FILTER(Events!C2:C,Events!B2:B="view"))),0)');
  formula('C6', '=IFERROR(COUNTA(UNIQUE(FILTER(Events!D2:D,Events!B2:B="view"))),0)');
  formula('D6', '=COUNTIF(Submissions!AF2:AF,"Complete")');
  formula('E6', '=IFERROR(D6/A6,0)');
  formula('F6', '=IFERROR(ROUND(AVERAGE(FILTER(Events!J2:J,Events!B2:B="exit",Events!J2:J<>"")),1),0)');
  formula('G6', '=IFERROR(ROUND(AVERAGE(FILTER(Submissions!Q2:Q,Submissions!Q2:Q<>"",Submissions!AF2:AF="Complete")),1),0)');
  formula('H6', '=MAX(A6-D6,0)');
  d.getRange('A6:H6').setFontSize(16).setFontWeight('bold');
  d.getRange('E6').setNumberFormat('0.0%');

  // ---- By trackable link ----
  d.getRange('A8:F8').merge();
  label('A8', 'BY TRACKABLE LINK (SOURCE)', { bold: true, bg: BRAND_WASH });
  var sourceHeaders = ['Source', 'Views', 'Submissions', 'Conversion %', 'Approx. Drop-offs', 'Avg. Question Reached at Exit (0–7)'];
  d.getRange('A9:F9').setValues([sourceHeaders]).setFontWeight('bold').setFontSize(9).setFontColor('#6C6560');
  formula('A10', '=IFERROR(SORT(UNIQUE(FILTER(Events!F2:F,Events!F2:F<>""))),"No trackable links yet - try sharing cab.picapool.tech/instagram")');
  formula('B10', '=ARRAYFORMULA(IF($A10:$A59="","",COUNTIFS(Events!$F$2:$F$9999,$A10:$A59,Events!$B$2:$B$9999,"view")))');
  formula('C10', '=ARRAYFORMULA(IF($A10:$A59="","",COUNTIFS(Submissions!$E$2:$E$9999,$A10:$A59,Submissions!$AF$2:$AF$9999,"Complete")))');
  formula('D10', '=ARRAYFORMULA(IF($A10:$A59="","",IFERROR($C10:$C59/$B10:$B59,0)))');
  formula('E10', '=ARRAYFORMULA(IF($A10:$A59="","",IF($B10:$B59-$C10:$C59>0,$B10:$B59-$C10:$C59,0)))');
  formula('F10', '=ARRAYFORMULA(IF($A10:$A59="","",IFERROR(AVERAGEIFS(Events!$H$2:$H$9999,Events!$F$2:$F$9999,$A10:$A59,Events!$B$2:$B$9999,"exit"),"")))');
  d.getRange('D10:D59').setNumberFormat('0.0%');

  // ---- Drop-off funnel ----
  d.getRange('A62:D62').merge();
  label('A62', 'DROP-OFF FUNNEL - where riders leave the form', { bold: true, bg: BRAND_WASH });
  formula('A63', '=IFERROR(QUERY(Events!A2:X,"select H, count(H) where B=\'exit\' group by H order by H asc label H \'Furthest Question Reached (0-7)\', count(H) \'Visitors\'"),"No drop-off data yet")');

  // ---- Recent activity (Complete + Partial, color-coded) ----
  d.getRange('A74:F74').merge();
  label('A74', 'RECENT ACTIVITY (live feed) - green = completed, red = partial', { bold: true, bg: BRAND_WASH });
  formula('A75', '=IFERROR(QUERY(Submissions!A2:AF,"select A,G,H,E,Q,AF order by A desc limit 15 label A \'When\', G \'Name\', H \'College\', E \'Source\', Q \'Took (s)\', AF \'Status\'"),"No activity yet")');
  var feedRange = d.getRange('A75:F91');
  var feedGreen = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F75="Complete"').setBackground('#E3F5E9').setRanges([feedRange]).build();
  var feedRed = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$F75="Partial"').setBackground('#FDE8E4').setRanges([feedRange]).build();
  d.setConditionalFormatRules([feedGreen, feedRed]);

  // ---- New vs returning, device breakdown ----
  d.getRange('A95:C95').merge();
  label('A95', 'NEW VS RETURNING VISITORS', { bold: true, bg: BRAND_WASH });
  d.getRange('D95:F95').merge();
  label('D95', 'VISITS BY DEVICE TYPE', { bold: true, bg: BRAND_WASH });
  formula('A96', '=IFERROR(QUERY(Events!A2:X,"select E, count(E) where B=\'view\' group by E label E \'New Visitor?\', count(E) \'Views\'"),"No data yet")');
  formula('D96', '=IFERROR(QUERY(Events!A2:X,"select O, count(O) where B=\'view\' and O is not null and O <> \'\' group by O order by count(O) desc label O \'Device\', count(O) \'Views\'"),"No data yet")');

  // ---- Country breakdown (variable height - placed last) ----
  d.getRange('A106:D106').merge();
  label('A106', 'VISITS BY COUNTRY (top locations)', { bold: true, bg: BRAND_WASH });
  formula('A107', '=IFERROR(QUERY(Events!A2:X,"select W, count(W) where B=\'view\' and W is not null and W <> \'\' group by W order by count(W) desc label W \'Country\', count(W) \'Views\'"),"No location data yet")');

  d.setColumnWidths(1, 8, 150);
  d.setFrozenRows(2);

  buildSourceSheet_(ss);
  buildReadme_(ss);
  cleanupDefaultSheet_(ss);
  SpreadsheetApp.flush();
}

/**
 * Builds (or rebuilds) the "Source Sheet" tab - one row per distinct traffic
 * source (trackable link or UTM source), with views, visitors, retention,
 * and fill-rate broken out. Genuinely unbounded: column A spills exactly as
 * many rows as there are distinct sources (via the `#` dynamic-array spill
 * reference), and every other column rides that same spill range - no
 * hardcoded cap on how many sources can show up.
 */
function buildSourceSheet_(ss) {
  var old = ss.getSheetByName(SHEET_SOURCE);
  if (old) ss.deleteSheet(old);
  var s = ss.insertSheet(SHEET_SOURCE, 1);

  function label(a1, text, opts) {
    var r = s.getRange(a1);
    r.setValue(text);
    if (opts && opts.bold) r.setFontWeight('bold');
    if (opts && opts.size) r.setFontSize(opts.size);
    if (opts && opts.bg) r.setBackground(opts.bg);
    if (opts && opts.color) r.setFontColor(opts.color);
    if (opts && opts.italic) r.setFontStyle('italic');
  }
  function formula(a1, f) { s.getRange(a1).setFormula(f); }

  s.getRange('A1:L1').merge();
  label('A1', '🔗 Source Sheet - traffic broken out by source', { bold: true, size: 18, bg: BRAND, color: '#FFFFFF' });
  s.getRange('A1:L1').setVerticalAlignment('middle');
  s.setRowHeight(1, 40);
  s.getRange('A2:L2').merge();
  label('A2', 'One row per trackable link / UTM source - as many as actually exist, no cap. Recalculates automatically. "Returning" = a visitor from that source who had already been here before (tracked via a device-level ID).', { italic: true, color: '#6C6560' });

  var headers = ['Source', 'Total Views', 'Unique Visitors', 'New Visitors', 'Returning Visitors', 'Retention Rate', 'First Seen', 'Last Seen', 'Filled (Complete)', 'Partial (Started, Left)', 'Not Filled', 'Conversion Rate'];
  s.getRange('A4:L4').setValues([headers]).setFontWeight('bold').setFontSize(9).setFontColor('#6C6560').setBackground(BRAND_WASH);

  // Every distinct source seen in either Events or Submissions, sorted
  // alphabetically - this single spilled column drives the row count for
  // every column to its right.
  formula('A5', '=IFERROR(SORT(UNIQUE(FILTER({Events!F2:F;Submissions!E2:E},{Events!F2:F;Submissions!E2:E}<>""))),"No trackable links yet - try sharing cab.picapool.tech/instagram")');

  formula('B5', '=ARRAYFORMULA(COUNTIFS(Events!$F$2:$F$9999,A5#,Events!$B$2:$B$9999,"view"))');
  formula('C5', '=MAP(A5#,LAMBDA(src,IFERROR(COUNTA(UNIQUE(FILTER(Events!$C$2:$C$9999,Events!$F$2:$F$9999=src,Events!$B$2:$B$9999="view"))),0)))');
  formula('D5', '=MAP(A5#,LAMBDA(src,IFERROR(COUNTA(UNIQUE(FILTER(Events!$C$2:$C$9999,Events!$F$2:$F$9999=src,Events!$B$2:$B$9999="view",Events!$E$2:$E$9999="Yes"))),0)))');
  formula('E5', '=ARRAYFORMULA(C5#-D5#)');
  formula('F5', '=ARRAYFORMULA(IFERROR(E5#/C5#,0))');
  formula('G5', '=MAP(A5#,LAMBDA(src,IFERROR(MIN(FILTER(Events!$A$2:$A$9999,Events!$F$2:$F$9999=src)),"")))');
  formula('H5', '=MAP(A5#,LAMBDA(src,IFERROR(MAX(FILTER(Events!$A$2:$A$9999,Events!$F$2:$F$9999=src)),"")))');
  formula('I5', '=ARRAYFORMULA(COUNTIFS(Submissions!$E$2:$E$9999,A5#,Submissions!$AF$2:$AF$9999,"Complete"))');
  formula('J5', '=ARRAYFORMULA(COUNTIFS(Submissions!$E$2:$E$9999,A5#,Submissions!$AF$2:$AF$9999,"Partial"))');
  formula('K5', '=ARRAYFORMULA(IF((B5#-I5#-J5#)<0,0,(B5#-I5#-J5#)))');
  formula('L5', '=ARRAYFORMULA(IFERROR(I5#/B5#,0))');

  s.getRange('F5:F5000').setNumberFormat('0.0%');
  s.getRange('L5:L5000').setNumberFormat('0.0%');
  s.getRange('G5:H5000').setNumberFormat('yyyy-mm-dd hh:mm');

  s.setColumnWidths(1, 12, 130);
  s.setFrozenRows(4);
}

/** Builds (or rebuilds) the "Read Me" tab - a plain-English guide to the sheet. */
function buildReadme_(ss) {
  var old = ss.getSheetByName(SHEET_README);
  if (old) ss.deleteSheet(old);
  var r = ss.insertSheet(SHEET_README, 0);

  function h(a1, text, opts) {
    var rg = r.getRange(a1);
    rg.setValue(text);
    if (opts && opts.bold) rg.setFontWeight('bold');
    if (opts && opts.size) rg.setFontSize(opts.size);
    if (opts && opts.bg) rg.setBackground(opts.bg);
    if (opts && opts.color) rg.setFontColor(opts.color);
    if (opts && opts.italic) rg.setFontStyle('italic');
  }

  r.getRange('A1:F1').merge();
  h('A1', '📋 Read Me - how this sheet works', { bold: true, size: 18, bg: BRAND, color: '#FFFFFF' });
  r.setRowHeight(1, 40);
  r.getRange('A2:F2').merge();
  h('A2', 'Auto-generated by the Apps Script behind the intent form. Safe to rebuild any time from the PicaPool menu above - it only touches Dashboard + this tab, never Submissions or Events.', { italic: true, color: '#6C6560' });

  // Each row: [term, description]. A blank term = spacer row. A term with an
  // empty description = a section header (spans the full width).
  var rows = [
    ['WHERE THE DATA COMES FROM', ''],
    ['Nothing here is typed in by hand.', 'Every row is written automatically by the form at cab.picapool.tech the moment someone loads the page, leaves without finishing, or submits.'],
    ['', ''],
    ['THE FIVE TABS', ''],
    ['Dashboard', 'Live numbers - views, unique visitors, conversion, drop-offs, per-link performance. All formulas, so it updates itself. Nothing to edit here.'],
    ['Source Sheet', 'One row per traffic source (trackable link or UTM source) - as many as actually exist, no cap. Views, unique/new/returning visitors, retention rate, first/last seen dates, and filled/partial/not-filled counts, per source.'],
    ['Submissions', 'One row per completed form, PLUS one row for anyone who answered at least one question and then left without finishing. Every answer, plus who/where/when/how they found us.'],
    ['Events', 'One row per page view, and one per drop-off ("exit" = left without submitting). This raw data is what the Dashboard and Source Sheet are built from.'],
    ['Read Me', 'This tab.'],
    ['', ''],
    ['PARTIAL vs. COMPLETE (the color coding)', ''],
    ['Green rows, Status = "Complete"', 'They finished the form and hit submit. This is a real rider ready to be matched.'],
    ['Red rows, Status = "Partial"', 'They answered at least one question but left before submitting. Worth a follow-up nudge (their phone/name may already be filled in if they got that far) - these are the leads most likely to convert with one WhatsApp message.'],
    ['', ''],
    ['TRACKABLE LINKS', ''],
    ['Anything after the domain’s slash becomes a Source, automatically.', 'e.g. cab.picapool.tech/instagram, cab.picapool.tech/prakash, cab.picapool.tech/whatsapp-status - no per-link setup. It shows up as its own row in the Dashboard’s "By Trackable Link" table the moment someone opens it.'],
    ['Classic ?utm_source=... links work too.', 'Logged alongside the path-based source.'],
    ['', ''],
    ['READING THE DROP-OFF NUMBERS', ''],
    ['"Furthest Question Reached" (0–7)', '0 = landed on the page but hadn’t scrolled to Q1 yet. 7 = reached the final "personal details" section. Captured the moment someone leaves without submitting.'],
    ['"Approx. Drop-offs"', 'Views minus Submissions - a simple estimate, not a guarantee every non-submitter actually engaged with the form.'],
    ['', ''],
    ['DATA QUALITY NOTES', ''],
    ['IP / City / Region / Country', 'Looked up client-side via a free public geo-IP service - it’s an approximation. VPNs, mobile carriers, and campus Wi-Fi proxies can all skew it. Treat it as a rough signal, not a precise location.'],
    ['Unique Visitors vs. Total Visits', 'A "visitor" is one device/browser (tracked via an ID saved on that device). A "visit" is one session - the same person opening the page on two different days is one visitor, two visits.'],
    ['', ''],
    ['MAINTENANCE', ''],
    ['Rebuild Dashboard + Read Me', 'Sheet menu bar → PicaPool → Rebuild dashboard. Safe to run any time.'],
    ['Change what gets tracked', 'Edit google-apps-script.gs and index.html in the GitHub repo, re-paste the .gs file into Extensions → Apps Script, then Deploy → Manage deployments → edit → new version (keeps the same URL).'],
    ['Frontend + backend source code', 'github.com/latharrr/cab-intent-pp']
  ];

  var startRow = 4;
  for (var i = 0; i < rows.length; i++) {
    var rowIndex = startRow + i;
    var term = rows[i][0];
    var desc = rows[i][1];
    if (!term) continue; // spacer row
    if (!desc) {
      r.getRange(rowIndex, 1, 1, 6).merge();
      h('A' + rowIndex, term, { bold: true, bg: BRAND_WASH });
    } else {
      r.getRange(rowIndex, 1).setValue(term).setFontWeight('bold').setVerticalAlignment('top');
      r.getRange(rowIndex, 2, 1, 5).merge();
      r.getRange(rowIndex, 2).setValue(desc).setWrap(true).setVerticalAlignment('top');
    }
  }

  r.setColumnWidth(1, 260);
  r.setColumnWidths(2, 5, 140);
  r.setFrozenRows(2);
}

/** Removes the default empty "Sheet1" tab Google adds to every new spreadsheet, if untouched. */
function cleanupDefaultSheet_(ss) {
  var def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && def.getLastColumn() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(def);
  }
}
