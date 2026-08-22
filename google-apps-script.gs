/**
 * PicaPool intent-form logger + live dashboard.
 *
 * This file does NOT run on Vercel — it runs inside a Google Sheet.
 * Setup (see README.md for the full walkthrough):
 *   1. Create a blank Google Sheet.
 *   2. Extensions > Apps Script, delete the placeholder code, paste this whole file.
 *   3. Run `setupDashboard` once from the editor toolbar (grant the permissions it asks for).
 *   4. Deploy > New deployment > type "Web app". Execute as "Me", access "Anyone".
 *   5. Copy the /exec URL it gives you into CONFIG.APPS_SCRIPT_URL in index.html.
 *
 * Re-run `setupDashboard` any time from the sheet's "PicaPool" menu to rebuild
 * the Dashboard tab (e.g. after you change the row layout below).
 */

var SHEET_SUBMISSIONS = 'Submissions';
var SHEET_EVENTS = 'Events';
var SHEET_DASHBOARD = 'Dashboard';
var BRAND = '#F03506';

var SUBMISSION_HEADERS = [
  'Timestamp', 'Visitor ID', 'Session ID', 'Is Returning Visitor', 'Source', 'Landing Path',
  'Full Name', 'College', 'Phone', 'Email', 'Travel Mode', 'Daily Spend', 'Commute Time',
  'Nearest Metro Station', 'Would Prefer Cab Share', 'Commute Effort', 'Time To Complete (s)',
  'Referrer', 'UTM Source', 'UTM Medium', 'UTM Campaign',
  'Device Type', 'Browser', 'OS', 'Language', 'Screen',
  'IP Address', 'City', 'Region', 'Country', 'ISP / Org'
]; // A..AE (31 cols)

var EVENT_HEADERS = [
  'Timestamp', 'Event Type', 'Visitor ID', 'Session ID', 'Is New Visitor', 'Source', 'Landing Path',
  'Furthest Question Reached', 'Percent Fields Filled', 'Time On Page (s)',
  'Referrer', 'UTM Source', 'UTM Medium', 'UTM Campaign',
  'Device Type', 'Browser', 'OS', 'Language', 'Screen',
  'IP Address', 'City', 'Region', 'Country', 'ISP / Org'
]; // A..X (24 cols)

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.kind === 'submission') {
      appendSubmission(p);
    } else if (p.kind === 'view' || p.kind === 'exit') {
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
  SpreadsheetApp.getUi().addMenu('PicaPool', [{ name: 'Rebuild dashboard', functionName: 'setupDashboard' }]);
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

function appendSubmission(p) {
  var sheet = getOrCreateSheet_(SHEET_SUBMISSIONS, SUBMISSION_HEADERS);
  sheet.appendRow([
    new Date(), p.visitorId || '', p.sessionId || '', p.isNewVisitor ? 'No' : 'Yes', p.source || '(direct)', p.path || '',
    p.fullName || '', p.college || '', p.phone || '', p.email || '', p.travelMode || '', p.spend || '', p.commuteTime || '',
    p.startArea || '', p.wouldPrefer || '', p.effort || '', p.timeToCompleteSec || '',
    p.referrer || '', p.utmSource || '', p.utmMedium || '', p.utmCampaign || '',
    p.deviceType || '', p.browser || '', p.os || '', p.language || '', p.screen || '',
    p.ip || '', p.city || '', p.region || '', p.country || '', p.isp || ''
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
 * Builds (or rebuilds) the Dashboard tab: headline metrics, per-trackable-link
 * funnel, drop-off-by-question, live submissions feed, and device/country
 * breakdowns — all formulas, so it stays live as new rows come in.
 */
function setupDashboard() {
  getOrCreateSheet_(SHEET_SUBMISSIONS, SUBMISSION_HEADERS);
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
  label('A1', '🚕 PicaPool — Live Dashboard', { bold: true, size: 18, bg: BRAND, color: '#FFFFFF' });
  d.getRange('A1:H1').setVerticalAlignment('middle');
  d.setRowHeight(1, 40);
  label('A2', 'Recalculates automatically whenever this sheet is open. Trackable links: share cab.picapool.tech/<anything> and it shows up below as a Source.', { italic: true, color: '#6C6560' });
  d.getRange('A2:H2').merge();

  // ---- Headline metrics ----
  d.getRange('A4:H4').merge();
  label('A4', 'HEADLINE METRICS', { bold: true, bg: '#FFE5DA' });
  var metricHeaders = ['Total Views', 'Unique Visitors', 'Total Visits (Sessions)', 'Total Submissions', 'Conversion Rate', 'Avg. Time on Page — Exits (s)', 'Avg. Time to Complete (s)', 'Approx. Drop-offs'];
  d.getRange('A5:H5').setValues([metricHeaders]).setFontWeight('bold').setFontSize(9).setFontColor('#6C6560');
  formula('A6', '=COUNTIF(Events!B:B,"view")');
  formula('B6', '=IFERROR(COUNTA(UNIQUE(FILTER(Events!C2:C,Events!B2:B="view"))),0)');
  formula('C6', '=IFERROR(COUNTA(UNIQUE(FILTER(Events!D2:D,Events!B2:B="view"))),0)');
  formula('D6', '=COUNTA(Submissions!B2:B)');
  formula('E6', '=IFERROR(D6/A6,0)');
  formula('F6', '=IFERROR(ROUND(AVERAGE(FILTER(Events!J2:J,Events!B2:B="exit",Events!J2:J<>"")),1),0)');
  formula('G6', '=IFERROR(ROUND(AVERAGE(FILTER(Submissions!Q2:Q,Submissions!Q2:Q<>"")),1),0)');
  formula('H6', '=MAX(A6-D6,0)');
  d.getRange('A6:H6').setFontSize(16).setFontWeight('bold');
  d.getRange('E6').setNumberFormat('0.0%');

  // ---- By trackable link ----
  d.getRange('A8:F8').merge();
  label('A8', 'BY TRACKABLE LINK (SOURCE)', { bold: true, bg: '#FFE5DA' });
  var sourceHeaders = ['Source', 'Views', 'Submissions', 'Conversion %', 'Approx. Drop-offs', 'Avg. Question Reached at Exit (0–7)'];
  d.getRange('A9:F9').setValues([sourceHeaders]).setFontWeight('bold').setFontSize(9).setFontColor('#6C6560');
  formula('A10', '=IFERROR(SORT(UNIQUE(FILTER(Events!F2:F,Events!F2:F<>""))),"No trackable links yet — try sharing cab.picapool.tech/instagram")');
  formula('B10', '=ARRAYFORMULA(IF($A10:$A59="","",COUNTIFS(Events!$F$2:$F$9999,$A10:$A59,Events!$B$2:$B$9999,"view")))');
  formula('C10', '=ARRAYFORMULA(IF($A10:$A59="","",COUNTIFS(Submissions!$E$2:$E$9999,$A10:$A59)))');
  formula('D10', '=ARRAYFORMULA(IF($A10:$A59="","",IFERROR($C10:$C59/$B10:$B59,0)))');
  formula('E10', '=ARRAYFORMULA(IF($A10:$A59="","",IF($B10:$B59-$C10:$C59>0,$B10:$B59-$C10:$C59,0)))');
  formula('F10', '=ARRAYFORMULA(IF($A10:$A59="","",IFERROR(AVERAGEIFS(Events!$H$2:$H$9999,Events!$F$2:$F$9999,$A10:$A59,Events!$B$2:$B$9999,"exit"),"")))');
  d.getRange('D10:D59').setNumberFormat('0.0%');

  // ---- Drop-off funnel ----
  label('A62', 'DROP-OFF FUNNEL — where riders leave the form', { bold: true, bg: '#FFE5DA' });
  d.getRange('A62:D62').merge();
  formula('A63', '=IFERROR(QUERY(Events!A2:X,"select H, count(H) where B=\'exit\' group by H order by H asc label H \'Furthest Question Reached (0-7)\', count(H) \'Visitors\'"),"No drop-off data yet")');

  // ---- Recent submissions ----
  label('A74', 'RECENT SUBMISSIONS (live feed)', { bold: true, bg: '#FFE5DA' });
  d.getRange('A74:F74').merge();
  formula('A75', '=IFERROR(QUERY(Submissions!A2:U,"select A,G,H,E,Q order by A desc limit 15 label A \'When\', G \'Name\', H \'College\', E \'Source\', Q \'Took (s)\'"),"No submissions yet")');

  // ---- New vs returning, device breakdown ----
  label('A95', 'NEW VS RETURNING VISITORS', { bold: true, bg: '#FFE5DA' });
  label('D95', 'VISITS BY DEVICE TYPE', { bold: true, bg: '#FFE5DA' });
  formula('A96', '=IFERROR(QUERY(Events!A2:X,"select E, count(E) where B=\'view\' group by E label E \'New Visitor?\', count(E) \'Views\'"),"No data yet")');
  formula('D96', '=IFERROR(QUERY(Events!A2:X,"select O, count(O) where B=\'view\' and O is not null and O <> \'\' group by O order by count(O) desc label O \'Device\', count(O) \'Views\'"),"No data yet")');

  // ---- Country breakdown (variable height — placed last) ----
  label('A106', 'VISITS BY COUNTRY (top locations)', { bold: true, bg: '#FFE5DA' });
  d.getRange('A106:D106').merge();
  formula('A107', '=IFERROR(QUERY(Events!A2:X,"select W, count(W) where B=\'view\' and W is not null and W <> \'\' group by W order by count(W) desc label W \'Country\', count(W) \'Views\'"),"No location data yet")');

  d.setColumnWidths(1, 8, 150);
  d.setFrozenRows(2);
  SpreadsheetApp.flush();
}
