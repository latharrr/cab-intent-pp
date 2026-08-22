// Same-origin relay for the intent-form's analytics beacon.
//
// The browser used to POST straight to script.google.com, which privacy
// blocklists (Brave Shields, uBlock Origin, Safari ITP, etc.) commonly block
// outright, since a Google Apps Script Web App used as a client-side beacon
// endpoint is a known ad-blocker-evasion pattern - so the request never left
// the visitor's browser and nothing landed in the Sheet, even though the
// form's success screen (which never depends on the beacon) still showed.
//
// The browser now POSTs here instead (same origin as cab.picapool.tech,
// which nothing blocks), and this function forwards to Apps Script
// server-side, where no browser extension can interfere.
var APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyT7erIbgapCQ7Emem0VtR-RIrKD99-LQy5bkqqmmY32HDQyYOM_SwlLT09lHBt23PE/exec';

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).end();
    return;
  }
  try {
    var body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body || {});
    await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
      body: body
    });
  } catch (err) {
    // analytics must never break the form
  }
  res.status(204).end();
};
