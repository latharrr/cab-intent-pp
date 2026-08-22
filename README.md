# Picapool Cab Sharing Intent Form

A single-page intent form for Picapool's cab-sharing launch, built on the
app's real design system (colors, type, radii, and component patterns
extracted from the Picapool Figma file).

Static HTML, no build step. `index.html` is the whole site. The backend
is a Google Sheet — no database, no server.

## Deploy

Zero-config on Vercel — import this repo, framework preset "Other",
no build command, output directory `/`. `vercel.json` rewrites every path
to `index.html`, which is what makes trackable links (below) work.

## Notes

- Fonts load from Google Fonts (Inter, IBM Plex Mono).
- The Picapool logo and favicon are inlined as base64 data URIs directly
  in `index.html` — no separate asset files, consistent with the
  single-file setup.

## Backend: Google Sheet + Apps Script

Every page view, form submission, and drop-off gets logged straight to a
Google Sheet via a small Apps Script Web App (`google-apps-script.gs`).
No server, no API keys, no cost.

**One-time setup (~5 minutes):**

1. Create a new blank Google Sheet — this is your dashboard, name it
   whatever you like (e.g. "Picapool Route Match — Responses").
2. In the Sheet: **Extensions > Apps Script**. Delete the placeholder
   `Code.gs` content and paste in the entire contents of
   [`google-apps-script.gs`](./google-apps-script.gs) from this repo.
3. In the Apps Script editor toolbar, select the `setupDashboard`
   function from the dropdown and click **Run**. The first run will ask
   you to authorize it (it's only touching this one Sheet) — approve it.
   This builds the `Submissions`, `Events`, and `Dashboard` tabs for you.
4. **Deploy > New deployment**. Type: **Web app**. Execute as: **Me**.
   Who has access: **Anyone** (must be "Anyone", not "Anyone with a
   Google account" — the form's visitors aren't logged into Google).
   Click Deploy and copy the `.../exec` URL it gives you.
5. In `index.html`, find `CONFIG.APPS_SCRIPT_URL` near the top of the
   `<script>` block and paste that URL in. Redeploy the site.

That's it — submissions and page-view events now land in the Sheet in
real time. Re-run `setupDashboard` any time from the Sheet's new
**PicaPool** menu if you want to rebuild the Dashboard tab from scratch.

**What's in the Sheet:**

- **Submissions** — one row per completed form: every answer, plus
  visitor/session IDs, source, referrer, device, time-to-complete, and
  approximate IP/city/region/country.
- **Events** — one row per page view and one per "exit" (someone who
  left without submitting), with how far they got (which question,
  what % of fields filled) — this is what powers the drop-off numbers.
- **Dashboard** — live formulas on top of the two tabs above: total
  views/unique visitors/submissions/conversion rate, a per-trackable-link
  breakdown (views, submissions, conversion %, drop-offs, average
  drop-off point), a drop-off funnel by question, a live feed of recent
  submissions, and new-vs-returning / device / country breakdowns.

**Trackable links:** anything after the domain's slash is captured as a
`Source`, no setup needed per-link. Share `cab.picapool.tech/instagram`,
`cab.picapool.tech/prakash`, `cab.picapool.tech/whatsapp` — each shows up
as its own row in the Dashboard's "By Trackable Link" table with its own
views/submissions/conversion/drop-off numbers. Classic `?utm_source=`
params work too and are logged alongside.

IP-based city/region/country is looked up client-side (via a free public
API) since Apps Script can't see a visitor's IP directly — treat it as
an approximation (VPNs, mobile carriers, and proxies will skew it), not
a precise location.
