# Picapool Cab Sharing Intent Form

A single-page intent form for Picapool's cab-sharing launch, built on the
app's real design system (colors, type, radii, and component patterns
extracted from the Picapool Figma file).

Static HTML, no build step, no backend. `index.html` is the whole site.

## Deploy

Zero-config on Vercel — import this repo, framework preset "Other",
no build command, output directory `/`.

## Notes

- Form submission is currently client-side only (shows a success state on
  submit). Wire it to a sheet, database, or WhatsApp API to actually
  collect responses.
- Fonts load from Google Fonts (Inter, IBM Plex Mono).
