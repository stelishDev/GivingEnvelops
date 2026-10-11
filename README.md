# Giving Envelopes

A small installable web app (PWA) for Mt Roskill Evangelical Fellowship. It reads the
envelopes spreadsheet, lets you tick the givers you want, and prints **4 freewill giving
envelopes per A4 sheet** with each giver's name under the fellowship title.

No server, no accounts, no dependencies. The spreadsheet never leaves the computer.

## Using it

1. Choose (or drag in) the spreadsheet — `.xlsx` or `.csv`. Columns are detected from the
   header row (`Christian Name:`, `Surname:`, `CONG`); change them in step 2 if needed.
2. Check the envelope date (defaults to the coming Sunday).
3. Tick the people you want (search, filter by cong. number, "Select shown", "Clear all").
4. **Print envelopes.** In the print dialog use **A4, landscape, scale 100% / Actual size,
   margins None**. Cut along the dashed guides (switch them off with *Cut guides*).

Each envelope shows: cong. number and name (`9  Abel Anderson`), the date, the six fund
lines with Cash / D/C boxes, and a Total line — matching `Freewill Giving Envelopes.docx`.

The last loaded list is remembered in the browser; *Forget saved list* removes it. Choose the
file again after you edit the spreadsheet.

## Running / hosting

It is plain static files. For testing: `python3 -m http.server 8000` then open
<http://localhost:8000>. Installing as a desktop app (and offline use) needs HTTPS or
localhost — GitHub Pages works well: repo **Settings → Pages → Deploy from branch → main / root**.
Then open the Pages URL in Chrome or Edge and use **Install app** (or the install icon in the
address bar).

Once installed, `.xlsx` / `.csv` files can also be opened with the app via *Open with*
(Chromium desktop).

## Files

| File | Purpose |
|---|---|
| `index.html`, `styles.css`, `app.js` | The app and the envelope/A4 print layout |
| `xlsx-reader.js` | Built-in `.xlsx`/`.csv` reader (no libraries) |
| `manifest.webmanifest`, `sw.js`, `icons/` | PWA install + offline support |
| `tests/` | `node tests/test-reader.mjs` checks the reader against a sample sheet |

When you change any app file, bump `CACHE` in `sw.js` so installed copies refresh.

## Adjusting the envelope

Sizes and text live in `styles.css` (`.sheet`, `.envelope`, `.row`) and the `FUNDS` list at the
top of `app.js`. The sheet is A4 landscape split 2 × 2, so each envelope is 148.5 × 105 mm.
