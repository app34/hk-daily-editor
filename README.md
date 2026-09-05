# HK Daily XLSM Editor

Web app that **updates your existing Housekeeping Daily `.xlsm` file**.

It does **not** create a new workbook. It opens the file you upload, changes only the cell values you edit, and saves **the same `.xlsm`** with:

- original colour fills / fonts
- merged cells, column widths, print setup
- drawings / logos
- VBA macros (`vbaProject.bin`)
- formulas (left alone unless you edit that cell)

Built as static files so you can put the folder on **GitHub Pages**.

## Why not “Save As” a new Excel file?

Libraries that rebuild an `.xlsx` from scratch drop or rewrite:

- VBA macros
- the original `styles.xml` (your colour codes)
- conditional formatting extensions
- drawings and printer settings

This app treats the `.xlsm` as a ZIP (that is what Excel files are), edits the worksheet XML + shared strings, and zips it back using the **same parts**.

## Deploy on GitHub

1. Create a repo (example name: `hk-daily-editor`).
2. Upload these files to the repo root:

   - `index.html`
   - `app.js`
   - `styles.css`
   - `README.md`

3. GitHub → **Settings → Pages → Deploy from branch → `main` / root**.
4. Open `https://YOUR_USER.github.io/hk-daily-editor/`.

Do **not** commit the live daily `.xlsm` if it has guest or staff phone data. Upload it in the browser each day.

## Daily use

1. Open the site (phone or desktop).
2. **Open XLSM** — pick `HK Daily AUGUST 2026.xlsm` (or the current month file).
3. Edit:
   - header / duty names
   - in-house counts
   - villa moves, arrivals, departures, honeymoon / birthday / anniversary
   - villa board statuses (ARR, DEP, OCC, B2B, VAC, OOO, …)
   - leave rows
   - MATCH and Supervisor Allocation sheets
4. **Download same XLSM** — same file name, same colours, macros kept.
5. Replace the file on the shared drive / WhatsApp it to the team.

Chrome / Edge / Safari / Android Chrome all work. No server and no login.

## Status colours in the app

These match the sheet legend and typical fills in your file:

| Code | Meaning | App colour |
|---|---|---|
| ARR | Arrival | green |
| DEP | Departure | red |
| OCC | Occupied | orange |
| B2B | Back to back | orange |
| VAC | Vacant | yellow |
| OOO | Out of order | dark red |
| OS | Out of service | grey |
| SV | Show villa | teal |
| STB | Stand by | blue |
| DU | Day use | purple |
| DC | Deep cleaning | navy |
| VM-ARR | Villa move arrival | green |

When you change a villa **STAT** value, the app also switches that cell to a style index already used in *your* file for that status, so Excel still shows the right fill.

## What the app will not do

- Run VBA inside the browser (Excel runs macros after you open the downloaded file).
- Recalculate every Excel formula in the browser. Totals that are formulas (`COUNTA`, `COUNTIF`, `TODAY()`) stay as formulas and update when opened in Excel.
- Invent a blank new workbook if you never upload a file.

## Local test

Open `index.html` in a browser, or from this folder:

```bash
python3 -m http.server 8080
```

Then visit `http://localhost:8080`.
