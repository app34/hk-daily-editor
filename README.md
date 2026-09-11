# HK Daily

Phone PWA for the orange **Housekeeping Daily .xlsm** — same pattern as Minibar.

Open the live workbook once. It stays on the phone. Edit villa status, duty names, forecast and allocation, then **Download** the same `.xlsm` (colours + macros kept).

**Live app:** https://app34.github.io/hk-daily-editor/

Install: open the link in Chrome / Safari → **Add to Home Screen**.

## Daily use

1. Open the app (home screen icon or the GitHub Pages link).
2. First time: **Choose HK Daily .xlsm**. After that, the last saved file opens by itself.
3. **To today** — roll yesterday ARR/B2B/VM-ARR → OCC and DEP → VAC.
4. **From text** — paste today’s arrival / departure / room-move lists.
5. Edit the full daily sheet, attendant cards, or status board.
6. **Save** stores the file on this phone. **Download** writes the same XLSM back.

Do **not** commit the live `.xlsm` to GitHub if it has staff or guest data.

## Why this does not build a new Excel file

The orange workbook is a template:

- VBA macros (`vbaProject.bin`)
- colour fills / legend
- merged headers, logos, print setup

The app treats the `.xlsm` as a ZIP, patches worksheet cell values, and zips the **same parts** back. Same approach as Minibar consumption.

## PWA files

| File | Role |
| --- | --- |
| `index.html` | Shell, install button, version check |
| `styles.css` | Same UI as the orange editor |
| `app.js` | Import / edit / IndexedDB / XLSM write-back |
| `manifest.json` | Add to Home Screen |
| `sw.js` | Offline app shell + SheetJS / JSZip cache |
| `version.json` | Force phones onto the latest build |
| `icons/` | Home screen icons |

Host as static files on **GitHub Pages**. No backend and no login.

## Deploy

Repo already has Pages on `main` / root:

https://app34.github.io/hk-daily-editor/

After a push, wait about a minute, then hard-refresh the phone (`?v=` is added automatically when `version.json` changes).

## Local test

```bash
python3 -m http.server 8080
```

Open `http://localhost:8080`.
