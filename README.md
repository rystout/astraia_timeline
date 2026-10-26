# Astraia — Master Timeline

An interactive, offline-capable browser page for the chronology of the Royal
Commonwealth of Zeythari, generated from `Astraia_Master_Timeline.xlsx`.

Open `index.html` — that's the whole site. No server, build step, framework,
font download or network access is involved.

## What's on the page

- **Timeline** — all 42 dated events, grouped under their era, with year range,
  duration, thread, certainty, peoples/places, notes and source. Search covers
  events, notes and sources; filters cover threads and certainty.
- **One switch for the reckoning** — every date can be read as A.S. (negative
  years as P.A.), A.K. (negatives as B.K.), years before present, or lifetimes ago.
- **Eras & spans** — the workbook's spans on a proportional A.S. axis; overlapping
  spans get their own line. Click a span to pull up the events inside it.
- **Converter** — year + reckoning in, A.S./A.K./years before present/lifetimes
  and the eras at that date out. The calibration panel (present year, A.S. of
  1 A.K., lifespan) recalculates those figures everywhere and is remembered in
  `localStorage`.
- **Open questions** — the nine unresolved items with why each matters.
- **Export** — the currently filtered events as CSV (spreadsheet-ready, UTF-8)
  or JSON (all fields plus the active filters and calibration).

## Files

| File | Purpose |
| --- | --- |
| `index.html` | The generated site. This is the only file GitHub Pages needs. |
| `build.py` | Reads the workbook and writes `index.html`. |
| `Astraia_Master_Timeline.xlsx` | Source of truth: Timeline, Eras & Spans, Settings, Converter, Open Questions. |

## Editing

Edit the spreadsheet, then regenerate the page:

```sh
python3 -m venv .venv
.venv/bin/pip install openpyxl
.venv/bin/python build.py
```

`build.py` takes optional input/output paths: `python3 build.py book.xlsx site.html`.

Do not hand-edit `index.html` — the generator overwrites it. Text, dates, notes
and spans all come from the workbook; the page's HTML/CSS/JS lives in `build.py`.

## Publishing on GitHub Pages

1. Push this repository to GitHub.
2. **Settings → Pages → Build and deployment**.
3. Source: **Deploy from a branch**; branch **main**, folder **/ (root)**; save.
4. The page is served at `https://<user>.github.io/<repo>/`.

No Actions workflow is needed: `index.html` is complete and self-contained, and
Jekyll passes it through untouched.

## Local screenshots (optional)

`.screenshots/` holds reference captures and sample exports and is git-ignored.
To take new ones, install Playwright's Chromium once, then use its CLI:

```sh
npx playwright install chromium
npx playwright screenshot --browser=chromium --viewport-size="390,844" \
  --full-page --wait-for-timeout=800 "file://$PWD/index.html" shot.png
```

Append `#eras`, `#converter` or `#questions` to the URL to capture another tab.
