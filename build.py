#!/usr/bin/env python3
"""Build a single self-contained, interactive timeline page from the workbook.

    python3 build.py Astraia_Master_Timeline.xlsx index.html

The output has no external dependencies: open it from disk or serve it from
GitHub Pages as-is. All data is embedded, so the page works offline.

Requires openpyxl:  pip install openpyxl
"""
import json
import sys

import openpyxl

TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Astraia — Master Timeline</title>
<style>
:root {
  --bg:#f3efe6; --card:#fffdf8; --ink:#1a1815; --muted:#6c6456; --faint:#8d8474;
  --line:#e4ddce; --line-soft:#efe9dc; --gold:#9a7a2f; --gold-soft:#c9a44a;
  --deep:#0b1020; --deep-2:#141a2c; --add:#2f6f4f;
  --serif:"Iowan Old Style","Palatino Linotype",Palatino,Georgia,"Times New Roman",serif;
  --sans:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
  --shadow:0 1px 2px rgba(26,24,21,.05),0 10px 30px -18px rgba(26,24,21,.25);
}
* { box-sizing:border-box; }
html { scroll-behavior:smooth; }
body { margin:0; background:var(--bg); color:var(--ink); font:16px/1.62 var(--sans);
  -webkit-font-smoothing:antialiased; }
a { color:inherit; }
button, input, select { font:inherit; color:inherit; }
:focus-visible { outline:2px solid var(--gold); outline-offset:2px; }

/* ---------- hero ---------- */
.hero { position:relative; overflow:hidden; background:
    radial-gradient(900px 420px at 82% -18%, rgba(201,164,74,.30), transparent 62%),
    radial-gradient(620px 320px at 4% 120%, rgba(96,120,196,.24), transparent 65%),
    linear-gradient(160deg,#0b1020 0%,#111829 58%,#0a0f1c 100%);
  color:#efe9da; }
.hero::before { content:""; position:absolute; inset:0; pointer-events:none; opacity:.55;
  background-image:
    radial-gradient(1.4px 1.4px at 12% 22%, #fff 50%, transparent 51%),
    radial-gradient(1.1px 1.1px at 28% 68%, #fff 50%, transparent 51%),
    radial-gradient(1.6px 1.6px at 44% 12%, #fff 50%, transparent 51%),
    radial-gradient(1px 1px at 58% 44%, #fff 50%, transparent 51%),
    radial-gradient(1.5px 1.5px at 71% 26%, #fff 50%, transparent 51%),
    radial-gradient(1.1px 1.1px at 84% 62%, #fff 50%, transparent 51%),
    radial-gradient(1.3px 1.3px at 92% 18%, #fff 50%, transparent 51%),
    radial-gradient(1px 1px at 38% 84%, #fff 50%, transparent 51%),
    radial-gradient(1.2px 1.2px at 66% 88%, #fff 50%, transparent 51%),
    radial-gradient(1px 1px at 6% 52%, #fff 50%, transparent 51%);
  background-size:520px 380px; }
.wrap { max-width:1120px; margin:0 auto; padding:0 26px; }
.hero .wrap { position:relative; padding-top:46px; padding-bottom:34px; }
.kicker { margin:0 0 10px; font-size:11.5px; letter-spacing:.24em; text-transform:uppercase;
  color:var(--gold-soft); }
.hero h1 { margin:0; font:400 clamp(38px,6vw,58px)/1.02 var(--serif); letter-spacing:.01em; }
.hero h1 em { font-style:italic; color:#f0dfb6; }
.lede { margin:14px 0 0; max-width:64ch; color:#c5c8d4; font-size:15.5px; }
.hero-stats { display:flex; flex-wrap:wrap; gap:26px 38px; margin-top:26px;
  padding-top:20px; border-top:1px solid rgba(255,255,255,.12); }
.hero-stats div { min-width:74px; }
.hero-stats .v { font:400 21px/1.2 var(--serif); color:#f6efdd; font-variant-numeric:tabular-nums; }
.hero-stats .k { font-size:10.5px; letter-spacing:.13em; text-transform:uppercase;
  color:#959db1; margin-top:3px; }
.hero-stats .now { margin-left:auto; text-align:right; font-size:12.5px; color:#a9b2c6; align-self:flex-end; }
.hero-stats .now b { color:#efe3c6; font-weight:600; }

/* ---------- tabs ---------- */
.tabbar { position:sticky; top:0; z-index:30; background:rgba(243,239,230,.92);
  backdrop-filter:blur(9px); border-bottom:1px solid var(--line); }
.tabbar .wrap { display:flex; gap:4px; overflow-x:auto; scrollbar-width:none; }
.tabbar .wrap::-webkit-scrollbar { display:none; }
.tab { appearance:none; background:none; border:0; border-bottom:2px solid transparent;
  padding:13px 13px 11px; font-size:13.5px; color:var(--muted); cursor:pointer; white-space:nowrap; }
.tab:hover { color:var(--ink); }
.tab[aria-selected="true"] { color:var(--ink); border-bottom-color:var(--gold); font-weight:600; }

main { padding:26px 0 70px; }
.view[hidden] { display:none; }
.panel { background:var(--card); border:1px solid var(--line); border-radius:13px;
  box-shadow:var(--shadow); padding:20px 22px; }

/* ---------- toolbar ---------- */
.toolbar { position:sticky; top:48px; z-index:20; background:rgba(243,239,230,.94);
  backdrop-filter:blur(8px); padding:10px 0 8px; }
.toolbar-row { display:flex; gap:9px; flex-wrap:wrap; align-items:center; }
.search { position:relative; flex:1 1 260px; min-width:190px; }
.search input { width:100%; padding:9px 12px 9px 34px; border:1px solid #d4cbb9;
  border-radius:9px; background:#fffdf8; font-size:14px; }
.search svg { position:absolute; left:11px; top:50%; transform:translateY(-50%);
  width:15px; height:15px; fill:none; stroke:#8d8474; stroke-width:1.8; }
select.mode { padding:9px 11px; border:1px solid #d4cbb9; border-radius:9px; background:#fffdf8;
  font-size:13.5px; }
.ghost { padding:8px 13px; border:1px solid #d4cbb9; border-radius:9px; background:transparent;
  font-size:13px; cursor:pointer; color:var(--muted); }
.ghost:hover { border-color:#a89c84; color:var(--ink); }
.ghost:disabled { opacity:.45; cursor:not-allowed; }
.ghost:disabled:hover { border-color:#d4cbb9; color:var(--muted); }
.menu-wrap { position:relative; }
.menu { position:absolute; right:0; top:calc(100% + 7px); z-index:40; min-width:238px;
  background:#fffdf8; border:1px solid var(--line); border-radius:11px; padding:5px;
  box-shadow:0 18px 40px -18px rgba(26,24,21,.45); }
.menu[hidden] { display:none; }
.menu .mhead { padding:7px 10px 6px; font-size:10px; letter-spacing:.1em; text-transform:uppercase;
  color:var(--faint); }
.menu button { display:flex; flex-direction:column; gap:1px; width:100%; text-align:left;
  background:none; border:0; border-radius:7px; padding:8px 10px; cursor:pointer; font-size:13px; }
.menu button:hover, .menu button:focus-visible { background:#f4efe3; }
.menu button small { font-size:11px; color:var(--muted); }
.chips { display:flex; flex-wrap:wrap; gap:6px; margin-top:9px; }
.chip { display:inline-flex; align-items:center; gap:7px; padding:4px 11px 4px 9px;
  border:1px solid #d8cfbd; border-radius:999px; background:#fffdf8; font-size:12.5px;
  cursor:pointer; color:#3a352c; }
.chip:hover { border-color:#b3a68c; }
.chip[aria-pressed="true"] { background:#1a1815; border-color:#1a1815; color:#f7f2e6; }
.chip i { width:8px; height:8px; border-radius:50%; background:#999; flex:none; }
.chip span { color:var(--faint); font-variant-numeric:tabular-nums; font-size:11.5px; }
.chip[aria-pressed="true"] span { color:#b9b2a3; }
.hits { font-size:12.5px; color:var(--muted); margin:14px 0 0; font-variant-numeric:tabular-nums; }
.hits b { color:var(--ink); font-weight:600; }

/* ---------- span filter banner ---------- */
.spanbar { display:flex; align-items:center; gap:12px; flex-wrap:wrap; margin:14px 0 0;
  padding:10px 14px; background:#fbf6e6; border:1px solid #e7d9ac; border-radius:10px;
  font-size:13.5px; }
.spanbar b { font-weight:600; }
.spanbar button { margin-left:auto; }

/* ---------- timeline ---------- */
.era-head { display:flex; align-items:baseline; gap:12px; flex-wrap:wrap;
  margin:38px 0 2px; padding-bottom:7px; border-bottom:2px solid #241f18; }
.era-head:first-of-type { margin-top:26px; }
.era-head h2 { margin:0; font:400 23px/1.2 var(--serif); letter-spacing:.01em; }
.era-head .range { font-size:12.5px; color:var(--muted); font-variant-numeric:tabular-nums; }
.era-head .count { margin-left:auto; font-size:11px; letter-spacing:.09em; text-transform:uppercase;
  color:var(--faint); }
.ev { display:grid; grid-template-columns:190px 1fr; gap:20px; padding:15px 4px 16px;
  border-bottom:1px solid var(--line-soft); }
.ev:hover { background:rgba(255,255,255,.55); }
.ev-when { text-align:right; padding-top:1px; }
.ev-when .when { display:inline-block; font:400 15px/1.35 var(--serif); color:#221e18;
  font-variant-numeric:tabular-nums; }
.ev-when .dur { display:block; margin-top:3px; font-size:11px; letter-spacing:.06em;
  text-transform:uppercase; color:var(--faint); }
.ev-tags { display:flex; align-items:center; gap:9px; flex-wrap:wrap; }
.thread { display:inline-flex; align-items:center; gap:6px; background:none; border:0; padding:0;
  font-size:11px; letter-spacing:.1em; text-transform:uppercase; color:var(--c,#5c554a);
  cursor:pointer; font-weight:600; }
.thread i { width:7px; height:7px; border-radius:50%; background:var(--c,#5c554a); }
.thread:hover { text-decoration:underline; text-underline-offset:2px; }
.cert { font-size:10px; letter-spacing:.09em; text-transform:uppercase; padding:1px 7px;
  border-radius:999px; border:1px solid currentColor; opacity:.72; }
.cert-fixed { color:#2f6f4f; }
.cert-approx { color:#8a6a20; }
.cert-estimated { color:#7a5c48; }
.cert-derived { color:#54617a; }
.eid { font-size:10.5px; color:#b6ac98; font-variant-numeric:tabular-nums; }
.ev-text { margin:6px 0 0; font-size:15.5px; line-height:1.55; font-weight:500; }
.ev-ppl { margin:5px 0 0; font-size:12.5px; color:var(--muted); }
.ev-note { margin:7px 0 0; padding-left:10px; border-left:2px solid #e6dcc6;
  font-size:13px; color:#5f5849; }
.ev-src { margin:6px 0 0; font-size:11.5px; color:#9a917f; }
.ev.is-now { background:linear-gradient(90deg,rgba(201,164,74,.14),transparent 62%);
  border-bottom-color:#e3d3a5; }
.ev.is-now .ev-text { color:#5c4510; }
mark { background:#f5e2a5; color:inherit; border-radius:2px; padding:0 1px; }
.empty { text-align:center; color:var(--muted); padding:52px 20px; font-size:14.5px; }
.empty b { display:block; font:400 20px/1.3 var(--serif); color:var(--ink); margin-bottom:4px; }

/* ---------- eras view ---------- */
.gantt { margin-top:6px; }
.gantt-axis { display:grid; grid-template-columns:220px 1fr; align-items:end; }
.gantt-axis .ticks { position:relative; height:30px; border-bottom:1px solid #d9d1bf; }
.gantt-axis .tick { position:absolute; bottom:4px; font-size:10.5px; color:var(--faint);
  font-variant-numeric:tabular-nums; white-space:nowrap; }
.gantt-axis .tick::after { content:""; position:absolute; left:50%; bottom:-5px; width:1px;
  height:4px; background:#d9d1bf; }
.gantt-row { display:grid; grid-template-columns:220px 1fr; align-items:center; }
.gantt-row .lbl { font-size:12.5px; color:#3a352c; padding:5px 14px 5px 0; text-align:right;
  line-height:1.35; }
.gantt-row .track { position:relative; height:28px; border-bottom:1px solid #f0ebdf; }
.bar { position:absolute; top:6px; height:16px; min-width:4px; border:0; border-radius:4px;
  background:var(--c,#8a7a5c); opacity:.85; cursor:pointer; padding:0; text-align:left;
  overflow:hidden; }
.bar:hover { opacity:1; box-shadow:0 0 0 2px rgba(154,122,47,.35); }
.bar span { position:absolute; left:5px; top:50%; transform:translateY(-50%); font-size:10px;
  color:#fff; white-space:nowrap; letter-spacing:.02em; font-variant-numeric:tabular-nums;
  text-shadow:0 1px 1px rgba(0,0,0,.35); }
.marker { position:absolute; top:0; bottom:0; width:0; border-left:1px dashed #c9b98f; }
.marker.ak { border-left-color:#b58a3a; }
.marker.now { border-left:1px solid #9a907c; }
.sw { display:inline-block; width:13px; height:0; border-top:1px dashed #b58a3a;
  vertical-align:middle; margin-right:5px; }
.sw.now { border-top:1px solid #9a907c; }
.gantt-axis .tick.m { display:none; }
@media (max-width:760px) { .gantt-axis .tick { display:none; }
  .gantt-axis .tick.m3 { display:block; } }
.gantt-key { display:flex; gap:18px; flex-wrap:wrap; font-size:11.5px; color:var(--muted);
  margin-top:12px; }
.gantt-key i { display:inline-block; width:14px; border-top:1px dashed #b58a3a; vertical-align:middle;
  margin-right:5px; }
table.spans { width:100%; border-collapse:collapse; font-size:13.5px; margin-top:6px; }
table.spans th { text-align:left; font-size:10.5px; letter-spacing:.09em; text-transform:uppercase;
  color:var(--faint); font-weight:600; padding:8px 10px; border-bottom:1px solid var(--line); }
table.spans td { padding:9px 10px; border-bottom:1px solid var(--line-soft); vertical-align:top; }
table.spans tr:last-child td { border-bottom:0; }
table.spans .num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
table.spans .dim { color:var(--muted); font-size:12.5px; }
h2.viewtitle { font:400 26px/1.2 var(--serif); margin:6px 0 2px; }
p.viewlede { margin:0 0 18px; color:var(--muted); font-size:14px; max-width:74ch; }

/* ---------- converter ---------- */
.conv-grid { display:grid; grid-template-columns:1.15fr .85fr; gap:16px; align-items:start; }
.conv-inputs { display:flex; gap:12px; flex-wrap:wrap; margin-top:4px; }
.field { display:flex; flex-direction:column; gap:5px; }
.field span { font-size:10.5px; letter-spacing:.1em; text-transform:uppercase; color:var(--faint); }
.field input, .field select { padding:9px 11px; border:1px solid #d4cbb9; border-radius:8px;
  background:#fff; font-size:15px; font-variant-numeric:tabular-nums; width:150px; }
.field select#cvRec { width:206px; font-size:14px; }
.field input[type=number] { -moz-appearance:textfield; }
.conv-out { margin-top:18px; border-top:1px solid var(--line); }
.conv-row { display:flex; justify-content:space-between; gap:16px; align-items:baseline;
  padding:11px 2px; border-bottom:1px solid var(--line-soft); }
.conv-row .k { font-size:12px; letter-spacing:.06em; text-transform:uppercase; color:var(--muted); }
.conv-row .v { font:400 19px/1.2 var(--serif); font-variant-numeric:tabular-nums; text-align:right; }
.conv-row .v small { font-family:var(--sans); font-size:11.5px; color:var(--muted); }
.hint { font-size:12px; color:var(--faint); margin:14px 0 0; }
.calib .field input { width:120px; }
.calib .field { margin-bottom:12px; }
.calib .field em { font-style:normal; font-size:11.5px; color:var(--faint); }

/* ---------- questions ---------- */
.q { display:grid; grid-template-columns:38px 1fr; gap:14px; padding:16px 0;
  border-bottom:1px solid var(--line-soft); }
.q:last-child { border-bottom:0; }
.q .n { font:400 17px/1.4 var(--serif); color:var(--gold); font-variant-numeric:tabular-nums; }
.q h3 { margin:0; font-size:15.5px; font-weight:600; }
.q p { margin:4px 0 0; font-size:13.5px; color:var(--muted); }
.q .where { margin-top:6px; font-size:11.5px; color:var(--faint); }

/* ---------- footer ---------- */
footer { border-top:1px solid var(--line); padding:22px 0 54px; color:var(--faint); font-size:12.5px; }
.legend { margin-bottom:14px; }
.legend summary { cursor:pointer; font-size:12.5px; color:var(--muted); }
.legend dl { display:grid; grid-template-columns:max-content 1fr; gap:5px 16px; margin:12px 0 0;
  font-size:12.5px; }
.legend dt { color:#4b4539; }
.legend dd { margin:0; color:var(--faint); }

@media (max-width:760px) {
  .wrap { padding:0 17px; }
  .ev { grid-template-columns:1fr; gap:6px; padding:15px 0; }
  .ev-when { text-align:left; }
  .ev-when .when { font:400 17px/1.3 var(--serif); }
  .ev-when .dur { display:inline; margin-left:8px; }
  .conv-grid { grid-template-columns:1fr; }
  .gantt-axis { display:block; }
  .gantt-axis .lbl { display:none; }
  .gantt-axis .ticks { height:26px; }
  .gantt-axis .tick { font-size:9px; }
  .gantt-row { display:block; padding-top:6px; }
  .gantt-row .lbl { text-align:left; padding:0 0 3px; font-size:11.5px; line-height:1.3; }
  .gantt-row .track { height:22px; border-bottom:0; }
  .bar { top:4px; height:14px; }
  .bar span { display:none; }
  .bar.wide span { display:block; font-size:9.5px; }
  table.spans thead { display:none; }
  table.spans tr { display:grid; grid-template-columns:max-content max-content 1fr;
    grid-template-areas:"name name name" "start end dur" "notes notes notes";
    padding:11px 0; border-bottom:1px solid var(--line-soft); }
  table.spans tr:last-child { border-bottom:0; }
  table.spans td { display:block; border:0; padding:1px 0; text-align:left; }
  table.spans td:nth-child(1) { grid-area:name; font-weight:600; font-size:13.5px;
    padding-bottom:2px; }
  table.spans td:nth-child(2) { grid-area:start; font-size:13px; }
  table.spans td:nth-child(3) { grid-area:end; font-size:13px; padding-left:5px; }
  table.spans td:nth-child(3)::before { content:"– "; color:var(--faint); }
  table.spans td:nth-child(4) { grid-area:dur; justify-self:end; font-size:12px;
    color:var(--muted); }
  table.spans td:nth-child(4)::after { content:" yrs"; font-size:10.5px; color:var(--faint); }
  table.spans td:nth-child(5) { grid-area:notes; font-size:12.5px; padding-top:5px; }
  .hero .wrap { padding-top:34px; padding-bottom:26px; }
  .hero-stats { display:grid; grid-template-columns:1fr 1fr; gap:16px 20px; }
  .hero-stats .now { margin-left:0; text-align:left; width:100%; grid-column:1 / -1; }
  .tab { padding:12px 9px 10px; }
  .toolbar { position:static; }
  .menu { left:0; right:auto; min-width:0; width:min(264px, calc(100vw - 36px)); }
}
@media print {
  .tabbar, .toolbar, .hero::before, .spanbar button, .gantt-key { display:none; }
  body { background:#fff; }
  .hero { color:#000; background:#fff; border-bottom:2px solid #000; }
  .hero h1, .hero-stats .v, .hero-stats .now b { color:#000; }
  .lede, .kicker, .hero-stats .k, .hero-stats .now { color:#333; }
  .panel, .ev { box-shadow:none; }
  .ev { break-inside:avoid; }
}
</style>
</head>
<body>

<header class="hero">
  <div class="wrap">
    <p class="kicker">Royal Commonwealth of Zeythari</p>
    <h1>Astraia <em>— Master Timeline</em></h1>
    <p class="lede">Every dated event from the Pre-Azyrra to the present day, in one
      browsable page: filter by thread, certainty or era, switch the reckoning on the fly,
      and convert any year between A.S., A.K., years ago and lifetimes.</p>
    <div class="hero-stats">
      <div><div class="v" id="statEvents">—</div><div class="k">Dated events</div></div>
      <div><div class="v" id="statThreads">—</div><div class="k">Threads</div></div>
      <div><div class="v" id="statSpan">—</div><div class="k">Years covered</div></div>
      <div><div class="v" id="statQuestions">—</div><div class="k">Open questions</div></div>
      <div class="now" id="heroNow"></div>
    </div>
  </div>
</header>

<nav class="tabbar" aria-label="Sections">
  <div class="wrap" role="tablist">
    <button class="tab" role="tab" data-tab="timeline" aria-selected="true">Timeline</button>
    <button class="tab" role="tab" data-tab="eras" aria-selected="false">Eras &amp; spans</button>
    <button class="tab" role="tab" data-tab="converter" aria-selected="false">Converter</button>
    <button class="tab" role="tab" data-tab="questions" aria-selected="false">Questions</button>
  </div>
</nav>

<main>
  <div class="wrap">

    <section class="view" id="view-timeline" role="tabpanel">
      <div class="toolbar">
        <div class="toolbar-row">
          <div class="search">
            <svg viewBox="0 0 20 20" aria-hidden="true"><circle cx="9" cy="9" r="6"/>
              <path d="M13.5 13.5 18 18"/></svg>
            <input id="q" type="search" placeholder="Search events, peoples, notes…  (press /)"
                   aria-label="Search the timeline">
          </div>
          <label class="field" style="gap:0">
            <select class="mode" id="mode" aria-label="Show years as">
              <option value="as">Years as A.S.</option>
              <option value="ak">Years as A.K.</option>
              <option value="ybp">Years before present</option>
              <option value="life">Lifetimes ago</option>
            </select>
          </label>
          <select class="mode" id="cert" aria-label="Filter by certainty">
            <option value="">Any certainty</option>
          </select>
          <button class="ghost" id="reset">Reset filters</button>
          <div class="menu-wrap">
            <button class="ghost" id="exportBtn" aria-haspopup="menu" aria-expanded="false"
                    aria-label="Export the currently filtered events"
                    title="Download the currently filtered events">Export &#9662;</button>
            <div class="menu" id="exportMenu" role="menu" hidden>
              <div class="mhead">Download <span id="exportCount"></span></div>
              <button type="button" role="menuitem" data-fmt="csv">CSV
                <small>Spreadsheet — one row per event</small></button>
              <button type="button" role="menuitem" data-fmt="json">JSON
                <small>Every field, plus export metadata</small></button>
            </div>
          </div>
        </div>
        <div class="chips" id="chips" aria-label="Filter by thread"></div>
      </div>
      <p class="hits" id="hits" aria-live="polite"></p>
      <div id="spanBanner"></div>
      <div id="evList"></div>
    </section>

    <section class="view" id="view-eras" role="tabpanel" hidden>
      <h2 class="viewtitle">Eras &amp; spans</h2>
      <p class="viewlede">Proportional to the A.S. axis. Spans overlap by design — they are not a
        single sequence — so each one gets its own line. Click a span to pull up the events inside it.</p>
      <div class="panel"><div class="gantt" id="gantt"></div></div>

      <h2 class="viewtitle" style="margin-top:34px">Span table</h2>
      <p class="viewlede">Duration and the two reckoning equivalents for every era, as in the
        workbook's “Eras &amp; Spans” tab.</p>
      <div class="panel" style="padding:8px 14px">
        <table class="spans">
          <thead><tr><th>Era / span</th><th class="num">Start</th><th class="num">End</th>
            <th class="num">Duration</th><th>Source / notes</th></tr></thead>
          <tbody id="spanRows"></tbody>
        </table>
      </div>
    </section>

    <section class="view" id="view-converter" role="tabpanel" hidden>
      <h2 class="viewtitle">Date converter</h2>
      <p class="viewlede">Type a year and pick its reckoning; the rest follows the same arithmetic as
        the workbook's Converter and Settings tabs.</p>
      <div class="conv-grid">
        <div class="panel">
          <div class="conv-inputs">
            <label class="field"><span>Year</span>
              <input id="cvYear" type="number" value="175" step="1"></label>
            <label class="field"><span>Reckoning</span>
              <select id="cvRec">
                <option value="A.S.">A.S. — After Sundering</option>
                <option value="A.K." selected>A.K. — Anno Kaironis</option>
                <option value="Years ago">Years ago (from the present)</option>
              </select></label>
          </div>
          <div class="conv-out" id="cvOut"></div>
          <p class="hint" id="cvHint"></p>
        </div>
        <div class="panel calib">
          <h2 class="viewtitle" style="font-size:19px;margin-top:0">Calibration</h2>
          <p class="hint" style="margin-top:-2px;margin-bottom:16px">These three numbers drive every
            “years before present” and “lifetimes ago” figure on the page. Event dates themselves stay
            fixed in A.S.</p>
          <label class="field"><span>Present year (A.K.)</span>
            <input id="sPresent" type="number" step="1" min="1"></label>
          <label class="field"><span>A.S. year of 1 A.K.</span>
            <input id="sVisitation" type="number" step="1"></label>
          <label class="field"><span>Lifespan for “lifetimes”</span>
            <input id="sLifespan" type="number" step="1" min="1">
            <em>Illustrai ≈ 200 years.</em></label>
          <button class="ghost" id="sReset">Restore workbook values</button>
        </div>
      </div>
    </section>

    <section class="view" id="view-questions" role="tabpanel" hidden>
      <h2 class="viewtitle">Open questions</h2>
      <p class="viewlede">Undated items, unresolved spans and assumptions still standing between the
        sources — from the workbook's “Open Questions” tab.</p>
      <div class="panel" id="qList"></div>
    </section>

  </div>
</main>

<footer>
  <div class="wrap">
    <details class="legend">
      <summary>How to read this page</summary>
      <dl>
        <dt>Fixed</dt><dd>Stated precisely in the source files.</dd>
        <dt>Approx</dt><dd>The sources give it as approximate (“~”).</dd>
        <dt>Estimated</dt><dd>A “years ago” figure or a range, so the exact year is uncertain.</dd>
        <dt>Derived</dt><dd>Computed from another reckoning — for example the retired F.Y. count.</dd>
        <dt>A.S.</dt><dd>After Sundering: year 0 is the Sundering; negative years are written P.A. (pre-Azyrra).</dd>
        <dt>A.K.</dt><dd>Anno Kaironis: 1 A.K. = 8,725 A.S. Negative A.K. years are written B.K. (before 1 A.K.).</dd>
        <dt>ybp</dt><dd>Years before present, measured back from the present year on the calibration panel.</dd>
      </dl>
    </details>
    <div id="genLine"></div>
  </div>
</footer>

<script type="application/json" id="data">__DATA__</script>
<script>
(function () {
  "use strict";

  var DATA = JSON.parse(document.getElementById("data").textContent);
  var LSKEY = "astraia.timeline.settings.v1";

  var state = {
    mode: "as",
    q: "",
    threads: new Set(),
    cert: "",
    span: null,
    settings: Object.assign({}, DATA.settings)
  };
  try {
    var saved = JSON.parse(localStorage.getItem(LSKEY) || "null");
    if (saved && typeof saved === "object") Object.assign(state.settings, saved);
  } catch (e) {}

  var THREAD_COLORS = {
    Cosmic: "#5b4bb7", Divine: "#b0791a", Maenari: "#0e7490", Maestrai: "#2f855a",
    Illustrai: "#b23b4e", Avesari: "#2b5fa8", Luminari: "#a8458f", Arts: "#8a4b24",
    Commonwealth: "#3d4a5c", Church: "#8a6d3b", Meta: "#6b7280"
  };
  var ERA_COLORS = ["#6b5b95", "#8d7b4f", "#9c6b4a", "#5f7d5a", "#7a6b8f", "#a67c3f",
                    "#4f6f8f", "#8a5a5a", "#6f7d68", "#9a7a2f"];

  var nf = new Intl.NumberFormat("en-US");
  function n(v) { return nf.format(v); }
  function cfg() {
    var s = state.settings;
    return { offset: s.visitation - 1, presentAS: s.presentAK + s.visitation - 1,
             lifespan: s.lifespan > 0 ? s.lifespan : 200 };
  }

  function fmtYear(as, mode) {
    if (as === null || as === undefined || as === "" || isNaN(as)) return "—";
    mode = mode || state.mode;
    var c = cfg();
    if (mode === "ak") {
      var ak = as - c.offset;
      return ak >= 0 ? n(ak) + " A.K." : n(-ak) + " B.K.";
    }
    if (mode === "ybp") {
      var d = c.presentAS - as;
      return d >= 0 ? n(d) + " ybp" : "in " + n(-d) + " yrs";
    }
    if (mode === "life") {
      var l = (c.presentAS - as) / c.lifespan;
      return (Math.round(l * 10) / 10).toLocaleString("en-US",
        { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + " lifetimes";
    }
    return as < 0 ? n(-as) + " P.A." : n(as) + " A.S.";
  }

  function fmtWhen(ev, mode) {
    if (ev.endAS === null || ev.endAS === undefined || ev.endAS === ev.startAS) {
      return fmtYear(ev.startAS, mode);
    }
    return fmtYear(ev.startAS, mode) + " – " + fmtYear(ev.endAS, mode);
  }

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined && text !== null) e.textContent = text;
    return e;
  }

  function highlight(node, text, q) {
    text = text === null || text === undefined ? "" : String(text);
    if (!q) { node.textContent = text; return; }
    var lower = text.toLowerCase(), ql = q.toLowerCase(), from = 0, i;
    while ((i = lower.indexOf(ql, from)) !== -1) {
      if (i > from) node.appendChild(document.createTextNode(text.slice(from, i)));
      var m = document.createElement("mark");
      m.textContent = text.slice(i, i + ql.length);
      node.appendChild(m);
      from = i + ql.length;
    }
    if (from < text.length) node.appendChild(document.createTextNode(text.slice(from)));
    if (!node.childNodes.length) node.textContent = text;
  }

  function threadColor(t) { return THREAD_COLORS[t] || "#6b6257"; }

  /* ---------------- timeline ---------------- */

  var threads = (function () {
    var seen = new Map();
    DATA.events.forEach(function (ev) {
      var key = ev.thread || "Unassigned";
      if (!seen.has(key)) seen.set(key, { name: key, count: 0, first: ev.startAS });
      var t = seen.get(key);
      t.count++;
      t.first = Math.min(t.first, ev.startAS);
    });
    return Array.from(seen.values()).sort(function (a, b) { return a.first - b.first; });
  })();

  var eraMeta = (function () {
    var map = new Map();
    DATA.events.forEach(function (ev) {
      var k = ev.era || "Unassigned";
      if (!map.has(k)) map.set(k, { name: k, min: ev.startAS, max: ev.endAS === null ? ev.startAS : ev.endAS, list: [] });
      var m = map.get(k);
      m.min = Math.min(m.min, ev.startAS);
      m.max = Math.max(m.max, ev.endAS === null ? ev.startAS : ev.endAS);
      m.list.push(ev);
    });
    return map;
  })();

  function matches(ev) {
    if (state.threads.size && !state.threads.has(ev.thread || "Unassigned")) return false;
    if (state.cert && ev.certainty !== state.cert) return false;
    if (state.span && !(ev.startAS >= state.span.start && ev.startAS <= state.span.end)) return false;
    if (state.q) {
      var hay = [ev.event, ev.peoples, ev.notes, ev.source, ev.era, ev.thread,
                 ev.reckoning, ev.written, ev.writtenEnd, ev.certainty, "#" + ev.id]
        .filter(function (x) { return x !== null && x !== undefined; }).join(" ").toLowerCase();
      if (hay.indexOf(state.q.toLowerCase()) === -1) return false;
    }
    return true;
  }

  function eventCard(ev) {
    var art = el("article", "ev cert-" + String(ev.certainty || "").toLowerCase());
    if (/^PRESENT DAY/.test(ev.event)) art.classList.add("is-now");

    var when = el("div", "ev-when");
    var w = el("span", "when");
    w.textContent = fmtWhen(ev);
    when.appendChild(w);
    if (ev.duration) when.appendChild(el("span", "dur", n(ev.duration) + " yrs"));
    art.appendChild(when);

    var main = el("div", "ev-main");
    var tags = el("div", "ev-tags");

    var th = el("button", "thread");
    th.type = "button";
    th.style.setProperty("--c", threadColor(ev.thread));
    th.title = "Filter by thread: " + (ev.thread || "Unassigned");
    th.appendChild(el("i"));
    th.appendChild(document.createTextNode(ev.thread || "Unassigned"));
    th.addEventListener("click", function () {
      if (!state.threads.has(ev.thread)) state.threads.add(ev.thread);
      else state.threads.delete(ev.thread);
      syncChips();
      renderTimeline();
    });
    tags.appendChild(th);

    tags.appendChild(el("span", "cert", ev.certainty || "—"));
    tags.appendChild(el("span", "eid", "#" + ev.id));
    main.appendChild(tags);

    var h = el("h3", "ev-text");
    highlight(h, ev.event, state.q);
    main.appendChild(h);

    if (ev.peoples) {
      var p = el("p", "ev-ppl");
      p.textContent = ev.peoples.split(";").map(function (s) { return s.trim(); }).join(" · ");
      main.appendChild(p);
    }
    if (ev.notes) {
      var note = el("p", "ev-note");
      highlight(note, ev.notes, state.q);
      main.appendChild(note);
    }
    var meta = [];
    if (ev.reckoning && ev.written !== null && ev.written !== undefined) {
      meta.push("As written: " + n(ev.written) +
        (ev.writtenEnd !== null && ev.writtenEnd !== undefined ? "–" + n(ev.writtenEnd) : "") +
        " " + ev.reckoning);
    }
    if (ev.source) meta.push("Source: " + ev.source);
    if (meta.length) main.appendChild(el("p", "ev-src", meta.join(" · ")));

    art.appendChild(main);
    return art;
  }

  function baseName(s) { return String(s).replace(/\s*\([^)]*\)\s*$/, "").trim(); }
  var eraRows = new Map(DATA.eras.map(function (e) { return [baseName(e.name), e]; }));

  function eraHeader(name, filtered, total) {
    var meta = eraMeta.get(name) || { min: null, max: null };
    var row = eraRows.get(baseName(name));
    var head = el("div", "era-head");
    head.appendChild(el("h2", null, name));
    if (row) {
      head.appendChild(el("span", "range", fmtYear(row.start) + " – " + fmtYear(row.end)));
    } else if (meta.min !== null) {
      head.appendChild(el("span", "range", fmtYear(meta.min) + " – " + fmtYear(meta.max)));
    }
    head.appendChild(el("span", "count", filtered === total
      ? total + (total === 1 ? " event" : " events")
      : filtered + " of " + total + (total === 1 ? " event" : " events")));
    return head;
  }

  function renderTimeline() {
    var list = document.getElementById("evList");
    var hits = document.getElementById("hits");
    var banner = document.getElementById("spanBanner");
    var shown = DATA.events.filter(matches);
    list.innerHTML = "";
    banner.innerHTML = "";

    if (state.span) {
      var bar = el("div", "spanbar");
      var txt = el("div");
      txt.appendChild(el("b", null, state.span.name));
      txt.appendChild(document.createTextNode(" · " + fmtYear(state.span.start) + " – " +
        fmtYear(state.span.end) + " — " + shown.length +
        (shown.length === 1 ? " event" : " events") + " inside this span."));
      bar.appendChild(txt);
      var clear = el("button", "ghost", "Clear span");
      clear.type = "button";
      clear.addEventListener("click", function () { state.span = null; renderTimeline(); });
      bar.appendChild(clear);
      banner.appendChild(bar);
    }

    hits.innerHTML = "";
    if (shown.length === DATA.events.length && !state.q && !state.threads.size &&
        !state.cert && !state.span) {
      hits.appendChild(document.createTextNode("Showing all " + DATA.events.length + " events."));
    } else {
      hits.innerHTML = "<b>" + shown.length + "</b> of " + DATA.events.length +
        " events match your filters.";
    }

    document.getElementById("exportCount").textContent = shown.length +
      (shown.length === 1 ? " event" : " events") + " of " + DATA.events.length;
    document.getElementById("exportBtn").disabled = shown.length === 0;

    if (!shown.length) {
      var empty = el("div", "empty");
      empty.appendChild(el("b", null, "Nothing matches those filters."));
      empty.appendChild(document.createTextNode("Try a different word, or reset the filters."));
      list.appendChild(empty);
      return;
    }

    if (state.span) {
      shown.slice().sort(function (a, b) { return a.startAS - b.startAS; })
        .forEach(function (ev) { list.appendChild(eventCard(ev)); });
      return;
    }

    var groups = new Map();
    DATA.events.forEach(function (ev) {
      var k = ev.era || "Unassigned";
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(ev);
    });
    var order = Array.from(groups.keys()).sort(function (a, b) {
      var am = eraMeta.get(a), bm = eraMeta.get(b);
      return am.min - bm.min;
    });

    var any = false;
    order.forEach(function (name) {
      var all = groups.get(name);
      var vis = all.filter(matches).sort(function (a, b) { return a.startAS - b.startAS; });
      if (!vis.length) return;
      any = true;
      list.appendChild(eraHeader(name, vis.length, all.length));
      vis.forEach(function (ev) { list.appendChild(eventCard(ev)); });
    });
    if (!any) list.appendChild(el("div", "empty", "Nothing matches those filters."));
  }

  /* ---------------- chips / controls ---------------- */

  function syncChips() {
    document.querySelectorAll("#chips .chip").forEach(function (c) {
      c.setAttribute("aria-pressed", String(state.threads.has(c.dataset.thread)));
    });
  }

  function buildChips() {
    var box = document.getElementById("chips");
    box.innerHTML = "";
    threads.forEach(function (t) {
      var chip = el("button", "chip");
      chip.type = "button";
      chip.dataset.thread = t.name;
      chip.setAttribute("aria-pressed", "false");
      var dot = el("i");
      dot.style.background = threadColor(t.name);
      chip.appendChild(dot);
      chip.appendChild(document.createTextNode(t.name));
      chip.appendChild(el("span", null, String(t.count)));
      chip.addEventListener("click", function () {
        if (state.threads.has(t.name)) state.threads.delete(t.name);
        else state.threads.add(t.name);
        syncChips();
        renderTimeline();
      });
      box.appendChild(chip);
    });
  }

  function buildCert() {
    var sel = document.getElementById("cert");
    var seen = [];
    DATA.events.forEach(function (ev) {
      if (ev.certainty && seen.indexOf(ev.certainty) === -1) seen.push(ev.certainty);
    });
    seen.forEach(function (c) {
      var o = document.createElement("option");
      o.value = c;
      o.textContent = c;
      sel.appendChild(o);
    });
    sel.addEventListener("change", function () {
      state.cert = sel.value;
      renderTimeline();
    });
  }

  function resetFilters() {
    state.q = "";
    state.threads = new Set();
    state.cert = "";
    state.span = null;
    document.getElementById("q").value = "";
    document.getElementById("cert").value = "";
    syncChips();
    renderTimeline();
  }

  /* ---------------- eras ---------------- */

  var AXIS = (function () {
    var lo = -2000, hi = 9999;
    DATA.events.forEach(function (ev) {
      lo = Math.min(lo, ev.startAS);
      hi = Math.max(hi, ev.endAS === null ? ev.startAS : ev.endAS);
    });
    DATA.eras.forEach(function (e) {
      lo = Math.min(lo, e.start); hi = Math.max(hi, e.end);
    });
    return { min: lo, max: hi };
  })();

  function pct(v) { return (v - AXIS.min) / (AXIS.max - AXIS.min) * 100; }

  function renderGantt() {
    var box = document.getElementById("gantt");
    box.innerHTML = "";
    var axis = el("div", "gantt-axis");
    axis.appendChild(el("div", "lbl"));
    var ticks = el("div", "ticks");
    var values = [];
    for (var v = Math.ceil(AXIS.min / 1000) * 1000; v <= AXIS.max; v += 1000) values.push(v);
    values.forEach(function (v, i) {
      var p = pct(v);
      if (p > 99.2) return;
      var t = el("div", "tick", fmtYear(v, "as"));
      if (v % 2000 !== 0) t.classList.add("m");
      if (v % 3000 === 0) t.classList.add("m3");
      if (i === 0) t.classList.add("edge");
      t.style.left = p + "%";
      if (p < 4) t.style.transform = "none";
      ticks.appendChild(t);
    });
    axis.appendChild(ticks);
    box.appendChild(axis);

    DATA.eras.forEach(function (e, i) {
      var row = el("div", "gantt-row");
      row.appendChild(el("div", "lbl", e.name));
      var track = el("div", "track");

      var left = Math.max(0, pct(e.start));
      var right = Math.min(100, pct(e.end));
      var width = Math.max(right - left, 0.7);

      var bar = el("button", "bar");
      bar.type = "button";
      bar.style.left = left + "%";
      bar.style.width = width + "%";
      bar.style.setProperty("--c", ERA_COLORS[i % ERA_COLORS.length]);
      bar.title = e.name + ": " + fmtYear(e.start, "as") + " – " + fmtYear(e.end, "as") +
        (e.duration ? " (" + n(e.duration) + " years)" : "") + " — click to show its events";
      if (width > 13) {
        bar.appendChild(el("span", null, fmtYear(e.start, state.mode)));
        if (width > 34) bar.classList.add("wide");
      }
      bar.addEventListener("click", function () {
        state.span = { name: e.name, start: e.start, end: e.end };
        renderTimeline();
        showView("timeline");
      });
      track.appendChild(bar);

      var c = cfg();
      [c.offset + 1, c.presentAS].forEach(function (mark, idx) {
        if (mark < AXIS.min || mark > AXIS.max) return;
        var m = el("div", "marker" + (idx === 0 ? " ak" : " now"));
        m.style.left = pct(mark) + "%";
        m.title = idx === 0 ? "1 A.K. (the Visitation)" : "Present day";
        track.appendChild(m);
      });

      row.appendChild(track);
      box.appendChild(row);
    });

    var key = el("div", "gantt-key");
    var k1 = el("div");
    k1.appendChild(el("span", "sw"));
    k1.appendChild(document.createTextNode("the Visitation (1 A.K.)"));
    key.appendChild(k1);
    var k1b = el("div");
    k1b.appendChild(el("span", "sw now"));
    k1b.appendChild(document.createTextNode("present day"));
    key.appendChild(k1b);
    var k2 = el("div", null, "Click any span to filter the timeline to the events inside it.");
    key.appendChild(k2);
    box.appendChild(key);
  }

  function renderSpans() {
    var body = document.getElementById("spanRows");
    body.innerHTML = "";
    DATA.eras.forEach(function (e) {
      var tr = el("tr");
      tr.appendChild(el("td", null, e.name));
      tr.appendChild(el("td", "num", fmtYear(e.start)));
      tr.appendChild(el("td", "num", fmtYear(e.end)));
      tr.appendChild(el("td", "num", e.duration ? n(e.duration) : "—"));
      tr.appendChild(el("td", "dim", e.notes || ""));
      body.appendChild(tr);
    });
  }

  /* ---------------- converter ---------------- */

  function toAS(year, rec) {
    var c = cfg();
    if (rec === "A.K.") return year + c.offset;
    if (rec === "Years ago") return c.presentAS - year;
    return year;
  }

  function renderConverter() {
    var input = document.getElementById("cvYear");
    var rec = document.getElementById("cvRec").value;
    var year = Number(input.value);
    var out = document.getElementById("cvOut");
    out.innerHTML = "";
    if (input.value === "" || isNaN(year)) {
      out.appendChild(el("p", "hint", "Type a year to convert it."));
      document.getElementById("cvHint").textContent = "";
      return;
    }
    var as = toAS(year, rec);
    var c = cfg();
    var rows = [
      ["A.S.", fmtYear(as, "as")],
      ["A.K. equivalent", fmtYear(as, "ak")],
      ["Years before present", (c.presentAS - as >= 0 ? n(c.presentAS - as) + " years"
        : "in " + n(as - c.presentAS) + " years")],
      ["Lifetimes ago", ((c.presentAS - as) / c.lifespan).toLocaleString("en-US",
        { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + " × " + n(c.lifespan) + "-year lifespans"],
      ["Eras at that date", DATA.eras.filter(function (e) { return as >= e.start && as <= e.end; })
        .map(function (e) { return e.name; }).join(" · ") || "—"]
    ];
    rows.forEach(function (r) {
      var row = el("div", "conv-row");
      row.appendChild(el("span", "k", r[0]));
      row.appendChild(el("span", "v", r[1]));
      out.appendChild(row);
    });
    document.getElementById("cvHint").textContent =
      "Workbook arithmetic: A.S. = A.K. + " + n(c.offset) + "; present = " +
      n(state.settings.presentAK) + " A.K. = " + n(c.presentAS) + " A.S.";
  }

  function renderCalibration() {
    document.getElementById("sPresent").value = state.settings.presentAK;
    document.getElementById("sVisitation").value = state.settings.visitation;
    document.getElementById("sLifespan").value = state.settings.lifespan;
  }

  function readCalibration() {
    var p = Number(document.getElementById("sPresent").value);
    var v = Number(document.getElementById("sVisitation").value);
    var l = Number(document.getElementById("sLifespan").value);
    if (!isNaN(p) && p > 0) state.settings.presentAK = p;
    if (!isNaN(v)) state.settings.visitation = v;
    if (!isNaN(l) && l > 0) state.settings.lifespan = l;
    try { localStorage.setItem(LSKEY, JSON.stringify(state.settings)); } catch (e) {}
    updateChrome();
    renderTimeline();
    renderSpans();
    renderGantt();
    renderConverter();
  }

  /* ---------------- questions ---------------- */

  function renderQuestions() {
    var box = document.getElementById("qList");
    box.innerHTML = "";
    DATA.questions.forEach(function (q) {
      var row = el("div", "q");
      row.appendChild(el("div", "n", "#" + q.id));
      var main = el("div");
      main.appendChild(el("h3", null, q.item));
      if (q.why) main.appendChild(el("p", null, q.why));
      if (q.where) main.appendChild(el("div", "where", "Appears in: " + q.where));
      row.appendChild(main);
      box.appendChild(row);
    });
  }

  /* ---------------- export ---------------- */

  var eraOrder = Array.from(eraMeta.keys()).sort(function (a, b) {
    return eraMeta.get(a).min - eraMeta.get(b).min;
  });
  var eraRank = new Map(eraOrder.map(function (k, i) { return [k, i]; }));

  function filteredSorted() {
    var list = DATA.events.filter(matches);
    list.sort(function (a, b) {
      if (!state.span) {
        var ra = eraRank.has(a.era) ? eraRank.get(a.era) : 99;
        var rb = eraRank.has(b.era) ? eraRank.get(b.era) : 99;
        if (ra !== rb) return ra - rb;
      }
      return a.startAS - b.startAS || a.id - b.id;
    });
    return list;
  }

  function exportRows(list) {
    var c = cfg();
    return list.map(function (ev) {
      var endAS = ev.endAS === null || ev.endAS === undefined ? null : ev.endAS;
      return {
        id: ev.id,
        era: ev.era,
        thread: ev.thread,
        event: ev.event,
        startAS: ev.startAS,
        endAS: endAS,
        startAK: ev.startAS - c.offset,
        endAK: endAS === null ? null : endAS - c.offset,
        yearsBeforePresent: c.presentAS - ev.startAS,
        lifetimesAgo: Math.round((c.presentAS - ev.startAS) / c.lifespan * 10) / 10,
        durationYears: ev.duration,
        reckoning: ev.reckoning,
        yearAsWritten: ev.written,
        endYearAsWritten: ev.writtenEnd,
        peoples: ev.peoples,
        certainty: ev.certainty,
        source: ev.source,
        notes: ev.notes
      };
    });
  }

  function csvCell(v) {
    if (v === null || v === undefined) return "";
    var s = String(v);
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  function toCSV(rows) {
    var c = cfg();
    var head = ["ID", "Era", "Thread", "Event", "Start (A.S.)", "End (A.S.)",
      "Start (A.K.)", "End (A.K.)", "Years before present",
      "Lifetimes ago (" + n(c.lifespan) + " yr)", "Duration (years)",
      "Reckoning (as written)", "Year (as written)", "End year (as written)",
      "Peoples / places", "Certainty", "Source", "Notes"];
    var keys = ["id", "era", "thread", "event", "startAS", "endAS", "startAK", "endAK",
      "yearsBeforePresent", "lifetimesAgo", "durationYears", "reckoning",
      "yearAsWritten", "endYearAsWritten", "peoples", "certainty", "source", "notes"];
    var lines = [head.map(csvCell).join(",")];
    rows.forEach(function (r) {
      lines.push(keys.map(function (k) { return csvCell(r[k]); }).join(","));
    });
    return lines.join("\r\n") + "\r\n";
  }

  function download(filename, mime, text) {
    var blob = new Blob([text], { type: mime });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(url);
      if (a.parentNode) a.parentNode.removeChild(a);
    }, 0);
  }

  function todayStamp() {
    var d = new Date();
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") +
      "-" + String(d.getDate()).padStart(2, "0");
  }

  function exportEvents(fmt) {
    var rows = exportRows(filteredSorted());
    var stamp = todayStamp();
    if (fmt === "json") {
      var c = cfg();
      var payload = {
        meta: {
          source: DATA.source,
          exported: new Date().toISOString(),
          count: rows.length,
          of: DATA.events.length,
          settings: {
            presentAK: state.settings.presentAK,
            visitationAS: state.settings.visitation,
            presentAS: c.presentAS,
            akOffset: c.offset,
            lifespan: c.lifespan
          },
          filters: {
            search: state.q || null,
            threads: Array.from(state.threads),
            certainty: state.cert || null,
            span: state.span
              ? { name: state.span.name, startAS: state.span.start, endAS: state.span.end }
              : null
          }
        },
        events: rows
      };
      download("astraia-timeline-" + stamp + ".json", "application/json",
        JSON.stringify(payload, null, 2) + "\n");
    } else {
      download("astraia-timeline-" + stamp + ".csv", "text/csv;charset=utf-8",
        "\ufeff" + toCSV(rows));
    }
    openExportMenu(false);
  }

  /* ---------------- chrome / routing ---------------- */

  function updateChrome() {
    var c = cfg();
    document.getElementById("statEvents").textContent = DATA.events.length;
    document.getElementById("statThreads").textContent = threads.length;
    document.getElementById("statSpan").textContent = n(AXIS.max - AXIS.min);
    document.getElementById("statQuestions").textContent = DATA.questions.length;
    document.getElementById("heroNow").innerHTML = "";
    document.getElementById("heroNow").appendChild(document.createTextNode("Present day "));
    var b = el("b", null, n(state.settings.presentAK) + " A.K. = " + n(c.presentAS) + " A.S.");
    document.getElementById("heroNow").appendChild(b);
    document.getElementById("heroNow").appendChild(
      document.createTextNode(" · " + n(c.lifespan) + "-year lifespan"));
  }

  function showView(name) {
    if (!document.getElementById("view-" + name)) name = "timeline";
    openExportMenu(false);
    ["timeline", "eras", "converter", "questions"].forEach(function (v) {
      document.getElementById("view-" + v).hidden = v !== name;
      var tab = document.querySelector('.tab[data-tab="' + v + '"]');
      tab.setAttribute("aria-selected", String(v === name));
    });
    if (location.hash.slice(1) !== name) {
      try { history.replaceState(null, "", "#" + name); } catch (e) { location.hash = name; }
    }
    window.scrollTo({ top: 0, behavior: "auto" });
  }

  /* ---------------- wiring ---------------- */

  var q = document.getElementById("q");
  q.addEventListener("input", function () {
    state.q = q.value.trim();
    renderTimeline();
  });

  document.getElementById("mode").addEventListener("change", function (e) {
    state.mode = e.target.value;
    renderTimeline();
    renderSpans();
    renderGantt();
  });

  document.getElementById("reset").addEventListener("click", resetFilters);

  var exportBtn = document.getElementById("exportBtn");
  var exportMenu = document.getElementById("exportMenu");

  function openExportMenu(open) {
    exportMenu.hidden = !open;
    exportBtn.setAttribute("aria-expanded", String(open));
    exportMenu.style.transform = "";
    if (!open) return;
    var pad = 10;
    var r = exportMenu.getBoundingClientRect();
    var shift = 0;
    if (r.right > window.innerWidth - pad) shift = window.innerWidth - pad - r.right;
    else if (r.left < pad) shift = pad - r.left;
    if (shift) exportMenu.style.transform = "translateX(" + Math.round(shift) + "px)";
  }

  exportBtn.addEventListener("click", function (e) {
    e.stopPropagation();
    openExportMenu(exportMenu.hidden);
  });

  document.addEventListener("click", function (e) {
    if (!exportMenu.hidden && !(e.target.closest && e.target.closest(".menu-wrap"))) {
      openExportMenu(false);
    }
  });

  document.querySelectorAll("#exportMenu [data-fmt]").forEach(function (b) {
    b.addEventListener("click", function () { exportEvents(b.dataset.fmt); });
  });

  document.querySelectorAll(".tab").forEach(function (tab) {
    tab.addEventListener("click", function () { showView(tab.dataset.tab); });
  });

  document.getElementById("cvYear").addEventListener("input", renderConverter);
  document.getElementById("cvRec").addEventListener("change", renderConverter);
  ["sPresent", "sVisitation", "sLifespan"].forEach(function (id) {
    document.getElementById(id).addEventListener("change", readCalibration);
  });
  document.getElementById("sReset").addEventListener("click", function () {
    state.settings = Object.assign({}, DATA.settings);
    try { localStorage.setItem(LSKEY, JSON.stringify(state.settings)); } catch (e) {}
    renderCalibration();
    readCalibration();
  });

  document.addEventListener("keydown", function (e) {
    if (e.key === "/" && document.activeElement !== q) {
      var view = document.querySelector('.view:not([hidden])');
      if (view && view.id === "view-timeline") { e.preventDefault(); q.focus(); q.select(); }
    }
    if (e.key === "Escape" && !exportMenu.hidden) openExportMenu(false);
    if (e.key === "Escape" && document.activeElement === q) {
      q.value = ""; state.q = ""; renderTimeline(); q.blur();
    }
  });

  window.addEventListener("hashchange", function () {
    var h = location.hash.slice(1);
    if (h) showView(h);
  });

  /* ---------------- init ---------------- */

  buildChips();
  buildCert();
  renderQuestions();
  renderCalibration();
  updateChrome();
  renderGantt();
  renderSpans();
  renderConverter();
  renderTimeline();
  document.getElementById("genLine").textContent =
    "Generated from " + DATA.source + " · no external scripts, fonts or network calls · " +
    DATA.events.length + " events, " + DATA.eras.length + " spans, " +
    DATA.questions.length + " open questions.";

  var startHash = location.hash.slice(1);
  if (startHash) showView(startHash);
})();
</script>
</body>
</html>
"""


def s(v):
    """Trim strings, keep numbers as-is."""
    if isinstance(v, str):
        return v.strip()
    return v


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    ws = wb["Timeline"]
    rows = list(ws.iter_rows(values_only=True))
    head = [str(h).strip() if h is not None else "" for h in rows[0]]
    col = {name: i for i, name in enumerate(head)}
    events, legend = [], {}
    for r in rows[1:]:
        rid = r[col["ID"]]
        if not isinstance(rid, int):
            for cell in r:
                if isinstance(cell, str) and cell.startswith(("Certainty:", "Reckoning:")):
                    legend[cell.split(":")[0].lower()] = cell.split(":", 1)[1].strip()
            continue
        events.append({
            "id": rid,
            "reckoning": s(r[col["Reckoning"]]),
            "written": s(r[col["Year (as written)"]]),
            "writtenEnd": s(r[col["End year (as written)"]]),
            "startAS": s(r[col["Start (A.S.)"]]),
            "endAS": s(r[col["End (A.S.)"]]),
            "duration": s(r[col["Duration (years)"]]),
            "era": s(r[col["Era"]]),
            "thread": s(r[col["Thread"]]),
            "event": s(r[col["Event"]]),
            "peoples": s(r[col["Peoples / places"]]),
            "certainty": s(r[col["Certainty"]]),
            "source": s(r[col["Source"]]),
            "notes": s(r[col["Notes"]]),
        })

    eras = []
    for r in list(wb["Eras & Spans"].iter_rows(values_only=True))[1:]:
        if not r[0] or not isinstance(r[1], (int, float)):
            continue
        eras.append({
            "name": s(r[0]), "start": int(r[1]), "end": int(r[2]),
            "duration": s(r[3]), "startAK": s(r[4]), "endYBP": s(r[5]),
            "notes": s(r[6]),
        })

    questions = []
    for r in list(wb["Open Questions"].iter_rows(values_only=True))[1:]:
        if not isinstance(r[0], int):
            continue
        questions.append({"id": r[0], "item": s(r[1]), "why": s(r[2]), "where": s(r[3])})

    st = wb["Settings"]
    settings = {
        "presentAK": int(st["B2"].value),
        "visitation": int(st["B3"].value),
        "lifespan": int(st["B6"].value),
    }

    return {
        "source": path.split("/")[-1],
        "settings": settings,
        "legend": legend,
        "events": events,
        "eras": eras,
        "questions": questions,
    }


def build(xlsx, out):
    data = load(xlsx)
    blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    page = TEMPLATE.replace("__DATA__", blob)
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(page)
    print(f"{xlsx} -> {out}: {len(page):,} bytes, {len(data['events'])} events, "
          f"{len(data['eras'])} spans, {len(data['questions'])} questions")


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "Astraia_Master_Timeline.xlsx",
          sys.argv[2] if len(sys.argv) > 2 else "index.html")
