#!/usr/bin/env python3
"""Render changes.json into a self-contained HTML change log."""
import html
import json
import re
import sys

E = html.escape

OP_LABEL = {
    "added": "Added", "removed": "Removed", "modified": "Changed",
    "format": "Formatting", "moved": "Moved",
}


def esc(t):
    return E(t).replace("\t", "    ")


def render_ops(ops, cls=""):
    out = []
    for op, txt in ops:
        s = esc(txt)
        if op == "=":
            out.append(s)
        elif op == "+":
            out.append(f'<ins class="ins">{s}</ins>')
        elif op == "-":
            out.append(f'<del class="del">{s}</del>')
    return f'<span class="diffline {cls}">' + "".join(out) + "</span>"


def short(t, n=110):
    t = " ".join(t.split())
    return t if len(t) <= n else t[: n - 1].rstrip() + "…"


def ref(e):
    a, b = e.get("old_ref"), e.get("new_ref")
    if a and b:
        return f"{a} → {b}"
    return b or a or ""


def table_rows_html(rows, mode):
    cells = []
    for r in rows[:40]:
        tds = "".join(f"<td>{esc(c['t'])}</td>" for c in r)
        cells.append(f'<tr class="{mode}">{tds}</tr>')
    if len(rows) > 40:
        cells.append(f'<tr class="more"><td colspan="99">… {len(rows) - 40} more rows</td></tr>')
    return "".join(cells)


def table_delta_html(e):
    parts = ['<div class="tblwrap"><table class="tbl tbl-diff">']
    hdr = e.get("header") or []
    if hdr and e["op"] == "modified":
        parts.append('<tr class="row-ctx">' + "".join(f"<td>{esc(c)}</td>" for c in hdr) + "</tr>")
    for ch in e.get("table", []):
        if ch["op"] == "added":
            parts.append(table_rows_html([ch["row"]], "row-add"))
        elif ch["op"] == "removed":
            parts.append(table_rows_html([ch["row"]], "row-del"))
        else:
            tds = []
            for c in ch["cells"]:
                if c["same"]:
                    tds.append(f"<td>{esc(c['text'])}</td>")
                else:
                    tds.append(f'<td class="cell-mod">{render_ops(c["diff"])}</td>')
            parts.append('<tr class="row-mod">' + "".join(tds) + "</tr>")
    parts.append("</table></div>")
    return "".join(parts)


def entry_html(e, idx):
    op = e["op"]
    kind = e["kind"]
    meta = []
    if e.get("style"):
        meta.append(esc(e["style"]))
    meta.append(ref(e))
    meta.append(f"{len(e.get('table') or [])} row edit(s)" if kind == "tbl" and op != "added" else "")
    if kind == "tbl":
        meta.append(f"{e.get('old_rows', 0) if op != 'added' else 0} → {e.get('new_rows', 0)} rows"
                    if op != "added" else f"{e.get('new_rows', 0)} rows")
    body = ""
    if kind == "tbl" and e.get("table"):
        body = table_delta_html(e)
    elif op == "format":
        body = f'<div class="diff">{render_ops(e["ops"])}</div>'
    elif op == "moved":
        body = f'<div class="diff quiet">{render_ops(e["ops"])}</div>'
    else:
        body = f'<div class="diff">{render_ops(e["ops"])}</div>'

    notes = ""
    if e.get("notes"):
        notes = '<ul class="notes">' + "".join(f"<li>{esc(n)}</li>" for n in e["notes"]) + "</ul>"
    tag = {"p": "¶", "tbl": "▦"}[kind]
    meta = [m for m in meta if m]
    return f'''
<article class="entry op-{op} kind-{kind}" data-op="{op}" data-kind="{kind}" id="c{idx}">
  <div class="entry-head">
    <span class="badge b-{op}">{OP_LABEL[op]}</span>
    <span class="tag">{tag}</span>
    <span class="meta">{' · '.join(meta)}</span>
    <a class="anchor" href="#c{idx}">#</a>
  </div>
  {notes}
  {body}
</article>'''


def main(changes_path, old_json, new_json, out_path, old_date="", new_date=""):
    data = json.load(open(changes_path))
    st = data["stats"]
    changes = data["changes"]

    sys.path.insert(0, ".")
    import diff as D

    mo, old = D.load(old_json)
    mn, new = D.load(new_json)
    sn = {}
    sn.update(mo["style_names"])
    sn.update(mn["style_names"])
    D.annotate_context(old, sn)
    D.annotate_context(new, sn)

    # ---- chapter word totals ------------------------------------------------
    def chapter_words(blocks):
        acc = {}
        for b in blocks:
            key = b["_h1"] or "(front matter)"
            n = len(b["_text"].split())
            acc[key] = acc.get(key, 0) + n
        return acc

    ow, nw = chapter_words(old), chapter_words(new)

    # ---- position keys so deletions land in the right place -----------------
    pairs = sorted((e["old_index"], e["new_index"]) for e in changes
                   if e.get("old_index") is not None and e.get("new_index") is not None)
    oi_sorted = [p[0] for p in pairs]

    def pos_key(e):
        if e.get("new_index") is not None:
            return (e["new_index"], 0)
        oi = e["old_index"]
        lo, hi = 0, len(oi_sorted)
        while lo < hi:
            mid = (lo + hi) // 2
            if oi_sorted[mid] <= oi:
                lo = mid + 1
            else:
                hi = mid
        base = pairs[lo - 1][1] if lo > 0 else -1
        return (base + 0.5, oi)

    # ---- group by chapter ---------------------------------------------------
    chapter_order = [b["_text"] for b in new if b["_lvl"] == 1]
    groups = {}
    for e in changes:
        ch = e["h1"] or (e["h1_old"] if e["op"] == "removed" else None) or "(front matter)"
        groups.setdefault(ch, []).append(e)
    for v in groups.values():
        v.sort(key=pos_key)

    extra = sorted(c for c in groups if c not in chapter_order)
    extra.sort(key=lambda c: (c != "(front matter)", c))
    chapter_list = chapter_order + extra                   # all chapters in overview
    section_list = [c for c in chapter_list if groups.get(c)]

    # ---- outline changes ----------------------------------------------------
    added_heads = [e for e in changes if e["op"] == "added" and e["level"]]
    removed_heads = [e for e in changes if e["op"] == "removed" and e["level"]]
    renames = [e for e in changes if e["op"] == "modified" and e["level"]]
    levels = [e for e in changes if e["op"] == "format" and e["level"]]

    def head_row(e):
        lvl = e["level"] or 0
        return (f'<tr><td><span class="hl hl{lvl}">H{lvl}</span></td>'
                f'<td>{render_ops(e["ops"])}</td>'
                f'<td class="dim">{esc(e["h1"] or e["h1_old"] or "")}</td>'
                f'<td class="dim">{ref(e)}</td></tr>')

    outline = ""
    if added_heads or removed_heads or renames or levels:
        outline = '<table class="grid"><thead><tr><th>Level</th><th>Heading</th>' \
                  '<th>Chapter</th><th>Ref</th></tr></thead><tbody>'
        outline += "".join(head_row(e) for e in added_heads)
        outline += "".join(head_row(e) for e in removed_heads)
        outline += "".join(head_row(e) for e in renames)
        outline += "".join(head_row(e) for e in levels)
        outline += "</tbody></table>"

    # ---- chapter overview table --------------------------------------------
    rows = []
    for ch in chapter_list:
        es = groups.get(ch, [])
        n_add = sum(1 for e in es if e["op"] == "added")
        n_del = sum(1 for e in es if e["op"] == "removed")
        n_mod = sum(1 for e in es if e["op"] == "modified")
        n_fmt = sum(1 for e in es if e["op"] in ("format", "moved"))
        wa = sum(len(t.split()) for e in es for o, t in e["ops"] if o == "+")
        wr = sum(len(t.split()) for e in es for o, t in e["ops"] if o == "-")
        o_c, n_c = ow.get(ch, 0), nw.get(ch, 0)
        delta = n_c - o_c
        dcls = "up" if delta > 0 else ("down" if delta < 0 else "")
        dtxt = f"{delta:+,}" if delta else "—"
        anchor = "ch-" + str(chapter_list.index(ch)) if es else ""
        name = f'<a href="#{anchor}">{esc(ch)}</a>' if es else esc(ch)
        rows.append(
            f'<tr><td>{name}</td>'
            f'<td class="num">{len(es)}</td>'
            f'<td class="num">{n_add}</td><td class="num">{n_del}</td>'
            f'<td class="num">{n_mod}</td><td class="num">{n_fmt}</td>'
            f'<td class="num">{wa:,}</td><td class="num">{wr:,}</td>'
            f'<td class="num {dcls}">{dtxt}</td></tr>')

    # ---- chapter sections ---------------------------------------------------
    sections = []
    total = 0
    counter = 0
    for ci, ch in enumerate(chapter_list):
        es = groups.get(ch, [])
        if not es:
            continue
        total += len(es)
        anchor = f"ch-{ci}"
        sub = {}
        for e in es:
            sec = e["h2"] or e["h2_old"] or "— opening —"
            sub.setdefault(sec, []).append(e)
        body = []
        for sec, items in sub.items():
            body.append(f'<h3 class="subsec">{esc(sec)}</h3>')
            for e in items:
                counter += 1
                body.append(entry_html(e, counter))
        sections.append(f'''
<section class="chapter" id="{anchor}">
  <h2 class="chap">{esc(ch)} <span class="count">{len(es)} change{'s' if len(es) != 1 else ''}</span></h2>
  {''.join(body)}
</section>''')

    # ---- media --------------------------------------------------------------
    media_rows = []
    om = {m["name"]: m["bytes"] for m in st["media_old"]}
    nm = {m["name"]: m["bytes"] for m in st["media_new"]}
    for name in sorted(set(om) | set(nm)):
        if name in om and name in nm:
            status = "unchanged" if om[name] == nm[name] else "replaced"
        elif name in nm:
            status = "added"
        else:
            status = "removed"
        a = f"{om[name]:,}" if name in om else "—"
        b = f"{nm[name]:,}" if name in nm else "—"
        cls = " class=\"down\"" if status == "removed" else (
            " class=\"up\"" if status == "added" else "")
        media_rows.append(
            f'<tr><td>{esc(name)}</td><td class="num">{a}</td>'
            f'<td class="num">{b}</td><td{cls}>{status}</td></tr>')
    media = ('<table class="grid"><thead><tr><th>Media part</th><th>Old bytes</th>'
             '<th>New bytes</th><th>Status</th></tr></thead><tbody>'
             + "".join(media_rows) + "</tbody></table>")

    # ---- tracked revisions still pending in the "after" document ----------
    revs = st.get("revisions_new") or []
    revisions = ""
    if revs:
        ins = [r for r in revs if r["kind"] == "inserted"]
        dele = [r for r in revs if r["kind"] == "deleted"]
        items = "".join(
            f'<tr><td class="num">{i}</td><td>{esc(r["at"])}</td>'
            f'<td><span class="badge b-{"added" if r["kind"] == "inserted" else "removed"}">'
            f'{r["kind"]}</span></td><td>{esc(r["text"])}</td></tr>'
            for i, r in enumerate(revs, 1))
        revisions = f'''<h2 class="sec">Tracked changes still pending in the “after” file &mdash; {len(ins)} insertions, {len(dele)} deletions</h2>
<p class="note">The merged file has not had its revisions accepted. Its text was read in the
final (revisions-applied) state for the diff above; these edits are listed separately because
they are still open in the document itself.</p>
<table class="grid revs"><thead><tr><th class="num">#</th><th>At</th><th>Revision</th><th>Text</th></tr></thead>
<tbody>{items}</tbody></table>'''

    nav = "".join(
        f'<a href="#ch-{i}">{esc(c)}<span>{len(groups.get(c, []))}</span></a>'
        for i, c in enumerate(chapter_list) if groups.get(c))

    n_mod_p = sum(1 for e in changes if e["op"] == "modified" and e["kind"] == "p")
    n_mod_t = sum(1 for e in changes if e["op"] == "modified" and e["kind"] == "tbl")

    page = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Change Log — Royal Commonwealth of Zeythari</title>
<style>
:root {{
  --ink:#14161a; --dim:#5f6773; --line:#e2e6ec; --bg:#f6f7f9; --card:#fff;
  --add:#0a7a35; --addbg:#e7f6ec; --del:#b3261e; --delbg:#fdecea;
  --mod:#8a5a00; --modbg:#fff6e0; --fmt:#4b4fa8; --fmtbg:#eef0fd;
}}
* {{ box-sizing:border-box; }}
body {{ margin:0; font:16px/1.6 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
  color:var(--ink); background:var(--bg); }}
header.top {{ background:#11151c; color:#f4f6f9; padding:34px 32px 26px; }}
header.top h1 {{ margin:0 0 6px; font-size:27px; letter-spacing:.2px; }}
header.top .sub {{ color:#a9b3c1; font-size:14px; }}
header.top .files {{ margin-top:16px; display:flex; gap:14px; flex-wrap:wrap; }}
header.top .files div {{ background:#1c2230; border:1px solid #2a3243; border-radius:9px;
  padding:9px 13px; font-size:12.5px; color:#c9d2de; }}
header.top .files b {{ color:#fff; font-weight:600; display:block; font-size:13px; }}
main {{ max-width:1180px; margin:0 auto; padding:26px 32px 80px; }}
.cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(148px,1fr)); gap:12px; margin:22px 0 8px; }}
.card {{ background:var(--card); border:1px solid var(--line); border-radius:11px; padding:13px 15px; }}
.card .k {{ font-size:11.5px; text-transform:uppercase; letter-spacing:.09em; color:var(--dim); }}
.card .v {{ font-size:23px; font-weight:650; margin-top:3px; }}
.card.add .v {{ color:var(--add); }} .card.del .v {{ color:var(--del); }}
.card.mod .v {{ color:var(--mod); }} .card.fmt .v {{ color:var(--fmt); }}
h2.sec {{ font-size:14px; text-transform:uppercase; letter-spacing:.1em; color:var(--dim);
  margin:34px 0 12px; padding-bottom:7px; border-bottom:1px solid var(--line); }}
table.grid {{ width:100%; border-collapse:collapse; background:var(--card);
  border:1px solid var(--line); border-radius:11px; overflow:hidden; font-size:14px; }}
table.grid th {{ text-align:left; font-size:11.5px; text-transform:uppercase; letter-spacing:.07em;
  color:var(--dim); background:#f1f3f6; padding:9px 12px; border-bottom:1px solid var(--line); }}
table.grid td {{ padding:9px 12px; border-bottom:1px solid #eef1f4; vertical-align:top; }}
table.grid tr:last-child td {{ border-bottom:0; }}
td.num, th.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
td.up {{ color:var(--add); }} td.down {{ color:var(--del); }}
td.dim, .dim {{ color:var(--dim); }}
p.note {{ font-size:13px; color:var(--dim); margin:0 0 10px; max-width:78ch; }}
table.revs td {{ font-size:13.5px; }}
nav.toc {{ display:flex; flex-wrap:wrap; gap:7px; margin:14px 0 6px; }}
nav.toc a {{ display:inline-flex; gap:7px; align-items:center; background:var(--card);
  border:1px solid var(--line); border-radius:999px; padding:5px 12px; font-size:13px;
  color:#2b3340; text-decoration:none; }}
nav.toc a:hover {{ border-color:#b8c1cd; }}
nav.toc a span {{ color:var(--dim); font-variant-numeric:tabular-nums; }}
.toolbar {{ position:sticky; top:0; z-index:9; background:rgba(246,247,249,.94);
  backdrop-filter:blur(6px); border-bottom:1px solid var(--line); padding:11px 0; margin:18px 0 6px;
  display:flex; gap:9px; flex-wrap:wrap; align-items:center; }}
.toolbar input {{ flex:1 1 210px; padding:7px 11px; border:1px solid #cfd6df; border-radius:8px;
  font-size:13.5px; background:#fff; }}
.chip {{ border:1px solid #cfd6df; background:#fff; border-radius:999px; padding:5px 12px;
  font-size:12.5px; cursor:pointer; color:#2b3340; user-select:none; }}
.chip[aria-pressed="true"] {{ background:#11151c; color:#fff; border-color:#11151c; }}
section.chapter {{ background:var(--card); border:1px solid var(--line); border-radius:13px;
  padding:6px 22px 18px; margin:20px 0; }}
h2.chap {{ font-size:20px; margin:16px 0 4px; display:flex; gap:12px; align-items:baseline; }}
h2.chap .count {{ font-size:12px; font-weight:500; color:var(--dim); background:#f1f3f6;
  border-radius:999px; padding:2px 10px; }}
h3.subsec {{ font-size:12.5px; text-transform:uppercase; letter-spacing:.09em; color:var(--dim);
  margin:22px 0 6px; border-left:3px solid #d7dce4; padding-left:9px; }}
.entry {{ border-top:1px solid #eef1f4; padding:10px 0 12px; }}
.entry-head {{ display:flex; gap:9px; align-items:center; flex-wrap:wrap; font-size:12.5px; }}
.badge {{ font-size:11px; font-weight:700; letter-spacing:.06em; text-transform:uppercase;
  border-radius:5px; padding:2px 7px; }}
.b-added {{ color:var(--add); background:var(--addbg); }}
.b-removed {{ color:var(--del); background:var(--delbg); }}
.b-modified {{ color:var(--mod); background:var(--modbg); }}
.b-format {{ color:var(--fmt); background:var(--fmtbg); }}
.b-moved {{ color:#4a5568; background:#eef1f5; }}
.tag {{ color:var(--dim); }}
.meta {{ color:var(--dim); font-variant-numeric:tabular-nums; }}
.anchor {{ margin-left:auto; color:#b9c1cd; text-decoration:none; font-size:12px; }}
.entry:hover .anchor {{ color:#6b7684; }}
.diff {{ margin-top:7px; white-space:pre-wrap; overflow-wrap:anywhere; font-size:15px; line-height:1.62; }}
.diff.quiet {{ color:#3d4550; }}
ins.ins {{ background:var(--addbg); color:var(--add); text-decoration:none;
  border-bottom:1px solid #a9dcbb; padding:1px 1px; }}
del.del {{ background:var(--delbg); color:var(--del); text-decoration:line-through;
  text-decoration-color:#e0a49f; padding:1px 1px; }}
ul.notes {{ margin:6px 0 0; padding-left:17px; font-size:12.5px; color:var(--dim); }}
ul.notes li {{ margin:1px 0; }}
.tblwrap {{ overflow-x:auto; margin-top:8px; }}
table.tbl {{ border-collapse:collapse; font-size:13.5px; background:#fff; }}
table.tbl td {{ border:1px solid #e6e9ee; padding:5px 9px; vertical-align:top; }}
tr.row-add td {{ background:var(--addbg); }}
tr.row-del td {{ background:var(--delbg); }}
tr.row-mod td {{ background:#fcfcfd; }}
td.cell-mod {{ background:var(--modbg); }}
tr.more td {{ color:var(--dim); font-style:italic; text-align:center; }}
.hl {{ font-size:10.5px; font-weight:700; color:#fff; background:#8794a5; border-radius:4px; padding:1px 5px; }}
.hl1 {{ background:#3d4a5c; }} .hl2 {{ background:#5b6b80; }}
footer {{ margin-top:44px; color:var(--dim); font-size:12.5px; border-top:1px solid var(--line); padding-top:16px; }}
body.hide-added .entry.op-added, body.hide-removed .entry.op-removed,
body.hide-modified .entry.op-modified, body.hide-format .entry.op-format,
body.hide-moved .entry.op-moved, body.hide-tbl .entry.kind-tbl {{ display:none; }}
.entry.nomatch, section.chapter.nomatch, h3.subsec.nomatch {{ display:none; }}
@media print {{ .toolbar, nav.toc, .anchor {{ display:none; }}
  section.chapter {{ break-inside:auto; border:0; }} body {{ background:#fff; }} }}
</style>
</head>
<body>
<header class="top">
  <h1>Change Log — Royal Commonwealth of Zeythari</h1>
  <div class="sub">Content-level diff of two Word documents: every paragraph, table row and
  heading compared, not a binary file comparison.</div>
  <div class="files">
    <div><b>Before &mdash; 16th Edition</b>{esc(st['old_file'])}<br>{st['old_paras']:,} paragraphs · {st['old_tables']:,} tables · {st['old_words']:,} words{(' · saved ' + old_date) if old_date else ''}</div>
    <div><b>After &mdash; Diff Merge</b>{esc(st['new_file'])}<br>{st['new_paras']:,} paragraphs · {st['new_tables']:,} tables · {st['new_words']:,} words{(' · saved ' + new_date) if new_date else ''}</div>
  </div>
</header>
<main>

<div class="cards">
  <div class="card mod"><div class="k">Paragraphs changed</div><div class="v">{n_mod_p}</div></div>
  <div class="card mod"><div class="k">Tables changed</div><div class="v">{n_mod_t}</div></div>
  <div class="card add"><div class="k">Blocks added</div><div class="v">{st['added']}</div></div>
  <div class="card del"><div class="k">Blocks removed</div><div class="v">{st['removed']}</div></div>
  <div class="card fmt"><div class="k">Formatting only</div><div class="v">{st['format']}</div></div>
  <div class="card"><div class="k">Moved</div><div class="v">{st['moved']}</div></div>
  <div class="card add"><div class="k">Words added</div><div class="v">+{st['words_added']:,}</div></div>
  <div class="card del"><div class="k">Words removed</div><div class="v">−{st['words_removed']:,}</div></div>
</div>

<h2 class="sec">Outline changes &mdash; {len(added_heads)} headings added, {len(removed_heads)} removed, {len(renames)} retitled, {len(levels)} restyled</h2>
{outline or '<p class="dim">No heading-level changes.</p>'}

<h2 class="sec">Chapter overview</h2>
<table class="grid">
<thead><tr><th>Chapter</th><th class="num">Changes</th><th class="num">Added</th><th class="num">Removed</th>
<th class="num">Changed</th><th class="num">Formatting</th><th class="num">+ words</th><th class="num">− words</th><th class="num">Δ words</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table>

<h2 class="sec">Media parts</h2>
{media}

{revisions}

<h2 class="sec">Jump to chapter</h2>
<nav class="toc">{nav}</nav>

<div class="toolbar">
  <input id="q" type="search" placeholder="Filter by word or phrase…" aria-label="Filter changes">
  <button class="chip" data-f="added" aria-pressed="true">Added</button>
  <button class="chip" data-f="removed" aria-pressed="true">Removed</button>
  <button class="chip" data-f="modified" aria-pressed="true">Changed</button>
  <button class="chip" data-f="format" aria-pressed="true">Formatting</button>
  <button class="chip" data-f="moved" aria-pressed="true">Moved</button>
  <button class="chip" data-f="tbl" aria-pressed="true">Tables</button>
</div>

<h2 class="sec">Detailed change log &mdash; {total} entries</h2>
{''.join(sections)}

<footer>
  Generated by a structural diff: both documents were converted to a body-order block sequence
  (paragraphs and tables, with style, list and run-formatting signatures). Blocks were aligned
  with a sequence matcher and, inside each rewritten region, re-paired by a similarity metric
  that blends character-level similarity with word overlap. Surviving pairs were diffed
  word-by-word (character-level inside single words); tables were diffed row-by-row and
  cell-by-cell. Blocks that disappeared from one place and reappeared unchanged elsewhere are
  reported as moves.
</footer>
</main>
<script>
const chips = document.querySelectorAll('.chip');
chips.forEach(c => c.addEventListener('click', () => {{
  const on = c.getAttribute('aria-pressed') === 'true';
  c.setAttribute('aria-pressed', String(!on));
  document.body.classList.toggle('hide-' + c.dataset.f, on);
}}));
const q = document.getElementById('q');
q.addEventListener('input', () => {{
  const v = q.value.trim().toLowerCase();
  document.querySelectorAll('.entry').forEach(el => {{
    el.classList.toggle('nomatch', !!v && !el.textContent.toLowerCase().includes(v));
  }});
  document.querySelectorAll('section.chapter').forEach(sec => {{
    const hits = [...sec.querySelectorAll('.entry')].filter(e => !e.classList.contains('nomatch'));
    sec.classList.toggle('nomatch', hits.length === 0);
    sec.querySelectorAll('h3.subsec').forEach(h3 => {{
      let n = h3.nextElementSibling, any = false;
      while (n && n.classList.contains('entry')) {{
        if (!n.classList.contains('nomatch')) any = true;
        n = n.nextElementSibling;
      }}
      h3.classList.toggle('nomatch', !any);
    }});
  }});
}});
</script>
</body>
</html>
"""
    with open(out_path, "w") as fh:
        fh.write(page)
    print(f"wrote {out_path}: {len(page):,} bytes, {total} entries, "
          f"{len(chapter_list)} chapters")


if __name__ == "__main__":
    main(*sys.argv[1:5],
         old_date=sys.argv[5] if len(sys.argv) > 5 else "",
         new_date=sys.argv[6] if len(sys.argv) > 6 else "")
