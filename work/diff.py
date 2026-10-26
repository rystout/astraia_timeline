#!/usr/bin/env python3
"""Structural diff between two extracted block models.

Produces changes.json: aligned blocks, inline word diffs, table row/cell
deltas, formatting-only changes, moves, and chapter/section context.
"""
import json
import re
import sys
from difflib import SequenceMatcher

TOKEN = re.compile(r"\w+|\s+|[^\w\s]")
PAIR_THRESHOLD = 0.35


# ---------------------------------------------------------------- text helpers
def tokenize(s):
    return TOKEN.findall(s)


def merge_ops(ops):
    out = []
    for op, txt in ops:
        if not txt:
            continue
        if out and out[-1][0] == op:
            out[-1] = (op, out[-1][1] + txt)
        else:
            out.append((op, txt))
    return out


def char_refine(x, y):
    """Char-level diff for a single replaced word, so typos stay readable."""
    if " " in x or " " in y or len(x) > 80 or len(y) > 80:
        return [("-", x), ("+", y)]
    sm = SequenceMatcher(None, x, y, autojunk=False)
    out = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            out.append(("=", x[i1:i2]))
        elif tag == "delete":
            out.append(("-", x[i1:i2]))
        elif tag == "insert":
            out.append(("+", y[j1:j2]))
        else:
            out.append(("-", x[i1:i2]))
            out.append(("+", y[j1:j2]))
    return out


def inline_diff(a, b):
    at, bt = tokenize(a), tokenize(b)
    sm = SequenceMatcher(None, at, bt, autojunk=False)
    ops = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            ops.append(("=", "".join(at[i1:i2])))
        elif tag == "delete":
            ops.append(("-", "".join(at[i1:i2])))
        elif tag == "insert":
            ops.append(("+", "".join(bt[j1:j2])))
        else:
            x, y = "".join(at[i1:i2]), "".join(bt[j1:j2])
            if x and y and len(x.split()) == 1 and len(y.split()) == 1:
                ops.extend(char_refine(x, y))
            elif len(x) <= 40 and len(y) <= 40:
                ops.extend(char_refine(x, y))
            else:
                ops.append(("-", x))
                ops.append(("+", y))
    return merge_ops(ops)


def ratio(a, b):
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    sm = SequenceMatcher(None, a, b, autojunk=False)
    if sm.real_quick_ratio() < PAIR_THRESHOLD:
        return 0.0
    if sm.quick_ratio() < PAIR_THRESHOLD:
        return 0.0
    return sm.ratio()


# ------------------------------------------------------------------ model prep
def key_of(block):
    if block["kind"] == "p":
        return block["text"]
    return "\n".join(" | ".join(c["t"] for c in row) for row in block["rows"])


def is_noise(block):
    if block["kind"] == "p":
        return not block["text"] and not block["image"]
    return not block["rows"] or all(not c["t"] for r in block["rows"] for c in r)


def load(path):
    data = json.load(open(path))
    blocks = [b for b in data["blocks"] if not is_noise(b)]
    return data["meta"], blocks


def heading_level(block, style_names):
    if block["kind"] != "p":
        return 0
    st = style_names.get(block.get("style") or "", "") or ""
    m = re.match(r"Heading (\d)", st)
    if m:
        return int(m.group(1))
    if st == "Title":
        return 1
    if block.get("outline") is not None:
        try:
            return int(block["outline"]) + 1
        except (TypeError, ValueError):
            return 0
    return 0


def annotate_context(blocks, style_names):
    h1 = h2 = None
    for b in blocks:
        lvl = heading_level(b, style_names)
        if lvl == 1:
            h1, h2 = b["text"], None
        elif lvl == 2:
            h2 = b["text"]
        b["_h1"], b["_h2"], b["_lvl"] = h1, h2, lvl
        if b["kind"] == "p":
            b["_text"] = b["text"]
        else:
            b["_text"] = key_of(b)


# ------------------------------------------------------------------- formatting
def fmt_notes(a, b, style_names):
    notes = []
    if a["kind"] != b["kind"]:
        return ["block type changed"]
    if a["kind"] == "p":
        sa = style_names.get(a.get("style") or "", a.get("style") or "") or "Normal"
        sb = style_names.get(b.get("style") or "", b.get("style") or "") or "Normal"
        if sa != sb:
            if a["_lvl"] or b["_lvl"]:
                notes.append(f"heading level {a['_lvl'] or '-'} → {b['_lvl'] or '-'}"
                             if a["_lvl"] != b["_lvl"] else f"style “{sa}” → “{sb}”")
            else:
                notes.append(f"style “{sa}” → “{sb}”")
        elif a["_lvl"] != b["_lvl"]:
            notes.append(f"outline level {a['_lvl']} → {b['_lvl']}")
        na = (a.get("num_abs"), a.get("ilvl"), a.get("num_fmt"))
        nb = (b.get("num_abs"), b.get("ilvl"), b.get("num_fmt"))
        if bool(a.get("numid")) != bool(b.get("numid")):
            notes.append("list marking " + ("added" if b.get("numid") else "removed"))
        elif a.get("numid") and (na[1], na[2]) != (nb[1], nb[2]):
            notes.append(f"list format {na[2]} level {na[1]} → {nb[2]} level {nb[1]}")
        if a.get("jc") != b.get("jc"):
            notes.append(f"alignment {a.get('jc') or 'default'} → {b.get('jc') or 'default'}")
        if a.get("ind") != b.get("ind"):
            notes.append("indentation changed")
        fa, fb = set(a.get("formats") or []), set(b.get("formats") or [])
        if fa != fb:
            add = sorted(fb - fa)
            rem = sorted(fa - fb)
            bits = []
            if add:
                bits.append("+ " + ", ".join(add))
            if rem:
                bits.append("− " + ", ".join(rem))
            notes.append("run emphasis: " + "; ".join(bits))
        if a.get("image") != b.get("image"):
            notes.append("inline image " + ("added" if b.get("image") else "removed"))
    else:
        if len(a["rows"]) != len(b["rows"]):
            notes.append(f"{len(a['rows'])} → {len(b['rows'])} rows")
    return notes


# ------------------------------------------------------------------- alignment
WORD = re.compile(r"\w+")


def words_of(s):
    return set(w.lower() for w in WORD.findall(s))


def jaccard(a, b):
    wa, wb = words_of(a), words_of(b)
    if not wa or not wb:
        return 0.0
    return len(wa & wb) / len(wa | wb)


def sim(a, b, style_names):
    """Blend of character similarity (edits) and word overlap (rewrites),
    so near-identical text pairs even when long, and unrelated text does not."""
    if a["kind"] != b["kind"]:
        return 0.0
    ca = ratio(a["_text"], b["_text"])
    j = jaccard(a["_text"], b["_text"])
    s = 0.6 * ca + 0.4 * min(1.0, j * 1.5)
    if ca >= 0.65:
        s = max(s, ca)
    if a["kind"] == "p" and a["_lvl"] and a["_lvl"] == b["_lvl"]:
        s = min(1.0, s + 0.12)
    elif a["kind"] == "p" and (a["_lvl"] or b["_lvl"]):
        s *= 0.85
    return s


def pair_span(old, new, i1, i2, j1, j2, style_names):
    """Optimal matching inside a replace span, maximising summed similarity."""
    n, m = i2 - i1, j2 - j1
    if n == 0 or m == 0:
        return [], list(range(i1, i2)), list(range(j1, j2))
    sims = [[0.0] * m for _ in range(n)]
    for x in range(n):
        for y in range(m):
            sims[x][y] = sim(old[i1 + x], new[j1 + y], style_names)
    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    for x in range(n - 1, -1, -1):
        for y in range(m - 1, -1, -1):
            best = max(dp[x + 1][y], dp[x][y + 1])
            s = sims[x][y]
            if s >= PAIR_THRESHOLD:
                best = max(best, s + dp[x + 1][y + 1])
            dp[x][y] = best
    pairs, uo, un, x, y = [], [], [], 0, 0
    while x < n and y < m:
        s = sims[x][y]
        if s >= PAIR_THRESHOLD and abs(dp[x][y] - (s + dp[x + 1][y + 1])) < 1e-9:
            pairs.append((i1 + x, j1 + y))
            x += 1
            y += 1
        elif dp[x][y] == dp[x + 1][y]:
            uo.append(i1 + x)
            x += 1
        else:
            un.append(j1 + y)
            y += 1
    uo.extend(range(i1 + x, i2))
    un.extend(range(j1 + y, j2))
    return pairs, uo, un


# ----------------------------------------------------------------------- tables
def table_delta(a, b, style_names):
    ra = [" | ".join(c["t"] for c in r) for r in a["rows"]]
    rb = [" | ".join(c["t"] for c in r) for r in b["rows"]]
    sm = SequenceMatcher(None, ra, rb, autojunk=False)
    changes = []
    row_map = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            row_map.extend([i1 + k, j1 + k] for k in range(i2 - i1))
            continue
        if tag == "insert":
            for j in range(j1, j2):
                changes.append({"op": "added", "new_row": j, "row": b["rows"][j]})
        elif tag == "delete":
            for i in range(i1, i2):
                changes.append({"op": "removed", "old_row": i, "row": a["rows"][i]})
        else:
            pseudo_a = [{"kind": "p", "_text": t, "_lvl": 0, "text": t} for t in ra[i1:i2]]
            pseudo_b = [{"kind": "p", "_text": t, "_lvl": 0, "text": t} for t in rb[j1:j2]]
            pairs, uo, un = pair_span(pseudo_a, pseudo_b, 0, len(ra[i1:i2]),
                                      0, len(rb[j1:j2]), style_names)
            row_map.extend([i1 + pi, j1 + pj] for pi, pj in pairs)
            for pi, pj in pairs:
                rowa, rowb = a["rows"][i1 + pi], b["rows"][j1 + pj]
                cells = []
                for k in range(max(len(rowa), len(rowb))):
                    ta = rowa[k]["t"] if k < len(rowa) else None
                    tb = rowb[k]["t"] if k < len(rowb) else None
                    if ta == tb:
                        cells.append({"same": True, "text": tb})
                    else:
                        cells.append({"same": False, "col": k, "old": ta or "",
                                      "new": tb or "",
                                      "diff": inline_diff(ta or "", tb or "")})
                changes.append({"op": "modified", "old_row": i1 + pi,
                                "new_row": j1 + pj, "cells": cells})
            for pi in uo:
                changes.append({"op": "removed", "old_row": i1 + pi,
                                "row": a["rows"][i1 + pi]})
            for pj in un:
                changes.append({"op": "added", "new_row": j1 + pj,
                                "row": b["rows"][j1 + pj]})
    return changes, row_map


# -------------------------------------------------------------------------- run
def main(old_path, new_path, out_path):
    meta_o, old = load(old_path)
    meta_n, new = load(new_path)
    meta_n["style_names"].update({})
    style_names = dict(meta_o["style_names"])
    style_names.update(meta_n["style_names"])
    annotate_context(old, style_names)
    annotate_context(new, style_names)

    sm = SequenceMatcher(None, [b["_text"] for b in old],
                         [b["_text"] for b in new], autojunk=False)
    entries = []
    used_old, used_new = set(), set()
    pairs = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                pairs.append((i1 + k, j1 + k))
        elif tag == "replace":
            ps, uo, un = pair_span(old, new, i1, i2, j1, j2, style_names)
            pairs.extend(ps)
            entry = {"unpaired_old": uo, "unpaired_new": un}
            entries.append(("replace", ps, uo, un))
        else:
            entries.append((tag, [], list(range(i1, i2)), list(range(j1, j2))))

    # moves / adds / deletes, with cross-span move detection
    deleted, added = [], []
    for item in entries:
        if item[0] in ("delete", "replace"):
            deleted.extend(item[2])
        if item[0] in ("insert", "replace"):
            added.extend(item[3])

    moved_pairs = []
    remaining_added = list(added)
    for oi in list(deleted):
        target = None
        for aj in remaining_added:
            if old[oi]["kind"] == new[aj]["kind"] and old[oi]["_text"] == new[aj]["_text"]:
                target = aj
                break
        if target is not None and old[oi]["_text"]:
            moved_pairs.append((oi, target))
            deleted.remove(oi)
            remaining_added.remove(target)

    # ---------------------------------------------------------- emit entries
    out = []
    for oi, ni in pairs:
        a, b = old[oi], new[ni]
        notes = fmt_notes(a, b, style_names)
        if a["_text"] != b["_text"]:
            op = "modified"
            ops = inline_diff(a["_text"], b["_text"])
        elif notes:
            op = "format"
            ops = [("=", b["_text"])]
        else:
            continue
        e = {
            "op": op, "kind": b["kind"],
            "old_ref": f"¶{a['i'] + 1}" if a["kind"] == "p" else f"table {a['i'] + 1}",
            "new_ref": f"¶{b['i'] + 1}" if b["kind"] == "p" else f"table {b['i'] + 1}",
            "old_index": oi, "new_index": ni,
            "style": style_names.get(b.get("style") or "", b.get("style") or ""),
            "level": b["_lvl"], "notes": notes, "ops": ops,
            "old_text": a["_text"], "new_text": b["_text"],
            "h1": b["_h1"], "h2": b["_h2"],
            "h1_old": a["_h1"], "h2_old": a["_h2"],
        }
        if b["kind"] == "tbl":
            e["table"], e["row_map"] = table_delta(a, b, style_names)
            e["old_rows"] = len(a["rows"])
            e["new_rows"] = len(b["rows"])
            e["header"] = [c["t"] for c in b["rows"][0]] if b["rows"] else []
        out.append(e)

    for oi, ni in moved_pairs:
        a, b = old[oi], new[ni]
        out.append({
            "op": "moved", "kind": b["kind"],
            "old_ref": f"¶{a['i'] + 1}" if a["kind"] == "p" else f"table {a['i'] + 1}",
            "new_ref": f"¶{b['i'] + 1}" if b["kind"] == "p" else f"table {b['i'] + 1}",
            "old_index": oi, "new_index": ni,
            "style": style_names.get(b.get("style") or "", b.get("style") or ""),
            "level": b["_lvl"], "notes": [], "ops": [("=", b["_text"])],
            "old_text": a["_text"], "new_text": b["_text"],
            "h1": b["_h1"], "h2": b["_h2"], "h1_old": a["_h1"], "h2_old": a["_h2"],
        })

    for oi in sorted(deleted):
        a = old[oi]
        e = {
            "op": "removed", "kind": a["kind"],
            "old_ref": f"¶{a['i'] + 1}" if a["kind"] == "p" else f"table {a['i'] + 1}",
            "new_ref": None, "old_index": oi, "new_index": None,
            "style": style_names.get(a.get("style") or "", a.get("style") or ""),
            "level": a["_lvl"], "notes": [], "ops": [("-", a["_text"])],
            "old_text": a["_text"], "new_text": "",
            "h1": a["_h1"], "h2": a["_h2"], "h1_old": a["_h1"], "h2_old": a["_h2"],
        }
        if a["kind"] == "tbl":
            e["table"] = [{"op": "removed", "row": r} for r in a["rows"]]
            e["old_rows"] = len(a["rows"])
            e["new_rows"] = 0
        out.append(e)

    for ni in sorted(remaining_added):
        b = new[ni]
        e = {
            "op": "added", "kind": b["kind"],
            "old_ref": None,
            "new_ref": f"¶{b['i'] + 1}" if b["kind"] == "p" else f"table {b['i'] + 1}",
            "old_index": None, "new_index": ni,
            "style": style_names.get(b.get("style") or "", b.get("style") or ""),
            "level": b["_lvl"], "notes": [], "ops": [("+", b["_text"])],
            "old_text": "", "new_text": b["_text"],
            "h1": b["_h1"], "h2": b["_h2"], "h1_old": None, "h2_old": None,
        }
        if b["kind"] == "tbl":
            e["table"] = [{"op": "added", "row": r} for r in b["rows"]]
            e["old_rows"] = 0
            e["new_rows"] = len(b["rows"])
        out.append(e)

    # -------------------------------------------------------------- statistics
    WRE = re.compile(r"[^\W_]+")

    def words(ops, sign):
        return sum(len(WRE.findall(t)) for o, t in ops if o == sign)

    stats = {
        "old_blocks": len(old), "new_blocks": len(new),
        "old_paras": meta_o["paras"], "new_paras": meta_n["paras"],
        "old_tables": meta_o["tables"], "new_tables": meta_n["tables"],
        "old_words": meta_o["words"], "new_words": meta_n["words"],
        "added": sum(1 for e in out if e["op"] == "added"),
        "removed": sum(1 for e in out if e["op"] == "removed"),
        "modified": sum(1 for e in out if e["op"] == "modified"),
        "format": sum(1 for e in out if e["op"] == "format"),
        "moved": sum(1 for e in out if e["op"] == "moved"),
        "tables_added": sum(1 for e in out if e["op"] == "added" and e["kind"] == "tbl"),
        "tables_removed": sum(1 for e in out if e["op"] == "removed" and e["kind"] == "tbl"),
        "words_added": sum(words(e["ops"], "+") for e in out),
        "words_removed": sum(words(e["ops"], "-") for e in out),
        "old_media": len(meta_o["media"]), "new_media": len(meta_n["media"]),
        "old_file": meta_o["file"], "new_file": meta_n["file"],
        "media_old": meta_o["media"], "media_new": meta_n["media"],
        "revisions_old": meta_o.get("revisions", []),
        "revisions_new": meta_n.get("revisions", []),
    }

    heads_added = [e for e in out if e["op"] == "added" and e["level"]]
    heads_removed = [e for e in out if e["op"] == "removed" and e["level"]]
    stats["headings_added"] = len(heads_added)
    stats["headings_removed"] = len(heads_removed)

    # ---------------------------------------------------------- full alignment
    pair_ops = {}
    for e in out:
        if e["op"] in ("modified", "format") and e["old_index"] is not None \
                and e["new_index"] is not None:
            pair_ops[(e["old_index"], e["new_index"])] = e["op"]
    alignment = [{"old": oi, "new": ni,
                  "op": pair_ops.get((oi, ni), "equal")} for oi, ni in pairs]
    alignment.sort(key=lambda a: a["new"])

    with open(out_path, "w") as fh:
        json.dump({"stats": stats, "changes": out, "alignment": alignment,
                   "moved_pairs": [[oi, ni] for oi, ni in moved_pairs],
                   "unpaired_old": sorted(deleted),
                   "unpaired_new": sorted(remaining_added)}, fh)

    print(json.dumps(stats, indent=1)[:2000])
    print("\nheadings added:", [e["new_text"][:60] for e in heads_added][:40])
    print("\nheadings removed:", [e["old_text"][:60] for e in heads_removed][:40])


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
