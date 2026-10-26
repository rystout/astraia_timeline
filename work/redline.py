#!/usr/bin/env python3
"""Write the diff between two .docx files into one document as Word revisions.

The 16th Edition is the base document and the merge is applied on top of it, so
accepting every revision produces the merged text and rejecting every revision
restores the original. Paragraph edits are word-level (character-level inside a
replaced word), table edits are row- and cell-level, and inserted/deleted
paragraphs carry their paragraph-mark revisions so accept/reject behave.
"""
import copy
import json
import os
import re
import sys
import zipfile

import docx
from lxml import etree

import diff as D
import extract as X

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
AUTHOR = "Codebuff Merge"
DATE = "2026-10-26T00:00:00Z"
SPECIAL = (W + "drawing", W + "pict", W + "fldSimple", W + "instrText",
           W + "object", W + "footnoteReference", W + "commentReference")


class Rev:
    def __init__(self):
        self.n = 1000

    def wrap(self, tag):
        el = etree.Element(W + tag)
        el.set(W + "id", str(self.n))
        el.set(W + "author", AUTHOR)
        el.set(W + "date", DATE)
        self.n += 1
        return el


# ------------------------------------------------------------------- text model
def segs_of(p):
    """Text-bearing nodes of a paragraph in order, with character offsets."""
    out, pos = [], 0
    for node in p.iter():
        tag = node.tag
        if tag == W + "t":
            txt = node.text or ""
        elif tag == W + "tab":
            txt = "\t"
        elif tag in (W + "br", W + "cr"):
            txt = "\n"
        elif tag == W + "noBreakHyphen":
            txt = "-"
        else:
            continue
        if not txt:
            continue
        run = node.getparent()
        while run is not None and run.tag != W + "r":
            run = run.getparent()
        out.append({"s": pos, "e": pos + len(txt), "tag": tag, "text": txt,
                    "run": run, "node": node})
        pos += len(txt)
    return out


def text_of(segs):
    return "".join(s["text"] for s in segs)


def run_range(segs, a, b):
    return [s for s in segs if s["e"] > a and s["s"] < b]


def clone_run(run_el, node, text):
    r = etree.Element(W + "r")
    if run_el is not None:
        rpr = run_el.find(W + "rPr")
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
    if node.tag == W + "t":
        t = etree.SubElement(r, W + "t")
        t.text = text
        t.set(XML_SPACE, "preserve")
    else:
        r.append(copy.deepcopy(node))
    return r


def emit(segs, a, b, mode, out, rev):
    """Emit runs covering [a,b) of `segs`, as-is, deleted, or inserted."""
    for s in run_range(segs, a, b):
        lo, hi = max(a, s["s"]), min(b, s["e"])
        txt = s["text"][lo - s["s"]:hi - s["s"]]
        if not txt:
            continue
        r = clone_run(s["run"], s["node"], txt)
        if mode == "del":
            for t in r.findall(W + "t"):
                t.tag = W + "delText"
            w = rev.wrap("del")
            w.append(r)
            out.append(w)
        elif mode == "ins":
            w = rev.wrap("ins")
            w.append(r)
            out.append(w)
        else:
            out.append(r)


def has_special(p):
    return any(p.find(".//" + tag) is not None for tag in SPECIAL)


def set_content(p, children):
    keep = [ch for ch in p if ch.tag in (W + "bookmarkStart", W + "bookmarkEnd",
                                         W + "commentRangeStart", W + "commentRangeEnd")]
    for ch in list(p):
        if ch.tag != W + "pPr":
            p.remove(ch)
    for ch in keep:
        p.append(ch)
    for ch in children:
        p.append(ch)


def mark_paragraph(p, tag, rev):
    """Mark the paragraph mark itself as inserted or deleted."""
    ppr = p.find(W + "pPr")
    if ppr is None:
        ppr = etree.Element(W + "pPr")
        p.insert(0, ppr)
    rpr = ppr.find(W + "rPr")
    if rpr is None:
        rpr = etree.Element(W + "rPr")
        # CT_PPr order: ... rPr, sectPr, pPrChange
        sect = ppr.find(W + "sectPr")
        if sect is not None:
            sect.addprevious(rpr)
        else:
            ppr.append(rpr)
    # CT_ParaRPr order: ins, del, moveFrom, moveTo, rStyle, ...
    rpr.insert(0, rev.wrap(tag))


# ------------------------------------------------------------- revision cleanup
def accept_revisions(el):
    """Unwrap w:ins, drop w:del, so copied XML carries plain final-state text."""
    for tag in ("del", "moveFrom"):
        for d in el.findall(".//" + W + tag):
            d.getparent().remove(d)
    for tag in ("ins", "moveTo"):
        for i in el.findall(".//" + W + tag):
            parent = i.getparent()
            idx = parent.index(i)
            for k, child in enumerate(list(i)):
                parent.insert(idx + k, child)
            parent.remove(i)


def wrap_all_runs(el, rev):
    """Wrap every run under `el` in a single w:ins (or w:del when tag='del')."""
    pass


def mark_inserted_block(el, rev):
    for p in el.iter(W + "p"):
        runs = [ch for ch in p if ch.tag == W + "r"]
        if runs:
            w = rev.wrap("ins")
            idx = p.index(runs[0])
            for r in runs:
                p.remove(r)
                w.append(r)
            p.insert(idx, w)
        mark_paragraph(p, "ins", rev)


def mark_deleted_block(el, rev):
    for p in el.iter(W + "p"):
        runs = [ch for ch in p if ch.tag == W + "r"]
        if runs:
            w = rev.wrap("del")
            idx = p.index(runs[0])
            for r in runs:
                p.remove(r)
                for t in r.findall(W + "t"):
                    t.tag = W + "delText"
                w.append(r)
            p.insert(idx, w)
        mark_paragraph(p, "del", rev)


def set_row_revision(tr, tag, rev):
    trpr = tr.find(W + "trPr")
    if trpr is None:
        trpr = etree.Element(W + "trPr")
        tr.insert(0, trpr)
    trpr.append(rev.wrap(tag))


# ---------------------------------------------------------------- paragraph ops
def replace_whole_paragraph(p_old, p_new, rev):
    """Delete every run of the old paragraph and insert every run of the new one.

    Used when a paragraph holds something runs-and-text logic would drop (images,
    fields, content controls) so that rejecting the change restores it intact.
    """
    pn = copy.deepcopy(p_new)
    accept_revisions(pn)
    out = []
    for r in [ch for ch in p_old if ch.tag == W + "r"]:
        r2 = copy.deepcopy(r)
        for t in r2.findall(W + "t"):
            t.tag = W + "delText"
        w = rev.wrap("del")
        w.append(r2)
        out.append(w)
    for r in [ch for ch in pn if ch.tag == W + "r"]:
        w = rev.wrap("ins")
        w.append(copy.deepcopy(r))
        out.append(w)
    set_content(p_old, out)


def rewrite_paragraph(p_old, p_new, rev):
    """Turn one changed paragraph into an insertion/deletion pair of runs."""
    if has_special(p_old) or has_special(p_new):
        replace_whole_paragraph(p_old, p_new, rev)
        return

    so, sn = segs_of(p_old), segs_of(p_new)
    a, b = text_of(so), text_of(sn)
    ops = D.inline_diff(a, b)
    out = []
    o = n = 0
    for op, txt in ops:
        if op == "=":
            emit(so, o, o + len(txt), "keep", out, rev)
            o += len(txt)
            n += len(txt)
        elif op == "-":
            emit(so, o, o + len(txt), "del", out, rev)
            o += len(txt)
        else:
            emit(sn, n, n + len(txt), "ins", out, rev)
            n += len(txt)
    set_content(p_old, out)


def rpr_signature(run):
    rpr = run.find(W + "rPr") if run is not None else None
    return etree.tostring(rpr) if rpr is not None else b""


def rewrite_formatting(p_old, p_new, rev):
    """Same text, different formatting: mark only the runs that changed."""
    so, sn = segs_of(p_old), segs_of(p_new)
    ro, rn = [], []
    for segs, acc in ((so, ro), (sn, rn)):
        cur = None
        for s in segs:
            if cur is not None and cur["run"] is s["run"]:
                cur["e"] = s["e"]
            else:
                cur = {"s": s["s"], "e": s["e"], "run": s["run"]}
                acc.append(cur)

    if len(ro) != len(rn):
        out = []
        emit(so, 0, len(text_of(so)), "del", out, rev)
        emit(sn, 0, len(text_of(sn)), "ins", out, rev)
        set_content(p_old, out)
        return

    out = []
    for o_r, n_r in zip(ro, rn):
        if rpr_signature(o_r["run"]) == rpr_signature(n_r["run"]):
            emit(so, o_r["s"], o_r["e"], "keep", out, rev)
        else:
            emit(so, o_r["s"], o_r["e"], "del", out, rev)
            emit(sn, n_r["s"], n_r["e"], "ins", out, rev)
    set_content(p_old, out)


# -------------------------------------------------------------------- table ops
def rewrite_cell(tc_old, tc_new, rev):
    po = tc_old.findall(W + "p")
    pn = tc_new.findall(W + "p")
    if len(po) == 1 and len(pn) == 1:
        rewrite_paragraph(po[0], pn[0], rev)
        return
    # fall back to replacing the whole cell content paragraph by paragraph
    for p in po[1:]:
        mark_deleted_block(p, rev)
    if po and pn:
        rewrite_paragraph(po[0], pn[0], rev)


def apply_table_change(tbl_old, tbl_new, entry, rev):
    rows_o = tbl_old.findall(W + "tr")
    rows_n = tbl_new.findall(W + "tr") if tbl_new is not None else []
    delta = entry.get("table") or []
    row_map = entry.get("row_map") or []
    new_to_old = {nm: om for om, nm in row_map}

    # 1. edits and deletions inside the existing rows
    for ch in delta:
        if ch["op"] == "removed":
            tr = rows_o[ch["old_row"]]
            mark_deleted_block(tr, rev)
            set_row_revision(tr, "del", rev)
        elif ch["op"] == "modified":
            tr_o = rows_o[ch["old_row"]]
            tr_n = rows_n[ch["new_row"]]
            cells_o = tr_o.findall(W + "tc")
            cells_n = tr_n.findall(W + "tc")
            for cell in ch["cells"]:
                if cell["same"]:
                    continue
                k = cell["col"]
                if k < len(cells_o) and k < len(cells_n):
                    rewrite_cell(cells_o[k], cells_n[k], rev)

    # 2. new rows, each anchored after the row it follows in the merged table
    last_after = {}
    added = sorted((c for c in delta if c["op"] == "added"),
                   key=lambda c: c["new_row"])
    for ch in added:
        tr = copy.deepcopy(rows_n[ch["new_row"]])
        accept_revisions(tr)
        mark_inserted_block(tr, rev)
        set_row_revision(tr, "ins", rev)
        anchor = None
        for nm in range(ch["new_row"] - 1, -1, -1):
            if nm in new_to_old:
                anchor = rows_o[new_to_old[nm]]
                break
        if anchor is None:
            first = tbl_old.find(W + "tr")
            if first is None:
                tbl_old.append(tr)
            else:
                first.addprevious(tr)
        else:
            prev = last_after.get(id(anchor), anchor)
            prev.addnext(tr)
            last_after[id(anchor)] = tr


# ------------------------------------------------------------------------- main
def build(old_docx, new_docx, changes_json, out_docx, workdir="work"):
    data = json.load(open(changes_json))
    changes = data["changes"]
    alignment = data["alignment"]
    moved = {tuple(m) for m in data.get("moved_pairs", [])}

    def blocks_of(name, docx_path):
        path = os.path.join(workdir, name, "word", "document.xml")
        if not os.path.exists(path):
            os.makedirs(os.path.join(workdir, name), exist_ok=True)
            with zipfile.ZipFile(docx_path) as z:
                z.extractall(os.path.join(workdir, name))
        tree = etree.parse(path)
        body = tree.getroot().find(W + "body")
        blocks, pi, ti = [], 0, 0
        for el in body:
            if el.tag == W + "p":
                blocks.append((X.para_info(el, pi), el))
                pi += 1
            elif el.tag == W + "tbl":
                blocks.append((X.table_info(el, ti), el))
                ti += 1
        return tree, body, [(info, el) for info, el in blocks if not D.is_noise(info)]

    tree_o, body_o, fo = blocks_of("oldpkg", old_docx)
    tree_n, body_n, fn = blocks_of("newpkg", new_docx)
    print(f"aligned blocks: old {len(fo)} new {len(fn)}")

    rev = Rev()
    # for an insertion at new index j, insert before the old element belonging
    # to the first alignment pair whose new index is greater than j
    al_sorted = sorted(alignment, key=lambda x: x["new"])
    al_new = [a["new"] for a in al_sorted]

    def anchor_for(j):
        import bisect
        k = bisect.bisect_right(al_new, j)
        return al_sorted[k]["old"] if k < len(al_sorted) else None

    # ---- in-place edits: changed paragraphs, changed tables, deletions ------
    ent_by_pair = {}
    for e in changes:
        if e["op"] in ("modified", "format"):
            ent_by_pair[(e["old_index"], e["new_index"])] = e

    for a in alignment:
        if a["op"] not in ("modified", "format"):
            continue
        e = ent_by_pair.get((a["old"], a["new"]))
        if e is None:
            continue
        info_o, el_o = fo[a["old"]]
        info_n, el_n = fn[a["new"]]
        if info_o["kind"] == "p" and info_n["kind"] == "p":
            if a["op"] == "format":
                rewrite_formatting(el_o, el_n, rev)
            else:
                rewrite_paragraph(el_o, el_n, rev)
        elif info_o["kind"] == "tbl" and info_n["kind"] == "tbl":
            apply_table_change(el_o, el_n, e, rev)

    for oi in data["unpaired_old"]:
        _, el_o = fo[oi]
        if el_o.tag == W + "p":
            mark_deleted_block(el_o, rev)
        else:
            for tr in el_o.findall(W + "tr"):
                mark_deleted_block(tr, rev)
                set_row_revision(tr, "del", rev)

    for oi, ni in moved:
        _, el_o = fo[oi]
        mark_deleted_block(el_o, rev)

    # ---- insertions ---------------------------------------------------------
    inserts = [(nj, nj) for nj in data["unpaired_new"]]
    inserts += [(ni, ni) for _, ni in moved]
    # ascending: each insertion goes immediately before its anchor, so later
    # insertions (whose anchors are later) land after earlier ones
    inserts.sort(key=lambda t: t[0])

    n_ins_par = n_ins_tbl = 0
    for _, nj in inserts:
        info_n, el_n = fn[nj]
        new_el = copy.deepcopy(el_n)
        accept_revisions(new_el)
        if info_n["kind"] == "p":
            mark_inserted_block(new_el, rev)
            n_ins_par += 1
        else:
            for tr in new_el.findall(W + "tr"):
                mark_inserted_block(tr, rev)
                set_row_revision(tr, "ins", rev)
            n_ins_tbl += 1
        anchor_old = anchor_for(nj)
        if anchor_old is None:
            sect = body_o.find(W + "sectPr")
            if sect is not None:
                sect.addprevious(new_el)
            else:
                body_o.append(new_el)
        else:
            body_o.insert(body_o.index(fo[anchor_old][1]), new_el)

    print(f"inserted paragraphs/tables: {n_ins_par}/{n_ins_tbl}")

    # ---- write the package --------------------------------------------------
    new_xml = etree.tostring(tree_o, xml_declaration=True, encoding="UTF-8",
                             standalone=True)
    if os.path.exists(out_docx):
        os.remove(out_docx)
    with zipfile.ZipFile(old_docx) as zin, \
            zipfile.ZipFile(out_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data_in = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data_in = new_xml
            elif item.filename == "docProps/core.xml":
                txt = data_in.decode("utf-8")
                txt = re.sub(r"<dc:title>[^<]*</dc:title>",
                             "<dc:title>Royal Commonwealth of Zeythari \u2014 "
                             "Merged (tracked changes)</dc:title>", txt)
                txt = re.sub(r"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)",
                             r"\g<1>2026-10-26T02:05:00Z\g<2>", txt)
                txt = re.sub(r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>",
                             "<cp:lastModifiedBy>Codebuff Merge</cp:lastModifiedBy>", txt)
                data_in = txt.encode("utf-8")
            zout.writestr(item, data_in)
    print("wrote", out_docx, os.path.getsize(out_docx), "bytes")
    return out_docx


# ------------------------------------------------------------------ validation
def view_text(path, mode):
    """Document text with revisions accepted or rejected."""
    tree = etree.parse(os.path.join(path, "word", "document.xml"))
    body = tree.getroot().find(W + "body")
    out = []
    for p in body.iter(W + "p"):
        buf = []
        for node in p.iter():
            if node.tag == W + "t":
                anc = [a.tag for a in node.iterancestors()]
                if mode == "accept" and W + "del" in anc:
                    continue
                if mode == "reject" and W + "ins" in anc:
                    continue
                buf.append(node.text or "")
            elif node.tag == W + "delText":
                if mode == "reject":
                    buf.append(node.text or "")
            elif node.tag == W + "tab":
                buf.append("\t")
            elif node.tag in (W + "br", W + "cr"):
                buf.append("\n")
        out.append("".join(buf))
    return out


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4],
          sys.argv[5] if len(sys.argv) > 5 else "work")
