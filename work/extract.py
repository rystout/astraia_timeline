#!/usr/bin/env python3
"""Extract a body-order block model from a .docx into JSON.

Each block is either a paragraph or a table. Formatting fingerprints are kept
so that formatting-only edits can be reported separately from text edits.
"""
import json
import re
import sys

import docx
from docx.oxml.ns import qn

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

WS = re.compile(r"\s+")


def norm(s):
    s = s.replace("\u00a0", " ").replace("\u200b", "").replace("\ufeff", "")
    s = WS.sub(" ", s)
    return s.strip()


def run_formats(r):
    """Compact formatting signature of a run."""
    f = []
    rpr = r.find(W + "rPr")
    if rpr is not None:
        for tag, code in (("b", "b"), ("i", "i"), ("u", "u"), ("strike", "s"),
                          ("caps", "c"), ("smallCaps", "sc"), ("vertAlign", "v")):
            el = rpr.find(W + tag)
            if el is not None and el.get(W + "val") not in ("0", "false", "none"):
                f.append(code)
        rf = rpr.find(W + "rFonts")
        if rf is not None:
            a = rf.get(W + "ascii")
            if a:
                f.append("font:" + a)
    return ",".join(f)


def para_text_and_runs(p):
    """Walk a paragraph's XML in document order collecting text and run formats."""
    out = []          # list of (text, fmt)
    for node in p.iter():
        if node.tag == W + "t":
            txt = node.text or ""
            if txt:
                r = node.getparent()
                while r is not None and r.tag != W + "r":   # runs nested in w:ins etc.
                    r = r.getparent()
                out.append((txt, run_formats(r) if r is not None else ""))
        elif node.tag == W + "tab":
            out.append(("\t", ""))
        elif node.tag == W + "br":
            out.append((" ", ""))
        elif node.tag == W + "noBreakHyphen":
            out.append(("-", ""))
    return out


def has_image(el):
    return el.find(".//" + W + "drawing") is not None or el.find(".//" + W + "pict") is not None


def numbering_maps(pkg_dir):
    """numId -> abstractNumId, and (abstractNumId, ilvl) -> numFmt."""
    import os
    from lxml import etree
    nmap, fmts = {}, {}
    path = os.path.join(pkg_dir, "word", "numbering.xml")
    if not os.path.exists(path):
        return nmap, fmts
    root = etree.parse(path).getroot()
    for num in root.findall(W + "num"):
        aid = num.find(W + "abstractNumId")
        if aid is not None:
            nmap[num.get(W + "numId")] = aid.get(W + "val")
    for an in root.findall(W + "abstractNum"):
        aid = an.get(W + "abstractNumId")
        for lvl in an.findall(W + "lvl"):
            nf = lvl.find(W + "numFmt")
            fmts[(aid, lvl.get(W + "ilvl"))] = nf.get(W + "val") if nf is not None else None
    return nmap, fmts


def para_info(p, idx):
    ppr = p.find(W + "pPr")
    style = None
    outline = None
    numid = ilvl = None
    jc = None
    ind = None
    if ppr is not None:
        ps = ppr.find(W + "pStyle")
        if ps is not None:
            style = ps.get(W + "val")
        ol = ppr.find(W + "outlineLvl")
        if ol is not None:
            outline = ol.get(W + "val")
        np = ppr.find(W + "numPr")
        if np is not None:
            n = np.find(W + "numId")
            l = np.find(W + "ilvl")
            numid = n.get(W + "val") if n is not None else None
            ilvl = l.get(W + "val") if l is not None else None
        j = ppr.find(W + "jc")
        if j is not None:
            jc = j.get(W + "val")
        i = ppr.find(W + "ind")
        if i is not None:
            ind = (i.get(W + "left"), i.get(W + "firstLine"), i.get(W + "hanging"))
    runs = para_text_and_runs(p)
    text = norm("".join(t for t, _ in runs))
    fmts = sorted({f for t, f in runs if t.strip() and f})
    return {
        "kind": "p",
        "i": idx,
        "style": style,
        "outline": outline,
        "numid": numid,
        "ilvl": ilvl,
        "jc": jc,
        "ind": ind,
        "text": text,
        "formats": fmts,
        "image": has_image(p),
    }


def cell_text(tc):
    parts = []
    for p in tc.findall(W + "p"):
        parts.append(para_info(p, -1)["text"])
    return norm(" ".join(parts))


def table_info(tbl, idx):
    rows = []
    for tr in tbl.findall(W + "tr"):
        tcs = tr.findall(W + "tc")
        # honour gridSpan: record span so column counts match visually
        row = []
        for tc in tcs:
            span = 1
            tcpr = tc.find(W + "tcPr")
            if tcpr is not None:
                gs = tcpr.find(W + "gridSpan")
                if gs is not None:
                    span = int(gs.get(W + "val") or 1)
            row.append({"t": cell_text(tc), "span": span})
        rows.append(row)
    style = None
    tblpr = tbl.find(W + "tblPr")
    if tblpr is not None:
        ts = tblpr.find(W + "tblStyle")
        if ts is not None:
            style = ts.get(W + "val")
    return {"kind": "tbl", "i": idx, "style": style, "rows": rows}


def revisions(block_el, label):
    """Tracked insertions/deletions still pending in the document."""
    out = []
    for tag, kind in (("ins", "inserted"), ("del", "deleted")):
        for el in block_el.iter(W + tag):
            ttag = W + "delText" if tag == "del" else W + "t"
            txt = norm("".join(t.text or "" for t in el.iter(ttag)))
            if txt:
                out.append({"at": label, "kind": kind, "text": txt})
    return out


def rel_images(path):
    """Media files referenced by the package, with sizes."""
    import os
    md = os.path.join(path, "word", "media")
    out = []
    if os.path.isdir(md):
        for f in sorted(os.listdir(md)):
            out.append({"name": f, "bytes": os.path.getsize(os.path.join(md, f))})
    return out


def main(docx_path, out_path, pkg_dir=""):
    d = docx.Document(docx_path)
    body = d.element.body
    blocks = []
    revs = []
    pi = ti = 0
    for el in body:
        if el.tag == W + "p":
            blocks.append(para_info(el, pi))
            revs.extend(revisions(el, f"¶{pi + 1}"))
            pi += 1
        elif el.tag == W + "tbl":
            blocks.append(table_info(el, ti))
            revs.extend(revisions(el, f"table {ti + 1}"))
            ti += 1
    styles = {}
    for s in d.styles:
        styles[s.style_id] = s.name
    nmap, fmts = numbering_maps(pkg_dir) if pkg_dir else ({}, {})
    for b in blocks:
        if b["kind"] == "p" and b.get("numid"):
            aid = nmap.get(b["numid"])
            b["num_abs"] = aid
            b["num_fmt"] = fmts.get((aid, b.get("ilvl") or "0"))
    meta = {
        "file": docx_path.split("/")[-1],
        "paras": pi,
        "tables": ti,
        "blocks": len(blocks),
        "style_names": styles,
        "media": rel_images(pkg_dir),
        "revisions": revs,
    }
    words = sum(len(b["text"].split()) for b in blocks if b["kind"] == "p")
    words += sum(len(c["t"].split()) for b in blocks if b["kind"] == "tbl"
                 for r in b["rows"] for c in r)
    meta["words"] = words
    with open(out_path, "w") as fh:
        json.dump({"meta": meta, "blocks": blocks}, fh)
    print(f"{docx_path}: {meta['blocks']} blocks, {meta['words']} words, "
          f"{meta['paras']} paras, {meta['tables']} tables, {len(meta['media'])} media")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "")
