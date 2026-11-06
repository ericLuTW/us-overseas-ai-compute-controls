#!/usr/bin/env python3
"""Regenerate the dashboard's embedded data blobs (window.__DATA_ZH / __DATA_EN)
from the source workbooks, and splice them into index.html.

The dashboard (index.html) is a self-contained Claude-Design React app whose
data lives in two inline <script> blocks:
    <script>window.__DATA_ZH = {...}</script>
    <script>window.__DATA_EN = {...}</script>

This script rebuilds those JSON objects from the ZH/EN workbooks. Three
structures are NOT derivable from the workbooks and are carried over verbatim
from the current index.html (they were authored at design time):
    GLOSS   - 66-entry abbreviation glossary (predates the v2 framework 詞彙表)
    nodes   - N1..N15 supply-chain node display names + group assignment
    GRP_EN  - node-group EN label map

Since 2026-11-06 the three design-time statics live in tools/statics.json
(written once by `extract-statics`); the old index.html is only the fallback.
Each blob also carries LAYERS (institutional-layer view), always computed from
the ZH workbook and copied into both language blobs.

Usage:
    python3 tools/gen_data_blobs.py extract-statics                      # index.html -> tools/statics.json (once)
    python3 tools/gen_data_blobs.py check   [--allow-path-regex R] <zh.xlsx> <en.xlsx>
    python3 tools/gen_data_blobs.py splice  <zh.xlsx> <en.xlsx>        # rewrite blobs present in index.html
    python3 tools/gen_data_blobs.py emit-js <zh.xlsx> <en.xlsx> <outdir>  # write data_zh.js / data_en.js
    python3 tools/gen_data_blobs.py inject  --asset-id ID [--html PATH] <zh.xlsx> <en.xlsx>
    python3 tools/gen_data_blobs.py bundle  [--html PATH]   # inline <dc-import> child .dc.html files (offline file://)
"""
import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
INDEX = ROOT / "index.html"
STATICS = ROOT / "tools" / "statics.json"

# LAYERS: reference-table term -> baseline entry it names (name identity)
LAYER_REF = {"ECRA": "BL-1", "FDPR": "BL-2", "ENC": "BL-5", "MEU": "BL-7", "EL": "BL-9"}
CODE_RE = r"(?:EW|EB|LC|M|BL)-\d+[a-e]?"
ZAI_RE = re.compile(r"載於\s*(" + CODE_RE + r"(?:\s*[、，,與及]\s*" + CODE_RE + r")*)")

# per-language column headers: blob field -> header name
POL_COLS = {
    "zh": {"code": "policy_id", "item": "條目（中文）", "id": "編號", "date": "日期",
           "prog": "進度", "who": "頒布者／提案人", "law": "法源", "bg": "背景與脈絡",
           "body": "內容", "links": "來源連結", "disp": "新碼"},
    "en": {"code": "policy_id", "item": "Entry (English)", "id": "Citation", "date": "Date",
           "prog": "Status", "who": "Issuer / Sponsor", "law": "Legal Basis",
           "bg": "Background & Context", "body": "Content", "links": "Source Links",
           "disp": "New Code"},
}
TOOL_COLS = {
    "zh": {"tool": "政策工具", "act": "工具作用", "appl": "適用對象", "dir": "作用方向",
           "macro": "工具大類", "grp": "工具群組", "trig": "觸發要件／門檻",
           "op": "操作方式", "spec": "具體標的", "law": "授權／制定主體",
           "impl": "實施機關", "note": "備註"},
    "en": {"tool": "Policy Tool", "act": "Tool Function", "appl": "Applicable Parties",
           "dir": "Direction", "macro": "Tool Category", "grp": "Tool Group",
           "trig": "Trigger Conditions / Thresholds", "op": "Mode of Operation",
           "spec": "Specific Targets", "law": "Authorizing / Issuing Body",
           "impl": "Implementing Agency", "note": "Notes"},
}
SHEETS = {
    "zh": {"pol": "政策條目", "tools": "政策工具盤點", "cat": "標的×工具長表",
           "node": "節點×工具長表", "polnode": "政策×節點長表", "vocab": "詞彙表"},
    "en": {"pol": "Policy Entries", "tools": "Policy Tool Inventory", "cat": "Targets × Tools",
           "node": "Nodes × Tools", "polnode": "Policies × Nodes", "vocab": "Glossary"},
}


def sheet_rows(ws):
    hdr = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    out = []
    for r in range(2, ws.max_row + 1):
        row = {h: ws.cell(row=r, column=c).value for h, c in hdr.items() if h}
        if any(v not in (None, "") for v in row.values()):
            out.append(row)
    return out


def s(v):
    if v is None:
        return ""
    return str(v).strip()


def extract_current_blobs(html):
    """Return {lang: blob} for the language blocks actually present in html."""
    blobs = {}
    for lang in ("ZH", "EN"):
        m = re.search(r"<script>window\.__DATA_" + lang + r" = ", html)
        if not m:
            continue
        start = m.end()
        end = html.index("</script>", start)
        blobs[lang.lower()] = json.loads(html[start:end].rstrip().rstrip(";"))
    return blobs


def build(path, lang, statics):
    wb = openpyxl.load_workbook(path, data_only=True)
    sh = SHEETS[lang]
    pc, tc = POL_COLS[lang], TOOL_COLS[lang]

    pol_rows = sheet_rows(wb[sh["pol"]])
    tool_rows = sheet_rows(wb[sh["tools"]])
    cat_rows = sheet_rows(wb[sh["cat"]])
    node_rows = sheet_rows(wb[sh["node"]])
    vocab_rows = sheet_rows(wb[sh["vocab"]])

    # policies
    pol, links, code2, track_of = [], {}, {}, {}
    pol_by_code = {}
    for r in pol_rows:
        code = s(r["policy_id"])
        disp = s(r[pc["disp"]])
        date = s(r[pc["date"]])
        p = {"code": code, "disp": disp, "item": s(r[pc["item"]]), "id": s(r[pc["id"]]),
             "year": date[:4], "date": date, "prog": s(r[pc["prog"]]),
             "who": s(r[pc["who"]]), "law": s(r[pc["law"]]), "bg": s(r[pc["bg"]]),
             "body": s(r[pc["body"]]), "track": disp.split("-")[0] if disp else ""}
        pol.append(p)
        pol_by_code[code] = p
        if s(r[pc["links"]]):
            # workbook separates multiple URLs with ' ｜ '; the blob uses newlines
            links[code] = re.sub(r"\s*｜\s*", "\n", s(r[pc["links"]]))
        code2[code] = disp
        track_of[code] = p["track"]

    # tools (row order defines numeric tool no 1..N)
    tools = {}
    code_of = {}
    for i, r in enumerate(tool_rows, 1):
        pid = s(r["policy_id"])
        parent = pol_by_code[pid]
        t = {"no": i, "tool": s(r[tc["tool"]]), "dir": s(r[tc["dir"]]),
             "macro": s(r[tc["macro"]]), "grp": s(r[tc["grp"]]),
             "prog": parent["prog"], "item": parent["item"],
             "spec": s(r[tc["spec"]]), "act": s(r[tc["act"]]), "op": s(r[tc["op"]]),
             "law": s(r[tc["law"]]), "appl": s(r[tc["appl"]]), "trig": s(r[tc["trig"]]),
             "impl": s(r[tc["impl"]]), "note": s(r[tc["note"]])}
        tools[s(r["tool_id"])] = t
        code_of[str(i)] = pid

    def light(t):
        return {"no": t["no"], "tool": t["tool"], "dir": t["dir"]}

    # cells: target_category ||| macro -> [full tool]
    cells = {}
    tool_cats = {}
    for r in cat_rows:
        t = tools[s(r["tool_id"])]
        cat = s(r["target_category"])
        cells.setdefault(cat + "|||" + t["macro"], []).append(t)
        tool_cats.setdefault(t["no"], []).append(cat)

    # nodecells: node ||| macro -> [light]; and tool -> nodes
    # node_id "0" marks tools without a supply-chain node -> excluded everywhere,
    # EXCEPT baseline (BL-track) tools: framework regimes apply chain-wide, so a
    # BL tool whose only mapping is "0" is presented on ALL nodes (Eric 2026-09-24).
    # BL tools with explicit node mappings use those nodes only.
    all_node_ids = [n["id"] for n in statics["nodes"]]
    raw_tool_nodes = {}
    for r in node_rows:
        t = tools[s(r["tool_id"])]
        raw_tool_nodes.setdefault(t["no"], []).append(s(r["node_id"]))

    def effective_nodes(t):
        nids = [n for n in raw_tool_nodes.get(t["no"], []) if n != "0"]
        if nids:
            return nids
        pid = code_of[str(t["no"])]
        if track_of.get(pid) == "BL":
            return list(all_node_ids)
        return []

    nodecells = {}
    tool_nodes = {}
    polcells = {}
    for t in sorted(tools.values(), key=lambda x: x["no"]):
        for nid in effective_nodes(t):
            nodecells.setdefault(nid + "|||" + t["macro"], []).append(light(t))
            tool_nodes.setdefault(t["no"], []).append(nid)
            pid = code_of[str(t["no"])]
            polcells.setdefault(pid + "|||" + nid, []).append(light(t))

    # bridge: target_category ||| node -> [light] (tools mapped to both)
    bridge = {}
    seen_bridge = set()
    for no, cats in tool_cats.items():
        for cat in cats:
            for nid in tool_nodes.get(no, []):
                key = cat + "|||" + nid
                if (key, no) in seen_bridge:
                    continue
                seen_bridge.add((key, no))
                t = next(t for t in tools.values() if t["no"] == no)
                bridge.setdefault(key, []).append(light(t))

    # vocab-driven lists
    def vocab(domain):
        return [r for r in vocab_rows if s(r["vocab_domain"]) == domain]

    name_key = "name_zh" if lang == "zh" else "name_en"
    macros = [s(r[name_key]) for r in vocab("tool_major")]
    cats = [s(r[name_key]) for r in vocab("target_category")]
    macros_en = [s(r["name_en"]) for r in vocab("tool_major")]
    cats_en = [s(r["name_en"]) for r in vocab("target_category")]
    tracks = [{"code": s(r["code"]), "zh": s(r["name_zh"]), "en": s(r["name_en"])}
              for r in vocab("track")]

    trk_idx = {tr["code"]: i for i, tr in enumerate(tracks)}
    pol.sort(key=lambda p: (trk_idx.get(p["track"], 99), p["disp"]))

    return {
        "D": {"macros": macros, "cats": cats, "nodes": statics["nodes"],
              "cells": cells, "nodecells": nodecells, "polcells": polcells,
              "bridge": bridge, "pol": pol},
        "CODE_OF": code_of, "LINKS": links, "GRP_EN": statics["GRP_EN"],
        "GLOSS": statics["GLOSS"], "trackOf": track_of, "TRACKS": tracks,
        "cats_en": cats_en, "macros_en": macros_en, "CODE2": code2,
    }


def load_statics(current):
    if STATICS.exists():
        return json.loads(STATICS.read_text(encoding="utf-8"))
    zh = current["zh"]
    return {"GLOSS": zh["GLOSS"], "GRP_EN": zh["GRP_EN"], "nodes": zh["D"]["nodes"]}


def build_layers(zh_path):
    """Institutional layers from the ZH workbook (codes are internal policy_id)."""
    wb = openpyxl.load_workbook(zh_path, data_only=True)
    pol_rows = sheet_rows(wb["政策條目"])
    new2old = {s(r["新碼"]): s(r["policy_id"]) for r in pol_rows}
    layer, basis, zai = {}, {}, {}
    for r in pol_rows:
        code = s(r["policy_id"])
        if not code.startswith("BL"):
            layer[code] = "targeted"
            continue
        layer[code] = "basis" if code in ("BL-1", "BL-10") else "baseline"
        if layer[code] == "baseline":
            law = s(r["法源"])
            b = (["BL-1"] if ("出口管制改革法" in law or "ECRA" in law) else []) \
                + (["BL-10"] if "瓦聖納" in law else [])
            if b:
                basis[code] = b
        found = []
        for m in ZAI_RE.finditer("\n".join(s(v) for v in r.values())):
            for disp in re.findall(CODE_RE, m.group(1)):
                if disp not in new2old:
                    raise ValueError(f"載於 code {disp} (in {code}) not found in 新碼")
                if new2old[disp] not in found:
                    found.append(new2old[disp])
        if found:
            zai[code] = found
    cites = {}
    for r in sheet_rows(wb["參照表"]):
        rid = s(r["ref_id"])
        tgt = LAYER_REF.get(rid) or (
            "BL-8" if s(r["ref_type"]) == "國家群組" and re.fullmatch(r"[A-E]:\d", rid) else None)
        if not tgt:
            continue
        for c in re.split(r"[、,\s]+", s(r["引用碼"])):
            if layer.get(c) == "targeted":
                cites.setdefault(tgt, set()).add(c)
    cites = {k: sorted(v) for k, v in sorted(cites.items())}
    linked = {c for v in cites.values() for c in v} | {c for v in zai.values() for c in v}
    unlinked = sorted(c for c, lv in layer.items() if lv == "targeted" and c not in linked)
    return {"layer": layer, "basis": basis, "zai": zai, "cites": cites, "unlinked": unlinked}


def layers_summary(L):
    n = Counter(L["layer"].values())
    cites = " ".join(f"{k} {len(v)}" for k, v in L["cites"].items())
    return (f"LAYERS summary: basis {n['basis']} baseline {n['baseline']} targeted {n['targeted']}; "
            f"zai {sum(len(v) for v in L['zai'].values())}; cites {cites}; unlinked {len(L['unlinked'])}")


def deep_diff(a, b, path="", out=None, limit=None):
    if out is None:
        out = []
    if limit is not None and len(out) >= limit:
        return out
    if type(a) is not type(b):
        out.append(f"{path}: type {type(a).__name__} != {type(b).__name__}")
    elif isinstance(a, dict):
        for k in a.keys() | b.keys():
            if k not in a:
                out.append(f"{path}.{k}: extra in generated")
            elif k not in b:
                out.append(f"{path}.{k}: missing in generated")
            else:
                deep_diff(a[k], b[k], f"{path}.{k}", out, limit)
    elif isinstance(a, list):
        if len(a) != len(b):
            out.append(f"{path}: len {len(a)} != {len(b)}")
        for i, (x, y) in enumerate(zip(a, b)):
            deep_diff(x, y, f"{path}[{i}]", out, limit)
    elif a != b:
        out.append(f"{path}: {str(a)[:60]!r} != {str(b)[:60]!r}")
    return out


def parse_args():
    ap = argparse.ArgumentParser(description="Regenerate dashboard data blobs from the workbooks.")
    sub = ap.add_subparsers(dest="mode", required=True)
    sub.add_parser("extract-statics")
    c = sub.add_parser("check")
    c.add_argument("--allow-path-regex")
    sp = sub.add_parser("splice")
    e = sub.add_parser("emit-js")
    i = sub.add_parser("inject")
    i.add_argument("--asset-id", required=True)
    i.add_argument("--html", default=str(INDEX))
    b = sub.add_parser("bundle")
    b.add_argument("--html", default=str(INDEX))
    for p in (c, sp, e, i):
        p.add_argument("zh")
        p.add_argument("en")
    e.add_argument("outdir")
    return ap.parse_args()


def blob_js(blob):
    return json.dumps(blob, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")


BUNDLE_RE = re.compile(r'<script id="dc-bundle">.*?</script>\n?', re.S)
SUPPORT_TAG = '<script src="./support.js"></script>'


def bundle(html_path):
    """Inline every <dc-import name="X"> child (X.dc.html beside html_path) through the
    runtime's window.__resources / __resourceBlobs hooks, so the page works from file://
    (browsers block fetch() there). Re-running replaces the existing block."""
    path = Path(html_path)
    html = path.read_text(encoding="utf-8")
    names = sorted(set(re.findall(r'<dc-import name="([A-Za-z0-9_]+)"', BUNDLE_RE.sub("", html))))
    missing = [n for n in names if not (path.parent / f"{n}.dc.html").exists()]
    if not names or missing:
        print("no <dc-import> found" if not names else f"missing child files: {missing}")
        sys.exit(1)
    files = {f"./{n}.dc.html": (path.parent / f"{n}.dc.html").read_text(encoding="utf-8") for n in names}
    tag = ('<script id="dc-bundle">(function(){var S=' + blob_js(files) + ';'
           'var R=window.__resources=window.__resources||{};'
           'var B=window.__resourceBlobs=window.__resourceBlobs||{};'
           'for(var k in S){R[k]=k;B[k]=new Blob([S[k]],{type:"text/html"});}})();</script>\n')
    if BUNDLE_RE.search(html):
        html = BUNDLE_RE.sub(lambda _: tag, html, count=1)
    elif html.count(SUPPORT_TAG) == 1:
        html = html.replace(SUPPORT_TAG, tag + SUPPORT_TAG)
    else:
        print(f"expected exactly 1 {SUPPORT_TAG}, found {html.count(SUPPORT_TAG)}")
        sys.exit(1)
    path.write_text(html, encoding="utf-8")
    print(f"bundled {len(names)} components into {path.name}: {', '.join(names)}")


def main():
    args = parse_args()
    if args.mode == "bundle":
        bundle(args.html)
        sys.exit(0)
    html = INDEX.read_text(encoding="utf-8")
    current = extract_current_blobs(html)

    if args.mode == "extract-statics":
        if STATICS.exists():
            print("statics.json exists; refuse to overwrite")
            sys.exit(2)
        zh = current["zh"]
        st = {"GLOSS": zh["GLOSS"], "GRP_EN": zh["GRP_EN"], "nodes": zh["D"]["nodes"]}
        STATICS.write_text(json.dumps(st, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
        print(f"wrote {STATICS.name}: GLOSS {len(st['GLOSS'])}, GRP_EN {len(st['GRP_EN'])}, "
              f"nodes {len(st['nodes'])}")
        sys.exit(0)

    statics = load_statics(current)
    gen = {"zh": build(args.zh, "zh", statics), "en": build(args.en, "en", statics)}
    layers = build_layers(args.zh)
    for lang in gen:
        gen[lang]["LAYERS"] = layers

    if args.mode == "check":
        if not current:
            print("no __DATA_ZH/__DATA_EN block found")
            sys.exit(1)
        rx = re.compile(args.allow_path_regex) if args.allow_path_regex else None
        ok = True
        for lang in ("zh", "en"):
            if lang not in current:
                print(f"[{lang}] skipped: not in index.html")
                continue
            g = dict(gen[lang])
            if "LAYERS" not in current[lang]:
                g.pop("LAYERS")
            diffs = deep_diff(current[lang], g, lang)
            allowed = sum(1 for d in diffs if rx and rx.search(d))
            outside = len(diffs) - allowed
            print(f"[{lang}] diffs: {len(diffs)} (allowed {allowed}, outside {outside})")
            for d in diffs[:40]:
                print("   ", d)
            ok = ok and outside == 0
        print(layers_summary(layers))
        sys.exit(0 if ok else 1)

    if args.mode == "splice":
        if not current:
            print("no __DATA_ZH/__DATA_EN block found")
            sys.exit(1)
        for lang in ("ZH", "EN"):
            if lang.lower() not in current:
                continue
            m = re.search(r"<script>window\.__DATA_" + lang + r" = ", html)
            start = m.end()
            end = html.index("</script>", start)
            html = html[:start] + blob_js(gen[lang.lower()]) + html[end:]
        INDEX.write_text(html, encoding="utf-8")
        print("spliced index.html")
        for lang in current:
            d = gen[lang]
            print(f"  {lang}: pol={len(d['D']['pol'])} tools={len(d['CODE_OF'])} "
                  f"cells={len(d['D']['cells'])} tracks={[t['code'] for t in d['TRACKS']]}")
        print(layers_summary(layers))
        sys.exit(0)

    if args.mode == "emit-js":
        out = Path(args.outdir)
        out.mkdir(parents=True, exist_ok=True)
        for lang in ("zh", "en"):
            (out / f"data_{lang}.js").write_text(
                f"window.__DATA_{lang.upper()} = {blob_js(gen[lang])};\n", encoding="utf-8")
        print(f"wrote {out}/data_zh.js, {out}/data_en.js")
        print(layers_summary(layers))
        sys.exit(0)

    if args.mode == "inject":
        target = Path(args.html)
        doc = target.read_text(encoding="utf-8")
        pat = re.compile(r'<script[^>]*\bsrc="[^"]*' + re.escape(args.asset_id) + r'[^"]*"[^>]*>\s*</script>')
        hits = pat.findall(doc)
        if len(hits) != 1:
            print(f"inject: found {len(hits)} script tags with asset id {args.asset_id}; expected 1")
            sys.exit(1)
        doc = pat.sub(lambda m: "<script>window.__DATA_ZH = " + blob_js(gen["zh"]) + "</script>", doc, count=1)
        target.write_text(doc, encoding="utf-8")
        print(f"injected __DATA_ZH into {target}")
        print(layers_summary(layers))
        sys.exit(0)


if __name__ == "__main__":
    main()
