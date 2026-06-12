#!/usr/bin/env python3
"""Aggregate results/<lib>.jsonl files into results/REPORT.md.

Reported per library, all percentages over the same manifest:
  load %        — files opened without error
  round-trip %  — files whose load->save->reload output also passed the
                  library-neutral OPC validation
  recalc clean %— of files with >=1 formula cell where recalc ran: files with
                  zero mismatches vs the cached values Excel stored in them
  cell match %  — formula cells matching cached values, across all such files
Libraries without a calculation engine show N/A (not 0%) for recalc columns.
"""
import glob
import json
import os
import sys
from collections import defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RESULTS = os.path.join(ROOT, "results-sb")
if len(sys.argv) > 1:
    RESULTS = os.path.abspath(sys.argv[1])


def pct(n: int, d: int) -> str:
    return f"{100.0 * n / d:.1f}%" if d else "—"


def main() -> int:
    libs = {}
    for path in sorted(glob.glob(os.path.join(RESULTS, "*.jsonl"))):
        name = os.path.basename(path).rsplit(".", 1)[0]
        if name.endswith(".raw") or name == "manifest":
            continue
        rows = [json.loads(l) for l in open(path) if l.strip()]
        if rows and "load" in rows[0]:  # skip non-benchmark jsonl (skips, sidecars)
            libs[name] = rows

    manifest_path = os.path.join(RESULTS, "manifest.jsonl")
    total = sum(1 for _ in open(manifest_path)) if os.path.exists(manifest_path) else 0

    # one universe for every tool: files Excel itself could process (truth
    # exists). Files Excel cannot open/save are excluded from numerator and
    # denominator of every metric for every library.
    truth_set = {
        os.path.basename(p)[:-5]
        for p in glob.glob(os.path.join(RESULTS, "excel-truth", "*.xlsx"))
        if not os.path.basename(p).startswith("~")
    }
    excluded = 0
    if truth_set:
        for lib in list(libs):
            before = len(libs[lib])
            libs[lib] = [r for r in libs[lib] if r["sha256"] in truth_set]
            excluded = max(excluded, before - len(libs[lib]))

    truth = {}
    for path in glob.glob(os.path.join(RESULTS, "truth-compare", "*.jsonl")):
        name = os.path.basename(path).rsplit(".", 1)[0]
        rows = [json.loads(l) for l in open(path) if l.strip()]
        ok = [r for r in rows if r.get("ok") and r.get("formula_cells", 0) > 0]
        if ok:
            truth[name] = {
                "files": len(ok),
                "clean": sum(1 for r in ok if r["mismatches"] == 0),
                "cells": sum(r["formula_cells"] for r in ok),
                "matched": sum(r["formula_cells"] - r["mismatches"] for r in ok),
            }

    universe_note = (
        f"Universe: {total - excluded:,} workbooks — the {total:,} unique files "
        f"minus {excluded} that Excel itself could not process; every library "
        "is measured on exactly this set, and a library failure inside it "
        "counts against that library rather than shrinking its denominator."
        if truth_set
        else f"Corpus: {total:,} unique workbooks (sha256-deduplicated)."
    )
    lines = [
        "# xlsx corpus benchmark — results",
        "",
        universe_note + " See README.md for methodology and definitions.",
        "",
        "Recalculation is judged against real Excel: Microsoft Excel "
        "recomputed every workbook (harness/excel_truth.py) and each "
        "engine's results are compared to Excel's by one shared comparator "
        "(harness/compare_truth.py). The two recalculation columns answer "
        "different questions: how many *workbooks* come out perfect "
        "(all-or-nothing per file), and how many *individual formulas* "
        "match (overall accuracy; a few big failing workbooks can move "
        "this a lot without moving the first).",
        "",
        "| library | workbooks | opens without error | survives open→save→reopen | workbooks recalculated 100% Excel-identical | formula cells matching Excel |",
        "|---|---|---|---|---|---|",
    ]

    summary = {}
    for lib, rows in sorted(libs.items()):
        n = len(rows)
        load_ok = sum(1 for r in rows if r["load"]["ok"])
        rt_ok = sum(1 for r in rows if r["roundtrip"].get("ok"))
        calc_rows = [
            r for r in rows
            if r["recalc"].get("supported") and r["recalc"].get("ok") and r["recalc"].get("formula_cells", 0) > 0
        ]
        supported = any(r["recalc"].get("supported") for r in rows)
        clean = sum(1 for r in calc_rows if r["recalc"].get("mismatches", 0) == 0)
        cells = sum(r["recalc"]["formula_cells"] for r in calc_rows)
        matched = sum(r["recalc"]["formula_cells"] - r["recalc"].get("mismatches", 0) for r in calc_rows)
        t = truth.get(lib)
        if t:
            truth_clean = f"{pct(t['clean'], t['files'])} of {t['files']:,}"
            truth_cells = f"{pct(t['matched'], t['cells'])} of {t['cells']:,}"
        elif supported:
            truth_clean = truth_cells = "(not yet compared)"
        else:
            reason = next(
                (r["recalc"].get("reason") for r in rows if r["recalc"].get("reason")),
                "no calculation engine",
            )
            truth_clean = truth_cells = f"N/A ({reason})"
        lines.append(
            f"| {lib} | {n:,} | {pct(load_ok, n)} | {pct(rt_ok, n)} | {truth_clean} | {truth_cells} |"
        )
        summary[lib] = {
            "files": n,
            "load_ok": load_ok,
            "roundtrip_ok": rt_ok,
            "recalc_supported": supported,
            "recalc_files_with_formulas": len(calc_rows),
            "recalc_clean_files": clean,
            "formula_cells": cells,
            "formula_cells_matched": matched,
        }

    # top load-error signatures per library, so failures are inspectable
    lines += ["", "## Top load-error signatures", ""]
    for lib, rows in sorted(libs.items()):
        errs = defaultdict(int)
        for r in rows:
            if not r["load"]["ok"] and r["load"].get("error"):
                errs[r["load"]["error"].split("\n")[0][:120]] += 1
        if not errs:
            continue
        lines.append(f"**{lib}**")
        for msg, c in sorted(errs.items(), key=lambda kv: -kv[1])[:5]:
            lines.append(f"- {c} × `{msg}`")
        lines.append("")

    out_md = os.path.join(RESULTS, "REPORT.md")
    with open(out_md, "w") as f:
        f.write("\n".join(lines) + "\n")
    with open(os.path.join(RESULTS, "summary.json"), "w") as f:
        json.dump(summary, f, indent=2)
    sync_readme(summary, truth)
    print("\n".join(lines))
    return 0


DISPLAY = {"witan": "witan", "openpyxl": "openpyxl", "epplus": "EPPlus",
           "closedxml": "ClosedXML", "libreoffice": "LibreOffice"}


def sync_readme(summary: dict, truth: dict) -> None:
    """Rewrite this corpus's results table in README.md between marker
    comments, so the README can never drift from the receipts. The best
    value per column is bolded — by rule, not editorial choice."""
    readme = os.path.join(ROOT, "README.md")
    tag = os.path.basename(RESULTS)
    start, end = f"<!-- table:{tag}:start -->", f"<!-- table:{tag}:end -->"
    if not os.path.exists(readme):
        return
    text = open(readme).read()
    if start not in text or end not in text:
        return

    rows = []
    for lib, s in summary.items():
        n = s["files"]
        t = truth.get(lib)
        rows.append({
            "lib": DISPLAY.get(lib, lib),
            "load": s["load_ok"] / n if n else 0.0,
            "rt": s["roundtrip_ok"] / n if n else 0.0,
            "clean": (t["clean"] / t["files"]) if t else None,
            "cells": (t["matched"] / t["cells"]) if t else None,
            "supported": s["recalc_supported"],
        })
    rows.sort(key=lambda r: (r["cells"] is None, -(r["cells"] or 0), -r["rt"]))
    best = {
        k: max((r[k] for r in rows if r[k] is not None), default=None)
        for k in ("load", "rt", "clean", "cells")
    }

    def cell(r, k):
        if r[k] is None:
            return "N/A — no calculation engine" if k == "clean" else ""
        v = f"{100 * r[k]:.1f}%"
        return f"**{v}**" if r[k] == best[k] else v

    denom = next((f"{t['cells']:,}" for t in truth.values()), "?")
    table = [
        f"| library | opens without error | survives open→save→reopen | workbooks recalculated 100% Excel-identical | formula cells matching Excel (of {denom}) |",
        "|---|---|---|---|---|",
    ]
    for r in rows:
        table.append(
            f"| {r['lib']} | {cell(r, 'load')} | {cell(r, 'rt')} | {cell(r, 'clean')} | {cell(r, 'cells')} |"
        )
    block = start + "\n" + "\n".join(table) + "\n" + end
    pre, rest = text.split(start, 1)
    _, post = rest.split(end, 1)
    with open(readme, "w") as f:
        f.write(pre + block + post)


if __name__ == "__main__":
    sys.exit(main())
