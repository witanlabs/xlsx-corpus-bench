#!/usr/bin/env python3
"""Corpus-wide diagnosis of witan recalc mismatches vs Excel truth.

For every file in results/truth-compare/witan.jsonl with mismatches,
re-produce witan's recalculated output (xlsx-serve calc), classify each
mismatching cell by the function names in its formula, and aggregate by
function signature across the corpus. Output: a team-facing markdown report
ranking root-cause candidates by impact (mismatched cells and files), with
example cells (address, formula, Excel's value, witan's value) and file
paths for repro.

Usage: recalc_gap_report.py --out results/WITAN-RECALC-GAPS.md
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cached_values import VOLATILE, extract, values_match  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_local = os.environ.get("WITAN_XLSX_SERVE")
WITAN_CALC = [_local, "calc"] if _local else ["witan", "xlsx", "calc"]
FN_RE = re.compile(r"(?:_xlfn\.)?([A-Z][A-Z0-9.]{2,})\s*\(")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--set",
        action="append",
        dest="sets",
        default=None,
        help="results_dir:corpus_dir pair, repeatable (default: both corpora)",
    )
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    sets = args.sets or ["results-sb:corpus-sb", "results-fuse:corpus-fuse"]

    # (corpus_dir, manifest_record, compare_record, truth_path) per unclean file
    unclean = []
    total_universe = 0
    for pair in sets:
        results_dir, corpus_dir = pair.split(":")
        compare = os.path.join(ROOT, results_dir, "truth-compare/witan.jsonl")
        if not os.path.exists(compare):
            print(f"skipping {pair}: no comparison yet", file=sys.stderr)
            continue
        by_sha = {
            json.loads(l)["sha256"]: json.loads(l)
            for l in open(os.path.join(ROOT, results_dir, "manifest.jsonl"))
        }
        for line in open(compare):
            comp = json.loads(line)
            if comp.get("ok") and comp.get("formula_cells", 0) > 0:
                total_universe += 1
                if comp.get("mismatches", 0) > 0:
                    unclean.append(
                        (
                            corpus_dir,
                            by_sha[comp["sha256"]],
                            comp,
                            os.path.join(ROOT, results_dir, "excel-truth", comp["sha256"] + ".xlsx"),
                        )
                    )
    print(f"{len(unclean)} files with mismatches (of {total_universe} compared)", file=sys.stderr)

    groups = defaultdict(lambda: {"cells": 0, "files": set(), "samples": []})
    total_cells = 0
    with tempfile.TemporaryDirectory() as tmp:
        for i, (corpus_dir, rec, comp, truth_path) in enumerate(unclean):
            work = os.path.join(tmp, "wb" + rec["ext"])
            shutil.copy(os.path.join(ROOT, corpus_dir, rec["path"]), work)
            try:
                subprocess.run(WITAN_CALC + [work, "--json"], capture_output=True, timeout=120)
                truth = extract(truth_path, with_formula=True)
                engine = extract(work)
            except Exception as e:
                print(f"  skip {rec['path']}: {e}", file=sys.stderr)
                continue
            for addr, t in truth.items():
                kind_t, val_t, formula = t[0], t[1], (t[2] if len(t) > 2 else "")
                if val_t is None or (formula and VOLATILE.search(formula)):
                    continue
                e = engine.get(addr)
                if e is None or e[1] is None:
                    key = "(cell missing from witan output)"
                elif not values_match((kind_t, val_t), (e[0], e[1])):
                    fns = sorted(set(FN_RE.findall(formula or "")))
                    key = ", ".join(fns[:4]) if fns else "(no function)"
                else:
                    continue
                g = groups[key]
                g["cells"] += 1
                g["files"].add(f"{corpus_dir}/{rec['path']}")
                total_cells += 1
                if len(g["samples"]) < 3:
                    g["samples"].append(
                        (f"{corpus_dir}/{rec['path']}", addr, (formula or "")[:90], repr(val_t)[:36],
                         repr(e[1])[:36] if e else "absent")
                    )
            os.unlink(work)
            if (i + 1) % 50 == 0:
                print(f"  {i + 1}/{len(unclean)}", file=sys.stderr)

    ranked = sorted(groups.items(), key=lambda kv: -kv[1]["cells"])
    lines = [
        "# witan recalc gaps vs Excel ground truth",
        "",
        f"Corpora: SpreadsheetBench + FUSE combined — {len(unclean)} workbooks "
        f"where witan's recalculation differs from Excel's, out of "
        f"{total_universe} formula-bearing workbooks compared (see "
        "results-sb/REPORT.md and results-fuse/REPORT.md). Ground truth: real "
        "Excel (forced full recalculation, see harness/excel_truth.py). "
        f"{total_cells} mismatched cells total, classified by the functions "
        "appearing in each cell's formula. Cascades inflate counts: one "
        "wrong upstream cell marks every dependent wrong too, so fix "
        "highest-impact groups first and re-run. Sample paths are relative "
        "to the xlsx-corpus-bench repo root (corpus-sb/ or corpus-fuse/).",
        "",
        "| functions in formula | mismatched cells | files |",
        "|---|---|---|",
    ]
    for key, g in ranked[:25]:
        lines.append(f"| `{key}` | {g['cells']} | {len(g['files'])} |")
    lines.append("")
    for key, g in ranked[:15]:
        lines.append(f"## `{key}` — {g['cells']} cells in {len(g['files'])} file(s)")
        lines.append("")
        for path, addr, formula, excel_v, witan_v in g["samples"]:
            lines.append(f"- `{addr}` in `{path}`")
            lines.append(f"  - formula: `{formula}`")
            lines.append(f"  - excel: `{excel_v}` · witan: `{witan_v}`")
        lines.append("")

    with open(args.out, "w") as f:
        f.write("\n".join(lines) + "\n")
    print(f"wrote {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
