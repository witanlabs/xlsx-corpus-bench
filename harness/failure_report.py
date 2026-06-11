#!/usr/bin/env python3
"""Generate a team-facing markdown report of one library's failures on a
corpus run: every file that failed load, round-trip, or recalc, grouped by
error signature, with repro commands.

Usage: failure_report.py --lib witan --results-dir results-fuse --corpus corpus-fuse --out FILE.md
"""
import argparse
import json
import os
import re
import sys
from collections import defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def signature(msg: str) -> str:
    """First meaningful line, with file-specific tokens generalized."""
    line = (msg or "(no detail)").strip().split("\n")[0]
    line = re.sub(r"xl/(drawings|diagrams)/_rels/\w+\.xml\.rels", r"xl/\1/_rels/<part>.xml.rels", line)
    return line[:140]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--lib", required=True)
    ap.add_argument("--results-dir", required=True)
    ap.add_argument("--corpus", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    results_dir = os.path.join(ROOT, args.results_dir)
    rows = [json.loads(l) for l in open(os.path.join(results_dir, f"{args.lib}.jsonl"))]
    n = len(rows)

    phases = {
        "load": [r for r in rows if not r["load"]["ok"]],
        "round-trip": [r for r in rows if r["load"]["ok"] and not r["roundtrip"].get("ok")],
        "recalc": [
            r for r in rows
            if r["load"]["ok"] and r["recalc"].get("supported") and not r["recalc"].get("ok")
        ],
    }

    def err_of(r, phase):
        if phase == "load":
            return r["load"].get("error")
        if phase == "round-trip":
            return r["roundtrip"].get("opc_error") or r["roundtrip"].get("error")
        return r["recalc"].get("error")

    lines = [
        f"# {args.lib} failures on the {args.results_dir.replace('results-', '')} corpus",
        "",
        f"Source: corpus benchmark run over {n} unique real-world workbooks "
        f"(see README.md in this repo for methodology). Raw per-file results: "
        f"`{args.results_dir}/{args.lib}.jsonl`. Corpus files are local under "
        f"`{args.corpus}/` (sha-named; FUSE files are public Common Crawl content).",
        "",
        "| phase | failures | rate |",
        "|---|---|---|",
    ]
    for phase, fails in phases.items():
        lines.append(f"| {phase} | {len(fails)} | {100 * len(fails) / n:.2f}% |")
    lines += [
        "",
        "Definitions: **load** = open + enumerate sheets; **round-trip** = "
        "load → save → reload, output must also pass library-neutral OPC "
        "structural validation (zip integrity, XML well-formedness, all "
        "relationship targets resolve); **recalc** = `calc --verify` "
        "completed (mismatches vs cached values are a separate metric, not "
        "listed here).",
        "",
        "Failed round-trip outputs are preserved under "
        f"`{args.results_dir}/failures/{args.lib}/<sha>.xlsx` for inspection.",
        "",
    ]

    for phase, fails in phases.items():
        if not fails:
            continue
        groups = defaultdict(list)
        for r in fails:
            groups[signature(err_of(r, phase))].append(r)
        lines.append(f"## {phase} failures ({len(fails)})")
        lines.append("")
        for sig, members in sorted(groups.items(), key=lambda kv: -len(kv[1])):
            lines.append(f"### {len(members)} × `{sig}`")
            lines.append("")
            for r in members:
                lines.append(f"- `{args.corpus}/{r['path']}`")
            lines.append("")
            example = members[0]["path"]
            cmd = {
                "load": f"xlsx-serve --json exec {args.corpus}/{example} --expr 'await xlsx.listSheets(wb)'",
                "round-trip": f"cp {args.corpus}/{example} /tmp/rt.xlsx && printf '{{\"id\":\"1\",\"workbook\":\"/tmp/rt.xlsx\",\"op\":\"open\",\"args\":{{}}}}\\n{{\"id\":\"2\",\"workbook\":\"/tmp/rt.xlsx\",\"op\":\"save\",\"args\":{{}}}}\\n' | xlsx-serve && python3 harness/opc_validate.py /tmp/rt.xlsx",
                "recalc": f"xlsx-serve --json calc {args.corpus}/{example} --verify",
            }[phase]
            lines += ["Repro (first file):", "", "```bash", cmd, "```", ""]

    with open(args.out, "w") as f:
        f.write("\n".join(lines) + "\n")
    print(f"wrote {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
