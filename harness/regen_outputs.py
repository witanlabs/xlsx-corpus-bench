#!/usr/bin/env python3
"""Regenerate each library's round-trip outputs (load -> save) and KEEP them,
for the Excel-reopen oracle (harness/excel_repair_check.py). The benchmark
runs deleted passing outputs after OPC validation; this pass recreates them
without re-running recalc.

Outputs land in <results-dir>/rt-outputs/<lib>/<sha><ext>. Resumable: files
already present are skipped. A library that cannot save a given workbook
produces no output, which the repair check counts as a failure for that
library (same no-credit-for-skipped-work rule as everywhere else).

Usage: regen_outputs.py --lib {openpyxl,witan,epplus,closedxml,libreoffice}
                        --results-dir results --corpus corpus
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import warnings

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "runners"))

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX_SERVE = os.environ.get(
    "WITAN_XLSX_SERVE", os.path.join(ROOT, "..", "witan-alfred", "bin", "publish", "xlsx-serve")
)
DOTNET_RUNNER = os.path.join(ROOT, "harness/runners/dotnet/bin/Release/net10.0/DotnetRunner")


def todo_records(results: str, out_dir: str):
    records = []
    for line in open(os.path.join(results, "manifest.jsonl")):
        rec = json.loads(line)
        # LibreOffice conversion always emits .xlsx regardless of source ext
        if not (
            os.path.exists(os.path.join(out_dir, rec["sha256"] + rec["ext"]))
            or os.path.exists(os.path.join(out_dir, rec["sha256"] + ".xlsx"))
        ):
            records.append(rec)
    return records


def openpyxl_outputs(corpus, records, out_dir):
    import openpyxl

    for i, rec in enumerate(records):
        out = os.path.join(out_dir, rec["sha256"] + rec["ext"])
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                wb = openpyxl.load_workbook(
                    os.path.join(corpus, rec["path"]), keep_vba=rec["ext"] == ".xlsm"
                )
                wb.save(out)
        except Exception:
            if os.path.exists(out):
                os.unlink(out)
        if (i + 1) % 1000 == 0:
            print(f"  openpyxl {i + 1}/{len(records)}", file=sys.stderr)


def witan_outputs(corpus, records, out_dir):
    for i, rec in enumerate(records):
        out = os.path.join(out_dir, rec["sha256"] + rec["ext"])
        try:
            shutil.copy(os.path.join(corpus, rec["path"]), out)
            rpc = (
                json.dumps({"id": "1", "workbook": out, "op": "open", "args": {}})
                + "\n"
                + json.dumps({"id": "2", "workbook": out, "op": "save", "args": {}})
                + "\n"
            )
            p = subprocess.run([XLSX_SERVE], input=rpc.encode(), capture_output=True, timeout=120)
            responses = [json.loads(l) for l in p.stdout.decode().splitlines() if l.strip()]
            if len(responses) < 2 or not all(r.get("ok") for r in responses):
                os.unlink(out)
        except Exception:
            if os.path.exists(out):
                os.unlink(out)
        if (i + 1) % 1000 == 0:
            print(f"  witan {i + 1}/{len(records)}", file=sys.stderr)


def dotnet_outputs(lib, corpus, records, out_dir):
    with tempfile.NamedTemporaryFile("w", suffix=".jsonl", delete=False) as mf:
        for rec in records:
            mf.write(json.dumps(rec) + "\n")
        manifest = mf.name
    progress = os.path.join(out_dir, "_progress.jsonl")
    cmd = [
        DOTNET_RUNNER, "--lib", lib, "--mode", "save-only",
        "--corpus", corpus, "--manifest", manifest,
        "--out", progress, "--out-dir", out_dir,
    ]
    for _attempt in range(1000):
        proc = subprocess.run(cmd)
        if proc.returncode != 3:
            break
        print(f"  {lib} save-only restarting after wedged file", file=sys.stderr)
    os.unlink(manifest)
    if os.path.exists(progress):
        os.unlink(progress)


def libreoffice_outputs(corpus, records, out_dir):
    from runner_libreoffice import RECALC_NEVER, convert_with_retry

    with tempfile.TemporaryDirectory(prefix="lo_regen_") as work:
        stage = os.path.join(work, "stage")
        os.makedirs(stage)
        for i in range(0, len(records), 25):
            batch = records[i : i + 25]
            staged = []
            for rec in batch:
                dst = os.path.join(stage, rec["sha256"] + rec["ext"])
                shutil.copy(os.path.join(corpus, rec["path"]), dst)
                staged.append(dst)
            convert_with_retry(staged, out_dir, work, f"r{i}", RECALC_NEVER)
            for f in staged:
                os.unlink(f)
            if (i // 25) % 20 == 0:
                print(f"  libreoffice {min(i + 25, len(records))}/{len(records)}", file=sys.stderr)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--lib", required=True, choices=["openpyxl", "witan", "epplus", "closedxml", "libreoffice"])
    ap.add_argument("--results-dir", required=True)
    ap.add_argument("--corpus", required=True)
    args = ap.parse_args()

    results = os.path.join(ROOT, args.results_dir)
    corpus = os.path.join(ROOT, args.corpus)
    out_dir = os.path.join(results, "rt-outputs", args.lib)
    os.makedirs(out_dir, exist_ok=True)
    records = todo_records(results, out_dir)
    print(f"{args.lib}: {len(records)} outputs to generate", file=sys.stderr)

    if args.lib == "openpyxl":
        openpyxl_outputs(corpus, records, out_dir)
    elif args.lib == "witan":
        witan_outputs(corpus, records, out_dir)
    elif args.lib == "libreoffice":
        libreoffice_outputs(corpus, records, out_dir)
    else:
        dotnet_outputs(args.lib, corpus, records, out_dir)
    n = len(os.listdir(out_dir))
    print(f"{args.lib}: {n} outputs present", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
