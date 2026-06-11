#!/usr/bin/env python3
"""Orchestrator: run one library's runner over the corpus, then OPC-validate
its round-trip outputs with the library-neutral validator and merge the
verdicts into results/<lib>.jsonl.

A file's round-trip only passes if BOTH hold:
  - the library reloaded its own output without error
  - the output passes harness/opc_validate.py (independent structural check)

Round-trip outputs are deleted after validation; failing outputs are kept
under results/failures/<lib>/ for inspection.

Usage: run.py --lib {openpyxl,witan,epplus,closedxml,libreoffice} [--manifest ...]
"""
import argparse
import json
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from opc_validate import validate  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CORPUS = os.path.join(ROOT, "corpus-sb")
RESULTS = os.path.join(ROOT, "results-sb")

RUNNERS = {
    "openpyxl": lambda a: ["python3", os.path.join(ROOT, "harness/runners/runner_openpyxl.py"), *a],
    "witan": lambda a: ["node", os.path.join(ROOT, "harness/runners/runner_witan.mjs"), *a],
    "epplus": lambda a: [
        os.path.join(ROOT, "harness/runners/dotnet/bin/Release/net10.0/DotnetRunner"),
        "--lib", "epplus", *a,
    ],
    "closedxml": lambda a: [
        os.path.join(ROOT, "harness/runners/dotnet/bin/Release/net10.0/DotnetRunner"),
        "--lib", "closedxml", *a,
    ],
    "libreoffice": lambda a: ["python3", os.path.join(ROOT, "harness/runners/runner_libreoffice.py"), *a],
}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--lib", required=True, choices=sorted(RUNNERS))
    ap.add_argument("--results-dir", default=RESULTS, help="keep separate corpora's results separate")
    ap.add_argument("--manifest", default=None)
    ap.add_argument("--corpus", default=CORPUS)
    args = ap.parse_args()

    results = args.results_dir
    os.makedirs(results, exist_ok=True)
    if args.manifest is None:
        args.manifest = os.path.join(results, "manifest.jsonl")

    raw = os.path.join(results, f"{args.lib}.raw.jsonl")
    final = os.path.join(results, f"{args.lib}.jsonl")
    out_dir = os.path.join(results, "out", args.lib)
    failures_dir = os.path.join(results, "failures", args.lib)
    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(failures_dir, exist_ok=True)

    cmd = RUNNERS[args.lib](
        ["--corpus", args.corpus, "--manifest", args.manifest, "--out", raw, "--out-dir", out_dir]
    )
    print(f"running: {' '.join(cmd)}", file=sys.stderr)
    # exit code 3 = runner recorded a wedged file and wants a clean process;
    # results are append-only keyed by sha so the relaunch resumes after it
    for _attempt in range(1000):
        proc = subprocess.run(cmd)
        if proc.returncode != 3:
            break
        print("runner restarting after wedged file", file=sys.stderr)
    if proc.returncode != 0:
        print(f"runner exited {proc.returncode}; continuing with partial results", file=sys.stderr)

    n = 0
    with open(raw) as f, open(final, "w") as out:
        for line in f:
            rec = json.loads(line)
            rt = rec["roundtrip"]
            out_path = rt.get("out")
            if rt.get("ok") and out_path and os.path.exists(out_path):
                opc_ok, opc_err = validate(out_path)
                if not opc_ok:
                    # baseline gate: if the ORIGINAL already fails validation,
                    # the library can't be blamed for round-tripping that
                    # defect — only penalize defects the library introduced
                    base_ok, _base_err = validate(os.path.join(args.corpus, rec["path"]))
                    if not base_ok:
                        rt["opc_ok"] = None
                        rt["opc_error"] = f"baseline also invalid: {opc_err}"
                    else:
                        rt["opc_ok"] = False
                        rt["opc_error"] = opc_err
                        rt["ok"] = False
                else:
                    rt["opc_ok"] = True
                    rt["opc_error"] = None
            elif rt.get("ok"):
                rt["opc_ok"] = None
                rt["opc_error"] = "output file missing at validation time"
                rt["ok"] = False
            if out_path and os.path.exists(out_path):
                if rt["ok"]:
                    os.unlink(out_path)
                else:
                    shutil.move(out_path, os.path.join(failures_dir, os.path.basename(out_path)))
            rt.pop("out", None)
            out.write(json.dumps(rec) + "\n")
            n += 1
    print(f"{args.lib}: {n} records -> {final}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
