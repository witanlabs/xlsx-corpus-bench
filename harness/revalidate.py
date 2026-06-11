#!/usr/bin/env python3
"""Re-judge round-trip OPC failures after a validator fix.

Files that previously PASSED can only keep passing under a relaxed validator,
so only records with opc_ok=false need re-checking — and their outputs were
preserved under <results>/failures/<lib>/. For each such record:
  - re-validate the preserved output with the current validator
  - apply the baseline gate (original also invalid -> not the library's fault)
  - restore the runner-level reload verdict from the raw jsonl
Newly-passing outputs are deleted from the failures dir; the final jsonl is
rewritten in place.

Usage: revalidate.py --results-dir results --corpus corpus [--lib LIB ...]
"""
import argparse
import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from opc_validate import validate  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--results-dir", required=True)
    ap.add_argument("--corpus", required=True)
    ap.add_argument("--lib", action="append", default=None)
    args = ap.parse_args()

    results = os.path.join(ROOT, args.results_dir)
    corpus = os.path.join(ROOT, args.corpus)
    libs = args.lib or [
        os.path.basename(p).rsplit(".", 1)[0]
        for p in glob.glob(os.path.join(results, "*.jsonl"))
        if not p.endswith(".raw.jsonl") and "manifest" not in p
    ]

    for lib in libs:
        final = os.path.join(results, f"{lib}.jsonl")
        raw = os.path.join(results, f"{lib}.raw.jsonl")
        if not (os.path.exists(final) and os.path.exists(raw)):
            continue
        runner_rt = {}
        for line in open(raw):
            r = json.loads(line)
            runner_rt[r["sha256"]] = (r["roundtrip"].get("ok"), r["roundtrip"].get("error"))

        flipped = gated = 0
        records = [json.loads(l) for l in open(final)]
        for rec in records:
            rt = rec["roundtrip"]
            if rt.get("opc_ok") is not False:
                continue
            ext = os.path.splitext(rec["path"])[1].lower()
            preserved = os.path.join(results, "failures", lib, rec["sha256"] + ext)
            if not os.path.exists(preserved):
                continue
            run_ok, run_err = runner_rt.get(rec["sha256"], (False, "raw record missing"))
            opc_ok, opc_err = validate(preserved)
            if opc_ok:
                rt.update({"ok": bool(run_ok), "error": run_err, "opc_ok": True, "opc_error": None})
                flipped += 1
                os.unlink(preserved)
            else:
                base_ok, _ = validate(os.path.join(corpus, rec["path"]))
                if not base_ok:
                    rt.update({
                        "ok": bool(run_ok),
                        "error": run_err,
                        "opc_ok": None,
                        "opc_error": f"baseline also invalid: {opc_err}",
                    })
                    gated += 1
                    os.unlink(preserved)
        with open(final, "w") as f:
            for rec in records:
                f.write(json.dumps(rec) + "\n")
        print(f"{args.results_dir}/{lib}: {flipped} false positives cleared, {gated} baseline-gated")
    return 0


if __name__ == "__main__":
    sys.exit(main())
