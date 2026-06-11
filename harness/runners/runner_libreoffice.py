#!/usr/bin/env python3
"""LibreOffice (headless) runner.

LibreOffice is an application, not a library, so it is driven through
`soffice --headless --convert-to xlsx`. A conversion is a full load of the
file into Calc's model plus a write of a new xlsx:

  load      — original converts without error, with recalc-on-load DISABLED
              (OOXMLRecalcMode=1) so the pass measures parsing only
  roundtrip — LibreOffice's own output converts again (recalc disabled);
              the orchestrator additionally OPC-validates outputs
  recalc    — original converts with recalc-on-load FORCED
              (OOXMLRecalcMode=0); the harness then extracts formula-cell
              cached values from the original and the recalculated output
              straight from sheet XML (harness/cached_values.py) and diffs
              them. LibreOffice 26's embedded Python has macOS launch
              constraints that kill external UNO clients, so the engine is
              exercised through load-time recalculation instead — same
              engine, no UNO.

Files are staged under their sha256 name to avoid basename collisions and
converted in batches to amortize soffice startup; files missing from a
batch's output (e.g. a crashed batch) are retried individually so one bad
file cannot fail its batchmates. Each batch gets a fresh user profile.
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from cached_values import compare, extract  # noqa: E402

SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
BATCH = 25
BATCH_TIMEOUT_S = 600
SINGLE_TIMEOUT_S = 120

PROFILE_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema">
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>{mode}</value></prop></item>
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>{mode}</value></prop></item>
</oor:items>
"""
RECALC_ALWAYS = 0
RECALC_NEVER = 1


def make_profile(base: str, name: str, mode: int) -> str:
    prof = os.path.join(base, name)
    os.makedirs(os.path.join(prof, "user"), exist_ok=True)
    with open(os.path.join(prof, "user", "registrymodifications.xcu"), "w") as f:
        f.write(PROFILE_XCU.format(mode=mode))
    return prof


def convert(files: list[str], out_dir: str, profile: str, timeout: int) -> None:
    os.makedirs(out_dir, exist_ok=True)
    cmd = [
        SOFFICE,
        "--headless",
        "--norestore",
        f"-env:UserInstallation=file://{profile}",
        "--convert-to",
        "xlsx",
        "--outdir",
        out_dir,
    ] + files
    try:
        subprocess.run(cmd, timeout=timeout, capture_output=True)
    except subprocess.TimeoutExpired:
        pass


def convert_with_retry(files: list[str], out_dir: str, work: str, tag: str, mode: int) -> None:
    """Batch-convert, then retry any file that produced no output one at a
    time with its own fresh profile."""
    if not files:
        return
    convert(files, out_dir, make_profile(work, f"prof_{tag}", mode), BATCH_TIMEOUT_S)
    for f in files:
        expected = os.path.join(out_dir, os.path.splitext(os.path.basename(f))[0] + ".xlsx")
        if not os.path.exists(expected):
            prof = make_profile(work, f"prof_{tag}_retry_{os.path.basename(f)}", mode)
            convert([f], out_dir, prof, SINGLE_TIMEOUT_S)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", required=True)
    ap.add_argument("--manifest", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--out-dir", required=True, help="dir for round-tripped files")
    args = ap.parse_args()

    done = set()
    if os.path.exists(args.out):
        with open(args.out) as f:
            for line in f:
                try:
                    done.add(json.loads(line)["sha256"])
                except Exception:
                    pass

    records = []
    with open(args.manifest) as f:
        for line in f:
            rec = json.loads(line)
            if rec["sha256"] not in done:
                records.append(rec)

    os.makedirs(args.out_dir, exist_ok=True)
    out = open(args.out, "a")
    with tempfile.TemporaryDirectory(prefix="lo_bench_") as work:
        stage = os.path.join(work, "stage")
        os.makedirs(stage)
        for i in range(0, len(records), BATCH):
            batch = records[i : i + BATCH]
            staged = []
            for rec in batch:
                dst = os.path.join(stage, rec["sha256"] + rec["ext"])
                shutil.copy(os.path.join(args.corpus, rec["path"]), dst)
                staged.append(dst)

            pass1 = os.path.join(work, f"pass1_{i}")  # load + roundtrip output
            pass2 = os.path.join(work, f"pass2_{i}")  # reload check
            pass3 = os.path.join(work, f"pass3_{i}")  # recalculated output
            t0 = time.monotonic()
            convert_with_retry(staged, pass1, work, f"{i}a", RECALC_NEVER)
            ms = int((time.monotonic() - t0) * 1000 / max(len(batch), 1))

            produced = [
                os.path.join(pass1, rec["sha256"] + ".xlsx")
                for rec in batch
                if os.path.exists(os.path.join(pass1, rec["sha256"] + ".xlsx"))
            ]
            convert_with_retry(produced, pass2, work, f"{i}b", RECALC_NEVER)

            t0 = time.monotonic()
            convert_with_retry(staged, pass3, work, f"{i}c", RECALC_ALWAYS)
            recalc_ms = int((time.monotonic() - t0) * 1000 / max(len(batch), 1))

            for rec, src in zip(batch, staged):
                loaded = os.path.exists(os.path.join(pass1, rec["sha256"] + ".xlsx"))
                reloaded = os.path.exists(os.path.join(pass2, rec["sha256"] + ".xlsx"))
                out_path = None
                if loaded:
                    out_path = os.path.join(args.out_dir, rec["sha256"] + ".xlsx")
                    shutil.copy(os.path.join(pass1, rec["sha256"] + ".xlsx"), out_path)

                recalc = {"supported": True, "ok": False, "error": None, "ms": recalc_ms}
                recalced = os.path.join(pass3, rec["sha256"] + ".xlsx")
                if os.path.exists(recalced):
                    try:
                        before = extract(src)
                        after = extract(recalced)
                        recalc.update(compare(before, after))
                        recalc["ok"] = True
                    except Exception as e:
                        recalc["error"] = f"{type(e).__name__}: {e}"[:500]
                else:
                    recalc["error"] = "recalc conversion produced no output"

                res = {
                    "sha256": rec["sha256"],
                    "path": rec["path"],
                    "lib": "libreoffice",
                    "load": {
                        "ok": loaded,
                        "ms": ms,
                        "error": None if loaded else "conversion produced no output",
                    },
                    "roundtrip": {
                        "ok": loaded and reloaded,
                        "ms": None,
                        "error": None if (loaded and reloaded) else ("no output to reload" if not loaded else "reload conversion failed"),
                        "out": out_path,
                    },
                    "recalc": recalc,
                }
                out.write(json.dumps(res) + "\n")
                os.unlink(src)
            out.flush()
            for d in (pass1, pass2, pass3):
                shutil.rmtree(d, ignore_errors=True)
            print(f"libreoffice: {min(i + BATCH, len(records))}/{len(records)}", file=sys.stderr)
    out.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
