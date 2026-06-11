#!/usr/bin/env python3
"""Excel-reopen oracle: does Microsoft Excel open each library's saved
output cleanly?

With `display alerts to false`, Excel REFUSES to open a file it would need
to repair (error -50) instead of showing the repair dialog — verified
empirically with a deliberately corrupted workbook. So the metric is
binary and dialog-free:

  opened  -> Excel accepts the library's output as-is
  openfail -> Excel cannot open it without user-mediated repair
  (no output) -> the library failed to save this workbook at all; counted
                 as a failure, same no-credit-for-skipped-work rule as
                 every other metric

Reads outputs from <results-dir>/rt-outputs/<lib>/ (harness/regen_outputs.py)
and writes one JSON line per manifest entry to
<results-dir>/excel-repair/<lib>.jsonl. Resumable. Excel lifecycle handling
(launch via Apple Events, AutoRecovery purge, responsiveness pings, chunk
retries) reuses the same machinery as truth generation.

Usage: excel_repair_check.py --lib LIB --results-dir results [--corpus corpus]
"""
import argparse
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import excel_truth  # noqa: E402  (restart_excel / ensure_excel / purge)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CHUNK = 10
CHUNK_TIMEOUT_S = 420
SINGLE_TIMEOUT_S = 90

PIN_PATH = None  # set in main()


def pin_manual_calc() -> bool:
    """Excel adopts its session calculation mode from the FIRST workbook
    opened. Open-and-close a tiny calcMode="manual" workbook so every
    subsequent open in this session skips recalculation (fullCalcOnLoad
    fires during open, so setting the mode afterwards is too late)."""
    import zipfile

    parts = dict(excel_truth.MINIMAL_XLSX_PARTS)
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
        b"</workbook>", b'<calcPr calcMode="manual"/></workbook>'
    )
    with zipfile.ZipFile(PIN_PATH, "w") as z:
        for name, data in parts.items():
            z.writestr(name, data)
    script = (
        'tell application "Microsoft Excel"\n'
        "set display alerts to false\n"
        f'open workbook workbook file name (POSIX file "{PIN_PATH}")\n'
        "set calculation to calculation manual\n"
        # verify BEFORE closing: with no workbook open the calculation
        # property does not report the session mode (though the mode itself
        # persists behaviorally for subsequent opens)
        'set pinOK to (calculation is calculation manual)\n'
        "close active workbook saving no\n"
        'if pinOK then return "PIN-OK"\n'
        'return "PIN-MISSING"\n'
        "end tell"
    )
    try:
        p = subprocess.run(["osascript", "-e", script], capture_output=True, timeout=60)
        return b"PIN-OK" in p.stdout
    except subprocess.TimeoutExpired:
        return False


def restart_excel_pinned() -> None:
    """The pin MUST verifiably stick: an unpinned session recalculates
    dynamic-array workbooks on open (~120s each for some 9 KB files)."""
    for _attempt in range(3):
        excel_truth.restart_excel()
        if pin_manual_calc():
            return
    raise SystemExit("cannot pin Excel to manual calculation; aborting")


def open_script(pairs) -> str:
    lines = [
        f"with timeout of {CHUNK_TIMEOUT_S - 30} seconds",
        'tell application "Microsoft Excel"',
        "set display alerts to false",
        # manual calculation: we measure open-without-repair, not values, and
        # openpyxl outputs set fullCalcOnLoad (its by-design behavior) which
        # otherwise makes Excel recompute entire workbooks per open. Setting
        # the property can fail with no workbook open — retry after each open.
        "try",
        "set calculation to calculation manual",
        "end try",
    ]
    for sha, path in pairs:
        lines += [
            "try",
            f'open workbook workbook file name (POSIX file "{path}") update links do not update links',
            "try",
            "set calculation to calculation manual",
            "end try",
            "close active workbook saving no",
            f'log "OPENED {sha}"',
            "on error m number n",
            f'log "OPENFAIL {sha} [" & n & "] " & m',
            "try",
            "close active workbook saving no",
            "end try",
            "end try",
        ]
    lines += ["end tell", "end timeout"]
    return "\n".join(lines)


def run_pairs(pairs, timeout) -> dict[str, str]:
    """Run a chunk; return sha -> 'opened' | 'openfail: detail'."""
    out: dict[str, str] = {}
    try:
        p = subprocess.run(
            ["osascript", "-e", open_script(pairs)], capture_output=True, timeout=timeout
        )
        for line in p.stderr.decode("utf-8", "replace").splitlines():
            line = line.strip()
            if line.startswith("OPENED "):
                out[line.split()[1]] = "opened"
            elif line.startswith("OPENFAIL "):
                sha = line.split()[1]
                out[sha] = "openfail: " + line.split(" ", 2)[2][:200]
    except subprocess.TimeoutExpired:
        restart_excel_pinned()
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--lib", required=True)
    ap.add_argument("--results-dir", required=True)
    args = ap.parse_args()

    results = os.path.join(ROOT, args.results_dir)
    out_dir = os.path.join(results, "rt-outputs", args.lib)
    excel_truth.CANARY_DIR = results  # save canary location for ensure_excel
    global PIN_PATH
    PIN_PATH = os.path.join(results, "calcpin.xlsx")
    repair_dir = os.path.join(results, "excel-repair")
    os.makedirs(repair_dir, exist_ok=True)
    out_path = os.path.join(repair_dir, f"{args.lib}.jsonl")

    done = set()
    if os.path.exists(out_path):
        with open(out_path) as f:
            for line in f:
                try:
                    done.add(json.loads(line)["sha256"])
                except Exception:
                    pass

    truth_dir = os.path.join(results, "excel-truth")
    todo = []
    for line in open(os.path.join(results, "manifest.jsonl")):
        rec = json.loads(line)
        if rec["sha256"] in done:
            continue
        # universe rule: skip files Excel can't process in original form
        if os.path.isdir(truth_dir) and not os.path.exists(
            os.path.join(truth_dir, rec["sha256"] + ".xlsx")
        ):
            continue
        path = os.path.join(out_dir, rec["sha256"] + rec["ext"])
        if not os.path.exists(path):
            path = os.path.join(out_dir, rec["sha256"] + ".xlsx")  # LO emits .xlsx
        todo.append((rec, path if os.path.exists(path) else None))

    print(f"{args.lib}: {len(todo)} outputs to check", file=sys.stderr)
    restart_excel_pinned()

    out = open(out_path, "a")
    n_done = 0
    for i in range(0, len(todo), CHUNK):
        if (i // CHUNK) % 50 == 49:
            restart_excel_pinned()  # preventive, long sessions decay
        chunk = todo[i : i + CHUNK]
        for rec, path in [(r, p) for r, p in chunk if p is None]:
            out.write(json.dumps({"sha256": rec["sha256"], "path": rec["path"], "result": "no-output"}) + "\n")
        pairs = [(rec["sha256"], os.path.abspath(path)) for rec, path in chunk if path]
        if pairs:
            excel_truth.ensure_excel()
            results_map = run_pairs(pairs, CHUNK_TIMEOUT_S)
            # whole chunk silent => Excel wedged, restart and retry singles
            if not results_map:
                restart_excel_pinned()
            for sha, path in pairs:
                verdict = results_map.get(sha)
                if verdict is None:
                    excel_truth.ensure_excel()
                    single = run_pairs([(sha, path)], SINGLE_TIMEOUT_S)
                    verdict = single.get(sha, "openfail: hang/timeout (modal or crash)")
                    if sha not in single:
                        restart_excel_pinned()
                rec = next(r for r, p in chunk if r["sha256"] == sha)
                out.write(json.dumps({"sha256": sha, "path": rec["path"], "result": verdict}) + "\n")
            out.flush()
        n_done += len(chunk)
        if (i // CHUNK) % 25 == 0:
            print(f"  {args.lib}: {n_done}/{len(todo)}", file=sys.stderr)
    out.close()
    subprocess.run(["pkill", "-9", "-f", "Microsoft Excel.app"], capture_output=True)
    print(f"{args.lib}: repair check complete", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
