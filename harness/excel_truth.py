#!/usr/bin/env python3
"""Generate Excel ground truth: what does Excel itself compute for every
formula cell in the corpus?

For each manifest entry:
  1. stage a copy with fullCalcOnLoad="1" injected into <calcPr> — this
     forces Excel to fully recalculate the workbook on open (ECMA-376
     18.2.2), with no reliance on scripting calc commands
  2. drive Excel via AppleScript: open -> save as results/excel-truth/<sha>.xlsx
     -> close (display alerts off; VBA disabled beforehand via
     `defaults write com.microsoft.Excel VisualBasicMacroExecutionState
     DisabledWithoutWarnings`)
  3. the saved file's cached values ARE Excel's freshly computed results;
     harness/cached_values.py extracts them for comparison

Files are processed in chunks of one osascript call each, with a per-file
try block inside; chunk-level hangs (modal dialogs) are recovered by killing
Excel and retrying that chunk's missing files one by one. Failures are
recorded in results/excel-truth-skips.jsonl. Resumable: existing truth files
and recorded skips are not redone.

Staging and truth dirs live under the repo (not /tmp): Excel's sandbox
rejects writes outside user directories.
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile

CHUNK = 10
CHUNK_TIMEOUT_S = 420
SINGLE_TIMEOUT_S = 90


def inject_full_calc(src: str, dst: str) -> bool:
    """Copy src to dst with fullCalcOnLoad="1" in xl/workbook.xml."""
    zin = zipfile.ZipFile(src)
    try:
        items = zin.infolist()
        wb = zin.read("xl/workbook.xml")
    except KeyError:
        zin.close()
        shutil.copy(src, dst)
        return False
    m = re.search(rb"<(?:\w+:)?calcPr\b[^>]*", wb)
    if m:
        tag = m.group(0)
        if b"fullCalcOnLoad" in tag:
            new_tag = re.sub(rb'fullCalcOnLoad="[^"]*"', b'fullCalcOnLoad="1"', tag)
        elif tag.endswith(b"/"):  # self-closing: insert attr before the slash
            new_tag = tag[:-1].rstrip() + b' fullCalcOnLoad="1"/'
        else:
            new_tag = tag + b' fullCalcOnLoad="1"'
        wb_new = wb[: m.start()] + new_tag + wb[m.end() :]
    else:
        m = re.search(rb"</(\w+:)?sheets>", wb)
        if not m:
            zin.close()
            shutil.copy(src, dst)
            return False
        prefix = m.group(1) or b""
        elem = b"<" + prefix + b'calcPr fullCalcOnLoad="1"/>'
        wb_new = wb[: m.end()] + elem + wb[m.end() :]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in items:
            data = wb_new if info.filename == "xl/workbook.xml" else zin.read(info.filename)
            zout.writestr(info, data)
    zin.close()
    return True


def applescript_chunk(pairs: list[tuple[str, str, str]]) -> str:
    lines = ["with timeout of %d seconds" % (CHUNK_TIMEOUT_S - 30), 'tell application "Microsoft Excel"', "set display alerts to false"]
    for sha, staged, truth in pairs:
        lines += [
            "try",
            # "update links do not update links" suppresses the external-links
            # security prompt (display alerts off does not cover it) and keeps
            # link-dependent formulas on their cached link values — the same
            # no-network semantics every engine under test gets
            f'open workbook workbook file name (POSIX file "{staged}") update links do not update links',
            f'log "OPEN-OK {sha}"',
            "on error errMsg number errNum",
            f'log "OPEN-FAIL {sha} [" & errNum & "] " & errMsg',
            "end try",
            "try",
            f'save workbook as active workbook filename "{truth}" file format Excel XML file format with overwrite',
            f'log "SAVE-OK {sha}"',
            "on error errMsg number errNum",
            f'log "SAVE-FAIL {sha} [" & errNum & "] " & errMsg',
            "end try",
            "try",
            "close active workbook saving no",
            "end try",
        ]
    lines += ["end tell", "end timeout"]
    return "\n".join(lines)


PURGE_DIRS = [
    os.path.expanduser(
        "~/Library/Containers/com.microsoft.Excel/Data/Library/"
        "Application Support/Microsoft/Office/Office 16 AutoRecovery"
    ),
    os.path.expanduser(
        "~/Library/Containers/com.microsoft.Excel/Data/Library/"
        "Saved Application State/com.microsoft.Excel.savedState"
    ),
]


def purge_autorecovery() -> None:
    """A corpus file that crashes Excel gets resurrected by AutoRecovery /
    saved-state on the next launch, wedging it in the recovery pane — it
    answers pings but refuses opens. Purge so relaunches come up empty."""
    for d in PURGE_DIRS:
        if os.path.isdir(d):
            for name in os.listdir(d):
                try:
                    path = os.path.join(d, name)
                    os.unlink(path) if os.path.isfile(path) else shutil.rmtree(path)
                except OSError:
                    pass


def excel_process_exists() -> bool:
    return (
        subprocess.run(
            ["pgrep", "-f", "Microsoft Excel.app/Contents/MacOS"], capture_output=True
        ).returncode
        == 0
    )


def excel_responsive() -> bool:
    """Only call when the process exists: `tell application` AUTO-LAUNCHES a
    dead Excel, resurrecting un-purged crash-recovery state as a side effect.
    All launching must go through restart_excel (purge first)."""
    try:
        p = subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft Excel" to return name'],
            capture_output=True,
            timeout=20,
        )
        return p.returncode == 0
    except subprocess.TimeoutExpired:
        return False


def restart_excel() -> None:
    try:
        subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft Excel" to quit saving no'],
            capture_output=True,
            timeout=20,
        )
    except subprocess.TimeoutExpired:
        pass
    # the process is named "Microsoft" (…/MacOS/Microsoft), so -x never
    # matches — match the full command line, and SIGKILL: a wedged Excel
    # ignores SIGTERM and a half-dead one poisons every Apple Events
    # connection with error -609
    subprocess.run(["pkill", "-9", "-f", "Microsoft Excel.app"], capture_output=True)
    time.sleep(3)
    purge_autorecovery()
    # launch via Apple Events, NOT `open -a`: from a sandboxed caller,
    # open(1) yields an Excel that can open files but gets Parameter error
    # -50 on every save; an AE launch goes through launchd as a normal app
    # launch. Purge above makes the AE auto-launch safe (no recovery state).
    try:
        subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft Excel" to launch'],
            capture_output=True,
            timeout=60,
        )
    except subprocess.TimeoutExpired:
        pass
    time.sleep(8)


MINIMAL_XLSX_PARTS = {
    "[Content_Types].xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    "_rels/.rels": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    "xl/workbook.xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    "xl/_rels/workbook.xml.rels": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    "xl/worksheets/sheet1.xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
}

CANARY_DIR = None  # set in main() to the truth dir's parent


def excel_can_save() -> bool:
    """Liveness is not enough: an Excel launched by open(1) from a sandboxed
    context answers pings and opens files but fails every save with
    Parameter error -50. Prove a save works before trusting the instance."""
    canary_in = os.path.join(CANARY_DIR, "canary_in.xlsx")
    canary_out = os.path.join(CANARY_DIR, "canary_out.xlsx")
    with zipfile.ZipFile(canary_in, "w") as z:
        for name, data in MINIMAL_XLSX_PARTS.items():
            z.writestr(name, data)
    for p in (canary_out,):
        if os.path.exists(p):
            os.unlink(p)
    script = (
        'tell application "Microsoft Excel"\n'
        "set display alerts to false\n"
        f'open workbook workbook file name (POSIX file "{canary_in}")\n'
        f'save workbook as active workbook filename "{canary_out}" file format Excel XML file format with overwrite\n'
        "close active workbook saving no\n"
        "end tell"
    )
    try:
        subprocess.run(["osascript", "-e", script], capture_output=True, timeout=60)
    except subprocess.TimeoutExpired:
        return False
    ok = os.path.exists(canary_out)
    for p in (canary_in, canary_out):
        if os.path.exists(p):
            os.unlink(p)
    return ok


def ensure_excel(verify_save: bool = False) -> None:
    """Some wild corpus files crash Excel outright (e.g. during save); make
    sure a RESPONSIVE instance exists before talking to it, or every
    subsequent Apple Event fails with -609 connection-invalid. A dead
    process goes straight to restart (purge + launch) — pinging it would
    auto-launch with crash-recovery state intact. With verify_save, also
    prove the instance can actually write a file (see excel_can_save)."""
    if not excel_process_exists() or not excel_responsive():
        restart_excel()
    if verify_save:
        for _attempt in range(3):
            if excel_can_save():
                return
            restart_excel()
        raise SystemExit("Excel cannot save even after restarts; aborting instead of mass-skipping")


def collect_saves(pairs, truth_dir: str) -> None:
    """Move Excel's temp saves from the parent dir into the truth dir."""
    for sha, _staged, tmp_save in pairs:
        if os.path.exists(tmp_save):
            shutil.move(tmp_save, os.path.join(truth_dir, sha + ".xlsx"))


def run_chunk(pairs) -> None:
    script = applescript_chunk(pairs)
    ensure_excel()
    try:
        p = subprocess.run(["osascript", "-e", script], capture_output=True, timeout=CHUNK_TIMEOUT_S)
        for line in p.stderr.decode("utf-8", "replace").splitlines():
            if "FAIL" in line:
                print(f"  {line.strip()[:200]}", file=sys.stderr)
    except subprocess.TimeoutExpired:
        print("  chunk timeout", file=sys.stderr)
        restart_excel()


def main() -> int:
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", default=os.path.join(root, "corpus-sb"))
    ap.add_argument("--manifest", default=os.path.join(root, "results-sb/manifest.jsonl"))
    ap.add_argument("--truth-dir", default=os.path.join(root, "results-sb/excel-truth"))
    ap.add_argument("--skips", default=os.path.join(root, "results-sb/excel-truth-skips.jsonl"))
    args = ap.parse_args()

    staging = os.path.join(os.path.dirname(os.path.abspath(args.truth_dir)), "excel-staging")
    global CANARY_DIR
    CANARY_DIR = os.path.dirname(os.path.abspath(args.truth_dir))
    os.makedirs(args.truth_dir, exist_ok=True)
    os.makedirs(staging, exist_ok=True)

    skipped = set()
    if os.path.exists(args.skips):
        with open(args.skips) as f:
            for line in f:
                try:
                    skipped.add(json.loads(line)["sha256"])
                except Exception:
                    pass

    todo = []
    with open(args.manifest) as f:
        for line in f:
            rec = json.loads(line)
            sha = rec["sha256"]
            if sha in skipped or os.path.exists(os.path.join(args.truth_dir, sha + ".xlsx")):
                continue
            todo.append(rec)

    # start from a provably save-capable Excel: a stale instance launched by
    # open(1) from a sandboxed context passes liveness but fails every save
    restart_excel()
    ensure_excel(verify_save=True)

    skips_out = open(args.skips, "a")
    done = 0
    for i in range(0, len(todo), CHUNK):
        # preventive restart: long Excel sessions accumulate state and can
        # wedge in ways that fail fast instead of timing out
        if (i // CHUNK) % 20 == 19:
            restart_excel()
        chunk = todo[i : i + CHUNK]
        pairs = []
        for rec in chunk:
            staged = os.path.join(staging, rec["sha256"] + rec["ext"])
            # Excel saves to a temp name in the truth dir's PARENT: saving
            # directly into the truth subdir draws Parameter error -50 from
            # background contexts (the parent is proven writable by the
            # canary); python moves results into place afterwards
            tmp_save = os.path.join(CANARY_DIR, "save_" + rec["sha256"] + ".xlsx")
            try:
                inject_full_calc(os.path.join(args.corpus, rec["path"]), staged)
                pairs.append((rec["sha256"], staged, tmp_save))
            except Exception as e:
                skips_out.write(json.dumps({"sha256": rec["sha256"], "error": f"staging: {e}"[:300]}) + "\n")
        run_chunk(pairs)
        collect_saves(pairs, args.truth_dir)

        # zero successes in a whole chunk means Excel itself is wedged or
        # write-restricted, not ten bad files in a row — get a provably
        # save-capable instance before judging any file individually
        if pairs and not any(
            os.path.exists(os.path.join(args.truth_dir, sha + ".xlsx")) for sha, _g, _t in pairs
        ):
            restart_excel()
            ensure_excel(verify_save=True)

        # retry chunk leftovers individually; a second miss is a recorded skip
        for sha, staged, tmp_save in pairs:
            final = os.path.join(args.truth_dir, sha + ".xlsx")
            if not os.path.exists(final):
                ensure_excel()
                detail = ""
                try:
                    p = subprocess.run(
                        ["osascript", "-e", applescript_chunk([(sha, staged, tmp_save)])],
                        capture_output=True,
                        timeout=SINGLE_TIMEOUT_S,
                    )
                    detail = p.stderr.decode("utf-8", "replace").strip()[:300]
                except subprocess.TimeoutExpired:
                    detail = "osascript timeout (modal dialog?)"
                    restart_excel()
                if os.path.exists(tmp_save):
                    shutil.move(tmp_save, final)
                if not os.path.exists(final):
                    skips_out.write(json.dumps({"sha256": sha, "error": detail or "no output, no error"}) + "\n")
                    skips_out.flush()
            os.unlink(staged)
        done += len(chunk)
        if (i // CHUNK) % 5 == 0:
            print(f"excel-truth: {done}/{len(todo)}", file=sys.stderr)

    subprocess.run(["pkill", "-x", "Microsoft Excel"], capture_output=True)
    print(f"excel-truth complete: {done} attempted", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
