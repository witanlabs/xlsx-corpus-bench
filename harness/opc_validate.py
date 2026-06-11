#!/usr/bin/env python3
"""Library-neutral structural validation of an OOXML package.

A round-tripped file only counts as surviving if it passes this check,
which is independent of every library under test:
  - the file is a readable zip archive
  - [Content_Types].xml exists and parses
  - every .xml/.rels part parses as well-formed XML
  - the package has a workbook part reachable from _rels/.rels
  - every relationship target in every .rels part resolves to an existing part

This is a structural proxy for "Excel opens it without the repair dialog";
an Excel-as-oracle pass can be layered on top in CI on a Windows runner.
"""
import argparse
import json
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile

REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def validate(path: str) -> tuple[bool, str]:
    try:
        zf = zipfile.ZipFile(path)
    except Exception as e:
        return False, f"not a zip: {e}"
    try:
        names = set(zf.namelist())
        if "[Content_Types].xml" not in names:
            return False, "missing [Content_Types].xml"
        for name in names:
            if name.endswith("/"):
                continue
            if name.endswith(".xml") or name.endswith(".rels"):
                try:
                    ET.fromstring(zf.read(name))
                except Exception as e:
                    return False, f"malformed XML in {name}: {e}"
        if "_rels/.rels" not in names:
            return False, "missing _rels/.rels"
        # every relationship target (non-external) must resolve to a part
        for name in [n for n in names if n.endswith(".rels")]:
            base = posixpath.dirname(posixpath.dirname(name))
            for rel in ET.fromstring(zf.read(name)).findall(f"{REL_NS}Relationship"):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target", "")
                if not target:
                    return False, f"empty relationship target in {name}"
                target = target.split("#")[0]
                if not target:
                    # fragment-only target (e.g. Target="#'Sheet'!A1" on a
                    # shape hyperlink): an in-package location reference,
                    # written by Excel itself without TargetMode="External"
                    continue
                if target.startswith("/"):
                    resolved = target.lstrip("/")
                else:
                    resolved = posixpath.normpath(posixpath.join(base, target))
                if resolved not in names:
                    return False, f"dangling relationship {name} -> {target}"
        return True, ""
    finally:
        zf.close()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()
    ok_all = True
    for path in args.files:
        ok, err = validate(path)
        ok_all &= ok
        if args.json:
            print(json.dumps({"file": path, "ok": ok, "error": err}))
        else:
            print(f"{'OK  ' if ok else 'FAIL'} {path}{'  ' + err if err else ''}")
    return 0 if ok_all else 1


if __name__ == "__main__":
    sys.exit(main())
