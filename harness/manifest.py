#!/usr/bin/env python3
"""Build a deduplicated manifest of corpus workbooks.

Walks the corpus directory, hashes every .xlsx/.xlsm file with sha256,
dedupes by content hash, and writes results/manifest.jsonl with one record
per unique workbook. The manifest is the single source of truth for what
the benchmark runs over — every runner consumes it, so every library sees
the exact same file set.
"""
import argparse
import hashlib
import json
import os
import sys

EXTENSIONS = {".xlsx", ".xlsm"}


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", required=True, help="corpus root directory")
    ap.add_argument("--out", required=True, help="output manifest.jsonl path")
    args = ap.parse_args()

    seen: dict[str, dict] = {}
    total = 0
    for root, _dirs, files in os.walk(args.corpus):
        for name in sorted(files):
            ext = os.path.splitext(name)[1].lower()
            if ext not in EXTENSIONS:
                continue
            path = os.path.join(root, name)
            total += 1
            digest = sha256_file(path)
            if digest in seen:
                seen[digest]["dupes"] += 1
            else:
                seen[digest] = {
                    "sha256": digest,
                    "path": os.path.relpath(path, args.corpus),
                    "size": os.path.getsize(path),
                    "ext": ext,
                    "dupes": 0,
                }

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "w") as f:
        for rec in sorted(seen.values(), key=lambda r: r["path"]):
            f.write(json.dumps(rec) + "\n")

    print(f"{total} files scanned, {len(seen)} unique by sha256", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
