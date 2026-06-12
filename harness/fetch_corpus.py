#!/usr/bin/env python3
"""Fetch a corpus end to end: download, extract, build the sha256 manifest.

  python3 harness/fetch_corpus.py spreadsheetbench
      ~100 MB download -> corpus-sb/ + results-sb/manifest.jsonl

  python3 harness/fetch_corpus.py fuse
      9.4 GB download (resumable), needs ~25 GB free and 7zz
      (`brew install sevenzip`) -> corpus-fuse/ + results-fuse/manifest.jsonl

Both are idempotent: finished stages are skipped on re-run.
"""
import argparse
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SB_URL = (
    "https://raw.githubusercontent.com/RUCKBReasoning/SpreadsheetBench/"
    "main/data/spreadsheetbench_912_v0.1.tar.gz"
)
FUSE_URL = "https://zenodo.org/records/581678/files/fuse.zip?download=1"


def run(cmd, **kw):
    print("+", " ".join(cmd), file=sys.stderr)
    subprocess.run(cmd, check=True, **kw)


def manifest(corpus_dir: str, results_dir: str) -> None:
    os.makedirs(results_dir, exist_ok=True)
    run([
        sys.executable, os.path.join(ROOT, "harness/manifest.py"),
        "--corpus", corpus_dir, "--out", os.path.join(results_dir, "manifest.jsonl"),
    ])


def fetch_spreadsheetbench() -> None:
    corpus = os.path.join(ROOT, "corpus-sb")
    if not os.path.isdir(os.path.join(corpus, "all_data_912_v0.1")):
        os.makedirs(corpus, exist_ok=True)
        tarball = os.path.join(corpus, "sb.tar.gz")
        run(["curl", "-fSL", "-C", "-", "-o", tarball, SB_URL])
        run(["tar", "-xzf", tarball, "-C", corpus])
        os.unlink(tarball)
    manifest(corpus, os.path.join(ROOT, "results-sb"))


def fetch_fuse() -> None:
    if not shutil.which("7zz"):
        sys.exit("7zz not found — install with `brew install sevenzip`")
    free_gb = shutil.disk_usage(ROOT).free / 1e9
    archive = os.path.join(ROOT, "fuse-archive")
    corpus = os.path.join(ROOT, "corpus-fuse")
    os.makedirs(archive, exist_ok=True)

    tarball = os.path.join(archive, "fuse-cc-binaries.tar.gz")
    if not os.path.exists(tarball):
        if free_gb < 25:
            sys.exit(f"only {free_gb:.0f} GB free; FUSE needs ~25 GB during extraction")
        zip_path = os.path.join(archive, "fuse.zip")
        if not os.path.exists(os.path.join(archive, "FUSE.7z.001")):
            run(["curl", "-fSL", "-C", "-", "-o", zip_path, FUSE_URL])
            run(["unzip", "-o", "-q", zip_path, "-x", "README.txt", "-d", archive])
            os.unlink(zip_path)
        # the zip contains a 140-part solid 7z; the spreadsheets are one
        # tar.gz member inside it
        run(["7zz", "e", os.path.join(archive, "FUSE.7z.001"),
             "fuse-cc-binaries.tar.gz", f"-o{archive}", "-y"])
        for name in os.listdir(archive):
            if name.startswith("FUSE.7z."):
                os.unlink(os.path.join(archive, name))

    if not os.path.isdir(corpus) or not os.listdir(corpus):
        run([sys.executable, os.path.join(ROOT, "harness/fuse_extract.py")])
    manifest(corpus, os.path.join(ROOT, "results-fuse"))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("corpus", choices=["spreadsheetbench", "fuse"])
    args = ap.parse_args()
    if args.corpus == "spreadsheetbench":
        fetch_spreadsheetbench()
    else:
        fetch_fuse()
    print("done", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
