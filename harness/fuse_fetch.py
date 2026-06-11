#!/usr/bin/env python3
"""Selectively extract xlsx members from the FUSE corpus zip on Zenodo
without downloading the whole 9.4 GB archive, using HTTP range requests:

  1. read the end-of-central-directory (+ zip64 records) from the file tail
  2. read the full central directory and list members
  3. download just the chosen members' byte ranges and inflate locally

Subcommands:
  list                 — parse the central directory, print summary stats
  fetch --out DIR      — download .xlsx/.xlsm members (optionally --limit N,
                         deterministically spread across the archive)
"""
import argparse
import os
import struct
import sys
import urllib.request
import zlib
from concurrent.futures import ThreadPoolExecutor

URL = "https://zenodo.org/records/581678/files/fuse.zip?download=1"
SIZE = 9392830118
CD_CACHE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "results", "fuse_cd.bin")


def fetch_range(start: int, end: int) -> bytes:
    req = urllib.request.Request(URL, headers={"Range": f"bytes={start}-{end}"})
    with urllib.request.urlopen(req, timeout=120) as resp:
        return resp.read()


def central_directory() -> bytes:
    if os.path.exists(CD_CACHE):
        return open(CD_CACHE, "rb").read()
    tail = fetch_range(SIZE - 65536, SIZE - 1)
    eocd = tail.rfind(b"PK\x05\x06")
    if eocd < 0:
        raise SystemExit("EOCD not found")
    cd_size, cd_offset = struct.unpack("<II", tail[eocd + 12 : eocd + 20])
    if cd_offset == 0xFFFFFFFF or cd_size == 0xFFFFFFFF:
        loc = tail.rfind(b"PK\x06\x07")  # zip64 EOCD locator
        if loc < 0:
            raise SystemExit("zip64 locator not found")
        (z64_eocd_off,) = struct.unpack("<Q", tail[loc + 8 : loc + 16])
        z64 = fetch_range(z64_eocd_off, z64_eocd_off + 55)
        if z64[:4] != b"PK\x06\x06":
            raise SystemExit("zip64 EOCD not found")
        cd_size, cd_offset = struct.unpack("<QQ", z64[40:56])
    print(f"central directory: {cd_size / 1e6:.1f} MB at offset {cd_offset}", file=sys.stderr)
    cd = b""
    chunk = 32 * 1024 * 1024
    for s in range(cd_offset, cd_offset + cd_size, chunk):
        cd += fetch_range(s, min(s + chunk, cd_offset + cd_size) - 1)
        print(f"  fetched {len(cd) / 1e6:.0f}/{cd_size / 1e6:.0f} MB", file=sys.stderr)
    os.makedirs(os.path.dirname(CD_CACHE), exist_ok=True)
    with open(CD_CACHE, "wb") as f:
        f.write(cd)
    return cd


def parse_members(cd: bytes):
    members, pos = [], 0
    while pos + 46 <= len(cd) and cd[pos : pos + 4] == b"PK\x01\x02":
        (method, csize, usize, nlen, elen, clen) = struct.unpack("<H8xIIHHH", cd[pos + 10 : pos + 34])
        (lho,) = struct.unpack("<I", cd[pos + 42 : pos + 46])
        name = cd[pos + 46 : pos + 46 + nlen].decode("utf-8", "replace")
        extra = cd[pos + 46 + nlen : pos + 46 + nlen + elen]
        # zip64 extra field overrides 0xFFFFFFFF fields, in order
        if 0xFFFFFFFF in (csize, usize, lho):
            e = 0
            while e + 4 <= len(extra):
                hid, hsize = struct.unpack("<HH", extra[e : e + 4])
                if hid == 1:
                    vals, v = extra[e + 4 : e + 4 + hsize], 0
                    if usize == 0xFFFFFFFF:
                        (usize,) = struct.unpack("<Q", vals[v : v + 8]); v += 8
                    if csize == 0xFFFFFFFF:
                        (csize,) = struct.unpack("<Q", vals[v : v + 8]); v += 8
                    if lho == 0xFFFFFFFF:
                        (lho,) = struct.unpack("<Q", vals[v : v + 8]); v += 8
                    break
                e += 4 + hsize
        members.append((name, method, csize, usize, lho))
        pos += 46 + nlen + elen + clen
    return members


def download_member(name, method, csize, lho, out_dir) -> str | None:
    out = os.path.join(out_dir, name.replace("/", "__"))
    if os.path.exists(out):
        return out
    try:
        header = fetch_range(lho, lho + 29)
        if header[:4] != b"PK\x03\x04":
            return None
        nlen, elen = struct.unpack("<HH", header[26:30])
        data_start = lho + 30 + nlen + elen
        raw = fetch_range(data_start, data_start + csize - 1)
        data = zlib.decompress(raw, -15) if method == 8 else raw
        with open(out, "wb") as f:
            f.write(data)
        return out
    except Exception as e:
        print(f"  failed {name}: {e}", file=sys.stderr)
        return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["list", "fetch"])
    ap.add_argument("--out", default=None)
    ap.add_argument("--limit", type=int, default=0, help="max files (0 = all)")
    ap.add_argument("--workers", type=int, default=12)
    args = ap.parse_args()

    members = parse_members(central_directory())
    from collections import Counter

    exts = Counter(os.path.splitext(m[0])[1].lower() for m in members)
    if args.cmd == "list":
        print(f"{len(members)} members")
        for ext, n in exts.most_common(15):
            size = sum(m[3] for m in members if os.path.splitext(m[0])[1].lower() == ext)
            print(f"  {ext or '(none)':12} {n:7}  {size / 1e9:.2f} GB uncompressed")
        return 0

    want = [m for m in members if os.path.splitext(m[0])[1].lower() in (".xlsx", ".xlsm")]
    want.sort(key=lambda m: m[0])
    if args.limit and len(want) > args.limit:
        step = len(want) / args.limit  # deterministic spread, no RNG
        want = [want[int(i * step)] for i in range(args.limit)]
    total = sum(m[2] for m in want)
    print(f"fetching {len(want)} members ({total / 1e9:.2f} GB compressed)", file=sys.stderr)
    os.makedirs(args.out, exist_ok=True)
    done = 0
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        for _ in pool.map(lambda m: download_member(m[0], m[1], m[2], m[4], args.out), want):
            done += 1
            if done % 500 == 0:
                print(f"  {done}/{len(want)}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
