#!/usr/bin/env python3
"""Stream fuse-cc-binaries.tar.gz and keep only OOXML spreadsheets.

FUSE stores files under sha1 names with no extensions, so format is detected
by content: a candidate must be a zip (PK magic) containing xl/workbook.xml.
Kept files are written as corpus-fuse/<tarname>.xlsx (or .xlsm if a
vbaProject part is present). Everything else (xls/OLE, csv, html mislabels)
is skipped and counted.
"""
import os
import sys
import tarfile
import zipfile
from collections import Counter
from io import BytesIO

SRC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fuse-archive", "fuse-cc-binaries.tar.gz")
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "corpus-fuse")


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    stats = Counter()
    tf = tarfile.open(SRC, "r|gz")  # streaming mode
    for member in tf:
        if not member.isfile():
            continue
        stats["files"] += 1
        f = tf.extractfile(member)
        head = f.read(4)
        if head != b"PK\x03\x04":
            stats["not-zip"] += 1
            continue
        data = head + f.read()
        try:
            zf = zipfile.ZipFile(BytesIO(data))
            names = set(zf.namelist())
            zf.close()
        except Exception:
            stats["broken-zip-skipped"] += 1
            continue
        if "xl/workbook.xml" not in names:
            stats["zip-not-xlsx"] += 1
            continue
        ext = ".xlsm" if "xl/vbaProject.bin" in names else ".xlsx"
        base = os.path.basename(member.name)
        with open(os.path.join(OUT, base + ext), "wb") as out:
            out.write(data)
        stats["kept"] += 1
        if stats["files"] % 10000 == 0:
            print(dict(stats), file=sys.stderr)
    tf.close()
    print("final:", dict(stats), file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
