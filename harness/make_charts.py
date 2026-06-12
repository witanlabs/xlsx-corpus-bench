#!/usr/bin/env python3
"""Render the results charts — with witan itself.

One scatter chart per corpus: x = round-trip survival, y = workbooks
recalculated 100% Excel-identical; one point per calculation engine
(openpyxl has no engine, hence no y value, hence no point — stated in the
chart title). Top-right is best. The workbook is authored through witan's
chart API and rasterized with witan's renderer, so the charts in the README
are themselves an artifact of the library under test.

Output: charts/charts.xlsx, charts/<corpus>.png

Usage: make_charts.py   (uses WITAN_XLSX_SERVE if set, else the public CLI)
"""
import json
import math
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_local = os.environ.get("WITAN_XLSX_SERVE")
WITAN = [_local] if _local else ["witan", "xlsx"]

ENGINES = ["witan", "libreoffice", "epplus", "closedxml"]
DISPLAY = {"witan": "witan", "libreoffice": "LibreOffice",
           "epplus": "EPPlus", "closedxml": "ClosedXML"}
MARKERS = {"witan": "circle", "libreoffice": "square",
           "epplus": "diamond", "closedxml": "triangle"}
CORPORA = [
    ("results-fuse", "FUSE", "FUSE — 10,544 wild-web workbooks"),
    ("results-sb", "SpreadsheetBench", "SpreadsheetBench — 5,426 forum workbooks"),
]


def aggregate(results_dir: str) -> dict:
    out = {}
    for lib in ENGINES:
        rows = [json.loads(l) for l in open(os.path.join(ROOT, results_dir, f"{lib}.jsonl"))]
        truth_dir = os.path.join(ROOT, results_dir, "excel-truth")
        rows = [r for r in rows if os.path.exists(os.path.join(truth_dir, r["sha256"] + ".xlsx"))]
        rt = sum(1 for r in rows if r["roundtrip"].get("ok")) / len(rows)
        tc = [json.loads(l) for l in open(os.path.join(ROOT, results_dir, "truth-compare", f"{lib}.jsonl"))]
        ok = [r for r in tc if r.get("ok") and r.get("formula_cells", 0) > 0]
        clean = sum(1 for r in ok if r["mismatches"] == 0) / len(ok)
        out[lib] = {"rt": rt, "clean": clean}
    return out


def main() -> int:
    charts_dir = os.path.join(ROOT, "charts")
    os.makedirs(charts_dir, exist_ok=True)
    wb_path = os.path.join(charts_dir, "charts.xlsx")
    if os.path.exists(wb_path):
        os.unlink(wb_path)

    data = {rd: aggregate(rd) for rd, _s, _t in CORPORA}

    script = []
    for rd, sheet, title in CORPORA:
        d = data[rd]
        xs = [d[l]["rt"] for l in ENGINES]
        ys = [d[l]["clean"] for l in ENGINES]
        ymin = max(0.0, math.floor((min(ys) - 0.02) * 20) / 20)
        script.append(f'await xlsx.addSheet(wb, {json.dumps(sheet)})')
        cells = [
            {"address": f"{sheet}!A1", "value": "library"},
            {"address": f"{sheet}!B1", "value": "round-trip survival"},
            {"address": f"{sheet}!C1", "value": "workbooks 100% Excel-identical"},
        ]
        for i, lib in enumerate(ENGINES):
            r = 2 + i
            cells += [
                {"address": f"{sheet}!A{r}", "value": DISPLAY[lib]},
                {"address": f"{sheet}!B{r}", "value": d[lib]["rt"]},
                {"address": f"{sheet}!C{r}", "value": d[lib]["clean"]},
            ]
        script.append(f"await xlsx.setCells(wb, {json.dumps(cells)})")
        series = []
        for i, lib in enumerate(ENGINES):
            r = 2 + i
            series.append({
                "name": {"ref": f"{sheet}!A{r}"},
                "xValues": f"{sheet}!B{r}:B{r}",
                "yValues": f"{sheet}!C{r}:C{r}",
                "marker": {"style": MARKERS[lib], "size": 11},
            })
        xmin = math.floor((min(xs) - 0.005) * 100) / 100
        chart = {
            "name": f"{sheet} results",
            "position": {"from": {"cell": "E2"}, "to": {"cell": "N22"}},
            "groups": [{"type": "scatter", "scatterStyle": "marker", "series": series}],
            "title": {"text": f"{title}\nxlsx engines — top-right is best"},
            "roundedCorners": False,
            "legend": {"position": "bottom"},
            "axes": {
                "category": {
                    "title": {"text": "round-trip survival (open→save→reopen)"},
                    "numberFormat": "0%", "numberFormatLinked": False,
                    "min": xmin, "max": 1.0, "majorUnit": 0.01, "majorGridlines": True,
                },
                "value": {
                    "title": {"text": "workbooks recalculated 100% Excel-identical"},
                    "numberFormat": "0%", "numberFormatLinked": False,
                    "min": ymin, "max": 1.0, "majorUnit": 0.05, "majorGridlines": True,
                },
            },
        }
        script.append(f"await xlsx.addChart(wb, {json.dumps(sheet)}, {json.dumps(chart)})")
    script.append('return "ok"')

    cmd = WITAN + ["exec", wb_path, "--create", "--save", "--stdin", "--json"]
    p = subprocess.run(cmd, input="\n".join(script).encode(), capture_output=True, timeout=300)
    resp = json.loads(p.stdout)
    if not resp.get("ok"):
        sys.exit(f"chart authoring failed: {resp.get('error')}")

    for rd, sheet, _t in CORPORA:
        png = os.path.join(charts_dir, f"{rd.replace('results-', '')}.png")
        subprocess.run(
            WITAN + ["render", wb_path, "-r", f"{sheet}!E2:M21", "-o", png],
            check=True, timeout=300,
        )
        print(f"rendered {png}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
