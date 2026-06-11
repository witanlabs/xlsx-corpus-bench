#!/usr/bin/env python3
"""Extract cached formula-cell values straight from xlsx sheet XML.

This gives the harness an engine-neutral view of what Excel (or whichever
application last saved the file) stored as each formula's result: no library
under test is involved in producing the ground truth.

extract(path) returns {"Sheet!A1": (kind, value), ...} for every cell that
has an <f> element (including shared/array formula members). kind is one of
"n" (number, value float), "s" (string, value str), "b" (bool, value float
0/1), "e" (error, value str like "#DIV/0!"). Cells with no cached <v> are
returned with value None.
"""
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _text(elem) -> str:
    return "".join(elem.itertext())


def extract(path: str, with_formula: bool = False):
    """Without with_formula: {"Sheet!A1": (kind, value)}.
    With it: {"Sheet!A1": (kind, value, formula_text)} — shared-formula
    members are resolved to the master's text via their si index."""
    zf = zipfile.ZipFile(path)
    names = set(zf.namelist())

    # workbook sheet name -> rId
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rid_attr = next(
        (k for el in wb.iter() if _local(el.tag) == "sheet" for k in el.attrib if k.endswith("}id")),
        None,
    )
    sheets = []  # (name, rId)
    for el in wb.iter():
        if _local(el.tag) == "sheet":
            sheets.append((el.get("name"), el.get(rid_attr) if rid_attr else None))

    # rId -> part path
    rels = {}
    if "xl/_rels/workbook.xml.rels" in names:
        for rel in ET.fromstring(zf.read("xl/_rels/workbook.xml.rels")):
            target = rel.get("Target", "")
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = posixpath.normpath(posixpath.join("xl", target))
            rels[rel.get("Id")] = resolved

    shared = []
    if "xl/sharedStrings.xml" in names:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
        shared = [_text(si) for si in root if _local(si.tag) == "si"]

    out = {}
    for sheet_name, rid in sheets:
        part = rels.get(rid)
        if not part or part not in names:
            continue
        shared_formulas: dict[str, str] = {}  # si index -> master formula text
        formulas: dict[str, str] = {}
        for _event, c in ET.iterparse(zf.open(part)):
            if _local(c.tag) != "c":
                continue
            has_f, v_text, is_text, f_text = False, None, None, ""
            for child in c:
                tag = _local(child.tag)
                if tag == "f":
                    has_f = True
                    f_text = child.text or ""
                    si = child.get("si")
                    if si is not None:
                        if f_text:
                            shared_formulas[si] = f_text
                        else:
                            f_text = shared_formulas.get(si, "")
                elif tag == "v":
                    v_text = child.text or ""
                elif tag == "is":
                    is_text = _text(child)
            if not has_f:
                c.clear()
                continue
            addr = f"{sheet_name}!{c.get('r')}"
            if with_formula:
                formulas[addr] = f_text
            t = c.get("t", "n")
            if v_text is None and is_text is None:
                out[addr] = (t if t in ("e", "b") else ("s" if t in ("s", "str", "inlineStr") else "n"), None)
            elif t == "n":
                try:
                    out[addr] = ("n", float(v_text))
                except ValueError:
                    out[addr] = ("s", v_text)
            elif t == "b":
                out[addr] = ("b", float(v_text in ("1", "true", "TRUE")))
            elif t == "e":
                out[addr] = ("e", v_text)
            elif t == "s":
                try:
                    out[addr] = ("s", shared[int(v_text)])
                except (ValueError, IndexError):
                    out[addr] = ("s", v_text)
            else:  # "str", "inlineStr", "d"
                out[addr] = ("s", is_text if is_text is not None else v_text)
            c.clear()
        if with_formula:
            for addr, f_text in formulas.items():
                if addr in out:
                    out[addr] = (*out[addr], f_text)
    zf.close()
    return out


VOLATILE = re.compile(r"\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY)\s*\(", re.I)
# external-workbook references ([1]Sheet!A1 / '[1]Name'!A1). These STAY in
# the comparison: xlsx caches external values (xl/externalLinks/) exactly so
# formulas evaluate offline, and the truth corpus reflects Excel's defined
# closed-workbook semantics (including the documented *IF/*IFS limitation of
# erroring with #VALUE! against closed sources). Counted per file so results
# can be stratified if a group is disputed.
EXTERNAL_REF = re.compile(r"\[\d+\]")


def compare_to_truth(truth: dict, engine: dict) -> dict:
    """Judge an engine's recalculated file against Excel's recalculated file.

    Denominator: truth formula cells where Excel produced a value and the
    formula is not volatile (NOW/TODAY/RAND* can never match across runs).
    A cell Excel computed but the engine left valueless or dropped counts as
    missing — engines don't get credit for skipping work."""
    comparable = mismatches = missing = errors = volatile_skipped = 0
    external_ref = external_mismatches = 0
    for addr, t in truth.items():
        kind_t, val_t, formula_t = t[0], t[1], (t[2] if len(t) > 2 else "")
        if val_t is None:
            continue
        if formula_t and VOLATILE.search(formula_t):
            volatile_skipped += 1
            continue
        is_external = bool(formula_t and EXTERNAL_REF.search(formula_t))
        if is_external:
            external_ref += 1
        e = engine.get(addr)
        if e is None or e[1] is None:
            missing += 1
            if is_external:
                external_mismatches += 1
            continue
        comparable += 1
        if e[0] == "e":
            errors += 1
        if not values_match((kind_t, val_t), (e[0], e[1])):
            mismatches += 1
            if is_external:
                external_mismatches += 1
    return {
        "formula_cells": comparable + missing,
        "mismatches": mismatches + missing,
        "missing": missing,
        "errors": errors,
        "volatile_skipped": volatile_skipped,
        "external_ref": external_ref,
        "external_mismatches": external_mismatches,
    }


def values_match(a: tuple[str, object], b: tuple[str, object]) -> bool:
    """Shared comparator: numbers/bools with 1e-9 relative tolerance,
    strings exact, errors exact by code."""
    (ka, va), (kb, vb) = a, b
    num_a, num_b = ka in ("n", "b"), kb in ("n", "b")
    if num_a and num_b:
        tol = max(1e-9, abs(va) * 1e-9)
        return abs(va - vb) <= tol
    if ka != kb:
        return False
    return va == vb


def compare(before: dict, after: dict) -> dict:
    """Compare two extractions. Only addresses present in both with a cached
    value on both sides are comparable; the rest are reported, not judged.

    An empty-string cache on the ground-truth side is also skipped: corpus
    files written by generator tools cache "" for formulas they never
    evaluated, which is indistinguishable from a genuine empty-string result,
    so judging an engine against it would count staleness as failure."""
    comparable = mismatches = errors = 0
    for addr, a in before.items():
        b = after.get(addr)
        if b is None or a[1] is None or b[1] is None:
            continue
        if a == ("s", ""):
            continue
        comparable += 1
        if b[0] == "e":
            errors += 1
        if not values_match(a, b):
            mismatches += 1
    return {
        "formula_cells": comparable,
        "mismatches": mismatches,
        "errors": errors,
        "uncomparable": len(before) - comparable,
    }
