"""
reconcile.py — deterministic pre-pass for K-2 Part II / Part X reconciliation.

Runs BEFORE the LLM. Reads the control grid (allocRec) and emits a small,
structured payload:

    {
      "confident":  [ ... rows safe to auto-apply (zero-judgment mechanical rules) ],
      "exceptions": [ ... rows that need LLM/preparer judgment ],
      "summary":    { ... }
    }

Design principle (hard-won): the reconciled workbook is MOSTLY the template's
existing column M/N defaults, with only a small curated set of overrides. An
engine that re-derives the full mapping from the Sch K-1 line number is WRONG —
it auto-"fixes" rows that were already correct and breaks tie-out.

So this engine is intentionally CONSERVATIVE:
  * It auto-applies ONLY rules that had zero counter-examples against the
    reconciled ground truth (100% precision). A wrong confident change breaks
    tie-out and is worse than flagging.
  * Everything else that looks inconsistent is FLAGGED as an exception for the
    LLM/preparer to resolve. The LLM sees only this short list, never the 21 MB
    workbook.

Part II <-> Part X are two views of the same income item and must form a valid
equivalence pair. The confident rules and the exception detector both key off
that relationship rather than off the Sch K-1 line number.

Authorized ranges only:
    allocRec      '(K.01) Sch K to K2 Control'!C10:GH1802   (edit cols M, N, H)
    k2Part2       'K2II-Part II'!A14:W3008                   (read-only)
    k2PartX       'K2X-Part X'!A8:U2496                      (read-only)
    part2Exclude  '(Z.00) Dynamic Dropdown List'!FN2:FO79    (read-only, static)
    partXExclude  '(Z.00) Dynamic Dropdown List'!FQ2:FR78    (read-only, static)
"""

from __future__ import annotations

import argparse
import json
import warnings
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as ci

warnings.simplefilter("ignore")

CONTROL_SHEET = "(K.01) Sch K to K2 Control"
DATA_FIRST_ROW = 18
DATA_LAST_ROW = 1802

# allocRec meta columns we read (1-based sheet columns).
COLS = {
    "SchKLineCode": ci("C"),
    "K3": ci("D"),
    "DetailCode": ci("E"),
    "Line": ci("F"),
    "Desc": ci("G"),
    "Active": ci("H"),
    "NaturalSign": ci("J"),
    "P2": ci("M"),   # Sch K2 - Part II Line  (edit target)
    "PX": ci("N"),   # Sch K2 - Part X Line   (edit target)
}

# Part II line labels.
P2_ORD_DIV = "7 - Ordinary dividends (exclude amount on line 8)"
P2_QUAL_DIV = "8 - Qualified dividends"
P2_LT_GAIN = "12 - Net long-term capital gain"

# Part X line labels.
PX_DIV = "7 - Dividends"
PX_LT_GAIN = "11 - Net long-term capital gain"

# Valid Part II <-> Part X equivalence pairs. A reconciled row's (P2, PX) must be
# one of these (or both blank). Anything else is an inconsistency to resolve.
VALID_PAIRS = {
    ("7 - Ordinary dividends (exclude amount on line 8)", "7 - Dividends"),
    ("8 - Qualified dividends", ""),
    ("12 - Net long-term capital gain", "11 - Net long-term capital gain"),
    ("20 - Other income (see instructions)", "2 - Other income"),
    ("15 - Net section 1231 gain", "14 - Net section 1231 gain"),
    ("14 - Unrecaptured section 1250 gain", "11 - Net long-term capital gain"),
    ("14 - Unrecaptured section 1250 gain", "14 - Net section 1231 gain"),
}

# Descriptions that flag an entity-specific / foreign / fund-level duplicate whose
# treatment depends on which row carries the amount -> always needs judgment.
VARIANT_TOKENS = ("climate", "tpg", "- foreign", "- us", "fund level")


def _norm(v) -> str:
    return "" if v is None else str(v).strip()


def confident_fix(line: str, desc: str, p2: str, px: str):
    """Return (new_p2, new_px, reason) for a row IF a zero-judgment rule fires,
    else None. These rules were validated to 100% precision (zero wrong) against
    the reconciled ground-truth workbook.
    """
    # R1 — Qualified dividends are reported on Part II line 8 ONLY; the same
    # amount must NOT also appear in Part X. If Part II already says line 8 and
    # Part X is populated, clear Part X.
    if p2 == P2_QUAL_DIV and px != "":
        return (p2, "", "qualified dividends: line 8 in Part II, cleared from Part X")

    # R2 — Long-term capital gain reported in Part X line 11 on Sch K-1 line 9a
    # must also be reflected in Part II line 12. If Part II is blank, set it.
    if line == "9a" and px == PX_LT_GAIN and p2 == "":
        return (P2_LT_GAIN, px, "Part II line 12 set to match Part X line 11 LT gain")

    # R3 — Ordinary dividends: a line 6a item (or any row whose Part X says line 7
    # Dividends) with Part II blank maps to Part II line 7 (Ordinary dividends).
    # True qualified dividends are handled by R4; "non-qualified" stays ordinary.
    is_qual = "qualified dividends" in desc.lower() and "non-qualified" not in desc.lower()
    if p2 == "" and px != "" and not is_qual and (px == PX_DIV or line == "6a"):
        return (P2_ORD_DIV, px, "Part II line 7 (ordinary dividends) set; Part X retained")

    # R4 — Qualified dividends carried only in Part X (line 7) with Part II blank:
    # reclass to Part II line 8 and clear Part X (qualified div is Part II only).
    if px == PX_DIV and p2 == "" and is_qual:
        return (P2_QUAL_DIV, "", "qualified dividends reclassed to Part II line 8, "
                                 "cleared from Part X")

    return None


def is_variant(desc: str) -> bool:
    d = desc.lower()
    return any(tok in d for tok in VARIANT_TOKENS)


def load_control(path: str):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[CONTROL_SHEET]
    rows = []
    for r, rv in enumerate(
        ws.iter_rows(min_row=DATA_FIRST_ROW, max_row=DATA_LAST_ROW, values_only=True),
        start=DATA_FIRST_ROW,
    ):
        rec = {k: (rv[idx - 1] if idx - 1 < len(rv) else None) for k, idx in COLS.items()}
        rec["row"] = r
        rows.append(rec)
    wb.close()
    return rows


def is_data_row(rec) -> bool:
    """A Schedule K-1 detail row: numeric K3 code and a Line value."""
    k3 = rec["K3"]
    if k3 is None:
        return False
    try:
        float(str(k3).strip())
    except (ValueError, TypeError):
        return False
    return _norm(rec["Line"]) != ""


def run(path: str):
    rows = load_control(path)
    confident, exceptions = [], []

    for rec in rows:
        if not is_data_row(rec):
            continue
        line = _norm(rec["Line"])
        desc = _norm(rec["Desc"])
        active = _norm(rec["Active"])
        p2 = _norm(rec["P2"])
        px = _norm(rec["PX"])

        base = {
            "row": rec["row"], "K3": rec["K3"], "line": line, "desc": desc,
            "active": active, "current_part2": p2, "current_partx": px,
        }

        # 1) Entity/foreign/fund-level variants always need judgment.
        if is_variant(desc):
            exceptions.append({**base, "proposed_part2": p2, "proposed_partx": px,
                               "why": "entity/foreign/fund-level variant — confirm which "
                                      "duplicate stays Active and clear the others"})
            continue

        # 2) Zero-judgment confident rule?
        fix = confident_fix(line, desc, p2, px)
        if fix is not None:
            new_p2, new_px, reason = fix
            if (new_p2, new_px) != (p2, px):
                confident.append({**base, "proposed_part2": new_p2,
                                  "proposed_partx": new_px, "reason": reason})
            continue

        # 3) Inconsistent (P2, PX) pair that no confident rule covers -> flag.
        if active.lower() != "no" and (p2 or px) and (p2, px) not in VALID_PAIRS:
            exceptions.append({**base, "proposed_part2": p2, "proposed_partx": px,
                               "why": "Part II / Part X lines are not a valid equivalence "
                                      "pair — preparer must choose the correct mapping"})

    payload = {
        "workbook": str(path),
        "confident": confident,
        "exceptions": exceptions,
        "summary": {
            "data_rows": sum(1 for r in rows if is_data_row(r)),
            "confident_changes": len(confident),
            "exceptions": len(exceptions),
        },
    }
    return payload


def main():
    ap = argparse.ArgumentParser(description="Deterministic K-2 Part II/X pre-pass")
    ap.add_argument("workbook", help="Path to the .xlsm workpaper")
    ap.add_argument("-o", "--output", default="reconcile_payload.json",
                    help="Output JSON payload for the LLM")
    args = ap.parse_args()

    payload = run(args.workbook)
    Path(args.output).write_text(json.dumps(payload, indent=2, default=str))

    s = payload["summary"]
    print(f"Wrote {args.output}")
    print(f"  data rows           : {s['data_rows']}")
    print(f"  confident changes   : {s['confident_changes']}")
    print(f"  exceptions for LLM  : {s['exceptions']}")


if __name__ == "__main__":
    main()
