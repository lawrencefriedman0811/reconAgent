---
name: k2-partx-recon
description: >
  Reconcile Schedule K to K-2 Part X in a federal partnership tax workpaper by
  assigning (or pruning) the correct K-2 Part X line for each active Schedule K-1 detail
  line in the (K.01) Sch K to K2 Control sheet, then validating that Part X ties out.
  WHEN: "reconcile Part X", "reconcile Part 10", "reconcile K-2 Part X", "Sch K to K2
  Part X", "assign Part X lines", "K2 Part X mapping", "fix Part X reconciliation",
  "Part X doesn't tie out", "map Schedule K to K-2 Part X". DO NOT USE FOR: Part II (use
  k2-part2-recon), or editing K-2 source amounts / exclude lists.
location: user
---

# K-2 Part X Reconciliation

Assigns (or clears) the correct **K-2 Part X line** for every active Schedule K-1 detail
row in a partnership workpaper, so that Schedule K gross income ties to K-2 Part X.
Reverse-engineered from a known unreconciled→reconciled example pair. Note: "Part 10" =
Part X (Roman numeral).

## When to use
- A workpaper's K-2 **Part X** is unreconciled (status cell shows the red ERROR) or does
  not tie out, and the Part X line assignments need to be set/corrected.
- Triggered by a SharePoint button → agent pipeline, or an explicit user request.

Use the sibling skill **`k2-part2-recon`** for Part II (column M). The two skills share
the `Active?` column — see "Coordination with Part II" below.

---

## Inputs

| Input | Description | Default |
|---|---|---|
| `workbook` | Path to the `.xlsm` workpaper | required |
| `entity` / `period` | Identify the workpaper instance (for logging/write-back) | optional |

### Authoritative ranges (read these only)
| Logical name | Location | Purpose |
|---|---|---|
| `allocRec`     | `'(K.01) Sch K to K2 Control'!C10:GH1802` | control grid — **edit here** |
| `k2PartX`      | `'K2X-Part X'!A8:U2496` | Part X source amounts (read-only) |
| `partXExclude` | `'(Z.00) Dynamic Dropdown List'!FQ2:FR78` | K3 codes excluded from tie-out (read-only, static) |

> Do **not** read or edit other ranges. `partXExclude` and `k2PartX` are never modified
> by this skill (confirmed byte-identical across the reconciled example).

### `allocRec` layout
- Header band: rows 10–16 (per-activity metadata + formulas; do not edit).
- Column headers: **row 17**. Data rows: **18 → 1802**.
- Meta columns (sheet → meaning):
  - `C` Schedule K Line Item Code · `D` **K-3 Code** · `E` Detail code · `F` **Line**
    (Sch K-1 line, e.g. `6a`, `6b`, `9a`) · `G` **Description** · `H` **Active?** ·
    `J` Natural Sign · `M` Part II Line · **`N` Sch K2 - Part X Line (EDIT TARGET)** ·
    `O` Total.
- **Validation cell:** `allocRec` N15 reads `"ERROR: Remove K-2 Selection (Red)"` when
  unreconciled and `"OK"` when reconciled.

---

## How it runs (script-first → LLM exceptions → re-validate)

Identical loop to the Part II skill. The deterministic script runs *before* the LLM so
the model only reasons over a short exceptions list, never the 21 MB workbook.

```
reconcile.py (deterministic)  ──►  exceptions payload  ──►  LLM judgment  ──►  re-validate
   auto-applies safe fixes          (small JSON)            proposes N/H edits     (recon.py)
```

`reconcile.py` auto-applies only zero-judgment rules (verified 100% precision against the
reconciled reference) and flags everything else. Part X reconciliation is largely
**subtractive** (clearing column N), and clearing is almost always entity-variant
judgment — so most Part X work lands in `exceptions`, not `confident`.

> Design lesson: do **not** re-derive Part X from the Sch K-1 line. The reconciled state
> is the template defaults minus a small curated set of removals. Auto-deriving breaks
> tie-out. Be conservative: auto-fix only the safe rule below, flag the rest.

---

## Procedure

### Step 1 — Run the deterministic pre-pass
`python reconcile.py <workbook> -o payload.json`. The only **confident** rule that
touches Part X (column N), verified with zero counter-examples:

| Rule | Condition | Action (col N) |
|---|---|---|
| **R4 Qualified dividends** | Part X = `7 - Dividends`, Part II blank, description contains `"Qualified Dividends"` (not "Non-Qualified") | clear N (and set Part II line 8) |

Qualified dividends are a subset of ordinary dividends and must not be double-counted in
Part X. Everything else affecting column N is flagged for judgment (Step 2).

### Step 2 — Exceptions the script flags for judgment
The script surfaces these (current values shown, **never auto-changed**):

1. **Entity / foreign / fund-level variant duplicates** (description contains `CLIMATE`,
   `TPG`, `- Foreign`, `- US`, or `FUND LEVEL`). In the reference, **all 25** Part X
   removals were exactly these variant duplicates being deactivated — which one stays
   Active depends on which row carries the amount. Pure judgment.
2. **Part II / Part X pair is not a valid equivalence** (see the equivalence table in the
   Part II skill). Cases like line `9c` Unrecaptured §1250 or line `11h` §951(a) need the
   preparer to pick the correct line.

> **Correction to earlier notes:** there is **no** "line 10 / 11s excluded for this
> entity" rule. In the reference, §1231 (line 10) and similar rows *keep* their Part X
> mapping; only specific **entity-variant** rows (e.g. one TPG line-10 duplicate) were
> cleared. Earlier "Rule X-3" was a misread — clearing is variant-specific, not
> line-level. Treat all Part X removals as Step 2.1 judgment unless proven mechanical.

### Step 3 — Active? coordination (judgment)
Only the row carrying each amount stays active; duplicates are deactivated and their
Part X line cleared. This column is **shared** with `k2-part2-recon` — coordinate so both
parts agree. These rows come through as exceptions.

### Step 4 — Validate (must pass before returning)
1. `allocRec` N15 == `"OK"`.
2. `grossIncome`: `Part 2 - Part X` Difference == 0 for every K3 code (gain "X" and
   loss "Y" buckets) — Part X gross income agrees with Part II.
3. `schKToK2PartXDiff`: per `K3_Code` not in `partXExclude`, summed `Part X` == the
   `Schedule K` total.

`recon.py` / `validate_recon.py` compute checks 2–3 directly — reuse them as the harness.

---

## Output / write-back
Return values to drop into the workpaper — **nothing else is modified**:
- `allocRec` column **N** (Part X Line) for each affected data row.
- `allocRec` column **H** (`Active?`) for rows toggled in Step 3.

Provide write-back as `{sheet, cell, value}` records (e.g. `(K.01) Sch K to K2 Control!N128`).

## Coordination with Part II
- `k2-part2-recon` edits column **M** of the **same** sheet and shares column **H**.
- Run a single shared `Active?` decision before both mapping passes, or let Part X honor
  the active set chosen by the Part II pass.
- Final tie-out (`grossIncome`) validates Part II and Part X **together** — run it once
  after both passes.

## Guardrails
- Never edit `partXExclude`, `k2PartX`, or the `allocRec` header band / amount columns.
- Never invent K3 codes or amounts.
- If tie-out cannot be reached, return the unreconciled rows with a proposed mapping and
  the specific Difference, rather than forcing a value.
