---
name: k2-part2-recon
description: >
  Reconcile Schedule K to K-2 Part II in a federal partnership tax workpaper by
  assigning each active Schedule K-1 detail line to the correct K-2 Part II line in
  the (K.01) Sch K to K2 Control sheet, then validating that Part II ties out. WHEN:
  "reconcile Part II", "reconcile K-2 Part II", "Sch K to K2 Part 2", "assign Part II
  lines", "K2 Part II mapping", "fix Part II reconciliation", "Part II doesn't tie
  out", "map Schedule K to K-2 Part II". DO NOT USE FOR: Part X (use k2-partx-recon),
  or editing K-2 source amounts / exclude lists.
location: user
---

# K-2 Part II Reconciliation

Assigns the correct **K-2 Part II line** to every active Schedule K-1 detail row in a
partnership workpaper, so that Schedule K gross income ties to K-2 Part II. This is the
deterministic + judgment logic reverse-engineered from a known unreconciled→reconciled
example pair.

## When to use
- A workpaper's K-2 **Part II** is unreconciled (status cell shows the red ERROR) or
  does not tie out, and the Part II line assignments need to be set/corrected.
- Triggered by a SharePoint button → agent pipeline, or an explicit user request.

Use the sibling skill **`k2-partx-recon`** for Part X (column N). The two skills share
the `Active?` column — see "Coordination with Part X" below.

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
| `k2Part2`      | `'K2II-Part II'!A14:W3008` | Part II source amounts (read-only) |
| `part2Exclude` | `'(Z.00) Dynamic Dropdown List'!FN2:FO79` | K3 codes excluded from tie-out (read-only, static) |

> Do **not** read or edit other ranges. `part2Exclude` and `k2Part2` are never modified
> by this skill (confirmed byte-identical across the reconciled example).

### `allocRec` layout
- Header band: rows 10–16 (per-activity metadata + formulas; do not edit).
- Column headers: **row 17**. Data rows: **18 → 1802**.
- Meta columns (sheet → meaning):
  - `C` Schedule K Line Item Code · `D` **K-3 Code** · `E` Detail code · `F` **Line**
    (Sch K-1 line, e.g. `6a`, `6b`, `9a`) · `G` **Description** · `H` **Active?** ·
    `J` Natural Sign · **`M` Sch K2 - Part II Line (EDIT TARGET)** · `N` Part X Line ·
    `O` Total.
- **Validation cell:** `allocRec` M15 reads `"ERROR: Remove K-2 Selection (Red)"` when
  unreconciled and `"OK"` when reconciled.

---

## How it runs (script-first → LLM exceptions → re-validate)

This skill is a **three-stage loop**. The deterministic script runs *before* the LLM
so the model never has to read the 21 MB workbook — it only reasons over a short
exceptions list.

```
reconcile.py (deterministic)  ──►  exceptions payload  ──►  LLM judgment  ──►  re-validate
   auto-applies safe fixes          (small JSON)            proposes M/H edits     (recon.py)
```

1. **`reconcile.py`** reads `allocRec`, **auto-applies only zero-judgment rules**
   (validated to 100% precision against the reconciled reference — a wrong confident
   change breaks tie-out, so the bar is "no observed counter-examples"), and **flags
   everything else** as exceptions. It emits:
   ```json
   { "confident": [ {row, K3, line, desc, proposed_part2, reason}, ... ],
     "exceptions": [ {row, K3, line, desc, current_part2, why}, ... ],
     "summary": { ... } }
   ```
2. **LLM** reasons over **only `exceptions`**, proposing column **M** / **H** edits.
3. **Re-validate** with `recon.py` / `validate_recon.py`; loop until tie-out.

> Design lesson: the reconciled workbook is *mostly* the template's existing column-M
> defaults with a small curated set of overrides. Do **not** re-derive the whole mapping
> from the Sch K-1 line number — that "fixes" rows already correct and breaks tie-out.
> Be conservative: auto-fix only the safe rules below, flag the rest.

---

## Procedure

### Step 1 — Run the deterministic pre-pass
`python reconcile.py <workbook> -o payload.json`. The confident rules currently
auto-applied for **Part II (column M)** — each verified with zero counter-examples:

| Rule | Condition | Action (col M) |
|---|---|---|
| **R3 Ordinary dividends** | line `6a` (or Part X = `7 - Dividends`), M blank, **not** qualified | M = `7 - Ordinary dividends (exclude amount on line 8)` |
| **R4 Qualified dividends** | Part X = `7 - Dividends`, M blank, description contains `"Qualified Dividends"` (not "Non-Qualified") | M = `8 - Qualified dividends` (and Part X cleared) |
| **R2 LT capital gain** | line `9a`, Part X = `11 - Net long-term capital gain`, M blank | M = `12 - Net long-term capital gain` |

These map a blank Part II line to the value implied by the row's Part X line, using the
**Part II ↔ Part X equivalence** (see table in Step 2). They never overwrite an existing
non-blank Part II value.

### Step 2 — Exceptions the script flags for judgment
Everything below is surfaced as an exception (current values shown, **never auto-changed**):

1. **Part II / Part X pair is not a valid equivalence.** A reconciled active row's
   `(M, N)` must be one of:

   | Part II (M) | Part X (N) |
   |---|---|
   | `7 - Ordinary dividends (exclude amount on line 8)` | `7 - Dividends` |
   | `8 - Qualified dividends` | *(blank)* |
   | `12 - Net long-term capital gain` | `11 - Net long-term capital gain` |
   | `14 - Unrecaptured section 1250 gain` | `11 - Net long-term capital gain` / `14 - Net section 1231 gain` |
   | `15 - Net section 1231 gain` | `14 - Net section 1231 gain` |
   | `20 - Other income (see instructions)` | `2 - Other income` |

   Anything else (e.g. line `9c` Unrecaptured §1250, line `11h` §951(a) inclusions) is
   preparer judgment — the correct line depends on the income's character, not on a
   mechanical rule.
2. **Entity / foreign / fund-level variant duplicates** (description contains `CLIMATE`,
   `TPG`, `- Foreign`, `- US`, or `FUND LEVEL`). Which duplicate stays Active and which
   are cleared depends on **which row carries the amount** — pure judgment.

> **Correction to earlier notes:** line `10` (§1231) and line `11s` are **not** "excluded
> for this entity." In the reference, lines 10/11 *keep* their mappings; only specific
> **entity-variant** rows were cleared. Treat clearing as variant-specific (Step 2.2),
> never as a line-level exclusion.

### Step 3 — Active? coordination (judgment)
Set `Active?` (`H`) so only the row that actually carries each amount is active; clear the
Part II line on deactivated duplicates. (Reference example flipped 3 rows: 30086
`6a Ordinary Dividends - US` Yes→No; 30212 `11a §988 Gain` No→Yes; 30305
`11zz Other Income - FUND LEVEL` Yes→No.) These come through as exceptions.

### Step 4 — Validate (must pass before returning)
1. `allocRec` M15 == `"OK"`.
2. `grossIncome`: `Part 2 - Part X` Difference == 0 for every K3 code (gain "X" and
   loss "Y" buckets).
3. `schKToK2Part2Diff`: per `K3_Code` not in `part2Exclude`, summed `Part 2` == the
   `Schedule K` total.

`recon.py` / `validate_recon.py` compute checks 2–3 directly — reuse them as the harness.

---

## Output / write-back
Return values to drop into the workpaper — **nothing else is modified**:
- `allocRec` column **M** (Part II Line) for each affected data row.
- `allocRec` column **H** (`Active?`) for rows toggled in Step 3.

Provide write-back as `{sheet, cell, value}` records (e.g. `(K.01) Sch K to K2 Control!M98`).

## Coordination with Part X
- `k2-partx-recon` edits column **N** of the **same** sheet and shares column **H**.
- Run a single shared `Active?` decision before both mapping passes, or run Part II
  first and let Part X honor its active set.
- Final tie-out (`grossIncome`) validates Part II and Part X **together** — run it once
  after both passes.

## Guardrails
- Never edit `part2Exclude`, `k2Part2`, or the `allocRec` header band / amount columns.
- Never invent K3 codes or amounts.
- If tie-out cannot be reached, return the unreconciled rows with a proposed mapping and
  the specific Difference, rather than forcing a value.
