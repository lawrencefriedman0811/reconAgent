# Part X Reconciliation — Reverse-Engineering Notes

Derived by diffing `examples/unreconciled.xlsm` against `examples/reconciled.xlsm`,
using only the five authorized ranges (see `PartII_reconciliation_notes.md` for the
range list). Read this *after* the Part II notes — the two parts are reconciled in
the **same control sheet** and share the same mechanics.

---

## 1. Where the Part X reconciliation happens

**All Part X edits occur in `allocRec` (`(K.01) Sch K to K2 Control`), column N.**

| Sheet col | Header (row 17) | Meta index in `recon.py` | Role |
|---|---|---|---|
| **N** | `Sch K2 - Part X Line` | meta col 11 → `Part X Line` | K-2 **Part X** line assignment |
| **H** | `Active?` | meta col 5 → `Active` | include/exclude the Sch K line |

Column N is the same `Part X Line` field that the `schK` query reads and that flows
into `schKToK2PartXDiff` and the Part X side of `grossIncome`. Part II uses col M;
Part X uses col N. A reconciliation pass typically touches both columns, but the
**Part X-specific work is entirely in column N** (plus the shared `Active?` flips).

### Not changed by Part X reconciliation
- `partXExclude` (`'(Z.00) Dynamic Dropdown List'!FQ2:FR78`) is **byte-identical**
  between the two files — static config, never edited.
- `k2PartX` source amounts (`'K2X-Part X'!A8:U2496`) are imported data, not edited.
- Part X header band and amount columns are formulas that recompute automatically.

---

## 2. What changed (this example: 25 column-N edits)

| Column | ADD (blank→value) | REMOVE (value→blank) | CHANGE |
|---|---|---|---|
| **N – Part X Line** | 0 | **25** | 0 |

**Every Part X change in this example is a REMOVE** — the default Part X line was
cleared. Part X reconciliation here is *purely subtractive*: the template
pre-populates a default Part X line for many rows, and reconciliation prunes the
ones that should not flow to Part X.

Grouped by Schedule K-1 line:

| Sch K-1 Line | Part X REMOVE count | What it is |
|---|---|---|
| **6b** | 16 | Qualified dividends (all detail codes) |
| **6a** | 7  | Ordinary-dividend variants (3 qualified + 4 entity/foreign duplicates) |
| **10** | 1  | Net §1231 gain |
| **11s** | 1 | Non-portfolio long-term capital gain |

---

## 3. The Part X reconciliation rule (the "judgment")

Start from the template's **default Part X line** (e.g. dividends default to
`7 - Dividends`, LT gain to `11 - Net long-term capital gain`, §1231 to
`14 - Net section 1231 gain`). Then **clear the Part X line** in these cases:

### Rule X-1 — Qualified dividends never flow to Part X (19 of 25 rows)
Any detail under line **6a or 6b** whose description contains *"Qualified Dividends"*
(including the `Line 6b - Custom N`, `… From Diversified/Non-Diversified`, `REIT
Dividends` under 6b) → **Part X line cleared (blank)**.
Part II still maps these to `8 - Qualified dividends`; only Part X is removed. (Part X
line 7 "Dividends" is meant to carry ordinary dividends; qualified is a subset already
represented, so it is excluded from Part X to avoid double-counting.)

### Rule X-2 — Deactivated ordinary-dividend duplicates → Part X cleared (4 rows)
Entity-scoped / foreign duplicates of ordinary dividends
(`Ordinary Dividends - US`, `… - US - CLIMATE …`, `… - US - TPG RISE …`,
`… - Foreign - TPG …`) that are turned **Active? = No** during reconciliation also
have their Part X line cleared. Only the single master "Ordinary Dividends - US" row
that stays active carries the Part X `7 - Dividends` amount; the duplicates are
zeroed out on both Part II and Part X. (See the 3 `Active?` flips documented in the
Part II notes — they drive this.)

### Rule X-3 — CORRECTED: there is no line-level §1231 exclusion
> **Correction (found during `reconcile.py` validation):** an earlier draft of these
> notes claimed lines **10** (§1231) and **11s** (non-portfolio LT gain) are "excluded
> from this entity's K-2 entirely." That was a **misread**. Re-checking the full
> control grid against the reconciled workbook shows §1231 / line-10 rows generally
> **keep** their Part X mapping; only specific **entity-variant duplicates** (e.g. a
> single TPG line-10 row) were cleared — and they were cleared because they are
> variant duplicates being deactivated (Rule X-2), **not** because of their line number.
>
> Do **not** apply a blanket "clear Part X for lines 10/11s" rule. Treat every Part X
> removal as an entity-variant / Active? judgment (Rule X-2), which is exactly how
> `reconcile.py` flags them (as exceptions, never auto-cleared).

### Note on additions
In this example Part X reconciliation only *removed* lines. In a workpaper where the
template under-populates Part X, the same logic runs in reverse: **add** the correct
Part X line for an active income row whose Part X cell is blank, using the default
Part X line for that Sch K-1 line (table in §3a of the Part II notes / §3 here).

---

## 4. How to tell when Part X is reconciled (success criteria)

1. **Status cell** `allocRec` N15 reads `"OK"` (not `"ERROR: Remove K-2 Selection
   (Red)"`).
2. **Gross-income tie-out** (`grossIncome`): for each K3 code, `Part 2 - Part X`
   Difference is 0 in both the gain ("X") and loss ("Y") buckets — i.e. Part X gross
   income agrees with Part II gross income.
3. **Sch K vs Part X** (`schKToK2PartXDiff`): the `Schedule K` total and summed
   `Part X` agree for each `K3_Code` not on `partXExclude`.

`recon.py` / `validate_recon.py` already compute these, so they are the skill's
automated Part X validation harness.

---

## 5. Implications for the Part X SKILL.md

**Deterministic (script) layer:**
- Read `allocRec`; auto-apply ONLY zero-judgment rules verified to 100% precision
  against the reconciled reference (a wrong confident change breaks tie-out):
  - **R4**: description contains "Qualified Dividends" (not "Non-Qualified") and Part X
    = `7 - Dividends` with Part II blank → clear Part X (and set Part II line 8).
- Do **not** auto-derive the full Part X mapping from the Sch K-1 line, and do **not**
  apply any "lines {10, 11s} → clear" rule (that earlier X-3 was a misread — see §3).
- After writing col N, recompute and run the three tie-out checks in §4.

**Reasoning (LLM) layer:**
- **X-2** duplicate/entity-variant handling: decide which ordinary-dividend variant
  stays Active and carries Part X vs. which duplicates are blanked — depends on which
  row holds the real dollar amount (cannot be inferred from line number alone).
- Classifying ambiguous descriptions (Foreign vs US, REIT, PFIC, §988) when no
  keyword matches a known rule.
- Any K3 code whose Part X still doesn't tie after the deterministic pass: explain the
  residual and propose the col-N / `Active?` change, then re-validate.

**Write-back targets** (what the agent returns to drop into the file):
- `allocRec` column **N** (Part X line) per data row.
- `allocRec` column **H** (`Active?`) for rows that must be toggled (shared with the
  Part II pass — coordinate so both parts agree on which rows are active).
- Nothing else — `partXExclude` and `k2PartX` source amounts stay untouched.

---

## 6. Relationship between the Part II and Part X skills

- Both write to the **same control sheet** (`allocRec`): Part II → col M, Part X → col N.
- They **share** the `Active?` column (col H). The 3 `Active?` flips in this example
  affect both parts simultaneously, so the two skills must agree on the active set
  (or one shared "activation" step should run before both mapping passes).
- They are validated **together** by `grossIncome` (Part 2 vs Part X must tie), so a
  combined validation step is recommended even if the mapping logic is split into two
  skills.
