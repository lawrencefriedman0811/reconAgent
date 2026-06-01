# Part II Reconciliation — Reverse-Engineering Notes

Derived by diffing `examples/unreconciled.xlsm` against `examples/reconciled.xlsm`.
Only the five authorized ranges were inspected:

| Logical name | Location |
|---|---|
| `k2Part2`      | `'K2II-Part II'!A14:W3008` |
| `k2PartX`      | `'K2X-Part X'!A8:U2496` |
| `part2Exclude` | `'(Z.00) Dynamic Dropdown List'!FN2:FO79` |
| `partXExclude` | `'(Z.00) Dynamic Dropdown List'!FQ2:FR78` |
| `allocRec`     | `'(K.01) Sch K to K2 Control'!C10:GH1802` |

---

## 1. Where the reconciliation happens

**All reconciliation edits occur in `allocRec` (`(K.01) Sch K to K2 Control`).**
Two columns are edited (occasionally a third):

| Sheet col | Header (row 17) | Meta index in `recon.py` | Role |
|---|---|---|---|
| **M** | `Sch K2 - Part II Line` | meta col 10 → `Part 2 Line` | K-2 **Part II** line assignment |
| **N** | `Sch K2 - Part X Line`  | meta col 11 → `Part X Line`  | K-2 **Part X** line assignment |
| **H** | `Active?`                | meta col 5 → `Active`        | include/exclude the Sch K line |

These are the *same* `Part 2 Line` / `Part X Line` columns the `schK` query reads,
so editing them directly changes what flows into `schKToK2Part2Diff` and `grossIncome`.

### Things the reconciliation does NOT change
- **`part2Exclude` and `partXExclude` are byte-identical** between the two files.
  The exclusion lists are static config, never touched during reconciliation.
- `k2Part2` / `k2PartX` source amounts are not edited (they are imported data).
- The header band (rows 10–16) and the amount columns (N17→ right) are formulas
  that **recompute automatically** once M/N/H are set. They are outputs, not inputs.

### Built-in validation signal
Row 15, columns M & N hold a status cell that flips when the mapping is complete:

```
unreconciled:  "ERROR: Remove K-2 Selection (Red)"
reconciled:    "OK"
```

The reconciled file's Part II grand total (row 17, first activity col) moved from
`-26,249` → `-1,133,307`, i.e. previously-unmapped income was pulled in.

---

## 2. What changed (this example: 64 data rows)

| Column | ADD (blank→value) | REMOVE (value→blank) | CHANGE | unchanged |
|---|---|---|---|---|
| **M – Part II Line** | 42 | 2 | 0 | 20 |
| **N – Part X Line**  | 0  | 25 | 0 | 39 |

Plus **3 rows where `Active?` was flipped**:

| Row | K3 code | Sch K Line | Description | Active? |
|---|---|---|---|---|
| 98  | 30086 | 6a   | Ordinary Dividends – US        | Yes → **No** |
| 216 | 30212 | 11a  | Section 988 Gain               | No → **Yes** |
| 305 | 30305 | 11zz | Other Income (Loss) – FUND LEVEL | Yes → **No** |

> All 64 changed rows are detail lines under a Schedule K-1 line (the `Line`
> column holds `6a`, `6b`, `9a`, `10`, `11s`, …). Each Schedule K-1 line explodes
> into many detail codes (one `K-3 Code` per row), and the preparer assigns the
> correct K-2 line to each.

---

## 3. The reconciliation rule (the "judgment")

For every **active** Schedule K-1 detail row, assign the K-2 **Part II** line (col M)
and K-2 **Part X** line (col N) that matches the *character of the income*. The
choice is driven first by the **Schedule K-1 line number** and then refined by the
**detail description**.

### 3a. Default mapping by Schedule K-1 line

| Sch K-1 Line | → Part II line (M) | → Part X line (N) |
|---|---|---|
| **6a** Ordinary dividends | `7 - Ordinary dividends (exclude amount on line 8)` | `7 - Dividends` |
| **6b** Qualified dividends | `8 - Qualified dividends` | *(blank)* |
| **9a** Net LT capital gain | `12 - Net long-term capital gain` | `11 - Net long-term capital gain` |
| **10** Net §1231 gain | `15 - Net section 1231 gain` | `14 - Net section 1231 gain` |
| **11s** Non-portfolio LT cap gain | *(varies — see correction note)* | *(varies — see correction note)* |
| **11a** §988 gain | `20 - Other income (see instructions)` | `2 - Other income` |
| **11zz** Other income | `20 - Other income (see instructions)` | `2 - Other income` |

> **Correction:** an earlier draft listed lines **10** and **11s** as "blank/blank —
> excluded for this entity." That was a misread. §1231 (line 10) rows generally **keep**
> their mapping in the reconciled file; only specific entity-variant duplicates were
> cleared (because they were deactivated, not because of the line number). Do not apply a
> line-level exclusion for 10/11s — treat clearing as entity-variant judgment.

### 3b. Description-driven overrides (applied on top of the default)

These are the actual judgment rules, keyed off the detail **description text**:

1. **Qualified dividends → Part II line 8, Part X blank.**
   Any 6a/6b detail whose description contains *"Qualified Dividends"* maps to
   Part II `8 - Qualified dividends` and the **Part X line is cleared** (qualified
   dividends are a subset already carried in ordinary dividends; clearing Part X
   prevents double-counting).

2. **Short-term capital gain dividends → reclassified for Part X.**
   A 6a detail *"Short-Term Capital Gain Dividend"* keeps Part II `7 - Ordinary
   dividends` but maps Part X to `1 - Net short-term capital gain` (not `7 - Dividends`).

3. **Foreign / entity-specific dividend variants → left unmapped (blank/blank).**
   6a details whose description ends in *"- Foreign"*, or is entity-scoped
   (e.g. *"Ordinary Dividends - US - CLIMATE…"*, *"… - TPG …"*) are **not** mapped
   to Part II/Part X here — they are handled by a different sourcing path. The
   generic *"Ordinary Dividends - US"* line is the one that carries the amount.

4. **Everything else under 6a** (Money Market, Non-Qualified, REIT, PFIC, U.S.
   Obligations, Consent, From Diversified/Non-Diversified, etc.) follows the
   default 6a mapping (Part II 7 / Part X 7).

### 3c. Active? toggling

`Active?` is flipped so exactly one variant of a duplicated line carries the amount:
- Turn **off** a US line when its amount is represented elsewhere (rows 98, 305).
- Turn **on** a line that does hold a real amount (row 216, §988 gain).

This is the mechanism that makes the reconciliation *tie out* without
double-counting; it pairs with the description-driven blanking in 3b.

---

## 4. How to tell when it's reconciled (success criteria)

1. **Status cell** `allocRec` M15/N15 reads `"OK"` (not the red ERROR text).
2. **Gross-income tie-out**: in `grossIncome`, `Part 2 - Part X` Difference is 0
   for every K3 code that should tie (gain bucket "X" and loss bucket "Y").
3. **Sch K = Part II per code**: in `schKToK2Part2Diff`, the `Schedule K` total and
   the summed `Part 2` agree for each `K3_Code` not on `part2Exclude`.

These three are exactly what `recon.py` / `validate_recon.py` already compute, so
they double as the skill's automated validation harness.

---

## 5. Implications for the SKILL.md (deterministic core vs. reasoning)

**Deterministic (script) layer** — safe to hard-code:
- Read `allocRec`; locate each active Sch K-1 detail row.
- Apply the **default line mapping** (table 3a) by Sch K-1 line number.
- Apply mechanical overrides that are pure lookups: Qualified→8 + clear Part X;
  Short-term cap gain dividend→Part X line 1.
- After writing M/N, recompute and run the three tie-out checks in §4.

**Reasoning (LLM) layer** — needs judgment, surface for review:
- Deciding which duplicated/variant lines to **deactivate** vs keep (§3c) so amounts
  don't double-count — depends on which line actually carries the dollar amount.
- Classifying ambiguous detail descriptions (Foreign vs US, entity-specific, REIT,
  PFIC) when they don't match a known keyword.
- Any residual difference that doesn't tie out after the deterministic pass: explain
  it and propose the mapping/Active change, then re-validate.

**Write-back targets** (what the agent returns to drop into the file):
- `allocRec` column **M** (Part II line) and column **N** (Part X line) per data row.
- `allocRec` column **H** (`Active?`) for the rows that must be toggled.
- Nothing else — exclude tables and source amounts stay untouched.
