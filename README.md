# ReconAgent — K-2 Reconciliation Agent

HTTP-triggered agent for Schedule K to K-2 reconciliation in federal partnership tax workpapers. Uses Microsoft Excel Graph API to fetch/write data, runs deterministic pre-pass logic, flags exceptions for LLM reasoning, then validates results.

## Architecture

```
HTTP POST /reconcile
    ↓
[src/api/handlers.reconcile_handler]
    ↓
[src/api/graph_client] fetches allocRec range
    ↓
[src/reconciliation/reconcile.run] — deterministic pre-pass (100% precision rules only)
    ↓
Returns: { confident: [...], exceptions: [...], summary: {...} }
    ↓
LLM reasons over exceptions only (not the 21MB workbook)
    ↓
HTTP POST /writeback
    ↓
[src/api/handlers.writeback_handler] writes changes back to Excel
    ↓
HTTP POST /validate
    ↓
[src/reconciliation/recon.run] — re-computes tie-out validation
```

## Directory Structure

```
ReconAgent/
├── src/
│   ├── reconciliation/        # Business logic (pure functions, no I/O)
│   │   ├── reconcile.py       # Pre-pass (deterministic rules)
│   │   ├── recon.py           # Tie-out computation (8 M-code queries)
│   │   └── validate_recon.py  # K3-agnostic validation harness
│   ├── api/                   # HTTP and Graph API layer
│   │   ├── graph_client.py    # Excel Graph API wrapper
│   │   └── handlers.py        # FastAPI request handlers
│   ├── llm/                   # LLM integration
│   │   ├── llm_reason.py      # Existing reasoning module
│   │   └── skills/            # SKILL.md documents
│   │       ├── k2-part2-recon/
│   │       └── k2-partx-recon/
│   ├── writeback/             # Write results back to workbook
│   └── etl/                   # Data transformation utilities
│
├── config/
│   ├── settings.py            # Environment config (GRAPH_CLIENT_ID, etc.)
│   └── __init__.py
│
├── data/
│   ├── raw/examples/          # Reference workbooks & documentation
│   ├── staging/               # Intermediate processing
│   └── processed/             # Final outputs
│
├── tests/                     # Test suite
│   ├── fixtures/              # Mock data
│   └── __init__.py
│
├── scripts/
│   └── main.py                # FastAPI server entry point
│
├── logs/                      # Runtime logs (git-ignored)
├── requirements.txt
└── README.md
```

## Setup

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure environment variables:**
   ```bash
   export GRAPH_CLIENT_ID="your-app-id"
   export GRAPH_TENANT_ID="your-tenant-id"
   export GRAPH_CLIENT_SECRET="your-secret"
   export GRAPH_SITE_ID="your-sharepoint-site-id"
   export GRAPH_WORKBOOK_ID="your-excel-workbook-id"
   export LOG_LEVEL="INFO"
   ```

3. **Start the server:**
   ```bash
   python scripts/main.py
   ```

## API Endpoints

### POST `/reconcile`
Run the deterministic pre-pass.
```json
{
  "entity": "Test Fund",
  "period": "2025-01-31"
}
```
Response: `{ "success": true, "data": { "confident": [...], "exceptions": [...], "summary": {...} } }`

### POST `/validate`
Validate after LLM proposes changes.
```json
{
  "entity": "Test Fund",
  "period": "2025-01-31"
}
```
Response: `{ "success": true, "data": { "status": "OK|FAIL", "grossIncome": {...}, ... } }`

### POST `/writeback`
Write LLM changes back to Excel.

Supports three payload formats:

1. Direct cell updates (existing format)
```json
{
  "updates": [
    { "sheet": "(K.01) Sch K to K2 Control", "cell": "M18", "value": "7 - Ordinary dividends (exclude amount on line 8)" },
    ...
  ],
  "entity": "Test Fund",
  "period": "2025-01-31"
}
```

2. K2II Part II row updates (mapped to sheet `K2II-Part II` by row keys)
```json
{
  "updates": [
    {
      "Activity Number": 101,
      "Section": "I",
      "Line": "6a",
      "Country Code (See Detail)": "US",
      "Detail": "Ordinary dividends",
      "(a)U.S. Source": 1200,
      "(b)Foreign branch category income": 0,
      "(c)Passive Category Income": 0,
      "(d) General Category Income": 0,
      "(e) Other (category code OTH)": 0,
      "(e) Other (Category code 901j)": 0,
      "Sourced by Partner": "No"
    }
  ],
  "entity": "TPG RISE",
  "period": "2025-01-31"
}
```

3. K2X Part X row updates (mapped to sheet `K2X-Part X` by row keys)
```json
{
  "updates": [
    {
      "Activity Number": 101,
      "Section": "I",
      "Line": "6a",
      "Details": "Ordinary dividends",
      "(b) Partner Dertmination": "Partner",
      "(c) U.S. Source": 1200,
      "(d) Foreign Source": 0,
      "(e) U.S. Source (FDAP)": 0,
      "(f) U.S. Source (Other)": 0,
      "(g) Foreign Source": 0
    }
  ],
  "entity": "TPG RISE",
  "period": "2025-01-31"
}
```
Response: `{ "success": true, "data": { "status": "success|partial|failure", "written": N, "failed": [...] } }`

When row-based payloads are used, updates are also written to dedicated sheets in the same workbook:
- Part II payloads → `Part2 Updates`
- Part X payloads → `PartX Updates`

### GET `/health`
Health check.
```
Response: { "status": "healthy" }
```

## Key Design Decisions

1. **Script-first architecture:** Deterministic rules run **before** the LLM, so the model never reads the 21MB workbook — it only reasons over exceptions.
2. **Conservative confidence:** Only auto-apply rules with **zero counter-examples** against ground truth. A wrong confident change breaks tie-out (worse than flagging).
3. **Part II ↔ Part X equivalence:** Reconciliation logic keys off the valid (P2, PX) pairs, not off Sch K-1 line numbers.
4. **Business logic separation:** `src/reconciliation/` modules are pure functions; `src/api/` handles all Graph API I/O.

## Modules

### `src/reconciliation/reconcile.py`
Deterministic pre-pass. Auto-applies only:
- **R3:** Ordinary dividends (line 6a, PX = `7 - Dividends`, P2 blank) → set P2 = `7 - Ordinary dividends`
- **R4:** Qualified dividends (description contains "Qualified Dividends", not "Non-Qualified", Part X populated, Part II blank) → set P2 = `8 - Qualified dividends`, clear PX
- **R2:** Line 9a LT gain (Part X = `11 - Net long-term capital gain`, Part II blank) → set P2 = `12 - Net long-term capital gain`

Everything else is flagged as an exception.

### `src/reconciliation/recon.py`
Python port of 8 M-code queries. Computes:
- `schKToK2Part2Diff`: Reconciliation of Schedule K to K-2 Part II
- `schKToK2PartXDiff`: Reconciliation of Schedule K to K-2 Part X
- `grossIncome`: Part II vs Part X comparison by K3 code

### `src/api/graph_client.py`
Microsoft Graph API wrapper. Handles:
- OAuth token acquisition
- Range fetching (sync & async)
- Cell writes (single & batch, sync & async)

### `src/api/handlers.py`
HTTP request handlers. Routes:
- `reconcile_handler`: Pre-pass
- `validate_handler`: Tie-out validation
- `writeback_handler`: Write changes back

### `config/settings.py`
Configuration from environment variables. Supports:
- Graph API credentials
- Logging config
- Site/workbook IDs

## Testing

```bash
# Run tests
pytest tests/

# Test reconciliation logic without Graph API
pytest tests/test_reconcile.py -v
```

## References

- **SKILL.md files:** `src/llm/skills/k2-part2-recon/SKILL.md` and `src/llm/skills/k2-partx-recon/SKILL.md`
- **Reverse-engineering notes:** `data/raw/examples/PartII_reconciliation_notes.md` and `PartX_reconciliation_notes.md`
- **Example workbooks:** `data/raw/examples/{unreconciled,reconciled}.xlsm`
