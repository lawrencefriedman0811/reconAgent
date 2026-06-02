# ReconAgent — K-2 Reconciliation Service

A simple **file-in / file-out** HTTP service for Schedule K to K-2 reconciliation
in federal partnership tax workpapers.

You upload the Excel workbook, the service runs a deterministic pre-pass over its
sheets, applies the confident (zero-judgment) fixes back into the control sheet,
adds an exceptions sheet listing rows that need preparer judgment, and returns the
modified workbook. It is designed to be triggered from Excel via a **Power Automate**
flow, but any HTTP client (curl, Postman) works the same way.

No Microsoft Graph API, no Entra app registration, no MCP/Copilot agent — just one
HTTP endpoint that takes a workbook and returns a workbook.

## How it works

```
Excel  ──(Power Automate: Get file content)──►  POST /run  (multipart upload)
                                                    │
                                          [scripts/main.py]
                                                    │
                                   [src/api/handlers.run_reconciliation]
                                                    │
                          [src/reconciliation/reconcile.run]  (deterministic pre-pass)
                                                    │
                   apply confident fixes to control sheet cols M / N
                   + add "ReconAgent Exceptions" sheet
                                                    │
Excel  ◄──(Power Automate: Update file content)──  modified .xlsm (binary response)
```

The reconciliation engine is intentionally conservative: it auto-applies **only**
rules with zero counter-examples against the reconciled ground truth (100% precision).
Anything that looks inconsistent but needs judgment is **flagged** on the exceptions
sheet rather than changed.

## What the service writes back

On the control sheet `(K.01) Sch K to K2 Control` (data rows 18–1802):

- **Confident fixes** are written into **column M** (Sch K-2 Part II line) and
  **column N** (Sch K-2 Part X line) for the rows where a zero-judgment rule fires.
- A sheet named **`ReconAgent Exceptions`** is created (or replaced) listing every
  row that needs judgment, with columns:
  `Control Row | K3 Code | Sch K-1 Line | Description | Current Part II | Current Part X | Proposed Part II | Proposed Part X | Why (needs judgment)`.

Everything else in the workbook (source amounts, macros, other sheets) is left
untouched. `.xlsm` macros are preserved.

### Confident rules (zero-judgment)

| Rule | Condition | Action |
|------|-----------|--------|
| **R1** | Qualified dividends already on Part II line 8, also present in Part X | Clear Part X (qualified div is Part II only) |
| **R2** | Sch K-1 line 9a, Part X = long-term capital gain (line 11), Part II blank | Set Part II = line 12 (Net LT capital gain) |
| **R3** | Part II blank, Part X populated, ordinary dividends (line 6a / Part X line 7), not qualified | Set Part II = line 7 (Ordinary dividends), retain Part X |
| **R4** | Qualified dividends carried only in Part X (line 7), Part II blank | Set Part II = line 8 (Qualified dividends), clear Part X |

## API

### `POST /run`

Run reconciliation on an uploaded workbook and return the modified file.

- **Request:** `multipart/form-data` with a single field **`file`** containing the
  `.xlsm` (or `.xlsx`) workbook.
- **Response:** the modified workbook as a binary download
  (`Content-Type: application/vnd.ms-excel.sheet.macroEnabled.12` for `.xlsm`).
- **Response headers** (summary counts):
  - `X-Recon-Confident-Applied` — number of confident fixes written
  - `X-Recon-Exceptions` — number of rows flagged on the exceptions sheet
  - `X-Recon-Data-Rows` — number of data rows scanned
- **Auth (optional):** if the `RECON_API_KEY` environment variable is set, the
  request must include an `X-API-Key` header with the matching value. When unset,
  no auth is required (local-only testing).

### `GET /health`

Health check. Returns `{ "status": "healthy" }`.

## Setup

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **(Optional) configure environment** — copy `.env.example` to `.env`:
   ```
   LOG_LEVEL=INFO
   LOG_FILE=logs/recon_agent.log
   PORT=8000
   RECON_API_KEY=        # set a shared secret before exposing publicly
   ```

3. **Start the server:**
   ```bash
   python scripts/main.py
   ```
   The server listens on `http://0.0.0.0:8000` (override with `PORT`).

## Local test with curl

```bash
# Upload a workbook and save the reconciled result
curl.exe -F "file=@data/raw/examples/unreconciled.xlsm" \
  http://localhost:8000/run -o out.xlsm

# With an API key configured:
curl.exe -H "X-API-Key: YOUR_SECRET" \
  -F "file=@data/raw/examples/unreconciled.xlsm" \
  http://localhost:8000/run -o out.xlsm
```

Open `out.xlsm` and confirm the control sheet columns M/N are updated and a
`ReconAgent Exceptions` sheet is present.

## Triggering from Excel via Power Automate

Because Power Automate runs in Microsoft's cloud, it **cannot reach `localhost`**.
For testing, expose the local server with a tunnel (e.g. `devtunnel host -p 8000`
or `ngrok http 8000`) and use the resulting public HTTPS URL. For production, host
the service on a reachable URL (Azure App Service, Container Apps, a VM, etc.).

> The **HTTP** action used below is a Power Automate **premium** connector.

### Flow steps

1. **Trigger** — choose how the flow starts, e.g.:
   - *For a selected file* (SharePoint/OneDrive), or
   - *Manually trigger a flow*, or
   - an Office Scripts / Excel button (`Run a flow`).

2. **Get file content** — OneDrive for Business or SharePoint **Get file content**.
   - *File:* the workbook to reconcile.
   - This outputs **File Content** (the workbook bytes).

3. **HTTP** action — send the workbook to the service:
   - **Method:** `POST`
   - **URI:** `https://<your-tunnel-or-host>/run`
   - **Headers:**
     - `Content-Type`: `multipart/form-data`
     - `X-API-Key`: `<your secret>`  *(only if `RECON_API_KEY` is set)*
   - **Body:** select **form-data** body type and add one part:
     - **Key/Name:** `file`
     - **Value:** the **File Content** output from step 2
     - **Filename:** the original file name (e.g. `workbook.xlsm`)

   If your Power Automate plan's HTTP action does not expose a form-data body
   builder, use the **multipart** body shape instead:
   ```json
   {
     "$content-type": "multipart/form-data",
     "$multipart": [
       {
         "headers": {
           "Content-Disposition": "form-data; name=\"file\"; filename=\"workbook.xlsm\""
         },
         "body": @{body('Get_file_content')}
       }
     ]
   }
   ```

4. **Update file content** — OneDrive/SharePoint **Update file content** (or
   *Create file*) to write the result back:
   - *File:* the same file (to overwrite) or a new name.
   - *File Content:* the **Body** output of the HTTP action (the returned workbook).

That's the whole flow: **Get file content → HTTP POST /run → Update file content.**
Optionally read the `X-Recon-*` response headers to post a summary (e.g. a Teams
message or email) of how many fixes were applied and how many exceptions remain.

## Directory structure

```
ReconAgent/
├── scripts/
│   └── main.py                 # FastAPI server: POST /run, GET /health
├── src/
│   ├── api/
│   │   ├── __init__.py         # exports run_reconciliation
│   │   └── handlers.py         # run_reconciliation(): file in → file out
│   └── reconciliation/
│       ├── __init__.py
│       └── reconcile.py        # deterministic pre-pass engine
├── config/
│   ├── __init__.py
│   └── settings.py             # logging config
├── data/
│   └── raw/examples/           # reference workbooks & notes (test fixtures)
├── tests/
├── .env.example
├── requirements.txt
└── README.md
```

## Dependencies

Only what the `/run` workflow needs:

```
fastapi, uvicorn, pydantic, python-multipart, openpyxl, python-dotenv
```

The reconciliation engine reads and writes the workbook with **openpyxl** only —
no pandas/numpy required.

## Key design decisions

1. **File-in / file-out.** The workbook travels through HTTP as bytes; no Graph API,
   no cloud auth plumbing. Power Automate (or any client) owns getting the file in
   and writing it back.
2. **Conservative confidence.** Only zero-judgment rules are auto-applied. A wrong
   confident change breaks tie-out, which is worse than flagging — so everything
   uncertain goes to the exceptions sheet.
3. **Part II ↔ Part X equivalence.** Reconciliation keys off the valid (Part II,
   Part X) pairing, not off the Sch K-1 line number.

## References

- Reverse-engineering notes: `data/raw/examples/PartII_reconciliation_notes.md`
  and `PartX_reconciliation_notes.md`
- Example workbooks: `data/raw/examples/{unreconciled,reconciled}.xlsm`
