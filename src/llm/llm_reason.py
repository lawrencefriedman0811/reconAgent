"""
llm_reason.py — chatBDO reasoning layer for K2 reconciliation.

Uses the BDO chatBDO platform (chattest.bdo.com / chatdev.bdo.com) as the LLM,
mirroring the auth + API pattern in index.js.

Auth:
  - Acquires a token via MSAL device code flow (local dev) or
    accepts a pre-acquired Bearer token forwarded from Power Automate (production).
  - Token scope: api://15343568-925f-4215-b9da-82e63fa8460a/access_as_user (test)

Config (env vars or .env):
  CHATBDO_ENV           "test" (default) or "dev"
  CHATBDO_CLIENT_ID     Azure AD app client ID (default: 6ee94015-...)
  CHATBDO_TENANT_ID     Azure AD tenant ID   (default: 6e57fc1a-...)
  CHATBDO_ACCESS_TOKEN  Pre-acquired Bearer token (skips MSAL when set)
"""

import json, os, pathlib, textwrap
import requests
import msal

import data as reconciliation_data

# ── Load skill files once at startup ─────────────────────────────────────────
_HERE = pathlib.Path(__file__).parent

def _load_skill_sections() -> str:
    """
    Auto-discovers every .md file in the skills/ subfolder and injects each one
    into the system prompt under a labeled block.

    Convention:
    - Files starting with '_' or named 'README' are skipped (easy disable).
    - Files are loaded in alphabetical order for deterministic prompt assembly.
    - To add a rule: drop a .md file in skills/. No code changes needed.
    - To disable a rule: rename the file with a leading underscore.
    """
    skills_dir = _HERE / "skills"
    if not skills_dir.is_dir():
        return ""

    blocks = []
    for path in sorted(skills_dir.glob("*.md")):
        name = path.stem
        if name.startswith("_") or name.upper() == "README":
            continue
        text = path.read_text(encoding="utf-8").strip()
        if text:
            blocks.append(f"=== SKILL: {name} ===\n{text}\n=== END: {name} ===")

    if not blocks:
        return ""

    joined = "\n\n".join(blocks)
    return f"\n\n--- SKILL GROUNDING ({len(blocks)} skills loaded) ---\n{joined}\n--- END SKILL GROUNDING ---"

_SKILL_GROUNDING = _load_skill_sections()

# ── Environment config (mirrors index.js ENV_CONFIG) ─────────────────────────
_ENVS = {
    "test": {
        "scope":      "api://15343568-925f-4215-b9da-82e63fa8460a/access_as_user",
        "persona_id": "28a5aa9f-818c-4d85-78e2-08dbc92f6fd5",
        "base_url":   "https://chattest.bdo.com",
    },
    "dev": {
        "scope":      "api://761a27e2-2c58-446e-b7bd-520b00d5cf63/access_as_user",
        "persona_id": "9f64db13-1691-4787-112c-08dbc605dd06",
        "base_url":   "https://chatdev.bdo.com",
    },
}

CLIENT_ID = os.getenv("CHATBDO_CLIENT_ID", "6ee94015-074f-4714-bd7f-bacefb7fed34")
TENANT_ID = os.getenv("CHATBDO_TENANT_ID", "6e57fc1a-413e-4050-91da-7d2dc8543e3c")

def _env_cfg() -> dict:
    key = os.getenv("CHATBDO_ENV", "test").lower()
    return _ENVS.get(key, _ENVS["test"])


# ── MSAL auth (device code — mirrors acquireTokenDeviceCode in index.js) ──────
_token_cache = msal.SerializableTokenCache()

def acquire_token_device_code() -> str:
    """
    Acquires a chatBDO access token via MSAL device code flow.
    Prints the sign-in URL + code to stdout. Caches the token for reuse.
    """
    cfg   = _env_cfg()
    scope = cfg["scope"]
    app   = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=_token_cache,
    )

    # Try silent first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent([scope], account=accounts[0])
        if result and "access_token" in result:
            print("✔  Acquired chatBDO token silently.")
            return result["access_token"]

    # Device code flow
    flow   = app.initiate_device_flow(scopes=[scope])
    print("\n" + flow["message"] + "\n")
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"MSAL auth failed: {result.get('error_description', result)}")
    return result["access_token"]


def get_token(forwarded_token: str | None = None) -> str:
    """
    Returns a valid chatBDO Bearer token.
    Priority:
      1. forwarded_token (passed from Power Automate's Authorization header)
      2. CHATBDO_ACCESS_TOKEN env var (CI / pre-configured service account)
      3. MSAL device code flow (local dev / first run)
    """
    if forwarded_token:
        return forwarded_token
    env_tok = os.getenv("CHATBDO_ACCESS_TOKEN", "")
    if env_tok:
        return env_tok
    return acquire_token_device_code()


# ── chatBDO REST helpers ──────────────────────────────────────────────────────

def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
    }

def _start_session(base_url: str, token: str, persona_id: str, content: str) -> dict:
    """POST /api/command/ChatSession/Start → returns value dict with chatSession + responseMessage."""
    resp = requests.post(
        f"{base_url}/api/command/ChatSession/Start",
        headers=_headers(token),
        json={
            "personaId":            persona_id,
            "content":              content,
            "name":                 None,
            "additionalEmbeddings": None,
            "fileLinks":            None,
        },
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()["value"]

def _stream_completion(base_url: str, token: str, chat_session_id: str,
                        response_message_id: str) -> str:
    """
    POST /api/command/ChatMessage/GenerateCompletionStream
    Collects all streamed chunks and returns the full response text.
    """
    resp = requests.post(
        f"{base_url}/api/command/ChatMessage/GenerateCompletionStream",
        headers=_headers(token),
        json={"chatSessionId": chat_session_id, "responseMessageId": response_message_id},
        stream=True,
        timeout=120,
    )
    resp.raise_for_status()
    chunks = []
    for chunk in resp.iter_content(chunk_size=None):
        if chunk:
            chunks.append(chunk.decode("utf-8"))
    return "".join(chunks)


# ── Prompt construction ───────────────────────────────────────────────────────
SYSTEM_CONTEXT = textwrap.dedent("""
You are a K-2 tax reconciliation specialist. You will receive one normalized JSON
payload describing a reconciliation difference between Schedule K and K-2 Part II.

RULES:
1. Do NOT invent rows or amounts — only reason over the provided data.
2. Use candidates[].af_status as the primary signal:
   OK → CONFIRM (already tied)  |  VAR → UPDATE (amount wrong)
   NO_MATCH → UPDATE (wrong tag)  |  BLANK → TAG (no relationship)
3. k01.required_part2_line (K.01 col L, H=Yes only) is the authoritative K2II Part II line.
   If a candidate row's current line contradicts k01.required_part2_line, that IS the fix.
   Cross-check against k01.control_rows[].k2_part2_line — they must agree.
4. k01.control_rows are the actual K.01 rows from load_k01() for this K3 code.
   Use k2_part2_line (col L) and amount_buckets (cols O+) to verify line assignments
   and isolate which activity/allocation bucket contains the variance.
5. issue.k3_is_inactive=true means the K3 code is H=No in K.01 — re-route to the active K3.
6. For VAR rows: if computed_variance equals issue.difference, you have isolated the row.
7. Confidence: High (clear AF signal + required_part2_line known) | Medium (T1/T2/inactive K3) | Low (no trace)
""").strip() + _SKILL_GROUNDING + textwrap.dedent("""

Return ONLY a JSON array — no markdown fences, no extra text:
[
  {
    "k2ii_row": <integer|null>,
    "action": "<CONFIRM|UPDATE|TAG|CREATE_OR_LOCATE>",
    "recommended_sch_k_line": "<exact Part II line for col AC>",
    "confidence": "<High|Medium|Low>",
    "reason_codes": ["<af_ok|af_var|af_no_match|untagged|inactive_k3|line_mismatch|amount_mismatch>"],
    "explanation": "<1-2 sentence plain-English rationale and fix>",
    "computed_variance": <integer|null>
  }
]
""").strip()

def _build_prompt(diff: dict) -> str:
    payload = reconciliation_data.normalize_difference(diff)
    return (
        SYSTEM_CONTEXT + "\n\n"
        "Normalized reconciliation difference:\n"
        + json.dumps(payload, default=str, ensure_ascii=True)
        + "\n\nReturn the JSON array now."
    )


# ── Core reasoning call ───────────────────────────────────────────────────────
def reason_over_diff(diff: dict, token: str, cfg: dict) -> list[dict]:
    """Starts a chatBDO session for one diff, streams the response, parses JSON array."""
    prompt   = _build_prompt(diff)
    session  = _start_session(cfg["base_url"], token, cfg["persona_id"], prompt)
    chat_id  = session["chatSession"]["id"]
    resp_id  = session["responseMessage"]["id"]
    raw_text = _stream_completion(cfg["base_url"], token, chat_id, resp_id)

    # Parse — model should return a bare JSON array
    raw_text = raw_text.strip()
    # Strip markdown fences if present
    if raw_text.startswith("```"):
        raw_text = "\n".join(
            line for line in raw_text.splitlines()
            if not line.startswith("```")
        ).strip()

    parsed = json.loads(raw_text)
    if isinstance(parsed, list):
        return parsed
    # Unwrap common wrapper keys
    for key in ("rows", "result", "items", "data"):
        if isinstance(parsed, dict) and key in parsed and isinstance(parsed[key], list):
            return parsed[key]
    return [parsed] if isinstance(parsed, dict) else []


# ── Batch reasoning ───────────────────────────────────────────────────────────
def reason_all(recon_output: dict, forwarded_token: str | None = None) -> dict:
    """
    Runs chatBDO reasoning over every normalized difference.
    Returns {by_row: {k2ii_row: rec}, flat_rows: [rec, ...]}.
    """
    token = get_token(forwarded_token)
    cfg   = _env_cfg()
    normalized_output = reconciliation_data.normalize_reconciliation_output(recon_output)
    diffs = normalized_output.get("differences", [])

    by_row:    dict[int | None, dict] = {}
    flat_rows: list[dict]             = []

    print(f"Calling chatBDO ({cfg['base_url']}) for {len(diffs)} differences…")
    for i, diff in enumerate(diffs, 1):
        diff_id = diff.get("diff_id", f"diff_{i}")
        issue = diff.get("issue", {})
        try:
            rows = reason_over_diff(diff, token, cfg)
            for rec in rows:
                rn = rec.get("k2ii_row")
                rec.update({
                    "diff_id":    diff_id,
                    "k3_code":    issue.get("k3_code"),
                    "sch_k_line": issue.get("sch_k_line"),
                    "difference": issue.get("difference"),
                    "root_cause": issue.get("root_cause_type"),
                })
                if rn not in by_row:
                    by_row[rn] = rec
                flat_rows.append(rec)
            print(f"  [{i}/{len(diffs)}] {diff_id}: {len(rows)} row(s) reasoned")
        except Exception as exc:
            print(f"  [{i}/{len(diffs)}] {diff_id}: chatBDO error — {exc}")
            fallback_rows = diff.get("engine_recommendation", {}).get("rows", [])[:1]
            if not fallback_rows:
                fallback_rows = diff.get("candidates", [])[:1]
            if not fallback_rows:
                fallback_rows = [{"k2ii_row": None}]
            for cand in fallback_rows:
                rn = cand.get("k2ii_row")
                fallback = {
                    "k2ii_row":               rn,
                    "action":                 "REVIEW",
                    "recommended_sch_k_line": (
                        cand.get("recommended_sch_k_line")
                        or cand.get("k01_required_line")
                        or diff.get("k01", {}).get("required_part2_line", "")
                    ),
                    "confidence":             "Low",
                    "reason_codes":           ["chatbdo_error_fallback"],
                    "explanation":            f"chatBDO call failed ({exc}). Review manually.",
                    "computed_variance":      cand.get("computed_variance"),
                    "diff_id": diff_id, "k3_code": issue.get("k3_code"),
                    "sch_k_line": issue.get("sch_k_line"),
                    "difference": issue.get("difference"),
                    "root_cause": issue.get("root_cause_type"),
                }
                if rn not in by_row:
                    by_row[rn] = fallback
                flat_rows.append(fallback)

    return {"by_row": by_row, "flat_rows": flat_rows, "normalized_output": normalized_output}
