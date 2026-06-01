"""HTTP request handlers for reconciliation endpoints."""

from __future__ import annotations

import logging
from typing import Any

from openpyxl.utils import get_column_letter

from src.api.graph_client import GraphClient
from src.reconciliation import reconcile, recon, validate_recon

logger = logging.getLogger(__name__)

# Authorized ranges (sheet_name, range_spec)
RANGES = {
    "allocRec": ("(K.01) Sch K to K2 Control", "C10:GH1802"),
    "k2Part2": ("K2II-Part II", "A14:W3008"),
    "k2PartX": ("K2X-Part X", "A8:U2496"),
    "part2Exclude": ("(Z.00) Dynamic Dropdown List", "FN2:FO79"),
    "partXExclude": ("(Z.00) Dynamic Dropdown List", "FQ2:FR78"),
}

PART2_SHEET = "K2II-Part II"
PART2_UPDATES_SHEET = "Part2 Updates"
PART2_FIRST_ROW = 14  # Header row for range A14:W3008
PART2_WRITEBACK_COLUMNS = [
    "Activity Number",
    "Section",
    "Line",
    "Country Code (See Detail)",
    "Detail",
    "(a)U.S. Source",
    "(b)Foreign branch category income",
    "(c)Passive Category Income",
    "(d) General Category Income",
    "(e) Other (category code OTH)",
    "(e) Other (Category code 901j)",
    "Sourced by Partner",
]
PART2_ROW_KEY_COLUMNS = [
    "Activity Number",
    "Section",
    "Line",
    "Country Code (See Detail)",
    "Detail",
]
PART2_HEADER_ALIASES = {
    "Country Code (See Detail)": [
        "Country Code (See Detail)",
        "Country Code (See Note)",
    ],
}

PARTX_SHEET = "K2X-Part X"
PARTX_UPDATES_SHEET = "PartX Updates"
PARTX_FIRST_ROW = 8  # Header row for range A8:U2496
PARTX_WRITEBACK_COLUMNS = [
    "Activity Number",
    "Section",
    "Line",
    "Details",
    "(b) Partner Dertmination",
    "(c) U.S. Source",
    "(d) Foreign Source",
    "(e) U.S. Source (FDAP)",
    "(f) U.S. Source (Other)",
    "(g) Foreign Source",
]
PARTX_ROW_KEY_COLUMNS = [
    "Activity Number",
    "Section",
    "Line",
    "Details",
]
PARTX_HEADER_ALIASES = {
    "Details": ["Details", "Detail"],
    "(b) Partner Dertmination": [
        "(b) Partner Dertmination",
        "(b) Partner Determination",
        "(b) Partner determination",
    ],
}


def _normalize_text(value: Any) -> str:
    """Normalize value for stable header/row matching."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _normalized_match_key(value: Any) -> str:
    """Normalized comparison key with collapsed whitespace and lowercase."""
    return " ".join(_normalize_text(value).lower().split())


def _resolve_part2_column_indices(headers: list[Any]) -> dict[str, int]:
    """Map required Part II payload columns to workbook column indexes."""
    normalized_headers = {
        _normalized_match_key(header): idx for idx, header in enumerate(headers)
    }
    resolved: dict[str, int] = {}

    for column_name in PART2_WRITEBACK_COLUMNS:
        candidates = PART2_HEADER_ALIASES.get(column_name, [column_name])
        resolved_idx = None
        for candidate in candidates:
            key = _normalized_match_key(candidate)
            if key in normalized_headers:
                resolved_idx = normalized_headers[key]
                break
        if resolved_idx is None:
            raise ValueError(
                f"Part II write-back column not found in sheet headers: '{column_name}'"
            )
        resolved[column_name] = resolved_idx

    return resolved


def _is_part2_row_update(update: dict[str, Any]) -> bool:
    """Detect new Part II row-write payload format."""
    return all(col in update for col in PART2_ROW_KEY_COLUMNS)


def _resolve_partx_column_indices(headers: list[Any]) -> dict[str, int]:
    """Map required Part X payload columns to workbook column indexes."""
    normalized_headers = {
        _normalized_match_key(header): idx for idx, header in enumerate(headers)
    }
    resolved: dict[str, int] = {}

    for column_name in PARTX_WRITEBACK_COLUMNS:
        candidates = PARTX_HEADER_ALIASES.get(column_name, [column_name])
        resolved_idx = None
        for candidate in candidates:
            key = _normalized_match_key(candidate)
            if key in normalized_headers:
                resolved_idx = normalized_headers[key]
                break
        if resolved_idx is None:
            raise ValueError(
                f"Part X write-back column not found in sheet headers: '{column_name}'"
            )
        resolved[column_name] = resolved_idx

    return resolved


def _is_partx_row_update(update: dict[str, Any]) -> bool:
    """Detect Part X row-write payload format."""
    return all(col in update for col in PARTX_ROW_KEY_COLUMNS)


def _find_part2_row_number(
    data_rows: list[list[Any]],
    col_indices: dict[str, int],
    update: dict[str, Any],
) -> int:
    """Find the unique worksheet row for a Part II update record."""
    matches: list[int] = []
    for data_idx, row in enumerate(data_rows[1:], start=1):
        row_matches = True
        for key in PART2_ROW_KEY_COLUMNS:
            col_idx = col_indices[key]
            sheet_value = row[col_idx] if col_idx < len(row) else None
            if _normalized_match_key(sheet_value) != _normalized_match_key(update.get(key)):
                row_matches = False
                break
        if row_matches:
            matches.append(PART2_FIRST_ROW + data_idx)

    if not matches:
        raise ValueError(
            "No matching K2II-Part II row found for keys: "
            + ", ".join(f"{k}={update.get(k)!r}" for k in PART2_ROW_KEY_COLUMNS)
        )
    if len(matches) > 1:
        raise ValueError(
            "Multiple K2II-Part II rows match keys: "
            + ", ".join(f"{k}={update.get(k)!r}" for k in PART2_ROW_KEY_COLUMNS)
        )
    return matches[0]


def _find_partx_row_number(
    data_rows: list[list[Any]],
    col_indices: dict[str, int],
    update: dict[str, Any],
) -> int:
    """Find the unique worksheet row for a Part X update record."""
    matches: list[int] = []
    for data_idx, row in enumerate(data_rows[1:], start=1):
        row_matches = True
        for key in PARTX_ROW_KEY_COLUMNS:
            col_idx = col_indices[key]
            sheet_value = row[col_idx] if col_idx < len(row) else None
            if _normalized_match_key(sheet_value) != _normalized_match_key(update.get(key)):
                row_matches = False
                break
        if row_matches:
            matches.append(PARTX_FIRST_ROW + data_idx)

    if not matches:
        raise ValueError(
            "No matching K2X-Part X row found for keys: "
            + ", ".join(f"{k}={update.get(k)!r}" for k in PARTX_ROW_KEY_COLUMNS)
        )
    if len(matches) > 1:
        raise ValueError(
            "Multiple K2X-Part X rows match keys: "
            + ", ".join(f"{k}={update.get(k)!r}" for k in PARTX_ROW_KEY_COLUMNS)
        )
    return matches[0]


def _expand_part2_updates(
    graph_client: GraphClient,
    updates: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Convert Part II row payload updates into concrete cell writes."""
    part2_data = graph_client.fetch_range(RANGES["k2Part2"][0], RANGES["k2Part2"][1])
    if not part2_data:
        raise ValueError("Unable to load K2II-Part II range for write-back.")

    headers = part2_data[0]
    col_indices = _resolve_part2_column_indices(headers)

    cell_updates: list[dict[str, Any]] = []
    failed: list[dict[str, Any]] = []

    for update in updates:
        missing = [c for c in PART2_WRITEBACK_COLUMNS if c not in update]
        if missing:
            failed.append(
                {
                    **update,
                    "error": "Missing required Part II columns: " + ", ".join(missing),
                }
            )
            continue

        try:
            row_num = _find_part2_row_number(part2_data, col_indices, update)
            for column_name in PART2_WRITEBACK_COLUMNS:
                col_idx = col_indices[column_name]
                cell_updates.append(
                    {
                        "sheet": PART2_SHEET,
                        "cell": f"{get_column_letter(col_idx + 1)}{row_num}",
                        "value": update[column_name],
                    }
                )
        except ValueError as exc:
            failed.append({**update, "error": str(exc)})

    return cell_updates, failed


def _expand_partx_updates(
    graph_client: GraphClient,
    updates: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Convert Part X row payload updates into concrete cell writes."""
    partx_data = graph_client.fetch_range(RANGES["k2PartX"][0], RANGES["k2PartX"][1])
    if not partx_data:
        raise ValueError("Unable to load K2X-Part X range for write-back.")

    headers = partx_data[0]
    col_indices = _resolve_partx_column_indices(headers)

    cell_updates: list[dict[str, Any]] = []
    failed: list[dict[str, Any]] = []

    for update in updates:
        missing = [c for c in PARTX_WRITEBACK_COLUMNS if c not in update]
        if missing:
            failed.append(
                {
                    **update,
                    "error": "Missing required Part X columns: " + ", ".join(missing),
                }
            )
            continue

        try:
            row_num = _find_partx_row_number(partx_data, col_indices, update)
            for column_name in PARTX_WRITEBACK_COLUMNS:
                col_idx = col_indices[column_name]
                cell_updates.append(
                    {
                        "sheet": PARTX_SHEET,
                        "cell": f"{get_column_letter(col_idx + 1)}{row_num}",
                        "value": update[column_name],
                    }
                )
        except ValueError as exc:
            failed.append({**update, "error": str(exc)})

    return cell_updates, failed


def _write_updates_table(
    graph_client: GraphClient,
    sheet_name: str,
    columns: list[str],
    updates: list[dict[str, Any]],
) -> None:
    """Write submitted row updates to a dedicated workbook sheet."""
    valid_updates = [u for u in updates if all(c in u for c in columns)]
    rows = [[u.get(col) for col in columns] for u in valid_updates]
    graph_client.write_table(sheet_name=sheet_name, headers=columns, rows=rows)


def reconcile_handler(
    graph_client: GraphClient,
    entity: str = "",
    period: str = "",
) -> dict[str, Any]:
    """Run the deterministic pre-pass reconciliation.

    Fetches allocRec from Excel via Graph API, runs reconcile.py, and returns
    the payload of confident changes + exceptions for the LLM.

    Args:
        graph_client: Authenticated GraphClient instance
        entity: Entity identifier (for logging)
        period: Period identifier (for logging)

    Returns:
        {
            "confident": [...],
            "exceptions": [...],
            "summary": {...}
        }
    """
    logger.info(f"[reconcile_handler] Starting pre-pass for {entity}/{period}")

    try:
        # Fetch allocRec range
        allocrec_data = graph_client.fetch_range(
            RANGES["allocRec"][0], RANGES["allocRec"][1]
        )
        logger.debug(f"Fetched allocRec: {len(allocrec_data)} rows")

        # Write to temp workbook file for reconcile.py to read
        # (or refactor reconcile.py to accept in-memory data)
        # For now, assume reconcile.py can work with Graph API data directly
        payload = reconcile.run_from_data(allocrec_data)

        logger.info(
            f"[reconcile_handler] Complete: "
            f"{payload['summary']['confident_changes']} confident, "
            f"{payload['summary']['exceptions']} exceptions"
        )
        return payload

    except Exception as e:
        logger.error(f"[reconcile_handler] Failed: {e}", exc_info=True)
        raise


def validate_handler(
    graph_client: GraphClient,
    entity: str = "",
    period: str = "",
) -> dict[str, Any]:
    """Run validation after LLM proposes changes.

    Fetches all authorized ranges (allocRec, k2Part2, k2PartX, excludes) and
    re-computes tie-out validation.

    Args:
        graph_client: Authenticated GraphClient instance
        entity: Entity identifier (for logging)
        period: Period identifier (for logging)

    Returns:
        {
            "schKToK2Part2Diff": DataFrame (or dict),
            "schKToK2PartXDiff": DataFrame (or dict),
            "grossIncome": DataFrame (or dict),
            "status": "OK" or "FAIL",
            "message": str
        }
    """
    logger.info(f"[validate_handler] Starting validation for {entity}/{period}")

    try:
        # Fetch all required ranges
        ranges_data = {}
        for range_name, (sheet, spec) in RANGES.items():
            ranges_data[range_name] = graph_client.fetch_range(sheet, spec)
            logger.debug(f"Fetched {range_name}: {len(ranges_data[range_name])} rows")

        # Run validation (refactor validate_recon to accept in-memory data)
        result = validate_recon.validate_from_data(ranges_data)

        status = "OK" if result["grossIncome"]["Difference"].sum() == 0 else "FAIL"
        logger.info(f"[validate_handler] Validation {status}")

        return {
            "schKToK2Part2Diff": result["schKToK2Part2Diff"].to_dict(orient="records"),
            "schKToK2PartXDiff": result["schKToK2PartXDiff"].to_dict(orient="records"),
            "grossIncome": result["grossIncome"].to_dict(orient="records"),
            "status": status,
            "message": (
                "Tie-out successful"
                if status == "OK"
                else "Reconciliation incomplete — tie-out failed"
            ),
        }

    except Exception as e:
        logger.error(f"[validate_handler] Failed: {e}", exc_info=True)
        raise


def writeback_handler(
    graph_client: GraphClient,
    updates: list[dict[str, Any]],
    entity: str = "",
    period: str = "",
) -> dict[str, Any]:
    """Write LLM-proposed changes back to Excel.

    Args:
        graph_client: Authenticated GraphClient instance
        updates: Either:
            1) List of {sheet, cell, value} dicts, or
            2) List of Part II row dicts using PART2_WRITEBACK_COLUMNS keys.
            3) List of Part X row dicts using PARTX_WRITEBACK_COLUMNS keys.
        entity: Entity identifier (for logging)
        period: Period identifier (for logging)

    Returns:
        {"status": "success" | "partial" | "failure", "written": N, "failed": [...]}
    """
    logger.info(f"[writeback_handler] Writing {len(updates)} changes for {entity}/{period}")

    written = []
    failed = []

    updates_to_write = []
    part2_updates = [
        u for u in updates if isinstance(u, dict) and _is_part2_row_update(u)
    ]
    partx_updates = [
        u
        for u in updates
        if isinstance(u, dict) and (not _is_part2_row_update(u)) and _is_partx_row_update(u)
    ]
    direct_updates = [
        u
        for u in updates
        if isinstance(u, dict)
        and (not _is_part2_row_update(u))
        and (not _is_partx_row_update(u))
    ]

    if part2_updates:
        try:
            part2_cell_updates, expansion_failures = _expand_part2_updates(
                graph_client, part2_updates
            )
            updates_to_write.extend(part2_cell_updates)
            failed.extend(expansion_failures)
        except Exception as e:
            logger.error(f"Failed to expand Part II row write-back payload: {e}")
            raise
        try:
            _write_updates_table(
                graph_client=graph_client,
                sheet_name=PART2_UPDATES_SHEET,
                columns=PART2_WRITEBACK_COLUMNS,
                updates=part2_updates,
            )
        except Exception as e:
            logger.error(f"Failed to write Part II updates sheet: {e}")
            failed.append({"sheet": PART2_UPDATES_SHEET, "error": str(e)})

    if partx_updates:
        try:
            partx_cell_updates, expansion_failures = _expand_partx_updates(
                graph_client, partx_updates
            )
            updates_to_write.extend(partx_cell_updates)
            failed.extend(expansion_failures)
        except Exception as e:
            logger.error(f"Failed to expand Part X row write-back payload: {e}")
            raise
        try:
            _write_updates_table(
                graph_client=graph_client,
                sheet_name=PARTX_UPDATES_SHEET,
                columns=PARTX_WRITEBACK_COLUMNS,
                updates=partx_updates,
            )
        except Exception as e:
            logger.error(f"Failed to write Part X updates sheet: {e}")
            failed.append({"sheet": PARTX_UPDATES_SHEET, "error": str(e)})

    updates_to_write.extend(direct_updates)

    for update in updates_to_write:
        if not all(k in update for k in ("sheet", "cell", "value")):
            failed.append({**update, "error": "Expected keys: sheet, cell, value"})
            continue
        try:
            graph_client.write_cell(
                update["sheet"], update["cell"], update["value"]
            )
            written.append(update)
        except Exception as e:
            logger.error(
                f"Failed to write {update.get('sheet')}!{update.get('cell')}: {e}"
            )
            failed.append({**update, "error": str(e)})

    status = "success" if not failed else ("partial" if written else "failure")
    logger.info(f"[writeback_handler] {status}: {len(written)} written, {len(failed)} failed")

    return {
        "status": status,
        "written": len(written),
        "failed": len(failed),
        "failures": failed,
    }
