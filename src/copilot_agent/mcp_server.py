"""MCP server that exposes ReconAgent HTTP operations to Copilot."""

from __future__ import annotations

import os
from typing import Any

import requests
from mcp.server.fastmcp import FastMCP

DEFAULT_API_BASE_URL = os.getenv("RECON_API_BASE_URL", "http://localhost:8000")
REQUEST_TIMEOUT_SECONDS = 120

mcp = FastMCP("ReconAgent")


def _clean_base_url(base_url: str | None) -> str:
    if base_url and base_url.strip():
        return base_url.rstrip("/")
    return DEFAULT_API_BASE_URL.rstrip("/")


def _post_json(base_url: str, path: str, payload: dict[str, Any]) -> dict[str, Any]:
    url = f"{_clean_base_url(base_url)}{path}"
    response = requests.post(url, json=payload, timeout=REQUEST_TIMEOUT_SECONDS)
    response.raise_for_status()
    data = response.json()
    if not isinstance(data, dict):
        raise RuntimeError(f"Unexpected response type from {url}: {type(data)}")
    return data


@mcp.tool()
def run_reconciliation(
    entity: str = "",
    period: str = "",
    base_url: str = DEFAULT_API_BASE_URL,
) -> dict[str, Any]:
    """Kick off reconciliation by calling POST /reconcile."""
    return _post_json(
        base_url=base_url,
        path="/reconcile",
        payload={"entity": entity, "period": period},
    )


@mcp.tool()
def writeback_reconciliation(
    updates: list[dict[str, Any]],
    entity: str = "",
    period: str = "",
    base_url: str = DEFAULT_API_BASE_URL,
) -> dict[str, Any]:
    """Write reconciliation updates by calling POST /writeback."""
    return _post_json(
        base_url=base_url,
        path="/writeback",
        payload={"entity": entity, "period": period, "updates": updates},
    )


@mcp.tool()
def validate_reconciliation(
    entity: str = "",
    period: str = "",
    base_url: str = DEFAULT_API_BASE_URL,
) -> dict[str, Any]:
    """Validate reconciliation by calling POST /validate."""
    return _post_json(
        base_url=base_url,
        path="/validate",
        payload={"entity": entity, "period": period},
    )


if __name__ == "__main__":
    mcp.run()
