"""ReconAgent HTTP server entry point."""

import asyncio
import logging
import os
from typing import Any

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

from config import settings
from src.api import GraphClient, reconcile_handler, validate_handler, writeback_handler

# Configure logging
logging.basicConfig(
    level=settings.logging_cfg.level,
    format=settings.logging_cfg.format,
    handlers=[
        logging.FileHandler(settings.logging_cfg.log_file),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

app = FastAPI(title="ReconAgent", version="1.0.0")


class ReconcileRequest(BaseModel):
    """Reconciliation request payload."""

    entity: str = ""
    period: str = ""


class ValidateRequest(BaseModel):
    """Validation request payload."""

    entity: str = ""
    period: str = ""


class WritebackRequest(BaseModel):
    """Write-back request payload."""

    updates: list[dict[str, Any]]
    entity: str = ""
    period: str = ""


@app.post("/reconcile")
async def reconcile_endpoint(req: ReconcileRequest):
    """Run deterministic pre-pass reconciliation."""
    try:
        client = GraphClient(
            client_id=settings.graph.client_id,
            tenant_id=settings.graph.tenant_id,
            client_secret=settings.graph.client_secret,
            site_id=settings.graph.site_id,
            workbook_id=settings.graph.workbook_id,
        )

        payload = reconcile_handler(
            client, entity=req.entity, period=req.period
        )
        return {"success": True, "data": payload}

    except Exception as e:
        logger.error(f"Reconciliation failed: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/validate")
async def validate_endpoint(req: ValidateRequest):
    """Run validation after LLM proposes changes."""
    try:
        client = GraphClient(
            client_id=settings.graph.client_id,
            tenant_id=settings.graph.tenant_id,
            client_secret=settings.graph.client_secret,
            site_id=settings.graph.site_id,
            workbook_id=settings.graph.workbook_id,
        )

        result = validate_handler(client, entity=req.entity, period=req.period)
        return {"success": True, "data": result}

    except Exception as e:
        logger.error(f"Validation failed: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/writeback")
async def writeback_endpoint(req: WritebackRequest):
    """Write LLM-proposed changes back to Excel."""
    try:
        client = GraphClient(
            client_id=settings.graph.client_id,
            tenant_id=settings.graph.tenant_id,
            client_secret=settings.graph.client_secret,
            site_id=settings.graph.site_id,
            workbook_id=settings.graph.workbook_id,
        )

        result = writeback_handler(
            client, req.updates, entity=req.entity, period=req.period
        )
        return {"success": result["status"] == "success", "data": result}

    except Exception as e:
        logger.error(f"Write-back failed: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/health")
async def health_check():
    """Health check endpoint."""
    return {"status": "healthy"}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=int(os.getenv("PORT", 8000)),
    )
