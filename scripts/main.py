"""ReconAgent HTTP server entry point.

Single-purpose file-in / file-out API for the Power Automate workflow:

    POST /run   (multipart file upload)  -> modified .xlsm/.xlsx (binary)
    GET  /health                          -> {"status": "healthy"}

The flow uploads the workbook, the server runs the deterministic pre-pass,
applies confident fixes to the control sheet (cols M / N), adds an exceptions
sheet, and returns the modified workbook for the flow to write back.
"""

import io
import logging
import os

from fastapi import Depends, FastAPI, File, Header, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from config import settings
from src.api import run_reconciliation

os.makedirs(os.path.dirname(settings.logging_cfg.log_file) or ".", exist_ok=True)

logging.basicConfig(
    level=settings.logging_cfg.level,
    format=settings.logging_cfg.format,
    handlers=[
        logging.FileHandler(settings.logging_cfg.log_file),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

app = FastAPI(title="ReconAgent", version="2.0.0")

API_KEY = os.getenv("RECON_API_KEY", "")


def require_api_key(x_api_key: str = Header(default="")) -> None:
    """Require a shared-secret header when RECON_API_KEY is configured.

    No-op when RECON_API_KEY is unset, so local dev is unaffected. Set it
    before exposing the server through a public tunnel.
    """
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="Invalid or missing API key")


@app.post("/run")
async def run_endpoint(
    file: UploadFile = File(...),
    _: None = Depends(require_api_key),
):
    """Run reconciliation on the uploaded workbook and return the modified file."""
    filename = file.filename or "workbook.xlsm"
    try:
        workbook_bytes = await file.read()
        if not workbook_bytes:
            raise HTTPException(status_code=400, detail="Empty file upload")

        result = run_reconciliation(workbook_bytes, filename=filename)
        summary = result["summary"]
        logger.info(
            "[/run] %s — applied %s confident, %s exceptions",
            filename,
            summary.get("confident_applied"),
            summary.get("exceptions"),
        )

        is_macro = filename.lower().endswith(".xlsm")
        media_type = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if is_macro
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        return StreamingResponse(
            io.BytesIO(result["workbook_bytes"]),
            media_type=media_type,
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"',
                "X-Recon-Confident-Applied": str(summary.get("confident_applied", 0)),
                "X-Recon-Exceptions": str(summary.get("exceptions", 0)),
                "X-Recon-Data-Rows": str(summary.get("data_rows", 0)),
            },
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error("[/run] Failed for %s: %s", filename, e, exc_info=True)
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
