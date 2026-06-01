"""Configuration for ReconAgent."""

import os
from dataclasses import dataclass


@dataclass
class GraphConfig:
    """Microsoft Graph API configuration."""

    client_id: str = os.getenv("GRAPH_CLIENT_ID", "")
    tenant_id: str = os.getenv("GRAPH_TENANT_ID", "")
    client_secret: str = os.getenv("GRAPH_CLIENT_SECRET", "")
    site_id: str = os.getenv("GRAPH_SITE_ID", "")
    workbook_id: str = os.getenv("GRAPH_WORKBOOK_ID", "")


@dataclass
class LoggingConfig:
    """Logging configuration."""

    level: str = os.getenv("LOG_LEVEL", "INFO")
    format: str = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    log_file: str = os.getenv("LOG_FILE", "logs/recon_agent.log")


# Load configs from environment
graph = GraphConfig()
logging_cfg = LoggingConfig()
