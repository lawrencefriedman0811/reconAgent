"""Configuration for ReconAgent."""

import os
from dataclasses import dataclass


@dataclass
class LoggingConfig:
    """Logging configuration."""

    level: str = os.getenv("LOG_LEVEL", "INFO")
    format: str = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    log_file: str = os.getenv("LOG_FILE", "logs/recon_agent.log")


# Load configs from environment
logging_cfg = LoggingConfig()
