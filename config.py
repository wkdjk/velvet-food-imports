"""
Configuration loader for Velvet Food Imports.

Reads all secrets and settings from environment variables only.
No literal credentials appear here or anywhere else in the codebase.

Usage:
    from config import cfg
    print(cfg.sheets_id)
"""

import os
from pathlib import Path

# Load .env if present (local development only; CI reads from GitHub Secrets)
try:
    from dotenv import load_dotenv
    _env_path = Path(__file__).parent / ".env"
    load_dotenv(dotenv_path=_env_path, override=False)
except ImportError:
    pass  # python-dotenv not installed; assume environment is pre-populated


class _Config:
    """Validated configuration. Raises on missing required variables."""

    # ------------------------------------------------------------------ #
    # Required secrets                                                     #
    # ------------------------------------------------------------------ #

    @property
    def service_account_json_str(self) -> str:
        """Raw JSON string for the Google service account.

        Must be a single-line minified JSON value in .env.
        See .env.example for the required format.
        """
        value = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        if not value:
            raise EnvironmentError(
                "GOOGLE_SERVICE_ACCOUNT_JSON is not set. "
                "Copy .env.example to .env and fill in the service account JSON "
                "(must be on ONE line — see .env.example for instructions)."
            )
        return value

    @property
    def sheets_id(self) -> str:
        """Google Sheets ID for the VFI master data file."""
        value = os.environ.get("GOOGLE_SHEETS_ID", "").strip()
        if not value:
            raise EnvironmentError(
                "GOOGLE_SHEETS_ID is not set. "
                "Add it to .env or register it as a GitHub Secret."
            )
        return value

    # ------------------------------------------------------------------ #
    # Optional settings with sensible defaults                             #
    # ------------------------------------------------------------------ #

    @property
    def historical_xlsx_path(self) -> Path:
        """Absolute path to the Commander's historical xlsx file."""
        raw = os.environ.get(
            "VFI_HISTORICAL_XLSX",
            "data/historical/Korean food imports list for deer velvet.xlsx",
        )
        return Path(raw)

    @property
    def mfds_dir(self) -> Path:
        """Directory where MFDS weekly xlsx files are placed."""
        raw = os.environ.get("VFI_MFDS_DIR", "data/mfds")
        return Path(raw)


# Singleton instance — import this everywhere
cfg = _Config()
