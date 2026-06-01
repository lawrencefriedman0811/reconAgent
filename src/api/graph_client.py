"""Excel Graph API client for fetching and writing Excel ranges."""

from __future__ import annotations

import logging
from typing import Any, Optional

import aiohttp
import requests
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)


class GraphClient:
    """Wrapper for Microsoft Graph API Excel operations."""

    def __init__(
        self,
        client_id: str,
        tenant_id: str,
        client_secret: str,
        site_id: str,
        workbook_id: str,
    ):
        """Initialize Graph client.

        Args:
            client_id: Azure AD app ID
            tenant_id: Azure AD tenant ID
            client_secret: Client secret for authentication
            site_id: SharePoint site ID
            workbook_id: Excel workbook ID
        """
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.client_secret = client_secret
        self.site_id = site_id
        self.workbook_id = workbook_id
        self.access_token: Optional[str] = None
        self.base_url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/"
            f"drive/items/{workbook_id}/workbook"
        )

    def _get_access_token(self) -> str:
        """Get OAuth access token from Azure AD."""
        if self.access_token:
            return self.access_token

        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        payload = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        }

        try:
            resp = requests.post(token_url, data=payload, timeout=10)
            resp.raise_for_status()
            self.access_token = resp.json()["access_token"]
            return self.access_token
        except Exception as e:
            logger.error(f"Failed to get access token: {e}")
            raise

    def _headers(self) -> dict[str, str]:
        """Build request headers with auth token."""
        return {
            "Authorization": f"Bearer {self._get_access_token()}",
            "Content-Type": "application/json",
        }

    def _worksheet_url(self, sheet_name: str) -> str:
        """Build worksheet URL, escaping single quotes in names."""
        escaped = sheet_name.replace("'", "''")
        return f"{self.base_url}/worksheets('{escaped}')"

    def fetch_range(self, sheet_name: str, range_spec: str) -> list[list[Any]]:
        """Fetch a named range or cell range from Excel.

        Args:
            sheet_name: Sheet name (e.g. '(K.01) Sch K to K2 Control')
            range_spec: Range (e.g. 'C10:GH1802')

        Returns:
            List of lists (rows x columns) with cell values
        """
        url = f"{self.base_url}/worksheets('{sheet_name}')/range(address='{range_spec}')"
        try:
            resp = requests.get(url, headers=self._headers(), timeout=30)
            resp.raise_for_status()
            data = resp.json()
            return data.get("values", [])
        except Exception as e:
            logger.error(f"Failed to fetch range {sheet_name}!{range_spec}: {e}")
            raise

    def write_cell(
        self, sheet_name: str, cell_address: str, value: Any
    ) -> None:
        """Write a single cell value.

        Args:
            sheet_name: Sheet name
            cell_address: Cell address (e.g. 'M18')
            value: Value to write
        """
        url = (
            f"{self.base_url}/worksheets('{sheet_name}')/"
            f"range(address='{cell_address}')"
        )
        payload = {"values": [[value]]}
        try:
            resp = requests.patch(
                url, json=payload, headers=self._headers(), timeout=30
            )
            resp.raise_for_status()
            logger.debug(f"Wrote {sheet_name}!{cell_address} = {value}")
        except Exception as e:
            logger.error(f"Failed to write {sheet_name}!{cell_address}: {e}")
            raise

    def write_cells_batch(
        self, sheet_name: str, updates: list[dict[str, Any]]
    ) -> None:
        """Write multiple cells in a batch.

        Args:
            sheet_name: Sheet name
            updates: List of {cell_address, value} dicts
        """
        url = f"{self.base_url}/worksheets('{sheet_name}')"
        # Graph API batch updates via session calls
        for update in updates:
            self.write_cell(sheet_name, update["cell_address"], update["value"])

    def ensure_worksheet(self, sheet_name: str) -> None:
        """Ensure worksheet exists; create it when missing."""
        url = self._worksheet_url(sheet_name)
        resp = requests.get(url, headers=self._headers(), timeout=30)
        if resp.status_code == 404:
            create_url = f"{self.base_url}/worksheets/add"
            create_resp = requests.post(
                create_url,
                json={"name": sheet_name},
                headers=self._headers(),
                timeout=30,
            )
            create_resp.raise_for_status()
            logger.info(f"Created worksheet: {sheet_name}")
            return
        resp.raise_for_status()

    def clear_worksheet(self, sheet_name: str) -> None:
        """Clear used range of a worksheet if present."""
        self.ensure_worksheet(sheet_name)
        url = f"{self._worksheet_url(sheet_name)}/usedRange/clear"
        resp = requests.post(
            url,
            json={"applyTo": "All"},
            headers=self._headers(),
            timeout=30,
        )
        if resp.status_code in (400, 404):
            return
        resp.raise_for_status()

    def write_table(
        self,
        sheet_name: str,
        headers: list[str],
        rows: list[list[Any]],
    ) -> None:
        """Replace worksheet content with a header row and table rows."""
        if not headers:
            raise ValueError("headers must not be empty")

        self.clear_worksheet(sheet_name)
        values = [headers] + rows
        end_col = get_column_letter(len(headers))
        end_row = len(values)
        range_spec = f"A1:{end_col}{end_row}"
        url = f"{self._worksheet_url(sheet_name)}/range(address='{range_spec}')"
        resp = requests.patch(
            url,
            json={"values": values},
            headers=self._headers(),
            timeout=30,
        )
        resp.raise_for_status()

    async def fetch_range_async(
        self, sheet_name: str, range_spec: str
    ) -> list[list[Any]]:
        """Async version of fetch_range."""
        url = f"{self.base_url}/worksheets('{sheet_name}')/range(address='{range_spec}')"
        try:
            async with aiohttp.ClientSession() as session:
                async with session.get(
                    url, headers=self._headers(), timeout=30
                ) as resp:
                    resp.raise_for_status()
                    data = await resp.json()
                    return data.get("values", [])
        except Exception as e:
            logger.error(f"Failed to fetch range async {sheet_name}!{range_spec}: {e}")
            raise

    async def write_cells_batch_async(
        self, sheet_name: str, updates: list[dict[str, Any]]
    ) -> None:
        """Async batch write."""
        for update in updates:
            url = (
                f"{self.base_url}/worksheets('{sheet_name}')/"
                f"range(address='{update['cell_address']}')"
            )
            payload = {"values": [[update["value"]]]}
            try:
                async with aiohttp.ClientSession() as session:
                    async with session.patch(
                        url, json=payload, headers=self._headers(), timeout=30
                    ) as resp:
                        resp.raise_for_status()
            except Exception as e:
                logger.error(
                    f"Failed to write {sheet_name}!{update['cell_address']}: {e}"
                )
                raise
