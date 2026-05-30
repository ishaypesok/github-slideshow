"""Firefly III API client for fetching financial data."""

import logging
from datetime import date, timedelta
from typing import Any
from urllib.parse import urljoin

import requests

logger = logging.getLogger(__name__)


class FireflyClient:
    """Client for the Firefly III REST API v1."""

    def __init__(self, base_url: str, access_token: str):
        self.base_url = base_url.rstrip("/")
        self.api_url = f"{self.base_url}/api/v1"
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/json",
                "Content-Type": "application/json",
            }
        )

    def _get(self, endpoint: str, params: dict | None = None) -> dict:
        url = f"{self.api_url}/{endpoint.lstrip('/')}"
        resp = self.session.get(url, params=params)
        resp.raise_for_status()
        return resp.json()

    def _get_all_pages(self, endpoint: str, params: dict | None = None) -> list[dict]:
        """Fetch all pages of a paginated endpoint."""
        params = dict(params or {})
        params["page"] = 1
        all_data = []
        while True:
            result = self._get(endpoint, params)
            data = result.get("data", [])
            if not data:
                break
            all_data.extend(data)
            meta = result.get("meta", {}).get("pagination", {})
            total_pages = meta.get("total_pages", 1)
            if params["page"] >= total_pages:
                break
            params["page"] += 1
        return all_data

    def get_about(self) -> dict:
        """Get Firefly III instance info."""
        return self._get("about")

    def get_accounts(self, account_type: str = "asset") -> list[dict]:
        """List accounts filtered by type."""
        return self._get_all_pages("accounts", {"type": account_type})

    def get_transactions(
        self,
        start: date | None = None,
        end: date | None = None,
        tx_type: str = "all",
    ) -> list[dict]:
        """List transactions within a date range."""
        params: dict[str, Any] = {"type": tx_type}
        if start:
            params["start"] = start.isoformat()
        if end:
            params["end"] = end.isoformat()
        return self._get_all_pages("transactions", params)

    def get_budgets(
        self, start: date | None = None, end: date | None = None
    ) -> list[dict]:
        """List budgets with optional spent amounts for a date range."""
        params: dict[str, Any] = {}
        if start:
            params["start"] = start.isoformat()
        if end:
            params["end"] = end.isoformat()
        return self._get_all_pages("budgets", params)

    def get_categories(
        self, start: date | None = None, end: date | None = None
    ) -> list[dict]:
        """List categories with optional spent/earned amounts."""
        params: dict[str, Any] = {}
        if start:
            params["start"] = start.isoformat()
        if end:
            params["end"] = end.isoformat()
        return self._get_all_pages("categories", params)

    def get_insight_expense_category(
        self, start: date, end: date
    ) -> list[dict]:
        """Get expense insights grouped by category."""
        return self._get(
            "insight/expense/category",
            {"start": start.isoformat(), "end": end.isoformat()},
        )

    def get_insight_income_category(
        self, start: date, end: date
    ) -> list[dict]:
        """Get income insights grouped by category."""
        return self._get(
            "insight/income/category",
            {"start": start.isoformat(), "end": end.isoformat()},
        )

    def get_insight_expense_total(self, start: date, end: date) -> list[dict]:
        """Get total expenses for a period."""
        return self._get(
            "insight/expense/total",
            {"start": start.isoformat(), "end": end.isoformat()},
        )

    def get_insight_income_total(self, start: date, end: date) -> list[dict]:
        """Get total income for a period."""
        return self._get(
            "insight/income/total",
            {"start": start.isoformat(), "end": end.isoformat()},
        )

    def get_insight_expense_budget(
        self, start: date, end: date
    ) -> list[dict]:
        """Get expense insights grouped by budget."""
        return self._get(
            "insight/expense/budget",
            {"start": start.isoformat(), "end": end.isoformat()},
        )
