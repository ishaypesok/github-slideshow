"""Write forecasting report data to Google Sheets."""

import logging
from datetime import date, timedelta
from typing import Any

from google.oauth2.service_account import Credentials
import gspread

from firefly_client import FireflyClient
from report_generator import (
    _resolve_account_ids,
    _get_filtered_expenses,
    _aggregate_by_category,
    _aggregate_total,
)

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _linear_forecast(values: list[float], periods: int) -> list[float]:
    """Simple linear regression forecast."""
    import numpy as np

    if not values or len(values) < 2:
        avg = values[0] if values else 0.0
        return [avg] * periods
    n = len(values)
    x = np.arange(n)
    y = np.array(values)
    coeffs = np.polyfit(x, y, 1)
    return [float(coeffs[0] * i + coeffs[1]) for i in range(n, n + periods)]


def _get_monthly_ranges(months_back: int) -> list[tuple[date, date, str]]:
    """Generate (start, end, label) for each of the past N months."""
    today = date.today()
    ranges = []
    for i in range(months_back, 0, -1):
        month = today.month - i
        year = today.year
        while month <= 0:
            month += 12
            year -= 1
        start = date(year, month, 1)
        if month == 12:
            end = date(year + 1, 1, 1) - timedelta(days=1)
        else:
            end = date(year, month + 1, 1) - timedelta(days=1)
        label = start.strftime("%b %Y")
        ranges.append((start, end, label))
    return ranges


def _forecast_label(monthly_ranges: list, offset: int) -> str:
    if monthly_ranges:
        last_start = monthly_ranges[-1][0]
        fm = last_start.month + offset + 1
        fy = last_start.year
        while fm > 12:
            fm -= 12
            fy += 1
        return date(fy, fm, 1).strftime("%b %Y") + " (F)"
    return f"Forecast {offset + 1}"


def _fmt_currency(val: float) -> str:
    return f"{val:,.2f}"


class SheetsReportWriter:
    """Writes Firefly III forecasting report to Google Sheets."""

    def __init__(
        self,
        client: FireflyClient,
        credentials_file: str,
        months_back: int = 12,
        forecast_months: int = 3,
        filter_config: dict | None = None,
    ):
        self.client = client
        self.months_back = months_back
        self.forecast_months = forecast_months
        self.filter_config = filter_config or {}
        self.use_filters = bool(
            self.filter_config.get("account_filters") or self.filter_config.get("recurring_only")
        )
        self._account_ids: list[str] | None = None
        if self.filter_config.get("account_filters"):
            self._account_ids = _resolve_account_ids(client, self.filter_config["account_filters"])
        creds = Credentials.from_service_account_file(credentials_file, scopes=SCOPES)
        self.gc = gspread.authorize(creds)

    def generate(self, spreadsheet_name: str, share_with: str | None = None) -> str:
        """Create a new Google Sheet with the forecasting report.

        Returns the spreadsheet URL.
        """
        logger.info("Creating Google Sheet: %s", spreadsheet_name)
        spreadsheet = self.gc.create(spreadsheet_name)

        if share_with:
            spreadsheet.share(share_with, perm_type="user", role="writer")
            logger.info("Shared with %s", share_with)

        # Build sheets (delete default Sheet1 after)
        self._build_summary(spreadsheet)
        self._build_expense_by_category(spreadsheet)
        self._build_income_by_category(spreadsheet)
        self._build_transactions(spreadsheet)
        self._build_budget_overview(spreadsheet)
        self._build_accounts(spreadsheet)

        # Remove the default empty sheet
        default_sheet = spreadsheet.sheet1
        if default_sheet.title == "Sheet1":
            spreadsheet.del_worksheet(default_sheet)

        url = spreadsheet.url
        logger.info("Report available at: %s", url)
        return url

    def write_to_existing(self, spreadsheet_url: str) -> str:
        """Write to an existing Google Sheet (clears and replaces content).

        Returns the spreadsheet URL.
        """
        spreadsheet = self.gc.open_by_url(spreadsheet_url)
        logger.info("Writing to existing sheet: %s", spreadsheet.title)

        # Clear all existing worksheets and rebuild
        existing = spreadsheet.worksheets()
        self._build_summary(spreadsheet)
        self._build_expense_by_category(spreadsheet)
        self._build_income_by_category(spreadsheet)
        self._build_transactions(spreadsheet)
        self._build_budget_overview(spreadsheet)
        self._build_accounts(spreadsheet)

        # Remove old sheets
        for ws in existing:
            try:
                spreadsheet.del_worksheet(ws)
            except Exception:
                pass  # Can't delete if it's the only sheet left

        return spreadsheet.url

    def _build_summary(self, spreadsheet):
        ws = spreadsheet.add_worksheet("Summary", rows=100, cols=10)
        monthly_ranges = _get_monthly_ranges(self.months_back)

        headers = ["Month", "Total Income", "Total Expenses", "Net (Income - Expenses)"]
        rows = [headers]

        incomes = []
        expenses = []

        for start, end, label in monthly_ranges:
            try:
                inc_data = self.client.get_insight_income_total(start, end)
                inc_total = sum(abs(float(i.get("difference_float", 0))) for i in inc_data)
            except Exception:
                inc_total = 0.0

            if self.use_filters:
                filtered_txs = _get_filtered_expenses(
                    self.client, start, end,
                    account_ids=self._account_ids,
                    recurring_only=self.filter_config.get("recurring_only", False),
                )
                exp_total = _aggregate_total(filtered_txs)
            else:
                try:
                    exp_data = self.client.get_insight_expense_total(start, end)
                    exp_total = sum(abs(float(i.get("difference_float", 0))) for i in exp_data)
                except Exception:
                    exp_total = 0.0

            incomes.append(inc_total)
            expenses.append(exp_total)
            rows.append([label, _fmt_currency(inc_total), _fmt_currency(exp_total), _fmt_currency(inc_total - exp_total)])

        # Forecast
        fc_inc = _linear_forecast(incomes, self.forecast_months)
        fc_exp = _linear_forecast(expenses, self.forecast_months)
        for i in range(self.forecast_months):
            label = _forecast_label(monthly_ranges, i)
            fi = max(0, fc_inc[i])
            fe = max(0, fc_exp[i])
            rows.append([label, _fmt_currency(fi), _fmt_currency(fe), _fmt_currency(fi - fe)])

        ws.update(range_name="A1", values=rows)
        ws.format("A1:D1", {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59}, "horizontalAlignment": "CENTER", "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}})

    def _build_expense_by_category(self, spreadsheet):
        ws = spreadsheet.add_worksheet("Expenses by Category", rows=200, cols=50)
        monthly_ranges = _get_monthly_ranges(self.months_back)

        category_monthly: dict[str, list[float]] = {}
        month_labels = []

        for start, end, label in monthly_ranges:
            month_labels.append(label)

            if self.use_filters:
                filtered_txs = _get_filtered_expenses(
                    self.client, start, end,
                    account_ids=self._account_ids,
                    recurring_only=self.filter_config.get("recurring_only", False),
                )
                cat_amounts = _aggregate_by_category(filtered_txs)
                seen = set()
                for name, amount in cat_amounts.items():
                    if name not in category_monthly:
                        category_monthly[name] = [0.0] * len(month_labels[:-1])
                    category_monthly[name].append(amount)
                    seen.add(name)
                for name in category_monthly:
                    if name not in seen:
                        category_monthly[name].append(0.0)
            else:
                try:
                    data = self.client.get_insight_expense_category(start, end)
                except Exception:
                    data = []
                seen = set()
                for item in data:
                    name = item.get("name", "Uncategorized")
                    amount = abs(float(item.get("difference_float", 0)))
                    if name not in category_monthly:
                        category_monthly[name] = [0.0] * len(month_labels[:-1])
                    category_monthly[name].append(amount)
                    seen.add(name)
                for name in category_monthly:
                    if name not in seen:
                        category_monthly[name].append(0.0)

        if not category_monthly:
            ws.update(range_name="A1", values=[["No expense data found."]])
            return

        categories = sorted(category_monthly.keys())
        forecast_labels = [_forecast_label(monthly_ranges, i) for i in range(self.forecast_months)]
        headers = ["Category"] + month_labels + forecast_labels
        rows = [headers]

        for cat in categories:
            values = category_monthly[cat]
            forecasts = [max(0, v) for v in _linear_forecast(values, self.forecast_months)]
            row = [cat] + [_fmt_currency(v) for v in values] + [_fmt_currency(f) for f in forecasts]
            rows.append(row)

        ws.update(range_name="A1", values=rows)
        ws.format("A1:{}1".format(chr(64 + len(headers))), {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59}, "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}})

    def _build_income_by_category(self, spreadsheet):
        ws = spreadsheet.add_worksheet("Income by Category", rows=200, cols=50)
        monthly_ranges = _get_monthly_ranges(self.months_back)

        category_monthly: dict[str, list[float]] = {}
        month_labels = []

        for start, end, label in monthly_ranges:
            month_labels.append(label)
            try:
                data = self.client.get_insight_income_category(start, end)
            except Exception:
                data = []
            seen = set()
            for item in data:
                name = item.get("name", "Uncategorized")
                amount = abs(float(item.get("difference_float", 0)))
                if name not in category_monthly:
                    category_monthly[name] = [0.0] * len(month_labels[:-1])
                category_monthly[name].append(amount)
                seen.add(name)
            for name in category_monthly:
                if name not in seen:
                    category_monthly[name].append(0.0)

        if not category_monthly:
            ws.update(range_name="A1", values=[["No income data found."]])
            return

        categories = sorted(category_monthly.keys())
        forecast_labels = [_forecast_label(monthly_ranges, i) for i in range(self.forecast_months)]
        headers = ["Category"] + month_labels + forecast_labels
        rows = [headers]

        for cat in categories:
            values = category_monthly[cat]
            forecasts = [max(0, v) for v in _linear_forecast(values, self.forecast_months)]
            row = [cat] + [_fmt_currency(v) for v in values] + [_fmt_currency(f) for f in forecasts]
            rows.append(row)

        ws.update(range_name="A1", values=rows)
        ws.format("A1:{}1".format(chr(64 + len(headers))), {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59}, "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}})

    def _build_transactions(self, spreadsheet):
        ws = spreadsheet.add_worksheet("Transactions", rows=2000, cols=10)
        monthly_ranges = _get_monthly_ranges(self.months_back)

        if monthly_ranges:
            start = monthly_ranges[0][0]
            end = monthly_ranges[-1][1]
        else:
            start = date.today() - timedelta(days=365)
            end = date.today()

        recurring_only = self.filter_config.get("recurring_only", False)

        all_transactions = []
        if self._account_ids:
            for aid in self._account_ids:
                try:
                    tx_groups = self.client.get_account_transactions(aid, start, end, tx_type="withdrawal")
                except Exception:
                    continue
                for group in tx_groups:
                    attrs = group.get("attributes", {})
                    for tx in attrs.get("transactions", []):
                        all_transactions.append(tx)
        else:
            try:
                tx_groups = self.client.get_transactions(start, end, tx_type="withdrawal")
            except Exception:
                tx_groups = []
            for group in tx_groups:
                attrs = group.get("attributes", {})
                for tx in attrs.get("transactions", []):
                    all_transactions.append(tx)

        if recurring_only:
            all_transactions = [
                tx for tx in all_transactions
                if tx.get("bill_id") or tx.get("recurrence_id")
            ]

        all_transactions.sort(key=lambda tx: tx.get("date", ""), reverse=True)

        headers = ["Date", "Amount", "Source Account", "Destination", "Category", "Description", "Bill/Recurring"]
        rows = [headers]

        for tx in all_transactions:
            tx_date = tx.get("date", "")[:10]
            amount = _fmt_currency(abs(float(tx.get("amount", 0))))
            source = tx.get("source_name", "")
            destination = tx.get("destination_name", "")
            category = tx.get("category_name") or ""
            description = tx.get("description", "")
            is_bill = "Yes" if tx.get("bill_id") else ""
            if tx.get("recurrence_id"):
                is_bill = "Yes (recurring)"
            rows.append([tx_date, amount, source, destination, category, description, is_bill])

        ws.update(range_name="A1", values=rows)
        ws.format("A1:G1", {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59}, "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}})

    def _build_budget_overview(self, spreadsheet):
        ws = spreadsheet.add_worksheet("Budget Overview", rows=100, cols=10)
        today = date.today()
        start = date(today.year, today.month, 1)
        if today.month == 12:
            end = date(today.year + 1, 1, 1) - timedelta(days=1)
        else:
            end = date(today.year, today.month + 1, 1) - timedelta(days=1)

        headers = ["Budget Name", "Budgeted Amount", "Spent", "Remaining", "% Used"]
        rows = [headers]

        try:
            budgets = self.client.get_budgets(start, end)
        except Exception:
            budgets = []

        for budget in budgets:
            attrs = budget.get("attributes", {})
            name = attrs.get("name", "Unknown")
            auto_amount = float(attrs.get("auto_budget_amount", 0) or 0)
            spent_list = attrs.get("spent", [])
            spent = sum(abs(float(s.get("amount", 0))) for s in spent_list) if spent_list else 0.0
            remaining = auto_amount - spent
            pct = f"{(spent / auto_amount * 100):.1f}%" if auto_amount > 0 else "0.0%"
            rows.append([name, _fmt_currency(auto_amount), _fmt_currency(spent), _fmt_currency(remaining), pct])

        ws.update(range_name="A1", values=rows)
        ws.format("A1:E1", {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59}, "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}})

    def _build_accounts(self, spreadsheet):
        ws = spreadsheet.add_worksheet("Accounts", rows=100, cols=10)
        headers = ["Account Name", "Type", "Currency", "Current Balance", "Active"]
        rows = [headers]

        try:
            accounts = self.client.get_accounts("asset")
        except Exception:
            accounts = []

        for acct in accounts:
            attrs = acct.get("attributes", {})
            rows.append([
                attrs.get("name", ""),
                attrs.get("account_role", attrs.get("type", "")),
                attrs.get("currency_code", ""),
                _fmt_currency(float(attrs.get("current_balance", 0) or 0)),
                "Yes" if attrs.get("active") else "No",
            ])

        ws.update(range_name="A1", values=rows)
        ws.format("A1:E1", {"textFormat": {"bold": True}, "backgroundColor": {"red": 0.18, "green": 0.33, "blue": 0.59}, "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}})
