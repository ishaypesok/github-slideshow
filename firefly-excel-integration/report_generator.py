"""Generate Excel forecasting reports from Firefly III data."""

import logging
from datetime import date, timedelta
from typing import Any

import numpy as np
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from firefly_client import FireflyClient

logger = logging.getLogger(__name__)

HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
SUBHEADER_FILL = PatternFill(
    start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"
)
FORECAST_FILL = PatternFill(
    start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
)
CURRENCY_FORMAT = '#,##0.00'
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def _style_header_row(ws, row: int, max_col: int):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER


def _style_data_cell(ws, row: int, col: int, is_currency: bool = False):
    cell = ws.cell(row=row, column=col)
    cell.border = THIN_BORDER
    if is_currency:
        cell.number_format = CURRENCY_FORMAT
    return cell


def _linear_forecast(values: list[float], periods: int) -> list[float]:
    """Simple linear regression forecast."""
    if not values or len(values) < 2:
        avg = values[0] if values else 0.0
        return [avg] * periods
    n = len(values)
    x = np.arange(n)
    y = np.array(values)
    coeffs = np.polyfit(x, y, 1)
    forecasts = []
    for i in range(n, n + periods):
        forecasts.append(float(coeffs[0] * i + coeffs[1]))
    return forecasts


def _get_monthly_ranges(months_back: int) -> list[tuple[date, date, str]]:
    """Generate (start, end, label) for each of the past N months."""
    today = date.today()
    ranges = []
    for i in range(months_back, 0, -1):
        # First day of month i months ago
        month = today.month - i
        year = today.year
        while month <= 0:
            month += 12
            year -= 1
        start = date(year, month, 1)
        # Last day of that month
        if month == 12:
            end = date(year + 1, 1, 1) - timedelta(days=1)
        else:
            end = date(year, month + 1, 1) - timedelta(days=1)
        label = start.strftime("%b %Y")
        ranges.append((start, end, label))
    return ranges


def _match_account(account_name: str, account_id: str, filters: list[str]) -> bool:
    """Check if an account matches any of the filter strings (partial match)."""
    for f in filters:
        f_lower = f.lower()
        if f_lower in account_name.lower() or f_lower == str(account_id):
            return True
    return False


def _resolve_account_ids(client: FireflyClient, filters: list[str]) -> list[str]:
    """Find account IDs matching the filter strings."""
    matched_ids = []
    for acct_type in ["asset", "expense"]:
        try:
            accounts = client.get_accounts(acct_type)
        except Exception:
            continue
        for acct in accounts:
            attrs = acct.get("attributes", {})
            name = attrs.get("name", "")
            aid = str(acct.get("id", ""))
            if _match_account(name, aid, filters):
                matched_ids.append(aid)
                logger.info("Matched account: %s (ID: %s)", name, aid)
    return matched_ids


def _get_filtered_expenses(
    client: FireflyClient,
    start: date,
    end: date,
    account_ids: list[str] | None = None,
    recurring_only: bool = False,
) -> list[dict]:
    """Get transactions filtered by accounts and/or recurring status.

    Returns a list of individual transaction dicts with keys:
    amount, category_name, source_name, destination_name, bill_id, etc.
    """
    transactions = []

    if account_ids:
        for aid in account_ids:
            try:
                tx_groups = client.get_account_transactions(aid, start, end, tx_type="withdrawal")
            except Exception:
                continue
            for group in tx_groups:
                attrs = group.get("attributes", {})
                for tx in attrs.get("transactions", []):
                    transactions.append(tx)
    else:
        try:
            tx_groups = client.get_transactions(start, end, tx_type="withdrawal")
        except Exception:
            tx_groups = []
        for group in tx_groups:
            attrs = group.get("attributes", {})
            for tx in attrs.get("transactions", []):
                transactions.append(tx)

    if recurring_only:
        transactions = [
            tx for tx in transactions
            if tx.get("bill_id") or tx.get("recurrence_id")
        ]

    return transactions


def _aggregate_by_category(transactions: list[dict]) -> dict[str, float]:
    """Sum transaction amounts by category."""
    by_cat: dict[str, float] = {}
    for tx in transactions:
        cat = tx.get("category_name") or "Uncategorized"
        amount = abs(float(tx.get("amount", 0)))
        by_cat[cat] = by_cat.get(cat, 0) + amount
    return by_cat


def _aggregate_total(transactions: list[dict]) -> float:
    """Sum all transaction amounts."""
    return sum(abs(float(tx.get("amount", 0))) for tx in transactions)


class ReportGenerator:
    """Generates Excel reports from Firefly III data."""

    def __init__(
        self,
        client: FireflyClient,
        months_back: int = 12,
        forecast_months: int = 3,
        filter_config: dict | None = None,
    ):
        self.client = client
        self.months_back = months_back
        self.forecast_months = forecast_months
        self.wb = Workbook()
        self.filter_config = filter_config or {}
        self.use_filters = bool(
            self.filter_config.get("account_filters") or self.filter_config.get("recurring_only")
        )
        self._account_ids: list[str] | None = None
        if self.filter_config.get("account_filters"):
            self._account_ids = _resolve_account_ids(client, self.filter_config["account_filters"])

    def generate(self, output_path: str):
        """Generate the full forecasting report."""
        logger.info("Generating forecasting report...")

        # Remove default sheet
        self.wb.remove(self.wb.active)

        self._build_summary_sheet()
        self._build_expense_by_category_sheet()
        self._build_income_sheet()
        self._build_budget_sheet()
        self._build_accounts_sheet()

        self.wb.save(output_path)
        logger.info("Report saved to %s", output_path)

    def _build_summary_sheet(self):
        """Monthly totals + forecast for income, expenses, net."""
        ws = self.wb.create_sheet("Summary")
        monthly_ranges = _get_monthly_ranges(self.months_back)

        # Headers
        ws.cell(row=1, column=1, value="Month")
        ws.cell(row=1, column=2, value="Total Income")
        ws.cell(row=1, column=3, value="Total Expenses")
        ws.cell(row=1, column=4, value="Net (Income - Expenses)")
        _style_header_row(ws, 1, 4)

        incomes = []
        expenses = []

        for idx, (start, end, label) in enumerate(monthly_ranges):
            row = idx + 2
            ws.cell(row=row, column=1, value=label)
            _style_data_cell(ws, row, 1)

            try:
                inc_data = self.client.get_insight_income_total(start, end)
                inc_total = sum(
                    abs(float(item.get("difference_float", 0)))
                    for item in inc_data
                )
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
                    exp_total = sum(
                        abs(float(item.get("difference_float", 0)))
                        for item in exp_data
                    )
                except Exception:
                    exp_total = 0.0

            incomes.append(inc_total)
            expenses.append(exp_total)

            _style_data_cell(ws, row, 2, is_currency=True).value = inc_total
            _style_data_cell(ws, row, 3, is_currency=True).value = exp_total
            _style_data_cell(ws, row, 4, is_currency=True).value = inc_total - exp_total

        # Forecast
        forecast_incomes = _linear_forecast(incomes, self.forecast_months)
        forecast_expenses = _linear_forecast(expenses, self.forecast_months)

        last_range = monthly_ranges[-1] if monthly_ranges else None
        for i in range(self.forecast_months):
            row = len(monthly_ranges) + 2 + i
            if last_range:
                forecast_month = last_range[0].month + i + 1
                forecast_year = last_range[0].year
                while forecast_month > 12:
                    forecast_month -= 12
                    forecast_year += 1
                label = date(forecast_year, forecast_month, 1).strftime("%b %Y") + " (F)"
            else:
                label = f"Forecast {i + 1}"

            ws.cell(row=row, column=1, value=label)
            _style_data_cell(ws, row, 1)

            fc_inc = max(0, forecast_incomes[i])
            fc_exp = max(0, forecast_expenses[i])

            cell_inc = _style_data_cell(ws, row, 2, is_currency=True)
            cell_inc.value = fc_inc
            cell_inc.fill = FORECAST_FILL

            cell_exp = _style_data_cell(ws, row, 3, is_currency=True)
            cell_exp.value = fc_exp
            cell_exp.fill = FORECAST_FILL

            cell_net = _style_data_cell(ws, row, 4, is_currency=True)
            cell_net.value = fc_inc - fc_exp
            cell_net.fill = FORECAST_FILL

        # Chart
        total_rows = len(monthly_ranges) + self.forecast_months + 1
        chart = LineChart()
        chart.title = "Income vs Expenses (with Forecast)"
        chart.y_axis.title = "Amount"
        chart.x_axis.title = "Month"
        chart.style = 10
        chart.width = 25
        chart.height = 15

        cats = Reference(ws, min_col=1, min_row=2, max_row=total_rows)
        income_ref = Reference(ws, min_col=2, min_row=1, max_row=total_rows)
        expense_ref = Reference(ws, min_col=3, min_row=1, max_row=total_rows)
        net_ref = Reference(ws, min_col=4, min_row=1, max_row=total_rows)

        chart.add_data(income_ref, titles_from_data=True)
        chart.add_data(expense_ref, titles_from_data=True)
        chart.add_data(net_ref, titles_from_data=True)
        chart.set_categories(cats)

        ws.add_chart(chart, "F2")

        # Auto-width
        for col in range(1, 5):
            ws.column_dimensions[get_column_letter(col)].width = 18

    def _build_expense_by_category_sheet(self):
        """Monthly expenses per category + forecast."""
        ws = self.wb.create_sheet("Expenses by Category")
        monthly_ranges = _get_monthly_ranges(self.months_back)

        # Collect all category data
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
                # Fill zeros for categories not in this month
                for name in category_monthly:
                    if name not in seen:
                        category_monthly[name].append(0.0)

        if not category_monthly:
            ws.cell(row=1, column=1, value="No expense data found.")
            return

        categories = sorted(category_monthly.keys())

        # Headers
        ws.cell(row=1, column=1, value="Category")
        for i, label in enumerate(month_labels):
            ws.cell(row=1, column=i + 2, value=label)
        for i in range(self.forecast_months):
            if monthly_ranges:
                last_start = monthly_ranges[-1][0]
                fm = last_start.month + i + 1
                fy = last_start.year
                while fm > 12:
                    fm -= 12
                    fy += 1
                flabel = date(fy, fm, 1).strftime("%b %Y") + " (F)"
            else:
                flabel = f"F{i+1}"
            ws.cell(row=1, column=len(month_labels) + i + 2, value=flabel)

        total_cols = len(month_labels) + self.forecast_months + 1
        _style_header_row(ws, 1, total_cols)

        # Data rows
        for row_idx, cat in enumerate(categories):
            row = row_idx + 2
            ws.cell(row=row, column=1, value=cat)
            _style_data_cell(ws, row, 1)

            values = category_monthly[cat]
            for col_idx, val in enumerate(values):
                _style_data_cell(ws, row, col_idx + 2, is_currency=True).value = val

            forecasts = _linear_forecast(values, self.forecast_months)
            for fi, fval in enumerate(forecasts):
                cell = _style_data_cell(
                    ws, row, len(values) + fi + 2, is_currency=True
                )
                cell.value = max(0, fval)
                cell.fill = FORECAST_FILL

        # Auto-width
        ws.column_dimensions["A"].width = 22
        for col in range(2, total_cols + 1):
            ws.column_dimensions[get_column_letter(col)].width = 14

    def _build_income_sheet(self):
        """Monthly income by category + forecast."""
        ws = self.wb.create_sheet("Income by Category")
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
            ws.cell(row=1, column=1, value="No income data found.")
            return

        categories = sorted(category_monthly.keys())

        ws.cell(row=1, column=1, value="Category")
        for i, label in enumerate(month_labels):
            ws.cell(row=1, column=i + 2, value=label)
        for i in range(self.forecast_months):
            if monthly_ranges:
                last_start = monthly_ranges[-1][0]
                fm = last_start.month + i + 1
                fy = last_start.year
                while fm > 12:
                    fm -= 12
                    fy += 1
                flabel = date(fy, fm, 1).strftime("%b %Y") + " (F)"
            else:
                flabel = f"F{i+1}"
            ws.cell(row=1, column=len(month_labels) + i + 2, value=flabel)

        total_cols = len(month_labels) + self.forecast_months + 1
        _style_header_row(ws, 1, total_cols)

        for row_idx, cat in enumerate(categories):
            row = row_idx + 2
            ws.cell(row=row, column=1, value=cat)
            _style_data_cell(ws, row, 1)

            values = category_monthly[cat]
            for col_idx, val in enumerate(values):
                _style_data_cell(ws, row, col_idx + 2, is_currency=True).value = val

            forecasts = _linear_forecast(values, self.forecast_months)
            for fi, fval in enumerate(forecasts):
                cell = _style_data_cell(
                    ws, row, len(values) + fi + 2, is_currency=True
                )
                cell.value = max(0, fval)
                cell.fill = FORECAST_FILL

        ws.column_dimensions["A"].width = 22
        for col in range(2, total_cols + 1):
            ws.column_dimensions[get_column_letter(col)].width = 14

    def _build_budget_sheet(self):
        """Budget overview with spent vs budgeted."""
        ws = self.wb.create_sheet("Budget Overview")
        today = date.today()
        start = date(today.year, today.month, 1)
        if today.month == 12:
            end = date(today.year + 1, 1, 1) - timedelta(days=1)
        else:
            end = date(today.year, today.month + 1, 1) - timedelta(days=1)

        ws.cell(row=1, column=1, value="Budget Name")
        ws.cell(row=1, column=2, value="Budgeted Amount")
        ws.cell(row=1, column=3, value="Spent")
        ws.cell(row=1, column=4, value="Remaining")
        ws.cell(row=1, column=5, value="% Used")
        _style_header_row(ws, 1, 5)

        try:
            budgets = self.client.get_budgets(start, end)
        except Exception:
            budgets = []

        for idx, budget in enumerate(budgets):
            row = idx + 2
            attrs = budget.get("attributes", {})
            name = attrs.get("name", "Unknown")
            auto_amount = float(attrs.get("auto_budget_amount", 0) or 0)

            spent_list = attrs.get("spent", [])
            spent = sum(
                abs(float(s.get("amount", 0))) for s in spent_list
            ) if spent_list else 0.0

            remaining = auto_amount - spent
            pct = (spent / auto_amount * 100) if auto_amount > 0 else 0.0

            ws.cell(row=row, column=1, value=name)
            _style_data_cell(ws, row, 1)
            _style_data_cell(ws, row, 2, is_currency=True).value = auto_amount
            _style_data_cell(ws, row, 3, is_currency=True).value = spent
            _style_data_cell(ws, row, 4, is_currency=True).value = remaining
            cell_pct = _style_data_cell(ws, row, 5)
            cell_pct.value = pct / 100
            cell_pct.number_format = "0.0%"

        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 18

    def _build_accounts_sheet(self):
        """Account balances overview."""
        ws = self.wb.create_sheet("Accounts")

        ws.cell(row=1, column=1, value="Account Name")
        ws.cell(row=1, column=2, value="Type")
        ws.cell(row=1, column=3, value="Currency")
        ws.cell(row=1, column=4, value="Current Balance")
        ws.cell(row=1, column=5, value="Active")
        _style_header_row(ws, 1, 5)

        try:
            accounts = self.client.get_accounts("asset")
        except Exception:
            accounts = []

        for idx, acct in enumerate(accounts):
            row = idx + 2
            attrs = acct.get("attributes", {})
            ws.cell(row=row, column=1, value=attrs.get("name", ""))
            _style_data_cell(ws, row, 1)
            _style_data_cell(ws, row, 2).value = attrs.get("account_role", attrs.get("type", ""))
            _style_data_cell(ws, row, 3).value = attrs.get("currency_code", "")
            _style_data_cell(ws, row, 4, is_currency=True).value = float(
                attrs.get("current_balance", 0) or 0
            )
            _style_data_cell(ws, row, 5).value = "Yes" if attrs.get("active") else "No"

        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 18
