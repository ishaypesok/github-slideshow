#!/usr/bin/env python3
"""CLI tool to generate forecasting reports from Firefly III data."""

import logging
import sys

import click

from firefly_client import FireflyClient

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


@click.command()
@click.option(
    "--url",
    envvar="FIREFLY_URL",
    required=True,
    help="Firefly III base URL (e.g. https://firefly.example.com). "
    "Can also be set via FIREFLY_URL env var.",
)
@click.option(
    "--token",
    envvar="FIREFLY_TOKEN",
    required=True,
    help="Firefly III Personal Access Token. "
    "Can also be set via FIREFLY_TOKEN env var.",
)
@click.option(
    "--months",
    default=12,
    show_default=True,
    help="Number of past months to include in the report.",
)
@click.option(
    "--forecast",
    default=3,
    show_default=True,
    help="Number of months to forecast.",
)
@click.option(
    "--format",
    "output_format",
    type=click.Choice(["excel", "sheets"], case_sensitive=False),
    default="sheets",
    show_default=True,
    help="Output format: 'excel' for .xlsx file, 'sheets' for Google Sheets.",
)
@click.option(
    "--output",
    default="firefly_forecast_report.xlsx",
    show_default=True,
    help="Output Excel file path (only used with --format=excel).",
)
@click.option(
    "--google-credentials",
    envvar="GOOGLE_CREDENTIALS_FILE",
    default=None,
    help="Path to Google service account JSON key file. "
    "Can also be set via GOOGLE_CREDENTIALS_FILE env var. "
    "Required when --format=sheets.",
)
@click.option(
    "--sheet-name",
    default="Firefly III Forecast Report",
    show_default=True,
    help="Name for the created Google Sheet.",
)
@click.option(
    "--share-with",
    default=None,
    help="Email address to share the Google Sheet with (grants editor access).",
)
@click.option(
    "--sheet-url",
    default=None,
    help="URL of an existing Google Sheet to write to (overwrites content). "
    "If not provided, a new sheet is created.",
)
@click.option(
    "--accounts",
    default=None,
    help="Comma-separated list of account names or IDs to filter expenses. "
    "Only expenses from these accounts will be included. "
    "Matches partial names (e.g. '7802,298' matches accounts containing those strings).",
)
@click.option(
    "--recurring-only",
    is_flag=True,
    default=False,
    help="Only include recurring/bill-linked expenses in the report.",
)
@click.option(
    "--list-accounts",
    is_flag=True,
    default=False,
    help="List all accounts from Firefly III and exit (helps you find account names/IDs).",
)
def main(
    url: str,
    token: str,
    months: int,
    forecast: int,
    output_format: str,
    output: str,
    google_credentials: str | None,
    sheet_name: str,
    share_with: str | None,
    sheet_url: str | None,
    accounts: str | None,
    recurring_only: bool,
    list_accounts: bool,
):
    """Generate a forecasting report from Firefly III data.

    Connects to your Firefly III instance, pulls transaction history,
    and generates a report with:

    \b
    - Summary: monthly income, expenses, net, with trend forecast
    - Expenses by Category: monthly breakdown + forecast per category
    - Income by Category: monthly breakdown + forecast per category
    - Budget Overview: current month budget vs actual spending
    - Accounts: asset account balances

    \b
    Supports two output formats:
    - Google Sheets (default): creates/updates a Google Sheet
    - Excel: generates a local .xlsx file

    \b
    Filtering options:
    - --accounts: only include expenses from specific accounts
    - --recurring-only: only include recurring/bill-linked expenses
    - --list-accounts: show all your accounts and exit
    """
    logger.info("Connecting to Firefly III at %s", url)

    client = FireflyClient(url, token)

    # Verify connection
    try:
        about = client.get_about()
        version = about.get("data", {}).get("version", "unknown")
        logger.info("Connected to Firefly III v%s", version)
    except Exception as e:
        logger.error("Failed to connect to Firefly III: %s", e)
        sys.exit(1)

    # List accounts mode
    if list_accounts:
        click.echo("\nYour Firefly III accounts:\n")
        click.echo(f"{'ID':<6} {'Type':<15} {'Name':<40} {'Currency':<6} {'Balance'}")
        click.echo("-" * 85)
        for acct_type in ["asset", "expense", "revenue", "liability"]:
            try:
                accts = client.get_accounts(acct_type)
            except Exception:
                continue
            for acct in accts:
                attrs = acct.get("attributes", {})
                aid = acct.get("id", "?")
                name = attrs.get("name", "")
                curr = attrs.get("currency_code", "")
                bal = attrs.get("current_balance", "0")
                atype = attrs.get("type", acct_type)
                click.echo(f"{aid:<6} {atype:<15} {name:<40} {curr:<6} {bal}")
        return

    # Parse account filters
    account_filters = None
    if accounts:
        account_filters = [a.strip() for a in accounts.split(",") if a.strip()]

    # Build filter config
    filter_config = {
        "account_filters": account_filters,
        "recurring_only": recurring_only,
    }

    if output_format == "sheets":
        if not google_credentials:
            click.echo(
                "Error: --google-credentials or GOOGLE_CREDENTIALS_FILE is required "
                "when using --format=sheets.",
                err=True,
            )
            sys.exit(1)

        from sheets_writer import SheetsReportWriter

        writer = SheetsReportWriter(
            client, google_credentials, months_back=months, forecast_months=forecast,
            filter_config=filter_config,
        )

        if sheet_url:
            result_url = writer.write_to_existing(sheet_url)
            click.echo(f"\nReport written to existing sheet: {result_url}")
        else:
            result_url = writer.generate(sheet_name, share_with=share_with)
            click.echo(f"\nGoogle Sheet created: {result_url}")

        click.echo(
            f"  - {months} months of historical data\n"
            f"  - {forecast} months of forecasted data\n"
            f"  - Sheets: Summary, Expenses by Category, Income by Category, "
            f"Budget Overview, Accounts"
        )
    else:
        from report_generator import ReportGenerator

        generator = ReportGenerator(
            client, months_back=months, forecast_months=forecast,
            filter_config=filter_config,
        )
        generator.generate(output)

        click.echo(f"\nExcel report generated: {output}")
        click.echo(
            f"  - {months} months of historical data\n"
            f"  - {forecast} months of forecasted data\n"
            f"  - Sheets: Summary, Expenses by Category, Income by Category, "
            f"Budget Overview, Accounts"
        )


if __name__ == "__main__":
    main()
