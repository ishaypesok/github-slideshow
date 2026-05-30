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
            client, google_credentials, months_back=months, forecast_months=forecast
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

        generator = ReportGenerator(client, months_back=months, forecast_months=forecast)
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
