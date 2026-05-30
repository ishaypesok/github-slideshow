#!/usr/bin/env python3
"""CLI tool to generate Excel forecasting reports from Firefly III data."""

import logging
import os
import sys

import click

from firefly_client import FireflyClient
from report_generator import ReportGenerator

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
    "--output",
    default="firefly_forecast_report.xlsx",
    show_default=True,
    help="Output Excel file path.",
)
def main(url: str, token: str, months: int, forecast: int, output: str):
    """Generate an Excel forecasting report from Firefly III data.

    Connects to your Firefly III instance, pulls transaction history,
    and generates an Excel workbook with:

    \b
    - Summary: monthly income, expenses, net, with trend forecast
    - Expenses by Category: monthly breakdown + forecast per category
    - Income by Category: monthly breakdown + forecast per category
    - Budget Overview: current month budget vs actual spending
    - Accounts: asset account balances
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

    generator = ReportGenerator(client, months_back=months, forecast_months=forecast)
    generator.generate(output)

    click.echo(f"\nReport generated: {output}")
    click.echo(
        f"  - {months} months of historical data\n"
        f"  - {forecast} months of forecasted data\n"
        f"  - Sheets: Summary, Expenses by Category, Income by Category, "
        f"Budget Overview, Accounts"
    )


if __name__ == "__main__":
    main()
