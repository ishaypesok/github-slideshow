# Firefly III → Forecasting Report (Google Sheets / Excel)

A Python CLI tool that connects to your [Firefly III](https://www.firefly-iii.org/) personal finance manager, pulls historical financial data, and generates a forecasting report in **Google Sheets** (default) or **Excel (.xlsx)**.

## Features

- **Monthly Summary** — Income, expenses, and net totals per month with linear trend forecast
- **Expenses by Category** — Monthly expense breakdown per category with per-category forecasts
- **Income by Category** — Monthly income breakdown per category with per-category forecasts
- **Budget Overview** — Current month budgeted vs. actual spending with remaining and % used
- **Accounts** — Asset account balances overview
- **Charts** — Embedded line chart for income vs expenses trend (Excel output)
- **Two output formats** — Google Sheets (default) or Excel (.xlsx)

## Requirements

- Python 3.10+
- A running Firefly III instance (v5.x or v6.x)
- A Firefly III [Personal Access Token](https://docs.firefly-iii.org/references/firefly-iii/api/authentication/)
- For Google Sheets output: a Google Cloud service account JSON key file

## Setup

```bash
cd firefly-excel-integration
pip install -r requirements.txt
```

## Usage

### Google Sheets (default)

```bash
python main.py \
  --url https://your-firefly-instance.com \
  --token YOUR_FIREFLY_TOKEN \
  --google-credentials /path/to/service-account.json \
  --share-with your-email@gmail.com
```

This creates a new Google Sheet named "Firefly III Forecast Report" and shares it with your email.

#### Write to an existing Google Sheet

```bash
python main.py \
  --url https://your-firefly-instance.com \
  --token YOUR_FIREFLY_TOKEN \
  --google-credentials /path/to/service-account.json \
  --sheet-url "https://docs.google.com/spreadsheets/d/SHEET_ID/edit"
```

### Excel output

```bash
python main.py \
  --url https://your-firefly-instance.com \
  --token YOUR_FIREFLY_TOKEN \
  --format excel \
  --output my_report.xlsx
```

### Using environment variables

```bash
export FIREFLY_URL=https://your-firefly-instance.com
export FIREFLY_TOKEN=YOUR_PERSONAL_ACCESS_TOKEN
export GOOGLE_CREDENTIALS_FILE=/path/to/service-account.json

python main.py --share-with your-email@gmail.com
```

### All Options

| Flag                    | Env Var                    | Default                        | Description                                       |
|------------------------|----------------------------|--------------------------------|---------------------------------------------------|
| `--url`                | `FIREFLY_URL`              | *(required)*                   | Firefly III base URL                              |
| `--token`              | `FIREFLY_TOKEN`            | *(required)*                   | Personal Access Token                             |
| `--format`             | —                          | `sheets`                       | Output format: `sheets` or `excel`                |
| `--months`             | —                          | `12`                           | Months of historical data to include              |
| `--forecast`           | —                          | `3`                            | Months to forecast ahead                          |
| `--google-credentials` | `GOOGLE_CREDENTIALS_FILE`  | —                              | Path to Google service account JSON key           |
| `--sheet-name`         | —                          | `Firefly III Forecast Report`  | Name for new Google Sheet                         |
| `--share-with`         | —                          | —                              | Email to share the Google Sheet with              |
| `--sheet-url`          | —                          | —                              | URL of existing Google Sheet to overwrite          |
| `--output`             | —                          | `firefly_forecast_report.xlsx` | Excel output file path (only with `--format=excel`)|

## Output Sheets

Both formats generate 5 sheets:

1. **Summary** — Monthly totals with income/expense/net columns, forecast rows labeled `(F)`
2. **Expenses by Category** — One row per expense category, one column per month, with forecast columns
3. **Income by Category** — Same layout as expenses, for income categories
4. **Budget Overview** — Budget name, budgeted amount, spent, remaining, and % used for the current month
5. **Accounts** — Account name, type, currency, current balance, and active status

In Excel output, forecast rows are highlighted yellow. In Google Sheets, headers are styled with blue backgrounds.

## How Forecasting Works

The tool uses **linear regression** on the historical monthly data to project future values. This is a simple trend-based approach — it works well for data with a consistent upward or downward trend.

## Setting Up Google Sheets Access

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project (or use an existing one)
3. Enable the **Google Sheets API** and **Google Drive API**
4. Go to **IAM & Admin → Service Accounts** → **Create Service Account**
5. Download the JSON key file
6. Use the key file path as `--google-credentials`
7. When using `--share-with`, the sheet is shared with the specified email as an editor

## Getting a Firefly III Personal Access Token

1. Log in to your Firefly III instance
2. Go to **Options → Profile → OAuth**
3. Under **Personal Access Tokens**, click **Create New Token**
4. Give it a name (e.g., "Sheets Report") and click **Create**
5. Copy the token — you won't be able to see it again

## License

MIT
