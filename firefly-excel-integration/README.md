# Firefly III → Excel Forecasting Report

A Python CLI tool that connects to your [Firefly III](https://www.firefly-iii.org/) personal finance manager, pulls historical financial data, and generates an Excel (.xlsx) forecasting report.

## Features

- **Monthly Summary** — Income, expenses, and net totals per month with linear trend forecast
- **Expenses by Category** — Monthly expense breakdown per category with per-category forecasts
- **Income by Category** — Monthly income breakdown per category with per-category forecasts
- **Budget Overview** — Current month budgeted vs. actual spending with remaining and % used
- **Accounts** — Asset account balances overview
- **Charts** — Embedded Excel line chart for income vs expenses trend

## Requirements

- Python 3.10+
- A running Firefly III instance (v5.x or v6.x)
- A Firefly III [Personal Access Token](https://docs.firefly-iii.org/references/firefly-iii/api/authentication/)

## Setup

```bash
# Clone and enter the directory
cd firefly-excel-integration

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Using command-line options

```bash
python main.py \
  --url https://your-firefly-instance.com \
  --token YOUR_PERSONAL_ACCESS_TOKEN \
  --months 12 \
  --forecast 3 \
  --output my_report.xlsx
```

### Using environment variables

```bash
export FIREFLY_URL=https://your-firefly-instance.com
export FIREFLY_TOKEN=YOUR_PERSONAL_ACCESS_TOKEN

python main.py
```

### Options

| Flag         | Env Var         | Default                        | Description                            |
|-------------|-----------------|--------------------------------|----------------------------------------|
| `--url`     | `FIREFLY_URL`   | *(required)*                   | Firefly III base URL                   |
| `--token`   | `FIREFLY_TOKEN` | *(required)*                   | Personal Access Token                  |
| `--months`  | —               | `12`                           | Months of historical data to include   |
| `--forecast`| —               | `3`                            | Months to forecast ahead               |
| `--output`  | —               | `firefly_forecast_report.xlsx` | Output file path                       |

## Output

The generated Excel file contains 5 sheets:

1. **Summary** — Monthly totals with income/expense/net columns, forecast rows highlighted in yellow, and an embedded line chart
2. **Expenses by Category** — One row per expense category, one column per month, with forecast columns
3. **Income by Category** — Same layout as expenses, for income categories
4. **Budget Overview** — Budget name, budgeted amount, spent, remaining, and % used for the current month
5. **Accounts** — Account name, type, currency, current balance, and active status

Forecast rows/columns are highlighted with a yellow background and labeled with `(F)`.

## How Forecasting Works

The tool uses **linear regression** on the historical monthly data to project future values. This is a simple trend-based approach — it works well for data with a consistent upward or downward trend. For more sophisticated forecasting, the data can be exported and used with dedicated forecasting tools.

## Getting a Firefly III Personal Access Token

1. Log in to your Firefly III instance
2. Go to **Options → Profile → OAuth**
3. Under **Personal Access Tokens**, click **Create New Token**
4. Give it a name (e.g., "Excel Report") and click **Create**
5. Copy the token — you won't be able to see it again

## License

MIT
