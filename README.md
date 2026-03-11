# PAX8 Microsoft License Report Generator

Generates a comprehensive Excel report of all Microsoft subscriptions across all PAX8 clients, including full subscription history.

## Output

The script produces an Excel file (`pax8_microsoft_license_report_YYYY-MM-DD.xlsx`) with two tabs:

- **Summary** — One row per client per Microsoft subscription with product details, status, quantity, pricing, and billing terms.
- **Subscription History** — One row per historical change event showing quantity changes over time.

## Prerequisites

- **Python 3.9+** — Check with `python3 --version`. If not installed, install via Homebrew:
  ```bash
  brew install python
  ```
- **PAX8 API credentials** — Created at [Settings > Integrations](https://app.pax8.com/integrations/credentials) in the PAX8 app. Requires Partner Admin or Primary Partner Admin role.

## Setup

1. **Clone or navigate to the project directory:**
   ```bash
   cd /path/to/Pax8Report
   ```

2. **Create and activate a virtual environment:**
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure your credentials:**
   ```bash
   cp .env.example .env
   ```
   Open `.env` in your editor and replace the placeholder values with your real PAX8 API Client ID and Client Secret.

5. **Run the script:**
   ```bash
   python3 pax8_report.py
   ```

6. **Find the output:**
   The Excel report is saved in the same directory as the script. The full path is printed to the terminal when the script completes.

## What It Does

1. Authenticates with the PAX8 API using OAuth2 client credentials.
2. Fetches all companies (clients) from your PAX8 account.
3. For each company, retrieves Active, Cancelled, and PendingCancel subscriptions.
4. Filters to Microsoft-only subscriptions by checking the product vendor name.
5. Fetches the full change history for each Microsoft subscription.
6. Exports everything to a formatted Excel workbook with filters, sorted alphabetically by company name.

## Troubleshooting

- **Authentication fails** — Double-check your `PAX8_CLIENT_ID` and `PAX8_CLIENT_SECRET` in `.env`. Ensure your account has Partner Admin access.
- **No companies found** — Your API credentials may not have the correct permissions.
- **Rate limiting** — The script handles rate limits automatically with exponential backoff and retries.
- **Partial failures** — If individual client or subscription calls fail, the script logs the error and continues. A summary of all errors is printed at the end.
