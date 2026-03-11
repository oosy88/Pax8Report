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

---

## Script 2: License Optimization Analyzer

Analyzes the report generated above and recommends how many licenses to convert from monthly to annual commitment to reduce costs.

### Pricing

The optimizer automatically connects to the PAX8 API to fetch live pricing for every product found in the report. It compares Monthly vs Annual commitment rates using your partner buy rates, so dollar calculations always reflect current pricing.

This uses the same `.env` credentials as the report generator — no additional setup needed. If credentials are missing, the analyzer still runs but skips dollar calculations.

### Run the Analyzer

```bash
# Pass the report file explicitly
python3 license_optimizer.py pax8_microsoft_license_report_2026-03-11.xlsx

# Or let it auto-detect the most recent report
python3 license_optimizer.py
```

### Output

Generates `license_optimization_YYYY-MM-DD.xlsx` with four tabs:

- **Recommendations** — Per client/product annual vs monthly split at three risk tiers (conservative, moderate, aggressive) with savings and risk exposure.
- **Client Trends** — Month-by-month license count grid for visual trend inspection.
- **Savings Summary** — Per-client and portfolio-wide savings potential at each tier.
- **Unpriced Products** — Products where PAX8 didn't return both monthly and annual pricing (e.g., products only available on one billing term).

### How Tiers Work

| Tier | Logic | Risk Level |
|---|---|---|
| Conservative | 12-month min minus 10% buffer | Lowest |
| Moderate | 12-month min (stable/growing trends only) | Medium |
| Aggressive | 6-month min (growing + low volatility + no recent decreases) | Higher |

---

## Troubleshooting

- **Authentication fails** — Double-check your `PAX8_CLIENT_ID` and `PAX8_CLIENT_SECRET` in `.env`. Ensure your account has Partner Admin access.
- **No companies found** — Your API credentials may not have the correct permissions.
- **Rate limiting** — The script handles rate limits automatically with exponential backoff and retries.
- **Partial failures** — If individual client or subscription calls fail, the script logs the error and continues. A summary of all errors is printed at the end.
