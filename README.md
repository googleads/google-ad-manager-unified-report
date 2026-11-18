# Google Ad Manager Unified Report

This Google Sheets and Apps Script tool combines Primary Programmatic revenue (Ad Exchange, Open Bidding, and SDK Bidding) with On-Platform Multicall (OPMC) revenue in one report. Run the setup function once in a blank Google Sheet to create all tabs, formulas, filters, buttons, and charts.

---

## 1. Workbook Tabs

* **Overview**: Run reports, filter by ad unit, country, or demand channel, review totals, and inspect five interactive charts.
* **Ad Unit Breakdown**: View combined revenue for each Parent Ad Unit (`A`) and its OPMC Child Ad Units (`A1` to `An`).
* **Instructions**: Read setup steps, button usage, filter rules, and key metric definitions inside the workbook.

---

## 2. One-Time Setup

> [!IMPORTANT]
> Copy **both** [`appsscript.json`](appsscript.json) and [`Code.js`](Code.js) into your Apps Script project. Without `appsscript.json`, Apps Script cannot access the Google Ad Manager API.

### Step A: Set Up Google Cloud

1. Open your Google Cloud project and enable the [Google Ad Manager API](https://developers.google.com/ad-manager/api/beta) (`admanager.googleapis.com`).
2. Open `[IAM & Admin] > [IAM]` and confirm your Google account holds the **Service Usage Consumer** role (`roles/serviceusage.serviceUsageConsumer`).
3. Open `[Dashboard]` and copy your **GCP Project ID** (or Project Number) from the **Project info** card.

### Step B: Set Up Your Google Sheet

1. Open a blank sheet at [sheets.new](https://sheets.new) and click `[Extensions] > [Apps Script]`.
2. Open `[Project Settings]` (gear icon) and check `[Show "appsscript.json" manifest file in editor]`.
3. Click `[Editor]` (`< >`), open `appsscript.json`, and paste the manifest below:
   ```json
   {
     "timeZone": "America/Los_Angeles",
     "dependencies": {},
     "exceptionLogging": "STACKDRIVER",
     "runtimeVersion": "V8",
     "oauthScopes": [
       "https://www.googleapis.com/auth/admanager",
       "https://www.googleapis.com/auth/spreadsheets",
       "https://www.googleapis.com/auth/script.external_request",
       "https://www.googleapis.com/auth/script.container.ui"
     ]
   }
   ```
4. Open `Code.gs`, replace its contents with [`Code.js`](Code.js) from this repository, and click `[Save project]`.
5. Select `authorizeAndVerifyGamScope` in the toolbar, click `[Run]`, and approve the permission prompt.
6. Return to your Google Sheet tab and refresh the page once (`Cmd+R` or `F5`) to activate the row 10 buttons.

---

## 3. Run a Report

1. Open the **Overview** tab.
2. Enter your Google Ad Manager **Network Code** in cell `C7` (and an optional query-time substring filter in cell `C8`) and press **Enter**.
3. Enter your **GCP Project ID** in cell `I7` and press **Enter**.
4. Double-click cell `F7` for the **Start Date** and cell `F8` for the **End Date** (`YYYY-MM-DD`). Date ranges over 62 days automatically split into chunks and bisect on timeouts.
5. Click `[Generate Report]` (`B10:C10`) and wait for the **Report complete** message.

---

## 4. Controls, Charts, and CSV Export

### Controls and Live Filters

| Control | Cell | Action |
| :--- | :--- | :--- |
| **Ad Unit Filter (Optional)** | `C8` | Optional query-time substring filter (`Ad unit all levels contains`) applied during report fetch so parent tags keep nested OPMC child revenue. |
| `[Generate Report]` | `B10:C10` | Fetch settled programmatic data from Google Ad Manager and update all charts and tables. |
| `[Reset Filters]` | `E10` | Reset the Ad Unit (`C12`), Country (`F12`), and Demand Channel (`I12`) filters to `ALL`. |
| `[Clear Data]` | `F10` | Clear all report rows from the workbook while preserving your settings and formulas. |
| `[Download CSV]` | `H10:I10` | Download a four-section BI-ready CSV with `YYYY-MM-DD` dates and plain numeric float columns. |
| **Ad Unit Filter** | `C12` | Filter **Overview** by one ad unit; selecting a Parent Ad Unit includes all its OPMC child units. |
| **Country Filter** | `F12` | Filter **Overview** by one country. |
| **Demand Channel** | `I12` | Filter **Overview** and **Ad Unit Breakdown** live by `ALL`, `Ad Exchange`, `Open Bidding`, or `SDK Bidding`. |

### Overview Charts (`v1.3`)

1. **Revenue Trends**: Daily `Total Unified Revenue`, `Primary Programmatic Revenue`, and `OPMC Revenue`.
2. **OPMC Revenue by Demand Channel**: Daily OPMC revenue split by `Ad Exchange`, `Open Bidding`, and `SDK Bidding`.
3. **eCPM Trends**: Daily eCPM for `Total`, `Primary Programmatic`, and `OPMC` traffic.
4. **Stacked Incremental Revenue**: Daily Primary baseline stacked with incremental OPMC revenue.
5. **Country Revenue Breakdown (Top 10 Countries)**: Top 10 countries ranked by total revenue.

### BI-Ready CSV Format

`[Download CSV]` exports four sections: `Parent Ad Unit Revenue Summary`, `Daily Incremental Revenue & eCPM`, `Daily Revenue Trends`, and `Country Revenue Breakdown`. Percentage share and uplift columns export as plain decimals without `%` or `+` characters alongside ISO `YYYY-MM-DD` dates so BI tools and data lakes detect numeric `FLOAT` and `DATE` types natively.

### Key Revenue Terms

* **Primary Programmatic Revenue**: Google-settled revenue from Ad Exchange, Open Bidding, and SDK Bidding (excludes Mediation and Ad Server to prevent double-counting).
* **OPMC Revenue**: Incremental Google-settled revenue earned on On-Platform Multicall (OPMC) child ad units.
* **Total Unified Revenue**: The sum of Primary Programmatic Revenue and OPMC Revenue (zero double-counting).
* **Parent Ad Unit Rollup**: Groups each Parent Ad Unit (`A`) with its OPMC Child Ad Units (`A1` to `An`) on **Ad Unit Breakdown**. Edit Column `D` in Table 2 to reassign any child unit.

---

## Terms and Security

Use of the Google Ad Manager API is subject to the [Google APIs Terms of Service](https://developers.google.com/terms) and the [Google Ad Manager API Terms and Conditions](https://developers.google.com/ad-manager/api/terms).

This project is intended for demonstration and reporting utility purposes.

Eligibility for the [Google Open Source Software Vulnerability Rewards Program](https://bughunters.google.com/open-source-security) is determined by the [Google Open Source Software Vulnerability Reward Program Rules](https://bughunters.google.com/about/rules/open-source/google-open-source-software-vulnerability-reward-program-rules).

## License

Copyright 2026 Google LLC. Licensed under the [Apache License, Version 2.0](LICENSE).
