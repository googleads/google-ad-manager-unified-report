# Google Ad Manager Unified Report

This Google Sheets and Apps Script tool combines your Ad Exchange revenue and On-Platform Multicall (OPMC) revenue into one report. Run the setup function once in a blank Google Sheet to create all tabs, formulas, filters, buttons, and charts.

---

## 1. Workbook Tabs

* **Overview**: Run reports, filter by ad unit or country, view revenue totals, and check daily revenue, daily eCPM, incremental uplift, and country charts.
* **Ad Unit Breakdown**: See combined revenue for each Parent Ad Unit (A) and its OPMC Child Ad Units (A1 to A4).
* **Instructions**: View the setup steps, button guide, and key revenue definitions inside your sheet.

---

## 2. One-Time Setup

> [!IMPORTANT]
> Copy **both** [`appsscript.json`](appsscript.json) and [`Code.js`](Code.js) into your Apps Script project. Without `appsscript.json`, Apps Script cannot get Google Ad Manager API access.

### Step A: Set Up Google Cloud

1. Open your Google Cloud project and enable the [Google Ad Manager API](https://developers.google.com/ad-manager/api/beta) (`admanager.googleapis.com`).
2. Open `[IAM & Admin] > [IAM]` and make sure your Google account has the **Service Usage Consumer** role (`roles/serviceusage.serviceUsageConsumer`).
3. Open `[Dashboard]` and copy your **GCP Project ID** (or Project Number) from the **Project info** card.

### Step B: Set Up Your Google Sheet

1. In a new browser tab, open [sheets.new](https://sheets.new) and click `[Extensions] > [Apps Script]`.
2. Click `[Project Settings]` (gear icon) and check `[Show "appsscript.json" manifest file in editor]`.
3. Click `[Editor]` (`< >`), open `appsscript.json`, and replace its contents with the JSON below:
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
6. Return to your Google Sheet tab and refresh the page once (`Cmd+R` or `F5`) to activate the buttons.

---

## 3. Run a Report

1. Open the **Overview** tab.
2. Type your Google Ad Manager **Network Code** in cell `C7` and press **Enter**.
3. Type your **GCP Project ID** in cell `I7` and press **Enter**.
4. Double-click cell `F7` to pick the **Start Date** and cell `F8` to pick the **End Date** (`YYYY-MM-DD`).
5. Click `[Generate Report]` (`B10:C10`) and wait for the **Report complete** message.

---

## 4. Controls and Key Terms

| Control | Cell | Action |
| :--- | :--- | :--- |
| `[Generate Report]` | `B10:C10` | Fetch fresh data from Google Ad Manager and update all charts and tables. |
| `[Reset Filters]` | `E10` | Reset both the Ad Unit filter and the Country filter to `ALL`. |
| `[Clear Data]` | `F10` | Clear all report rows from the workbook while keeping your settings and formulas. |
| `[Download CSV]` | `H10:I10` | Download the combined report as a CSV file. |
| **Ad Unit Filter** | `C12` | Show revenue and impressions for one ad unit from the dropdown list. |
| **Country Filter** | `F12` | Show revenue and impressions for one country from the dropdown list. |

### Key Revenue Terms

* **Ad Exchange Revenue**: Revenue from standard Google Ad Exchange auctions on your ad units.
* **OPMC Revenue**: Incremental revenue from On-Platform Multicall (OPMC) mediation ad units.
* **Total Revenue**: The sum of Ad Exchange Revenue and OPMC Revenue.
* **Daily eCPM**: Revenue per 1,000 impressions (`(Revenue / Impressions) * 1000`) for Total, Ad Exchange, and OPMC traffic.
* **Parent Ad Unit Rollup**: Groups each Parent Ad Unit (A) with its OPMC Child Ad Units (A1 to A4) on **Ad Unit Breakdown**. Change any parent name in Table 2 (Column D) to update Table 1 automatically.

---

## Terms and Security

Use of the Google Ad Manager API is subject to the [Google APIs Terms of Service](https://developers.google.com/terms) and the [Google Ad Manager API Terms and Conditions](https://developers.google.com/ad-manager/api/terms).

This project is intended for demonstration and reporting utility purposes.

Eligibility for the [Google Open Source Software Vulnerability Rewards Program](https://bughunters.google.com/open-source-security) is determined by the [Google Open Source Software Vulnerability Reward Program Rules](https://bughunters.google.com/about/rules/open-source/google-open-source-software-vulnerability-reward-program-rules).

## License

Copyright 2026 Google LLC. Licensed under the [Apache License, Version 2.0](LICENSE).
