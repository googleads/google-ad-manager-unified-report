# Google Ad Manager Unified Report

This Google Apps Script and Google Sheets solution combines your Ad Exchange revenue and your On-Platform Multicall (OPMC) revenue into one report. Use this workbook to view daily trends, compare country performance, and see total revenue for each ad unit.

## Official Google Sheet Template

* **Template Link**: [[MAKE A COPY] Google Ad Manager Unified Report](https://docs.google.com/spreadsheets/d/1N0-iTPXxIw4mv-SkZQc44j55_m9w_OLqE1GMtD2Q19I/copy)
* Make a copy of the template in your own Google Drive, or attach `Code.js` and `appsscript.json` to an existing Google Sheet.

---

## 1. Workbook Tabs

* **Instructions**: Shows the one-time setup guide, report steps, button descriptions, and key revenue terms.
* **Overview**: Runs reports, filters data by ad unit or country, shows total revenue cards, and displays daily, incremental, and country charts.
* **Ad Unit Breakdown**: Shows revenue for each Parent Ad Unit (A) combined with its OPMC Child Ad Units (A1, A2, A3, A4).

---

## 2. One-Time Setup: Google Cloud API and Apps Script

Complete these six steps one time before you run your first report.

1. **Enable the Google Ad Manager API (`Google Cloud Console -> APIs & Services`)**:
   Open your Google Cloud project and enable the [Google Ad Manager API](https://developers.google.com/ad-manager/api/beta) (`admanager.googleapis.com`).
2. **Verify IAM Permission (`Google Cloud Console -> IAM & Admin -> IAM`)**:
   Ensure your Google account has the **Service Usage Consumer** role (`roles/serviceusage.serviceUsageConsumer`) on that GCP project.
3. **Copy your GCP Project ID (`Google Cloud Console -> Dashboard`)**:
   Copy your GCP Project ID (or Project Number) from the **Project info** card on the Cloud Console Dashboard.
4. **Enter your GCP Project ID (`Google Sheet -> Overview Tab -> Cell I8`)**:
   Open the **Overview** tab, enter your GCP Project ID in cell **I8**, and press **Enter** to save the cell.
5. **Open the script editor (`Google Sheet -> Extensions -> Apps Script`)**:
   In the top menu of the sheet, click **Extensions** and then click **Apps Script** to open the script editor.
6. **Authorize the script (`Apps Script -> Editor < > -> Toolbar`)**:
   Select `authorizeAndVerifyGamScope` in the top toolbar, click **Run**, and approve the permission prompt.

---

## 3. How to Generate a Report (Step-by-Step)

1. **Open the Overview tab**:
   Click the **Overview** tab at the bottom of the sheet.
2. **Enter Network Code and Project ID**:
   Type your Google Ad Manager Network Code in cell **C7** and your GCP Project ID in cell **I8**, then press **Enter**.
3. **Select your date range**:
   Double-click cell **F7** to select the Start Date and double-click cell **F8** to select the End Date (`YYYY-MM-DD`).
4. **Run the report**:
   Click the blue **Generate Report** button in row 10. Wait for the **Report complete** message to appear.

---

## 4. Overview Tab Parameter Reference (`B7:I8`)

| Parameter Name | Label Cell | Input Cell | Description |
| :--- | :--- | :--- | :--- |
| **Network Code** | `B7` | `C7` | Your Google Ad Manager Network Code. |
| **Start Date (YYYY-MM-DD)** | `E7` | `F7` | First date of the report. Double-click the cell to open the calendar picker. |
| **End Date (YYYY-MM-DD)** | `E8` | `F8` | Last date of the report. Double-click the cell to open the calendar picker. |
| **GCP Project ID (Quota)** | `H8` | `I8` | Your Google Cloud Project ID or Project Number where the Ad Manager API is enabled. |

---

## 5. How to Use the Buttons and Filters in the Overview Tab

| Button or Filter | Location | How to Use It |
| :--- | :--- | :--- |
| **Generate Report** | `B10:C10` | Fetch new data from Google Ad Manager and update all charts and tables. |
| **Reset Filters** | `E10` | Reset both the Ad Unit filter and the Country filter back to `ALL`. |
| **Clear Data** | `F10` | Wipe all fetched report data from all sheets while preserving your settings and formulas. |
| **Download CSV** | `H10:I10` | Download the full report as a CSV file. |
| **Filter by Ad Unit** | `C12` | Select an ad unit from the dropdown list in cell `C12` to show revenue for that ad unit only. |
| **Filter by Country** | `F12` | Select a country from the dropdown list in cell `F12` to show revenue for that country only. |

---

## 6. Key Revenue Terms

* **Ad Exchange Revenue**: Revenue from standard Google Ad Exchange auctions on your ad units.
* **OPMC Revenue**: Additional revenue from On-Platform Multicall (OPMC) mediation ad units.
* **Total Revenue**: The sum of Ad Exchange Revenue plus OPMC Revenue.
* **Parent Ad Unit Rollup**: On the **Ad Unit Breakdown** tab, Table 1 groups each Parent Ad Unit (A) with its OPMC Child Ad Units (A1 to A4). If you change a parent name in Table 2 (Column D), Table 1 updates automatically.

---

## Terms and Security

Use of the Google Ad Manager API is subject to the [Google APIs Terms of Service](https://developers.google.com/terms) and the [Google Ad Manager API Terms and Conditions](https://developers.google.com/ad-manager/api/terms).

This project is intended for demonstration and reporting utility purposes.

Eligibility for the [Google Open Source Software Vulnerability Rewards Program](https://bughunters.google.com/open-source-security) is determined by the [Google Open Source Software Vulnerability Reward Program Rules](https://bughunters.google.com/about/rules/open-source/google-open-source-software-vulnerability-reward-program-rules).

## License

Copyright 2026 Google LLC. Licensed under the [Apache License, Version 2.0](LICENSE).
