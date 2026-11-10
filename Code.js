/**
 * Copyright 2026 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Google Ad Manager Unified Report
 *
 * This script gets Ad Exchange and On-Platform Multicall (OPMC) report data
 * from the Google Ad Manager API and writes the combined results to Google Sheets.
 */

// Base URLs and default project settings for the Google Ad Manager API.
const PUBLIC_GAM_BASE_URL = 'https://admanager.googleapis.com/v1';
const DEFAULT_QUOTA_PROJECT = '';
const FALLBACK_QUOTA_PROJECTS = [];

// Currency symbols mapped by three-letter ISO currency code.
const CURRENCY_SYMBOLS = {
  USD: '$',
  EUR: '€',
  GBP: '£',
  JPY: '¥',
  CAD: 'CA$',
  AUD: 'A$',
  CHF: 'CHF ',
  BRL: 'R$',
  INR: '₹',
  ILS: '₪',
  SGD: 'S$',
  HKD: 'HK$',
  NZD: 'NZ$',
  KRW: '₩',
  PLN: 'zł ',
  SEK: 'kr ',
  NOK: 'kr ',
  DKK: 'kr ',
  TRY: '₺',
  MXN: 'Mex$',
  ZAR: 'R ',
  TWD: 'NT$',
  THB: '฿',
  PHP: '₱',
  IDR: 'Rp ',
  MYR: 'RM ',
  VND: '₫',
  AED: 'AED ',
  SAR: 'SAR ',
  CZK: 'Kč ',
  HUF: 'Ft ',
  CLP: 'CLP$',
  COP: 'COL$',
  PEN: 'S/.',
  ARS: 'ARS$',
  UAH: '₴'
};

// Prefix and 1x1 transparent PNG data for the action buttons on the Overview tab.
const BUTTON_MARKER_PREFIX = 'gam-unified-btn:';
const TRANSPARENT_1X1_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVQI12NgAAIAAAUAAeImBZsAAAAASUVORK5CYII=';

// Common ad format and placement words to ignore when matching parent and child ad units.
const FORMAT_AND_NOISE_WORDS = {
  opmc: true,
  multicall: true,
  waterfall: true,
  native: true,
  banner: true,
  interstitial: true,
  rewarded: true,
  video: true,
  appopen: true,
  placement: true,
  production: true,
  prod: true,
  direct: true,
  staging: true,
  screen: true,
  atf: true,
  btf: true,
  multisize: true,
  img: true,
  no: true,
  timer: true
};

/**
 * Builds the workbook tabs (Overview, Instructions, Ad Unit Breakdown, and hidden data tabs),
 * requests OAuth authorization for the Google Ad Manager API, checks the granted scopes,
 * and adds the action buttons to the Overview tab.
 */
function authorizeAndVerifyGamScope() {
  try {
    if (typeof ScriptApp.requireAllScopes === 'function') {
      ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
    }
    if (typeof ScriptApp.requireScopes === 'function') {
      ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
        'https://www.googleapis.com/auth/admanager'
      ]);
    }
  } catch (e) {}

  try {
    setupWorkbook_();
  } catch (setupErr) {
    Logger.log('Workbook setup note: ' + setupErr);
  }

  try {
    installDashboardButtons_();
  } catch (e) {}

  Logger.log('✅ Workbook tabs, formulas, buttons, and charts verified.');

  const tokInfo = acquireValidGamToken_('');
  Logger.log('✅ SUCCESS! Google Ad Manager OAuth token acquired via ' + tokInfo.source);
  Logger.log('Granted scopes: ' + tokInfo.scopes);
  Logger.log('👉 Next step: Switch to your Google Sheet tab and refresh the page once (Cmd+R / F5) to activate the row 10 buttons.');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const overview = ss && ss.getSheetByName('Overview');
    if (overview) {
      ss.setActiveSheet(overview);
      ss.toast(
        'Please refresh this browser tab once (Cmd+R or F5) to activate the row 10 buttons.',
        '✅ Setup Complete — Refresh Tab Once',
        -1
      );
      SpreadsheetApp.flush();
    }
  } catch (e) {}

  return tokInfo;
}

/**
 * Public helper to build all workbook tabs, formulas, dropdowns, charts, and buttons
 * in a blank Google Sheet without overwriting an already configured workbook.
 */
function setupWorkbook() {
  setupWorkbook_();
  installDashboardButtons_();
}

/**
 * Adjusts a sheet's grid dimensions to match targetRows and targetCols.
 */
function ensureSheetGridSize_(sheet, targetRows, targetCols) {
  if (!sheet) return;
  const maxCols = sheet.getMaxColumns();
  if (maxCols > targetCols) {
    sheet.deleteColumns(targetCols + 1, maxCols - targetCols);
  } else if (maxCols < targetCols) {
    sheet.insertColumnsAfter(maxCols, targetCols - maxCols);
  }
  const maxRows = sheet.getMaxRows();
  if (maxRows > targetRows) {
    sheet.deleteRows(targetRows + 1, maxRows - targetRows);
  } else if (maxRows < targetRows) {
    sheet.insertRowsAfter(maxRows, targetRows - maxRows);
  }
}

/**
 * Draws an outer card border around a range and an optional thick top accent bar.
 */
function styleCardBox_(range, borderHex, topAccentHex, borderStyle) {
  if (!range) return;
  const bStyle = borderStyle || SpreadsheetApp.BorderStyle.SOLID;
  range.setBorder(true, true, true, true, false, false, borderHex || '#dadce0', bStyle);
  if (topAccentHex) {
    const sheet = range.getSheet();
    sheet
      .getRange(range.getRow(), range.getColumn(), 1, range.getNumColumns())
      .setBorder(true, null, null, null, null, null, topAccentHex, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

/**
 * Builds the Daily eCPM Trends line chart anchored on Overview at (anchorRow, 2).
 */
function buildDailyEcpmChart_(overview, helperSheet, anchorRow, numRows) {
  const maxRow = Math.max(numRows || 400, 100);
  return overview
    .newChart()
    .asLineChart()
    .addRange(helperSheet.getRange('A1:A' + maxRow))
    .addRange(helperSheet.getRange('Q1:S' + maxRow))
    .setNumHeaders(1)
    .setPosition(anchorRow || 38, 2, 8, 4)
    .setOption('title', 'Daily eCPM Trends')
    .setOption('width', 1190)
    .setOption('height', 355)
    .setOption('legend', { position: 'top', textStyle: { fontName: 'Roboto', fontSize: 11 } })
    .setOption('curveType', 'function')
    .setOption('hAxis', { title: 'Date', textStyle: { fontName: 'Roboto', fontSize: 10 } })
    .setOption('vAxis', { title: 'eCPM', textStyle: { fontName: 'Roboto', fontSize: 10 } })
    .setOption('colors', ['#0f9d58', '#1a73e8', '#f29900'])
    .setOption('series', {
      0: { lineWidth: 3, color: '#0f9d58' },
      1: { lineWidth: 3, color: '#1a73e8' },
      2: { lineWidth: 3, lineDashStyle: [6, 3], color: '#f29900' }
    })
    .setOption('useFirstColumnAsDomain', true)
    .build();
}

/**
 * Upgrades an existing 3-chart Overview workbook to the 4-chart layout by inserting
 * the Daily eCPM Trends section at row 37 without overwriting user inputs or data.
 */
function upgradeToFourChartLayoutIfNeeded_(ss) {
  if (!ss) return;
  const overview = ss.getSheetByName('Overview');
  const helperSheet = ss.getSheetByName('Chart_Helper');
  const dailySheet = ss.getSheetByName('Daily_Data');
  if (!overview || !helperSheet || !dailySheet) return;

  const b37 = String(overview.getRange('B37').getDisplayValue() || '').trim();
  if (b37.toLowerCase().indexOf('ecpm') !== -1) return;
  if (b37.toLowerCase().indexOf('stacked') === -1) return;

  if (dailySheet.getMaxColumns() < 8) {
    dailySheet.insertColumnsAfter(dailySheet.getMaxColumns(), 8 - dailySheet.getMaxColumns());
  }
  dailySheet.getRange('F1:H1')
    .setValues([['Ad Exchange Impressions', 'On-platform multicall Impressions', 'Total Impressions']])
    .setBackground('#0f9d58')
    .setFontColor('#ffffff')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (helperSheet.getMaxColumns() < 20) {
    helperSheet.insertColumnsAfter(helperSheet.getMaxColumns(), 20 - helperSheet.getMaxColumns());
  }
  if (helperSheet.getMaxRows() < 405) {
    helperSheet.insertRowsAfter(helperSheet.getMaxRows(), 405 - helperSheet.getMaxRows());
  }
  helperSheet.getRange('N1:S1').setValues([[
    'Ad Exchange Impressions', 'On-platform multicall Impressions', 'Total Impressions',
    'Total eCPM', 'Ad Exchange eCPM', 'On-platform multicall eCPM'
  ]]);
  if (String(helperSheet.getRange('Q2').getDisplayValue() || '') === '') {
    helperSheet.getRange('N2:S2').setValues([[0, 0, 0, 0, 0, 0]]);
  }
  helperSheet.getRange('N2:P405').setNumberFormat('#,##0');
  helperSheet.getRange('Q2:S405').setNumberFormat('$#,##0.00');

  overview.insertRowsAfter(36, 19);
  overview.getRange('A37:J55')
    .setBackground('#f1f3f4')
    .setFontColor('#202124')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setVerticalAlignment('middle');
  overview.setRowHeight(37, 34);
  overview.setRowHeights(38, 17, 22);
  overview.setRowHeight(55, 18);

  overview.getRange('B37:G37').merge();
  overview.getRange('H37:I37').merge();
  overview.getRange('B37').setValue('▎ Daily eCPM Trends');
  overview.getRange('H37').setFormula(
    '="Active Filter: "&IF(C12<>"ALL", "Ad Unit = "&C12, IF(F12<>"ALL", "Country = "&F12, "All Inventory & Countries"))'
  );
  overview.getRange('B37:I54').setBackground('#ffffff');
  styleCardBox_(overview.getRange('B37:I54'), '#dadce0');
  overview.getRange('B37:G37').setFontColor('#1a73e8').setFontSize(12).setFontWeight('bold');
  overview.getRange('H37:I37').setFontColor('#5f6368').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right');

  overview.insertChart(buildDailyEcpmChart_(overview, helperSheet, 38, 400));
}

/**
 * Automatically creates and formats all 7 workbook tabs (Instructions, Overview,
 * Ad Unit Breakdown, Ad_Unit_Data, Country_Data, Daily_Data, and Chart_Helper),
 * formulas, dropdowns, and embedded charts when run on a blank Google Sheet.
 * Upgrades 3-chart sheets to 4 charts if needed; otherwise no-op when tabs exist.
 */
function setupWorkbook_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  if (
    ss.getSheetByName('Overview') &&
    ss.getSheetByName('Chart_Helper') &&
    ss.getSheetByName('Instructions') &&
    ss.getSheetByName('Ad Unit Breakdown')
  ) {
    try {
      upgradeToFourChartLayoutIfNeeded_(ss);
    } catch (e) {}
    return;
  }

  // 1. Create or configure hidden data and helper sheets first so formula & validation ranges resolve cleanly.
  let unitSheet = ss.getSheetByName('Ad_Unit_Data') || ss.insertSheet('Ad_Unit_Data');
  ensureSheetGridSize_(unitSheet, 200, 10);
  unitSheet.setTabColor('#1a73e8');
  unitSheet.setFrozenRows(1);
  unitSheet.getRange('A1:I1').setValues([[
    'Date', 'Raw_Date', 'Ad_Unit', 'AdX_Revenue', 'OPMC_Revenue',
    'Total_Revenue', 'AdX_Impressions', 'OPMC_Impressions', 'Total_Impressions'
  ]]);
  unitSheet.getRange('A1:I1')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  unitSheet.getRange('D2:F200').setNumberFormat('$#,##0.00');
  unitSheet.getRange('G2:I200').setNumberFormat('#,##0');

  let countrySheet = ss.getSheetByName('Country_Data') || ss.insertSheet('Country_Data');
  ensureSheetGridSize_(countrySheet, 200, 10);
  countrySheet.setTabColor('#f29900');
  countrySheet.setFrozenRows(1);
  countrySheet.getRange('A1:I1').setValues([[
    'Date', 'Raw_Date', 'Country', 'AdX_Revenue', 'OPMC_Revenue',
    'Total_Revenue', 'AdX_Impressions', 'OPMC_Impressions', 'Total_Impressions'
  ]]);
  countrySheet.getRange('A1:I1')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  countrySheet.getRange('D2:F200').setNumberFormat('$#,##0.00');
  countrySheet.getRange('G2:I200').setNumberFormat('#,##0');

  let dailySheet = ss.getSheetByName('Daily_Data') || ss.insertSheet('Daily_Data');
  ensureSheetGridSize_(dailySheet, 405, 8);
  dailySheet.setTabColor('#0f9d58');
  dailySheet.setFrozenRows(1);
  dailySheet.getRange('A1:H1').setValues([[
    'Date', 'Raw_Date', 'Ad Exchange Revenue', 'On-platform multicall Revenue', 'Total Revenue',
    'Ad Exchange Impressions', 'On-platform multicall Impressions', 'Total Impressions'
  ]]);
  dailySheet.getRange('A1:H1')
    .setBackground('#0f9d58')
    .setFontColor('#ffffff')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  dailySheet.getRange('C2:E405').setNumberFormat('$#,##0.00');
  dailySheet.getRange('F2:H405').setNumberFormat('#,##0');

  let helperSheet = ss.getSheetByName('Chart_Helper') || ss.insertSheet('Chart_Helper');
  ensureSheetGridSize_(helperSheet, 405, 20);
  helperSheet.setTabColor('#9aa0a6');
  helperSheet.setFrozenRows(1);
  helperSheet.getRange('A1:S2').setValues([
    [
      'Date', 'Total Revenue', 'Ad Exchange Revenue', 'On-platform multicall Revenue', '',
      'Country', 'Ad Exchange Revenue', 'On-platform multicall Revenue', 'Total Revenue', '',
      'Ad_Unit_Options', 'Country_Options', 'Currency_Code',
      'Ad Exchange Impressions', 'On-platform multicall Impressions', 'Total Impressions',
      'Total eCPM', 'Ad Exchange eCPM', 'On-platform multicall eCPM'
    ],
    ['—', 0, 0, 0, '', '—', 0, 0, 0, '', 'ALL', 'ALL', 'USD', 0, 0, 0, 0, 0, 0]
  ]);
  helperSheet.getRange('B2:D405').setNumberFormat('$#,##0.00');
  helperSheet.getRange('G2:I15').setNumberFormat('$#,##0.00');
  helperSheet.getRange('N2:P405').setNumberFormat('#,##0');
  helperSheet.getRange('Q2:S405').setNumberFormat('$#,##0.00');

  // 2. Build the Instructions tab.
  let instSheet = ss.getSheetByName('Instructions') || ss.insertSheet('Instructions');
  ensureSheetGridSize_(instSheet, 47, 10);
  instSheet.setTabColor('#1a73e8');
  instSheet.setHiddenGridlines(true);

  const instColWidths = [26, 95, 155, 155, 165, 165, 165, 165, 165, 26];
  for (let c = 0; c < instColWidths.length; c++) {
    instSheet.setColumnWidth(c + 1, instColWidths[c]);
  }

  instSheet.getRange('A1:J47')
    .setBackground('#f1f3f4')
    .setFontColor('#202124')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  const instRows = [];
  for (let r = 0; r < 47; r++) {
    instRows.push(['', '', '', '', '', '', '', '', '', '']);
  }
  instRows[1][1] = 'Google Ad Manager Unified Report: User Guide';
  instRows[2][1] = 'Google Sheets & Google Apps Script Reporting Tool';
  instRows[3][1] = 'This tool combines your Ad Exchange revenue and your On-Platform Multicall (OPMC) revenue into one report. Use this workbook to view daily trends, compare country performance, and see total revenue for each ad unit.';

  instRows[5][1] = '▎ 1. Workbook Tabs';
  instRows[6][1] = '📊 Overview';
  instRows[6][3] = 'Run reports, filter data by ad unit or country, view total revenue cards, and check daily revenue, daily eCPM, incremental, and country charts.';
  instRows[7][1] = '🧩 Ad Unit Breakdown';
  instRows[7][3] = 'View revenue for each Parent Ad Unit (A) combined with its OPMC Child Ad Units (A1, A2, A3, A4).';

  instRows[9][1] = '▎ 2. One-Time Setup: Google Cloud API and Apps Script';
  instRows[10][1] = 'Step';
  instRows[10][2] = 'Where to Go';
  instRows[10][4] = 'What to Do';
  instRows[11][1] = 'Step 1';
  instRows[11][2] = 'Google Cloud Console -> APIs & Services';
  instRows[11][4] = 'Open your Google Cloud project and enable the Google Ad Manager API (admanager.googleapis.com).';
  instRows[12][1] = 'Step 2';
  instRows[12][2] = 'Google Cloud Console -> IAM & Admin -> IAM';
  instRows[12][4] = 'Ensure your Google account has the Service Usage Consumer role (roles/serviceusage.serviceUsageConsumer) on that GCP project.';
  instRows[13][1] = 'Step 3';
  instRows[13][2] = 'Google Cloud Console -> Dashboard';
  instRows[13][4] = 'Copy your GCP Project ID (or Project Number) from the Project info card on the Cloud Console Dashboard.';
  instRows[14][1] = 'Step 4';
  instRows[14][2] = 'Google Sheet -> Extensions -> Apps Script';
  instRows[14][4] = 'In a blank Google Sheet, click Extensions -> Apps Script, enable Show "appsscript.json" manifest file in editor under Project Settings, paste Code.js and appsscript.json, and save.';
  instRows[15][1] = 'Step 5';
  instRows[15][2] = 'Apps Script -> Editor (< >) -> Toolbar';
  instRows[15][4] = 'Select authorizeAndVerifyGamScope in the top toolbar, click Run, and approve the permission prompt to build all workbook tabs.';
  instRows[16][1] = 'Step 6';
  instRows[16][2] = 'Google Sheet -> Refresh Page & Cell I7';
  instRows[16][4] = 'Refresh the Google Sheet browser tab once (Cmd+R or F5) to activate the buttons, enter your GCP Project ID in cell I7, and press Enter.';

  instRows[18][1] = '▎ 3. How to Generate a Report (Step-by-Step)';
  instRows[19][1] = 'Step';
  instRows[19][2] = 'Action';
  instRows[19][4] = 'Instructions';
  instRows[20][1] = 'Step 1';
  instRows[20][2] = 'Open the Overview tab';
  instRows[20][4] = 'Click the Overview tab at the bottom of the sheet.';
  instRows[21][1] = 'Step 2';
  instRows[21][2] = 'Enter Network Code and Project ID';
  instRows[21][4] = 'Type your Google Ad Manager Network Code in cell C7 and your GCP Project ID in cell I7, then press Enter.';
  instRows[22][1] = 'Step 3';
  instRows[22][2] = 'Select your date range';
  instRows[22][4] = 'Double-click cell F7 to select the Start Date and double-click cell F8 to select the End Date (YYYY-MM-DD).';
  instRows[23][1] = 'Step 4';
  instRows[23][2] = 'Run the report';
  instRows[23][4] = 'Click the Generate Report button in row 10. Wait for the Report complete message to appear.';

  instRows[25][1] = '▎ 4. How to Use the Buttons and Filters in the "Overview" tab';
  instRows[26][1] = 'Button or Filter';
  instRows[26][3] = 'How to Use It';
  instRows[27][1] = '⚡ Generate Report';
  instRows[27][3] = 'Fetch new data from Google Ad Manager and update all charts and tables.';
  instRows[28][1] = '🔄 Reset Filters';
  instRows[28][3] = 'Reset both the Ad Unit filter and the Country filter back to ALL.';
  instRows[29][1] = '📥 Download CSV';
  instRows[29][3] = 'Download the full report as a CSV file.';
  instRows[30][1] = '🧹 Clear Data';
  instRows[30][3] = 'Wipe all fetched report data from all sheets while preserving your settings and formulas.';
  instRows[31][1] = '📱 Filter by Ad Unit';
  instRows[31][3] = 'Select an ad unit from the dropdown list in cell C12 to show revenue for that ad unit only.';
  instRows[32][1] = '🌍 Filter by Country';
  instRows[32][3] = 'Select a country from the dropdown list in cell F12 to show revenue for that country only.';

  instRows[34][1] = '▎ 5. Key Revenue Terms';
  instRows[35][1] = 'Term';
  instRows[35][3] = 'Meaning';
  instRows[36][1] = 'Ad Exchange Revenue';
  instRows[36][3] = 'Revenue from standard Google Ad Exchange auctions on your ad units.';
  instRows[37][1] = 'OPMC Revenue';
  instRows[37][3] = 'Additional revenue from On-Platform Multicall (OPMC) mediation ad units.';
  instRows[38][1] = 'Total Revenue';
  instRows[38][3] = 'The sum of Ad Exchange Revenue plus OPMC Revenue.';
  instRows[39][1] = 'Parent Ad Unit Rollup';
  instRows[39][3] = 'On the Ad Unit Breakdown tab, Table 1 groups each Parent Ad Unit (A) with its OPMC Child Ad Units (A1 to A4). If you change a parent name in Table 2 (Column D), Table 1 updates automatically.';

  instRows[41][1] = '▎ 6. Version History & Changelog';
  instRows[42][1] = 'Version';
  instRows[42][2] = 'Release Date';
  instRows[42][3] = 'Key Features & Enhancements';
  instRows[43][1] = 'v1.1';
  instRows[43][2] = 'November 2026';
  instRows[43][3] =
    'Daily eCPM Trends & Self-Service Setup\n' +
    '• Daily eCPM Trends Chart: Adds a dedicated Daily eCPM Trends chart (Total, Ad Exchange, and OPMC eCPM) on Overview that updates with the Ad Unit and Country filters.\n' +
    '• Self-Service Workbook Setup: Automatically creates and formats all workbook tabs, formulas, buttons, and charts on first run in a blank Google Sheet.\n' +
    '• Executive Scorecard Subtitles: Shows Share of Total (%) on Ad Exchange and OPMC cards and Incremental Uplift (+%) on Total Revenue.\n' +
    '• Dynamic Chart Scaling & CSV eCPM: Scales daily chart axes to the exact report date range and adds daily eCPM columns to CSV exports.';
  instRows[44][1] = 'v1.0';
  instRows[44][2] = 'October 2026';
  instRows[44][3] =
    'Initial Release (Google Sheets and Apps Script)\n' +
    '• Google Ad Manager API Integration: Gets Ad Exchange and OPMC report data in parallel from admanager.googleapis.com.\n' +
    '• Parent Ad Unit Rollup: Groups each Parent Ad Unit (A) with its OPMC Child Ad Units (A1 to A4) on the Ad Unit Breakdown tab.\n' +
    '• Overview Charts: Shows three charts: Daily Revenue Trends, Stacked Incremental Revenue, and Country Revenue Breakdown.\n' +
    '• Live Filters: Filters the Overview tab by Ad Unit (C12) or Country (F12) and resets filters with Reset Filters (E10).\n' +
    '• CSV Export & Clear Data: Exports a four-section CSV via downloadCsvFromSheet() and clears report data safely via clearReportData().';

  instSheet.getRange(1, 1, instRows.length, 10).setValues(instRows);

  instSheet.setRowHeight(1, 16);
  instSheet.setRowHeight(2, 38);
  instSheet.setRowHeight(3, 24);
  instSheet.setRowHeight(4, 44);
  instSheet.setRowHeight(5, 16);
  instSheet.setRowHeight(6, 32);
  instSheet.setRowHeights(7, 2, 30);
  instSheet.setRowHeight(9, 16);
  instSheet.setRowHeight(10, 32);
  instSheet.setRowHeight(11, 28);
  instSheet.setRowHeights(12, 6, 34);
  instSheet.setRowHeight(18, 16);
  instSheet.setRowHeight(19, 32);
  instSheet.setRowHeight(20, 28);
  instSheet.setRowHeights(21, 4, 30);
  instSheet.setRowHeight(25, 16);
  instSheet.setRowHeight(26, 32);
  instSheet.setRowHeight(27, 28);
  instSheet.setRowHeights(28, 6, 30);
  instSheet.setRowHeight(34, 16);
  instSheet.setRowHeight(35, 32);
  instSheet.setRowHeight(36, 28);
  instSheet.setRowHeights(37, 3, 30);
  instSheet.setRowHeight(40, 44);
  instSheet.setRowHeight(41, 16);
  instSheet.setRowHeight(42, 32);
  instSheet.setRowHeight(43, 28);
  instSheet.setRowHeight(44, 128);
  instSheet.setRowHeight(45, 134);

  // Merge and style Instructions cards.
  ['B2:I2', 'B3:I3', 'B4:I4', 'B6:I6', 'B10:I10', 'B19:I19', 'B26:I26', 'B35:I35', 'B42:I42'].forEach(function(a1) {
    instSheet.getRange(a1).merge();
  });
  [7, 8, 27, 28, 29, 30, 31, 32, 33, 36, 37, 38, 39, 40].forEach(function(r) {
    instSheet.getRange(r, 2, 1, 2).merge();
    instSheet.getRange(r, 4, 1, 6).merge();
  });
  [11, 12, 13, 14, 15, 16, 17, 20, 21, 22, 23, 24].forEach(function(r) {
    instSheet.getRange(r, 3, 1, 2).merge();
    instSheet.getRange(r, 5, 1, 5).merge();
  });
  instSheet.getRange('D43:I43').merge();
  instSheet.getRange('D44:I44').merge();
  instSheet.getRange('D45:I45').merge();

  instSheet.getRange('B2:I4').setBackground('#ffffff');
  instSheet.getRange('B2').setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8');
  instSheet.getRange('B3').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368');
  instSheet.getRange('B4:I4').setBackground('#f8f9fa').setFontSize(10).setFontColor('#3c4043').setWrap(true);
  styleCardBox_(instSheet.getRange('B2:I4'), '#dadce0', '#1a73e8');

  [
    { hdr: 'B6:I6', box: 'B6:I8', tblHdr: null },
    { hdr: 'B10:I10', box: 'B10:I17', tblHdr: 'B11:I11' },
    { hdr: 'B19:I19', box: 'B19:I24', tblHdr: 'B20:I20' },
    { hdr: 'B26:I26', box: 'B26:I33', tblHdr: 'B27:I27' },
    { hdr: 'B35:I35', box: 'B35:I40', tblHdr: 'B36:I36' },
    { hdr: 'B42:I42', box: 'B42:I45', tblHdr: 'B43:I43' }
  ].forEach(function(sec) {
    instSheet.getRange(sec.box).setBackground('#ffffff').setWrap(true);
    instSheet.getRange(sec.box).setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    instSheet.getRange(sec.hdr).setBackground('#ffffff').setFontColor('#1a73e8').setFontSize(12).setFontWeight('bold');
    if (sec.tblHdr) {
      instSheet.getRange(sec.tblHdr).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    }
  });
  instSheet.getRange('B7:C8').setFontWeight('bold');
  instSheet.getRange('B12:B17').setFontWeight('bold').setHorizontalAlignment('center');
  instSheet.getRange('B21:B24').setFontWeight('bold').setHorizontalAlignment('center');
  instSheet.getRange('B28:C33').setFontWeight('bold');
  instSheet.getRange('B37:C40').setFontWeight('bold');
  instSheet.getRange('B44:C45').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('top');
  instSheet.getRange('D44:I45').setVerticalAlignment('top');

  // 3. Build the Overview tab (95 rows, 4 chart sections).
  let overview = ss.getSheetByName('Overview') || ss.insertSheet('Overview');
  ensureSheetGridSize_(overview, 95, 10);
  overview.setTabColor('#1a73e8');
  overview.setHiddenGridlines(true);

  const ovColWidths = [26, 185, 185, 48, 185, 185, 48, 185, 185, 26];
  for (let c = 0; c < ovColWidths.length; c++) {
    overview.setColumnWidth(c + 1, ovColWidths[c]);
  }

  overview.getRange('A1:J95')
    .setBackground('#f1f3f4')
    .setFontColor('#202124')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  const todayIso = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'UTC', 'yyyy-MM-dd');
  const ovValues = [];
  for (let r = 0; r < 95; r++) {
    ovValues.push(['', '', '', '', '', '', '', '', '', '']);
  }

  ovValues[1][1] = 'Google Ad Manager Unified Report';
  ovValues[3][1] =
    'Copyright Google LLC. Supported by Google LLC and/or its affiliate(s). This solution, including any related sample code or data, is made available on an “as is,” “as available,” and “with all faults” basis, solely for illustrative purposes, and without warranty or representation of any kind. This solution is experimental, unsupported and provided solely for your convenience. Your use of it is subject to your agreements with Google, as applicable, and may constitute a beta feature as defined under those agreements. To the extent that you make any data available to Google in connection with your use of the solution, you represent and warrant that you have all necessary and appropriate rights, consents and permissions to permit Google to use and process that data. By using any portion of this solution, you acknowledge, assume and accept all risks, known and unknown, associated with its usage and any processing of data by Google, including with respect to your deployment of any portion of this solution in your systems, or usage in connection with your business, if at all. In connection with this solution, you will not provide to Google any personally identifiable information, personal information or personal data.';

  ovValues[5][1] = '▎ Report Parameters';
  ovValues[5][7] = '=IF(C12<>"ALL", "Filtered dataset records (Ad Unit: "&C12&"): "&TEXT(COUNTIF(Ad_Unit_Data!C2:C, C12), "#,##0"), IF(F12<>"ALL", "Filtered dataset records (Country: "&F12&"): "&TEXT(COUNTIF(Country_Data!C2:C, F12), "#,##0"), "Filtered dataset records (All Ad Units): "&TEXT(COUNTA(Ad_Unit_Data!A2:A), "#,##0")))';

  ovValues[6][1] = 'Network Code';
  ovValues[6][2] = '<INSERT_HERE>';
  ovValues[6][4] = 'Start Date (YYYY-MM-DD)';
  ovValues[6][5] = todayIso;
  ovValues[6][7] = 'GCP Project ID (Quota)';
  ovValues[6][8] = '<INSERT_HERE>';

  ovValues[7][4] = 'End Date (YYYY-MM-DD)';
  ovValues[7][5] = todayIso;

  ovValues[9][1] = '⚡ Generate Report';
  ovValues[9][4] = '🔄 Reset Filters';
  ovValues[9][5] = '🧹 Clear Data';
  ovValues[9][7] = '📥 Download CSV';

  ovValues[10][1] = '📱 Filter by Ad Unit (Select below)';
  ovValues[10][4] = '🌍 Filter by Country (Select below)';
  ovValues[10][7] = '📊 Total Filtered Impressions';

  ovValues[11][1] = 'Ad Unit Filter:';
  ovValues[11][2] = 'ALL';
  ovValues[11][4] = 'Country Filter:';
  ovValues[11][5] = 'ALL';
  ovValues[11][7] = '=IF(C12<>"ALL", SUMIF(Ad_Unit_Data!C2:C, C12, Ad_Unit_Data!I2:I), IF(F12<>"ALL", SUMIF(Country_Data!C2:C, F12, Country_Data!I2:I), SUM(Ad_Unit_Data!I2:I)))';

  ovValues[13][1] = 'AD EXCHANGE REVENUE';
  ovValues[13][3] = '="+"';
  ovValues[13][4] = 'OPMC REVENUE';
  ovValues[13][6] = '="="';
  ovValues[13][7] = 'TOTAL REVENUE';

  ovValues[14][1] = '=IF(C12<>"ALL", SUMIF(Ad_Unit_Data!C2:C, C12, Ad_Unit_Data!D2:D), IF(F12<>"ALL", SUMIF(Country_Data!C2:C, F12, Country_Data!D2:D), SUM(Ad_Unit_Data!D2:D)))';
  ovValues[14][4] = '=IF(C12<>"ALL", SUMIF(Ad_Unit_Data!C2:C, C12, Ad_Unit_Data!E2:E), IF(F12<>"ALL", SUMIF(Country_Data!C2:C, F12, Country_Data!E2:E), SUM(Ad_Unit_Data!E2:E)))';
  ovValues[14][7] = '=B15+E15';

  ovValues[15][1] = '="Share of Total: "&TEXT(IF(H15>0, B15/H15, 0), "0.0%")';
  ovValues[15][4] = '="Share of Total: "&TEXT(IF(H15>0, E15/H15, 0), "0.0%")';
  ovValues[15][7] = '="Incremental Uplift: +"&TEXT(IF(B15>0, E15/B15, 0), "0.0%")';

  ovValues[17][1] = '▎ Daily Revenue Trends';
  ovValues[17][7] = '="Active Filter: "&IF(C12<>"ALL", "Ad Unit = "&C12, IF(F12<>"ALL", "Country = "&F12, "All Inventory & Countries"))';

  ovValues[36][1] = '▎ Daily eCPM Trends';
  ovValues[36][7] = '="Active Filter: "&IF(C12<>"ALL", "Ad Unit = "&C12, IF(F12<>"ALL", "Country = "&F12, "All Inventory & Countries"))';

  ovValues[55][1] = '▎ Stacked Incremental Revenue';

  ovValues[74][1] = '▎ Country Revenue Breakdown (Top 10 Countries)';

  // Format plain-text ID inputs before setting values so IDs never convert to scientific notation.
  overview.getRange('C7:C8').setNumberFormat('@');
  overview.getRange('I7:I8').setNumberFormat('@');
  overview.getRange(1, 1, ovValues.length, 10).setValues(ovValues);
  overview.getRange('F7:F8').setNumberFormat('yyyy-mm-dd');

  // Set exact Overview row heights.
  overview.setRowHeight(1, 16);
  overview.setRowHeights(2, 2, 26);
  overview.setRowHeight(4, 70);
  overview.setRowHeight(5, 16);
  overview.setRowHeight(6, 32);
  overview.setRowHeights(7, 2, 34);
  overview.setRowHeight(9, 14);
  overview.setRowHeight(10, 44);
  overview.setRowHeight(11, 26);
  overview.setRowHeight(12, 34);
  overview.setRowHeight(13, 18);
  overview.setRowHeight(14, 28);
  overview.setRowHeight(15, 52);
  overview.setRowHeight(16, 28);
  overview.setRowHeight(17, 18);
  overview.setRowHeight(18, 34);
  overview.setRowHeights(19, 17, 22);
  overview.setRowHeight(36, 18);
  overview.setRowHeight(37, 34);
  overview.setRowHeights(38, 17, 22);
  overview.setRowHeight(55, 18);
  overview.setRowHeight(56, 34);
  overview.setRowHeights(57, 17, 22);
  overview.setRowHeight(74, 18);
  overview.setRowHeight(75, 34);
  overview.setRowHeights(76, 17, 22);
  overview.setRowHeights(93, 3, 18);

  // Merge Overview ranges.
  [
    'B2:I3', 'B4:I4',
    'B6:G6', 'H6:I6',
    'B7:B8', 'C7:C8', 'H7:H8', 'I7:I8',
    'B10:C10', 'H10:I10',
    'B11:C11', 'E11:F11', 'H11:I11', 'H12:I12',
    'B14:C14', 'B15:C15', 'B16:C16', 'D14:D16',
    'E14:F14', 'E15:F15', 'E16:F16', 'G14:G16',
    'H14:I14', 'H15:I15', 'H16:I16',
    'B18:G18', 'H18:I18',
    'B37:G37', 'H37:I37',
    'B56:G56', 'H56:I56',
    'B75:G75', 'H75:I75'
  ].forEach(function(a1) {
    overview.getRange(a1).merge();
  });

  // Style Header Card (B2:I4).
  overview.getRange('B2:I3').setBackground('#ffffff').setFontColor('#202124').setFontSize(16).setFontWeight('bold');
  overview.getRange('B4:I4').setBackground('#f8f9fa').setFontColor('#5f6368').setFontSize(9).setWrap(true);
  styleCardBox_(overview.getRange('B2:I4'), '#dadce0', '#1a73e8');

  // Style Report Parameters Card (B6:I8).
  overview.getRange('B6:I8').setBackground('#ffffff');
  styleCardBox_(overview.getRange('B6:I8'), '#dadce0');
  overview.getRange('B6:G6').setFontColor('#1a73e8').setFontSize(12).setFontWeight('bold');
  overview.getRange('H6:I6').setBackground('#e8f0fe').setFontColor('#1967d2').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');

  ['B7:B8', 'E7:E8', 'H7:H8'].forEach(function(lblRange) {
    overview.getRange(lblRange).setFontColor('#5f6368').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right');
  });
  ['C7:C8', 'F7', 'F8', 'I7:I8'].forEach(function(inpRange) {
    const r = overview.getRange(inpRange);
    r.setBackground('#f8f9fa').setFontColor('#202124').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');
    r.setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  });

  // Style Row 10 Action Buttons.
  overview.getRange('B10:C10')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, '#0d47a1', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  overview.getRange('E10')
    .setBackground('#ffffff')
    .setFontColor('#1a73e8')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  overview.getRange('F10')
    .setBackground('#ffffff')
    .setFontColor('#d93025')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  overview.getRange('H10:I10')
    .setBackground('#137333')
    .setFontColor('#ffffff')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, '#0b5323', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Style Filters Card (B11:I12).
  overview.getRange('B11:I12').setBackground('#ffffff');
  styleCardBox_(overview.getRange('B11:I12'), '#dadce0');
  overview.getRange('B11:F11').setFontColor('#5f6368').setFontSize(9).setFontWeight('bold');
  overview.getRange('H11:I11').setFontColor('#5f6368').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
  overview.getRange('B12').setFontColor('#5f6368').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right');
  overview.getRange('C12')
    .setBackground('#e8f0fe')
    .setFontColor('#1967d2')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBorder(true, true, true, true, false, false, '#1a73e8', SpreadsheetApp.BorderStyle.SOLID);
  overview.getRange('E12').setFontColor('#5f6368').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right');
  overview.getRange('F12')
    .setBackground('#fef7e0')
    .setFontColor('#b06000')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBorder(true, true, true, true, false, false, '#f29900', SpreadsheetApp.BorderStyle.SOLID);
  overview.getRange('H12:I12')
    .setBackground('#f8f9fa')
    .setFontColor('#202124')
    .setFontSize(12)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setNumberFormat('#,##0');

  // Style Executive Scorecards Row (B14:I16).
  overview.getRange('B14:C16').setBackground('#ffffff');
  overview.getRange('B14:C14').setFontColor('#5f6368').setFontSize(10).setFontWeight('bold');
  overview.getRange('B15:C15').setFontColor('#202124').setFontSize(22).setFontWeight('bold').setNumberFormat('$#,##0.00');
  overview.getRange('B16:C16').setFontColor('#5f6368').setFontSize(8);
  styleCardBox_(overview.getRange('B14:C16'), '#dadce0', '#1a73e8');

  overview.getRange('D14:D16').setBackground('#f1f3f4').setFontColor('#9aa0a6').setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');

  overview.getRange('E14:F16').setBackground('#ffffff');
  overview.getRange('E14:F14').setFontColor('#5f6368').setFontSize(10).setFontWeight('bold');
  overview.getRange('E15:F15').setFontColor('#202124').setFontSize(22).setFontWeight('bold').setNumberFormat('$#,##0.00');
  overview.getRange('E16:F16').setFontColor('#5f6368').setFontSize(8);
  styleCardBox_(overview.getRange('E14:F16'), '#dadce0', '#f29900');

  overview.getRange('G14:G16').setBackground('#f1f3f4').setFontColor('#9aa0a6').setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');

  overview.getRange('H14:I16').setBackground('#fafffa');
  overview.getRange('H14:I14').setFontColor('#5f6368').setFontSize(10).setFontWeight('bold');
  overview.getRange('H15:I15').setFontColor('#0f9d58').setFontSize(22).setFontWeight('bold').setNumberFormat('$#,##0.00');
  overview.getRange('H16:I16').setFontColor('#137333').setFontSize(8).setFontWeight('bold');
  styleCardBox_(overview.getRange('H14:I16'), '#ceead6', '#0f9d58');

  // Style the 4 Chart Containers.
  [
    { box: 'B18:I35', title: 'B18:G18', sub: 'H18:I18', color: '#1a73e8' },
    { box: 'B37:I54', title: 'B37:G37', sub: 'H37:I37', color: '#1a73e8' },
    { box: 'B56:I73', title: 'B56:G56', sub: 'H56:I56', color: '#0f9d58' },
    { box: 'B75:I92', title: 'B75:G75', sub: 'H75:I75', color: '#f29900' }
  ].forEach(function(ch) {
    overview.getRange(ch.box).setBackground('#ffffff');
    styleCardBox_(overview.getRange(ch.box), '#dadce0');
    overview.getRange(ch.title).setFontColor(ch.color).setFontSize(12).setFontWeight('bold');
    overview.getRange(ch.sub).setFontColor('#5f6368').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right');
  });

  // Apply Data Validations on Overview (Calendar pickers on F7:F8 and dynamic dropdowns on C12 & F12).
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();
  overview.getRange('F7:F8').setDataValidation(dateRule);

  const unitRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(helperSheet.getRange('K2:K500'), true)
    .setAllowInvalid(true)
    .build();
  overview.getRange('C12').setDataValidation(unitRule);

  const countryRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(helperSheet.getRange('L2:L300'), true)
    .setAllowInvalid(true)
    .build();
  overview.getRange('F12').setDataValidation(countryRule);

  // Add the 4 Embedded Charts on Overview if not already present.
  if (overview.getCharts().length === 0) {
    const chart1 = overview.newChart()
      .asLineChart()
      .addRange(helperSheet.getRange('A1:D400'))
      .setNumHeaders(1)
      .setPosition(19, 2, 8, 4)
      .setOption('title', 'Daily Revenue Trends')
      .setOption('width', 1190)
      .setOption('height', 355)
      .setOption('legend', { position: 'top', textStyle: { fontName: 'Roboto', fontSize: 11 } })
      .setOption('curveType', 'function')
      .setOption('hAxis', { title: 'Date', textStyle: { fontName: 'Roboto', fontSize: 10 } })
      .setOption('vAxis', { title: 'Revenue', textStyle: { fontName: 'Roboto', fontSize: 10 } })
      .setOption('colors', ['#0f9d58', '#1a73e8', '#f29900'])
      .setOption('series', {
        0: { lineWidth: 3, color: '#0f9d58' },
        1: { lineWidth: 3, color: '#1a73e8' },
        2: { lineWidth: 3, lineDashStyle: [6, 3], color: '#f29900' }
      })
      .setOption('useFirstColumnAsDomain', true)
      .build();
    overview.insertChart(chart1);

    const chartEcpm = buildDailyEcpmChart_(overview, helperSheet, 38, 400);
    overview.insertChart(chartEcpm);

    const chart2 = overview.newChart()
      .asColumnChart()
      .setStacked()
      .addRange(helperSheet.getRange('A1:A400'))
      .addRange(helperSheet.getRange('C1:D400'))
      .setNumHeaders(1)
      .setPosition(57, 2, 8, 4)
      .setOption('width', 1190)
      .setOption('height', 355)
      .setOption('legend', { position: 'top', textStyle: { fontName: 'Roboto', fontSize: 11 } })
      .setOption('hAxis', { title: 'Date', textStyle: { fontName: 'Roboto', fontSize: 10 } })
      .setOption('vAxis', { title: 'Combined Revenue (AdX + OPMC)', textStyle: { fontName: 'Roboto', fontSize: 10 } })
      .setOption('colors', ['#1a73e8', '#f29900'])
      .setOption('useFirstColumnAsDomain', true)
      .build();
    overview.insertChart(chart2);

    const chart3 = overview.newChart()
      .asColumnChart()
      .setStacked()
      .addRange(helperSheet.getRange('F1:H11'))
      .setNumHeaders(1)
      .setPosition(76, 2, 8, 4)
      .setOption('width', 1190)
      .setOption('height', 355)
      .setOption('legend', { position: 'top', textStyle: { fontName: 'Roboto', fontSize: 11 } })
      .setOption('hAxis', { title: 'Country', textStyle: { fontName: 'Roboto', fontSize: 10 } })
      .setOption('vAxis', { title: 'Revenue', textStyle: { fontName: 'Roboto', fontSize: 10 } })
      .setOption('colors', ['#1a73e8', '#f29900'])
      .setOption('useFirstColumnAsDomain', true)
      .build();
    overview.insertChart(chart3);
  }

  // 4. Build and style the Ad Unit Breakdown tab.
  let breakdown = ss.getSheetByName('Ad Unit Breakdown') || ss.insertSheet('Ad Unit Breakdown');
  ensureSheetGridSize_(breakdown, 60, 10);
  breakdown.setTabColor('#0f9d58');
  breakdown.setHiddenGridlines(true);

  const bdColWidths = [26, 340, 360, 290, 165, 165, 130, 130, 165, 26];
  for (let c = 0; c < bdColWidths.length; c++) {
    breakdown.setColumnWidth(c + 1, bdColWidths[c]);
  }
  breakdown.getRange('A1:J60')
    .setBackground('#f1f3f4')
    .setFontColor('#202124')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  breakdown.setRowHeight(1, 16);
  breakdown.setRowHeights(2, 2, 24);
  breakdown.setRowHeight(4, 46);
  breakdown.setRowHeight(5, 16);
  breakdown.setRowHeight(6, 26);
  breakdown.setRowHeight(7, 44);
  breakdown.setRowHeight(8, 26);
  breakdown.setRowHeight(9, 16);
  breakdown.setRowHeight(10, 32);

  [
    'B2:I3', 'B4:I4',
    'B6:C6', 'D6:E6', 'F6:G6', 'H6:I6',
    'B7:C7', 'D7:E7', 'F7:G7', 'H7:I7',
    'B8:C8', 'D8:E8', 'F8:G8', 'H8:I8',
    'B10:G10', 'H10:I10'
  ].forEach(function(a1) {
    breakdown.getRange(a1).merge();
  });

  breakdown.getRange('B2:I3').setBackground('#ffffff').setFontColor('#202124').setFontSize(16).setFontWeight('bold');
  breakdown.getRange('B4:I4').setBackground('#f8f9fa').setFontColor('#5f6368').setFontSize(9).setWrap(true);
  styleCardBox_(breakdown.getRange('B2:I4'), '#dadce0', '#0f9d58');

  [
    { box: 'B6:C8', lbl: 'B6:C6', val: 'B7:C7', sub: 'B8:C8', accent: '#5f6368', valSize: 13, valColor: '#202124', bg: '#ffffff' },
    { box: 'D6:E8', lbl: 'D6:E6', val: 'D7:E7', sub: 'D8:E8', accent: '#1a73e8', valSize: 18, valColor: '#202124', bg: '#ffffff', curr: true },
    { box: 'F6:G8', lbl: 'F6:G6', val: 'F7:G7', sub: 'F8:G8', accent: '#f29900', valSize: 18, valColor: '#202124', bg: '#ffffff', curr: true },
    { box: 'H6:I8', lbl: 'H6:I6', val: 'H7:I7', sub: 'H8:I8', accent: '#0f9d58', valSize: 18, valColor: '#0f9d58', bg: '#fafffa', curr: true }
  ].forEach(function(kpi) {
    breakdown.getRange(kpi.box).setBackground(kpi.bg);
    breakdown.getRange(kpi.lbl).setFontColor('#5f6368').setFontSize(9).setFontWeight('bold');
    breakdown.getRange(kpi.val).setFontColor(kpi.valColor).setFontSize(kpi.valSize).setFontWeight('bold');
    if (kpi.curr) {
      breakdown.getRange(kpi.val).setNumberFormat('$#,##0.00');
    }
    breakdown.getRange(kpi.sub).setFontColor(kpi.accent === '#0f9d58' ? '#137333' : '#5f6368').setFontSize(8).setFontWeight(kpi.accent === '#0f9d58' ? 'bold' : 'normal');
    styleCardBox_(breakdown.getRange(kpi.box), kpi.accent === '#0f9d58' ? '#ceead6' : '#dadce0', kpi.accent);
  });

  breakdown.getRange('B10:G10').setBackground('#ffffff').setFontColor('#1a73e8').setFontSize(12).setFontWeight('bold');
  breakdown.getRange('H10:I10').setBackground('#e8f0fe').setFontColor('#1967d2').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');

  writeParentAdUnitBreakdownSheet_(ss, {
    ad_unit_rows: [],
    currency_code: 'USD',
    currency_symbol: '$'
  });

  // 5. Remove any leftover empty default sheet (e.g., Sheet1) and order tabs cleanly.
  const managedNames = {
    Instructions: true,
    Overview: true,
    'Ad Unit Breakdown': true,
    Ad_Unit_Data: true,
    Country_Data: true,
    Daily_Data: true,
    Chart_Helper: true
  };
  ss.getSheets().forEach(function(s) {
    const nm = s.getName();
    if (!managedNames[nm] && s.getLastRow() === 0 && s.getLastColumn() === 0 && s.getCharts().length === 0) {
      try {
        ss.deleteSheet(s);
      } catch (e) {}
    }
  });

  const desiredOrder = [
    instSheet,
    overview,
    breakdown,
    unitSheet,
    countrySheet,
    dailySheet,
    helperSheet
  ];
  desiredOrder.forEach(function(s, idx) {
    if (s) {
      try {
        ss.setActiveSheet(s);
        ss.moveActiveSheet(idx + 1);
      } catch (e) {}
    }
  });

  [unitSheet, countrySheet, dailySheet, helperSheet].forEach(function(s) {
    if (s) {
      try {
        s.hideSheet();
      } catch (e) {}
    }
  });

  ss.setActiveSheet(overview);
  SpreadsheetApp.flush();
}

/**
 * If the four action buttons do not exist on the Overview tab, adds them to row 10:
 * 1. B10:C10: Generate Report (runReportFromSheetParams)
 * 2. E10: Reset Filters (resetSheetFilters)
 * 3. F10: Clear Data (clearReportData)
 * 4. H10:I10: Download CSV (downloadCsvFromSheet)
 */
function installDashboardButtons_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  const sheet = ss.getSheetByName('Overview') || ss.getSheetByName('Dashboard');
  if (!sheet) return;

  try {
    const g14 = sheet.getRange('G14');
    if (String(g14.getDisplayValue() || '').indexOf('#ERROR') === 0) {
      sheet.getRange('G14:G16').setNumberFormat('General');
      g14.setFormula('="="');
    }
  } catch (e) {}

  const defs = [
    { key: 'generate_report', col: 2, row: 10, width: 345, height: 44, fn: 'runReportFromSheetParams', desc: 'Generate Report' },
    { key: 'reset_filters', col: 5, row: 10, width: 170, height: 44, fn: 'resetSheetFilters', desc: 'Reset Sheet Filters' },
    { key: 'clear_data', col: 6, row: 10, width: 170, height: 44, fn: 'clearReportData', desc: 'Clear Report Data' },
    { key: 'download_csv', col: 8, row: 10, width: 345, height: 44, fn: 'downloadCsvFromSheet', desc: 'Download CSV' }
  ];

  try {
    const existingImages = sheet.getImages();
    const foundKeys = {};
    existingImages.forEach(function(img) {
      const title = String(img.getAltTextTitle() || '');
      if (title.indexOf(BUTTON_MARKER_PREFIX) === 0) {
        foundKeys[title] = true;
      }
    });

    // If all four buttons already exist on the Overview tab, stop here.
    if (
      foundKeys[BUTTON_MARKER_PREFIX + 'generate_report'] &&
      foundKeys[BUTTON_MARKER_PREFIX + 'reset_filters'] &&
      foundKeys[BUTTON_MARKER_PREFIX + 'clear_data'] &&
      foundKeys[BUTTON_MARKER_PREFIX + 'download_csv'] &&
      existingImages.length === 4
    ) {
      return;
    }

    existingImages.forEach(function(img) {
      const title = String(img.getAltTextTitle() || '');
      const fn = String(img.getScript() || '');
      if (
        title.indexOf(BUTTON_MARKER_PREFIX) === 0 ||
        fn === 'runReportFromSheetParams' ||
        fn === 'resetSheetFilters' ||
        fn === 'clearReportData' ||
        fn === 'downloadCsvFromSheet'
      ) {
        img.remove();
      }
    });

    const rawBytes = Utilities.base64Decode(TRANSPARENT_1X1_PNG_BASE64);
    defs.forEach(function(d) {
      const blob = Utilities.newBlob(rawBytes, 'image/png', d.key + '.png');
      const img = sheet.insertImage(blob, d.col, d.row);
      img.setWidth(d.width).setHeight(d.height);
      img.setAltTextTitle(BUTTON_MARKER_PREFIX + d.key);
      img.setAltTextDescription(d.desc);
      img.assignScript(d.fn);
    });
  } catch (err) {}
}

/**
 * Adds the action buttons to the Overview tab when a user opens the spreadsheet.
 */
function onOpen() {
  try {
    installDashboardButtons_();
  } catch (e) {}
}

/**
 * Resets the Ad Unit filter (C12) and the Country filter (F12) on the Overview tab to "ALL".
 */
function resetSheetFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = getTargetDashboardSheet_(ss);
  if (!dash) return;
  dash.getRange('C12').setValue('ALL');
  dash.getRange('F12').setValue('ALL');
  SpreadsheetApp.flush();
}

/**
 * Clears all report data from the visible and hidden tabs.
 * Keeps all formulas, report parameters, dropdown lists, and sheet formatting.
 */
function clearReportData() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Clear All Report Data',
    'Are you sure you want to clear all fetched report data from all sheets?\n\n' +
    'This will wipe all data rows, scorecards, and charts while preserving your formulas, report settings, and workbook structure.',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  // 1. Clear data rows in the hidden Ad_Unit_Data tab and keep the header row.
  const unitSheet = ss.getSheetByName('Ad_Unit_Data');
  if (unitSheet) {
    const lastRow = unitSheet.getLastRow();
    if (lastRow > 1) {
      unitSheet.getRange(2, 1, lastRow - 1, unitSheet.getLastColumn()).clearContent();
    }
  }

  // 2. Clear data rows in the hidden Country_Data tab and keep the header row.
  const countrySheet = ss.getSheetByName('Country_Data');
  if (countrySheet) {
    const lastRow = countrySheet.getLastRow();
    if (lastRow > 1) {
      countrySheet.getRange(2, 1, lastRow - 1, countrySheet.getLastColumn()).clearContent();
    }
  }

  // 3. Clear data rows in the hidden Daily_Data tab and keep the header row.
  const dailySheet = ss.getSheetByName('Daily_Data');
  if (dailySheet) {
    const lastRow = dailySheet.getLastRow();
    if (lastRow > 1) {
      dailySheet.getRange(2, 1, lastRow - 1, dailySheet.getLastColumn()).clearContent();
    }
  }

  // 4. Reset the hidden Chart_Helper tab.
  const helperSheet = ss.getSheetByName('Chart_Helper');
  if (helperSheet) {
    // Clear daily helper rows in columns A to D and eCPM columns N to S, and seed row 2 with 0s so chart series stay bound.
    const clearRows = Math.max(helperSheet.getLastRow(), 400);
    helperSheet.getRange(2, 1, clearRows, 4).clearContent();
    helperSheet.getRange('A2:D2').setValues([['—', 0, 0, 0]]);
    helperSheet.getRange('B2:D400').setNumberFormat('$#,##0.00');

    if (helperSheet.getMaxColumns() >= 19) {
      helperSheet.getRange(2, 14, clearRows, 6).clearContent();
      helperSheet.getRange('N1:S2').setValues([
        [
          'Ad Exchange Impressions', 'On-platform multicall Impressions', 'Total Impressions',
          'Total eCPM', 'Ad Exchange eCPM', 'On-platform multicall eCPM'
        ],
        [0, 0, 0, 0, 0, 0]
      ]);
      helperSheet.getRange('N2:P400').setNumberFormat('#,##0');
      helperSheet.getRange('Q2:S400').setNumberFormat('$#,##0.00');
    }

    // Reset the top 10 countries table in columns F to I and seed row 2 with 0s.
    helperSheet.getRange('F1:I1').setValues([['Country', 'Ad Exchange Revenue', 'On-platform multicall Revenue', 'Total Revenue']]);
    helperSheet.getRange(2, 6, 15, 4).clearContent();
    helperSheet.getRange('F2:I2').setValues([['—', 0, 0, 0]]);
    helperSheet.getRange('G2:I15').setNumberFormat('$#,##0.00');

    const dashForCharts = getTargetDashboardSheet_(ss);
    if (dashForCharts) {
      syncOverviewCharts_(dashForCharts, helperSheet, 0, 'USD');
    }

    // Reset filter dropdown lists in columns K and L to "ALL".
    helperSheet.getRange(1, 11, Math.max(helperSheet.getLastRow(), 150), 1).clearContent();
    helperSheet.getRange('K1:K2').setValues([['Ad_Unit_Options'], ['ALL']]);

    helperSheet.getRange(1, 12, Math.max(helperSheet.getLastRow(), 250), 1).clearContent();
    helperSheet.getRange('L1:L2').setValues([['Country_Options'], ['ALL']]);

    // Set the default currency code in column M to "USD".
    helperSheet.getRange('M1:M2').setValues([['Currency_Code'], ['USD']]);
  }

  // 5. Clear data rows in the Ad Unit Breakdown tab.
  try {
    writeParentAdUnitBreakdownSheet_(ss, {
      ad_unit_rows: [],
      currency_code: 'USD',
      currency_symbol: '$'
    });
  } catch (err) {}

  // 6. Reset the filter cells and scorecard currency format on the Overview tab.
  const dash = getTargetDashboardSheet_(ss);
  if (dash) {
    dash.getRange('C12').setValue('ALL');
    dash.getRange('F12').setValue('ALL');
    try {
      dash.getRange('B15').setNumberFormat('$#,##0.00');
      dash.getRange('E15').setNumberFormat('$#,##0.00');
      dash.getRange('H15').setNumberFormat('$#,##0.00');
    } catch (e) {}
  }

  // 7. Hide the data and helper tabs.
  [unitSheet, countrySheet, dailySheet, helperSheet].forEach(function(s) {
    if (s) {
      try { s.hideSheet(); } catch (e) {}
    }
  });

  SpreadsheetApp.flush();

  ui.alert(
    'Report Data Cleared',
    'All publisher report data has been cleared from all sheets.\n\n' +
    'The workbook formulas, settings, and structure are preserved and ready for the next report.',
    ui.ButtonSet.OK
  );
}

/**
 * Runs the matching function when a user selects a button cell in row 10 on the Overview tab.
 */
function onSelectionChange(e) {
  if (!e || !e.range) return;
  try {
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    if (row !== 10) return;

    const sheet = range.getSheet();
    const sName = sheet.getName();
    if (sName !== 'Overview' && sName !== 'Dashboard') {
      const b7 = String(sheet.getRange('B7').getDisplayValue() || '').trim();
      if (b7 !== 'Network Code') return;
    }

    let topCol = col;
    if (col === 2 || col === 3) topCol = 2;
    else if (col === 5) topCol = 5;
    else if (col === 6) topCol = 6;
    else if (col === 8 || col === 9) topCol = 8;
    else return;

    const cellText = String(sheet.getRange(10, topCol).getDisplayValue() || '').trim();
    if (cellText.indexOf('Reset Filters') !== -1) {
      sheet.getRange('C12').setValue('ALL');
      sheet.getRange('F12').setValue('ALL');
      sheet.getRange('C12').activate();
      SpreadsheetApp.flush();
    } else if (
      cellText.indexOf('Clear') !== -1 ||
      cellText.indexOf('Download') !== -1 ||
      cellText.indexOf('Generate Report') !== -1
    ) {
      // Simple triggers run in AuthMode.LIMITED and cannot call ScriptApp.getOAuthToken() or UI modals.
      // If clicking row 10 selected the underlying cell instead of the floating button, prompt a one-time refresh.
      sheet.getRange('C7').activate();
      const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
      if (ss) {
        ss.toast(
          'Please refresh this Google Sheet tab once (Cmd+R or F5) to activate the row 10 action buttons.',
          '🔄 Refresh Tab Once',
          10
        );
      }
    }
  } catch (err) {}
}

/**
 * Checks the OAuth2 access token and returns its granted scopes.
 */
function inspectTokenScopes_(accessToken) {
  if (!accessToken) return { valid: false, hasAdManagerScope: false, scopes: '' };
  try {
    const res = UrlFetchApp.fetch(
      'https://www.googleapis.com/oauth2/v1/tokeninfo?access_token=' + encodeURIComponent(accessToken),
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() === 200) {
      const info = JSON.parse(res.getContentText() || '{}');
      const scopes = String(info.scope || '');
      const scopeList = scopes.split(/\s+/);
      const hasGam = scopeList.indexOf('https://www.googleapis.com/auth/admanager') !== -1;
      return {
        valid: true,
        hasAdManagerScope: hasGam,
        scopes: scopes,
        expiresIn: info.expires_in
      };
    }
  } catch (e) {}
  return { valid: false, hasAdManagerScope: false, scopes: '' };
}

/**
 * Gets the active user's OAuth2 access token and checks that it includes the Ad Manager scope.
 */
function acquireValidGamToken_(explicitCustomToken) {
  const custom = String(explicitCustomToken || '').trim();
  if (custom && custom.indexOf('{') !== 0 && custom.indexOf('1//') !== 0) {
    const checkCustom = inspectTokenScopes_(custom);
    if (checkCustom.valid && checkCustom.hasAdManagerScope) {
      return { token: custom, source: 'Custom OAuth Token', scopes: checkCustom.scopes };
    }
  }

  let scriptToken = '';
  try {
    scriptToken = ScriptApp.getOAuthToken() || '';
  } catch (e) {}

  if (scriptToken) {
    const checkScript = inspectTokenScopes_(scriptToken);
    if (checkScript.hasAdManagerScope) {
      return { token: scriptToken, source: 'ScriptApp.getOAuthToken()', scopes: checkScript.scopes };
    }
    throw new Error(
      'MISSING appsscript.json MANIFEST (OAUTH SCOPE MISSING):\n' +
        'Your workbook tabs were created, but Apps Script only has default scopes: [' + (checkScript.scopes || 'none') + ']\n' +
        'and is missing "https://www.googleapis.com/auth/admanager".\n' +
        'To fix this:\n' +
        '1. Click Project Settings (gear icon on the left) and check Show "appsscript.json" manifest file in editor.\n' +
        '2. Return to Editor (< >), open appsscript.json, paste the repository appsscript.json, and click Save.\n' +
        '3. Click ▷ Run on authorizeAndVerifyGamScope again and approve the permission prompt.'
    );
  }

  throw new Error('Could not acquire any OAuth2 token. Please authorize the script.');
}

/**
 * Cleans the Google Cloud Project ID value from cell I7 (or legacy cell I8).
 * Removes spaces, strips any "projects/" prefix, and converts text to lowercase.
 */
function normalizeQuotaProject_(val) {
  return String(val || '')
    .trim()
    .replace(/^projects\//i, '')
    .trim()
    .toLowerCase();
}

/**
 * Checks whether a value is a valid Google Cloud Project ID and not a placeholder or three-letter currency code.
 */
function isValidQuotaProject_(val) {
  const raw = String(val || '').trim();
  if (
    !raw ||
    raw.toUpperCase() === '<INSERT_HERE>' ||
    /^[A-Z]{3}$/.test(raw) ||
    CURRENCY_SYMBOLS[raw.toUpperCase()]
  ) {
    return false;
  }
  const s = normalizeQuotaProject_(raw);
  if (!s || s.length < 4) return false;
  return /^[a-z0-9:._-]+$/.test(s);
}

/**
 * Reads a date cell and returns the date string in YYYY-MM-DD format.
 */
function readDateCellIso_(range, ss) {
  if (!range) return '';
  const disp = String(range.getDisplayValue() || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(disp)) {
    return disp;
  }
  try {
    const raw = range.getValue();
    if (raw instanceof Date && !isNaN(raw.getTime())) {
      const tz = (ss && ss.getSpreadsheetTimeZone()) || Session.getScriptTimeZone() || 'UTC';
      return Utilities.formatDate(raw, tz, 'yyyy-MM-dd');
    }
  } catch (e) {}
  return disp;
}

/**
 * Gets the Overview tab from the active spreadsheet.
 */
function getTargetDashboardSheet_(ss) {
  const targetSs = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!targetSs) return null;
  try {
    const active = targetSs.getActiveSheet();
    if (active && (active.getName() === 'Overview' || active.getName() === 'Dashboard')) {
      return active;
    }
  } catch (e) {}
  const overview = targetSs.getSheetByName('Overview') || targetSs.getSheetByName('Dashboard');
  if (overview) return overview;
  try {
    const active = targetSs.getActiveSheet();
    if (active && String(active.getRange('B7').getDisplayValue() || '').trim() === 'Network Code') {
      return active;
    }
  } catch (e) {}
  return null;
}

/**
 * Creates a CSV file from the current report data and starts the download in the browser.
 */
function downloadCsvFromSheet() {
  const cached = readExistingSheetDataForModal_();
  const params = readParamsFromDashboardSheet_();
  if (!cached || !cached.ad_unit_rows || cached.ad_unit_rows.length === 0) {
    SpreadsheetApp.getUi().alert('No Report Data', 'Please generate a report first before downloading CSV.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const curr = cached.currency_code || 'USD';
  const netCode = String(params.network_code || 'report').trim();
  const fileName = 'ad_manager_unified_report_' + netCode + '.csv';

  const rollup = buildParentAdUnitRollup_(cached.ad_unit_rows || []);
  const lines = [];
  lines.push('Parent Ad Unit Revenue Summary');
  lines.push('Parent Ad Unit,OPMC Child Units Count,Associated OPMC Ad Units,Ad Exchange Revenue (' + curr + '),OPMC Revenue (' + curr + '),Total Revenue (' + curr + '),AdX Share (%),OPMC Share (%),Total Impressions');
  (rollup.parentRows || []).forEach(function(pr) {
    lines.push([
      '"' + String(pr.parentName || '').replace(/"/g, '""') + '"',
      pr.childCount,
      '"' + String(pr.childSummary || '').replace(/"/g, '""') + '"',
      pr.adxRev,
      pr.opmcRev,
      pr.totalRev,
      (pr.adxShare * 100).toFixed(1) + '%',
      (pr.opmcShare * 100).toFixed(1) + '%',
      pr.totalImps
    ].join(','));
  });

  // Group daily revenue and impression totals by date for Stacked Incremental Revenue & eCPM.
  const dailyMap = {};
  (cached.ad_unit_rows || []).forEach(function(r) {
    const dKey = String(r.raw_date || r.date || '').trim();
    if (!dKey) return;
    if (!dailyMap[dKey]) {
      dailyMap[dKey] = {
        date: r.date,
        raw_date: dKey,
        adxRev: 0,
        opmcRev: 0,
        totalRev: 0,
        adxImps: 0,
        opmcImps: 0,
        totalImps: 0
      };
    }
    const ar = Number(r.adx_revenue) || 0;
    const or = Number(r.opmc_revenue) || 0;
    const tr = Number(r.total_revenue) || (ar + or);
    const ai = Number(r.adx_impressions) || 0;
    const oi = Number(r.opmc_impressions) || 0;
    const ti = Number(r.total_impressions) || (ai + oi);
    dailyMap[dKey].adxRev += ar;
    dailyMap[dKey].opmcRev += or;
    dailyMap[dKey].totalRev += tr;
    dailyMap[dKey].adxImps += ai;
    dailyMap[dKey].opmcImps += oi;
    dailyMap[dKey].totalImps += ti;
  });
  const sortedDailyKeys = Object.keys(dailyMap).sort();

  lines.push('');
  lines.push('Daily Incremental Revenue & eCPM');
  lines.push('Date,Ad Exchange Revenue (' + curr + '),OPMC Incremental Revenue (' + curr + '),Total Revenue (' + curr + '),AdX Share (%),OPMC Share (%),OPMC Incremental Uplift (%),Ad Exchange eCPM (' + curr + '),OPMC eCPM (' + curr + '),Total eCPM (' + curr + '),Total Impressions');
  sortedDailyKeys.forEach(function(k) {
    const d = dailyMap[k];
    const ar = Math.round(d.adxRev * 100) / 100;
    const or = Math.round(d.opmcRev * 100) / 100;
    const tr = Math.round((ar + or) * 100) / 100;
    const adxPct = tr > 0 ? ((ar / tr) * 100).toFixed(1) + '%' : '0.0%';
    const opmcPct = tr > 0 ? ((or / tr) * 100).toFixed(1) + '%' : '0.0%';
    const upliftPct = ar > 0 ? '+' + ((or / ar) * 100).toFixed(1) + '%' : '0.0%';
    const adxEcpm = d.adxImps > 0 ? ((d.adxRev / d.adxImps) * 1000).toFixed(2) : '0.00';
    const opmcEcpm = d.opmcImps > 0 ? ((d.opmcRev / d.opmcImps) * 1000).toFixed(2) : '0.00';
    const totEcpm = d.totalImps > 0 ? ((d.totalRev / d.totalImps) * 1000).toFixed(2) : '0.00';
    lines.push([
      d.date,
      ar,
      or,
      tr,
      adxPct,
      opmcPct,
      upliftPct,
      adxEcpm,
      opmcEcpm,
      totEcpm,
      Math.round(d.totalImps)
    ].join(','));
  });

  lines.push('');
  lines.push('Daily Revenue Trends');
  lines.push('Date,Ad Unit Name,Ad Exchange Revenue (' + curr + '),OPMC Revenue (' + curr + '),Total Revenue (' + curr + '),Total Impressions');
  (cached.ad_unit_rows || []).forEach(function(r) {
    lines.push([
      r.date,
      '"' + String(r.ad_unit || '').replace(/"/g, '""') + '"',
      r.adx_revenue,
      r.opmc_revenue,
      r.total_revenue,
      r.total_impressions
    ].join(','));
  });

  lines.push('');
  lines.push('Country Revenue Breakdown');
  lines.push('Date,Country,Ad Exchange Revenue (' + curr + '),OPMC Revenue (' + curr + '),Total Revenue (' + curr + '),Total Impressions');
  (cached.country_rows || []).forEach(function(r) {
    lines.push([
      r.date,
      '"' + String(r.country || '').replace(/"/g, '""') + '"',
      r.adx_revenue,
      r.opmc_revenue,
      r.total_revenue,
      r.total_impressions
    ].join(','));
  });

  const csvString = lines.join('\n');
  const b64Csv = Utilities.base64Encode(Utilities.newBlob(csvString, 'text/csv', fileName).getBytes());

  const htmlContent =
    '<!DOCTYPE html><html><head><style>' +
    'body { font-family: "Google Sans", Arial, sans-serif; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; margin: 0; color: #202124; }' +
    '.msg { font-size: 14px; font-weight: 600; color: #137333; margin-bottom: 8px; }' +
    '.sub { font-size: 12px; color: #5f6368; }' +
    '</style></head><body>' +
    '<div class="msg">📥 Downloading CSV...</div>' +
    '<div class="sub">' + fileName + '</div>' +
    '<script>' +
    '(function() {' +
    '  var raw = atob("' + b64Csv + '");' +
    '  var bytes = new Uint8Array(raw.length);' +
    '  for (var i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);' +
    '  var blob = new Blob([bytes], { type: "text/csv;charset=utf-8;" });' +
    '  var url = URL.createObjectURL(blob);' +
    '  var a = document.createElement("a");' +
    '  a.href = url;' +
    '  a.download = "' + fileName + '";' +
    '  document.body.appendChild(a);' +
    '  a.click();' +
    '  document.body.removeChild(a);' +
    '  setTimeout(function() { google.script.host.close(); }, 1200);' +
    '})();' +
    '</script></body></html>';

  const htmlOut = HtmlService.createHtmlOutput(htmlContent).setWidth(360).setHeight(130);
  SpreadsheetApp.getUi().showModalDialog(htmlOut, 'Download CSV');
}

/**
 * If a user selects an Ad Unit filter (C12) or a Country filter (F12), resets the other filter to "ALL".
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sName = sheet.getName();
  if (sName !== 'Overview' && sName !== 'Dashboard') {
    const b7 = String(sheet.getRange('B7').getDisplayValue() || '').trim();
    if (b7 !== 'Network Code') return;
  }
  const a1 = e.range.getA1Notation();
  const val = String(e.range.getValue() || '').trim();
  if (a1 === 'C12' && val && val !== 'ALL') {
    sheet.getRange('F12').setValue('ALL');
  } else if (a1 === 'F12' && val && val !== 'ALL') {
    sheet.getRange('C12').setValue('ALL');
  }
}

/**
 * Reads the report parameters from the Overview tab:
 * - C7: Network Code
 * - F7: Start Date (YYYY-MM-DD)
 * - F8: End Date (YYYY-MM-DD)
 * - I7 (or legacy I8): Google Cloud Project ID
 */
function readParamsFromDashboardSheet_() {
  const defaults = {
    sheet_name: 'Overview',
    network_code: '',
    start_date: '',
    end_date: '',
    api_mode: 'PUBLIC',
    quota_project: '',
    custom_token: ''
  };

  try {
    SpreadsheetApp.flush();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return defaults;
    const dash = getTargetDashboardSheet_(ss);
    if (!dash) return defaults;

    const rawNetCode = String(dash.getRange('C7').getDisplayValue() || '').trim();
    const netCode = rawNetCode.toUpperCase() === '<INSERT_HERE>' ? '' : rawNetCode;
    const startVal = readDateCellIso_(dash.getRange('F7'), ss);
    const e8Label = String(dash.getRange('E8').getDisplayValue() || '').toUpperCase();
    const isStackedDatesLayout = e8Label.indexOf('END') !== -1;
    const h7Label = String(dash.getRange('H7').getDisplayValue() || '').toUpperCase();

    const endVal = readDateCellIso_(dash.getRange(isStackedDatesLayout ? 'F8' : 'I7'), ss);
    let qpVal = String(dash.getRange('I8').getDisplayValue() || '').trim();
    if (
      !isValidQuotaProject_(qpVal) &&
      isStackedDatesLayout &&
      (h7Label.indexOf('PROJECT') !== -1 || h7Label.indexOf('QUOTA') !== -1)
    ) {
      qpVal = String(dash.getRange('I7').getDisplayValue() || '').trim();
    }
    if (!isValidQuotaProject_(qpVal)) {
      const overviewSheet = ss.getSheetByName('Overview');
      if (overviewSheet && overviewSheet.getSheetId() !== dash.getSheetId()) {
        qpVal = String(overviewSheet.getRange('I8').getDisplayValue() || '').trim();
        if (!isValidQuotaProject_(qpVal)) {
          const ovE8 = String(overviewSheet.getRange('E8').getDisplayValue() || '').toUpperCase();
          const ovH7 = String(overviewSheet.getRange('H7').getDisplayValue() || '').toUpperCase();
          if (
            ovE8.indexOf('END') !== -1 &&
            (ovH7.indexOf('PROJECT') !== -1 || ovH7.indexOf('QUOTA') !== -1)
          ) {
            qpVal = String(overviewSheet.getRange('I7').getDisplayValue() || '').trim();
          }
        }
      }
    }

    const apiMode = 'PUBLIC';

    const resolvedQuotaProject = isValidQuotaProject_(qpVal) ? normalizeQuotaProject_(qpVal) : '';

    return {
      sheet_name: dash.getName(),
      network_code: netCode,
      start_date: startVal,
      end_date: endVal,
      api_mode: apiMode,
      quota_project: resolvedQuotaProject,
      custom_token: ''
    };
  } catch (err) {
    return defaults;
  }
}

/**
 * Reads the parameters from the Overview tab, checks that required values exist, and runs the report.
 */
function runReportFromSheetParams() {
  try {
    const params = readParamsFromDashboardSheet_();
    if (!params.network_code) {
      throw new Error('Missing Network Code: Please enter your Google Ad Manager Network Code in cell C7 on the Overview tab.');
    }
    if (!params.start_date || !params.end_date) {
      throw new Error('Missing Date Range: Please select both Start Date (cell F7) and End Date (cell F8) on the Overview tab.');
    }
    if (!params.quota_project) {
      throw new Error(
        'Missing GCP Project ID (Quota): Please enter your GCP Project ID in cell I7 on the Overview tab (and press Enter to commit the cell) before clicking Generate Report.'
      );
    }
    runUnifiedReportPipeline(params);
    SpreadsheetApp.getUi().alert(
      'Report complete',
      'Unified Report generated successfully',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert('Report Generation Error', String(err.message || err), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Builds the API configuration and HTTP headers using the OAuth2 token and Google Cloud Project ID.
 */
function resolveApiConfig_(params) {
  const tokObj = acquireValidGamToken_(params && params.custom_token);
  const token = tokObj.token;


  const rawQp = String((params && params.quota_project) || '').trim();
  const quotaProject = isValidQuotaProject_(rawQp) ? normalizeQuotaProject_(rawQp) : '';

  if (!quotaProject) {
    throw new Error(
      'Missing GCP Project ID (Quota): Please enter your GCP Project ID in cell I7 on the Overview tab (and press Enter to commit the cell) before running the report.'
    );
  }

  const headers = {
    Authorization: 'Bearer ' + token,
    'x-goog-user-project': quotaProject
  };

  return {

    quotaProject: quotaProject,
    baseUrl: PUBLIC_GAM_BASE_URL,
    headers: headers,
    tokenSource: tokObj.source,
    tokenScopes: tokObj.scopes,
    modeLabel: 'Public API Mode (Publisher OAuth • Quota: ' + quotaProject + ')'
  };
}

/**
 * Builds an HTTP request object for UrlFetchApp.
 */
function buildFetchRequest_(url, method, payloadObj, apiConfig) {
  const req = {
    url: url,
    method: (method || 'get').toLowerCase(),
    headers: Object.assign({}, apiConfig.headers),
    contentType: 'application/json',
    muteHttpExceptions: true
  };
  if (payloadObj !== null && payloadObj !== undefined) {
    req.payload = JSON.stringify(payloadObj);
  }
  return req;
}

/**
 * Sends a batch of HTTP requests to the Google Ad Manager API and retries failed requests.
 */
function executeBatchRequests_(requests, apiConfig, retries) {
  const maxRetries = Math.max(retries || 3, 1);
  let currentRequests = requests.map(function(r) {
    return Object.assign({}, r);
  });
  const configuredQuotaProject =
    (apiConfig && apiConfig.quotaProject) ||
    (currentRequests.length > 0 && currentRequests[0].headers && currentRequests[0].headers['x-goog-user-project']) ||
    '<none>';

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const responses = UrlFetchApp.fetchAll(currentRequests);
    const parsed = [];
    let shouldRetry = false;
    let lastErrorMsg = '';

    for (let i = 0; i < responses.length; i++) {
      const res = responses[i];
      const code = res.getResponseCode();
      const bodyText = res.getContentText() || '';
      if (code >= 200 && code < 300) {
        try {
          parsed.push(bodyText ? JSON.parse(bodyText) : {});
        } catch (jsonErr) {
          shouldRetry = true;
          lastErrorMsg = 'Non-JSON response (' + code + '): ' + bodyText.substring(0, 300);
        }
      } else if (code === 429 || code >= 500) {
        shouldRetry = true;
        lastErrorMsg = 'HTTP ' + code + ': ' + bodyText.substring(0, 400);
      } else {
        if (
          (code === 403 || code === 400) &&
          (bodyText.indexOf('USER_PROJECT') !== -1 ||
            bodyText.indexOf('not found or deleted') !== -1 ||
            bodyText.indexOf('serviceUsageConsumer') !== -1 ||
            bodyText.indexOf('permission to use project') !== -1 ||
            bodyText.indexOf('has been deleted') !== -1 ||
            bodyText.indexOf('SERVICE_DISABLED') !== -1 ||
            bodyText.indexOf('has not been used in project') !== -1)
        ) {
          throw new Error(
            'GAM API Quota Project Error (HTTP ' + code + ' on ' +
              currentRequests[i].url +
              '):\nConfigured GCP Project ID (Overview!I7): ' +
              configuredQuotaProject +
              '\nToken Scopes: [' +
              ((apiConfig && apiConfig.tokenScopes) || 'none') +
              ']\n\nRaw API response: ' +
              bodyText.substring(0, 500)
          );
        }
        if (code === 403 && bodyText.indexOf('ACCESS_TOKEN_SCOPE_INSUFFICIENT') !== -1) {
          throw new Error(
            'GAM API Error (HTTP 403 ACCESS_TOKEN_SCOPE_INSUFFICIENT on ' +
              currentRequests[i].url +
              '):\nToken Source: ' +
              ((apiConfig && apiConfig.tokenSource) || 'Unknown') +
              '\nToken Scopes: [' +
              ((apiConfig && apiConfig.tokenScopes) || 'none') +
              ']\nMake sure appsscript.json has "https://www.googleapis.com/auth/admanager" (without dfp), Save, and click ▷ Run on authorizeAndVerifyGamScope.\nRaw API response: ' +
              bodyText.substring(0, 400)
          );
        }
        throw new Error('GAM API Error (HTTP ' + code + ' on ' + currentRequests[i].url + '): ' + bodyText.substring(0, 500));
      }
    }

    if (!shouldRetry && parsed.length === currentRequests.length) {
      return parsed;
    }
    if (attempt < maxRetries) {
      Utilities.sleep(1000 * attempt);
    } else {
      throw new Error('GAM API batch request failed after ' + maxRetries + ' attempts: ' + lastErrorMsg);
    }
  }
}

/**
 * Gets the primitive value from a Google Ad Manager API value object.
 */
function extractVal_(vDict) {
  if (!vDict || typeof vDict !== 'object') return null;
  const keys = ['stringValue', 'intValue', 'doubleValue', 'boolValue', 'microsValue'];
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    if (vDict[k] !== undefined && vDict[k] !== null) {
      return k === 'intValue' ? parseInt(vDict[k], 10) : vDict[k];
    }
  }
  return null;
}

/**
 * Formats a date string (YYYYMMDD or YYYY-MM-DD) as DD-Mon-YYYY for chart labels.
 */
function formatDateLabel_(cleanDateStr) {
  const s = String(cleanDateStr || '').trim();
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  if (/^\d{8}$/.test(s)) {
    const y = s.substring(0, 4);
    const m = parseInt(s.substring(4, 6), 10);
    const d = s.substring(6, 8);
    if (m >= 1 && m <= 12) return d + '-' + months[m - 1] + '-' + y;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const parts = s.split('-');
    const m = parseInt(parts[1], 10);
    if (m >= 1 && m <= 12) return parts[2] + '-' + months[m - 1] + '-' + parts[0];
  }
  return s;
}

/**
 * Converts a YYYY-MM-DD string into an object with year, month, and day numbers.
 */
function parseDateObj_(dateStr) {
  const parts = String(dateStr || '').trim().split('-');
  if (parts.length !== 3) {
    throw new Error('Invalid date format "' + dateStr + '". Expected YYYY-MM-DD.');
  }
  return {
    year: parseInt(parts[0], 10),
    month: parseInt(parts[1], 10),
    day: parseInt(parts[2], 10)
  };
}

/**
 * Splits an ad unit name into lowercase words to match parent and child ad units.
 */
function tokenizeAdUnitForParentMatch_(name) {
  let s = String(name || '').replace(/\$\d+(?:\.\d+)?/g, ' ');
  s = s.replace(/([a-z])([A-Z])/g, '$1 $2');
  s = s.replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2');
  const raw = s.split(/[^a-zA-Z0-9]+/).filter(Boolean).map(function(t) { return t.toLowerCase(); });
  const core = raw.filter(function(t) {
    if (FORMAT_AND_NOISE_WORDS[t]) return false;
    if (/^\d+$/.test(t)) return false;
    if (/^\d+x\d+$/.test(t)) return false;
    return true;
  });
  return {
    tokens: core,
    joined: core.join('')
  };
}

/**
 * Sorts child ad unit names by their base name and trailing number.
 */
function compareOpmcChildUnitsNaturally_(a, b) {
  const sa = String(a || '');
  const sb = String(b || '');
  const stemA = sa.replace(/[\s_|/-]*\$?\d+(?:\.\d+)?$/, '').trim();
  const stemB = sb.replace(/[\s_|/-]*\$?\d+(?:\.\d+)?$/, '').trim();
  if (stemA !== stemB) {
    return stemA < stemB ? -1 : 1;
  }
  const ma = sa.match(/(\d+(?:\.\d+)?)\s*$/);
  const mb = sb.match(/(\d+(?:\.\d+)?)\s*$/);
  if (ma && mb) {
    const na = parseFloat(ma[1]);
    const nb = parseFloat(mb[1]);
    if (na !== nb) return na - nb;
  }
  return sa < sb ? -1 : (sa > sb ? 1 : 0);
}

/**
 * Groups child ad units under their parent ad units and calculates total revenue and impressions.
 */
function buildParentAdUnitRollup_(adUnitRows) {
  const unitAgg = {};
  (adUnitRows || []).forEach(function(r) {
    const u = String(r.ad_unit || '').trim();
    if (!u) return;
    if (!unitAgg[u]) {
      unitAgg[u] = { adxRev: 0, opmcRev: 0, totRev: 0, adxImps: 0, opmcImps: 0, totImps: 0 };
    }
    const ar = Number(r.adx_revenue) || 0;
    const or = Number(r.opmc_revenue) || 0;
    const tr = Number(r.total_revenue) || (ar + or);
    const ai = Number(r.adx_impressions) || 0;
    const oi = Number(r.opmc_impressions) || 0;
    const ti = Number(r.total_impressions) || (ai + oi);
    unitAgg[u].adxRev += ar;
    unitAgg[u].opmcRev += or;
    unitAgg[u].totRev += tr;
    unitAgg[u].adxImps += ai;
    unitAgg[u].opmcImps += oi;
    unitAgg[u].totImps += ti;
  });

  const allUnits = Object.keys(unitAgg);
  const parents = [];
  const children = [];

  allUnits.forEach(function(u) {
    const v = unitAgg[u];
    const isOpmc = u.toLowerCase().indexOf('opmc') !== -1 || v.opmcRev > 0 || v.opmcImps > 0;
    if (isOpmc) {
      children.push(u);
    } else {
      parents.push(u);
    }
  });

  const parentInfo = {};
  parents.forEach(function(p) {
    const t = tokenizeAdUnitForParentMatch_(p);
    const tokMap = {};
    t.tokens.forEach(function(tok) { tokMap[tok] = true; });
    const lowerP = p.toLowerCase();
    parentInfo[p] = {
      tokens: t.tokens,
      tokenSet: tokMap,
      joined: t.joined,
      adxRev: unitAgg[p].adxRev,
      isStagingOrDirect: lowerP.indexOf('staging') !== -1 || lowerP.indexOf('direct') !== -1
    };
  });

  function matchChildToParent(cName) {
    let stem = String(cName || '').replace(/^opmc[\s_|-]*/i, '');
    stem = stem.replace(/[\s_|-]*\$\d+(?:\.\d+)?$/, '');
    stem = stem.replace(/[\s_|-]*\d+$/, '').trim();
    for (let i = 0; i < parents.length; i++) {
      if (parents[i].toLowerCase() === stem.toLowerCase()) {
        return parents[i];
      }
    }

    const ct = tokenizeAdUnitForParentMatch_(cName);
    if (!ct.tokens || ct.tokens.length === 0) {
      return String(cName || '').replace(/[\s_|/-]*\$\d+(?:\.\d+)?$/, '').trim() + ' (Shared OPMC)';
    }

    let bestP = null;
    let bestExact = -1;
    let bestRatio = -1;
    let bestNotStaging = -1;
    let bestAdxRev = -1;

    for (let j = 0; j < parents.length; j++) {
      const p = parents[j];
      const pInfo = parentInfo[p];
      if (!pInfo.tokens || pInfo.tokens.length === 0) continue;

      let exactCore = 0;
      let overlapRatio = 0;

      if (ct.joined === pInfo.joined) {
        exactCore = 2;
        overlapRatio = 1.0;
      } else {
        let common = 0;
        for (let k = 0; k < ct.tokens.length; k++) {
          if (pInfo.tokenSet[ct.tokens[k]]) common++;
        }
        const minJoinedLen = Math.min(ct.joined.length, pInfo.joined.length);
        if (
          common === 0 &&
          minJoinedLen >= 5 &&
          (ct.joined.indexOf(pInfo.joined) !== -1 || pInfo.joined.indexOf(ct.joined) !== -1)
        ) {
          exactCore = 1;
          overlapRatio = 0.85;
        } else {
          if (common === 0) continue;
          const childCov = common / ct.tokens.length;
          const parentCov = common / pInfo.tokens.length;
          if (childCov < 1.0) continue;
          overlapRatio = Math.round((childCov * 0.7 + parentCov * 0.3) * 10000) / 10000;
          exactCore = childCov === 1.0 && parentCov === 1.0 ? 1 : 0;
        }
      }

      const notStaging = pInfo.isStagingOrDirect ? 0 : 1;
      const adxRev = pInfo.adxRev;

      if (
        exactCore > bestExact ||
        (exactCore === bestExact && overlapRatio > bestRatio) ||
        (exactCore === bestExact && overlapRatio === bestRatio && notStaging > bestNotStaging) ||
        (exactCore === bestExact && overlapRatio === bestRatio && notStaging === bestNotStaging && adxRev > bestAdxRev)
      ) {
        bestExact = exactCore;
        bestRatio = overlapRatio;
        bestNotStaging = notStaging;
        bestAdxRev = adxRev;
        bestP = p;
      }
    }

    return bestP || (String(cName || '').replace(/[\s_|/-]*\$\d+(?:\.\d+)?$/, '').trim() + ' (Shared OPMC)');
  }

  const parentGroups = {};
  parents.forEach(function(p) {
    parentGroups[p] = { children: [] };
  });
  children.sort(compareOpmcChildUnitsNaturally_);
  children.forEach(function(c) {
    const p = matchChildToParent(c);
    if (!parentGroups[p]) {
      parentGroups[p] = { children: [] };
    }
    parentGroups[p].children.push(c);
  });

  const parentRows = [];
  Object.keys(parentGroups).forEach(function(p) {
    const clist = parentGroups[p].children;
    const pStats = unitAgg[p] || { adxRev: 0, opmcRev: 0, totRev: 0, totImps: 0 };
    let cAdx = 0;
    let cOpmc = 0;
    let cImps = 0;
    clist.forEach(function(c) {
      const cs = unitAgg[c] || { adxRev: 0, opmcRev: 0, totImps: 0 };
      cAdx += cs.adxRev;
      cOpmc += cs.opmcRev;
      cImps += cs.totImps;
    });
    const totAdx = Math.round((pStats.adxRev + cAdx) * 100) / 100;
    const totOpmc = Math.round((pStats.opmcRev + cOpmc) * 100) / 100;
    const totRev = Math.round((totAdx + totOpmc) * 100) / 100;
    const totImps = Math.round(pStats.totImps + cImps);

    let childSummary = 'None (Parent Ad Unit Only)';
    if (clist.length === 1) {
      childSummary = '1 OPMC unit (' + clist[0] + ')';
    } else if (clist.length > 1) {
      childSummary = clist.length + ' OPMC units (' + clist[0] + ' .. ' + clist[clist.length - 1] + ')';
    }

    parentRows.push({
      parentName: p,
      childCount: clist.length,
      children: clist,
      childSummary: childSummary,
      adxRev: totAdx,
      opmcRev: totOpmc,
      totalRev: totRev,
      adxShare: totRev > 0 ? totAdx / totRev : 0,
      opmcShare: totRev > 0 ? totOpmc / totRev : 0,
      totalImps: totImps
    });
  });

  parentRows.sort(function(a, b) {
    if (b.totalRev !== a.totalRev) return b.totalRev - a.totalRev;
    return b.totalImps - a.totalImps;
  });

  const detailRows = [];
  parentRows.forEach(function(pr) {
    const p = pr.parentName;
    if (unitAgg[p]) {
      const ps = unitAgg[p];
      const pAr = Math.round(ps.adxRev * 100) / 100;
      const pOr = Math.round(ps.opmcRev * 100) / 100;
      const pTr = Math.round((pAr + pOr) * 100) / 100;
      detailRows.push({
        displayUnit: '🏠 ' + p,
        rawUnit: p,
        role: 'Parent Ad Unit (A)',
        assignedParent: p,
        adxRev: pAr,
        opmcRev: pOr,
        totalRev: pTr,
        opmcShare: pTr > 0 ? pOr / pTr : 0,
        totalImps: Math.round(ps.totImps)
      });
    }
    pr.children.forEach(function(c) {
      const cs = unitAgg[c] || { adxRev: 0, opmcRev: 0, totImps: 0 };
      const cAr = Math.round(cs.adxRev * 100) / 100;
      const cOr = Math.round(cs.opmcRev * 100) / 100;
      const cTr = Math.round((cAr + cOr) * 100) / 100;
      detailRows.push({
        displayUnit: '   ↳ ' + c,
        rawUnit: c,
        role: 'OPMC Child Unit (A1..An)',
        assignedParent: p,
        adxRev: cAr,
        opmcRev: cOr,
        totalRev: cTr,
        opmcShare: cTr > 0 ? cOr / cTr : 0,
        totalImps: Math.round(cs.totImps)
      });
    });
  });

  return {
    parentRows: parentRows,
    detailRows: detailRows,
    totalParents: parents.length,
    totalChildren: children.length
  };
}

/**
 * Writes the summary table (Table 1) and the detailed mapping table (Table 2) to the Ad Unit Breakdown tab.
 */
function writeParentAdUnitBreakdownSheet_(ss, data) {
  if (!ss || !data) return;
  let sheet = ss.getSheetByName('Ad Unit Breakdown');
  if (!sheet) {
    sheet = ss.insertSheet('Ad Unit Breakdown', 1);
  }
  try {
    sheet.showSheet();
  } catch (e) {}

  const rollup = buildParentAdUnitRollup_(data.ad_unit_rows || []);
  const parentRows = rollup.parentRows || [];
  const detailRows = rollup.detailRows || [];

  const numParents = parentRows.length;
  const t1HeaderRow = 11;
  const t1TotalRow = 12;
  const t1StartRow = 13;
  const t1EndRow = Math.max(t1StartRow, t1StartRow + numParents - 1);

  const t2SectionRow = t1EndRow + 3;
  const t2HeaderRow = t2SectionRow + 1;
  const t2StartRow = t2HeaderRow + 1;
  const t2EndRow = Math.max(t2StartRow, t2StartRow + detailRows.length - 1);

  const totalGridRows = t2EndRow + 4;
  const grid = [];
  for (let r = 0; r < totalGridRows; r++) {
    grid.push(['', '', '', '', '', '', '', '', '', '']);
  }

  let existingTop = null;
  try {
    if (sheet.getLastRow() >= 10) {
      existingTop = sheet.getRange(1, 1, 10, 10).getValues();
    }
  } catch (e) {}

  if (existingTop && existingTop.length === 10) {
    // Keep existing header cells in rows 1 to 10.
    for (let r = 0; r < 10; r++) {
      for (let c = 0; c < 10; c++) {
        grid[r][c] = existingTop[r][c] !== undefined ? existingTop[r][c] : '';
      }
    }
  } else {
    grid[1][1] = 'Parent Ad Unit Revenue Breakdown';
    grid[3][1] =
      'Table 1 combines Ad Exchange revenue for each Parent Ad Unit (A) with OPMC revenue from its Child Ad Units (A1 to A4). Table 2 lists each ad unit and its assigned Parent Ad Unit. To change a parent group, edit Column D in Table 2 and Table 1 updates automatically.';
    grid[5][1] = 'PARENT & OPMC INVENTORY';
    grid[5][3] = 'TOTAL AD EXCHANGE REVENUE';
    grid[5][5] = 'TOTAL OPMC REVENUE';
    grid[5][7] = 'TOTAL UNIFIED REVENUE';
    grid[7][1] = 'Automatic Hierarchy & Naming Match (A ↔ A1..An)';
    grid[9][1] = '▎ Table 1: Parent Ad Unit Revenue Summary (Parent A + Associated OPMC Units A1..An)';
    grid[9][7] = 'Live Formula Rollup from Table 2';
  }

  // Write summary formulas in rows 7 and 8.
  grid[6][1] = numParents + ' Parent Groups  •  ' + rollup.totalChildren + ' OPMC Units';
  grid[6][3] = '=D' + t1TotalRow;
  grid[6][5] = '=E' + t1TotalRow;
  grid[6][7] = '=F' + t1TotalRow;

  grid[7][3] = '="AdX Share: "&TEXT(G' + t1TotalRow + ', "0.0%")';
  grid[7][5] = '="OPMC Share: "&TEXT(H' + t1TotalRow + ', "0.0%")';
  grid[7][7] = '="OPMC Uplift over AdX: +"&TEXT(IF(D' + t1TotalRow + '>0, E' + t1TotalRow + '/D' + t1TotalRow + ', 0), "0.0%")';

  grid[t1HeaderRow - 1][1] = 'Parent Ad Unit (A)';
  grid[t1HeaderRow - 1][2] = 'Associated OPMC Child Units (A1..An)';
  grid[t1HeaderRow - 1][3] = 'Ad Exchange Revenue';
  grid[t1HeaderRow - 1][4] = 'OPMC Revenue (A1..An)';
  grid[t1HeaderRow - 1][5] = 'Total Unified Revenue';
  grid[t1HeaderRow - 1][6] = 'AdX Share (%)';
  grid[t1HeaderRow - 1][7] = 'OPMC Share (%)';
  grid[t1HeaderRow - 1][8] = 'Total Impressions';

  // Write the grand total row for Table 1 (row 12).
  grid[t1TotalRow - 1][1] = '🏆 GRAND TOTAL (ALL PARENT AD UNITS)';
  grid[t1TotalRow - 1][2] = rollup.totalChildren + ' Total OPMC Child Ad Units';
  if (numParents > 0) {
    grid[t1TotalRow - 1][3] = '=SUM(D' + t1StartRow + ':D' + t1EndRow + ')';
    grid[t1TotalRow - 1][4] = '=SUM(E' + t1StartRow + ':E' + t1EndRow + ')';
    grid[t1TotalRow - 1][5] = '=D' + t1TotalRow + '+E' + t1TotalRow;
    grid[t1TotalRow - 1][6] = '=IF(F' + t1TotalRow + '>0, D' + t1TotalRow + '/F' + t1TotalRow + ', 0)';
    grid[t1TotalRow - 1][7] = '=IF(F' + t1TotalRow + '>0, E' + t1TotalRow + '/F' + t1TotalRow + ', 0)';
    grid[t1TotalRow - 1][8] = '=SUM(I' + t1StartRow + ':I' + t1EndRow + ')';
  } else {
    grid[t1TotalRow - 1][3] = 0;
    grid[t1TotalRow - 1][4] = 0;
    grid[t1TotalRow - 1][5] = 0;
    grid[t1TotalRow - 1][6] = 0;
    grid[t1TotalRow - 1][7] = 0;
    grid[t1TotalRow - 1][8] = 0;
  }

  // Write the parent ad unit rows for Table 1.
  for (let i = 0; i < numParents; i++) {
    const rNum = t1StartRow + i;
    const pr = parentRows[i];
    grid[rNum - 1][1] = pr.parentName;
    grid[rNum - 1][2] = pr.childSummary;
    grid[rNum - 1][3] = '=SUMIF($D$' + t2StartRow + ':$D$' + t2EndRow + ', B' + rNum + ', $E$' + t2StartRow + ':$E$' + t2EndRow + ')';
    grid[rNum - 1][4] = '=SUMIF($D$' + t2StartRow + ':$D$' + t2EndRow + ', B' + rNum + ', $F$' + t2StartRow + ':$F$' + t2EndRow + ')';
    grid[rNum - 1][5] = '=D' + rNum + '+E' + rNum;
    grid[rNum - 1][6] = '=IF(F' + rNum + '>0, D' + rNum + '/F' + rNum + ', 0)';
    grid[rNum - 1][7] = '=IF(F' + rNum + '>0, E' + rNum + '/F' + rNum + ', 0)';
    grid[rNum - 1][8] = '=SUMIF($D$' + t2StartRow + ':$D$' + t2EndRow + ', B' + rNum + ', $I$' + t2StartRow + ':$I$' + t2EndRow + ')';
  }

  // Write the section header and column headers for Table 2.
  grid[t2SectionRow - 1][1] = '▎ Table 2: Detailed Parent → OPMC Child Breakdown & Mapping (A & Individual A1, A2, A3, A4 Units)';

  grid[t2HeaderRow - 1][1] = 'Individual Ad Unit Name (A or A1..An)';
  grid[t2HeaderRow - 1][2] = 'Ad Unit Role';
  grid[t2HeaderRow - 1][3] = 'Assigned Parent Ad Unit (A)';
  grid[t2HeaderRow - 1][4] = 'Ad Exchange Revenue';
  grid[t2HeaderRow - 1][5] = 'OPMC Revenue';
  grid[t2HeaderRow - 1][6] = 'Total Revenue';
  grid[t2HeaderRow - 1][7] = 'OPMC Share (%)';
  grid[t2HeaderRow - 1][8] = 'Total Impressions';

  for (let j = 0; j < detailRows.length; j++) {
    const rNum = t2StartRow + j;
    const dr = detailRows[j];
    grid[rNum - 1][1] = dr.displayUnit;
    grid[rNum - 1][2] = dr.role;
    grid[rNum - 1][3] = dr.assignedParent;
    grid[rNum - 1][4] = dr.adxRev;
    grid[rNum - 1][5] = dr.opmcRev;
    grid[rNum - 1][6] = '=E' + rNum + '+F' + rNum;
    grid[rNum - 1][7] = '=IF(G' + rNum + '>0, F' + rNum + '/G' + rNum + ', 0)';
    grid[rNum - 1][8] = dr.totalImps;
  }

  sheet.clearContents();
  ensureSheetGridSize_(sheet, Math.max(totalGridRows, 40), 10);
  sheet.getRange(1, 1, grid.length, 10).setValues(grid);

  try {
    const cSym = String(data.currency_symbol || '$').replace(/"/g, '');
    const currFmt = cSym === '$' ? '$#,##0.00' : '"' + cSym + '"#,##0.00';

    // Update top KPI currency format for non-USD networks.
    sheet.getRange('D7').setNumberFormat(currFmt);
    sheet.getRange('F7').setNumberFormat(currFmt);
    sheet.getRange('H7').setNumberFormat(currFmt);

    // Reset dynamic table area below row 10 so resized reports leave zero stale borders or fills.
    const resetRows = Math.max(sheet.getMaxRows() - 10, 1);
    const dynamicArea = sheet.getRange(11, 1, resetRows, 10);
    dynamicArea.breakApart();
    dynamicArea
      .setBackground('#f1f3f4')
      .setFontColor('#202124')
      .setFontFamily('Google Sans')
      .setFontSize(10)
      .setFontWeight('normal')
      .setBorder(false, false, false, false, false, false);

    // Style Table 1 header (row 11) and Grand Total row (row 12).
    sheet.setRowHeight(t1HeaderRow, 32);
    sheet
      .getRange(t1HeaderRow, 2, 1, 8)
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, true, true, '#1557b0', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(t1HeaderRow, 2, 1, 2).setHorizontalAlignment('left');
    sheet.getRange(t1HeaderRow, 4, 1, 6).setHorizontalAlignment('right');

    sheet.setRowHeight(t1TotalRow, 32);
    sheet
      .getRange(t1TotalRow, 2, 1, 8)
      .setBackground('#e6f4ea')
      .setFontColor('#137333')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, true, true, '#ceead6', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(t1TotalRow, 2, 1, 2).setHorizontalAlignment('left');
    sheet.getRange(t1TotalRow, 4, 1, 6).setHorizontalAlignment('right');

    const t1FormattedRows = numParents + 1;
    sheet.getRange(t1TotalRow, 4, t1FormattedRows, 3).setNumberFormat(currFmt);
    sheet.getRange(t1TotalRow, 7, t1FormattedRows, 2).setNumberFormat('0.0%');
    sheet.getRange(t1TotalRow, 9, t1FormattedRows, 1).setNumberFormat('#,##0');

    if (numParents > 0) {
      sheet.setRowHeights(t1StartRow, numParents, 28);
      const t1Body = sheet.getRange(t1StartRow, 2, numParents, 8);
      t1Body
        .setBackground('#ffffff')
        .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(t1StartRow, 2, numParents, 1).setFontWeight('bold').setHorizontalAlignment('left');
      sheet.getRange(t1StartRow, 3, numParents, 1).setFontColor('#5f6368').setHorizontalAlignment('left');
      sheet.getRange(t1StartRow, 4, numParents, 6).setHorizontalAlignment('right');
      sheet.getRange(t1StartRow, 6, numParents, 1).setFontWeight('bold').setFontColor('#0f9d58');
    }

    // Style Table 2 section banner and header row.
    sheet.setRowHeight(t2SectionRow, 32);
    const t2Banner = sheet.getRange(t2SectionRow, 2, 1, 8);
    t2Banner.merge();
    t2Banner
      .setBackground('#ffffff')
      .setFontColor('#1a73e8')
      .setFontSize(12)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);

    sheet.setRowHeight(t2HeaderRow, 32);
    sheet
      .getRange(t2HeaderRow, 2, 1, 8)
      .setBackground('#202124')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, true, true, '#3c4043', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(t2HeaderRow, 2, 1, 3).setHorizontalAlignment('left');
    sheet.getRange(t2HeaderRow, 5, 1, 5).setHorizontalAlignment('right');

    if (detailRows.length > 0) {
      sheet.setRowHeights(t2StartRow, detailRows.length, 26);
      const t2Bgs = [];
      const t2Weights = [];
      for (let dIdx = 0; dIdx < detailRows.length; dIdx++) {
        const isParentRow = detailRows[dIdx].role.indexOf('Parent') === 0;
        const rowBg = isParentRow ? '#f8f9fa' : '#ffffff';
        const rowWeight = isParentRow ? 'bold' : 'normal';
        t2Bgs.push([rowBg, rowBg, '#e8f0fe', rowBg, rowBg, rowBg, rowBg, rowBg]);
        t2Weights.push([rowWeight, rowWeight, 'bold', rowWeight, rowWeight, 'bold', rowWeight, rowWeight]);
      }
      const t2Body = sheet.getRange(t2StartRow, 2, detailRows.length, 8);
      t2Body
        .setBackgrounds(t2Bgs)
        .setFontWeights(t2Weights)
        .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(t2StartRow, 2, detailRows.length, 3).setHorizontalAlignment('left');
      sheet.getRange(t2StartRow, 4, detailRows.length, 1).setFontColor('#1967d2');
      sheet.getRange(t2StartRow, 5, detailRows.length, 5).setHorizontalAlignment('right');
      sheet.getRange(t2StartRow, 5, detailRows.length, 3).setNumberFormat(currFmt);
      sheet.getRange(t2StartRow, 8, detailRows.length, 1).setNumberFormat('0.0%');
      sheet.getRange(t2StartRow, 9, detailRows.length, 1).setNumberFormat('#,##0');
    }
  } catch (fmtErr) {}
}

/**
 * Creates and runs the Ad Exchange and OPMC reports, waits for completion,
 * gets all result rows, and writes the combined data to the spreadsheet.
 */
function runUnifiedReportPipeline(params) {
  const effectiveParams = params || readParamsFromDashboardSheet_();
  const networkCode = String(effectiveParams.network_code || '').trim();
  const startDateStr = String(effectiveParams.start_date || '').trim();
  const endDateStr = String(effectiveParams.end_date || '').trim();

  if (!networkCode) {
    throw new Error('Missing Network Code: Please enter your Google Ad Manager Network Code in cell C7 on the Overview tab.');
  }
  if (!startDateStr || !endDateStr) {
    throw new Error('Missing Date Range: Please select both Start Date (cell F7) and End Date (cell F8) on the Overview tab.');
  }

  const apiConfig = resolveApiConfig_(effectiveParams);
  const startObj = parseDateObj_(startDateStr);
  const endObj = parseDateObj_(endDateStr);
  const baseUrl = apiConfig.baseUrl;
  const ts = new Date().getTime();

  const reportSpecs = [
    {
      key: 'h_unit',
      reportType: 'HISTORICAL',
      dimensions: ['DATE', 'AD_UNIT_NAME'],
      metrics: ['AD_EXCHANGE_IMPRESSIONS', 'AD_EXCHANGE_REVENUE'],
      expandedCompatibility: true
    },
    {
      key: 'o_unit',
      reportType: 'ON_PLATFORM_MULTICALL',
      dimensions: ['DATE', 'AD_UNIT_NAME'],
      metrics: ['ON_PLATFORM_MULTIPLE_CALL_IMPRESSIONS', 'ON_PLATFORM_MULTIPLE_CALL_REVENUE'],
      expandedCompatibility: false
    },
    {
      key: 'h_country',
      reportType: 'HISTORICAL',
      dimensions: ['DATE', 'COUNTRY_NAME'],
      metrics: ['AD_EXCHANGE_IMPRESSIONS', 'AD_EXCHANGE_REVENUE'],
      expandedCompatibility: true
    },
    {
      key: 'o_country',
      reportType: 'ON_PLATFORM_MULTICALL',
      dimensions: ['DATE', 'COUNTRY_NAME'],
      metrics: ['ON_PLATFORM_MULTIPLE_CALL_IMPRESSIONS', 'ON_PLATFORM_MULTIPLE_CALL_REVENUE'],
      expandedCompatibility: false
    }
  ];

  let currencyCode = 'USD';
  try {
    const netReq = buildFetchRequest_(baseUrl + '/networks/' + networkCode, 'GET', null, apiConfig);
    const netRes = executeBatchRequests_([netReq], apiConfig, 2)[0];
    if (netRes && netRes.currencyCode) {
      currencyCode = String(netRes.currencyCode).trim().toUpperCase();
    }
  } catch (e) {
    const errMsg = String((e && e.message) || e || '');
    if (
      errMsg.indexOf('GAM API Quota Project Error') !== -1 ||
      errMsg.indexOf('ACCESS_TOKEN_SCOPE_INSUFFICIENT') !== -1 ||
      errMsg.indexOf('OAUTH SCOPE MISSING') !== -1
    ) {
      throw e;
    }
    currencyCode = 'USD';
  }
  const currSymbol = CURRENCY_SYMBOLS[currencyCode] || (currencyCode + ' ');

  const createUrl = baseUrl + '/networks/' + networkCode + '/reports';
  const createRequests = reportSpecs.map(function(spec, idx) {
    const reportDef = {
      dimensions: spec.dimensions,
      metrics: spec.metrics,
      dateRange: {
        fixed: {
          startDate: startObj,
          endDate: endObj
        }
      },
      reportType: spec.reportType
    };
    if (spec.expandedCompatibility) {
      reportDef.expandedCompatibility = true;
    }
    const payload = {
      displayName: 'AppsScript_' + spec.reportType + '_' + idx + '_' + ts,
      reportDefinition: reportDef
    };
    return buildFetchRequest_(createUrl, 'POST', payload, apiConfig);
  });

  const createdReports = executeBatchRequests_(createRequests, apiConfig, 3);
  const reportNames = createdReports.map(function(r) {
    if (!r || !r.name) {
      throw new Error('Failed to create report: ' + JSON.stringify(r));
    }
    return r.name;
  });

  const runRequests = reportNames.map(function(reportName) {
    return buildFetchRequest_(baseUrl + '/' + reportName + ':run', 'POST', {}, apiConfig);
  });
  let operations = executeBatchRequests_(runRequests, apiConfig, 3);

  const maxPolls = 90;
  for (let poll = 0; poll < maxPolls; poll++) {
    const pendingIndices = [];
    const pollRequests = [];
    for (let i = 0; i < operations.length; i++) {
      if (!operations[i].done) {
        pendingIndices.push(i);
        pollRequests.push(buildFetchRequest_(baseUrl + '/' + operations[i].name, 'GET', null, apiConfig));
      }
    }
    if (pendingIndices.length === 0) break;
    Utilities.sleep(2000);
    const pollResponses = executeBatchRequests_(pollRequests, apiConfig, 3);
    for (let j = 0; j < pendingIndices.length; j++) {
      operations[pendingIndices[j]] = pollResponses[j];
    }
  }

  const resultNames = operations.map(function(op, i) {
    if (!op.done) {
      throw new Error('Report operation timed out: ' + op.name);
    }
    if (op.error) {
      throw new Error('Report operation failed: ' + JSON.stringify(op.error));
    }
    const respData = op.response || {};
    const opParts = String(op.name || '').split('/');
    const opId = opParts[opParts.length - 1];
    return respData.reportResult || (reportNames[i] + '/results/' + opId);
  });

  const allRowsByReport = [[], [], [], []];
  let pageTokens = [null, null, null, null];
  let activeReportIndices = [0, 1, 2, 3];

  while (activeReportIndices.length > 0) {
    const fetchRequests = activeReportIndices.map(function(idx) {
      let fetchUrl = baseUrl + '/' + resultNames[idx] + ':fetchRows?pageSize=5000';
      if (pageTokens[idx]) {
        fetchUrl += '&pageToken=' + encodeURIComponent(pageTokens[idx]);
      }
      return buildFetchRequest_(fetchUrl, 'GET', null, apiConfig);
    });

    const pageResults = executeBatchRequests_(fetchRequests, apiConfig, 3);
    const nextActive = [];

    for (let k = 0; k < activeReportIndices.length; k++) {
      const rIdx = activeReportIndices[k];
      const resData = pageResults[k] || {};
      const rows = resData.rows || [];
      for (let r = 0; r < rows.length; r++) {
        allRowsByReport[rIdx].push(rows[r]);
      }
      if (resData.nextPageToken) {
        pageTokens[rIdx] = resData.nextPageToken;
        nextActive.push(rIdx);
      }
    }
    activeReportIndices = nextActive;
  }

  const output = aggregateStreams_(
    currencyCode,
    currSymbol,
    allRowsByReport[0],
    allRowsByReport[1],
    allRowsByReport[2],
    allRowsByReport[3],
    apiConfig.modeLabel
  );

  try {
    writeResultsToSheets_(output, {
      network_code: networkCode,
      start_date: startDateStr,
      end_date: endDateStr,
      api_mode: params && params.api_mode,
      quota_project: apiConfig.quotaProject
    });
  } catch (sheetErr) {
    output.sheetSyncWarning = sheetErr.message;
  }

  return output;
}

/**
 * Combines the Ad Exchange and OPMC report rows by date, ad unit, and country.
 */
function aggregateStreams_(currencyCode, currSymbol, rawHUnit, rawOUnit, rawHCountry, rawOCountry, modeLabel) {
  const adUnitDict = {};

  function addAdUnitRow(r, isOpmc) {
    const dimVals = (r.dimensionValues || []).map(extractVal_);
    const date = dimVals.length > 0 && dimVals[0] !== null ? String(dimVals[0]) : 'All';
    const adUnit = dimVals.length > 1 && dimVals[1] !== null ? String(dimVals[1]) : 'All';

    const metricGroups = r.metricValueGroups || [];
    const mets = metricGroups.length > 0 ? (metricGroups[0].primaryValues || []).map(extractVal_) : [0, 0];
    const imps = mets.length > 0 && mets[0] !== null ? Number(mets[0]) : 0.0;
    const rev = mets.length > 1 && mets[1] !== null ? Number(mets[1]) : 0.0;

    const key = date + '|||' + adUnit;
    if (!adUnitDict[key]) {
      adUnitDict[key] = { date: date, ad_unit: adUnit, adx_rev: 0.0, opmc_rev: 0.0, adx_imps: 0.0, opmc_imps: 0.0 };
    }
    if (isOpmc) {
      adUnitDict[key].opmc_rev += rev;
      adUnitDict[key].opmc_imps += imps;
    } else {
      adUnitDict[key].adx_rev += rev;
      adUnitDict[key].adx_imps += imps;
    }
  }

  (rawHUnit || []).forEach(function(r) { addAdUnitRow(r, false); });
  (rawOUnit || []).forEach(function(r) { addAdUnitRow(r, true); });

  const countryDict = {};

  function addCountryRow(r, isOpmc) {
    const dimVals = (r.dimensionValues || []).map(extractVal_);
    const date = dimVals.length > 0 && dimVals[0] !== null ? String(dimVals[0]) : 'All';
    const country = dimVals.length > 1 && dimVals[1] !== null ? String(dimVals[1]) : 'Unknown';

    const metricGroups = r.metricValueGroups || [];
    const mets = metricGroups.length > 0 ? (metricGroups[0].primaryValues || []).map(extractVal_) : [0, 0];
    const imps = mets.length > 0 && mets[0] !== null ? Number(mets[0]) : 0.0;
    const rev = mets.length > 1 && mets[1] !== null ? Number(mets[1]) : 0.0;

    const key = date + '|||' + country;
    if (!countryDict[key]) {
      countryDict[key] = { date: date, country: country, adx_rev: 0.0, opmc_rev: 0.0, adx_imps: 0.0, opmc_imps: 0.0 };
    }
    if (isOpmc) {
      countryDict[key].opmc_rev += rev;
      countryDict[key].opmc_imps += imps;
    } else {
      countryDict[key].adx_rev += rev;
      countryDict[key].adx_imps += imps;
    }
  }

  (rawHCountry || []).forEach(function(r) { addCountryRow(r, false); });
  (rawOCountry || []).forEach(function(r) { addCountryRow(r, true); });

  let totalAdxRev = 0.0;
  let totalOpmcRev = 0.0;
  let totalAdxImps = 0.0;
  let totalOpmcImps = 0.0;

  const dailyDict = {};
  const adUnitRows = [];
  const countryRows = [];
  const adUnitSet = {};
  const countrySet = {};

  Object.keys(adUnitDict).forEach(function(k) {
    const val = adUnitDict[k];
    const adxRev = val.adx_rev;
    const opmcRev = val.opmc_rev;
    const totRev = adxRev + opmcRev;
    const adxImp = val.adx_imps;
    const opmcImp = val.opmc_imps;
    const totImp = adxImp + opmcImp;

    totalAdxRev += adxRev;
    totalOpmcRev += opmcRev;
    totalAdxImps += adxImp;
    totalOpmcImps += opmcImp;

    const cleanD = String(val.date).trim();
    if (!dailyDict[cleanD]) {
      dailyDict[cleanD] = { adx: 0.0, opmc: 0.0, tot: 0.0, adx_imps: 0.0, opmc_imps: 0.0, tot_imps: 0.0 };
    }
    dailyDict[cleanD].adx += adxRev;
    dailyDict[cleanD].opmc += opmcRev;
    dailyDict[cleanD].tot += totRev;
    dailyDict[cleanD].adx_imps += adxImp;
    dailyDict[cleanD].opmc_imps += opmcImp;
    dailyDict[cleanD].tot_imps += totImp;

    if (val.ad_unit) adUnitSet[val.ad_unit] = true;

    adUnitRows.push({
      date: formatDateLabel_(cleanD),
      raw_date: cleanD,
      ad_unit: val.ad_unit,
      adx_revenue: Math.round(adxRev * 100) / 100,
      opmc_revenue: Math.round(opmcRev * 100) / 100,
      total_revenue: Math.round(totRev * 100) / 100,
      adx_impressions: Math.round(adxImp),
      opmc_impressions: Math.round(opmcImp),
      total_impressions: Math.round(totImp)
    });
  });

  Object.keys(countryDict).forEach(function(k) {
    const val = countryDict[k];
    const adxRev = val.adx_rev;
    const opmcRev = val.opmc_rev;
    const totRev = adxRev + opmcRev;
    const adxImp = val.adx_imps;
    const opmcImp = val.opmc_imps;
    const totImp = adxImp + opmcImp;

    const cleanD = String(val.date).trim();
    if (val.country && val.country !== 'Unknown') {
      countrySet[val.country] = true;
    }

    countryRows.push({
      date: formatDateLabel_(cleanD),
      raw_date: cleanD,
      country: val.country,
      adx_revenue: Math.round(adxRev * 100) / 100,
      opmc_revenue: Math.round(opmcRev * 100) / 100,
      total_revenue: Math.round(totRev * 100) / 100,
      adx_impressions: Math.round(adxImp),
      opmc_impressions: Math.round(opmcImp),
      total_impressions: Math.round(totImp)
    });
  });

  adUnitRows.sort(function(a, b) {
    if (a.raw_date !== b.raw_date) return a.raw_date < b.raw_date ? -1 : 1;
    return b.total_revenue - a.total_revenue;
  });

  countryRows.sort(function(a, b) {
    if (a.raw_date !== b.raw_date) return a.raw_date < b.raw_date ? -1 : 1;
    return b.total_revenue - a.total_revenue;
  });

  const sortedDates = Object.keys(dailyDict).sort();
  const dailyLabels = sortedDates.map(function(d) { return formatDateLabel_(d); });
  const dailyAdx = sortedDates.map(function(d) { return Math.round(dailyDict[d].adx * 100) / 100; });
  const dailyOpmc = sortedDates.map(function(d) { return Math.round(dailyDict[d].opmc * 100) / 100; });
  const dailyTotal = sortedDates.map(function(d) { return Math.round(dailyDict[d].tot * 100) / 100; });
  const dailyAdxImps = sortedDates.map(function(d) { return Math.round(dailyDict[d].adx_imps); });
  const dailyOpmcImps = sortedDates.map(function(d) { return Math.round(dailyDict[d].opmc_imps); });
  const dailyTotalImps = sortedDates.map(function(d) { return Math.round(dailyDict[d].tot_imps); });

  const allAdUnits = Object.keys(adUnitSet).sort();
  const allCountries = Object.keys(countrySet).sort();

  const totCombined = totalAdxRev + totalOpmcRev;
  const adxSharePct = totCombined > 0 ? Math.round((totalAdxRev / totCombined) * 1000) / 10 : 0.0;
  const opmcSharePct = totCombined > 0 ? Math.round((totalOpmcRev / totCombined) * 1000) / 10 : 0.0;
  const totalImpsCombined = Math.round(totalAdxImps + totalOpmcImps);

  function fmtCurrency(n) {
    return currSymbol + Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  return {
    success: true,
    api_mode_label: modeLabel,
    currency_code: currencyCode,
    currency_symbol: currSymbol,
    summary: {
      currency_code: currencyCode,
      currency_symbol: currSymbol,
      adx_revenue: fmtCurrency(totalAdxRev),
      opmc_revenue: fmtCurrency(totalOpmcRev),
      grand_total: fmtCurrency(totCombined),
      raw_adx: Math.round(totalAdxRev * 100) / 100,
      raw_opmc: Math.round(totalOpmcRev * 100) / 100,
      raw_total: Math.round(totCombined * 100) / 100,
      adx_share_pct: adxSharePct,
      opmc_share_pct: opmcSharePct,
      total_impressions: totalImpsCombined.toLocaleString('en-US'),
      total_rows: adUnitRows.length
    },
    daily: {
      raw_dates: sortedDates,
      labels: dailyLabels,
      adx: dailyAdx,
      opmc: dailyOpmc,
      total: dailyTotal,
      adx_impressions: dailyAdxImps,
      opmc_impressions: dailyOpmcImps,
      total_impressions: dailyTotalImps
    },
    ad_units: allAdUnits,
    countries: allCountries,
    ad_unit_rows: adUnitRows,
    country_rows: countryRows
  };
}

/**
 * Writes the combined report data to the hidden data tabs, updates the Ad Unit Breakdown tab,
 * and updates the formulas in the Chart_Helper tab.
 */
function writeResultsToSheets_(data, params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  let unitSheet = ss.getSheetByName('Ad_Unit_Data') || ss.insertSheet('Ad_Unit_Data');
  unitSheet.clearContents();
  const unitHeader = [[
    'Date', 'Raw_Date', 'Ad_Unit', 'AdX_Revenue', 'OPMC_Revenue', 'Total_Revenue',
    'AdX_Impressions', 'OPMC_Impressions', 'Total_Impressions'
  ]];
  const unitValues = (data.ad_unit_rows || []).map(function(r) {
    return [
      r.date, r.raw_date, r.ad_unit, r.adx_revenue, r.opmc_revenue, r.total_revenue,
      r.adx_impressions || 0, r.opmc_impressions || 0, r.total_impressions
    ];
  });
  const allUnitRows = unitHeader.concat(unitValues);
  if (unitSheet.getMaxRows() < allUnitRows.length + 5) {
    unitSheet.insertRowsAfter(unitSheet.getMaxRows(), allUnitRows.length + 5 - unitSheet.getMaxRows());
  }
  unitSheet.getRange(1, 1, allUnitRows.length, 9).setValues(allUnitRows);

  let countrySheet = ss.getSheetByName('Country_Data') || ss.insertSheet('Country_Data');
  countrySheet.clearContents();
  const countryHeader = [[
    'Date', 'Raw_Date', 'Country', 'AdX_Revenue', 'OPMC_Revenue', 'Total_Revenue',
    'AdX_Impressions', 'OPMC_Impressions', 'Total_Impressions'
  ]];
  const countryValues = (data.country_rows || []).map(function(r) {
    return [
      r.date, r.raw_date, r.country, r.adx_revenue, r.opmc_revenue, r.total_revenue,
      r.adx_impressions || 0, r.opmc_impressions || 0, r.total_impressions
    ];
  });
  const allCountryRows = countryHeader.concat(countryValues);
  if (countrySheet.getMaxRows() < allCountryRows.length + 5) {
    countrySheet.insertRowsAfter(countrySheet.getMaxRows(), allCountryRows.length + 5 - countrySheet.getMaxRows());
  }
  countrySheet.getRange(1, 1, allCountryRows.length, 9).setValues(allCountryRows);

  const cSym = String(data.currency_symbol || '$').replace(/"/g, '');
  const currFmt = cSym === '$' ? '$#,##0.00' : '"' + cSym + '"#,##0.00';

  let dailySheet = ss.getSheetByName('Daily_Data') || ss.insertSheet('Daily_Data');
  dailySheet.clearContents();
  if (dailySheet.getMaxColumns() < 8) {
    dailySheet.insertColumnsAfter(dailySheet.getMaxColumns(), 8 - dailySheet.getMaxColumns());
  }
  const dailyHeader = [[
    'Date', 'Raw_Date', 'Ad Exchange Revenue', 'On-platform multicall Revenue', 'Total Revenue',
    'Ad Exchange Impressions', 'On-platform multicall Impressions', 'Total Impressions'
  ]];
  const dailyValues = (data.daily.labels || []).map(function(lbl, idx) {
    return [
      lbl,
      (data.daily.raw_dates && data.daily.raw_dates[idx]) || lbl,
      data.daily.adx[idx] || 0,
      data.daily.opmc[idx] || 0,
      data.daily.total[idx] || 0,
      (data.daily.adx_impressions && data.daily.adx_impressions[idx]) || 0,
      (data.daily.opmc_impressions && data.daily.opmc_impressions[idx]) || 0,
      (data.daily.total_impressions && data.daily.total_impressions[idx]) || 0
    ];
  });
  const allDailyRows = dailyHeader.concat(dailyValues);
  if (dailySheet.getMaxRows() < allDailyRows.length + 5) {
    dailySheet.insertRowsAfter(dailySheet.getMaxRows(), allDailyRows.length + 5 - dailySheet.getMaxRows());
  }
  dailySheet.getRange(1, 1, allDailyRows.length, 8).setValues(allDailyRows);
  try {
    if (dailyValues.length > 0) {
      dailySheet.getRange(2, 3, dailyValues.length, 3).setNumberFormat(currFmt);
      dailySheet.getRange(2, 6, dailyValues.length, 3).setNumberFormat('#,##0');
    }
  } catch (e) {}

  // Update the Ad Unit Breakdown tab.
  try {
    writeParentAdUnitBreakdownSheet_(ss, data);
  } catch (parentErr) {}

  const dash = getTargetDashboardSheet_(ss);
  if (dash) {
    const e8Label = String(dash.getRange('E8').getDisplayValue() || '').toUpperCase();
    const isStackedDatesLayout = e8Label.indexOf('END') !== -1;
    if (!isStackedDatesLayout) {
      dash.getRange('I8').setValue(data.currency_code || 'USD');
    }
    try {
      dash.getRange('B15').setNumberFormat(currFmt);
      dash.getRange('E15').setNumberFormat(currFmt);
      dash.getRange('H15').setNumberFormat(currFmt);
    } catch (e) {}
  }

  // Update daily revenue & eCPM rows, top 10 countries, filter dropdown lists, and currency code in Chart_Helper.
  const helperSheet = ss.getSheetByName('Chart_Helper');
  if (helperSheet) {
    try {
      if (helperSheet.getMaxColumns() < 20) {
        helperSheet.insertColumnsAfter(helperSheet.getMaxColumns(), 20 - helperSheet.getMaxColumns());
      }
      const numDaily = (data.daily && data.daily.labels) ? data.daily.labels.length : 0;
      const neededRows = Math.max(numDaily + 5, 405);
      if (helperSheet.getMaxRows() < neededRows) {
        helperSheet.insertRowsAfter(helperSheet.getMaxRows(), neededRows - helperSheet.getMaxRows());
      }

      helperSheet.getRange('M1:M2').setValues([['Currency_Code'], [data.currency_code || 'USD']]);
      helperSheet.getRange('N1:S1').setValues([[
        'Ad Exchange Impressions', 'On-platform multicall Impressions', 'Total Impressions',
        'Total eCPM', 'Ad Exchange eCPM', 'On-platform multicall eCPM'
      ]]);

      const clearSpan = Math.max(helperSheet.getLastRow(), 400);
      helperSheet.getRange(2, 1, clearSpan, 4).clearContent();
      helperSheet.getRange(2, 14, clearSpan, 6).clearContent();

      if (numDaily > 0) {
        const helperDailyFormulas = [];
        const helperEcpmFormulas = [];
        for (let dIdx = 0; dIdx < numDaily; dIdx++) {
          const rNum = dIdx + 2;
          helperDailyFormulas.push([
            '=Daily_Data!A' + rNum,
            '=C' + rNum + '+D' + rNum,
            '=IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$D$2:$D, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12), IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$D$2:$D, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12), Daily_Data!C' + rNum + '))',
            '=IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$E$2:$E, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12), IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$E$2:$E, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12), Daily_Data!D' + rNum + '))'
          ]);
          helperEcpmFormulas.push([
            '=IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$G$2:$G, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12), IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$G$2:$G, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12), Daily_Data!F' + rNum + '))',
            '=IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$H$2:$H, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12), IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$H$2:$H, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12), Daily_Data!G' + rNum + '))',
            '=N' + rNum + '+O' + rNum,
            '=IF(P' + rNum + '>0, (B' + rNum + '/P' + rNum + ')*1000, 0)',
            '=IF(N' + rNum + '>0, (C' + rNum + '/N' + rNum + ')*1000, 0)',
            '=IF(O' + rNum + '>0, (D' + rNum + '/O' + rNum + ')*1000, 0)'
          ]);
        }
        helperSheet.getRange(2, 1, numDaily, 4).setFormulas(helperDailyFormulas);
        helperSheet.getRange(2, 14, numDaily, 6).setFormulas(helperEcpmFormulas);
        try {
          const fmtSpan = Math.max(numDaily, 400);
          helperSheet.getRange(2, 2, fmtSpan, 3).setNumberFormat(currFmt);
          helperSheet.getRange(2, 14, fmtSpan, 3).setNumberFormat('#,##0');
          helperSheet.getRange(2, 17, fmtSpan, 3).setNumberFormat(currFmt);
        } catch (fmtErr) {}
      } else {
        const emptyLbl = (params && params.start_date && formatDateLabel_(params.start_date)) || '—';
        helperSheet.getRange('A2:D2').setValues([[emptyLbl, 0, 0, 0]]);
        helperSheet.getRange('N2:S2').setValues([[0, 0, 0, 0, 0, 0]]);
        try {
          helperSheet.getRange(2, 2, 400, 3).setNumberFormat(currFmt);
          helperSheet.getRange(2, 14, 400, 3).setNumberFormat('#,##0');
          helperSheet.getRange(2, 17, 400, 3).setNumberFormat(currFmt);
        } catch (fmtErr) {}
      }

      const cTotals = {};
      (data.country_rows || []).forEach(function(r) {
        const c = r.country || 'Unknown';
        if (c === 'Unknown') return;
        cTotals[c] = (cTotals[c] || 0) + (Number(r.total_revenue) || 0);
      });
      const top10Countries = Object.keys(cTotals)
        .sort(function(a, b) { return cTotals[b] - cTotals[a]; })
        .slice(0, 10);
      helperSheet.getRange('F1:I1').setValues([['Country', 'Ad Exchange Revenue', 'On-platform multicall Revenue', 'Total Revenue']]);
      const countryNames = [];
      const countryFormulas = [];
      for (let t = 0; t < 10; t++) {
        const rNum = t + 2;
        countryNames.push([top10Countries[t] || '']);
        countryFormulas.push([
          '=IF(F' + rNum + '="", 0, IF(OR(Overview!$F$12="ALL", Overview!$F$12=F' + rNum + '), SUMIF(Country_Data!$C$2:$C, F' + rNum + ', Country_Data!$D$2:$D), 0))',
          '=IF(F' + rNum + '="", 0, IF(OR(Overview!$F$12="ALL", Overview!$F$12=F' + rNum + '), SUMIF(Country_Data!$C$2:$C, F' + rNum + ', Country_Data!$E$2:$E), 0))',
          '=G' + rNum + '+H' + rNum
        ]);
      }
      helperSheet.getRange(2, 6, 10, 1).setValues(countryNames);
      helperSheet.getRange(2, 7, 10, 3).setFormulas(countryFormulas);
      try {
        helperSheet.getRange('G2:I11').setNumberFormat(currFmt);
      } catch (fmtErr) {}

      const unitOpts = [['Ad_Unit_Options'], ['ALL']].concat((data.ad_units || []).map(function(u) { return [u]; }));
      helperSheet.getRange(1, 11, Math.max(unitOpts.length, 150), 1).clearContent();
      helperSheet.getRange(1, 11, unitOpts.length, 1).setValues(unitOpts);

      const countryOpts = [['Country_Options'], ['ALL']].concat((data.countries || []).map(function(c) { return [c]; }));
      helperSheet.getRange(1, 12, Math.max(countryOpts.length, 250), 1).clearContent();
      helperSheet.getRange(1, 12, countryOpts.length, 1).setValues(countryOpts);

      if (dash) {
        syncOverviewCharts_(dash, helperSheet, numDaily, data.currency_code || 'USD');
      }
    } catch (helperErr) {}
  }

  // Hide the raw data tabs and the helper tab.
  [unitSheet, countrySheet, dailySheet, helperSheet].forEach(function(s) {
    if (s) {
      try {
        s.hideSheet();
      } catch (e) {}
    }
  });

  SpreadsheetApp.flush();
}

/**
 * Updates the exact daily row range and currency title on the Overview charts.
 */
function syncOverviewCharts_(overview, helperSheet, numDaily, currencyCode) {
  if (!overview || !helperSheet) return;
  try {
    const endRow = Math.max((Number(numDaily) || 0) + 1, 2);
    const curr = String(currencyCode || 'USD').trim().toUpperCase() || 'USD';
    const charts = overview.getCharts() || [];
    charts.forEach(function(ch) {
      try {
        const info = ch.getContainerInfo();
        const r = info ? info.getAnchorRow() : 0;
        if (r >= 18 && r <= 22) {
          const updated = ch
            .modify()
            .clearRanges()
            .addRange(helperSheet.getRange(1, 1, endRow, 4))
            .setOption('title', 'Daily Revenue Trends (' + curr + ')')
            .setOption('vAxis', { title: 'Revenue (' + curr + ')', textStyle: { fontName: 'Roboto', fontSize: 10 } })
            .build();
          overview.updateChart(updated);
        } else if (r >= 37 && r <= 41) {
          const b37 = String(overview.getRange('B37').getDisplayValue() || '').toLowerCase();
          if (b37.indexOf('ecpm') !== -1) {
            const updated = ch
              .modify()
              .clearRanges()
              .addRange(helperSheet.getRange(1, 1, endRow, 1))
              .addRange(helperSheet.getRange(1, 17, endRow, 3))
              .setOption('title', 'Daily eCPM Trends (' + curr + ')')
              .setOption('vAxis', { title: 'eCPM (' + curr + ')', textStyle: { fontName: 'Roboto', fontSize: 10 } })
              .build();
            overview.updateChart(updated);
          } else {
            const updated = ch
              .modify()
              .clearRanges()
              .addRange(helperSheet.getRange(1, 1, endRow, 1))
              .addRange(helperSheet.getRange(1, 3, endRow, 2))
              .build();
            overview.updateChart(updated);
          }
        } else if (r >= 56 && r <= 60) {
          const b56 = String(overview.getRange('B56').getDisplayValue() || '').toLowerCase();
          if (b56.indexOf('stacked') !== -1) {
            const updated = ch
              .modify()
              .clearRanges()
              .addRange(helperSheet.getRange(1, 1, endRow, 1))
              .addRange(helperSheet.getRange(1, 3, endRow, 2))
              .build();
            overview.updateChart(updated);
          }
        }
      } catch (innerErr) {}
    });
  } catch (err) {}
}

/**
 * Formats a date cell or raw date value as YYYY-MM-DD without a time value.
 */
function formatCsvDateOnly_(cellVal, rawVal) {
  if (rawVal && !(rawVal instanceof Date) && typeof rawVal.getFullYear !== 'function') {
    const rs = String(rawVal).trim();
    if (/^\d{8}$/.test(rs)) {
      return rs.substring(0, 4) + '-' + rs.substring(4, 6) + '-' + rs.substring(6, 8);
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(rs)) {
      return rs;
    }
  }
  if (cellVal && (cellVal instanceof Date || typeof cellVal.getFullYear === 'function')) {
    const y = cellVal.getFullYear();
    const m = ('0' + (cellVal.getMonth() + 1)).slice(-2);
    const d = ('0' + cellVal.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  const s = String(cellVal || '').trim();
  if (/^\d{8}$/.test(s)) {
    return s.substring(0, 4) + '-' + s.substring(4, 6) + '-' + s.substring(6, 8);
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return s.substring(0, 10);
  }
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime()) && s.indexOf(':') !== -1) {
    const y = parsed.getFullYear();
    const m = ('0' + (parsed.getMonth() + 1)).slice(-2);
    const d = ('0' + parsed.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return s.replace(/\s+\d{1,2}:\d{2}(:\d{2})?.*$/, '').trim();
}

/**
 * Reads the current report rows from the hidden data tabs for CSV export.
 */
function readExistingSheetDataForModal_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return null;
    const unitSheet = ss.getSheetByName('Ad_Unit_Data');
    const countrySheet = ss.getSheetByName('Country_Data');
    const helperSheet = ss.getSheetByName('Chart_Helper');
    const dash = getTargetDashboardSheet_(ss);
    if (!unitSheet || !countrySheet) return null;

    const uVals = unitSheet.getDataRange().getValues();
    const cVals = countrySheet.getDataRange().getValues();
    if (uVals.length <= 1) return null;

    let rawCurr = helperSheet ? String(helperSheet.getRange('M2').getDisplayValue() || '').trim() : '';
    if (!/^[A-Z]{3}$/.test(rawCurr) && dash) {
      rawCurr = String(dash.getRange('I8').getDisplayValue() || 'USD').trim();
    }
    const currencyCode = /^[A-Z]{3}$/.test(rawCurr) ? rawCurr : 'USD';
    const currSymbol = CURRENCY_SYMBOLS[currencyCode] || '$';

    const adUnitRows = [];
    for (let i = 1; i < uVals.length; i++) {
      const r = uVals[i];
      if (!r[0] && !r[2]) continue;
      const d = formatCsvDateOnly_(r[0], r[1]);
      const rawD = String(r[1] || d);
      const unit = String(r[2]);
      const adx = Number(r[3]) || 0;
      const opmc = Number(r[4]) || 0;
      const tot = Number(r[5]) || (adx + opmc);
      const adxImp = Number(r[6]) || 0;
      const opmcImp = Number(r[7]) || 0;
      const imp = Number(r[8] || (adxImp + opmcImp)) || 0;

      adUnitRows.push({
        date: d,
        raw_date: rawD,
        ad_unit: unit,
        adx_revenue: Math.round(adx * 100) / 100,
        opmc_revenue: Math.round(opmc * 100) / 100,
        total_revenue: Math.round(tot * 100) / 100,
        adx_impressions: Math.round(adxImp),
        opmc_impressions: Math.round(opmcImp),
        total_impressions: Math.round(imp)
      });
    }

    const countryRows = [];
    for (let j = 1; j < cVals.length; j++) {
      const cr = cVals[j];
      if (!cr[0] && !cr[2]) continue;
      const d = formatCsvDateOnly_(cr[0], cr[1]);
      const rawD = String(cr[1] || d);
      const country = String(cr[2]);
      const adx = Number(cr[3]) || 0;
      const opmc = Number(cr[4]) || 0;
      const tot = Number(cr[5]) || (adx + opmc);
      const imp = Number(cr[8] || cr[6]) || 0;

      countryRows.push({
        date: d,
        raw_date: rawD,
        country: country,
        adx_revenue: Math.round(adx * 100) / 100,
        opmc_revenue: Math.round(opmc * 100) / 100,
        total_revenue: Math.round(tot * 100) / 100,
        total_impressions: Math.round(imp)
      });
    }

    return {
      success: true,
      currency_code: currencyCode,
      currency_symbol: currSymbol,
      ad_unit_rows: adUnitRows,
      country_rows: countryRows
    };
  } catch (err) {
    return null;
  }
}
// === END OF CODE.GS ===