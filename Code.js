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
 * Google Ad Manager Unified Report (v1.3)
 *
 * This script gets Primary Programmatic (Ad Exchange, Open Bidding, SDK Bidding)
 * and On-Platform Multicall (OPMC) report data from the Google Ad Manager API
 * and writes the combined results to Google Sheets.
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

// Google-settled programmatic demand channels (excludes Mediation & Ad Server to avoid double-counting).
const SETTLED_DEMAND_CHANNELS = ['AD_EXCHANGE', 'OPEN_BIDDING', 'SDK_BIDDING'];
const DEMAND_CHANNEL_LABELS = {
  AD_EXCHANGE: 'Ad Exchange',
  OPEN_BIDDING: 'Open Bidding',
  SDK_BIDDING: 'SDK Bidding'
};
const DEMAND_CHANNEL_OPTIONS = ['Ad Exchange', 'Open Bidding', 'SDK Bidding'];

// Common noise, format, and platform words used to match parent and OPMC child ad units.
const NOISE_WORDS = {
  opmc: true,
  multicall: true,
  waterfall: true,
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
  timer: true,
  final: true
};

const FORMAT_WORDS = {
  native: true,
  banner: true,
  interstitial: true,
  inter: true,
  mrec: true,
  rewarded: true,
  video: true,
  appopen: true
};

const PLATFORM_WORDS = {
  a: true,
  i: true,
  w: true,
  m: true,
  d: true,
  android: true,
  ios: true,
  web: true,
  mobile: true,
  desktop: true,
  app: true
};

const ACTIVE_FILTER_FORMULA =
  '="Active Filter: "&IF(C12<>"ALL", "Ad Unit = "&C12, IF(F12<>"ALL", "Country = "&F12, IF(AND(C8<>"", UPPER(C8)<>"ALL", C8<>"<INSERT_HERE>"), "Query = """&C8&"""", "All Inventory")))&IF(AND(I12<>"ALL", I12<>""), "  •  Channel = "&I12, "  •  Channels: AdX + OB + SDKB")';

/**
 * Maps a GAM DEMAND_CHANNEL enum or label to its canonical channel name.
 */
function normalizeDemandChannel_(rawChannel) {
  const channelText = String(rawChannel || '').trim();
  if (!channelText) return 'Ad Exchange';
  const upperKey = channelText.toUpperCase().replace(/\s+/g, '_');
  if (DEMAND_CHANNEL_LABELS[upperKey]) return DEMAND_CHANNEL_LABELS[upperKey];
  const lowerText = channelText.toLowerCase();
  if (lowerText.indexOf('open') !== -1 && lowerText.indexOf('bidding') !== -1) return 'Open Bidding';
  if (lowerText.indexOf('sdk') !== -1 && lowerText.indexOf('bidding') !== -1) return 'SDK Bidding';
  if (lowerText.indexOf('exchange') !== -1 || lowerText.indexOf('adx') !== -1) return 'Ad Exchange';
  return channelText;
}

/**
 * Escapes Google Sheets SUMIF/SUMIFS criteria wildcards and double quotes.
 */
function gsCriteriaLiteral_(rawText) {
  return String(rawText || '')
    .replace(/~/g, '~~')
    .replace(/\*/g, '~*')
    .replace(/\?/g, '~?')
    .replace(/"/g, '""');
}

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
 * Inserts the 5 v1.3 embedded charts on Overview using explicit range and currency styling.
 */
function insertOverviewCharts_(overview, helperSheet, endRow, currencyCode) {
  const maxRow = Math.max((Number(endRow) || 0) + 1, 2);
  const curr = String(currencyCode || 'USD').trim().toUpperCase() || 'USD';
  (overview.getCharts() || []).forEach(function(ch) {
    try {
      overview.removeChart(ch);
    } catch (e) {}
  });

  const baseLegend = { position: 'top', textStyle: { fontName: 'Roboto', fontSize: 11 } };
  const axisStyle = { fontName: 'Roboto', fontSize: 10 };

  const chartDefs = [
    {
      type: 'LINE',
      anchorRow: 19,
      ranges: [helperSheet.getRange('A1:D' + maxRow)],
      title: 'Revenue Trends (' + curr + ')',
      hTitle: 'Date',
      vTitle: 'Revenue (' + curr + ')',
      colors: ['#0f9d58', '#1a73e8', '#f29900'],
      series: {
        0: { lineWidth: 3, color: '#0f9d58' },
        1: { lineWidth: 3, color: '#1a73e8' },
        2: { lineWidth: 3, lineDashStyle: [6, 3], color: '#f29900' }
      }
    },
    {
      type: 'LINE',
      anchorRow: 38,
      ranges: [helperSheet.getRange('A1:A' + maxRow), helperSheet.getRange('U1:W' + maxRow)],
      title: '',
      hTitle: 'Date',
      vTitle: 'OPMC Revenue (' + curr + ')',
      colors: ['#1a73e8', '#9334e6', '#f29900'],
      series: {
        0: { lineWidth: 3, color: '#1a73e8' },
        1: { lineWidth: 3, color: '#9334e6' },
        2: { lineWidth: 3, color: '#f29900' }
      }
    },
    {
      type: 'LINE',
      anchorRow: 57,
      ranges: [helperSheet.getRange('A1:A' + maxRow), helperSheet.getRange('Q1:S' + maxRow)],
      title: 'eCPM Trends (' + curr + ')',
      hTitle: 'Date',
      vTitle: 'eCPM (' + curr + ')',
      colors: ['#0f9d58', '#1a73e8', '#f29900'],
      series: {
        0: { lineWidth: 3, color: '#0f9d58' },
        1: { lineWidth: 3, color: '#1a73e8' },
        2: { lineWidth: 3, lineDashStyle: [6, 3], color: '#f29900' }
      }
    },
    {
      type: 'STACKED_COLUMN',
      anchorRow: 76,
      ranges: [helperSheet.getRange('A1:A' + maxRow), helperSheet.getRange('C1:D' + maxRow)],
      title: '',
      hTitle: 'Date',
      vTitle: 'Combined Revenue (' + curr + ')',
      colors: ['#1a73e8', '#f29900']
    },
    {
      type: 'STACKED_COLUMN',
      anchorRow: 95,
      ranges: [helperSheet.getRange('F1:H11')],
      title: '',
      hTitle: 'Country',
      vTitle: 'Revenue (' + curr + ')',
      colors: ['#1a73e8', '#f29900']
    }
  ];

  chartDefs.forEach(function(def) {
    let builder = overview.newChart();
    if (def.type === 'LINE') {
      builder = builder.asLineChart().setOption('curveType', 'function');
    } else {
      builder = builder.asColumnChart().setStacked();
    }
    def.ranges.forEach(function(rng) {
      builder = builder.addRange(rng);
    });
    builder = builder
      .setNumHeaders(1)
      .setPosition(def.anchorRow, 2, 8, 4)
      .setOption('width', 1190)
      .setOption('height', 355)
      .setOption('legend', baseLegend)
      .setOption('hAxis', { title: def.hTitle, textStyle: axisStyle })
      .setOption('vAxis', { title: def.vTitle, textStyle: axisStyle })
      .setOption('colors', def.colors)
      .setOption('useFirstColumnAsDomain', true);
    if (def.title) {
      builder = builder.setOption('title', def.title);
    }
    if (def.series) {
      builder = builder.setOption('series', def.series);
    }
    overview.insertChart(builder.build());
  });
}

/**
 * Returns the 43x10 cell value matrix for the v1.3 Instructions user guide.
 */
function buildInstructionsRows_() {
  const rows = [];
  for (let r = 0; r < 43; r++) {
    rows.push(['', '', '', '', '', '', '', '', '', '']);
  }
  rows[1][1] = 'Google Ad Manager Unified Report: User Guide';
  rows[2][1] = 'Google Sheets & Google Apps Script Reporting Tool';
  rows[3][1] =
    'This tool combines your Primary Google-settled programmatic revenue (Ad Exchange, Open Bidding, and SDK Bidding) with your On-Platform Multicall (OPMC) revenue into one report. Use this workbook to view daily trends, compare demand channels, and track revenue by ad unit.';

  rows[5][1] = '▎ 1. Workbook Tabs';
  rows[6][1] = '📊 Overview';
  rows[6][3] =
    'Run reports, filter data by ad unit, country, or demand channel, view total revenue cards, and check Revenue Trends, OPMC Revenue by Demand Channel, eCPM Trends, Stacked Incremental, and Country charts.';
  rows[7][1] = '🧩 Ad Unit Breakdown';
  rows[7][3] =
    'View revenue for each Parent Ad Unit (A) combined with its OPMC Child Ad Units (A1 to An), filtered live by Demand Channel.';

  rows[9][1] = '▎ 2. One-Time Setup: Google Cloud API and Apps Script';
  rows[10][1] = 'Step';
  rows[10][2] = 'Where to Go';
  rows[10][4] = 'What to Do';
  const setupSteps = [
    ['Step 1', 'Google Cloud Console -> APIs & Services', 'Open your Google Cloud project and enable the Google Ad Manager API (admanager.googleapis.com).'],
    ['Step 2', 'Google Cloud Console -> IAM & Admin -> IAM', 'Ensure your Google account has the Service Usage Consumer role (roles/serviceusage.serviceUsageConsumer) on that GCP project.'],
    ['Step 3', 'Google Cloud Console -> Dashboard', 'Copy your GCP Project ID (or Project Number) from the Project info card on the Cloud Console Dashboard.'],
    ['Step 4', 'Google Sheet -> Extensions -> Apps Script', 'In a blank Google Sheet, click Extensions -> Apps Script, enable Show "appsscript.json" manifest file in editor under Project Settings, paste Code.js and appsscript.json, and save.'],
    ['Step 5', 'Apps Script -> Editor (< >) -> Toolbar', 'Select authorizeAndVerifyGamScope in the top toolbar, click Run, and approve the permission prompt to build all workbook tabs.'],
    ['Step 6', 'Google Sheet -> Refresh Page & Cell I7', 'Refresh the Google Sheet browser tab once (Cmd+R or F5) to activate the buttons, enter your GCP Project ID in cell I7, and press Enter.']
  ];
  setupSteps.forEach(function(st, idx) {
    rows[11 + idx][1] = st[0];
    rows[11 + idx][2] = st[1];
    rows[11 + idx][4] = st[2];
  });

  rows[18][1] = '▎ 3. How to Generate a Report (Step-by-Step)';
  rows[19][1] = 'Step';
  rows[19][2] = 'Action';
  rows[19][4] = 'Instructions';
  const runSteps = [
    ['Step 1', 'Open the Overview tab', 'Click the Overview tab at the bottom of the sheet.'],
    ['Step 2', 'Enter Network Code, Optional Filter, and Project ID', 'Type your Google Ad Manager Network Code in cell C7 (optional: enter an Ad Unit substring filter in cell C8) and your GCP Project ID in cell I7, then press Enter.'],
    ['Step 3', 'Select your date range', 'Double-click cell F7 to select the Start Date and double-click cell F8 to select the End Date (YYYY-MM-DD).'],
    ['Step 4', 'Run the report', 'Click the Generate Report button in row 10. Wait for the Report complete message to appear.']
  ];
  runSteps.forEach(function(st, idx) {
    rows[20 + idx][1] = st[0];
    rows[20 + idx][2] = st[1];
    rows[20 + idx][4] = st[2];
  });

  rows[25][1] = '▎ 4. How to Use the Buttons and Filters in the "Overview" tab';
  rows[26][1] = 'Button or Filter';
  rows[26][3] = 'How to Use It';
  const controls = [
    ['⚡ Generate Report', 'Fetch new data from Google Ad Manager and update all charts and tables.'],
    ['🔄 Reset Filters', 'Reset the Ad Unit, Country, and Demand Channel filters back to ALL.'],
    ['📥 Download CSV', 'Download the full report as a CSV file.'],
    ['🧹 Clear Data', 'Wipe all fetched report data from all sheets while preserving your settings and formulas.'],
    ['📱 Filter by Ad Unit', 'Select an ad unit from cell C12 to show revenue for that unit. Selecting a Parent Ad Unit includes all associated OPMC child units.'],
    ['🌍 Filter by Country', 'Select a country from the dropdown list in cell F12 to show revenue for that country only.'],
    ['📡 Filter by Demand Channel', 'Select ALL, Ad Exchange, Open Bidding, or SDK Bidding in cell I12 to filter both Overview and Ad Unit Breakdown live.']
  ];
  controls.forEach(function(ctrl, idx) {
    rows[27 + idx][1] = ctrl[0];
    rows[27 + idx][3] = ctrl[1];
  });

  rows[35][1] = '▎ 5. Key Revenue Terms';
  rows[36][1] = 'Term';
  rows[36][3] = 'Meaning';
  const terms = [
    ['Primary Programmatic Revenue', 'Google-settled primary revenue from Ad Exchange, Open Bidding, and SDK Bidding (excludes Mediation and Ad Server to prevent double-counting).'],
    ['OPMC Revenue', 'Additional Google-settled revenue (Ad Exchange, Open Bidding, and SDK Bidding) earned on On-Platform Multicall (OPMC) child ad units.'],
    ['Total Unified Revenue', 'The sum of Primary Programmatic Revenue plus OPMC Revenue (100% mutually exclusive with zero double-counting).'],
    ['Parent Ad Unit Rollup', 'On the Ad Unit Breakdown tab, Table 1 groups each Parent Ad Unit (A) with its OPMC Child Ad Units (A1 to An). Edit Column D in Table 2 to reassign any child unit.']
  ];
  terms.forEach(function(term, idx) {
    rows[37 + idx][1] = term[0];
    rows[37 + idx][3] = term[1];
  });

  return rows;
}

/**
 * Formats and writes the 43x10 Instructions tab cleanly.
 */
function styleInstructionsSheet_(instSheet) {
  if (!instSheet) return;
  instSheet.clearContents();
  instSheet.getRange(1, 1, instSheet.getMaxRows(), instSheet.getMaxColumns()).breakApart();
  ensureSheetGridSize_(instSheet, 43, 10);
  instSheet.setTabColor('#1a73e8');
  instSheet.setHiddenGridlines(true);

  const instColWidths = [26, 95, 155, 155, 165, 165, 165, 165, 165, 26];
  for (let c = 0; c < instColWidths.length; c++) {
    instSheet.setColumnWidth(c + 1, instColWidths[c]);
  }

  instSheet.getRange('A1:J43')
    .setBackground('#f1f3f4')
    .setFontColor('#202124')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontWeight('normal')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);

  const instRows = buildInstructionsRows_();
  instSheet.getRange(1, 1, 43, 10).setValues(instRows);

  const rowHeights = [
    [1, 1, 16], [2, 1, 38], [3, 1, 24], [4, 1, 44], [5, 1, 16],
    [6, 1, 32], [7, 2, 34], [9, 1, 16], [10, 1, 32], [11, 1, 28],
    [12, 6, 34], [18, 1, 16], [19, 1, 32], [20, 1, 28], [21, 4, 30],
    [25, 1, 16], [26, 1, 32], [27, 1, 28], [28, 7, 30], [35, 1, 16],
    [36, 1, 32], [37, 1, 28], [38, 3, 34], [41, 1, 44], [42, 2, 18]
  ];
  rowHeights.forEach(function(rh) {
    if (rh[1] === 1) {
      instSheet.setRowHeight(rh[0], rh[2]);
    } else {
      instSheet.setRowHeights(rh[0], rh[1], rh[2]);
    }
  });

  ['B2:I2', 'B3:I3', 'B4:I4', 'B6:I6', 'B10:I10', 'B19:I19', 'B26:I26', 'B36:I36'].forEach(function(a1) {
    instSheet.getRange(a1).merge();
  });
  [7, 8, 27, 28, 29, 30, 31, 32, 33, 34, 37, 38, 39, 40, 41].forEach(function(r) {
    instSheet.getRange(r, 2, 1, 2).merge();
    instSheet.getRange(r, 4, 1, 6).merge();
  });
  [11, 12, 13, 14, 15, 16, 17, 20, 21, 22, 23, 24].forEach(function(r) {
    instSheet.getRange(r, 3, 1, 2).merge();
    instSheet.getRange(r, 5, 1, 5).merge();
  });

  instSheet.getRange('B2:I4').setBackground('#ffffff');
  instSheet.getRange('B2').setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8');
  instSheet.getRange('B3').setFontSize(10).setFontWeight('bold').setFontColor('#5f6368');
  instSheet.getRange('B4:I4').setBackground('#f8f9fa').setFontSize(10).setFontColor('#3c4043').setWrap(true);
  styleCardBox_(instSheet.getRange('B2:I4'), '#dadce0', '#1a73e8');

  [
    { hdr: 'B6:I6', box: 'B6:I8', tblHdr: null },
    { hdr: 'B10:I10', box: 'B10:I17', tblHdr: 'B11:I11' },
    { hdr: 'B19:I19', box: 'B19:I24', tblHdr: 'B20:I20' },
    { hdr: 'B26:I26', box: 'B26:I34', tblHdr: 'B27:I27' },
    { hdr: 'B36:I36', box: 'B36:I41', tblHdr: 'B37:I37' }
  ].forEach(function(sec) {
    instSheet.getRange(sec.box)
      .setBackground('#ffffff')
      .setWrap(true)
      .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
    instSheet.getRange(sec.hdr).setBackground('#ffffff').setFontColor('#1a73e8').setFontSize(12).setFontWeight('bold');
    if (sec.tblHdr) {
      instSheet.getRange(sec.tblHdr).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    }
  });
  instSheet.getRange('B7:C8').setFontWeight('bold');
  instSheet.getRange('B12:B17').setFontWeight('bold').setHorizontalAlignment('center');
  instSheet.getRange('B21:B24').setFontWeight('bold').setHorizontalAlignment('center');
  instSheet.getRange('B28:C34').setFontWeight('bold');
  instSheet.getRange('B38:C41').setFontWeight('bold');
}

/**
 * Applies the v1.3 Overview formulas, styles, validations, and 5-chart card containers.
 */
function styleOverviewV13Body_(overview, helperSheet, currFmt) {
  const fmt = currFmt || '$#,##0.00';
  const chCritLocal = 'IF(OR(I12="ALL",I12=""),"*",I12)';

  overview.getRange('B7:I8').breakApart();
  overview.getRange('H7:H8').merge();
  overview.getRange('I7:I8').merge();
  overview.getRange('C7:C8').setNumberFormat('@');
  overview.getRange('I7:I8').setNumberFormat('@');
  overview.getRange('F7:F8').setNumberFormat('yyyy-mm-dd');
  overview.getRange('B8').setValue('Ad Unit Filter (Optional)');
  overview.getRange('H6').setFormula(
    '=IF(C12<>"ALL", "Filtered records (Ad Unit: "&C12&"): "&TEXT(COUNTIFS(Ad_Unit_Data!C2:C, C12, Ad_Unit_Data!K2:K, ' +
      chCritLocal +
      ') + COUNTIFS(Ad_Unit_Data!J2:J, C12, Ad_Unit_Data!K2:K, ' +
      chCritLocal +
      '), "#,##0"), IF(F12<>"ALL", "Filtered records (Country: "&F12&"): "&TEXT(COUNTIFS(Country_Data!C2:C, F12, Country_Data!J2:J, ' +
      chCritLocal +
      '), "#,##0"), IF(AND(C8<>"", UPPER(C8)<>"ALL", C8<>"<INSERT_HERE>"), "Filtered records (Query: "&C8&"): "&TEXT(COUNTIF(Ad_Unit_Data!K2:K, ' +
      chCritLocal +
      '), "#,##0"), "Dataset records ("&IF(OR(I12="ALL",I12=""),"All Channels",I12)&"): "&TEXT(COUNTIF(Ad_Unit_Data!K2:K, ' +
      chCritLocal +
      '), "#,##0"))))'
  );
  ['B7:B8', 'E7:E8', 'H7:H8'].forEach(function(lblRange) {
    overview.getRange(lblRange).setFontColor('#5f6368').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right');
  });
  ['C7', 'C8', 'F7', 'F8', 'I7:I8'].forEach(function(inpRange) {
    const r = overview.getRange(inpRange);
    r.setBackground('#f8f9fa').setFontColor('#202124').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('left');
    r.setBorder(true, true, true, true, false, false, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
  });

  overview.getRange('H11:I12').breakApart();
  overview.getRange('H11:I11').merge();
  overview.getRange('H11').setValue('📡 Filter by Demand Channel (Select below)');
  overview.getRange('H12').setValue('Demand Channel:');
  overview.getRange('I12').setValue('ALL');
  overview.getRange('H11:I11').setBackground('#ffffff').setFontColor('#5f6368').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('left');
  overview.getRange('H12').setBackground('#ffffff').setFontColor('#5f6368').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right').setNumberFormat('@');
  overview.getRange('I12')
    .setBackground('#e6f4ea')
    .setFontColor('#137333')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setNumberFormat('@')
    .setBorder(true, true, true, true, false, false, '#0f9d58', SpreadsheetApp.BorderStyle.SOLID);
  styleCardBox_(overview.getRange('B11:I12'), '#dadce0');

  overview.getRange('B14').setFormula('=IF(OR(I12="ALL",I12=""), "PRIMARY REVENUE (ADX + OB + SDKB)", UPPER(I12)&" PRIMARY REVENUE")');
  overview.getRange('E14').setFormula('=IF(OR(I12="ALL",I12=""), "OPMC REVENUE (ADX + OB + SDKB)", "OPMC "&UPPER(I12)&" REVENUE")');
  overview.getRange('H14').setValue('TOTAL UNIFIED REVENUE');
  overview.getRange('B15').setFormula(
    '=IF(C12<>"ALL", SUMIFS(Ad_Unit_Data!D2:D, Ad_Unit_Data!C2:C, C12, Ad_Unit_Data!K2:K, ' +
      chCritLocal +
      ') + SUMIFS(Ad_Unit_Data!D2:D, Ad_Unit_Data!J2:J, C12, Ad_Unit_Data!K2:K, ' +
      chCritLocal +
      '), IF(F12<>"ALL", SUMIFS(Country_Data!D2:D, Country_Data!C2:C, F12, Country_Data!J2:J, ' +
      chCritLocal +
      '), SUMIFS(Ad_Unit_Data!D2:D, Ad_Unit_Data!K2:K, ' +
      chCritLocal +
      ')))'
  );
  overview.getRange('E15').setFormula(
    '=IF(C12<>"ALL", SUMIFS(Ad_Unit_Data!E2:E, Ad_Unit_Data!C2:C, C12, Ad_Unit_Data!K2:K, ' +
      chCritLocal +
      ') + SUMIFS(Ad_Unit_Data!E2:E, Ad_Unit_Data!J2:J, C12, Ad_Unit_Data!K2:K, ' +
      chCritLocal +
      '), IF(F12<>"ALL", SUMIFS(Country_Data!E2:E, Country_Data!C2:C, F12, Country_Data!J2:J, ' +
      chCritLocal +
      '), SUMIFS(Ad_Unit_Data!E2:E, Ad_Unit_Data!K2:K, ' +
      chCritLocal +
      ')))'
  );
  overview.getRange('H15').setFormula('=B15+E15');
  overview.getRange('B15:C15').setNumberFormat(fmt);
  overview.getRange('E15:F15').setNumberFormat(fmt);
  overview.getRange('H15:I15').setNumberFormat(fmt);
  overview.getRange('B16').setFormula('="Share: "&TEXT(IF(H15>0, B15/H15, 0), "0.0%")&"  •  "&TEXT(SUM(Chart_Helper!N2:N), "#,##0")&" imps"');
  overview.getRange('E16').setFormula('="Share: "&TEXT(IF(H15>0, E15/H15, 0), "0.0%")&"  •  "&TEXT(SUM(Chart_Helper!O2:O), "#,##0")&" imps"');
  overview.getRange('H16').setFormula('="Uplift: +"&TEXT(IF(B15>0, E15/B15, 0), "0.0%")&"  •  "&TEXT(SUM(Chart_Helper!P2:P), "#,##0")&" imps"');

  ensureSheetGridSize_(overview, 114, 10);
  const canvasRange = overview.getRange('A17:J114');
  canvasRange.breakApart();
  canvasRange.clearContent();
  canvasRange
    .setBackground('#f1f3f4')
    .setFontColor('#202124')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontWeight('normal')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(false, false, false, false, false, false);

  const chartContainers = [
    { row: 18, box: 'B18:I35', titleRange: 'B18:G18', subRange: 'H18:I18', title: '▎ Revenue Trends', color: '#1a73e8' },
    { row: 37, box: 'B37:I54', titleRange: 'B37:G37', subRange: 'H37:I37', title: '▎ OPMC Revenue by Demand Channel', color: '#9334e6' },
    { row: 56, box: 'B56:I73', titleRange: 'B56:G56', subRange: 'H56:I56', title: '▎ eCPM Trends', color: '#1a73e8' },
    { row: 75, box: 'B75:I92', titleRange: 'B75:G75', subRange: 'H75:I75', title: '▎ Stacked Incremental Revenue', color: '#0f9d58' },
    { row: 94, box: 'B94:I111', titleRange: 'B94:G94', subRange: 'H94:I94', title: '▎ Country Revenue Breakdown (Top 10 Countries)', color: '#f29900' }
  ];
  chartContainers.forEach(function(ch) {
    overview.setRowHeight(ch.row - 1, 18);
    overview.setRowHeight(ch.row, 34);
    overview.setRowHeights(ch.row + 1, 17, 22);
    overview.getRange(ch.titleRange).merge();
    overview.getRange(ch.subRange).merge();
    overview.getRange(ch.box).setBackground('#ffffff');
    styleCardBox_(overview.getRange(ch.box), '#dadce0');
    overview.getRange(ch.row, 2).setValue(ch.title);
    overview.getRange(ch.row, 8).setFormula(ACTIVE_FILTER_FORMULA);
    overview.getRange(ch.titleRange).setFontColor(ch.color).setFontSize(12).setFontWeight('bold');
    overview.getRange(ch.subRange).setFontColor('#5f6368').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right');
  });
  overview.setRowHeights(112, 3, 18);

  if (helperSheet) {
    overview.getRange('I12').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(helperSheet.getRange('T2:T5'), true).setAllowInvalid(true).build()
    );
  }
}

/**
 * Upgrades an existing pre-v1.3 workbook in-place to the 5-chart v1.3 schema and layout.
 */
function upgradeToV13LayoutIfNeeded_(ss, currencySymbol) {
  if (!ss) return;
  const overview = ss.getSheetByName('Overview');
  const helperSheet = ss.getSheetByName('Chart_Helper');
  const unitSheet = ss.getSheetByName('Ad_Unit_Data');
  const countrySheet = ss.getSheetByName('Country_Data');
  const dailySheet = ss.getSheetByName('Daily_Data');
  const instSheet = ss.getSheetByName('Instructions');
  if (!overview || !helperSheet || !unitSheet || !countrySheet || !dailySheet) return;

  const h12 = String(overview.getRange('H12').getDisplayValue() || '').trim();
  const b37 = String(overview.getRange('B37').getDisplayValue() || '').toLowerCase();
  const alreadyV13 =
    h12 === 'Demand Channel:' &&
    b37.indexOf('demand channel') !== -1 &&
    overview.getMaxRows() >= 114 &&
    helperSheet.getMaxColumns() >= 26 &&
    unitSheet.getMaxColumns() >= 11 &&
    countrySheet.getMaxColumns() >= 10 &&
    dailySheet.getMaxColumns() >= 14;
  if (alreadyV13) return;

  const cSym = String(currencySymbol || '$').replace(/"/g, '');
  const currFmt = cSym === '$' ? '$#,##0.00' : '"' + cSym + '"#,##0.00';

  if (unitSheet.getMaxColumns() < 11) {
    unitSheet.insertColumnsAfter(unitSheet.getMaxColumns(), 11 - unitSheet.getMaxColumns());
  }
  unitSheet.getRange('A1:K1').setValues([[
    'Date', 'Raw_Date', 'Ad_Unit', 'Primary_Revenue', 'OPMC_Revenue', 'Total_Revenue',
    'Primary_Impressions', 'OPMC_Impressions', 'Total_Impressions', 'Parent_Ad_Unit', 'Demand_Channel'
  ]]).setBackground('#1a73e8').setFontColor('#ffffff').setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');

  if (countrySheet.getMaxColumns() < 10) {
    countrySheet.insertColumnsAfter(countrySheet.getMaxColumns(), 10 - countrySheet.getMaxColumns());
  }
  countrySheet.getRange('A1:J1').setValues([[
    'Date', 'Raw_Date', 'Country', 'Primary_Revenue', 'OPMC_Revenue', 'Total_Revenue',
    'Primary_Impressions', 'OPMC_Impressions', 'Total_Impressions', 'Demand_Channel'
  ]]).setBackground('#1a73e8').setFontColor('#ffffff').setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');

  if (dailySheet.getMaxColumns() < 14) {
    dailySheet.insertColumnsAfter(dailySheet.getMaxColumns(), 14 - dailySheet.getMaxColumns());
  }
  dailySheet.getRange('A1:N1').setValues([[
    'Date', 'Raw_Date', 'Primary Programmatic Revenue', 'OPMC Revenue', 'Total Revenue',
    'Primary Impressions', 'OPMC Impressions', 'Total Impressions',
    'OPMC Ad Exchange Revenue', 'OPMC Open Bidding Revenue', 'OPMC SDK Bidding Revenue',
    'Primary Ad Exchange Revenue', 'Primary Open Bidding Revenue', 'Primary SDK Bidding Revenue'
  ]]).setBackground('#0f9d58').setFontColor('#ffffff').setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');

  if (helperSheet.getMaxColumns() < 26) {
    helperSheet.insertColumnsAfter(helperSheet.getMaxColumns(), 26 - helperSheet.getMaxColumns());
  }
  if (helperSheet.getMaxRows() < 405) {
    helperSheet.insertRowsAfter(helperSheet.getMaxRows(), 405 - helperSheet.getMaxRows());
  }
  helperSheet.getRange('T1:T5').setValues([['Demand_Channel_Options'], ['ALL'], ['Ad Exchange'], ['Open Bidding'], ['SDK Bidding']]);
  helperSheet.getRange('U1:W1').setValues([['OPMC — Ad Exchange', 'OPMC — Open Bidding', 'OPMC — SDK Bidding']]);
  if (String(helperSheet.getRange('U2').getDisplayValue() || '') === '') {
    helperSheet.getRange('U2:W2').setValues([[0, 0, 0]]);
  }

  if (instSheet) {
    styleInstructionsSheet_(instSheet);
  }

  styleOverviewV13Body_(overview, helperSheet, currFmt);
  const currCode = String(helperSheet.getRange('M2').getDisplayValue() || 'USD').trim() || 'USD';
  insertOverviewCharts_(overview, helperSheet, 400, currCode);
}

/**
 * Automatically creates and formats all 7 workbook tabs (Instructions, Overview,
 * Ad Unit Breakdown, Ad_Unit_Data, Country_Data, Daily_Data, and Chart_Helper),
 * formulas, dropdowns, and embedded charts when run on a blank Google Sheet.
 * Upgrades pre-v1.3 sheets to v1.3 when needed; otherwise no-op when already v1.3.
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
      upgradeToV13LayoutIfNeeded_(ss, '$');
    } catch (e) {}
    return;
  }

  // 1. Create or configure hidden data and helper sheets first so formula & validation ranges resolve cleanly.
  const unitSheet = ss.getSheetByName('Ad_Unit_Data') || ss.insertSheet('Ad_Unit_Data');
  ensureSheetGridSize_(unitSheet, 200, 11);
  unitSheet.setTabColor('#1a73e8');
  unitSheet.setFrozenRows(1);
  unitSheet.getRange('A1:K1').setValues([[
    'Date', 'Raw_Date', 'Ad_Unit', 'Primary_Revenue', 'OPMC_Revenue', 'Total_Revenue',
    'Primary_Impressions', 'OPMC_Impressions', 'Total_Impressions', 'Parent_Ad_Unit', 'Demand_Channel'
  ]]);
  unitSheet.getRange('A1:K1')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  unitSheet.getRange('D2:F200').setNumberFormat('$#,##0.00');
  unitSheet.getRange('G2:I200').setNumberFormat('#,##0');

  const countrySheet = ss.getSheetByName('Country_Data') || ss.insertSheet('Country_Data');
  ensureSheetGridSize_(countrySheet, 200, 10);
  countrySheet.setTabColor('#f29900');
  countrySheet.setFrozenRows(1);
  countrySheet.getRange('A1:J1').setValues([[
    'Date', 'Raw_Date', 'Country', 'Primary_Revenue', 'OPMC_Revenue', 'Total_Revenue',
    'Primary_Impressions', 'OPMC_Impressions', 'Total_Impressions', 'Demand_Channel'
  ]]);
  countrySheet.getRange('A1:J1')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  countrySheet.getRange('D2:F200').setNumberFormat('$#,##0.00');
  countrySheet.getRange('G2:I200').setNumberFormat('#,##0');

  const dailySheet = ss.getSheetByName('Daily_Data') || ss.insertSheet('Daily_Data');
  ensureSheetGridSize_(dailySheet, 405, 14);
  dailySheet.setTabColor('#0f9d58');
  dailySheet.setFrozenRows(1);
  dailySheet.getRange('A1:N1').setValues([[
    'Date', 'Raw_Date', 'Primary Programmatic Revenue', 'OPMC Revenue', 'Total Revenue',
    'Primary Impressions', 'OPMC Impressions', 'Total Impressions',
    'OPMC Ad Exchange Revenue', 'OPMC Open Bidding Revenue', 'OPMC SDK Bidding Revenue',
    'Primary Ad Exchange Revenue', 'Primary Open Bidding Revenue', 'Primary SDK Bidding Revenue'
  ]]);
  dailySheet.getRange('A1:N1')
    .setBackground('#0f9d58')
    .setFontColor('#ffffff')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  dailySheet.getRange('C2:E405').setNumberFormat('$#,##0.00');
  dailySheet.getRange('F2:H405').setNumberFormat('#,##0');
  dailySheet.getRange('I2:N405').setNumberFormat('$#,##0.00');

  const helperSheet = ss.getSheetByName('Chart_Helper') || ss.insertSheet('Chart_Helper');
  ensureSheetGridSize_(helperSheet, 405, 26);
  helperSheet.setTabColor('#9aa0a6');
  helperSheet.setFrozenRows(1);
  helperSheet.getRange('A1:W2').setValues([
    [
      'Date', 'Total Revenue', 'Primary Programmatic Revenue', 'OPMC Revenue', '',
      'Country', 'Primary Programmatic Revenue', 'OPMC Revenue', 'Total Revenue', '',
      'Ad_Unit_Options', 'Country_Options', 'Currency_Code',
      'Primary Impressions', 'OPMC Impressions', 'Total Impressions',
      'Total eCPM', 'Primary eCPM', 'OPMC eCPM',
      'Demand_Channel_Options', 'OPMC — Ad Exchange', 'OPMC — Open Bidding', 'OPMC — SDK Bidding'
    ],
    ['—', 0, 0, 0, '', '—', 0, 0, 0, '', 'ALL', 'ALL', 'USD', 0, 0, 0, 0, 0, 0, 'ALL', 0, 0, 0]
  ]);
  helperSheet.getRange('T3:T5').setValues([['Ad Exchange'], ['Open Bidding'], ['SDK Bidding']]);
  helperSheet.getRange('B2:D405').setNumberFormat('$#,##0.00');
  helperSheet.getRange('G2:I15').setNumberFormat('$#,##0.00');
  helperSheet.getRange('N2:P405').setNumberFormat('#,##0');
  helperSheet.getRange('Q2:S405').setNumberFormat('$#,##0.00');
  helperSheet.getRange('U2:W405').setNumberFormat('$#,##0.00');

  // 2. Build the Instructions tab.
  const instSheet = ss.getSheetByName('Instructions') || ss.insertSheet('Instructions');
  styleInstructionsSheet_(instSheet);

  // 3. Build the Overview tab (114 rows, 5 chart sections).
  const overview = ss.getSheetByName('Overview') || ss.insertSheet('Overview');
  ensureSheetGridSize_(overview, 114, 10);
  overview.setTabColor('#1a73e8');
  overview.setHiddenGridlines(true);

  const ovColWidths = [26, 185, 185, 48, 185, 185, 48, 185, 185, 26];
  for (let c = 0; c < ovColWidths.length; c++) {
    overview.setColumnWidth(c + 1, ovColWidths[c]);
  }

  overview.getRange('A1:J114')
    .setBackground('#f1f3f4')
    .setFontColor('#202124')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  const todayIso = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'UTC', 'yyyy-MM-dd');
  const ovValues = [];
  for (let r = 0; r < 114; r++) {
    ovValues.push(['', '', '', '', '', '', '', '', '', '']);
  }

  ovValues[1][1] = 'Google Ad Manager Unified Report';
  ovValues[3][1] =
    'Copyright Google LLC. Supported by Google LLC and/or its affiliate(s). This solution, including any related sample code or data, is made available on an “as is,” “as available,” and “with all faults” basis, solely for illustrative purposes, and without warranty or representation of any kind. This solution is experimental, unsupported and provided solely for your convenience. Your use of it is subject to your agreements with Google, as applicable, and may constitute a beta feature as defined under those agreements. To the extent that you make any data available to Google in connection with your use of the solution, you represent and warrant that you have all necessary and appropriate rights, consents and permissions to permit Google to use and process that data. By using any portion of this solution, you acknowledge, assume and accept all risks, known and unknown, associated with its usage and any processing of data by Google, including with respect to your deployment of any portion of this solution in your systems, or usage in connection with your business, if at all. In connection with this solution, you will not provide to Google any personally identifiable information, personal information or personal data.';

  ovValues[5][1] = '▎ Report Parameters';
  ovValues[6][1] = 'Network Code';
  ovValues[6][2] = '<INSERT_HERE>';
  ovValues[6][4] = 'Start Date (YYYY-MM-DD)';
  ovValues[6][5] = todayIso;
  ovValues[6][7] = 'GCP Project ID (Quota)';
  ovValues[6][8] = '<INSERT_HERE>';

  ovValues[7][1] = 'Ad Unit Filter (Optional)';
  ovValues[7][2] = '';
  ovValues[7][4] = 'End Date (YYYY-MM-DD)';
  ovValues[7][5] = todayIso;

  ovValues[9][1] = '⚡ Generate Report';
  ovValues[9][4] = '🔄 Reset Filters';
  ovValues[9][5] = '🧹 Clear Data';
  ovValues[9][7] = '📥 Download CSV';

  ovValues[10][1] = '📱 Filter by Ad Unit (Select below)';
  ovValues[10][4] = '🌍 Filter by Country (Select below)';
  ovValues[10][7] = '📡 Filter by Demand Channel (Select below)';

  ovValues[11][1] = 'Ad Unit Filter:';
  ovValues[11][2] = 'ALL';
  ovValues[11][4] = 'Country Filter:';
  ovValues[11][5] = 'ALL';
  ovValues[11][7] = 'Demand Channel:';
  ovValues[11][8] = 'ALL';

  ovValues[13][3] = '="+"';
  ovValues[13][6] = '="="';
  ovValues[13][7] = 'TOTAL UNIFIED REVENUE';

  overview.getRange('C7:C8').setNumberFormat('@');
  overview.getRange('I7:I8').setNumberFormat('@');
  overview.getRange('I12').setNumberFormat('@');
  overview.getRange(1, 1, ovValues.length, 10).setValues(ovValues);
  overview.getRange('F7:F8').setNumberFormat('yyyy-mm-dd');

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

  [
    'B2:I3', 'B4:I4',
    'B6:G6', 'H6:I6',
    'B10:C10', 'H10:I10',
    'B11:C11', 'E11:F11', 'H11:I11',
    'B14:C14', 'B15:C15', 'B16:C16', 'D14:D16',
    'E14:F14', 'E15:F15', 'E16:F16', 'G14:G16',
    'H14:I14', 'H15:I15', 'H16:I16'
  ].forEach(function(a1) {
    overview.getRange(a1).merge();
  });

  overview.getRange('B2:I3').setBackground('#ffffff').setFontColor('#202124').setFontSize(16).setFontWeight('bold');
  overview.getRange('B4:I4').setBackground('#f8f9fa').setFontColor('#5f6368').setFontSize(9).setWrap(true);
  styleCardBox_(overview.getRange('B2:I4'), '#dadce0', '#1a73e8');

  overview.getRange('B6:I8').setBackground('#ffffff');
  styleCardBox_(overview.getRange('B6:I8'), '#dadce0');
  overview.getRange('B6:G6').setFontColor('#1a73e8').setFontSize(12).setFontWeight('bold');
  overview.getRange('H6:I6').setBackground('#e8f0fe').setFontColor('#1967d2').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');

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

  overview.getRange('B11:I12').setBackground('#ffffff');
  overview.getRange('B11:F11').setFontColor('#5f6368').setFontSize(9).setFontWeight('bold');
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

  overview.getRange('F7:F8').setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build()
  );
  overview.getRange('C12').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(helperSheet.getRange('K2:K500'), true).setAllowInvalid(true).build()
  );
  overview.getRange('F12').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(helperSheet.getRange('L2:L300'), true).setAllowInvalid(true).build()
  );

  styleOverviewV13Body_(overview, helperSheet, '$#,##0.00');
  insertOverviewCharts_(overview, helperSheet, 400, 'USD');

  // 4. Build and style the Ad Unit Breakdown tab.
  const breakdown = ss.getSheetByName('Ad Unit Breakdown') || ss.insertSheet('Ad Unit Breakdown');
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
 * Resets the Ad Unit filter (C12), Country filter (F12), and Demand Channel filter (I12) to "ALL".
 */
function resetSheetFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = getTargetDashboardSheet_(ss);
  if (!dash) return;
  dash.getRange('C12').setValue('ALL');
  dash.getRange('F12').setValue('ALL');
  dash.getRange('I12').setValue('ALL');
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

  // 1. Clear data rows in hidden data tabs while keeping headers.
  ['Ad_Unit_Data', 'Country_Data', 'Daily_Data'].forEach(function(tabName) {
    const sh = ss.getSheetByName(tabName);
    if (sh) {
      const lastRow = sh.getLastRow();
      if (lastRow > 1) {
        sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
      }
    }
  });
  const unitSheet = ss.getSheetByName('Ad_Unit_Data');
  const countrySheet = ss.getSheetByName('Country_Data');
  const dailySheet = ss.getSheetByName('Daily_Data');

  // 2. Reset the hidden Chart_Helper tab while keeping chart ranges cleanly seeded.
  const helperSheet = ss.getSheetByName('Chart_Helper');
  if (helperSheet) {
    const clearRows = Math.max(helperSheet.getLastRow(), 400);
    helperSheet.getRange(2, 1, clearRows, 4).clearContent();
    helperSheet.getRange('A1:D2').setValues([
      ['Date', 'Total Revenue', 'Primary Programmatic Revenue', 'OPMC Revenue'],
      ['—', 0, 0, 0]
    ]);
    helperSheet.getRange('B2:D400').setNumberFormat('$#,##0.00');

    if (helperSheet.getMaxColumns() >= 26) {
      helperSheet.getRange(2, 14, clearRows, 6).clearContent();
      helperSheet.getRange('N1:S2').setValues([
        [
          'Primary Impressions', 'OPMC Impressions', 'Total Impressions',
          'Total eCPM', 'Primary eCPM', 'OPMC eCPM'
        ],
        [0, 0, 0, 0, 0, 0]
      ]);
      helperSheet.getRange('N2:P400').setNumberFormat('#,##0');
      helperSheet.getRange('Q2:S400').setNumberFormat('$#,##0.00');

      helperSheet.getRange('T1:T5').setValues([['Demand_Channel_Options'], ['ALL'], ['Ad Exchange'], ['Open Bidding'], ['SDK Bidding']]);
      helperSheet.getRange(2, 21, clearRows, 3).clearContent();
      helperSheet.getRange('U1:W2').setValues([
        ['OPMC — Ad Exchange', 'OPMC — Open Bidding', 'OPMC — SDK Bidding'],
        [0, 0, 0]
      ]);
      helperSheet.getRange('U2:W400').setNumberFormat('$#,##0.00');
    }

    helperSheet.getRange('F1:I1').setValues([['Country', 'Primary Programmatic Revenue', 'OPMC Revenue', 'Total Revenue']]);
    helperSheet.getRange(2, 6, 15, 4).clearContent();
    helperSheet.getRange('F2:I2').setValues([['—', 0, 0, 0]]);
    helperSheet.getRange('G2:I15').setNumberFormat('$#,##0.00');

    const dashForCharts = getTargetDashboardSheet_(ss);
    if (dashForCharts) {
      syncOverviewCharts_(dashForCharts, helperSheet, 0, 'USD');
    }

    helperSheet.getRange(1, 11, Math.max(helperSheet.getLastRow(), 150), 1).clearContent();
    helperSheet.getRange('K1:K2').setValues([['Ad_Unit_Options'], ['ALL']]);

    helperSheet.getRange(1, 12, Math.max(helperSheet.getLastRow(), 250), 1).clearContent();
    helperSheet.getRange('L1:L2').setValues([['Country_Options'], ['ALL']]);

    helperSheet.getRange('M1:M2').setValues([['Currency_Code'], ['USD']]);
  }

  // 3. Clear data rows in the Ad Unit Breakdown tab.
  try {
    writeParentAdUnitBreakdownSheet_(ss, {
      ad_unit_rows: [],
      currency_code: 'USD',
      currency_symbol: '$'
    });
  } catch (err) {}

  // 4. Reset the filter cells and scorecard currency format on the Overview tab.
  const dash = getTargetDashboardSheet_(ss);
  if (dash) {
    dash.getRange('C12').setValue('ALL');
    dash.getRange('F12').setValue('ALL');
    dash.getRange('I12').setValue('ALL');
    try {
      dash.getRange('B15').setNumberFormat('$#,##0.00');
      dash.getRange('E15').setNumberFormat('$#,##0.00');
      dash.getRange('H15').setNumberFormat('$#,##0.00');
    } catch (e) {}
  }

  // 5. Hide the data and helper tabs.
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
      sheet.getRange('I12').setValue('ALL');
      sheet.getRange('C12').activate();
      SpreadsheetApp.flush();
    } else if (
      cellText.indexOf('Clear') !== -1 ||
      cellText.indexOf('Download') !== -1 ||
      cellText.indexOf('Generate Report') !== -1
    ) {
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
  const childToParent = rollup.childToParent || {};
  const lines = [];
  lines.push('Parent Ad Unit Revenue Summary');
  lines.push('Parent Ad Unit,OPMC Child Units Count,Associated OPMC Ad Units,Primary Programmatic Revenue (' + curr + '),OPMC Revenue (' + curr + '),Total Revenue (' + curr + '),Primary Share (%),OPMC Share (%),Total Impressions');
  (rollup.parentRows || []).forEach(function(pr) {
    lines.push([
      '"' + String(pr.parentName || '').replace(/"/g, '""') + '"',
      pr.childCount,
      '"' + String(pr.childSummary || '').replace(/"/g, '""') + '"',
      pr.adxRev,
      pr.opmcRev,
      pr.totalRev,
      ((pr.adxShare || 0) * 100).toFixed(1),
      ((pr.opmcShare || 0) * 100).toFixed(1),
      pr.totalImps
    ].join(','));
  });

  // Group daily revenue and impression totals by date across all settled channels.
  const dailyMap = {};
  (cached.ad_unit_rows || []).forEach(function(r) {
    const isoDate = formatCsvDateOnly_(r.date, r.raw_date);
    const dKey = String(r.raw_date || isoDate || '').trim();
    if (!dKey) return;
    if (!dailyMap[dKey]) {
      dailyMap[dKey] = {
        date: isoDate,
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
  lines.push('Date,Primary Programmatic Revenue (' + curr + '),OPMC Incremental Revenue (' + curr + '),Total Revenue (' + curr + '),Primary Share (%),OPMC Share (%),OPMC Incremental Uplift (%),Primary eCPM (' + curr + '),OPMC eCPM (' + curr + '),Total eCPM (' + curr + '),Total Impressions');
  sortedDailyKeys.forEach(function(k) {
    const d = dailyMap[k];
    const ar = Math.round(d.adxRev * 100) / 100;
    const or = Math.round(d.opmcRev * 100) / 100;
    const tr = Math.round((ar + or) * 100) / 100;
    const primPct = tr > 0 ? ((ar / tr) * 100).toFixed(1) : '0.0';
    const opmcPct = tr > 0 ? ((or / tr) * 100).toFixed(1) : '0.0';
    const upliftPct = ar > 0 ? ((or / ar) * 100).toFixed(1) : '0.0';
    const primEcpm = d.adxImps > 0 ? ((d.adxRev / d.adxImps) * 1000).toFixed(2) : '0.00';
    const opmcEcpm = d.opmcImps > 0 ? ((d.opmcRev / d.opmcImps) * 1000).toFixed(2) : '0.00';
    const totEcpm = d.totalImps > 0 ? ((d.totalRev / d.totalImps) * 1000).toFixed(2) : '0.00';
    lines.push([
      d.date,
      ar,
      or,
      tr,
      primPct,
      opmcPct,
      upliftPct,
      primEcpm,
      opmcEcpm,
      totEcpm,
      Math.round(d.totalImps)
    ].join(','));
  });

  lines.push('');
  lines.push('Daily Revenue Trends');
  lines.push('Date,Ad Unit Name,Assigned Parent Ad Unit,Demand Channel,Primary Programmatic Revenue (' + curr + '),OPMC Revenue (' + curr + '),Total Revenue (' + curr + '),Total Impressions');
  (cached.ad_unit_rows || []).forEach(function(r) {
    const mappedParent = childToParent[r.ad_unit] || r.parent_ad_unit || '';
    const assignedParent = mappedParent !== r.ad_unit ? mappedParent : '';
    lines.push([
      formatCsvDateOnly_(r.date, r.raw_date),
      '"' + String(r.ad_unit || '').replace(/"/g, '""') + '"',
      '"' + String(assignedParent).replace(/"/g, '""') + '"',
      '"' + String(r.demand_channel || 'Ad Exchange').replace(/"/g, '""') + '"',
      r.adx_revenue,
      r.opmc_revenue,
      r.total_revenue,
      r.total_impressions
    ].join(','));
  });

  lines.push('');
  lines.push('Country Revenue Breakdown');
  lines.push('Date,Country,Demand Channel,Primary Programmatic Revenue (' + curr + '),OPMC Revenue (' + curr + '),Total Revenue (' + curr + '),Total Impressions');
  (cached.country_rows || []).forEach(function(r) {
    lines.push([
      formatCsvDateOnly_(r.date, r.raw_date),
      '"' + String(r.country || '').replace(/"/g, '""') + '"',
      '"' + String(r.demand_channel || 'Ad Exchange').replace(/"/g, '""') + '"',
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
    ad_unit_filter: '',
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
    const b8Label = String(dash.getRange('B8').getDisplayValue() || '').toUpperCase();
    let rawAdUnitFilter = '';
    if (b8Label.indexOf('FILTER') !== -1 || !dash.getRange('C7:C8').isPartOfMerge()) {
      rawAdUnitFilter = String(dash.getRange('C8').getDisplayValue() || '').trim();
    }
    const cleanAdUnitFilter = normalizeAdUnitFilter_(rawAdUnitFilter);
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
      ad_unit_filter: cleanAdUnitFilter,
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
  if (vDict.stringListValue && Array.isArray(vDict.stringListValue.values)) {
    let vals = vDict.stringListValue.values
      .map(function(x) { return String(x || '').trim(); })
      .filter(Boolean);
    if (vals.length > 1 && vals[0].toLowerCase() === 'default') {
      vals = vals.slice(1);
    }
    return vals.length > 0 ? vals.join(' » ') : null;
  }
  return null;
}

/**
 * Extracts the ordered [parent, ..., leaf] ad unit path from AD_UNIT_NAME_ALL_LEVEL.
 */
function extractAdUnitHierarchy_(vDict) {
  if (!vDict || typeof vDict !== 'object') return ['All'];
  if (vDict.stringListValue && Array.isArray(vDict.stringListValue.values)) {
    let vals = vDict.stringListValue.values
      .map(function(x) { return String(x || '').trim(); })
      .filter(Boolean);
    if (vals.length > 1 && vals[0].toLowerCase() === 'default') {
      vals = vals.slice(1);
    }
    if (vals.length > 0) return vals;
  }
  const rawVal = extractVal_(vDict);
  if (rawVal === null || rawVal === undefined) return ['All'];
  const text = String(rawVal).trim();
  if (!text) return ['All'];
  let parts = text
    .split(' » ')
    .map(function(p) { return p.trim(); })
    .filter(Boolean);
  if (parts.length > 1 && parts[0].toLowerCase() === 'default') {
    parts = parts.slice(1);
  }
  return parts.length > 0 ? parts : ['All'];
}

/**
 * Normalizes an optional ad unit substring filter, ignoring blank/ALL/<INSERT_HERE> placeholders.
 */
function normalizeAdUnitFilter_(rawFilter) {
  const val = String(rawFilter || '').trim();
  const upper = val.toUpperCase();
  if (!val || upper === 'ALL' || upper === '<INSERT_HERE>') return '';
  return val;
}

/**
 * Returns true if an ad unit name carries an explicit OPMC or mc1..mcN marker.
 */
function isOpmcUnitName_(name) {
  const lower = String(name || '').toLowerCase();
  if (lower.indexOf('opmc') !== -1 || lower.indexOf('multicall') !== -1) return true;
  return /(?:^|[\s_|/-])mc\d+(?:$|[\s_|/-])/.test(lower);
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
  const base = raw.filter(function(t) {
    if (NOISE_WORDS[t]) return false;
    if (/^\d+$/.test(t)) return false;
    if (/^\d+x\d+$/.test(t)) return false;
    if (/^(?:op)?mc\d+$/.test(t)) return false;
    return true;
  });
  const withoutFmt = base.filter(function(t) { return !FORMAT_WORDS[t]; });
  const hasSpecificPlacement = withoutFmt.some(function(t) { return !PLATFORM_WORDS[t]; });
  const core = hasSpecificPlacement ? withoutFmt : base;
  const formats = {};
  base.forEach(function(t) {
    if (FORMAT_WORDS[t]) formats[t] = true;
  });
  return {
    tokens: core,
    baseTokens: base,
    joined: core.join(''),
    formats: formats,
    isFormatLevel: !hasSpecificPlacement
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
  const unitAncestors = {};
  const manualOverrides = {};
  (adUnitRows || []).forEach(function(r) {
    const u = String(r.ad_unit || '').trim();
    if (!u) return;
    if (!unitAgg[u]) {
      unitAgg[u] = { adxRev: 0, opmcRev: 0, totRev: 0, adxImps: 0, opmcImps: 0, totImps: 0 };
    }
    if (r.manual_parent && String(r.manual_parent).trim()) {
      manualOverrides[u] = String(r.manual_parent).trim();
    }
    let ancList = [];
    if (Array.isArray(r.ancestors) && r.ancestors.length > 0) {
      ancList = r.ancestors.map(function(x) { return String(x || '').trim(); }).filter(Boolean);
    } else if (r.parent_ad_unit && String(r.parent_ad_unit).trim()) {
      ancList = String(r.parent_ad_unit).split(' » ').map(function(x) { return x.trim(); }).filter(Boolean);
    }
    if (ancList.length > 0 && (!unitAncestors[u] || ancList.length > unitAncestors[u].length)) {
      unitAncestors[u] = ancList;
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
  const parentSet = {};
  const parentLowerMap = {};

  allUnits.forEach(function(u) {
    const v = unitAgg[u];
    const isOpmc = isOpmcUnitName_(u) || v.opmcRev > 0 || v.opmcImps > 0;
    if (isOpmc) {
      children.push(u);
    } else {
      parents.push(u);
      parentSet[u] = true;
      parentLowerMap[u.toLowerCase()] = u;
    }
  });

  const parentInfo = {};
  parents.forEach(function(p) {
    const t = tokenizeAdUnitForParentMatch_(p);
    const tokMap = {};
    t.tokens.forEach(function(tok) { tokMap[tok] = true; });
    const baseTokMap = {};
    t.baseTokens.forEach(function(tok) { baseTokMap[tok] = true; });
    const lowerP = p.toLowerCase();
    parentInfo[p] = {
      tokens: t.tokens,
      tokenSet: tokMap,
      baseTokens: t.baseTokens,
      baseTokenSet: baseTokMap,
      formats: t.formats,
      joined: t.joined,
      adxRev: unitAgg[p].adxRev,
      isStagingOrDirect:
        lowerP.indexOf('staging') !== -1 ||
        lowerP.indexOf('direct') !== -1 ||
        lowerP.indexOf('fallback') !== -1 ||
        lowerP.indexOf('dummy') !== -1 ||
        lowerP.indexOf('fake') !== -1
    };
  });

  function isCandidateBetter_(cand, best) {
    for (let i = 0; i < cand.length; i++) {
      if (cand[i] !== best[i]) return cand[i] > best[i];
    }
    return false;
  }

  function matchChildToParent(cName) {
    if (manualOverrides[cName]) {
      return manualOverrides[cName];
    }
    const ancs = unitAncestors[cName] || [];
    if (ancs.length > 0) {
      for (let idx = ancs.length; idx >= 1; idx--) {
        for (let start = 0; start < idx; start++) {
          const candidate = ancs.slice(start, idx).join(' » ');
          if (parentSet[candidate]) return candidate;
          const lowerMatch = parentLowerMap[candidate.toLowerCase()];
          if (lowerMatch) return lowerMatch;
        }
      }
      return ancs.join(' » ');
    }

    const leafParts = String(cName || '').split(' » ');
    const leafOnly = leafParts[leafParts.length - 1].trim();
    let stem = leafOnly.replace(/^opmc[\s_|-]*/i, '');
    stem = stem.replace(/[\s_|-]*\$\d+(?:\.\d+)?$/, '');
    stem = stem.replace(/[\s_|-]*(?:op)?mc\d+$/i, '');
    stem = stem.replace(/[\s_|-]*\d+$/, '').trim();
    for (let i = 0; i < parents.length; i++) {
      if (parents[i].toLowerCase() === stem.toLowerCase()) {
        return parents[i];
      }
    }

    const ct = tokenizeAdUnitForParentMatch_(leafOnly);
    if (!ct.tokens || ct.tokens.length === 0) {
      return leafOnly.replace(/[\s_|/-]*\$\d+(?:\.\d+)?$/, '').trim() + ' (Shared OPMC)';
    }

    let bestP = null;
    let bestTuple = [-1, -1, -1, -1, -1];
    const ctFmtKeys = Object.keys(ct.formats);

    for (let j = 0; j < parents.length; j++) {
      const p = parents[j];
      const pInfo = parentInfo[p];
      const pToks = ct.isFormatLevel ? pInfo.baseTokens : pInfo.tokens;
      const pSet = ct.isFormatLevel ? pInfo.baseTokenSet : pInfo.tokenSet;
      if (!pToks || pToks.length === 0) continue;
      const pJoined = pToks.join('');

      let exactCore = 0;
      let overlapRatio = 0;

      if (ct.joined === pJoined) {
        exactCore = 2;
        overlapRatio = 1.0;
      } else {
        let common = 0;
        for (let k = 0; k < ct.tokens.length; k++) {
          if (pSet[ct.tokens[k]]) common++;
        }
        const minJoinedLen = Math.min(ct.joined.length, pJoined.length);
        if (
          common === 0 &&
          minJoinedLen >= 5 &&
          (ct.joined.indexOf(pJoined) !== -1 || pJoined.indexOf(ct.joined) !== -1)
        ) {
          exactCore = 1;
          overlapRatio = 0.85;
        } else {
          if (common === 0) continue;
          const childCov = common / ct.tokens.length;
          const parentCov = common / pToks.length;
          if (childCov < 1.0) continue;
          overlapRatio = ct.isFormatLevel
            ? 1.0
            : Math.round((childCov * 0.7 + parentCov * 0.3) * 10000) / 10000;
          exactCore = childCov === 1.0 && parentCov === 1.0 ? 1 : 0;
        }
      }

      const pFmtKeys = Object.keys(pInfo.formats);
      let fmtCompat = 0;
      if (ctFmtKeys.length === 0 || pFmtKeys.length === 0) {
        fmtCompat = 2;
      } else {
        const isSubset = ctFmtKeys.every(function(f) { return Boolean(pInfo.formats[f]); });
        if (isSubset && ctFmtKeys.length === pFmtKeys.length) {
          fmtCompat = 2;
        } else if (isSubset) {
          fmtCompat = 1;
        } else {
          fmtCompat = 0;
        }
      }

      const notStaging = pInfo.isStagingOrDirect ? 0 : 1;
      const candTuple = [fmtCompat, exactCore, overlapRatio, notStaging, pInfo.adxRev];
      if (isCandidateBetter_(candTuple, bestTuple)) {
        bestTuple = candTuple;
        bestP = p;
      }
    }

    return bestP || (String(cName || '').replace(/[\s_|/-]*\$\d+(?:\.\d+)?$/, '').trim() + ' (Shared OPMC)');
  }

  const parentGroups = {};
  parents.forEach(function(p) {
    parentGroups[p] = { children: [] };
  });
  const childToParent = {};
  children.sort(compareOpmcChildUnitsNaturally_);
  children.forEach(function(c) {
    const p = matchChildToParent(c);
    childToParent[c] = p;
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
    childToParent: childToParent,
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

  grid[1][1] = 'Parent Ad Unit Revenue Breakdown';
  grid[3][1] =
    'Table 1 combines Primary programmatic revenue (Ad Exchange, Open Bidding, and SDK Bidding) on each Parent Ad Unit (A) with OPMC revenue from its Child Ad Units (A1 to An). Table 2 lists each ad unit and its assigned Parent Ad Unit, updating automatically when you change the Demand Channel filter on Overview (cell I12) or edit Column D.';
  grid[5][1] = 'PARENT & OPMC INVENTORY';
  grid[5][3] = 'TOTAL PRIMARY REVENUE';
  grid[5][5] = 'TOTAL OPMC REVENUE';
  grid[5][7] = 'TOTAL UNIFIED REVENUE';
  grid[7][1] = 'Automatic Hierarchy & Naming Match (A ↔ A1..An)';
  grid[9][1] = '▎ Table 1: Parent Ad Unit Revenue Summary (Parent A + Associated OPMC Units A1..An)';
  grid[9][7] = '="Channel: "&IF(OR(Overview!I12="ALL",Overview!I12=""),"All (AdX+OB+SDKB)",Overview!I12)&" • Live Rollup"';

  // Write summary formulas in rows 7 and 8.
  grid[6][1] = numParents + ' Parent Groups  •  ' + rollup.totalChildren + ' OPMC Units';
  grid[6][3] = '=D' + t1TotalRow;
  grid[6][5] = '=E' + t1TotalRow;
  grid[6][7] = '=F' + t1TotalRow;

  grid[7][3] = '="Primary Share: "&TEXT(G' + t1TotalRow + ', "0.0%")';
  grid[7][5] = '="OPMC Share: "&TEXT(H' + t1TotalRow + ', "0.0%")';
  grid[7][7] = '="OPMC Uplift over Primary: +"&TEXT(IF(D' + t1TotalRow + '>0, E' + t1TotalRow + '/D' + t1TotalRow + ', 0), "0.0%")';

  grid[t1HeaderRow - 1][1] = 'Parent Ad Unit (A)';
  grid[t1HeaderRow - 1][2] = 'Associated OPMC Child Units (A1..An)';
  grid[t1HeaderRow - 1][3] = 'Primary Revenue';
  grid[t1HeaderRow - 1][4] = 'OPMC Revenue (A1..An)';
  grid[t1HeaderRow - 1][5] = 'Total Unified Revenue';
  grid[t1HeaderRow - 1][6] = 'Primary Share (%)';
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
  grid[t2SectionRow - 1][1] = '▎ Table 2: Detailed Parent → OPMC Child Breakdown & Mapping (A & Individual A1..An Units)';

  grid[t2HeaderRow - 1][1] = 'Individual Ad Unit Name (A or A1..An)';
  grid[t2HeaderRow - 1][2] = 'Ad Unit Role';
  grid[t2HeaderRow - 1][3] = 'Assigned Parent Ad Unit (A)';
  grid[t2HeaderRow - 1][4] = 'Primary Revenue';
  grid[t2HeaderRow - 1][5] = 'OPMC Revenue';
  grid[t2HeaderRow - 1][6] = 'Total Revenue';
  grid[t2HeaderRow - 1][7] = 'OPMC Share (%)';
  grid[t2HeaderRow - 1][8] = 'Total Impressions';

  const chCritExpr = 'IF(OR(Overview!$I$12="ALL",Overview!$I$12=""),"*",Overview!$I$12)';
  for (let j = 0; j < detailRows.length; j++) {
    const rNum = t2StartRow + j;
    const dr = detailRows[j];
    const crit = gsCriteriaLiteral_(dr.rawUnit);
    grid[rNum - 1][1] = dr.displayUnit;
    grid[rNum - 1][2] = dr.role;
    grid[rNum - 1][3] = dr.assignedParent;
    grid[rNum - 1][4] = '=SUMIFS(Ad_Unit_Data!$D$2:$D, Ad_Unit_Data!$C$2:$C, "' + crit + '", Ad_Unit_Data!$K$2:$K, ' + chCritExpr + ')';
    grid[rNum - 1][5] = '=SUMIFS(Ad_Unit_Data!$E$2:$E, Ad_Unit_Data!$C$2:$C, "' + crit + '", Ad_Unit_Data!$K$2:$K, ' + chCritExpr + ')';
    grid[rNum - 1][6] = '=E' + rNum + '+F' + rNum;
    grid[rNum - 1][7] = '=IF(G' + rNum + '>0, F' + rNum + '/G' + rNum + ', 0)';
    grid[rNum - 1][8] = '=SUMIFS(Ad_Unit_Data!$I$2:$I, Ad_Unit_Data!$C$2:$C, "' + crit + '", Ad_Unit_Data!$K$2:$K, ' + chCritExpr + ')';
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
/**
 * Builds GAM v1 report filters: settled DEMAND_CHANNEL IN + optional AD_UNIT_NAME_ALL_LEVEL CONTAINS.
 */
function buildReportFilters_(cleanAdUnitFilter) {
  const filters = [{
    fieldFilter: {
      field: { dimension: 'DEMAND_CHANNEL' },
      operation: 'IN',
      values: SETTLED_DEMAND_CHANNELS.map(function(ch) {
        return { stringValue: ch };
      })
    }
  }];
  if (cleanAdUnitFilter) {
    filters.push({
      fieldFilter: {
        field: { dimension: 'AD_UNIT_NAME_ALL_LEVEL' },
        operation: 'CONTAINS',
        values: [{ stringValue: cleanAdUnitFilter }]
      }
    });
  }
  return filters;
}

/**
 * Converts a UTC millisecond timestamp into a GAM API Date object and ISO string.
 */
function utcMsToDateInfo_(utcMs) {
  const dt = new Date(utcMs);
  const y = dt.getUTCFullYear();
  const m = dt.getUTCMonth() + 1;
  const d = dt.getUTCDate();
  return {
    str: y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2),
    obj: { year: y, month: m, day: d }
  };
}

/**
 * Splits a YYYY-MM-DD range into contiguous windows of at most maxChunkDays (default 62).
 */
function splitDateRangeChunks_(startDateStr, endDateStr, maxChunkDays) {
  const chunkDays = Math.max(Number(maxChunkDays) || 62, 1);
  const sObj = parseDateObj_(startDateStr);
  const eObj = parseDateObj_(endDateStr);
  const startMs = Date.UTC(sObj.year, sObj.month - 1, sObj.day);
  const endMs = Date.UTC(eObj.year, eObj.month - 1, eObj.day);
  if (endMs < startMs) {
    throw new Error('Invalid Date Range: End Date (' + endDateStr + ') is before Start Date (' + startDateStr + ').');
  }

  const dayMs = 86400000;
  const windows = [];
  let curMs = startMs;
  while (curMs <= endMs) {
    const wEndMs = Math.min(endMs, curMs + (chunkDays - 1) * dayMs);
    const sInfo = utcMsToDateInfo_(curMs);
    const eInfo = utcMsToDateInfo_(wEndMs);
    windows.push({
      startMs: curMs,
      endMs: wEndMs,
      spanDays: Math.round((wEndMs - curMs) / dayMs),
      startStr: sInfo.str,
      endStr: eInfo.str,
      startObj: sInfo.obj,
      endObj: eInfo.obj
    });
    curMs = wEndMs + dayMs;
  }
  return windows;
}

/**
 * Returns true if an API error message represents a transient or timeout failure.
 */
function isTransientReportError_(errText) {
  const msg = String(errText || '');
  return (
    msg.indexOf('DEADLINE_EXCEEDED') !== -1 ||
    msg.indexOf('UNAVAILABLE') !== -1 ||
    msg.indexOf('INTERNAL') !== -1 ||
    msg.indexOf('timed out') !== -1 ||
    msg.indexOf('SERVER_ERROR') !== -1
  );
}

/**
 * Executes a wave of report window tasks (create -> run -> poll -> fetchRows) with bisection support.
 */
function runUnifiedReportPipeline(params) {
  const effectiveParams = params || readParamsFromDashboardSheet_();
  const networkCode = String(effectiveParams.network_code || '').trim();
  const startDateStr = String(effectiveParams.start_date || '').trim();
  const endDateStr = String(effectiveParams.end_date || '').trim();
  const cleanAdUnitFilter = normalizeAdUnitFilter_(effectiveParams.ad_unit_filter);

  if (!networkCode) {
    throw new Error('Missing Network Code: Please enter your Google Ad Manager Network Code in cell C7 on the Overview tab.');
  }
  if (!startDateStr || !endDateStr) {
    throw new Error('Missing Date Range: Please select both Start Date (cell F7) and End Date (cell F8) on the Overview tab.');
  }

  const apiConfig = resolveApiConfig_(effectiveParams);
  const windows = splitDateRangeChunks_(startDateStr, endDateStr, 62);
  const baseUrl = apiConfig.baseUrl;
  const ts = new Date().getTime();
  const commonFilters = buildReportFilters_(cleanAdUnitFilter);

  const reportSpecs = [
    {
      key: 'h_unit',
      reportType: 'HISTORICAL',
      dimensions: ['DATE', 'AD_UNIT_NAME_ALL_LEVEL', 'DEMAND_CHANNEL'],
      metrics: ['IMPRESSIONS', 'REVENUE_WITHOUT_CPD'],
      expandedCompatibility: false,
      filters: commonFilters
    },
    {
      key: 'o_unit',
      reportType: 'ON_PLATFORM_MULTICALL',
      dimensions: ['DATE', 'AD_UNIT_NAME_ALL_LEVEL', 'DEMAND_CHANNEL'],
      metrics: ['ON_PLATFORM_MULTIPLE_CALL_IMPRESSIONS', 'ON_PLATFORM_MULTIPLE_CALL_REVENUE'],
      expandedCompatibility: false,
      filters: commonFilters
    },
    {
      key: 'h_country',
      reportType: 'HISTORICAL',
      dimensions: ['DATE', 'COUNTRY_NAME', 'DEMAND_CHANNEL'],
      metrics: ['IMPRESSIONS', 'REVENUE_WITHOUT_CPD'],
      expandedCompatibility: false,
      filters: commonFilters
    },
    {
      key: 'o_country',
      reportType: 'ON_PLATFORM_MULTICALL',
      dimensions: ['DATE', 'COUNTRY_NAME', 'DEMAND_CHANNEL'],
      metrics: ['ON_PLATFORM_MULTIPLE_CALL_IMPRESSIONS', 'ON_PLATFORM_MULTIPLE_CALL_REVENUE'],
      expandedCompatibility: false,
      filters: commonFilters
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

  const taskQueue = [];
  windows.forEach(function(win, wIdx) {
    reportSpecs.forEach(function(spec, specIdx) {
      taskQueue.push({
        specIdx: specIdx,
        spec: spec,
        startMs: win.startMs,
        endMs: win.endMs,
        startObj: win.startObj,
        endObj: win.endObj,
        tag: 'w' + wIdx,
        attempt: 0
      });
    });
  });

  const allRowsByReport = [[], [], [], []];
  const createUrl = baseUrl + '/networks/' + networkCode + '/reports';
  const maxPolls = 90;
  const maxWaveSize = 16;

  while (taskQueue.length > 0) {
    const wave = taskQueue.splice(0, maxWaveSize);
    const createRequests = wave.map(function(task, idx) {
      const reportDef = {
        dimensions: task.spec.dimensions,
        metrics: task.spec.metrics,
        dateRange: {
          fixed: {
            startDate: task.startObj,
            endDate: task.endObj
          }
        },
        reportType: task.spec.reportType
      };
      if (task.spec.expandedCompatibility) {
        reportDef.expandedCompatibility = true;
      }
      if (task.spec.filters && task.spec.filters.length > 0) {
        reportDef.filters = task.spec.filters;
      }
      return buildFetchRequest_(
        createUrl,
        'POST',
        {
          displayName: 'AppsScript_' + task.spec.reportType + '_' + task.specIdx + '_' + task.tag + '_' + idx + '_' + ts,
          reportDefinition: reportDef
        },
        apiConfig
      );
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
    const operations = executeBatchRequests_(runRequests, apiConfig, 3);

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

    const completedTasks = [];
    for (let i = 0; i < wave.length; i++) {
      const task = wave[i];
      const op = operations[i];
      if (op && op.done && !op.error) {
        const respData = op.response || {};
        const opParts = String(op.name || '').split('/');
        const opId = opParts[opParts.length - 1];
        completedTasks.push({
          task: task,
          resultName: respData.reportResult || (reportNames[i] + '/results/' + opId),
          pageToken: null
        });
      } else {
        const errStr = op && op.error
          ? JSON.stringify(op.error)
          : ('Report operation timed out: ' + (op ? op.name : task.tag));
        const dayMs = 86400000;
        const spanDays = Math.round((task.endMs - task.startMs) / dayMs);
        if (isTransientReportError_(errStr) && spanDays >= 7) {
          const midMs = task.startMs + Math.floor(spanDays / 2) * dayMs;
          const nextMs = midMs + dayMs;
          const leftEnd = utcMsToDateInfo_(midMs);
          const rightStart = utcMsToDateInfo_(nextMs);
          taskQueue.unshift(
            {
              specIdx: task.specIdx,
              spec: task.spec,
              startMs: task.startMs,
              endMs: midMs,
              startObj: task.startObj,
              endObj: leftEnd.obj,
              tag: task.tag + 'a',
              attempt: 0
            },
            {
              specIdx: task.specIdx,
              spec: task.spec,
              startMs: nextMs,
              endMs: task.endMs,
              startObj: rightStart.obj,
              endObj: task.endObj,
              tag: task.tag + 'b',
              attempt: 0
            }
          );
        } else if (isTransientReportError_(errStr) && task.attempt < 2) {
          Utilities.sleep(2000 * (task.attempt + 1));
          taskQueue.unshift({
            specIdx: task.specIdx,
            spec: task.spec,
            startMs: task.startMs,
            endMs: task.endMs,
            startObj: task.startObj,
            endObj: task.endObj,
            tag: task.tag,
            attempt: task.attempt + 1
          });
        } else {
          throw new Error('Report operation failed (' + task.spec.key + '): ' + errStr);
        }
      }
    }

    let activeFetch = completedTasks;
    while (activeFetch.length > 0) {
      const fetchRequests = activeFetch.map(function(entry) {
        let fetchUrl = baseUrl + '/' + entry.resultName + ':fetchRows?pageSize=5000';
        if (entry.pageToken) {
          fetchUrl += '&pageToken=' + encodeURIComponent(entry.pageToken);
        }
        return buildFetchRequest_(fetchUrl, 'GET', null, apiConfig);
      });
      const pageResults = executeBatchRequests_(fetchRequests, apiConfig, 3);
      const nextActive = [];
      for (let k = 0; k < activeFetch.length; k++) {
        const entry = activeFetch[k];
        const resData = pageResults[k] || {};
        const rows = resData.rows || [];
        for (let r = 0; r < rows.length; r++) {
          allRowsByReport[entry.task.specIdx].push(rows[r]);
        }
        if (resData.nextPageToken) {
          entry.pageToken = resData.nextPageToken;
          nextActive.push(entry);
        }
      }
      activeFetch = nextActive;
    }
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
      ad_unit_filter: cleanAdUnitFilter,
      api_mode: params && params.api_mode,
      quota_project: apiConfig.quotaProject
    });
  } catch (sheetErr) {
    output.sheetSyncWarning = sheetErr.message;
  }

  return output;
}

/**
 * Combines the Primary Programmatic and OPMC report rows by date, ad unit, country, and demand channel.
 */
function aggregateStreams_(currencyCode, currSymbol, rawHUnit, rawOUnit, rawHCountry, rawOCountry, modeLabel) {
  const rawAdUnitDict = {};
  const leafToPaths = {};

  function addAdUnitRow(r, isOpmc) {
    const rawDims = r.dimensionValues || [];
    const dateVal = rawDims.length > 0 ? extractVal_(rawDims[0]) : null;
    const date = dateVal !== null ? String(dateVal) : 'All';
    const pathList = rawDims.length > 1 ? extractAdUnitHierarchy_(rawDims[1]) : ['All'];
    const pathKey = pathList.join(' » ');
    const leaf = pathList[pathList.length - 1];
    if (!leafToPaths[leaf]) leafToPaths[leaf] = {};
    leafToPaths[leaf][pathKey] = true;
    const rawCh = rawDims.length > 2 ? extractVal_(rawDims[2]) : 'AD_EXCHANGE';
    const demandChannel = normalizeDemandChannel_(rawCh);

    const metricGroups = r.metricValueGroups || [];
    const mets = metricGroups.length > 0 ? (metricGroups[0].primaryValues || []).map(extractVal_) : [0, 0];
    const imps = mets.length > 0 && mets[0] !== null ? Number(mets[0]) : 0.0;
    const rev = mets.length > 1 && mets[1] !== null ? Number(mets[1]) : 0.0;

    const key = date + '|||' + pathKey + '|||' + demandChannel;
    if (!rawAdUnitDict[key]) {
      rawAdUnitDict[key] = {
        date: date,
        path: pathList,
        demand_channel: demandChannel,
        adx_rev: 0.0,
        opmc_rev: 0.0,
        adx_imps: 0.0,
        opmc_imps: 0.0
      };
    }
    if (isOpmc) {
      rawAdUnitDict[key].opmc_rev += rev;
      rawAdUnitDict[key].opmc_imps += imps;
    } else {
      rawAdUnitDict[key].adx_rev += rev;
      rawAdUnitDict[key].adx_imps += imps;
    }
  }

  (rawHUnit || []).forEach(function(r) { addAdUnitRow(r, false); });
  (rawOUnit || []).forEach(function(r) { addAdUnitRow(r, true); });

  const adUnitDict = {};
  Object.keys(rawAdUnitDict).forEach(function(rawKey) {
    const item = rawAdUnitDict[rawKey];
    const pathList = item.path;
    const leaf = pathList[pathList.length - 1];
    const ancestors = pathList.slice(0, pathList.length - 1);
    const hasCollision = Object.keys(leafToPaths[leaf] || {}).length > 1;
    const adUnitDisplay = hasCollision ? pathList.join(' » ') : leaf;
    const demandChannel = item.demand_channel;
    const key = item.date + '|||' + adUnitDisplay + '|||' + demandChannel;
    if (!adUnitDict[key]) {
      adUnitDict[key] = {
        date: item.date,
        ad_unit: adUnitDisplay,
        ancestors: ancestors,
        parent_ad_unit: ancestors.join(' » '),
        demand_channel: demandChannel,
        adx_rev: 0.0,
        opmc_rev: 0.0,
        adx_imps: 0.0,
        opmc_imps: 0.0
      };
    }
    adUnitDict[key].adx_rev += item.adx_rev;
    adUnitDict[key].opmc_rev += item.opmc_rev;
    adUnitDict[key].adx_imps += item.adx_imps;
    adUnitDict[key].opmc_imps += item.opmc_imps;
  });

  const countryDict = {};

  function addCountryRow(r, isOpmc) {
    const dimVals = (r.dimensionValues || []).map(extractVal_);
    const date = dimVals.length > 0 && dimVals[0] !== null ? String(dimVals[0]) : 'All';
    const country = dimVals.length > 1 && dimVals[1] !== null ? String(dimVals[1]) : 'Unknown';
    const rawCh = dimVals.length > 2 && dimVals[2] !== null ? dimVals[2] : 'AD_EXCHANGE';
    const demandChannel = normalizeDemandChannel_(rawCh);

    const metricGroups = r.metricValueGroups || [];
    const mets = metricGroups.length > 0 ? (metricGroups[0].primaryValues || []).map(extractVal_) : [0, 0];
    const imps = mets.length > 0 && mets[0] !== null ? Number(mets[0]) : 0.0;
    const rev = mets.length > 1 && mets[1] !== null ? Number(mets[1]) : 0.0;

    const key = date + '|||' + country + '|||' + demandChannel;
    if (!countryDict[key]) {
      countryDict[key] = {
        date: date,
        country: country,
        demand_channel: demandChannel,
        adx_rev: 0.0,
        opmc_rev: 0.0,
        adx_imps: 0.0,
        opmc_imps: 0.0
      };
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
    const demandChannel = val.demand_channel || 'Ad Exchange';

    totalAdxRev += adxRev;
    totalOpmcRev += opmcRev;
    totalAdxImps += adxImp;
    totalOpmcImps += opmcImp;

    const cleanD = String(val.date).trim();
    if (!dailyDict[cleanD]) {
      dailyDict[cleanD] = {
        adx: 0.0,
        opmc: 0.0,
        tot: 0.0,
        adx_imps: 0.0,
        opmc_imps: 0.0,
        tot_imps: 0.0,
        opmc_adx: 0.0,
        opmc_ob: 0.0,
        opmc_sdkb: 0.0,
        prim_adx: 0.0,
        prim_ob: 0.0,
        prim_sdkb: 0.0
      };
    }
    dailyDict[cleanD].adx += adxRev;
    dailyDict[cleanD].opmc += opmcRev;
    dailyDict[cleanD].tot += totRev;
    dailyDict[cleanD].adx_imps += adxImp;
    dailyDict[cleanD].opmc_imps += opmcImp;
    dailyDict[cleanD].tot_imps += totImp;
    if (demandChannel === 'Ad Exchange') {
      dailyDict[cleanD].opmc_adx += opmcRev;
      dailyDict[cleanD].prim_adx += adxRev;
    } else if (demandChannel === 'Open Bidding') {
      dailyDict[cleanD].opmc_ob += opmcRev;
      dailyDict[cleanD].prim_ob += adxRev;
    } else if (demandChannel === 'SDK Bidding') {
      dailyDict[cleanD].opmc_sdkb += opmcRev;
      dailyDict[cleanD].prim_sdkb += adxRev;
    }

    if (val.ad_unit) adUnitSet[val.ad_unit] = true;

    adUnitRows.push({
      date: formatDateLabel_(cleanD),
      raw_date: cleanD,
      ad_unit: val.ad_unit,
      parent_ad_unit: val.parent_ad_unit || '',
      ancestors: val.ancestors || [],
      demand_channel: demandChannel,
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
      demand_channel: val.demand_channel || 'Ad Exchange',
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
  const dailyOpmcAdx = sortedDates.map(function(d) { return Math.round(dailyDict[d].opmc_adx * 100) / 100; });
  const dailyOpmcOb = sortedDates.map(function(d) { return Math.round(dailyDict[d].opmc_ob * 100) / 100; });
  const dailyOpmcSdkb = sortedDates.map(function(d) { return Math.round(dailyDict[d].opmc_sdkb * 100) / 100; });
  const dailyPrimAdx = sortedDates.map(function(d) { return Math.round(dailyDict[d].prim_adx * 100) / 100; });
  const dailyPrimOb = sortedDates.map(function(d) { return Math.round(dailyDict[d].prim_ob * 100) / 100; });
  const dailyPrimSdkb = sortedDates.map(function(d) { return Math.round(dailyDict[d].prim_sdkb * 100) / 100; });

  const rollupPreview = buildParentAdUnitRollup_(adUnitRows);
  const childToParent = rollupPreview.childToParent || {};
  (rollupPreview.parentRows || []).forEach(function(pr) {
    const pName = String(pr.parentName || '').trim();
    if (pName) adUnitSet[pName] = true;
  });

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
    child_to_parent: childToParent,
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
      total_impressions: dailyTotalImps,
      opmc_adx: dailyOpmcAdx,
      opmc_ob: dailyOpmcOb,
      opmc_sdkb: dailyOpmcSdkb,
      prim_adx: dailyPrimAdx,
      prim_ob: dailyPrimOb,
      prim_sdkb: dailyPrimSdkb
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

  try {
    upgradeToV13LayoutIfNeeded_(ss, data.currency_symbol || '$');
  } catch (upgErr) {}

  const childToParent = data.child_to_parent || (buildParentAdUnitRollup_(data.ad_unit_rows || []).childToParent || {});

  const unitSheet = ss.getSheetByName('Ad_Unit_Data') || ss.insertSheet('Ad_Unit_Data');
  unitSheet.clearContents();
  if (unitSheet.getMaxColumns() < 11) {
    unitSheet.insertColumnsAfter(unitSheet.getMaxColumns(), 11 - unitSheet.getMaxColumns());
  }
  const unitHeader = [[
    'Date', 'Raw_Date', 'Ad_Unit', 'Primary_Revenue', 'OPMC_Revenue', 'Total_Revenue',
    'Primary_Impressions', 'OPMC_Impressions', 'Total_Impressions', 'Parent_Ad_Unit', 'Demand_Channel'
  ]];
  const unitValues = (data.ad_unit_rows || []).map(function(r) {
    let resolvedParent = childToParent[r.ad_unit] || '';
    if (resolvedParent === r.ad_unit) resolvedParent = '';
    return [
      r.date, r.raw_date, r.ad_unit, r.adx_revenue, r.opmc_revenue, r.total_revenue,
      r.adx_impressions || 0, r.opmc_impressions || 0, r.total_impressions, resolvedParent, r.demand_channel || 'Ad Exchange'
    ];
  });
  const allUnitRows = unitHeader.concat(unitValues);
  if (unitSheet.getMaxRows() < allUnitRows.length + 5) {
    unitSheet.insertRowsAfter(unitSheet.getMaxRows(), allUnitRows.length + 5 - unitSheet.getMaxRows());
  }
  unitSheet.getRange(1, 1, allUnitRows.length, 11).setValues(allUnitRows);

  const countrySheet = ss.getSheetByName('Country_Data') || ss.insertSheet('Country_Data');
  countrySheet.clearContents();
  if (countrySheet.getMaxColumns() < 10) {
    countrySheet.insertColumnsAfter(countrySheet.getMaxColumns(), 10 - countrySheet.getMaxColumns());
  }
  const countryHeader = [[
    'Date', 'Raw_Date', 'Country', 'Primary_Revenue', 'OPMC_Revenue', 'Total_Revenue',
    'Primary_Impressions', 'OPMC_Impressions', 'Total_Impressions', 'Demand_Channel'
  ]];
  const countryValues = (data.country_rows || []).map(function(r) {
    return [
      r.date, r.raw_date, r.country, r.adx_revenue, r.opmc_revenue, r.total_revenue,
      r.adx_impressions || 0, r.opmc_impressions || 0, r.total_impressions, r.demand_channel || 'Ad Exchange'
    ];
  });
  const allCountryRows = countryHeader.concat(countryValues);
  if (countrySheet.getMaxRows() < allCountryRows.length + 5) {
    countrySheet.insertRowsAfter(countrySheet.getMaxRows(), allCountryRows.length + 5 - countrySheet.getMaxRows());
  }
  countrySheet.getRange(1, 1, allCountryRows.length, 10).setValues(allCountryRows);

  const cSym = String(data.currency_symbol || '$').replace(/"/g, '');
  const currFmt = cSym === '$' ? '$#,##0.00' : '"' + cSym + '"#,##0.00';

  const dailySheet = ss.getSheetByName('Daily_Data') || ss.insertSheet('Daily_Data');
  dailySheet.clearContents();
  if (dailySheet.getMaxColumns() < 14) {
    dailySheet.insertColumnsAfter(dailySheet.getMaxColumns(), 14 - dailySheet.getMaxColumns());
  }
  const dailyHeader = [[
    'Date', 'Raw_Date', 'Primary Programmatic Revenue', 'OPMC Revenue', 'Total Revenue',
    'Primary Impressions', 'OPMC Impressions', 'Total Impressions',
    'OPMC Ad Exchange Revenue', 'OPMC Open Bidding Revenue', 'OPMC SDK Bidding Revenue',
    'Primary Ad Exchange Revenue', 'Primary Open Bidding Revenue', 'Primary SDK Bidding Revenue'
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
      (data.daily.total_impressions && data.daily.total_impressions[idx]) || 0,
      (data.daily.opmc_adx && data.daily.opmc_adx[idx]) || 0,
      (data.daily.opmc_ob && data.daily.opmc_ob[idx]) || 0,
      (data.daily.opmc_sdkb && data.daily.opmc_sdkb[idx]) || 0,
      (data.daily.prim_adx && data.daily.prim_adx[idx]) || 0,
      (data.daily.prim_ob && data.daily.prim_ob[idx]) || 0,
      (data.daily.prim_sdkb && data.daily.prim_sdkb[idx]) || 0
    ];
  });
  const allDailyRows = dailyHeader.concat(dailyValues);
  if (dailySheet.getMaxRows() < allDailyRows.length + 5) {
    dailySheet.insertRowsAfter(dailySheet.getMaxRows(), allDailyRows.length + 5 - dailySheet.getMaxRows());
  }
  dailySheet.getRange(1, 1, allDailyRows.length, 14).setValues(allDailyRows);
  try {
    if (dailyValues.length > 0) {
      dailySheet.getRange(2, 3, dailyValues.length, 3).setNumberFormat(currFmt);
      dailySheet.getRange(2, 6, dailyValues.length, 3).setNumberFormat('#,##0');
      dailySheet.getRange(2, 9, dailyValues.length, 6).setNumberFormat(currFmt);
    }
  } catch (e) {}

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
    dash.getRange('C12').setValue('ALL');
    dash.getRange('F12').setValue('ALL');
    dash.getRange('I12').setValue('ALL');
    try {
      dash.getRange('B15:C15').setNumberFormat(currFmt);
      dash.getRange('E15:F15').setNumberFormat(currFmt);
      dash.getRange('H15:I15').setNumberFormat(currFmt);
    } catch (e) {}
  }

  // Update Chart_Helper daily formulas, channel breakdown, top 10 countries, options, and currency code.
  const helperSheet = ss.getSheetByName('Chart_Helper');
  if (helperSheet) {
    try {
      if (helperSheet.getMaxColumns() < 26) {
        helperSheet.insertColumnsAfter(helperSheet.getMaxColumns(), 26 - helperSheet.getMaxColumns());
      }
      const numDaily = (data.daily && data.daily.labels) ? data.daily.labels.length : 0;
      const neededRows = Math.max(numDaily + 5, 405);
      if (helperSheet.getMaxRows() < neededRows) {
        helperSheet.insertRowsAfter(helperSheet.getMaxRows(), neededRows - helperSheet.getMaxRows());
      }

      helperSheet.getRange('M1:M2').setValues([['Currency_Code'], [data.currency_code || 'USD']]);
      helperSheet.getRange('T1:T5').setValues([['Demand_Channel_Options'], ['ALL'], ['Ad Exchange'], ['Open Bidding'], ['SDK Bidding']]);
      helperSheet.getRange('A1:D1').setValues([['Date', 'Total Revenue', 'Primary Programmatic Revenue', 'OPMC Revenue']]);
      helperSheet.getRange('N1:S1').setValues([[
        'Primary Impressions', 'OPMC Impressions', 'Total Impressions',
        'Total eCPM', 'Primary eCPM', 'OPMC eCPM'
      ]]);
      helperSheet.getRange('U1:W1').setValues([['OPMC — Ad Exchange', 'OPMC — Open Bidding', 'OPMC — SDK Bidding']]);

      const clearSpan = Math.max(helperSheet.getLastRow(), 400);
      helperSheet.getRange(2, 1, clearSpan, 4).clearContent();
      helperSheet.getRange(2, 14, clearSpan, 6).clearContent();
      helperSheet.getRange(2, 21, clearSpan, 3).clearContent();

      const chCrit = 'IF(OR(Overview!$I$12="ALL",Overview!$I$12=""),"*",Overview!$I$12)';
      if (numDaily > 0) {
        const helperDailyFormulas = [];
        const helperEcpmFormulas = [];
        const helperOpmcChFormulas = [];
        for (let dIdx = 0; dIdx < numDaily; dIdx++) {
          const rNum = dIdx + 2;
          helperDailyFormulas.push([
            '=Daily_Data!A' + rNum,
            '=C' + rNum + '+D' + rNum,
            '=IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$D$2:$D, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12, Ad_Unit_Data!$K$2:$K, ' + chCrit + ') + SUMIFS(Ad_Unit_Data!$D$2:$D, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$J$2:$J, Overview!$C$12, Ad_Unit_Data!$K$2:$K, ' + chCrit + '), IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$D$2:$D, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12, Country_Data!$J$2:$J, ' + chCrit + '), IF(OR(Overview!$I$12="ALL",Overview!$I$12=""), Daily_Data!C' + rNum + ', SUMIFS(Ad_Unit_Data!$D$2:$D, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$K$2:$K, Overview!$I$12))))',
            '=IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$E$2:$E, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12, Ad_Unit_Data!$K$2:$K, ' + chCrit + ') + SUMIFS(Ad_Unit_Data!$E$2:$E, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$J$2:$J, Overview!$C$12, Ad_Unit_Data!$K$2:$K, ' + chCrit + '), IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$E$2:$E, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12, Country_Data!$J$2:$J, ' + chCrit + '), IF(OR(Overview!$I$12="ALL",Overview!$I$12=""), Daily_Data!D' + rNum + ', SUMIFS(Ad_Unit_Data!$E$2:$E, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$K$2:$K, Overview!$I$12))))'
          ]);
          helperEcpmFormulas.push([
            '=IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$G$2:$G, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12, Ad_Unit_Data!$K$2:$K, ' + chCrit + ') + SUMIFS(Ad_Unit_Data!$G$2:$G, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$J$2:$J, Overview!$C$12, Ad_Unit_Data!$K$2:$K, ' + chCrit + '), IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$G$2:$G, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12, Country_Data!$J$2:$J, ' + chCrit + '), IF(OR(Overview!$I$12="ALL",Overview!$I$12=""), Daily_Data!F' + rNum + ', SUMIFS(Ad_Unit_Data!$G$2:$G, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$K$2:$K, Overview!$I$12))))',
            '=IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$H$2:$H, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12, Ad_Unit_Data!$K$2:$K, ' + chCrit + ') + SUMIFS(Ad_Unit_Data!$H$2:$H, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$J$2:$J, Overview!$C$12, Ad_Unit_Data!$K$2:$K, ' + chCrit + '), IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$H$2:$H, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12, Country_Data!$J$2:$J, ' + chCrit + '), IF(OR(Overview!$I$12="ALL",Overview!$I$12=""), Daily_Data!G' + rNum + ', SUMIFS(Ad_Unit_Data!$H$2:$H, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$K$2:$K, Overview!$I$12))))',
            '=N' + rNum + '+O' + rNum,
            '=IF(P' + rNum + '>0, (B' + rNum + '/P' + rNum + ')*1000, 0)',
            '=IF(N' + rNum + '>0, (C' + rNum + '/N' + rNum + ')*1000, 0)',
            '=IF(O' + rNum + '>0, (D' + rNum + '/O' + rNum + ')*1000, 0)'
          ]);
          const chPairs = [['Ad Exchange', 'I'], ['Open Bidding', 'J'], ['SDK Bidding', 'K']];
          helperOpmcChFormulas.push(chPairs.map(function(pair) {
            const chName = pair[0];
            const dailyCol = pair[1];
            return (
              '=IF(AND(Overview!$I$12<>"ALL", Overview!$I$12<>"", Overview!$I$12<>"' + chName + '"), 0, ' +
              'IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$E$2:$E, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12, Ad_Unit_Data!$K$2:$K, "' + chName + '") + ' +
              'SUMIFS(Ad_Unit_Data!$E$2:$E, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$J$2:$J, Overview!$C$12, Ad_Unit_Data!$K$2:$K, "' + chName + '"), ' +
              'IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$E$2:$E, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12, Country_Data!$J$2:$J, "' + chName + '"), ' +
              'Daily_Data!' + dailyCol + rNum + ')))'
            );
          }));
        }
        helperSheet.getRange(2, 1, numDaily, 4).setFormulas(helperDailyFormulas);
        helperSheet.getRange(2, 14, numDaily, 6).setFormulas(helperEcpmFormulas);
        helperSheet.getRange(2, 21, numDaily, 3).setFormulas(helperOpmcChFormulas);
        try {
          const fmtSpan = Math.max(numDaily, 400);
          helperSheet.getRange(2, 2, fmtSpan, 3).setNumberFormat(currFmt);
          helperSheet.getRange(2, 14, fmtSpan, 3).setNumberFormat('#,##0');
          helperSheet.getRange(2, 17, fmtSpan, 3).setNumberFormat(currFmt);
          helperSheet.getRange(2, 21, fmtSpan, 3).setNumberFormat(currFmt);
        } catch (fmtErr) {}
      } else {
        const emptyLbl = (params && params.start_date && formatDateLabel_(params.start_date)) || '—';
        helperSheet.getRange('A2:D2').setValues([[emptyLbl, 0, 0, 0]]);
        helperSheet.getRange('N2:S2').setValues([[0, 0, 0, 0, 0, 0]]);
        helperSheet.getRange('U2:W2').setValues([[0, 0, 0]]);
        try {
          helperSheet.getRange(2, 2, 400, 3).setNumberFormat(currFmt);
          helperSheet.getRange(2, 14, 400, 3).setNumberFormat('#,##0');
          helperSheet.getRange(2, 17, 400, 3).setNumberFormat(currFmt);
          helperSheet.getRange(2, 21, 400, 3).setNumberFormat(currFmt);
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
      helperSheet.getRange('F1:I1').setValues([['Country', 'Primary Programmatic Revenue', 'OPMC Revenue', 'Total Revenue']]);
      const countryNames = [];
      const countryFormulas = [];
      for (let t = 0; t < 10; t++) {
        const rNum = t + 2;
        countryNames.push([top10Countries[t] || '']);
        countryFormulas.push([
          '=IF(F' + rNum + '="", 0, IF(OR(Overview!$F$12="ALL", Overview!$F$12=F' + rNum + '), SUMIFS(Country_Data!$D$2:$D, Country_Data!$C$2:$C, F' + rNum + ', Country_Data!$J$2:$J, ' + chCrit + '), 0))',
          '=IF(F' + rNum + '="", 0, IF(OR(Overview!$F$12="ALL", Overview!$F$12=F' + rNum + '), SUMIFS(Country_Data!$E$2:$E, Country_Data!$C$2:$C, F' + rNum + ', Country_Data!$J$2:$J, ' + chCrit + '), 0))',
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
        dash.getRange('C12').setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInRange(helperSheet.getRange(2, 11, Math.max(unitOpts.length, 500), 1), true).setAllowInvalid(true).build()
        );
        dash.getRange('F12').setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInRange(helperSheet.getRange(2, 12, Math.max(countryOpts.length, 300), 1), true).setAllowInvalid(true).build()
        );
        dash.getRange('I12').setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInRange(helperSheet.getRange('T2:T5'), true).setAllowInvalid(true).build()
        );
        syncOverviewCharts_(dash, helperSheet, numDaily, data.currency_code || 'USD');
      }
    } catch (helperErr) {}
  }

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
 * Updates the exact daily row range and currency titles across the 5 v1.3 Overview charts.
 */
function syncOverviewCharts_(overview, helperSheet, numDaily, currencyCode) {
  if (!overview || !helperSheet) return;
  try {
    const endRow = Math.max((Number(numDaily) || 0) + 1, 2);
    const curr = String(currencyCode || 'USD').trim().toUpperCase() || 'USD';
    const charts = overview.getCharts() || [];
    if (charts.length !== 5) {
      insertOverviewCharts_(overview, helperSheet, Number(numDaily) || 0, curr);
      return;
    }
    const vAxisStyle = { fontName: 'Roboto', fontSize: 10 };
    charts.forEach(function(ch) {
      try {
        const info = ch.getContainerInfo();
        const r = info ? info.getAnchorRow() : 0;
        if (r >= 18 && r <= 22) {
          const updated1 = ch
            .modify()
            .clearRanges()
            .addRange(helperSheet.getRange(1, 1, endRow, 4))
            .setOption('title', 'Revenue Trends (' + curr + ')')
            .setOption('vAxis', { title: 'Revenue (' + curr + ')', textStyle: vAxisStyle })
            .build();
          overview.updateChart(updated1);
        } else if (r >= 37 && r <= 41) {
          const updated2 = ch
            .modify()
            .clearRanges()
            .addRange(helperSheet.getRange(1, 1, endRow, 1))
            .addRange(helperSheet.getRange(1, 21, endRow, 3))
            .setOption('vAxis', { title: 'OPMC Revenue (' + curr + ')', textStyle: vAxisStyle })
            .build();
          overview.updateChart(updated2);
        } else if (r >= 56 && r <= 60) {
          const updated3 = ch
            .modify()
            .clearRanges()
            .addRange(helperSheet.getRange(1, 1, endRow, 1))
            .addRange(helperSheet.getRange(1, 17, endRow, 3))
            .setOption('title', 'eCPM Trends (' + curr + ')')
            .setOption('vAxis', { title: 'eCPM (' + curr + ')', textStyle: vAxisStyle })
            .build();
          overview.updateChart(updated3);
        } else if (r >= 75 && r <= 79) {
          const updated4 = ch
            .modify()
            .clearRanges()
            .addRange(helperSheet.getRange(1, 1, endRow, 1))
            .addRange(helperSheet.getRange(1, 3, endRow, 2))
            .setOption('vAxis', { title: 'Combined Revenue (' + curr + ')', textStyle: vAxisStyle })
            .build();
          overview.updateChart(updated4);
        } else if (r >= 94 && r <= 98) {
          const updated5 = ch
            .modify()
            .clearRanges()
            .addRange(helperSheet.getRange(1, 6, 11, 3))
            .setOption('vAxis', { title: 'Revenue (' + curr + ')', textStyle: vAxisStyle })
            .build();
          overview.updateChart(updated5);
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

    const manualParentMap = {};
    try {
      const breakdownSheet = ss.getSheetByName('Ad Unit Breakdown');
      if (breakdownSheet && breakdownSheet.getLastRow() > 14) {
        const bdVals = breakdownSheet.getDataRange().getValues();
        for (let b = 0; b < bdVals.length; b++) {
          const disp = String(bdVals[b][1] || '');
          const role = String(bdVals[b][2] || '');
          const assigned = String(bdVals[b][3] || '').trim();
          if (role.indexOf('OPMC Child Unit') !== -1 && assigned) {
            const rawChild = disp.replace(/^[\s↳]+/, '').trim();
            if (rawChild) manualParentMap[rawChild] = assigned;
          }
        }
      }
    } catch (bdErr) {}

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
      const parentStr = String(r[9] || '').trim();
      const demandChannel = String(r[10] || 'Ad Exchange').trim() || 'Ad Exchange';

      adUnitRows.push({
        date: d,
        raw_date: rawD,
        ad_unit: unit,
        parent_ad_unit: parentStr,
        manual_parent: manualParentMap[unit] || '',
        demand_channel: demandChannel,
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
      const adxImp = Number(cr[6]) || 0;
      const opmcImp = Number(cr[7]) || 0;
      const imp = Number(cr[8] || (adxImp + opmcImp)) || 0;
      const demandChannel = String(cr[9] || 'Ad Exchange').trim() || 'Ad Exchange';

      countryRows.push({
        date: d,
        raw_date: rawD,
        country: country,
        demand_channel: demandChannel,
        adx_revenue: Math.round(adx * 100) / 100,
        opmc_revenue: Math.round(opmc * 100) / 100,
        total_revenue: Math.round(tot * 100) / 100,
        adx_impressions: Math.round(adxImp),
        opmc_impressions: Math.round(opmcImp),
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