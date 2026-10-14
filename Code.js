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
 * Author: Francesco Fresta (fresta@google.com)
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
 * Requests OAuth authorization for the Google Ad Manager API, checks the
 * granted scopes, and adds the action buttons to the Overview tab.
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
  } catch (e) {
    Logger.log('Note: requireScopes check: ' + e);
  }

  const tokInfo = acquireValidGamToken_('');
  Logger.log('✅ SUCCESS! Google Ad Manager OAuth token acquired via ' + tokInfo.source);
  Logger.log('Granted scopes: ' + tokInfo.scopes);
  try {
    installDashboardButtons_();
  } catch (e) {}
  return tokInfo;
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
    // Clear daily helper rows in columns A to D.
    helperSheet.getRange(2, 1, 150, 4).clearContent();

    // Reset the top 10 countries table in columns F to I.
    helperSheet.getRange('F1:I1').setValues([['Country', 'Ad Exchange Revenue', 'On-platform multicall Revenue', 'Total Revenue']]);
    helperSheet.getRange(2, 6, 15, 4).clearContent();

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

  // 6. Reset the filter cells on the Overview tab to "ALL".
  const dash = getTargetDashboardSheet_(ss);
  if (dash) {
    dash.getRange('C12').setValue('ALL');
    dash.getRange('F12').setValue('ALL');
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
    } else if (cellText.indexOf('Clear') !== -1) {
      sheet.getRange('C12').activate();
      clearReportData();
    } else if (cellText.indexOf('Download') !== -1) {
      sheet.getRange('C12').activate();
      downloadCsvFromSheet();
    } else if (cellText.indexOf('Generate Report') !== -1) {
      sheet.getRange('C7').activate();
      runReportFromSheetParams();
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
      'OAUTH SCOPE MISSING (HTTP 403 ACCESS_TOKEN_SCOPE_INSUFFICIENT):\n' +
        'Your Apps Script token currently has scopes: [' + (checkScript.scopes || 'none') + ']\n' +
        'and is missing the required canonical scope "https://www.googleapis.com/auth/admanager".\n' +
        'In the Apps Script Editor, open appsscript.json, ensure "https://www.googleapis.com/auth/admanager" is in "oauthScopes" (without dfp), Save, and click ▷ Run on authorizeAndVerifyGamScope.'
    );
  }

  throw new Error('Could not acquire any OAuth2 token. Please authorize the script.');
}

/**
 * Cleans the Google Cloud Project ID value from cell I8.
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
 * Checks whether a value is a valid Google Cloud Project ID and not a three-letter currency code.
 */
function isValidQuotaProject_(val) {
  const raw = String(val || '').trim();
  if (!raw || /^[A-Z]{3}$/.test(raw) || CURRENCY_SYMBOLS[raw.toUpperCase()]) return false;
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

  // Group daily revenue totals by date for the Stacked Incremental Revenue section.
  const dailyMap = {};
  (cached.ad_unit_rows || []).forEach(function(r) {
    const dKey = String(r.raw_date || r.date || '').trim();
    if (!dKey) return;
    if (!dailyMap[dKey]) {
      dailyMap[dKey] = { date: r.date, raw_date: dKey, adxRev: 0, opmcRev: 0, totalRev: 0, totalImps: 0 };
    }
    const ar = Number(r.adx_revenue) || 0;
    const or = Number(r.opmc_revenue) || 0;
    const tr = Number(r.total_revenue) || (ar + or);
    const ti = Number(r.total_impressions) || 0;
    dailyMap[dKey].adxRev += ar;
    dailyMap[dKey].opmcRev += or;
    dailyMap[dKey].totalRev += tr;
    dailyMap[dKey].totalImps += ti;
  });
  const sortedDailyKeys = Object.keys(dailyMap).sort();

  lines.push('');
  lines.push('Stacked Incremental Revenue');
  lines.push('Date,Ad Exchange Revenue (' + curr + '),OPMC Incremental Revenue (' + curr + '),Total Revenue (' + curr + '),AdX Share (%),OPMC Share (%),OPMC Incremental Uplift (%),Total Impressions');
  sortedDailyKeys.forEach(function(k) {
    const d = dailyMap[k];
    const ar = Math.round(d.adxRev * 100) / 100;
    const or = Math.round(d.opmcRev * 100) / 100;
    const tr = Math.round((ar + or) * 100) / 100;
    const adxPct = tr > 0 ? ((ar / tr) * 100).toFixed(1) + '%' : '0.0%';
    const opmcPct = tr > 0 ? ((or / tr) * 100).toFixed(1) + '%' : '0.0%';
    const upliftPct = ar > 0 ? '+' + ((or / ar) * 100).toFixed(1) + '%' : '0.0%';
    lines.push([
      d.date,
      ar,
      or,
      tr,
      adxPct,
      opmcPct,
      upliftPct,
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
 * - I8: Google Cloud Project ID
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

    const netCode = String(dash.getRange('C7').getDisplayValue() || '').trim();
    const apiModeVal = String(dash.getRange('C8').getDisplayValue() || '').trim();
    const startVal = readDateCellIso_(dash.getRange('F7'), ss);
    const e8Label = String(dash.getRange('E8').getDisplayValue() || '').toUpperCase();
    const isStackedDatesLayout = e8Label.indexOf('END') !== -1;

    const endVal = readDateCellIso_(dash.getRange(isStackedDatesLayout ? 'F8' : 'I7'), ss);
    let i8Val = String(dash.getRange('I8').getDisplayValue() || '').trim();
    if (!isValidQuotaProject_(i8Val)) {
      const overviewSheet = ss.getSheetByName('Overview');
      if (overviewSheet && overviewSheet.getSheetId() !== dash.getSheetId()) {
        i8Val = String(overviewSheet.getRange('I8').getDisplayValue() || '').trim();
      }
    }

    const apiMode = 'PUBLIC';

    const resolvedQuotaProject = isValidQuotaProject_(i8Val) ? normalizeQuotaProject_(i8Val) : '';

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
        'Missing GCP Project ID (Quota): Please enter your GCP Project ID in cell I8 on the Overview tab (and press Enter to commit the cell) before clicking Generate Report.'
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
      'Missing GCP Project ID (Quota): Please enter your GCP Project ID in cell I8 on the Overview tab (and press Enter to commit the cell) before running the report.'
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
              '):\nConfigured GCP Project ID (Overview!I8): ' +
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
  sheet.getRange(1, 1, grid.length, 10).setValues(grid);

  try {
    const cSym = String(data.currency_symbol || '$').replace(/"/g, '');
    const currFmt = cSym === '$' ? '$#,##0.00' : '"' + cSym + '"#,##0.00';
    if (numParents > 0) {
      sheet.getRange(t1TotalRow, 4, numParents + 1, 3).setNumberFormat(currFmt);
      sheet.getRange(t1TotalRow, 7, numParents + 1, 2).setNumberFormat('0.0%');
      sheet.getRange(t1TotalRow, 9, numParents + 1, 1).setNumberFormat('#,##0');
    }
    if (detailRows.length > 0) {
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
      dailyDict[cleanD] = { adx: 0.0, opmc: 0.0, tot: 0.0 };
    }
    dailyDict[cleanD].adx += adxRev;
    dailyDict[cleanD].opmc += opmcRev;
    dailyDict[cleanD].tot += totRev;

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
      total: dailyTotal
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
  countrySheet.getRange(1, 1, allCountryRows.length, 9).setValues(allCountryRows);

  let dailySheet = ss.getSheetByName('Daily_Data') || ss.insertSheet('Daily_Data');
  dailySheet.clearContents();
  const dailyHeader = [['Date', 'Raw_Date', 'Ad Exchange Revenue', 'On-platform multicall Revenue', 'Total Revenue']];
  const dailyValues = (data.daily.labels || []).map(function(lbl, idx) {
    return [
      lbl,
      (data.daily.raw_dates && data.daily.raw_dates[idx]) || lbl,
      data.daily.adx[idx] || 0,
      data.daily.opmc[idx] || 0,
      data.daily.total[idx] || 0
    ];
  });
  const allDailyRows = dailyHeader.concat(dailyValues);
  dailySheet.getRange(1, 1, allDailyRows.length, 5).setValues(allDailyRows);

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
  }

  // Update daily rows, top 10 countries, filter dropdown lists, and currency code in Chart_Helper.
  const helperSheet = ss.getSheetByName('Chart_Helper');
  if (helperSheet) {
    try {
      helperSheet.getRange('M1:M2').setValues([['Currency_Code'], [data.currency_code || 'USD']]);

      const numDaily = (data.daily && data.daily.labels) ? data.daily.labels.length : 0;
      helperSheet.getRange(2, 1, 100, 4).clearContent();
      if (numDaily > 0) {
        const helperDailyFormulas = [];
        for (let dIdx = 0; dIdx < numDaily; dIdx++) {
          const rNum = dIdx + 2;
          helperDailyFormulas.push([
            '=Daily_Data!A' + rNum,
            '=C' + rNum + '+D' + rNum,
            '=IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$D$2:$D, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12), IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$D$2:$D, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12), Daily_Data!C' + rNum + '))',
            '=IF(Overview!$C$12<>"ALL", SUMIFS(Ad_Unit_Data!$E$2:$E, Ad_Unit_Data!$A$2:$A, A' + rNum + ', Ad_Unit_Data!$C$2:$C, Overview!$C$12), IF(Overview!$F$12<>"ALL", SUMIFS(Country_Data!$E$2:$E, Country_Data!$A$2:$A, A' + rNum + ', Country_Data!$C$2:$C, Overview!$F$12), Daily_Data!D' + rNum + '))'
          ]);
        }
        helperSheet.getRange(2, 1, numDaily, 4).setFormulas(helperDailyFormulas);
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
        helperSheet.getRange('G2:I11').setNumberFormat('$#,##0.00');
      } catch (fmtErr) {}

      const unitOpts = [['Ad_Unit_Options'], ['ALL']].concat((data.ad_units || []).map(function(u) { return [u]; }));
      helperSheet.getRange(1, 11, Math.max(unitOpts.length, 150), 1).clearContent();
      helperSheet.getRange(1, 11, unitOpts.length, 1).setValues(unitOpts);

      const countryOpts = [['Country_Options'], ['ALL']].concat((data.countries || []).map(function(c) { return [c]; }));
      helperSheet.getRange(1, 12, Math.max(countryOpts.length, 250), 1).clearContent();
      helperSheet.getRange(1, 12, countryOpts.length, 1).setValues(countryOpts);
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