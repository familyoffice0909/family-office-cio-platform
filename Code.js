const FO_CONFIG = {
  LEDGER_SPREADSHEET_ID: '1_NIOTk1bC0QilRDfo8nKshoLh9Xdm1FWdDybsNvAo8k',
  DASHBOARD_SPREADSHEET_ID: '13jHJ0N1Gzbia7B4FIHkdTqf2tCp1tkXSFicQ2ti8M1w',
  PLATFORM_VERSION: 'v1.0.0',
  RELEASE_NAME: 'Operational Analytics',
  ENGINE_VERSION: 'v2.5',
  BUILD: '2026.07.08',
  BASELINE: 'CB-002',
  BASE_CURRENCY: 'CAD'
};

const FO_TABS = {
  RECOMMENDATIONS: 'Recommendations',
  MARKET_ACCESS: 'Canadian Market Access Library',
  ORCHESTRATION_LOG: 'Orchestration Log',
  REPORT_ARCHIVE: 'Report Archive',
  DATA_INTEGRITY: 'Data Integrity'
};

function foLedger_() {
  return SpreadsheetApp.openById(FO_CONFIG.LEDGER_SPREADSHEET_ID);
}

function foDashboard_() {
  return SpreadsheetApp.openById(FO_CONFIG.DASHBOARD_SPREADSHEET_ID);
}

function foEnsureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0 && headers) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function foNowId_(prefix) {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  return prefix + '-' + ts;
}

function foLog_(channel, action, status, message) {
  const sheet = foEnsureSheet_(foLedger_(), FO_TABS.ORCHESTRATION_LOG,
    ['Timestamp', 'Run ID', 'Channel', 'Action', 'Status', 'Message', 'Version']);
  sheet.appendRow([new Date(), foNowId_('EVT'), channel, action, status, message, FO_CONFIG.VERSION]);
}

function getVal_(row, headers, name) {
  const index = headers.indexOf(name);
  return index >= 0 ? row[index] : '';
}

function num_(value) {
  if (value === '' || value === null || value === undefined) return '';
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/[$,%\s,]/g, '');
  const n = Number(cleaned);
  return isNaN(n) ? '' : n;
}

function foSetupPhaseB() {
  const ledger = foLedger_();

  foEnsureSheet_(ledger, FO_TABS.ORCHESTRATION_LOG,
    ['Timestamp', 'Run ID', 'Channel', 'Action', 'Status', 'Message', 'Version']);

  foEnsureSheet_(ledger, FO_TABS.REPORT_ARCHIVE,
    ['Timestamp', 'Report ID', 'Report Type', 'Title', 'Materiality Score', 'Action Plan Changed?', 'Report Link', 'Notes']);

  foEnsureSheet_(ledger, FO_TABS.DATA_INTEGRITY,
    ['Timestamp', 'Check Type', 'Status', 'Details', 'Resolved?']);

  foEnsureSheet_(ledger, FO_TABS.MARKET_ACCESS,
    ['Ticker', 'Company', 'Native Listing', 'Canadian Common Share', 'Canadian CDR', 'Canadian ETF Alternative', 'Preferred Vehicle', 'Market Access Type', 'Last Reviewed', 'Liquidity Assessment', 'Tracking Assessment', 'CIO Notes']);

  foLog_('setup', 'Setup Phase B', 'SUCCESS', 'Service layer initialized.');
}

function foAppendRecommendationEvent(event) {
  const headers = [
    'Recommendation ID','Event ID','Event Timestamp','Event Type','CIO Version','Investment','Ticker','Preferred Vehicle','Market Access Type','Account(s)','Recommendation Category','Status at Issuance','Current Status','Lifecycle Status','Materiality Score','Conviction Score','Buy Zone Confidence','CIO Readiness %','Native Price','Native Currency','CAD Equivalent','Prepare Zone CAD','Buy Zone CAD','Exceptional Zone CAD','Target / Horizon','Planned Review Trigger','Review Date','Benchmark','Rationale','Source / Notes','Outcome Status','Evidence Status'
  ];

  const sheet = foEnsureSheet_(foLedger_(), FO_TABS.RECOMMENDATIONS, headers);
  const recId = event.recommendationId || foNowId_('REC');
  const eventId = foNowId_('EVT');

  sheet.appendRow([
    recId, eventId, new Date(), event.eventType || 'New Recommendation', FO_CONFIG.VERSION,
    event.investment || '', event.ticker || '', event.preferredVehicle || '', event.marketAccessType || '', event.accounts || '',
    event.recommendationCategory || '', event.statusAtIssuance || '', event.currentStatus || '', event.lifecycleStatus || 'Active',
    event.materialityScore || '', event.convictionScore || '', event.buyZoneConfidence || '', event.cioReadiness || '',
    event.nativePrice || '', event.nativeCurrency || '', event.cadEquivalent || '', event.prepareZoneCAD || '', event.buyZoneCAD || '', event.exceptionalZoneCAD || '',
    event.targetHorizon || '', event.plannedReviewTrigger || '', event.reviewDate || '', event.benchmark || '', event.rationale || '', event.sourceNotes || '',
    event.outcomeStatus || 'Evidence Pending', event.evidenceStatus || 'Evidence Pending'
  ]);

  foLog_(event.channel || 'ledger', 'Append Recommendation Event', 'SUCCESS', recId + ' / ' + eventId);
  return { recommendationId: recId, eventId: eventId };
}

function foUpsertMarketAccess(record) {
  const sheet = foEnsureSheet_(foLedger_(), FO_TABS.MARKET_ACCESS,
    ['Ticker', 'Company', 'Native Listing', 'Canadian Common Share', 'Canadian CDR', 'Canadian ETF Alternative', 'Preferred Vehicle', 'Market Access Type', 'Last Reviewed', 'Liquidity Assessment', 'Tracking Assessment', 'CIO Notes']);

  const values = sheet.getDataRange().getValues();
  let targetRow = -1;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toUpperCase() === String(record.ticker).toUpperCase()) {
      targetRow = i + 1;
      break;
    }
  }

  const row = [
    record.ticker || '',
    record.company || '',
    record.nativeListing || '',
    record.canadianCommonShare || '',
    record.canadianCDR || '',
    record.canadianETF || '',
    record.preferredVehicle || '',
    record.marketAccessType || '',
    new Date(),
    record.liquidityAssessment || '',
    record.trackingAssessment || '',
    record.notes || ''
  ];

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  foLog_('market-access', 'Upsert Market Access', 'SUCCESS', record.ticker || 'unknown');
}

function foSeedKnownCDRs() {
  foUpsertMarketAccess({
    ticker: 'MU',
    company: 'Micron',
    nativeListing: 'MU US',
    canadianCDR: 'Micron HG CDR',
    preferredVehicle: 'Micron HG CDR',
    marketAccessType: 'CDR',
    liquidityAssessment: 'Pending',
    trackingAssessment: 'Pending',
    notes: 'Seeded CDR candidate.'
  });

  foUpsertMarketAccess({
    ticker: 'AVGO',
    company: 'Broadcom',
    nativeListing: 'AVGO US',
    canadianCDR: 'Broadcom CDR',
    preferredVehicle: 'Broadcom CDR',
    marketAccessType: 'CDR',
    liquidityAssessment: 'Pending',
    trackingAssessment: 'Pending',
    notes: 'Seeded CDR candidate.'
  });

  foUpsertMarketAccess({
    ticker: 'QCOM',
    company: 'Qualcomm',
    nativeListing: 'QCOM US',
    canadianCDR: 'Qualcomm HG CDR',
    preferredVehicle: 'Qualcomm HG CDR',
    marketAccessType: 'CDR',
    liquidityAssessment: 'Pending',
    trackingAssessment: 'Pending',
    notes: 'Seeded CDR candidate.'
  });

  foUpsertMarketAccess({
    ticker: 'CLS',
    company: 'Celestica',
    nativeListing: 'CLS US / TSX context to verify',
    canadianCDR: 'Celestica CDR',
    preferredVehicle: 'Canadian access to be evaluated',
    marketAccessType: 'CDR/Common Share Review',
    liquidityAssessment: 'Pending',
    trackingAssessment: 'Pending',
    notes: 'Seeded CDR candidate.'
  });
}

function foAppendDashboardMetric(metric) {
  const sheet = foEnsureSheet_(foDashboard_(), 'Executive Dashboard',
    ['Timestamp', 'Metric', 'Value', 'Status', 'Notes']);

  sheet.appendRow([new Date(), metric.name || '', metric.value || '', metric.status || '', metric.notes || '']);
  foLog_('dashboard', 'Append Dashboard Metric', 'SUCCESS', metric.name || 'metric');
}

function foRunLedgerIntegrityCheck() {
  const rec = foLedger_().getSheetByName(FO_TABS.RECOMMENDATIONS);
  let status = 'PASS';
  const issues = [];

  if (!rec) {
    status = 'FAIL';
    issues.push('Recommendations tab missing.');
  } else {
    const data = rec.getDataRange().getValues();
    const headers = data[0] || [];

    if (headers.indexOf('Recommendation ID') < 0) {
      status = 'FAIL';
      issues.push('Recommendation ID column missing.');
    }

    if (headers.indexOf('Event ID') < 0) {
      status = 'FAIL';
      issues.push('Event ID column missing.');
    }

    if (headers.indexOf('Planned Review Trigger') < 0) {
      status = 'FAIL';
      issues.push('Planned Review Trigger column missing.');
    }

    if (data.length <= 1) {
      issues.push('No recommendation events recorded yet.');
    }
  }

  const sheet = foEnsureSheet_(foLedger_(), FO_TABS.DATA_INTEGRITY,
    ['Timestamp', 'Check Type', 'Status', 'Details', 'Resolved?']);

  sheet.appendRow([new Date(), 'Ledger Integrity Check', status, issues.join(' | ') || 'No structural issues detected.', status === 'PASS' ? 'Yes' : 'No']);

  foLog_('integrity', 'Ledger Integrity Check', status === 'PASS' ? 'SUCCESS' : 'WARNING', issues.join(' | '));

  return { status: status, issues: issues };
}

function foArchiveReport(report) {
  const sheet = foEnsureSheet_(foLedger_(), FO_TABS.REPORT_ARCHIVE,
    ['Timestamp', 'Report ID', 'Report Type', 'Title', 'Materiality Score', 'Action Plan Changed?', 'Report Link', 'Notes']);

  const reportId = report.reportId || foNowId_('RPT');

  sheet.appendRow([
    new Date(),
    reportId,
    report.reportType || '',
    report.title || '',
    report.materialityScore || '',
    report.actionPlanChanged || '',
    report.reportLink || '',
    report.notes || ''
  ]);

  foLog_(report.reportType || 'report', 'Archive Report', 'SUCCESS', reportId);
  return reportId;
}

function foRunPhaseBSmokeTest() {
  foSetupPhaseB();
  foSeedKnownCDRs();

  foAppendRecommendationEvent({
    channel: 'smoke-test',
    investment: 'Sample Only',
    ticker: 'SAMPLE',
    preferredVehicle: 'Sample Vehicle',
    marketAccessType: 'Test',
    accounts: 'TFSA',
    recommendationCategory: 'Watch for Trigger',
    statusAtIssuance: 'Sample',
    currentStatus: 'Sample',
    materialityScore: 0,
    convictionScore: 0,
    buyZoneConfidence: 0,
    cioReadiness: 0,
    plannedReviewTrigger: 'Manual Review',
    rationale: 'Smoke test for Phase B service layer.',
    sourceNotes: 'Generated by Apps Script smoke test.'
  });

  foAppendDashboardMetric({
    name: 'Phase B Smoke Test',
    value: 'Completed',
    status: 'PASS',
    notes: 'Dashboard write path validated.'
  });

  foRunLedgerIntegrityCheck();

  foArchiveReport({
    reportType: 'Smoke Test',
    title: 'Phase B Orchestration Smoke Test',
    materialityScore: 0,
    actionPlanChanged: 'No',
    notes: 'Validated service calls.'
  });
}

function foLoadPortfolioMaster_() {
  const dashboard = foDashboard_();
  const sheet = dashboard.getSheetByName('Portfolio Master');
  const map = {};

  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return map;

  const headers = values[0].map(h => String(h).trim());

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const ticker = String(getVal_(row, headers, 'Ticker')).trim();

    if (!ticker) continue;

    map[ticker.toUpperCase()] = {
      company: getVal_(row, headers, 'Company / Fund'),
      preferredVehicle: getVal_(row, headers, 'Preferred Vehicle'),
      marketAccessType: getVal_(row, headers, 'Market Access Type'),
      assetClass: getVal_(row, headers, 'Asset Class'),
      theme: getVal_(row, headers, 'Theme'),
      sector: getVal_(row, headers, 'Sector'),
      nativeCurrency: getVal_(row, headers, 'Native Currency'),
      targetWeight: getVal_(row, headers, 'Target Weight'),
      riskTier: getVal_(row, headers, 'Risk Tier'),
      dataQuality: getVal_(row, headers, 'Data Quality')
    };
  }

  return map;
}

function foBuildPortfolioStateSnapshot() {
  foRebuildPortfolioStateFromSourceSheets();
}

function foRebuildPortfolioStateFromSourceSheets() {
  const dashboard = foDashboard_();

  const stateSheet = foEnsureSheet_(dashboard, 'Portfolio State', [
    'Timestamp', 'Account', 'Ticker', 'Name', 'Vehicle', 'Market Access Type',
    'Quantity', 'Native Currency', 'Native Price', 'FX Rate',
    'Market Value CAD', 'Cost Basis CAD', 'Unrealized P&L CAD', 'Unrealized P&L %',
    'Asset Class', 'Theme', 'Sector', 'Target Weight', 'Current Weight',
    'Drift', 'Status', 'Notes'
  ]);

  const masterMap = foLoadPortfolioMaster_();

  if (stateSheet.getLastRow() > 1) {
    stateSheet.getRange(2, 1, stateSheet.getLastRow() - 1, 22).clearContent();
  }

  const now = new Date();
  const rows = [];

  const sources = [
    { sheetName: 'TFSA Holdings', account: 'TFSA', status: 'Active' },
    { sheetName: 'LIRA Holdings', account: 'LIRA', status: 'Active' },
    { sheetName: 'Interactive Brokers', account: 'Interactive Brokers', status: 'Active' },
    { sheetName: 'Watchlists', account: 'Watchlist', status: 'Watchlist' }
  ];

  sources.forEach(source => {
    const sheet = dashboard.getSheetByName(source.sheetName);
    if (!sheet) return;

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return;

    const headers = values[0].map(h => String(h).trim());

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const ticker = String(getVal_(row, headers, 'Ticker')).trim();

      if (!ticker) continue;

      const master = masterMap[ticker.toUpperCase()] || {};

      const quantity =
        num_(getVal_(row, headers, 'Quantity')) ||
        num_(getVal_(row, headers, 'Shares'));

      const nativePrice =
        num_(getVal_(row, headers, 'Native Price')) ||
        num_(getVal_(row, headers, 'Current Price'));

      const fxRate =
        num_(getVal_(row, headers, 'FX Rate')) ||
        (master.nativeCurrency === 'USD' ? 1.36 : 1);

      const existingMarketValue =
        num_(getVal_(row, headers, 'Market Value CAD')) ||
        num_(getVal_(row, headers, 'Market Value'));

      const marketValueCAD =
        quantity && nativePrice
          ? quantity * nativePrice * fxRate
          : existingMarketValue;

      const costBasisCAD =
        num_(getVal_(row, headers, 'Cost Basis CAD')) ||
        num_(getVal_(row, headers, 'Book Value'));

      const unrealizedCAD =
        marketValueCAD !== '' && costBasisCAD !== ''
          ? marketValueCAD - costBasisCAD
          : '';

      const unrealizedPct =
        unrealizedCAD !== '' && costBasisCAD
          ? unrealizedCAD / costBasisCAD
          : '';

      const targetWeight =
        num_(getVal_(row, headers, 'Target Weight')) ||
        num_(master.targetWeight);

      rows.push([
        now,
        source.account,
        ticker,
        getVal_(row, headers, 'Name') || getVal_(row, headers, 'Security') || master.company || '',
        getVal_(row, headers, 'Vehicle') || master.preferredVehicle || '',
        getVal_(row, headers, 'Market Access Type') || master.marketAccessType || '',
        quantity,
        getVal_(row, headers, 'Native Currency') || master.nativeCurrency || 'CAD',
        nativePrice,
        fxRate,
        marketValueCAD,
        costBasisCAD,
        unrealizedCAD,
        unrealizedPct,
        getVal_(row, headers, 'Asset Class') || master.assetClass || '',
        getVal_(row, headers, 'Theme') || master.theme || '',
        getVal_(row, headers, 'Sector') || master.sector || '',
        targetWeight,
        '',
        '',
        source.status,
        master.dataQuality
          ? 'Master enriched from Portfolio Master | Data Quality: ' + master.dataQuality
          : 'Master enriched from Portfolio Master'
      ]);
    }
  });

  const totalMarketValue = rows.reduce((sum, row) => sum + (Number(row[10]) || 0), 0);

  rows.forEach(row => {
    const marketValueCAD = Number(row[10]) || 0;
    const targetWeight = Number(row[17]);

    const currentWeight =
      totalMarketValue > 0 ? marketValueCAD / totalMarketValue : '';

    const drift =
      currentWeight !== '' && !isNaN(targetWeight)
        ? currentWeight - targetWeight
        : '';

    row[18] = currentWeight;
    row[19] = drift;
  });

  if (rows.length > 0) {
    stateSheet.getRange(2, 1, rows.length, 22).setValues(rows);
  }

  foLog_(
    'portfolio-state',
    'Rebuild Portfolio State from Master',
    'SUCCESS',
    rows.length + ' rows loaded. Total market value CAD: ' + totalMarketValue
  );
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Family Office CIO')
    .addItem('Setup Phase B Services', 'foSetupPhaseB')
    .addItem('Run Phase B Smoke Test', 'foRunPhaseBSmokeTest')
    .addItem('Seed Known CDRs', 'foSeedKnownCDRs')
    .addItem('Run Ledger Integrity Check', 'foRunLedgerIntegrityCheck')
    .addItem('Rebuild Portfolio State', 'foRebuildPortfolioStateFromSourceSheets')
    .addSeparator()
    .addItem('Automation Bootstrap', 'foAutomationBootstrap')
    .addItem('Run Platform Health Check', 'foRunPlatformHealthCheck')
    .addItem('Show Automation Version', 'foShowAutomationVersion')
    .addToUi();
}

/************************************************************
 * Wave 1C.1 — Automation Framework Extension
 * Safe additive extension. Do not delete existing code.
 ************************************************************/

const FO_AUTOMATION = {
  AUTOMATION_VERSION: 'v1.0.0',
  BASELINE: 'CB-002',
  RELEASE: 'Operational Analytics v1.0',
  HEALTH_SHEET: 'Platform Health',
  AUTOMATION_LOG: 'Automation Log',
  REQUIRED_DASHBOARD_SHEETS: [
    'Portfolio Master',
    'Portfolio Dashboard',
    'Ledger',
    'Recommendation Ledger',
    'Recommendation Performance',
    'Portfolio Attribution',
    'Knowledge Base',
    'Executive Dashboard',
    'Daily CIO Report',
    'Weekly Executive Report',
    'Investment Committee Summary'
  ]
};

function foAutomationLog_(level, module, action, message) {
  const sheet = foEnsureSheet_(foDashboard_(), FO_AUTOMATION.AUTOMATION_LOG, [
    'Timestamp',
    'Level',
    'Module',
    'Action',
    'Message',
    'Platform Version',
    'Automation Version',
    'Baseline',
    'User'
  ]);

  sheet.appendRow([
    new Date(),
    level,
    module,
    action,
    message,
    FO_CONFIG.VERSION,
    FO_AUTOMATION.AUTOMATION_VERSION,
    FO_AUTOMATION.BASELINE,
    Session.getActiveUser().getEmail()
  ]);
}

function foAutomationInfo_(module, action, message) {
  foAutomationLog_('INFO', module, action, message);
}

function foAutomationWarn_(module, action, message) {
  foAutomationLog_('WARNING', module, action, message);
}

function foAutomationError_(module, action, error) {
  const message = error && error.stack ? error.stack : String(error);
  foAutomationLog_('ERROR', module, action, message);
}

function foGetAutomationVersion() {
  return {
    platformVersion: FO_CONFIG.VERSION,
    automationVersion: FO_AUTOMATION.AUTOMATION_VERSION,
    baseline: FO_AUTOMATION.BASELINE,
    release: FO_AUTOMATION.RELEASE,
    dashboardSpreadsheetId: FO_CONFIG.DASHBOARD_SPREADSHEET_ID,
    ledgerSpreadsheetId: FO_CONFIG.LEDGER_SPREADSHEET_ID,
    baseCurrency: FO_CONFIG.BASE_CURRENCY
  };
}

function foRunPlatformHealthCheck() {
  const module = 'PlatformHealthCheck';

  try {
    foAutomationInfo_(module, 'Start', 'Platform health check started.');

    const dashboard = foDashboard_();
    const rows = [];
    let failures = 0;

    FO_AUTOMATION.REQUIRED_DASHBOARD_SHEETS.forEach(function(name) {
      const sheet = dashboard.getSheetByName(name);
      const status = sheet ? 'PASS' : 'FAIL';

      if (!sheet) failures++;

      rows.push([
        new Date(),
        name,
        status,
        sheet ? 'Worksheet found.' : 'Worksheet missing.',
        FO_CONFIG.VERSION,
        FO_AUTOMATION.BASELINE
      ]);
    });

    const healthSheet = foEnsureSheet_(dashboard, FO_AUTOMATION.HEALTH_SHEET, [
      'Timestamp',
      'Check',
      'Status',
      'Details',
      'Platform Version',
      'Baseline'
    ]);

    if (healthSheet.getLastRow() > 1) {
      healthSheet.getRange(2, 1, healthSheet.getLastRow() - 1, 6).clearContent();
    }

    healthSheet.getRange(2, 1, rows.length, 6).setValues(rows);

    const result = failures === 0 ? 'PASS' : 'FAIL';

    foAutomationLog_(
      result === 'PASS' ? 'INFO' : 'ERROR',
      module,
      'Complete',
      result === 'PASS'
        ? 'All required dashboard worksheets found.'
        : failures + ' required worksheet(s) missing.'
    );

    return {
      status: result,
      checkedWorksheets: FO_AUTOMATION.REQUIRED_DASHBOARD_SHEETS.length,
      failures: failures
    };

  } catch (error) {
    foAutomationError_(module, 'Failure', error);
    throw error;
  }
}

function foAutomationBootstrap() {
  const module = 'AutomationBootstrap';

  try {
    foAutomationInfo_(module, 'Start', 'Automation bootstrap started.');

    foEnsureSheet_(foDashboard_(), FO_AUTOMATION.AUTOMATION_LOG, [
      'Timestamp',
      'Level',
      'Module',
      'Action',
      'Message',
      'Platform Version',
      'Automation Version',
      'Baseline',
      'User'
    ]);

    foEnsureSheet_(foDashboard_(), FO_AUTOMATION.HEALTH_SHEET, [
      'Timestamp',
      'Check',
      'Status',
      'Details',
      'Platform Version',
      'Baseline'
    ]);

    const health = foRunPlatformHealthCheck();

    foAutomationInfo_(
      module,
      'Complete',
      'Automation bootstrap completed. Health status: ' + health.status
    );

    return {
      status: 'SUCCESS',
      version: foGetAutomationVersion(),
      health: health
    };

  } catch (error) {
    foAutomationError_(module, 'Failure', error);
    throw error;
  }
}

function foShowAutomationVersion() {
  const v = foGetAutomationVersion();

  SpreadsheetApp.getUi().alert(
    'Family Office CIO Platform' +
    '\nPlatform Version: ' + v.platformVersion +
    '\nAutomation Version: ' + v.automationVersion +
    '\nBaseline: ' + v.baseline +
    '\nRelease: ' + v.release +
    '\nBase Currency: ' + v.baseCurrency
  );
}



