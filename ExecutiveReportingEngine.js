
/************************************************************
 * ExecutiveReportingEngine.gs
 * Sprint 2.6.0 — Governed Executive Reporting
 * Sprint 3.1.0 — Portfolio Scenario Intelligence integration
 * Sprint 3.2.0 — Risk Budget Intelligence integration
 ************************************************************/

const FO_MORNING_BRIEF_REQUIRED_DASHBOARD_SHEETS = [
  'Executive Dashboard',
  'Portfolio Master',
  'Portfolio Snapshot',
  'Portfolio Performance Positions',
  'Portfolio Valuation Summary',
  'Portfolio Optimization Summary',
  'Portfolio Scenario Summary',
  'Risk Budget Summary',
  'Investment Decision Support',
  'Executive Decision State A233',
  'Executive CIO Report',
  'Automation Log',
  'Executive Report Archive',
  'Knowledge Base'
];

const FO_MORNING_BRIEF_REQUIRED_LEDGER_SHEETS = [
  'Version History',
  'Canadian Market Access Library',
  'Outcomes',
  'Lessons Learned',
  'Orchestration Log',
  'Report Archive'
];

function foRunExecutiveReportEngine() {
  const module = 'ExecutiveReportingEngine';

  try {
    foInfo_(module, 'Start', 'Executive Report Engine started.');

    const preflight = foRunMorningBriefPreflight_();
    const dashboard = preflight.dashboard;
    const integrationA233Base =
      typeof foRunExecutiveDecisionIntegrationA233 === 'function'
        ? foRunExecutiveDecisionIntegrationA233()
        : null;

    if (!integrationA233Base) {
      throw new Error(
        'Executive Decision Integration A233 is unavailable. Run A2.3.3 first.'
      );
    }

    const integrationA233 =
      typeof foApplyPortfolioScenarioExecutiveIntegration_ === 'function'
        ? foApplyPortfolioScenarioExecutiveIntegration_(
          integrationA233Base,
          dashboard
        )
        : integrationA233Base;

    const decisions = foReadGovernedExecutiveDecisions_(
      dashboard,
      integrationA233
    );

    if (!decisions.length) {
      throw new Error(
        'Investment Decision Support has no governed rows to report.'
      );
    }

    const reportId = foNowId_('EXEC-RPT');

    const valuationSummary = foReadMetricSheet_(
      dashboard,
      'Portfolio Valuation Summary',
      'Metric',
      'Value'
    );

    const summary = foBuildExecutiveSummary_(
      decisions,
      valuationSummary
    );

    const output = foEnsureSheet_(dashboard, 'Executive CIO Report', [
      'Section',
      'Metric / Ticker',
      'Value / Action',
      'Priority',
      'Risk',
      'Notes',
      'Report ID',
      'Platform Version',
      'Baseline',
      'Timestamp'
    ]);

    if (output.getLastRow() > 1) {
      output.getRange(2, 1, output.getLastRow() - 1, 10).clearContent();
    }

    const rows = [];

    foAppendExecutiveDecisionStateRowsA233_(
      rows,
      integrationA233,
      reportId
    );

    rows.push([
      'Executive Summary',
      'Overall CIO Readiness',
      summary.averageReadiness,
      summary.overallPriority,
      summary.portfolioRisk,
      summary.executiveNarrative,
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);

    rows.push([
      'Executive Summary',
      'Total Market Value',
      summary.totalMarketValue,
      '',
      '',
      'Governed Portfolio Valuation Summary — valued positions only.',
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);

    foAppendPortfolioValuationExecutiveRows_(
      rows,
      valuationSummary,
      reportId
    );

    foAppendPortfolioOptimizationExecutiveRows_(dashboard, rows, reportId);
    foAppendPortfolioScenarioExecutiveRows_(
      dashboard,
      rows,
      reportId,
      integrationA233
    );
    foAppendRiskBudgetExecutiveRows_(dashboard, rows, reportId);

    rows.push([
      'Executive Summary',
      'Actions Requiring Review',
      summary.reviewCount,
      '',
      '',
      'Items where CIO review is required before execution.',
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);

    foAppendMaterialActionExplainability_(
      rows,
      integrationA233,
      reportId
    );

    foAppendDecisionSectionA233_(rows, 'Deploy Capital', decisions, ['DEPLOY NOW'], reportId, integrationA233);
    foAppendDecisionSectionA233_(rows, 'Buy / Add', decisions, ['BUY'], reportId, integrationA233);
    foAppendDecisionSectionA233_(rows, 'Accumulate', decisions, ['ACCUMULATE'], reportId, integrationA233);
    foAppendDecisionSectionA233_(rows, 'Hold', decisions, ['HOLD'], reportId, integrationA233);
    foAppendDecisionSectionA233_(rows, 'Watch / Review', decisions, ['WATCH', 'REFRESH DATA'], reportId, integrationA233);
    foAppendDecisionSectionA233_(rows, 'No Action', decisions, ['AVOID'], reportId, integrationA233);

    if (rows.length > 0) {
      output.getRange(2, 1, rows.length, 10).setValues(rows);
    }

    foArchiveExecutiveReport_(dashboard, reportId, summary);

    foInfo_(module, 'Complete', 'Executive report generated: ' + reportId);

    return {
      status: 'SUCCESS',
      reportId: reportId,
      rowsWritten: rows.length,
      averageReadiness: summary.averageReadiness,
      preferredPortfolioScenario:
        integrationA233.portfolioScenario &&
        integrationA233.portfolioScenario.available
          ? integrationA233.portfolioScenario.preferredScenario
          : 'NOT AVAILABLE'
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}


function foValidateRequiredSheets_(spreadsheet, requiredSheets, workbookName) {
  if (!spreadsheet) {
    throw new Error(
      workbookName + ' workbook is unavailable.'
    );
  }

  const missingSheets = requiredSheets.filter(function(sheetName) {
    return !spreadsheet.getSheetByName(sheetName);
  });

  if (missingSheets.length) {
    throw new Error(
      workbookName +
      ' is missing required sheet(s): ' +
      missingSheets.join(', ')
    );
  }

  return {
    status: 'SUCCESS',
    workbookName: workbookName,
    requiredSheetCount: requiredSheets.length,
    missingSheets: []
  };
}

function foValidateWorkbookTitle_(spreadsheet, expectedTitle) {
  if (!spreadsheet) {
    throw new Error(expectedTitle + ' workbook is unavailable.');
  }

  const actualTitle = String(spreadsheet.getName() || '').trim();

  if (actualTitle !== expectedTitle) {
    throw new Error(
      'Workbook title mismatch. Expected "' +
      expectedTitle +
      '" but found "' +
      actualTitle +
      '".'
    );
  }

  return {
    status: 'SUCCESS',
    expectedTitle: expectedTitle,
    actualTitle: actualTitle,
    verified: true
  };
}


function foRunMorningBriefPreflight_() {
  const dashboard = foDashboard_();
  const ledger = foLedger_();

  const dashboardTitleValidation = foValidateWorkbookTitle_(
    dashboard,
    'Family Office Portfolio Dashboard'
  );

  const ledgerTitleValidation = foValidateWorkbookTitle_(
    ledger,
    'Family Office Investment Ledger'
  );

  const dashboardValidation = foValidateRequiredSheets_(
    dashboard,
    FO_MORNING_BRIEF_REQUIRED_DASHBOARD_SHEETS,
    'Family Office Portfolio Dashboard'
  );

  const ledgerValidation = foValidateRequiredSheets_(
    ledger,
    FO_MORNING_BRIEF_REQUIRED_LEDGER_SHEETS,
    'Family Office Investment Ledger'
  );

  return {
    status: 'SUCCESS',
    dataAccessStatus: 'LIVE',
    workbookAccessStatus: 'LIVE',
    workbookTitleStatus: 'VERIFIED',
    requiredSheetsStatus: 'VERIFIED',
    governedEvidenceStatus: 'AVAILABLE',
    persistenceDependenciesStatus: 'VERIFIED',
    dashboard: dashboard,
    ledger: ledger,
    dashboardTitleValidation: dashboardTitleValidation,
    ledgerTitleValidation: ledgerTitleValidation,
    dashboardValidation: dashboardValidation,
    ledgerValidation: ledgerValidation,
    checkedAt: new Date()
  };
}


function foReadGovernedExecutiveDecisions_(dashboard, integrationA233) {
  const decisionSheet = dashboard.getSheetByName(
    FO_SHEETS.INVESTMENT_DECISION_SUPPORT
  );
  if (!decisionSheet || decisionSheet.getLastRow() < 2) return [];

  const values = decisionSheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const marketValues = foExecutiveMarketValueMap_(dashboard);
  const cards = integrationA233 && integrationA233.actionCards
    ? integrationA233.actionCards : [];

  return values.slice(1).map(function(row) {
    const ticker = String(
      foGetVal_(row, headers, 'Ticker') || ''
    ).trim().toUpperCase();
    const account = String(
      foGetVal_(row, headers, 'Account') || ''
    ).trim();
    if (!ticker) return null;

    const card = typeof foA233FindCard_ === 'function'
      ? foA233FindCard_(cards, ticker, account)
      : null;
    const risk = foNum_(foGetVal_(row, headers, 'Risk'));
    const qualityScore = foNum_(
      foGetVal_(row, headers, 'Recommendation Quality Score')
    );
    const qualityGrade = String(
      foGetVal_(row, headers, 'Recommendation Quality Grade') ||
      'NOT ASSESSED'
    ).trim().toUpperCase();
    const contradictionStatus = String(
      foGetVal_(row, headers, 'Contradiction Status') ||
      'NOT ASSESSED'
    ).trim().toUpperCase();
    const materiality = foNum_(
      foGetVal_(row, headers, 'Materiality Score')
    );
    const action = String(
      foGetVal_(row, headers, 'Action') || ''
    ).trim().toUpperCase();
    const qualityRationale = String(
      foGetVal_(row, headers, 'Quality Rationale') || ''
    ).trim();
    const executiveReason = String(
      foGetVal_(row, headers, 'Executive Reason') || ''
    ).trim();

    return {
      ticker: ticker,
      company: marketValues.companies[ticker] || '',
      account: account,
      marketValue: foExecutiveMarketValue_(marketValues, ticker, account),
      buyZoneConfidence: foNum_(
        foGetVal_(row, headers, 'Confidence')
      ),
      convictionScore: foNum_(
        foGetVal_(row, headers, 'Conviction')
      ),
      materialityScore: materiality,
      riskRating: risk > 50 ? 'HIGH' : (risk > 35 ? 'MEDIUM' : 'LOW'),
      marketRecommendation: String(
        foGetVal_(row, headers, 'Recommendation') || ''
      ).trim(),
      cioReadiness: qualityScore,
      cioAction: action,
      priority: contradictionStatus === 'BLOCKED' || materiality >= 85
        ? 'CRITICAL'
        : (qualityGrade === 'LOW' || qualityGrade === 'INSUFFICIENT DATA'
          ? 'HIGH'
          : (materiality >= 70 ? 'HIGH' : 'NORMAL')),
      deploymentGuidance: card ? card.executionStatus : '',
      requiresReview:
        contradictionStatus !== 'CLEAR' ||
        qualityGrade === 'LOW' ||
        qualityGrade === 'INSUFFICIENT DATA' ||
        (card && String(card.executionStatus).indexOf('BLOCKED') === 0)
          ? 'YES' : 'NO',
      recommendationQualityScore: qualityScore,
      recommendationQualityGrade: qualityGrade,
      evidenceBalance: String(
        foGetVal_(row, headers, 'Evidence Balance') || 'NOT ASSESSED'
      ).trim(),
      contradictionStatus: contradictionStatus,
      rationale: [qualityRationale, executiveReason]
        .filter(function(value) { return value; })
        .join(' | ')
    };
  }).filter(function(item) {
    return item !== null;
  });
}

function foExecutiveMarketValueMap_(dashboard) {
  const sheet = dashboard.getSheetByName('Portfolio Performance Positions');
  const result = {exact: {}, ticker: {}, companies: {}};
  if (!sheet || sheet.getLastRow() < 2) return result;

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  values.slice(1).forEach(function(row) {
    const ticker = String(
      foGetVal_(row, headers, 'Ticker') || ''
    ).trim().toUpperCase();
    const account = String(
      foGetVal_(row, headers, 'Account') || ''
    ).trim().toUpperCase();
    const marketValue = foNum_(
      foGetVal_(row, headers, 'Market Value')
    );
    if (!ticker) return;
    const key = ticker + '|' + account;
    result.exact[key] = (result.exact[key] || 0) + marketValue;
    result.ticker[ticker] = (result.ticker[ticker] || 0) + marketValue;
    result.companies[ticker] = String(
      foGetVal_(row, headers, 'Company') || result.companies[ticker] || ''
    ).trim();
  });
  return result;
}

function foExecutiveMarketValue_(marketValues, ticker, account) {
  const key = String(ticker || '').trim().toUpperCase() + '|' +
    String(account || '').trim().toUpperCase();
  if (Object.prototype.hasOwnProperty.call(marketValues.exact, key)) {
    return marketValues.exact[key];
  }
  return marketValues.ticker[String(ticker || '').trim().toUpperCase()] || 0;
}

function foBuildExecutiveSummary_(decisions, valuationSummary) {
  const metrics = valuationSummary || {};
  const hasGovernedMarketValue =
    Object.prototype.hasOwnProperty.call(
      metrics,
      'Valued-Position Market Value'
    ) ||
    Object.prototype.hasOwnProperty.call(
      metrics,
      'Total Market Value'
    );

  if (!hasGovernedMarketValue) {
    throw new Error(
      'Portfolio Valuation Summary does not contain a governed market-value metric.'
    );
  }

  const totalMarketValue = Number(
    metrics['Valued-Position Market Value'] !== undefined
      ? metrics['Valued-Position Market Value']
      : metrics['Total Market Value']
  );

  if (!Number.isFinite(totalMarketValue)) {
    throw new Error(
      'Governed portfolio market value is not numeric.'
    );
  }

  const readinessValues = decisions
    .map(function(d) { return Number(d.cioReadiness || 0); })
    .filter(function(v) { return v > 0; });

  const averageReadiness =
    readinessValues.length > 0
      ? Math.round(readinessValues.reduce(function(a, b) { return a + b; }, 0) / readinessValues.length)
      : 0;

  const reviewCount = decisions.filter(function(d) {
    return String(d.requiresReview || '').toUpperCase() === 'YES';
  }).length;

  const criticalCount = decisions.filter(function(d) {
    return String(d.priority || '').toUpperCase() === 'CRITICAL';
  }).length;

  const highRiskCount = decisions.filter(function(d) {
    return String(d.riskRating || '').toUpperCase() === 'HIGH';
  }).length;

  let portfolioRisk = 'Low';
  if (highRiskCount >= 3) portfolioRisk = 'High';
  else if (highRiskCount >= 1) portfolioRisk = 'Medium';

  let overallPriority = 'Normal';
  if (criticalCount > 0) overallPriority = 'Critical';
  else if (averageReadiness >= 85) overallPriority = 'High';

  let narrative = 'Portfolio remains stable. No urgent action required.';

  if (criticalCount > 0) {
    narrative = 'Critical opportunity detected. CIO review required before capital deployment.';
  } else if (averageReadiness >= 85) {
    narrative = 'Portfolio opportunity set is strong. Selective deployment is supported.';
  } else if (portfolioRisk === 'High') {
    narrative = 'Portfolio contains elevated risk exposures. Review before adding capital.';
  }

  return {
    totalMarketValue: totalMarketValue,
    averageReadiness: averageReadiness,
    reviewCount: reviewCount,
    criticalCount: criticalCount,
    highRiskCount: highRiskCount,
    portfolioRisk: portfolioRisk,
    overallPriority: overallPriority,
    executiveNarrative: narrative
  };
}

function foAppendDecisionSection_(rows, sectionName, decisions, actions, reportId) {
  const filtered = decisions.filter(function(d) {
    return actions.indexOf(String(d.cioAction || '').toUpperCase()) >= 0;
  });

  filtered.forEach(function(d) {
    rows.push([
      sectionName,
      d.ticker,
      d.cioAction,
      d.priority,
      d.riskRating,
      d.rationale || d.deploymentGuidance || '',
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);
  });
}


function foAppendPortfolioValuationExecutiveRows_(rows, metrics, reportId) {
  metrics = metrics || {};

  const valuedMarketValue = Number(
    metrics['Valued-Position Market Value'] ||
    metrics['Total Market Value'] ||
    0
  );
  const totalCostBasis = Number(metrics['Total Cost Basis'] || 0);
  const comparableCostBasis = Number(metrics['Comparable Cost Basis'] || 0);
  const comparableGainLoss = Number(metrics['Comparable Unrealized Gain/Loss'] || 0);
  const comparableGainLossPct = Number(metrics['Comparable Unrealized Gain/Loss %'] || 0);
  const fullReturnEligible = String(metrics['Full Portfolio Return Eligible'] || 'NO').toUpperCase() === 'YES';
  const fullGainLoss = fullReturnEligible ? Number(metrics['Unrealized Gain/Loss'] || 0) : 'SUPPRESSED';
  const fullGainLossPct = fullReturnEligible ? Number(metrics['Unrealized Gain/Loss %'] || 0) : 'SUPPRESSED';
  const priceCoverage = Number(metrics['Price Coverage %'] || 0);
  const costBasisCoverage = Number(metrics['Cost Basis Coverage %'] || 0);
  const missingPriceCount = Number(metrics['Missing Price Count'] || 0);
  const missingPriceTickers = String(metrics['Missing Price Tickers'] || 'NONE');
  const completeness = String(metrics['Valuation Completeness Status'] || 'UNAVAILABLE');
  const reconciliationVariance = Number(metrics['Reconciliation Variance'] || 0);
  const reconciliationStatus = String(metrics['Reconciliation Status'] || 'NOT AVAILABLE');
  const certificationStatus = String(metrics['Certification Status'] || 'NOT AVAILABLE');
  const valuationTimestamp = metrics['Valuation Timestamp'] || 'NOT AVAILABLE';
  const latestPriceTimestamp = metrics['Latest Price Timestamp'] || 'NOT AVAILABLE';
  const priceBasis = String(metrics['Price Basis'] || 'NOT AVAILABLE');
  const reconciliationPassed = reconciliationStatus.toUpperCase() === 'RECONCILED' || reconciliationStatus.toUpperCase() === 'PASS';

  function addValuationMetric_(metric, value, priority, risk, notes) {
    rows.push(['Portfolio Valuation Evidence', metric, value, priority || '', risk || '', notes || '', reportId, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE, new Date()]);
  }

  addValuationMetric_('Certification Status', certificationStatus,
    certificationStatus.toUpperCase() === 'CERTIFIED' ? 'NORMAL' : 'CRITICAL',
    certificationStatus.toUpperCase() === 'CERTIFIED' ? 'LOW' : 'HIGH',
    'Certification requires complete valuation evidence and successful reconciliation.');
  addValuationMetric_('Valuation Completeness', completeness,
    completeness === 'COMPLETE' ? 'NORMAL' : 'CRITICAL',
    completeness === 'COMPLETE' ? 'LOW' : 'HIGH',
    'Completeness is separate from reconciliation.');
  addValuationMetric_('Reconciliation Status', reconciliationStatus,
    reconciliationPassed ? 'NORMAL' : 'CRITICAL',
    reconciliationPassed && Math.abs(reconciliationVariance) <= 0.01 ? 'LOW' : 'HIGH',
    'Reconciliation variance: C$' + reconciliationVariance + '; governed tolerance is within C$0.01.');
  addValuationMetric_('Valued-Position Market Value', valuedMarketValue, '', '',
    'Market value for positions with supported current or persisted-fallback valuation evidence.');
  addValuationMetric_('Total Cost Basis', totalCostBasis,
    costBasisCoverage >= 1 ? 'NORMAL' : 'HIGH',
    costBasisCoverage >= 1 ? 'LOW' : 'MEDIUM',
    'Documented cost basis across all active positions.');
  addValuationMetric_('Comparable Cost Basis', comparableCostBasis, '', '',
    'Cost basis only for positions included in valued-position market value.');
  addValuationMetric_('Full Portfolio Unrealized Gain / Loss', fullGainLoss,
    fullReturnEligible ? '' : 'CRITICAL',
    fullReturnEligible ? '' : 'HIGH',
    fullReturnEligible ? 'Complete-portfolio variance.' : 'SUPPRESSED because price coverage is incomplete.');
  addValuationMetric_('Full Portfolio Unrealized Gain / Loss %', fullGainLossPct,
    fullReturnEligible ? '' : 'CRITICAL',
    fullReturnEligible ? '' : 'HIGH',
    fullReturnEligible ? 'Complete-portfolio return.' : 'SUPPRESSED because price coverage is incomplete.');
  addValuationMetric_('Comparable Unrealized Gain / Loss', comparableGainLoss, '', '',
    'Like-for-like variance for valued positions only.');
  addValuationMetric_('Comparable Unrealized Gain / Loss %', comparableGainLossPct, '', '',
    'Like-for-like return for valued positions only.');
  addValuationMetric_('Price Coverage %', priceCoverage,
    priceCoverage >= 1 ? 'NORMAL' : 'HIGH',
    priceCoverage >= 1 ? 'LOW' : 'MEDIUM',
    'Coverage ratio across active positions.');


  addValuationMetric_(
    'Valued Positions',
    Number(metrics['Valued Positions'] || 0),
    '',
    '',
    'Positions successfully valued using governed valuation evidence.'
  );

  addValuationMetric_(
    'Total Active Positions',
    Number(metrics['Total Active Positions'] || 0),
    '',
    '',
    'Total active portfolio positions considered during valuation.'
  );

  addValuationMetric_(
    'Missing Cost Basis Count',
    Number(metrics['Missing Cost Basis Count'] || 0),
    Number(metrics['Missing Cost Basis Count'] || 0) === 0 ? 'NORMAL' : 'HIGH',
    Number(metrics['Missing Cost Basis Count'] || 0) === 0 ? 'LOW' : 'MEDIUM',
    'Positions without documented cost basis.'
  );
  addValuationMetric_('Cost Basis Coverage %', costBasisCoverage,
    costBasisCoverage >= 1 ? 'NORMAL' : 'HIGH',
    costBasisCoverage >= 1 ? 'LOW' : 'MEDIUM',
    'Coverage ratio across active positions.');
  addValuationMetric_('Missing Price Count', missingPriceCount,
    missingPriceCount === 0 ? 'NORMAL' : 'CRITICAL',
    missingPriceCount === 0 ? 'LOW' : 'HIGH',
    'Missing-price tickers: ' + missingPriceTickers + '.');
  addValuationMetric_('Valuation Timestamp', valuationTimestamp,
    valuationTimestamp === 'NOT AVAILABLE' ? 'HIGH' : 'NORMAL',
    valuationTimestamp === 'NOT AVAILABLE' ? 'MEDIUM' : 'LOW',
    'Portfolio valuation execution timestamp.');
  addValuationMetric_('Latest Price Timestamp', latestPriceTimestamp,
    latestPriceTimestamp === 'NOT AVAILABLE' ? 'HIGH' : 'NORMAL',
    latestPriceTimestamp === 'NOT AVAILABLE' ? 'MEDIUM' : 'LOW',
    'Most recent supported price timestamp used in valuation.');
  addValuationMetric_('Price Basis', priceBasis,
    priceBasis === 'NOT AVAILABLE' ? 'HIGH' : 'NORMAL',
    priceBasis === 'NOT AVAILABLE' ? 'MEDIUM' : 'LOW',
    'Portfolio-level basis: LIVE, DELAYED, PRIOR_CLOSE, PERSISTED_FALLBACK, ESTIMATED, or MIXED.');

  ['TFSA', 'LIRA', 'IBKR'].forEach(function(account) {
    const key = account + ' Market Value';
    if (Object.prototype.hasOwnProperty.call(metrics, key)) {
      addValuationMetric_(key, Number(metrics[key] || 0), '', '', 'Account-level governed valuation evidence.');
    }
  });
  if (Object.prototype.hasOwnProperty.call(metrics, 'IBKR Cash Included')) {
    addValuationMetric_('IBKR Cash Included', String(metrics['IBKR Cash Included']), '', '',
      'YES only when a recognized cash ticker is included in valuation evidence.');
  }
}

function foAppendPortfolioOptimizationExecutiveRows_(
  dashboard,
  rows,
  reportId
) {
  const sheet = dashboard.getSheetByName(
    FO_SHEETS.PORTFOLIO_OPTIMIZATION_SUMMARY
  );

  if (!sheet || sheet.getLastRow() < 2) {
    rows.push([
      'Portfolio Optimization Intelligence',
      'Optimization Status',
      'NOT AVAILABLE',
      'REVIEW',
      'UNKNOWN',
      'Run Portfolio Optimization Intelligence before executive reporting.',
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);
    return;
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const metrics = {};

  values.slice(1).forEach(function(row) {
    const metric = String(
      foGetVal_(row, headers, 'Metric') || ''
    ).trim();

    if (metric) {
      metrics[metric] = foGetVal_(row, headers, 'Value');
    }
  });

  function addOptimizationMetric_(metric, value, notes) {
    rows.push([
      'Portfolio Optimization Intelligence',
      metric,
      value,
      '',
      '',
      notes || '',
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);
  }

  addOptimizationMetric_(
    'Portfolio Directive',
    metrics['Portfolio Directive'] || 'NOT AVAILABLE',
    'Governed deterministic weight-based optimization directive.'
  );

  addOptimizationMetric_(
    'Candidates Reviewed',
    Number(metrics['Candidate Count'] || 0),
    'Candidates assessed by Portfolio Optimization Intelligence.'
  );

  addOptimizationMetric_(
    'Eligible Candidates',
    Number(metrics['Eligible Candidate Count'] || 0),
    'Candidates eligible after governed upstream controls.'
  );

  addOptimizationMetric_(
    'Funded Candidates',
    Number(metrics['Funded Candidate Count'] || 0),
    'Candidates receiving positive optimized incremental weight.'
  );

  addOptimizationMetric_(
    'Constrained Candidates',
    Number(metrics['Constrained Candidate Count'] || 0),
    'Candidates capped or blocked by governed allocation constraints.'
  );

  addOptimizationMetric_(
    'Optimized Incremental Weight',
    Number(metrics['Optimized Incremental Weight'] || 0),
    'Aggregate optimized incremental portfolio weight.'
  );

  addOptimizationMetric_(
    'Largest Optimized Target Weight',
    Number(metrics['Largest Optimized Target Weight'] || 0),
    'Largest position weight produced by the governed optimization result.'
  );
}

function foAppendPortfolioScenarioExecutiveRows_(
  dashboard,
  rows,
  reportId,
  integrationA233
) {
  let scenario =
    integrationA233 &&
    integrationA233.portfolioScenario
      ? integrationA233.portfolioScenario
      : null;

  if (
    (!scenario || !scenario.available) &&
    typeof foReadPreferredPortfolioScenarioExecutive_ === 'function'
  ) {
    scenario = foReadPreferredPortfolioScenarioExecutive_(dashboard);
  }

  if (!scenario || !scenario.available) {
    rows.push([
      'Portfolio Scenario Intelligence',
      'Preferred Scenario',
      'NOT AVAILABLE',
      'REVIEW',
      'UNKNOWN',
      'Run Portfolio Scenario Intelligence after Portfolio Optimization.',
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);
    return;
  }

  const detail = foReadPreferredScenarioExplainability_(dashboard);

  function addScenarioMetric_(metric, value, priority, risk, notes) {
    rows.push([
      'Portfolio Scenario Intelligence',
      metric,
      value,
      priority || '',
      risk || '',
      notes || '',
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);
  }

  addScenarioMetric_(
    'Preferred Scenario',
    scenario.preferredScenario,
    scenario.scenarioScore >= 75 ? 'HIGH' : 'NORMAL',
    scenario.portfolioRiskLevel,
    scenario.rationale
  );

  addScenarioMetric_(
    'Scenario Score',
    scenario.scenarioScore,
    '',
    scenario.stressContext,
    'Deterministic comparison score; not a return forecast.'
  );

  addScenarioMetric_(
    'Proposed Incremental Weight',
    Number(scenario.totalIncrementalWeight || 0),
    '',
    '',
    'Preferred-scenario proposed incremental portfolio weight.'
  );

  addScenarioMetric_(
    'Funded Candidates',
    Number(scenario.fundedCandidateCount || 0),
    '',
    '',
    'Candidates receiving positive proposed incremental weight.'
  );

  addScenarioMetric_(
    'Largest Projected Position',
    Number(scenario.largestTargetWeight || 0),
    '',
    '',
    'Largest target weight under the preferred scenario.'
  );

  const showExplainability =
    String(scenario.portfolioRiskLevel || '').toUpperCase() ===
      'CRITICAL' ||
    Number(detail.constraintBreachCount || 0) > 0 ||
    Number(detail.constraintComplianceScore || 100) < 100;

  if (showExplainability) {
    addScenarioMetric_(
      'Deployment Alignment Score',
      detail.deploymentAlignmentScore,
      '',
      '',
      'Alignment with governed optimized deployment potential.'
    );

    addScenarioMetric_(
      'Diversification Score',
      detail.diversificationScore,
      '',
      '',
      'Concentration-aware diversification assessment.'
    );

    addScenarioMetric_(
      'Risk Discipline Score',
      detail.riskDisciplineScore,
      '',
      '',
      'Risk discipline under the preferred scenario.'
    );

    addScenarioMetric_(
      'Stress Discipline Score',
      detail.stressDisciplineScore,
      '',
      '',
      'Stress discipline under enabled stress scenarios.'
    );

    addScenarioMetric_(
      'Constraint Compliance Score',
      detail.constraintComplianceScore,
      Number(detail.constraintBreachCount || 0) > 0
        ? 'CRITICAL'
        : 'NORMAL',
      '',
      'Compliance with governed scenario constraints.'
    );

    addScenarioMetric_(
      'Scenario Constraint Breaches',
      detail.constraintBreachCount,
      Number(detail.constraintBreachCount || 0) > 0
        ? 'CRITICAL'
        : 'NORMAL',
      '',
      'Preferred-scenario constraint breach count.'
    );
  }

  addScenarioMetric_(
    'Scenario Recommendation',
    scenario.executiveRecommendation,
    scenario.portfolioRiskLevel === 'CRITICAL'
      ? 'CRITICAL'
      : 'NORMAL',
    scenario.portfolioRiskLevel,
    'Advisory scenario output; all execution controls remain authoritative.'
  );
}

function foReadPreferredScenarioExplainability_(dashboard) {
  const sheet = dashboard.getSheetByName(
    FO_SHEETS.PORTFOLIO_SCENARIO_SUMMARY
  );

  const unavailable = {
    deploymentAlignmentScore: 0,
    diversificationScore: 0,
    riskDisciplineScore: 0,
    stressDisciplineScore: 0,
    constraintComplianceScore: 0,
    constraintBreachCount: 0
  };

  if (!sheet || sheet.getLastRow() < 2) return unavailable;

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);

  const preferred = values.slice(1).filter(function(row) {
    return String(
      foGetVal_(row, headers, 'Preferred') || ''
    ).trim().toUpperCase() === 'YES';
  })[0];

  if (!preferred) return unavailable;

  return {
    deploymentAlignmentScore: Number(
      foGetVal_(
        preferred,
        headers,
        'Deployment Alignment Score'
      ) || 0
    ),
    diversificationScore: Number(
      foGetVal_(preferred, headers, 'Diversification Score') || 0
    ),
    riskDisciplineScore: Number(
      foGetVal_(preferred, headers, 'Risk Discipline Score') || 0
    ),
    stressDisciplineScore: Number(
      foGetVal_(preferred, headers, 'Stress Discipline Score') || 0
    ),
    constraintComplianceScore: Number(
      foGetVal_(
        preferred,
        headers,
        'Constraint Compliance Score'
      ) || 0
    ),
    constraintBreachCount: Number(
      foGetVal_(preferred, headers, 'Constraint Breach Count') || 0
    )
  };
}

function foAppendRiskBudgetExecutiveRows_(
  dashboard,
  rows,
  reportId
) {
  const sheet = dashboard.getSheetByName(
    FO_SHEETS.RISK_BUDGET_SUMMARY
  );

  if (!sheet || sheet.getLastRow() < 2) {
    rows.push([
      'Risk Budget Intelligence',
      'Risk Budget Status',
      'NOT AVAILABLE',
      'REVIEW',
      'UNKNOWN',
      'Run Risk Budget Intelligence after Portfolio Scenario Intelligence.',
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);
    return;
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const metrics = {};

  values.slice(1).forEach(function(row) {
    const metric = String(
      foGetVal_(row, headers, 'Metric') || ''
    ).trim();

    if (metric) {
      metrics[metric] = {
        value: foGetVal_(row, headers, 'Value'),
        status: foGetVal_(row, headers, 'Status'),
        rationale: foGetVal_(row, headers, 'Rationale')
      };
    }
  });

  function metric(name) {
    return metrics[name] || {};
  }

  function addRiskBudgetMetric_(name, displayName, priority, notes) {
    const item = metric(name);

    rows.push([
      'Risk Budget Intelligence',
      displayName,
      item.value !== undefined && item.value !== ''
        ? item.value
        : 0,
      priority || '',
      item.status || '',
      item.rationale || notes || '',
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);
  }

  const overallStatus = String(
    metric('Overall Risk Budget Status').value || 'NOT AVAILABLE'
  ).toUpperCase();

  const overallConstraintStatus = String(
    metric('Overall Constraint Status').value || 'NOT AVAILABLE'
  ).toUpperCase();

  addRiskBudgetMetric_(
    'Overall Risk Budget Status',
    'Risk Budget Status',
    overallStatus === 'BREACH' ? 'CRITICAL' : 'NORMAL'
  );

  addRiskBudgetMetric_(
    'Portfolio Risk Budget Utilization',
    'Portfolio Budget Utilization',
    ''
  );

  addRiskBudgetMetric_(
    'Remaining Risk Capacity',
    'Remaining Risk Capacity',
    ''
  );

  addRiskBudgetMetric_(
    'Risk Budget Breach Count',
    'Capacity Breaches',
    Number(metric('Risk Budget Breach Count').value || 0) > 0
      ? 'CRITICAL'
      : 'NORMAL'
  );

  addRiskBudgetMetric_(
    'Constrained Allocation Count',
    'Constrained Allocations',
    Number(metric('Constrained Allocation Count').value || 0) > 0
      ? 'HIGH'
      : 'NORMAL'
  );

  addRiskBudgetMetric_(
    'Overall Constraint Status',
    'Overall Constraint Status',
    overallConstraintStatus === 'BLOCKED'
      ? 'CRITICAL'
      : 'NORMAL'
  );

  addRiskBudgetMetric_(
    'Primary Blocker',
    'Primary Blocker',
    String(metric('Primary Blocker').value || 'NONE').toUpperCase() ===
      'NONE'
      ? 'NORMAL'
      : 'HIGH'
  );

  addRiskBudgetMetric_(
    'Blocked Position Count',
    'Blocked Positions',
    Number(metric('Blocked Position Count').value || 0) > 0
      ? 'HIGH'
      : 'NORMAL'
  );

  if (
    overallStatus === 'BREACH' ||
    overallConstraintStatus === 'BLOCKED'
  ) {
    addRiskBudgetMetric_(
      'Recommendation Control Blocks',
      'Recommendation-Control Blocks',
      ''
    );

    addRiskBudgetMetric_(
      'Confidence Blocks',
      'Confidence Blocks',
      ''
    );

    addRiskBudgetMetric_(
      'Allocation Eligibility Blocks',
      'Allocation-Eligibility Blocks',
      ''
    );

    addRiskBudgetMetric_(
      'Market Data Blocks',
      'Market-Data Blocks',
      ''
    );

    addRiskBudgetMetric_(
      'Other Upstream Blocks',
      'Other Upstream Blocks',
      ''
    );
  }

  addRiskBudgetMetric_(
    'Executive Risk Budget Directive',
    'Executive Directive',
    overallStatus === 'BREACH' ? 'CRITICAL' : ''
  );
}

function foAppendMaterialActionExplainability_(
  rows,
  integrationA233,
  reportId
) {
  const cards =
    integrationA233 &&
    integrationA233.actionCards
      ? integrationA233.actionCards.slice()
      : [];

  if (!cards.length) return;

  const materialCards = cards.filter(function(card) {
    const executionStatus = String(
      card.executionStatus || ''
    ).toUpperCase();

    return (
      executionStatus.indexOf('BLOCKED') === 0 ||
      String(card.priorityLevel || '').toUpperCase() === 'CRITICAL' ||
      String(card.priorityLevel || '').toUpperCase() === 'HIGH' ||
      Number(card.materialityScore || 0) >= 70
    );
  }).sort(function(a, b) {
    const aBlocked =
      String(a.executionStatus || '').indexOf('BLOCKED') === 0
        ? 1
        : 0;

    const bBlocked =
      String(b.executionStatus || '').indexOf('BLOCKED') === 0
        ? 1
        : 0;

    if (aBlocked !== bBlocked) return bBlocked - aBlocked;

    return (
      Number(b.executivePriorityScore || b.materialityScore || 0) -
      Number(a.executivePriorityScore || a.materialityScore || 0)
    );
  }).slice(0, 3);

  materialCards.forEach(function(card) {
    const executionStatus =
      String(card.executionStatus || 'INFORMATIONAL ONLY');

    const trigger =
      String(card.trigger || '').trim() ||
      'Refresh governed inputs before reconsideration.';

    const invalidation =
      String(card.invalidationCondition || '').trim();

    const notes = [
      'Security type: ' +
        String(card.securityType || 'NOT AVAILABLE'),
      'Execution status: ' + executionStatus,
      'Reconsider when: ' + trigger,
      invalidation
        ? 'Invalidation: ' + invalidation
        : ''
    ].filter(function(value) {
      return value;
    }).join(' | ');

    rows.push([
      'Material Action Explainability',
      String(card.ticker || 'NOT AVAILABLE'),
      String(card.action || 'REVIEW') +
        ' | ' +
        executionStatus,
      String(card.priorityLevel || ''),
      String(card.riskImpact || ''),
      notes,
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);
  });
}

function foArchiveExecutiveReport_(dashboard, reportId, summary) {
  const archive = foEnsureSheet_(dashboard, 'Executive Report Archive', [
    'Timestamp',
    'Report ID',
    'Average CIO Readiness',
    'Total Market Value',
    'Portfolio Risk',
    'Overall Priority',
    'Review Count',
    'Narrative',
    'Platform Version',
    'Baseline'
  ]);

  archive.appendRow([
    new Date(),
    reportId,
    summary.averageReadiness,
    summary.totalMarketValue,
    summary.portfolioRisk,
    summary.overallPriority,
    summary.reviewCount,
    summary.executiveNarrative,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE
  ]);
}

function foRunExecutiveReportSmokeTest() {
  const module = 'ExecutiveReportingEngine';

  try {
    foInfo_(module, 'Start', 'Executive Report smoke test started.');

    const result = foRunExecutiveReportEngine();

    foInfo_(module, 'Complete', 'Executive Report smoke test completed: ' + result.reportId);

    return result;

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}
