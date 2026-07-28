
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
  'Portfolio Scenario Summary',
  'Risk Budget Summary',
  'Investment Decision Support',
  'Executive Decision State A233',
  'Automation Log',
  'Executive Report Archive'
];

const FO_MORNING_BRIEF_REQUIRED_LEDGER_SHEETS = [
  'Version History',
  'Knowledge Base',
  'Canadian Market Access Library',
  'Outcomes',
  'Lessons Learned',
  'Orchestration Log'
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
    const summary = foBuildExecutiveSummary_(decisions);

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
      'Based on Portfolio Performance Positions market value data.',
      reportId,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE,
      new Date()
    ]);

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

function foRunMorningBriefPreflight_() {
  const dashboard = foDashboard_();
  const ledger = foLedger_();

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
    dashboard: dashboard,
    ledger: ledger,
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

function foBuildExecutiveSummary_(decisions) {
  const totalMarketValue = decisions.reduce(function(sum, d) {
    return sum + (Number(d.marketValue) || 0);
  }, 0);

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


function foAppendPortfolioOptimizationExecutiveRows_(dashboard, rows, reportId) {
  const sheet = dashboard.getSheetByName(
    FO_SHEETS.PORTFOLIO_OPTIMIZATION_SUMMARY
  );
  if (!sheet || sheet.getLastRow() < 2) return;

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const metrics = {};
  values.slice(1).forEach(function(row) {
    const metric = String(foGetVal_(row, headers, 'Metric') || '').trim();
    if (metric) metrics[metric] = foGetVal_(row, headers, 'Value');
  });

  rows.push([
    'Portfolio Optimization Summary',
    'Portfolio Directive',
    metrics['Portfolio Directive'] || 'NOT AVAILABLE',
    '',
    '',
    'Weight-based optimization; no cash or trade execution assumed.',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);

  rows.push([
    'Portfolio Optimization Summary',
    'Funded / Eligible Candidates',
    String(metrics['Funded Candidate Count'] || 0) + ' / ' +
      String(metrics['Eligible Candidate Count'] || 0),
    '',
    '',
    'Candidates receiving a positive incremental target weight.',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);

  rows.push([
    'Portfolio Optimization Summary',
    'Optimized Incremental Weight',
    metrics['Optimized Incremental Weight'] || 0,
    '',
    '',
    'Aggregate recommended incremental portfolio weight across eligible candidates.',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);
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

  rows.push([
    'Portfolio Scenario Intelligence',
    'Preferred Scenario',
    scenario.preferredScenario,
    scenario.scenarioScore >= 75 ? 'HIGH' : 'NORMAL',
    scenario.portfolioRiskLevel,
    scenario.rationale,
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);

  rows.push([
    'Portfolio Scenario Intelligence',
    'Scenario Score',
    scenario.scenarioScore,
    '',
    scenario.stressContext,
    'Deterministic comparison score; not a return forecast.',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);

  rows.push([
    'Portfolio Scenario Intelligence',
    'Deployment / Funded Candidates',
    String(
      Math.round(
        (Number(scenario.totalIncrementalWeight) || 0) *
        10000
      ) / 100
    ) + '% / ' +
      String(scenario.fundedCandidateCount || 0),
    '',
    '',
    'Proposed incremental portfolio weight and funded candidate count.',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);

  rows.push([
    'Portfolio Scenario Intelligence',
    'Scenario Recommendation',
    scenario.executiveRecommendation,
    scenario.portfolioRiskLevel === 'CRITICAL'
      ? 'CRITICAL'
      : 'NORMAL',
    scenario.portfolioRiskLevel,
    'Preferred scenario is advisory and remains subject to all execution controls.',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);
}


function foAppendRiskBudgetExecutiveRows_(dashboard, rows, reportId) {
  const sheet = dashboard.getSheetByName(FO_SHEETS.RISK_BUDGET_SUMMARY);
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
    const metric = String(foGetVal_(row, headers, 'Metric') || '').trim();
    if (metric) {
      metrics[metric] = {
        value: foGetVal_(row, headers, 'Value'),
        status: foGetVal_(row, headers, 'Status'),
        rationale: foGetVal_(row, headers, 'Rationale')
      };
    }
  });

  const overall = metrics['Overall Risk Budget Status'] || {};
  const utilization = metrics['Portfolio Risk Budget Utilization'] || {};
  const breaches = metrics['Risk Budget Breach Count'] || {};
  const blocker = metrics['Primary Blocker'] || {};
  const blockedPositions = metrics['Blocked Position Count'] || {};
  const directive = metrics['Executive Risk Budget Directive'] || {};

  rows.push([
    'Risk Budget Intelligence',
    'Risk Budget Status',
    overall.value || 'NOT AVAILABLE',
    String(overall.value || '').toUpperCase() === 'BREACH' ? 'CRITICAL' : 'NORMAL',
    overall.status || '',
    overall.rationale || '',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);

  rows.push([
    'Risk Budget Intelligence',
    'Portfolio Budget Utilization',
    utilization.value || 0,
    '',
    utilization.status || '',
    utilization.rationale || 'Proposed target weight divided by governed position-capacity total.',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);

  rows.push([
    'Risk Budget Intelligence',
    'Breach Count',
    breaches.value || 0,
    Number(breaches.value || 0) > 0 ? 'CRITICAL' : 'NORMAL',
    breaches.status || '',
    breaches.rationale || '',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);


  rows.push([
    'Risk Budget Intelligence',
    'Primary Blocker',
    blocker.value || 'NONE',
    String(blocker.value || 'NONE').toUpperCase() === 'NONE' ? 'NORMAL' : 'HIGH',
    blocker.status || '',
    blocker.rationale || 'Highest-priority blocker under the governed deterministic hierarchy.',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);

  rows.push([
    'Risk Budget Intelligence',
    'Blocked Positions',
    blockedPositions.value || 0,
    Number(blockedPositions.value || 0) > 0 ? 'HIGH' : 'NORMAL',
    blockedPositions.status || '',
    blockedPositions.rationale || '',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);

  rows.push([
    'Risk Budget Intelligence',
    'Executive Directive',
    directive.value || 'REVIEW RISK BUDGET',
    '',
    directive.status || '',
    directive.rationale || '',
    reportId,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.BASELINE,
    new Date()
  ]);
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
