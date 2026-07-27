/************************************************************
 * RiskBudgetIntelligence.js
 * Sprint 3.2.0 — Risk Budget Intelligence
 *
 * Deterministic governance layer. It does not forecast returns,
 * volatility, drawdowns, losses, or trade quantities. Existing upstream
 * optimization, scenario, risk, exposure, and constraint outputs remain
 * authoritative.
 ************************************************************/

const FO_RISK_BUDGET_ENGINE_VERSION = '1.0.0';

function foSetupRiskBudgetIntelligence() {
  const dashboard = foDashboard_();
  foEnsureRiskBudgetSheetContract_(
    dashboard,
    FO_SHEETS.RISK_BUDGET_ASSESSMENT,
    foRiskBudgetAssessmentHeaders_()
  );
  foEnsureRiskBudgetSheetContract_(
    dashboard,
    FO_SHEETS.RISK_BUDGET_SUMMARY,
    foRiskBudgetSummaryHeaders_()
  );
  return {status: 'SUCCESS', engineVersion: FO_RISK_BUDGET_ENGINE_VERSION};
}

function foRunRiskBudgetIntelligence() {
  const module = 'RiskBudgetIntelligence';
  try {
    foInfo_(module, 'Start', 'Risk Budget Intelligence started.');
    foSetupRiskBudgetIntelligence();
    const dashboard = foDashboard_();
    const preferred = foReadPreferredScenarioRiskBudgetRows_(dashboard);
    const result = foEvaluateRiskBudget_(preferred.allocations, preferred.summary);
    foWriteRiskBudgetAssessment_(dashboard, result.assessments);
    foWriteRiskBudgetSummary_(dashboard, result.summary);
    foInfo_(module, 'Complete', 'Risk Budget Intelligence completed: ' + result.summary.overallStatus);
    return {
      status: 'SUCCESS',
      engineVersion: FO_RISK_BUDGET_ENGINE_VERSION,
      overallStatus: result.summary.overallStatus,
      utilization: result.summary.utilization,
      remainingCapacity: result.summary.remainingCapacity,
      breachCount: result.summary.breachCount,
      constrainedCount: result.summary.constrainedCount,
      directive: result.summary.directive
    };
  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foReadPreferredScenarioRiskBudgetRows_(dashboard) {
  const detail = dashboard.getSheetByName(FO_SHEETS.PORTFOLIO_SCENARIOS);
  const summarySheet = dashboard.getSheetByName(FO_SHEETS.PORTFOLIO_SCENARIO_SUMMARY);
  if (!detail || detail.getLastRow() < 2 || !summarySheet || summarySheet.getLastRow() < 2) {
    throw new Error('Portfolio Scenario Intelligence contains no governed preferred scenario.');
  }
  const detailValues = detail.getDataRange().getValues();
  const headers = detailValues[0].map(String);
  const allocations = detailValues.slice(1).filter(function(row) {
    return foRiskBudgetBoolean_(foRiskBudgetVal_(row, headers, 'Preferred'));
  }).map(function(row) {
    return {
      scenarioRank: foRiskBudgetNumber_(foRiskBudgetVal_(row, headers, 'Scenario Rank'), 999999),
      scenarioId: String(foRiskBudgetVal_(row, headers, 'Scenario ID') || '').trim(),
      scenarioName: String(foRiskBudgetVal_(row, headers, 'Scenario Name') || '').trim(),
      ticker: String(foRiskBudgetVal_(row, headers, 'Ticker') || '').trim().toUpperCase(),
      account: String(foRiskBudgetVal_(row, headers, 'Account') || '').trim(),
      deploymentDecision: String(foRiskBudgetVal_(row, headers, 'Deployment Decision') || '').trim(),
      currentWeight: foRiskBudgetWeight_(foRiskBudgetVal_(row, headers, 'Current Portfolio Weight')),
      proposedIncrementalWeight: foRiskBudgetWeight_(foRiskBudgetVal_(row, headers, 'Proposed Incremental Weight')),
      proposedTargetWeight: foRiskBudgetWeight_(foRiskBudgetVal_(row, headers, 'Proposed Target Weight')),
      maximumPositionWeight: foRiskBudgetWeight_(foRiskBudgetVal_(row, headers, 'Maximum Position Weight')),
      upstreamConstraintStatus: String(foRiskBudgetVal_(row, headers, 'Upstream Constraint Status') || '').trim().toUpperCase(),
      upstreamConstraintReason: String(foRiskBudgetVal_(row, headers, 'Upstream Constraint Reason') || '').trim(),
      riskDisciplineScore: foRiskBudgetScore_(foRiskBudgetVal_(row, headers, 'Risk Discipline Score'), 0),
      constraintComplianceScore: foRiskBudgetScore_(foRiskBudgetVal_(row, headers, 'Constraint Compliance Score'), 0)
    };
  }).filter(function(item) { return item.ticker; });

  const summaryValues = summarySheet.getDataRange().getValues();
  const summaryHeaders = summaryValues[0].map(String);
  const preferredSummary = summaryValues.slice(1).find(function(row) {
    return foRiskBudgetBoolean_(foRiskBudgetVal_(row, summaryHeaders, 'Preferred'));
  });
  if (!preferredSummary || !allocations.length) {
    throw new Error('Preferred portfolio scenario is missing or has no allocations.');
  }
  return {
    allocations: allocations,
    summary: {
      scenarioRank: foRiskBudgetNumber_(foRiskBudgetVal_(preferredSummary, summaryHeaders, 'Scenario Rank'), 999999),
      scenarioId: String(foRiskBudgetVal_(preferredSummary, summaryHeaders, 'Scenario ID') || '').trim(),
      scenarioName: String(foRiskBudgetVal_(preferredSummary, summaryHeaders, 'Scenario Name') || '').trim(),
      portfolioRiskLevel: String(foRiskBudgetVal_(preferredSummary, summaryHeaders, 'Portfolio Risk Level') || 'UNKNOWN').trim().toUpperCase(),
      riskDisciplineScore: foRiskBudgetScore_(foRiskBudgetVal_(preferredSummary, summaryHeaders, 'Risk Discipline Score'), 0),
      constraintComplianceScore: foRiskBudgetScore_(foRiskBudgetVal_(preferredSummary, summaryHeaders, 'Constraint Compliance Score'), 0),
      upstreamBreachCount: foRiskBudgetNumber_(foRiskBudgetVal_(preferredSummary, summaryHeaders, 'Constraint Breach Count'), 0)
    }
  };
}

function foEvaluateRiskBudget_(allocations, scenarioSummary) {
  const safe = allocations || [];
  if (!safe.length) throw new Error('Risk Budget Intelligence requires at least one preferred-scenario allocation.');
  const assessments = safe.map(function(item) {
    const capacity = Math.max(0, item.maximumPositionWeight || 0);
    const proposed = Math.max(0, item.proposedTargetWeight || 0);
    const utilization = capacity > 0 ? proposed / capacity : (proposed > 0 ? 1 : 0);
    const upstreamFailed = item.upstreamConstraintStatus && item.upstreamConstraintStatus !== 'PASS';
    const overBudget = capacity <= 0 ? proposed > 0 : proposed > capacity + 0.0000001;
    let status = 'WITHIN BUDGET';
    let reason = 'Proposed target remains within the governed maximum position weight.';
    if (upstreamFailed || overBudget) {
      status = 'BREACH';
      reason = upstreamFailed
        ? (item.upstreamConstraintReason || 'Upstream constraint is not PASS.')
        : 'Proposed target weight exceeds the governed maximum position weight.';
    } else if (utilization >= 0.90) {
      status = 'CONSTRAINED';
      reason = 'Proposed target uses at least 90% of the governed position capacity.';
    }
    return Object.assign({}, item, {
      riskBudgetCapacity: foRiskBudgetRoundWeight_(capacity),
      riskBudgetUtilization: foRiskBudgetRoundWeight_(utilization),
      remainingRiskCapacity: foRiskBudgetRoundWeight_(Math.max(0, capacity - proposed)),
      budgetStatus: status,
      breachReason: reason,
      executiveDirective: status === 'BREACH'
        ? 'DO NOT INCREASE; REMEDIATE CONSTRAINT'
        : (status === 'CONSTRAINED' ? 'LIMIT ADDITIONAL DEPLOYMENT' : 'WITHIN GOVERNED CAPACITY')
    });
  });
  const totalCapacity = assessments.reduce(function(sum, item) { return sum + item.riskBudgetCapacity; }, 0);
  const totalProposed = assessments.reduce(function(sum, item) { return sum + item.proposedTargetWeight; }, 0);
  const breachCount = assessments.filter(function(item) { return item.budgetStatus === 'BREACH'; }).length;
  const constrainedCount = assessments.filter(function(item) { return item.budgetStatus === 'CONSTRAINED'; }).length;
  const utilization = totalCapacity > 0 ? totalProposed / totalCapacity : 0;
  let overallStatus = 'WITHIN BUDGET';
  if (breachCount > 0 || (scenarioSummary && scenarioSummary.upstreamBreachCount > 0)) overallStatus = 'BREACH';
  else if (constrainedCount > 0 || utilization >= 0.90) overallStatus = 'CONSTRAINED';
  const directive = overallStatus === 'BREACH'
    ? 'PAUSE INCREMENTAL DEPLOYMENT AND REMEDIATE RISK-BUDGET BREACHES.'
    : (overallStatus === 'CONSTRAINED'
      ? 'DEPLOY SELECTIVELY; PRESERVE REMAINING RISK CAPACITY.'
      : 'PROPOSED ALLOCATION REMAINS WITHIN GOVERNED RISK CAPACITY.');
  return {
    assessments: assessments,
    summary: {
      overallStatus: overallStatus,
      totalCapacity: foRiskBudgetRoundWeight_(totalCapacity),
      totalProposed: foRiskBudgetRoundWeight_(totalProposed),
      utilization: foRiskBudgetRoundWeight_(utilization),
      remainingCapacity: foRiskBudgetRoundWeight_(Math.max(0, totalCapacity - totalProposed)),
      breachCount: breachCount,
      constrainedCount: constrainedCount,
      scenarioRiskLevel: scenarioSummary ? scenarioSummary.portfolioRiskLevel : 'UNKNOWN',
      riskDisciplineScore: scenarioSummary ? scenarioSummary.riskDisciplineScore : 0,
      constraintComplianceScore: scenarioSummary ? scenarioSummary.constraintComplianceScore : 0,
      directive: directive,
      rationale: 'Risk budget uses preferred-scenario target weights and existing governed maximum-position and upstream-constraint outputs; it is not predictive.'
    }
  };
}

function foWriteRiskBudgetAssessment_(dashboard, assessments) {
  const headers = foRiskBudgetAssessmentHeaders_();
  const sheet = foEnsureRiskBudgetSheetContract_(dashboard, FO_SHEETS.RISK_BUDGET_ASSESSMENT, headers);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
  const now = new Date();
  const rows = assessments.map(function(item) { return [
    item.scenarioRank, item.scenarioId, item.scenarioName, item.ticker, item.account,
    item.deploymentDecision, item.currentWeight, item.proposedIncrementalWeight,
    item.proposedTargetWeight, item.riskBudgetCapacity, item.riskBudgetUtilization,
    item.remainingRiskCapacity, item.budgetStatus, item.breachReason,
    item.upstreamConstraintStatus, item.upstreamConstraintReason,
    item.riskDisciplineScore, item.constraintComplianceScore,
    item.executiveDirective, now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE
  ]; });
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  foFormatRiskBudgetSheet_(sheet, headers, rows.length, [
    'Current Portfolio Weight','Proposed Incremental Weight','Proposed Target Weight',
    'Risk Budget Capacity','Risk Budget Utilization','Remaining Risk Capacity'
  ]);
}

function foWriteRiskBudgetSummary_(dashboard, summary) {
  const headers = foRiskBudgetSummaryHeaders_();
  const sheet = foEnsureRiskBudgetSheetContract_(dashboard, FO_SHEETS.RISK_BUDGET_SUMMARY, headers);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
  const now = new Date();
  const rows = [
    ['Overall Risk Budget Status', summary.overallStatus, summary.overallStatus, summary.rationale, now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    ['Portfolio Risk Budget Capacity', summary.totalCapacity, summary.overallStatus, 'Aggregate governed maximum-position capacity for preferred-scenario candidates.', now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    ['Proposed Target Weight', summary.totalProposed, summary.overallStatus, 'Aggregate preferred-scenario proposed target weight.', now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    ['Portfolio Risk Budget Utilization', summary.utilization, summary.overallStatus, 'Proposed target weight divided by governed position-capacity total.', now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    ['Remaining Risk Capacity', summary.remainingCapacity, summary.overallStatus, 'Non-negative residual governed capacity.', now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    ['Risk Budget Breach Count', summary.breachCount, summary.overallStatus, 'Allocations breaching upstream constraints or maximum-position capacity.', now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    ['Constrained Allocation Count', summary.constrainedCount, summary.overallStatus, 'Allocations using at least 90% of governed position capacity.', now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    ['Portfolio Risk Level', summary.scenarioRiskLevel, summary.overallStatus, 'Reused from the preferred Portfolio Scenario summary.', now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    ['Risk Discipline Score', summary.riskDisciplineScore, summary.overallStatus, 'Reused from the preferred Portfolio Scenario summary.', now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    ['Constraint Compliance Score', summary.constraintComplianceScore, summary.overallStatus, 'Reused from the preferred Portfolio Scenario summary.', now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    ['Executive Risk Budget Directive', summary.directive, summary.overallStatus, summary.rationale, now, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE]
  ];
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  foFormatRiskBudgetSheet_(sheet, headers, rows.length, []);
  const valueColumn = headers.indexOf('Value') + 1;
  sheet.getRange(3, valueColumn, 4, 1).setNumberFormat('0.00%');
}

function foEnsureRiskBudgetSheetContract_(dashboard, sheetName, headers) {
  let sheet = dashboard.getSheetByName(sheetName);
  if (!sheet) {
    sheet = dashboard.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }
  if (sheet.getLastColumn() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }
  const actual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function(value) { return String(value || '').trim(); });
  const matches = actual.length === headers.length && headers.every(function(header, index) { return actual[index] === header; });
  if (!matches) throw new Error('Risk Budget contract mismatch for "' + sheetName + '". Expected: ' + JSON.stringify(headers) + ' Actual: ' + JSON.stringify(actual));
  return sheet;
}

function foFormatRiskBudgetSheet_(sheet, headers, rowCount, percentHeaders) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  percentHeaders.forEach(function(header) {
    const column = headers.indexOf(header) + 1;
    if (column > 0) sheet.getRange(2, column, Math.max(rowCount, 1), 1).setNumberFormat('0.00%');
  });
  sheet.autoResizeColumns(1, headers.length);
  ['Breach Reason','Upstream Constraint Reason','Executive Directive','Rationale'].forEach(function(header) {
    const column = headers.indexOf(header) + 1;
    if (column > 0) sheet.setColumnWidth(column, 520);
  });
}

function foRiskBudgetAssessmentHeaders_() { return [
  'Scenario Rank','Scenario ID','Scenario Name','Ticker','Account','Deployment Decision',
  'Current Portfolio Weight','Proposed Incremental Weight','Proposed Target Weight',
  'Risk Budget Capacity','Risk Budget Utilization','Remaining Risk Capacity',
  'Budget Status','Breach Reason','Upstream Constraint Status','Upstream Constraint Reason',
  'Risk Discipline Score','Constraint Compliance Score','Executive Directive',
  'Timestamp','Platform Version','Baseline'
]; }
function foRiskBudgetSummaryHeaders_() { return ['Metric','Value','Status','Rationale','Timestamp','Platform Version','Baseline']; }
function foRiskBudgetVal_(row, headers, name) { const index = headers.indexOf(name); return index >= 0 ? row[index] : ''; }
function foRiskBudgetNumber_(value, fallback) { const number = Number(value); return isFinite(number) ? number : fallback; }
function foRiskBudgetScore_(value, fallback) { return Math.max(0, Math.min(100, foRiskBudgetNumber_(value, fallback))); }
function foRiskBudgetWeight_(value) { const number = foRiskBudgetNumber_(value, 0); return number > 1 ? number / 100 : Math.max(0, number); }
function foRiskBudgetBoolean_(value) { return value === true || String(value || '').trim().toUpperCase() === 'TRUE' || String(value || '').trim().toUpperCase() === 'YES'; }
function foRiskBudgetRoundWeight_(value) { return Math.round((Number(value) || 0) * 1000000) / 1000000; }

function foRunRiskBudgetIntelligenceSmokeTest() {
  const setup = foSetupRiskBudgetIntelligence();
  const result = foRunRiskBudgetIntelligence();
  if (setup.status !== 'SUCCESS' || result.status !== 'SUCCESS') throw new Error('Risk Budget Intelligence smoke test failed: ' + JSON.stringify(result));
  return result;
}
