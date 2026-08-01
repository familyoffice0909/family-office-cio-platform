/************************************************************
 * ExecutiveDashboardEngine.gs
 * Wave 1C.8 — Executive Dashboard Engine
 ************************************************************/

function foRunExecutiveDashboardEngine() {
  const module = 'ExecutiveDashboardEngine';

  try {
    foInfo_(module, 'Start', 'Executive Dashboard Engine started.');

    const dashboard = foDashboard_();

    const portfolioSummary = foReadMetricSheet_(
      dashboard,
      'Portfolio Engine Summary',
      'Metric',
      'Value'
    );

    const valuationSummary = foReadMetricSheet_(
      dashboard,
      'Portfolio Valuation Summary',
      'Metric',
      'Value'
    );

    const executiveReportSummary = foReadExecutiveReportSummary_(dashboard);
    const cioDecisions = foReadGovernedDashboardDecisionRows_(dashboard);

    const dashboardSheet = foEnsureSheet_(dashboard, 'Executive Dashboard', [
      'Metric',
      'Value',
      'Status',
      'Notes'
    ]);

    dashboardSheet.clear();

    const rows = foBuildExecutiveDashboardRows_(
      portfolioSummary,
      valuationSummary,
      executiveReportSummary,
      cioDecisions
    );

    dashboardSheet.getRange(1, 1, rows.length, 4).setValues(rows);

    foFormatExecutiveDashboard_(dashboardSheet);

    foInfo_(module, 'Complete', 'Executive Dashboard refreshed.');

    return {
      status: 'SUCCESS',
      rowsWritten: rows.length
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foReadMetricSheet_(spreadsheet, sheetName, metricHeader, valueHeader) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  const map = {};

  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return map;

  const headers = values[0].map(String);
  const metricIndex = headers.indexOf(metricHeader);
  const valueIndex = headers.indexOf(valueHeader);

  if (metricIndex < 0 || valueIndex < 0) return map;

  for (let r = 1; r < values.length; r++) {
    const metric = String(values[r][metricIndex] || '').trim();
    if (!metric) continue;

    map[metric] = values[r][valueIndex];
  }

  return map;
}

function foReadExecutiveReportSummary_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('Executive CIO Report');
  const summary = {};

  if (!sheet) return summary;

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return summary;

  const headers = values[0].map(String);

  const sectionIndex = headers.indexOf('Section');
  const metricIndex = headers.indexOf('Metric / Ticker');
  const valueIndex = headers.indexOf('Value / Action');
  const priorityIndex = headers.indexOf('Priority');
  const riskIndex = headers.indexOf('Risk');
  const notesIndex = headers.indexOf('Notes');

  for (let r = 1; r < values.length; r++) {
    const section = String(values[r][sectionIndex] || '');
    const metric = String(values[r][metricIndex] || '');

    if (section !== 'Executive Summary') continue;

    summary[metric] = {
      value: values[r][valueIndex],
      priority: priorityIndex >= 0 ? values[r][priorityIndex] : '',
      risk: riskIndex >= 0 ? values[r][riskIndex] : '',
      notes: notesIndex >= 0 ? values[r][notesIndex] : ''
    };
  }

  return summary;
}

function foReadGovernedDashboardDecisionRows_(spreadsheet) {
  const decisionSheet = spreadsheet.getSheetByName(
    FO_SHEETS.INVESTMENT_DECISION_SUPPORT
  );

  if (!decisionSheet || decisionSheet.getLastRow() < 2) {
    return [];
  }

  const holdingValues = foExecutiveDashboardMarketValueMap_(spreadsheet);
  const actionCards = foExecutiveDashboardActionCardMap_(spreadsheet);

  const values = decisionSheet.getDataRange().getValues();
  const headers = values[0].map(String);

  return values.slice(1).map(function(row) {
    const ticker = String(
      foGetVal_(row, headers, 'Ticker') || ''
    ).trim().toUpperCase();

    const account = String(
      foGetVal_(row, headers, 'Account') || ''
    ).trim();

    if (!ticker) return null;

    const key = ticker + '|' + account.toUpperCase();
    const card = actionCards[key] || actionCards[ticker + '|'] || {};

    const risk = foNum_(foGetVal_(row, headers, 'Risk'));
    const qualityScore = foNum_(
      foGetVal_(row, headers, 'Recommendation Quality Score')
    );
    const materialityScore = foNum_(
      foGetVal_(row, headers, 'Materiality Score')
    );

    const contradictionStatus = String(
      card.contradictionStatus ||
      foGetVal_(row, headers, 'Contradiction Status') ||
      'NOT ASSESSED'
    ).trim().toUpperCase();

    const qualityGrade = String(
      card.recommendationQualityGrade ||
      foGetVal_(row, headers, 'Recommendation Quality Grade') ||
      'NOT ASSESSED'
    ).trim().toUpperCase();

    const executionStatus = String(
      card.executionStatus || ''
    ).trim().toUpperCase();

    return {
      ticker: ticker,
      company: foGetVal_(row, headers, 'Company') || '',
      account: account,
      marketValue:
        Object.prototype.hasOwnProperty.call(holdingValues.exact, key)
          ? holdingValues.exact[key]
          : (holdingValues.ticker[ticker] || 0),
      cioReadiness: qualityScore,
      cioAction: foGetVal_(row, headers, 'Action'),
      priority:
        contradictionStatus === 'BLOCKED' || materialityScore >= 85
          ? 'CRITICAL'
          : (
            qualityGrade === 'LOW' ||
            qualityGrade === 'INSUFFICIENT DATA' ||
            materialityScore >= 70
              ? 'HIGH'
              : 'NORMAL'
          ),
      riskRating: risk > 50 ? 'HIGH' : (risk > 35 ? 'MEDIUM' : 'LOW'),
      requiresReview:
        contradictionStatus !== 'CLEAR' ||
        qualityGrade === 'LOW' ||
        qualityGrade === 'INSUFFICIENT DATA' ||
        executionStatus.indexOf('BLOCKED') === 0
          ? 'YES'
          : 'NO',
      buyZoneConfidence: foNum_(
        foGetVal_(row, headers, 'Confidence')
      ),
      convictionScore: foNum_(
        foGetVal_(row, headers, 'Conviction')
      ),
      materialityScore: materialityScore,
      recommendation: foGetVal_(row, headers, 'Recommendation'),
      executionStatus: executionStatus,
      securityType: card.securityType || ''
    };
  }).filter(function(item) {
    return item !== null;
  });
}

function foExecutiveDashboardActionCardMap_(spreadsheet) {
  const result = {};
  const sheet = spreadsheet.getSheetByName(
    FO_SHEETS.REPORT_ACTION_CARDS_A233
  );

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

    if (!ticker) return;

    result[ticker + '|' + account] = {
      executionStatus: foGetVal_(row, headers, 'Execution Status'),
      securityType: foGetVal_(row, headers, 'Security Type'),
      recommendationQualityGrade: foGetVal_(
        row,
        headers,
        'Recommendation Quality Grade'
      ),
      contradictionStatus: foGetVal_(
        row,
        headers,
        'Contradiction Status'
      )
    };
  });

  return result;
}

function foExecutiveDashboardMarketValueMap_(spreadsheet) {
  const result = {exact: {}, ticker: {}};
  const sheet = spreadsheet.getSheetByName(
    'Portfolio Performance Positions'
  );

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

    if (!ticker) return;

    const marketValue = foNum_(
      foGetVal_(row, headers, 'Market Value')
    );

    const key = ticker + '|' + account;

    result.exact[key] = (result.exact[key] || 0) + marketValue;
    result.ticker[ticker] =
      (result.ticker[ticker] || 0) + marketValue;
  });

  return result;
}

function foBuildExecutiveDashboardRows_(portfolioSummary, valuationSummary, executiveReportSummary, cioDecisions) {
  const totalPositions = Number(
    portfolioSummary['Total Positions'] ||
    cioDecisions.length ||
    0
  );

  const hasGovernedMarketValue =
    Object.prototype.hasOwnProperty.call(
      valuationSummary,
      'Valued-Position Market Value'
    ) ||
    Object.prototype.hasOwnProperty.call(
      valuationSummary,
      'Total Market Value'
    );

  const valuationTotalMarketValue = hasGovernedMarketValue
    ? Number(
      valuationSummary['Valued-Position Market Value'] !== undefined
        ? valuationSummary['Valued-Position Market Value']
        : valuationSummary['Total Market Value']
    )
    : 'NOT AVAILABLE';

  const totalCostBasis = Number(
    valuationSummary['Total Cost Basis'] ||
    valuationSummary['Total Documented Book Value'] ||
    0
  );

  const fullReturnEligible = String(
    valuationSummary['Full Portfolio Return Eligible'] || 'NO'
  ).trim().toUpperCase() === 'YES';

  const unrealizedGainLoss = fullReturnEligible
    ? Number(
      valuationSummary['Unrealized Gain/Loss'] ||
      valuationSummary['Unrealized Variance'] ||
      0
    )
    : 'SUPPRESSED';

  const unrealizedGainLossPct = fullReturnEligible
    ? Number(
      valuationSummary['Unrealized Gain/Loss %'] ||
      valuationSummary['Unrealized Variance %'] ||
      0
    )
    : 'SUPPRESSED';

  const priceCoverage = Number(
    valuationSummary['Price Coverage %'] ||
    valuationSummary['Market Price Coverage %'] ||
    0
  );

  const costBasisCoverage = Number(
    valuationSummary['Cost Basis Coverage %'] ||
    valuationSummary['Book Value Coverage %'] ||
    0
  );

  const missingPriceCount = Number(
    valuationSummary['Missing Price Count'] ||
    valuationSummary['Missing Prices'] ||
    0
  );

  const valuedPositions = Number(
    valuationSummary['Valued Positions'] || 0
  );

  const totalActivePositions = Number(
    valuationSummary['Total Active Positions'] || 0
  );

  const missingCostBasisCount = Number(
    valuationSummary['Missing Cost Basis Count'] || 0
  );


  const reconciliationVariance = Number(
    valuationSummary['Reconciliation Variance'] ||
    valuationSummary['Portfolio Reconciliation Variance'] ||
    0
  );

  const reconciliationStatus = String(
    valuationSummary['Reconciliation Status'] || 'NOT AVAILABLE'
  );

  const certificationStatus = String(
    valuationSummary['Certification Status'] ||
    valuationSummary['Portfolio Certification'] ||
    'NOT AVAILABLE'
  );

  const valuationTimestamp =
    valuationSummary['Valuation Timestamp'] ||
    valuationSummary['Holdings Timestamp'] ||
    '';

  const priceBasis = String(
    valuationSummary['Price Basis'] ||
    valuationSummary['Market Price Basis'] ||
    'NOT AVAILABLE'
  );

  const readinessMetric = executiveReportSummary['Overall CIO Readiness'] || {};
  const reviewMetric = executiveReportSummary['Actions Requiring Review'] || {};

  const avgReadiness =
    readinessMetric.value !== undefined && readinessMetric.value !== ''
      ? readinessMetric.value
      : foAverage_(cioDecisions.map(function(d) { return Number(d.cioReadiness || 0); }));

  const criticalActions = cioDecisions.filter(function(d) {
    return String(d.priority || '').toUpperCase() === 'CRITICAL';
  }).length;

  const highPriorityActions = cioDecisions.filter(function(d) {
    return String(d.priority || '').toUpperCase() === 'HIGH';
  }).length;

  const reviewCount = cioDecisions.filter(function(d) {
    return String(d.requiresReview || '').toUpperCase() === 'YES';
  }).length;

  const deployCapital = cioDecisions.filter(function(d) {
    return String(d.cioAction || '').toUpperCase().indexOf('DEPLOY') >= 0;
  });

  const buyAdd = cioDecisions.filter(function(d) {
    return String(d.cioAction || '').toUpperCase() === 'BUY / ADD';
  });

  const watchReview = cioDecisions.filter(function(d) {
    return String(d.cioAction || '').toUpperCase().indexOf('WATCH') >= 0;
  });

  const topOpportunity = foFindTopOpportunity_(cioDecisions);
  const largestPosition = foFindLargestPosition_(cioDecisions);
  const highestRisk = foFindHighestRisk_(cioDecisions);

  const rows = [];

  rows.push(['Metric', 'Value', 'Status', 'Notes']);
  rows.push(['Portfolio Health', 'Operational', '🟢', 'Dashboard generated by Executive Dashboard Engine.']);
  rows.push(['Platform Version', FO_CONFIG.PLATFORM_VERSION, '🟢', FO_CONFIG.BASELINE]);
  rows.push(['Last Refresh', new Date(), '🟢', 'Automated dashboard refresh timestamp.']);
  rows.push(['', '', '', '']);

  rows.push(['1. Executive Summary', '', '', '']);
  rows.push([
    'Total Portfolio Value',
    valuationTotalMarketValue,
    certificationStatus.toUpperCase() === 'CERTIFIED' ? '🟢' : '🟡',
    hasGovernedMarketValue
      ? 'Governed Portfolio Valuation Summary — valued positions only.'
      : 'Governed valuation evidence is unavailable.'
  ]);
  rows.push(['Total Positions', totalPositions, '🟢', 'Active holdings included in Portfolio Snapshot.']);
  rows.push(['Overall CIO Readiness', avgReadiness, foReadinessStatus_(avgReadiness), readinessMetric.notes || 'Average readiness across CIO decisions.']);
  rows.push(['Actions Requiring Review', reviewMetric.value || reviewCount, reviewCount > 0 ? '🟡' : '🟢', 'Items requiring manual CIO review.']);
  rows.push(['Critical Actions', criticalActions, criticalActions > 0 ? '🔴' : '🟢', 'Critical-priority CIO decisions.']);
  rows.push(['High Priority Actions', highPriorityActions, highPriorityActions > 0 ? '🟡' : '🟢', 'High-priority CIO actions.']);
  rows.push(['', '', '', '']);

  rows.push(['2. Portfolio Valuation Evidence','','','']);

  rows.push(['Certification Status',certificationStatus,
    certificationStatus.toUpperCase()==='CERTIFIED'?'🟢':'🔴',
    'Governed certification status.']);

  rows.push(['Reconciliation Status',reconciliationStatus,
    reconciliationStatus.toUpperCase()==='RECONCILED'?'🟢':'🔴',
    'Portfolio reconciliation control.']);

  rows.push(['Reconciliation Variance',
    reconciliationVariance,
    Math.abs(reconciliationVariance)<=0.01?'🟢':'🔴',
    'Expected tolerance ≤ C$0.01.']);

  rows.push(['Total Cost Basis',
    totalCostBasis,
    costBasisCoverage >= 1 ? '🟢' : '🟡',
    'Documented cost basis.']);

  rows.push(['Unrealized Gain / Loss',
    unrealizedGainLoss,
    fullReturnEligible ? '' : '🟡',
    fullReturnEligible
      ? 'Complete-portfolio market value minus cost basis.'
      : 'SUPPRESSED because full-portfolio return is not eligible.']);

  rows.push(['Unrealized Gain / Loss %',
    unrealizedGainLossPct,
    fullReturnEligible ? '' : '🟡',
    fullReturnEligible
      ? 'Complete-portfolio percentage variance versus cost basis.'
      : 'SUPPRESSED because full-portfolio return is not eligible.']);

  rows.push(['Valuation Price Coverage %',
    priceCoverage,
    priceCoverage >= 1 ? '🟢' : '🟡',
    'Supported valuation evidence across active positions; distinct from decision price freshness.']);

  rows.push(['Valued Positions',
    valuedPositions,
    valuedPositions === totalActivePositions ? '🟢' : '🟡',
    'Positions with supported market valuation.']);

  rows.push(['Total Active Positions',
    totalActivePositions,
    '🟢',
    'Active holdings included in valuation.']);

  rows.push(['Cost Basis Coverage %',
    costBasisCoverage,
    costBasisCoverage >= 1 ? '🟢' : '🟡',
    'Documented valuation cost-basis coverage across active positions; distinct from return-attribution coverage.']);

  rows.push(['Missing Cost Basis Count',
    missingCostBasisCount,
    missingCostBasisCount===0 ? '🟢' : '🔴',
    'Active holdings without documented cost basis.']);

  rows.push(['Missing Price Count',
    missingPriceCount,
    missingPriceCount===0?'🟢':'🔴',
    'Active holdings without supported pricing.']);

  rows.push(['Valuation Timestamp',
    valuationTimestamp||'NOT AVAILABLE',
    valuationTimestamp?'🟢':'🟡',
    'Persisted valuation evidence timestamp.']);

  rows.push(['Price Basis',
    priceBasis,
    priceBasis==='NOT AVAILABLE'?'🟡':'🟢',
    'Live, delayed, prior close, persisted fallback or estimated.']);

  rows.push(['','','','']);

  rows.push(['3. Capital Deployment', '', '', '']);
  rows.push(['Deploy Capital Count', deployCapital.length, deployCapital.length > 0 ? '🔴' : '🟢', foJoinTickers_(deployCapital)]);
  rows.push(['Buy / Add Count', buyAdd.length, buyAdd.length > 0 ? '🟡' : '🟢', foJoinTickers_(buyAdd)]);
  rows.push(['Watch / Review Count', watchReview.length, watchReview.length > 0 ? '🟡' : '🟢', foJoinTickers_(watchReview)]);
  rows.push(['Top Opportunity', topOpportunity.ticker || 'N/A', topOpportunity.status || '', topOpportunity.notes || '']);
  rows.push(['', '', '', '']);

  rows.push(['4. Risk Monitoring', '', '', '']);
  rows.push(['Largest Position', largestPosition.ticker || 'N/A', '', largestPosition.notes || '']);
  rows.push(['Highest Risk Holding', highestRisk.ticker || 'N/A', highestRisk.status || '', highestRisk.notes || '']);
  rows.push(['High Risk Holdings Count', foCountRisk_(cioDecisions, 'HIGH'), foCountRisk_(cioDecisions, 'HIGH') > 0 ? '🟡' : '🟢', 'Count of holdings rated High risk.']);
  rows.push(['', '', '', '']);

  rows.push(['5. Market Intelligence', '', '', '']);
  rows.push(['Average Buy Zone Confidence', foAverage_(cioDecisions.map(function(x) { return Number(x.buyZoneConfidence || 0); })), '', 'Average across scored holdings.']);
  rows.push(['Average Conviction Score', foAverage_(cioDecisions.map(function(x) { return Number(x.convictionScore || 0); })), '', 'Average across scored holdings.']);
  rows.push(['Average Materiality Score', foAverage_(cioDecisions.map(function(x) { return Number(x.materialityScore || 0); })), '', 'Average across scored holdings.']);

  return rows;
}

function foAverage_(values) {
  const clean = values.filter(function(v) {
    return !isNaN(v) && Number(v) > 0;
  });

  if (clean.length === 0) return 0;

  return Math.round(clean.reduce(function(a, b) {
    return a + b;
  }, 0) / clean.length);
}

function foReadinessStatus_(readiness) {
  const score = Number(readiness || 0);

  if (score >= 85) return '🟢';
  if (score >= 70) return '🟡';
  return '🔴';
}

function foJoinTickers_(rows) {
  if (!rows || rows.length === 0) return 'None';

  const counts = {};

  rows.forEach(function(r) {
    const ticker = String(r.ticker || r.Ticker || '').trim();
    counts[ticker] = (counts[ticker] || 0) + 1;
  });

  return rows.map(function(r) {
    const ticker = String(r.ticker || r.Ticker || '').trim();

    if (counts[ticker] === 1) {
      return ticker;
    }

    const account = foExecutiveDashboardAccountLabel_(r);
    return ticker + ' (' + account + ')';
  }).join(', ');
}

function foExecutiveDashboardAccountLabel_(row) {
  const account =
    row.account ||
    row.Account ||
    row.accountName ||
    row.account_name ||
    row.accountType ||
    row.account_type ||
    row.portfolioAccount ||
    row.portfolio_account ||
    row.sourceAccount ||
    row.source_account ||
    '';

  const normalized = String(account).trim();

  if (normalized) {
    return normalized;
  }

  return 'Account not provided';
}

function foFindTopOpportunity_(decisions) {
  if (!decisions || decisions.length === 0) {
    return { ticker: '', status: '', notes: '' };
  }

  const sorted = decisions.slice().sort(function(a, b) {
    return Number(b.cioReadiness || 0) - Number(a.cioReadiness || 0);
  });

  const top = sorted[0];

  return {
    ticker: top.ticker,
    status: String(top.priority || '').toUpperCase() === 'CRITICAL' ? '🔴' : '🟢',
    notes: 'Readiness: ' + top.cioReadiness + ' | Action: ' + top.cioAction
  };
}

function foFindLargestPosition_(decisions) {
  if (!decisions || decisions.length === 0) {
    return { ticker: '', notes: '' };
  }

  const sorted = decisions.slice().sort(function(a, b) {
    return Number(b.marketValue || 0) - Number(a.marketValue || 0);
  });

  const top = sorted[0];

  return {
    ticker: top.ticker,
    notes: 'Market Value: ' + top.marketValue
  };
}

function foFindHighestRisk_(decisions) {
  const highRisk = decisions.filter(function(d) {
    return String(d.riskRating || '').toUpperCase() === 'HIGH';
  });

  if (highRisk.length === 0) {
    return { ticker: 'None', status: '🟢', notes: 'No high-risk holdings detected.' };
  }

  const sorted = highRisk.slice().sort(function(a, b) {
    return Number(b.cioReadiness || 0) - Number(a.cioReadiness || 0);
  });

  return {
    ticker: sorted[0].ticker,
    status: '🟡',
    notes: 'High-risk holding with readiness ' + sorted[0].cioReadiness
  };
}

function foCountRisk_(decisions, riskLevel) {
  return decisions.filter(function(d) {
    return String(d.riskRating || '').toUpperCase() === String(riskLevel || '').toUpperCase();
  }).length;
}

function foFormatExecutiveDashboard_(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange('A1:D1').setFontWeight('bold');
  sheet.autoResizeColumns(1, 4);

  const values = sheet.getDataRange().getValues();

  for (let r = 1; r <= values.length; r++) {
    const label = String(values[r - 1][0] || '');

    if (label.indexOf('.') > 0 || label === 'Metric') {
      sheet.getRange(r, 1, 1, 4).setFontWeight('bold');
    }
  }
}

function foRunExecutiveDashboardSmokeTest() {
  const module = 'ExecutiveDashboardEngine';

  try {
    foInfo_(module, 'Start', 'Executive Dashboard smoke test started.');

    const result = foRunExecutiveDashboardEngine();

    foInfo_(module, 'Complete', 'Executive Dashboard smoke test completed.');

    return result;

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}