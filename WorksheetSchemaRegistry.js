/**
 * Worksheet Schema Registry
 * Wave 2.6.0-R2 — Production Certification
 */

function foProductionSchemaRegistry_() {
  return [
    {
      sheetName: FO_SHEETS.CAPITAL_DEPLOYMENT_HISTORY,
      schemaVersion: '1.1',
      headers: [
        'Timestamp',
        'Run ID',
        'Portfolio Directive',
        'Top Ticker',
        'Top Account',
        'Top Decision',
        'Top Deployment Score',
        'Deployable Candidates',
        'Blocked Candidates',
        'Portfolio Materiality Score',
        'Platform Version',
        'Baseline',
        'State Signature'
      ]
    },
    {
      sheetName: FO_SHEETS.CAPITAL_DEPLOYMENT_PRIORITIES,
      schemaVersion: '1.1',
      headers: [
        'Rank',
        'Ticker',
        'Account',
        'Deployment Decision',
        'Deployment Score',
        'Recommendation',
        'Action',
        'Allocation Band',
        'Trend',
        'Conviction',
        'Risk',
        'Confidence',
        'Materiality Score',
        'Base Priority Score',
        'Executive Priority Score',
        'Priority Level',
        'Overall Trajectory',
        'Reversal Status',
        'Trend Evidence Strength',
        'Price Freshness',
        'Zone Position',
        'Distance to Entry %',
        'Current Price',
        'Target Entry Price',
        'Blocked',
        'Blockers',
        'Executive Reason',
        'Portfolio Directive',
        'Portfolio Materiality Score',
        'Timestamp',
        'Platform Version',
        'Baseline',
        'Recommendation Quality Score',
        'Recommendation Quality Grade',
        'Evidence Balance',
        'Contradiction Status',
        'Quality Rationale'
      ]
    },
    {
      sheetName: FO_SHEETS.PORTFOLIO_SCENARIOS,
      schemaVersion: '1.0',
      headers: [
        'Scenario Rank',
        'Scenario ID',
        'Scenario Name',
        'Scenario Type',
        'Preferred',
        'Ticker',
        'Account',
        'Deployment Decision',
        'Deployment Score',
        'Current Portfolio Weight',
        'Optimized Incremental Weight',
        'Proposed Incremental Weight',
        'Proposed Target Weight',
        'Maximum Position Weight',
        'Upstream Constraint Status',
        'Upstream Constraint Reason',
        'Scenario Score',
        'Deployment Alignment Score',
        'Diversification Score',
        'Risk Discipline Score',
        'Stress Discipline Score',
        'Constraint Compliance Score',
        'Executive Recommendation',
        'Scenario Rationale',
        'Timestamp',
        'Platform Version',
        'Baseline'
      ]
    },
    {
      sheetName: FO_SHEETS.PORTFOLIO_SCENARIO_SUMMARY,
      schemaVersion: '1.0',
      headers: [
        'Scenario Rank',
        'Scenario ID',
        'Scenario Name',
        'Scenario Type',
        'Preferred',
        'Scenario Score',
        'Total Incremental Weight',
        'Funded Candidate Count',
        'Largest Target Weight',
        'Deployment Alignment Score',
        'Diversification Score',
        'Risk Discipline Score',
        'Stress Discipline Score',
        'Constraint Compliance Score',
        'Constraint Breach Count',
        'Portfolio Risk Level',
        'Stress Context',
        'Executive Recommendation',
        'Scenario Rationale',
        'Timestamp',
        'Platform Version',
        'Baseline'
      ]
    },
    {
      sheetName: FO_SHEETS.PORTFOLIO_MATERIALITY,
      schemaVersion: '1.1',
      headers: [
        'Timestamp',
        'Portfolio Materiality Score',
        'Portfolio Materiality Level',
        'Primary Portfolio Driver',
        'Top Contributing Securities',
        'Material Securities',
        'Significant Securities',
        'Critical Securities',
        'Improving Securities',
        'Deteriorating Securities',
        'Stale or Missing Prices',
        'Priority Review Securities',
        'Suppressed Securities',
        'Downward Reversals',
        'Opportunity Signals',
        'Top Executive Priority',
        'Primary Attention Type',
        'Recommended CIO Response',
        'Executive Summary',
        'Platform Version',
        'Baseline'
      ]
    },
    {
      sheetName: FO_SHEETS.INVESTMENT_DECISION_SUPPORT,
      schemaVersion: '2.5.3',
      requiredHeaders: [
        'Ticker',
        'Account',
        'Action',
        'Recommendation',
        'Materiality Score',
        'Priority Score',
        'Trend',
        'Conviction',
        'Risk',
        'Confidence',
        'Current Price',
        'Target Entry Price',
        'Platform Version',
        'Baseline'
      ]
    },
    {
      sheetName: FO_SHEETS.BUY_ZONE_INTELLIGENCE,
      schemaVersion: '2.5.3',
      requiredHeaders: [
        'Ticker',
        'Account',
        'Current Price',
        'Price Freshness',
        'Target Entry Price',
        'Zone Position',
        'Conviction Score',
        'Risk Score',
        'Buy Zone Confidence',
        'Recommendation',
        'Platform Version',
        'Baseline'
      ]
    }
  ];
}

function foValidateRegisteredSchemas_(dashboard) {
  return foProductionSchemaRegistry_().map(function(definition) {
    const sheet = dashboard.getSheetByName(definition.sheetName);

    if (!sheet) {
      return {
        sheetName: definition.sheetName,
        schemaVersion: definition.schemaVersion,
        status: 'FAIL',
        issues: ['Worksheet missing']
      };
    }

    if (sheet.getLastColumn() < 1) {
      return {
        sheetName: definition.sheetName,
        schemaVersion: definition.schemaVersion,
        status: 'FAIL',
        issues: ['Worksheet has no columns']
      };
    }

    const actualHeaders = sheet.getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    ).getValues()[0].map(function(value) {
      return String(value || '').trim();
    });

    const issues = [];

    if (definition.headers) {
      definition.headers.forEach(function(expected, index) {
        if (actualHeaders[index] !== expected) {
          issues.push(
            'Column ' + (index + 1) +
            ' expected "' + expected +
            '" but found "' + (actualHeaders[index] || '') + '"'
          );
        }
      });
    }

    if (definition.requiredHeaders) {
      definition.requiredHeaders.forEach(function(expected) {
        if (actualHeaders.indexOf(expected) === -1) {
          issues.push('Required header missing: ' + expected);
        }
      });
    }

    return {
      sheetName: definition.sheetName,
      schemaVersion: definition.schemaVersion,
      status: issues.length ? 'FAIL' : 'PASS',
      issues: issues
    };
  });
}
