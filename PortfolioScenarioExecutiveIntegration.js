/**
 * PortfolioScenarioExecutiveIntegration.js
 * Sprint 3.1.0 — Portfolio Scenario Executive Decision Integration
 *
 * Additive adapter that enriches the governed A2.3.3 executive decision result
 * without changing the frozen A2.3.3 worksheet contracts.
 */

function foApplyPortfolioScenarioExecutiveIntegration_(
  integrationA233,
  dashboard
) {
  const base = integrationA233 || {};
  const scenario = foReadPreferredPortfolioScenarioExecutive_(dashboard);
  const enriched = {};

  Object.keys(base).forEach(function(key) {
    enriched[key] = base[key];
  });
  enriched.portfolioScenario = scenario;

  return enriched;
}

function foReadPreferredPortfolioScenarioExecutive_(dashboard) {
  const sheet = dashboard.getSheetByName(
    FO_SHEETS.PORTFOLIO_SCENARIO_SUMMARY
  );

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      available: false,
      preferredScenarioId: '',
      preferredScenario: 'NOT AVAILABLE',
      scenarioType: '',
      scenarioScore: 0,
      totalIncrementalWeight: 0,
      fundedCandidateCount: 0,
      largestTargetWeight: 0,
      portfolioRiskLevel: 'UNKNOWN',
      stressContext: 'NOT AVAILABLE',
      executiveRecommendation:
        'RUN PORTFOLIO SCENARIO INTELLIGENCE',
      rationale:
        'Portfolio Scenario Summary has no governed comparison results.'
    };
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const rows = values.slice(1).map(function(row) {
    return {
      rank: foScenarioExecutiveNumber_(
        foScenarioExecutiveVal_(row, headers, 'Scenario Rank'),
        999999
      ),
      preferred: String(
        foScenarioExecutiveVal_(row, headers, 'Preferred') || ''
      ).trim().toUpperCase(),
      scenarioId: String(
        foScenarioExecutiveVal_(row, headers, 'Scenario ID') || ''
      ).trim(),
      scenarioName: String(
        foScenarioExecutiveVal_(row, headers, 'Scenario Name') || ''
      ).trim(),
      scenarioType: String(
        foScenarioExecutiveVal_(row, headers, 'Scenario Type') || ''
      ).trim(),
      scenarioScore: foScenarioExecutiveNumber_(
        foScenarioExecutiveVal_(row, headers, 'Scenario Score'),
        0
      ),
      totalIncrementalWeight:
        foScenarioExecutiveNormalizeWeight_(
          foScenarioExecutiveVal_(
            row,
            headers,
            'Total Incremental Weight'
          )
        ) || 0,
      fundedCandidateCount: foScenarioExecutiveNumber_(
        foScenarioExecutiveVal_(
          row,
          headers,
          'Funded Candidate Count'
        ),
        0
      ),
      largestTargetWeight:
        foScenarioExecutiveNormalizeWeight_(
          foScenarioExecutiveVal_(
            row,
            headers,
            'Largest Target Weight'
          )
        ) || 0,
      portfolioRiskLevel: String(
        foScenarioExecutiveVal_(
          row,
          headers,
          'Portfolio Risk Level'
        ) || 'UNKNOWN'
      ).trim(),
      stressContext: String(
        foScenarioExecutiveVal_(row, headers, 'Stress Context') ||
        'NOT AVAILABLE'
      ).trim(),
      executiveRecommendation: String(
        foScenarioExecutiveVal_(
          row,
          headers,
          'Executive Recommendation'
        ) || 'NOT AVAILABLE'
      ).trim(),
      rationale: String(
        foScenarioExecutiveVal_(
          row,
          headers,
          'Scenario Rationale'
        ) || ''
      ).trim()
    };
  }).sort(function(a, b) {
    const aPreferred = a.preferred === 'YES' ? 0 : 1;
    const bPreferred = b.preferred === 'YES' ? 0 : 1;
    if (aPreferred !== bPreferred) return aPreferred - bPreferred;
    return a.rank - b.rank;
  });

  const preferred = rows[0];
  if (!preferred || !preferred.scenarioName) {
    return {
      available: false,
      preferredScenarioId: '',
      preferredScenario: 'NOT AVAILABLE',
      scenarioType: '',
      scenarioScore: 0,
      totalIncrementalWeight: 0,
      fundedCandidateCount: 0,
      largestTargetWeight: 0,
      portfolioRiskLevel: 'UNKNOWN',
      stressContext: 'NOT AVAILABLE',
      executiveRecommendation:
        'RUN PORTFOLIO SCENARIO INTELLIGENCE',
      rationale:
        'Portfolio Scenario Summary does not contain a valid preferred scenario.'
    };
  }

  return {
    available: true,
    preferredScenarioId: preferred.scenarioId,
    preferredScenario: preferred.scenarioName,
    scenarioType: preferred.scenarioType,
    scenarioScore: preferred.scenarioScore,
    totalIncrementalWeight: preferred.totalIncrementalWeight,
    fundedCandidateCount: preferred.fundedCandidateCount,
    largestTargetWeight: preferred.largestTargetWeight,
    portfolioRiskLevel: preferred.portfolioRiskLevel,
    stressContext: preferred.stressContext,
    executiveRecommendation: preferred.executiveRecommendation,
    rationale: preferred.rationale
  };
}

function foScenarioExecutiveVal_(row, headers, name) {
  const index = headers.indexOf(name);
  return index >= 0 ? row[index] : '';
}

function foScenarioExecutiveNumber_(value, fallback) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return fallback;
  }
  const number = Number(String(value).replace(/[$,%\s]/g, ''));
  return Number.isFinite(number) ? number : fallback;
}

function foScenarioExecutiveNormalizeWeight_(value) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return null;
  }
  const text = String(value).trim();
  const percentMarked = text.indexOf('%') >= 0;
  const number = Number(text.replace(/[,%\s]/g, ''));
  if (!Number.isFinite(number)) return null;
  if (percentMarked || Math.abs(number) > 1) return number / 100;
  return number;
}
