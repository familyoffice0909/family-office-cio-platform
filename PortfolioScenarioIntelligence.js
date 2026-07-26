/**
 * PortfolioScenarioIntelligence.js
 * Sprint 3.1.0 — Portfolio Scenario Intelligence
 *
 * Compares deterministic allocation alternatives before capital deployment.
 * This module is not predictive and does not estimate investment returns.
 * Existing optimization, risk, materiality, confidence and execution controls
 * remain authoritative.
 */

const FO_PORTFOLIO_SCENARIO_ENGINE_VERSION = '1.0.0';

function foSetupPortfolioScenarioIntelligence() {
  const dashboard = foDashboard_();
  const scenarioSheet = foEnsurePortfolioScenarioContract_(
    dashboard,
    FO_SHEETS.PORTFOLIO_SCENARIOS,
    foPortfolioScenarioHeaders_()
  );
  const summarySheet = foEnsurePortfolioScenarioContract_(
    dashboard,
    FO_SHEETS.PORTFOLIO_SCENARIO_SUMMARY,
    foPortfolioScenarioSummaryHeaders_()
  );

  return {
    status: 'SUCCESS',
    worksheetsEnsured: 2,
    scenarioWorksheet: scenarioSheet.getName(),
    summaryWorksheet: summarySheet.getName(),
    engineVersion: FO_PORTFOLIO_SCENARIO_ENGINE_VERSION,
    platformVersion: FO_CONFIG.PLATFORM_VERSION,
    baseline: FO_CONFIG.BASELINE
  };
}

function foRunPortfolioScenarioIntelligence() {
  const module = 'PortfolioScenarioIntelligence';

  try {
    foInfo_(module, 'Start', 'Portfolio scenario comparison started.');

    const dashboard = foDashboard_();
    foSetupPortfolioScenarioIntelligence();

    const candidates = foReadPortfolioScenarioCandidates_(dashboard);
    const risk = foReadPortfolioScenarioRiskContext_(dashboard);
    const stress = foReadPortfolioScenarioStressContext_(dashboard);
    const comparison = foBuildPortfolioScenarios_(candidates, {
      risk: risk,
      stress: stress
    });

    foWritePortfolioScenarioIntelligence_(dashboard, comparison);

    const preferred = comparison.scenarios[0] || null;
    foInfo_(
      module,
      'Complete',
      preferred
        ? 'Portfolio scenarios completed. Preferred: ' +
          preferred.scenarioName + ' (' + preferred.scenarioScore + ').'
        : 'Portfolio scenarios completed with no comparison result.'
    );

    return {
      status: preferred ? 'SUCCESS' : 'NO_DATA',
      scenarioCount: comparison.scenarios.length,
      preferredScenarioId: preferred ? preferred.scenarioId : '',
      preferredScenario: preferred ? preferred.scenarioName : '',
      preferredScenarioScore: preferred ? preferred.scenarioScore : 0,
      recommendedIncrementalWeight:
        preferred ? preferred.totalIncrementalWeight : 0,
      fundedCandidateCount: preferred ? preferred.fundedCandidateCount : 0,
      executiveRecommendation:
        preferred ? preferred.executiveRecommendation : 'NO SCENARIO AVAILABLE',
      engineVersion: FO_PORTFOLIO_SCENARIO_ENGINE_VERSION
    };
  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foEnsurePortfolioScenarioContract_(dashboard, sheetName, headers) {
  let sheet = dashboard.getSheetByName(sheetName);
  if (!sheet) {
    sheet = dashboard.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const actual = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(function(value) { return String(value || '').trim(); });
  const matches =
    actual.length === headers.length &&
    headers.every(function(header, index) {
      return actual[index] === header;
    });

  if (!matches) {
    throw new Error(
      'Portfolio Scenario contract mismatch for "' + sheetName +
      '". Expected: ' + JSON.stringify(headers) +
      ' Actual: ' + JSON.stringify(actual)
    );
  }

  return sheet;
}

function foReadPortfolioScenarioCandidates_(dashboard) {
  const sheet = dashboard.getSheetByName(FO_SHEETS.PORTFOLIO_OPTIMIZATION);
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error(
      'Portfolio Optimization contains no results. Run portfolio optimization first.'
    );
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);

  return values.slice(1).map(function(row) {
    return {
      rank: foScenarioNumber_(
        foScenarioVal_(row, headers, 'Rank'),
        999999
      ),
      ticker: String(
        foScenarioVal_(row, headers, 'Ticker') || ''
      ).trim().toUpperCase(),
      account: String(
        foScenarioVal_(row, headers, 'Account') || ''
      ).trim(),
      deploymentDecision: String(
        foScenarioVal_(row, headers, 'Deployment Decision') || ''
      ).trim().toUpperCase(),
      deploymentScore: foScenarioNumber_(
        foScenarioVal_(row, headers, 'Deployment Score'),
        0
      ),
      currentPortfolioWeight: foScenarioNormalizeWeight_(
        foScenarioVal_(row, headers, 'Current Portfolio Weight')
      ) || 0,
      optimizedIncrementalWeight: foScenarioNormalizeWeight_(
        foScenarioVal_(row, headers, 'Optimized Incremental Weight')
      ) || 0,
      optimizedTargetWeight: foScenarioNormalizeWeight_(
        foScenarioVal_(row, headers, 'Optimized Target Weight')
      ) || 0,
      maximumPositionWeight: foScenarioNormalizeWeight_(
        foScenarioVal_(row, headers, 'Maximum Position Weight')
      ) || 1,
      constraintStatus: String(
        foScenarioVal_(row, headers, 'Constraint Status') || ''
      ).trim().toUpperCase(),
      constraintReason: String(
        foScenarioVal_(row, headers, 'Constraint Reason') || ''
      ).trim()
    };
  }).filter(function(item) {
    return item.ticker;
  }).sort(function(a, b) {
    if (a.rank !== b.rank) return a.rank - b.rank;
    if (b.deploymentScore !== a.deploymentScore) {
      return b.deploymentScore - a.deploymentScore;
    }
    const aKey = a.ticker + '|' + a.account;
    const bKey = b.ticker + '|' + b.account;
    return aKey < bKey ? -1 : (aKey > bKey ? 1 : 0);
  });
}

function foReadPortfolioScenarioRiskContext_(dashboard) {
  const sheet = dashboard.getSheetByName(FO_SHEETS.PORTFOLIO_RISK);
  if (!sheet || sheet.getLastRow() < 2) {
    return {
      riskScore: 50,
      diversificationScore: 50,
      largestPositionWeight: 0,
      topFiveWeight: 0,
      stressTestScore: 50,
      overallRisk: 'UNKNOWN',
      critical: false,
      sourceAvailable: false
    };
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const row = values[values.length - 1];

  const riskScore = foScenarioScore_(
    foScenarioVal_(row, headers, 'Risk Score'),
    50
  );
  const overallRisk = String(
    foScenarioVal_(row, headers, 'Overall Risk') || 'UNKNOWN'
  ).trim().toUpperCase();

  return {
    riskScore: riskScore,
    diversificationScore: foScenarioScore_(
      foScenarioVal_(row, headers, 'Diversification Score'),
      50
    ),
    largestPositionWeight: foScenarioNormalizeWeight_(
      foScenarioVal_(row, headers, 'Largest Position %')
    ) || 0,
    topFiveWeight: foScenarioNormalizeWeight_(
      foScenarioVal_(row, headers, 'Top 5 %')
    ) || 0,
    stressTestScore: foScenarioScore_(
      foScenarioVal_(row, headers, 'Stress Test Score'),
      50
    ),
    overallRisk: overallRisk,
    critical:
      overallRisk === 'CRITICAL' ||
      riskScore >= 80,
    sourceAvailable: true
  };
}

function foReadPortfolioScenarioStressContext_(dashboard) {
  const sheet = dashboard.getSheetByName(FO_SHEETS.STRESS_SCENARIOS);
  if (!sheet || sheet.getLastRow() < 2) {
    return {
      enabledScenarioCount: 0,
      stressPressureScore: 50,
      highestSeverity: 'UNKNOWN',
      scenarioNames: [],
      sourceAvailable: false
    };
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const enabled = values.slice(1).map(function(row) {
    const enabledValue = foScenarioVal_(row, headers, 'Enabled');
    const isEnabled =
      enabledValue === true ||
      String(enabledValue || '').trim().toUpperCase() === 'TRUE';

    return {
      scenario: String(
        foScenarioVal_(row, headers, 'Scenario') || ''
      ).trim(),
      severity: String(
        foScenarioVal_(row, headers, 'Severity') || 'UNKNOWN'
      ).trim().toUpperCase(),
      shock: foScenarioNumber_(
        foScenarioVal_(row, headers, 'Shock %'),
        0
      ),
      enabled: isEnabled
    };
  }).filter(function(item) {
    return item.enabled;
  });

  if (!enabled.length) {
    return {
      enabledScenarioCount: 0,
      stressPressureScore: 0,
      highestSeverity: 'NONE',
      scenarioNames: [],
      sourceAvailable: true
    };
  }

  const severityMap = {
    LOW: 20,
    MEDIUM: 40,
    HIGH: 70,
    CRITICAL: 90,
    EXTREME: 100,
    UNKNOWN: 50
  };
  const severityRank = {
    NONE: 0,
    LOW: 1,
    MEDIUM: 2,
    HIGH: 3,
    CRITICAL: 4,
    EXTREME: 5,
    UNKNOWN: 1
  };

  const pressure = enabled.reduce(function(sum, item) {
    const severityScore = severityMap[item.severity] || severityMap.UNKNOWN;
    const shockScore = Math.min(Math.abs(item.shock), 35) / 35 * 100;
    return sum + severityScore * 0.65 + shockScore * 0.35;
  }, 0) / enabled.length;

  const highestSeverity = enabled.reduce(function(highest, item) {
    return (severityRank[item.severity] || 0) >
      (severityRank[highest] || 0)
      ? item.severity
      : highest;
  }, 'NONE');

  return {
    enabledScenarioCount: enabled.length,
    stressPressureScore: foScenarioRoundScore_(pressure),
    highestSeverity: highestSeverity,
    scenarioNames: enabled.map(function(item) { return item.scenario; }),
    sourceAvailable: true
  };
}

function foBuildPortfolioScenarios_(candidates, context) {
  const safeCandidates = (candidates || []).map(function(candidate) {
    return Object.assign({}, candidate, {
      currentPortfolioWeight:
        foScenarioRoundWeight_(candidate.currentPortfolioWeight),
      optimizedIncrementalWeight:
        Math.max(0, foScenarioRoundWeight_(
          candidate.optimizedIncrementalWeight
        )),
      optimizedTargetWeight:
        foScenarioRoundWeight_(candidate.optimizedTargetWeight),
      maximumPositionWeight:
        foScenarioRoundWeight_(candidate.maximumPositionWeight || 1)
    });
  });
  const risk = context && context.risk ? context.risk : {
    riskScore: 50,
    diversificationScore: 50,
    largestPositionWeight: 0,
    stressTestScore: 50,
    overallRisk: 'UNKNOWN',
    critical: false
  };
  const stress = context && context.stress ? context.stress : {
    enabledScenarioCount: 0,
    stressPressureScore: 50,
    highestSeverity: 'UNKNOWN',
    scenarioNames: []
  };

  const definitions = [
    {
      scenarioId: 'SCN-CURRENT',
      scenarioName: 'Maintain Current Allocation',
      scenarioType: 'BASELINE',
      mode: 'CURRENT'
    },
    {
      scenarioId: 'SCN-OPTIMIZED',
      scenarioName: 'Optimized Selective Deployment',
      scenarioType: 'OPTIMIZATION',
      mode: 'OPTIMIZED'
    },
    {
      scenarioId: 'SCN-DIVERSIFIED',
      scenarioName: 'Diversified Deployment',
      scenarioType: 'DIVERSIFICATION',
      mode: 'DIVERSIFIED'
    },
    {
      scenarioId: 'SCN-FOCUSED',
      scenarioName: 'Focused High-Priority Deployment',
      scenarioType: 'PRIORITY',
      mode: 'FOCUSED'
    },
    {
      scenarioId: 'SCN-DEFENSIVE',
      scenarioName: 'Defensive Risk-Controlled Allocation',
      scenarioType: 'RISK CONTROL',
      mode: 'DEFENSIVE'
    }
  ];

  const scenarios = definitions.map(function(definition) {
    const allocations = foBuildScenarioAllocations_(
      definition.mode,
      safeCandidates,
      {risk: risk, stress: stress}
    );
    return foScorePortfolioScenario_(
      definition,
      allocations,
      safeCandidates,
      {risk: risk, stress: stress}
    );
  }).sort(function(a, b) {
    if (b.scenarioScore !== a.scenarioScore) {
      return b.scenarioScore - a.scenarioScore;
    }
    return a.scenarioId < b.scenarioId ? -1 : 1;
  });

  scenarios.forEach(function(scenario, index) {
    scenario.rank = index + 1;
    scenario.preferred = index === 0;
  });

  return {
    scenarios: scenarios,
    candidateCount: safeCandidates.length,
    risk: risk,
    stress: stress
  };
}

function foBuildScenarioAllocations_(mode, candidates, context) {
  const funded = candidates.filter(function(candidate) {
    return candidate.optimizedIncrementalWeight > 0;
  });
  const totalOptimized = funded.reduce(function(sum, candidate) {
    return sum + candidate.optimizedIncrementalWeight;
  }, 0);
  const equalShare = funded.length ? totalOptimized / funded.length : 0;
  const focusedKeys = funded.slice().sort(function(a, b) {
    if (b.deploymentScore !== a.deploymentScore) {
      return b.deploymentScore - a.deploymentScore;
    }
    return a.rank - b.rank;
  }).slice(0, 2).reduce(function(result, candidate) {
    result[foScenarioCandidateKey_(candidate)] = true;
    return result;
  }, {});

  return candidates.map(function(candidate) {
    let increment = 0;

    if (mode === 'OPTIMIZED') {
      increment = candidate.optimizedIncrementalWeight;
    } else if (mode === 'DIVERSIFIED') {
      increment = Math.min(
        candidate.optimizedIncrementalWeight,
        equalShare
      );
    } else if (mode === 'FOCUSED') {
      increment = focusedKeys[foScenarioCandidateKey_(candidate)]
        ? candidate.optimizedIncrementalWeight
        : 0;
    } else if (mode === 'DEFENSIVE') {
      const blocksDeployment =
        context.risk.critical ||
        context.stress.stressPressureScore >= 75;
      increment = blocksDeployment
        ? 0
        : candidate.optimizedIncrementalWeight * 0.25;
    }

    increment = Math.min(
      Math.max(0, increment),
      candidate.optimizedIncrementalWeight,
      Math.max(
        0,
        candidate.maximumPositionWeight -
          candidate.currentPortfolioWeight
      )
    );

    return {
      rank: candidate.rank,
      ticker: candidate.ticker,
      account: candidate.account,
      deploymentDecision: candidate.deploymentDecision,
      deploymentScore: candidate.deploymentScore,
      currentPortfolioWeight: candidate.currentPortfolioWeight,
      optimizedIncrementalWeight:
        candidate.optimizedIncrementalWeight,
      proposedIncrementalWeight: foScenarioRoundWeight_(increment),
      proposedTargetWeight: foScenarioRoundWeight_(
        candidate.currentPortfolioWeight + increment
      ),
      maximumPositionWeight: candidate.maximumPositionWeight,
      upstreamConstraintStatus: candidate.constraintStatus,
      upstreamConstraintReason: candidate.constraintReason
    };
  });
}

function foScorePortfolioScenario_(
  definition,
  allocations,
  candidates,
  context
) {
  const funded = allocations.filter(function(item) {
    return item.proposedIncrementalWeight > 0;
  });
  const totalIncrement = allocations.reduce(function(sum, item) {
    return sum + item.proposedIncrementalWeight;
  }, 0);
  const optimizedWeightedPotential = candidates.reduce(function(sum, item) {
    return sum +
      item.optimizedIncrementalWeight *
      Math.max(item.deploymentScore, 0);
  }, 0);
  const proposedWeightedDeployment = allocations.reduce(function(sum, item) {
    return sum +
      item.proposedIncrementalWeight *
      Math.max(item.deploymentScore, 0);
  }, 0);
  const deploymentAlignmentScore = optimizedWeightedPotential > 0
    ? foScenarioRoundScore_(
      proposedWeightedDeployment / optimizedWeightedPotential * 100
    )
    : (totalIncrement === 0 ? 100 : 0);

  const largestTargetWeight = allocations.reduce(function(maximum, item) {
    return Math.max(maximum, item.proposedTargetWeight);
  }, context.risk.largestPositionWeight || 0);

  const diversificationGain =
    funded.length * 3 -
    Math.max(
      0,
      largestTargetWeight -
        (context.risk.largestPositionWeight || 0)
    ) * 200;
  const diversificationScore = foScenarioRoundScore_(
    foScenarioClamp_(
      (context.risk.diversificationScore || 50) + diversificationGain,
      0,
      100
    )
  );

  const deploymentStress =
    totalIncrement * 100 *
    (context.stress.stressPressureScore || 0) / 100;
  let riskDisciplineScore = foScenarioClamp_(
    100 -
      (context.risk.riskScore || 0) -
      deploymentStress * 1.5,
    0,
    100
  );
  let stressDisciplineScore = foScenarioClamp_(
    100 - deploymentStress * 2,
    0,
    100
  );

  if (definition.mode === 'DEFENSIVE') {
    riskDisciplineScore = foScenarioClamp_(
      riskDisciplineScore + 20,
      0,
      100
    );
    stressDisciplineScore = foScenarioClamp_(
      stressDisciplineScore + 20,
      0,
      100
    );
  }

  const constraintBreaches = allocations.filter(function(item) {
    return (
      item.proposedIncrementalWeight >
        item.optimizedIncrementalWeight + 0.000001 ||
      item.proposedTargetWeight >
        item.maximumPositionWeight + 0.000001
    );
  }).length;
  const constraintComplianceScore =
    constraintBreaches > 0 ? 0 : 100;

  let contextAdjustment = 0;
  if (
    definition.mode === 'OPTIMIZED' &&
    !context.risk.critical &&
    context.stress.stressPressureScore < 75
  ) {
    contextAdjustment += 8;
  }
  if (
    definition.mode === 'DIVERSIFIED' &&
    (
      context.risk.diversificationScore < 65 ||
      context.risk.largestPositionWeight >= 0.20
    )
  ) {
    contextAdjustment += 12;
  }
  if (
    definition.mode === 'FOCUSED' &&
    context.risk.riskScore < 50 &&
    context.stress.stressPressureScore < 60
  ) {
    contextAdjustment += 6;
  }
  if (definition.mode === 'DEFENSIVE') {
    contextAdjustment += (
      context.risk.critical ||
      context.stress.stressPressureScore >= 75
    ) ? 30 : -10;
  }
  if (
    definition.mode === 'CURRENT' &&
    !candidates.some(function(item) {
      return item.optimizedIncrementalWeight > 0;
    })
  ) {
    contextAdjustment += 20;
  }

  let score =
    deploymentAlignmentScore * 0.35 +
    diversificationScore * 0.25 +
    riskDisciplineScore * 0.20 +
    stressDisciplineScore * 0.10 +
    constraintComplianceScore * 0.10 +
    contextAdjustment;

  if (
    (context.risk.critical ||
      context.stress.stressPressureScore >= 80) &&
    totalIncrement > 0
  ) {
    score -= 35;
  }

  score = foScenarioRoundScore_(foScenarioClamp_(score, 0, 100));

  return {
    rank: 0,
    scenarioId: definition.scenarioId,
    scenarioName: definition.scenarioName,
    scenarioType: definition.scenarioType,
    preferred: false,
    scenarioScore: score,
    totalIncrementalWeight: foScenarioRoundWeight_(totalIncrement),
    fundedCandidateCount: funded.length,
    largestTargetWeight: foScenarioRoundWeight_(largestTargetWeight),
    deploymentAlignmentScore: deploymentAlignmentScore,
    diversificationScore: diversificationScore,
    riskDisciplineScore: foScenarioRoundScore_(riskDisciplineScore),
    stressDisciplineScore:
      foScenarioRoundScore_(stressDisciplineScore),
    constraintComplianceScore: constraintComplianceScore,
    constraintBreachCount: constraintBreaches,
    portfolioRiskLevel: context.risk.overallRisk || 'UNKNOWN',
    stressContext:
      String(context.stress.highestSeverity || 'UNKNOWN') +
      ' / ' +
      String(context.stress.enabledScenarioCount || 0) +
      ' ENABLED',
    executiveRecommendation:
      foPortfolioScenarioRecommendation_(
        definition.mode,
        totalIncrement,
        context
      ),
    scenarioRationale:
      foPortfolioScenarioRationale_(
        definition,
        score,
        totalIncrement,
        funded.length,
        largestTargetWeight,
        diversificationScore,
        riskDisciplineScore,
        context
      ),
    allocations: allocations
  };
}

function foPortfolioScenarioRecommendation_(mode, totalIncrement, context) {
  if (mode === 'DEFENSIVE') {
    return (
      context.risk.critical ||
      context.stress.stressPressureScore >= 75
    )
      ? 'DEFER NEW DEPLOYMENT; PRESERVE RISK CAPACITY'
      : 'USE LIMITED RISK-CONTROLLED DEPLOYMENT';
  }
  if (mode === 'DIVERSIFIED') {
    return 'PREFER BROADER, CONCENTRATION-AWARE DEPLOYMENT';
  }
  if (mode === 'FOCUSED') {
    return 'LIMIT DEPLOYMENT TO THE HIGHEST-PRIORITY CANDIDATES';
  }
  if (mode === 'OPTIMIZED') {
    return totalIncrement > 0
      ? 'PROCEED WITH SELECTIVE OPTIMIZED DEPLOYMENT'
      : 'NO ELIGIBLE OPTIMIZED DEPLOYMENT';
  }
  return 'MAINTAIN CURRENT ALLOCATION';
}

function foPortfolioScenarioRationale_(
  definition,
  score,
  totalIncrement,
  fundedCount,
  largestTargetWeight,
  diversificationScore,
  riskDisciplineScore,
  context
) {
  return definition.scenarioName + ' scored ' + score +
    '/100. Proposed incremental weight: ' +
    foScenarioPercentText_(totalIncrement) +
    '; funded candidates: ' + fundedCount +
    '; largest projected position: ' +
    foScenarioPercentText_(largestTargetWeight) +
    '; diversification score: ' + diversificationScore +
    '; risk-discipline score: ' +
    foScenarioRoundScore_(riskDisciplineScore) +
    '; portfolio risk: ' +
    String(context.risk.overallRisk || 'UNKNOWN') +
    '; stress pressure: ' +
    foScenarioRoundScore_(
      context.stress.stressPressureScore || 0
    ) + '/100.';
}

function foWritePortfolioScenarioIntelligence_(dashboard, comparison) {
  const scenarioHeaders = foPortfolioScenarioHeaders_();
  const scenarioSheet = foEnsurePortfolioScenarioContract_(
    dashboard,
    FO_SHEETS.PORTFOLIO_SCENARIOS,
    scenarioHeaders
  );
  const summaryHeaders = foPortfolioScenarioSummaryHeaders_();
  const summarySheet = foEnsurePortfolioScenarioContract_(
    dashboard,
    FO_SHEETS.PORTFOLIO_SCENARIO_SUMMARY,
    summaryHeaders
  );
  const timestamp = new Date();

  scenarioSheet.clearContents();
  scenarioSheet.getRange(
    1,
    1,
    1,
    scenarioHeaders.length
  ).setValues([scenarioHeaders]);

  const detailRows = [];
  comparison.scenarios.forEach(function(scenario) {
    scenario.allocations.forEach(function(item) {
      detailRows.push([
        scenario.rank,
        scenario.scenarioId,
        scenario.scenarioName,
        scenario.scenarioType,
        scenario.preferred ? 'YES' : 'NO',
        item.ticker,
        item.account,
        item.deploymentDecision,
        item.deploymentScore,
        item.currentPortfolioWeight,
        item.optimizedIncrementalWeight,
        item.proposedIncrementalWeight,
        item.proposedTargetWeight,
        item.maximumPositionWeight,
        item.upstreamConstraintStatus,
        item.upstreamConstraintReason,
        scenario.scenarioScore,
        scenario.deploymentAlignmentScore,
        scenario.diversificationScore,
        scenario.riskDisciplineScore,
        scenario.stressDisciplineScore,
        scenario.constraintComplianceScore,
        scenario.executiveRecommendation,
        scenario.scenarioRationale,
        timestamp,
        FO_CONFIG.PLATFORM_VERSION,
        FO_CONFIG.BASELINE
      ]);
    });
  });

  if (detailRows.length) {
    scenarioSheet.getRange(
      2,
      1,
      detailRows.length,
      scenarioHeaders.length
    ).setValues(detailRows);
  }

  summarySheet.clearContents();
  summarySheet.getRange(
    1,
    1,
    1,
    summaryHeaders.length
  ).setValues([summaryHeaders]);

  const summaryRows = comparison.scenarios.map(function(scenario) {
    return [
      scenario.rank,
      scenario.scenarioId,
      scenario.scenarioName,
      scenario.scenarioType,
      scenario.preferred ? 'YES' : 'NO',
      scenario.scenarioScore,
      scenario.totalIncrementalWeight,
      scenario.fundedCandidateCount,
      scenario.largestTargetWeight,
      scenario.deploymentAlignmentScore,
      scenario.diversificationScore,
      scenario.riskDisciplineScore,
      scenario.stressDisciplineScore,
      scenario.constraintComplianceScore,
      scenario.constraintBreachCount,
      scenario.portfolioRiskLevel,
      scenario.stressContext,
      scenario.executiveRecommendation,
      scenario.scenarioRationale,
      timestamp,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE
    ];
  });

  if (summaryRows.length) {
    summarySheet.getRange(
      2,
      1,
      summaryRows.length,
      summaryHeaders.length
    ).setValues(summaryRows);
  }

  foFormatPortfolioScenarioSheet_(
    scenarioSheet,
    scenarioHeaders,
    detailRows.length,
    [
      'Current Portfolio Weight',
      'Optimized Incremental Weight',
      'Proposed Incremental Weight',
      'Proposed Target Weight',
      'Maximum Position Weight'
    ]
  );
  foFormatPortfolioScenarioSheet_(
    summarySheet,
    summaryHeaders,
    summaryRows.length,
    [
      'Total Incremental Weight',
      'Largest Target Weight'
    ]
  );
}

function foFormatPortfolioScenarioSheet_(
  sheet,
  headers,
  rowCount,
  percentHeaders
) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff');

  percentHeaders.forEach(function(header) {
    const column = headers.indexOf(header) + 1;
    if (column > 0) {
      sheet.getRange(
        2,
        column,
        Math.max(rowCount, 1),
        1
      ).setNumberFormat('0.00%');
    }
  });

  sheet.autoResizeColumns(1, headers.length);
  [
    'Upstream Constraint Reason',
    'Executive Recommendation',
    'Scenario Rationale'
  ].forEach(function(header) {
    const column = headers.indexOf(header) + 1;
    if (column > 0) sheet.setColumnWidth(column, 520);
  });
}

function foPortfolioScenarioHeaders_() {
  return [
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
  ];
}

function foPortfolioScenarioSummaryHeaders_() {
  return [
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
  ];
}

function foRunPortfolioScenarioIntelligenceSmokeTest() {
  const setup = foSetupPortfolioScenarioIntelligence();
  const result = foRunPortfolioScenarioIntelligence();

  if (result.status !== 'SUCCESS') {
    throw new Error(
      'Portfolio Scenario Intelligence smoke test failed: ' +
      JSON.stringify(result)
    );
  }
  if (!result.preferredScenario) {
    throw new Error('Preferred portfolio scenario is missing.');
  }

  return {
    status: 'PASS',
    setup: setup,
    result: result
  };
}

function foScenarioCandidateKey_(candidate) {
  return String(candidate.ticker || '').trim().toUpperCase() +
    '|' +
    String(candidate.account || '').trim().toUpperCase();
}

function foScenarioVal_(row, headers, name) {
  const index = headers.indexOf(name);
  return index >= 0 ? row[index] : '';
}

function foScenarioNumber_(value, fallback) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return fallback;
  }
  const number = Number(String(value).replace(/[$,%\s]/g, ''));
  return Number.isFinite(number) ? number : fallback;
}

function foScenarioNormalizeWeight_(value) {
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

function foScenarioScore_(value, fallback) {
  const number = foScenarioNumber_(value, fallback);
  if (!Number.isFinite(number)) return fallback;
  if (number >= 0 && number <= 1) return number * 100;
  return foScenarioClamp_(number, 0, 100);
}

function foScenarioClamp_(value, minimum, maximum) {
  return Math.min(maximum, Math.max(minimum, Number(value) || 0));
}

function foScenarioRoundWeight_(value) {
  return Math.round((Number(value) || 0) * 1000000) / 1000000;
}

function foScenarioRoundScore_(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function foScenarioPercentText_(value) {
  return (foScenarioRoundWeight_(value) * 100).toFixed(2) + '%';
}
