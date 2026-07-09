/************************************************************
 * RecommendationEngine.gs
 * Wave 1C.3 — Recommendation Engine
 ************************************************************/

function foCreateRecommendation(input) {
  const module = 'RecommendationEngine';

  try {
    foInfo_(module, 'Start', 'Creating recommendation.');

    const validation = foValidateRecommendationInput_(input);

    if (!validation.valid) {
      foWarn_(module, 'Validation Failed', validation.issues.join(' | '));
      throw new Error('Recommendation validation failed: ' + validation.issues.join(' | '));
    }

    const recommendationId = input.recommendationId || foGenerateRecommendationId_(input.ticker);
    const score = foCalculateInitialRecommendationScore_(input);
    const recommendationType = foClassifyRecommendationType_(input, score);

    const result = foAppendRecommendationEvent({
      recommendationId: recommendationId,
      ticker: input.ticker,
      company: input.company,
      assetClass: input.assetClass || 'Equity',
      sector: input.sector || '',
      investmentTheme: input.investmentTheme || '',
      recommendationType: recommendationType,
      recommendationAction: input.recommendationAction || 'Watch',
      recommendationStatus: input.recommendationStatus || 'Proposed',
      priority: input.priority || foAssignPriority_(score),
      materialityScore: input.materialityScore || score.materialityScore,
      confidenceScore: input.confidenceScore || score.confidenceScore,
      entryPrice: input.entryPrice || '',
      buyZoneLow: input.buyZoneLow || '',
      buyZoneHigh: input.buyZoneHigh || '',
      targetPrice: input.targetPrice || '',
      stopLoss: input.stopLoss || '',
      timeHorizon: input.timeHorizon || '',
      suggestedAllocation: input.suggestedAllocation || score.suggestedAllocation,
      portfolioAccount: input.portfolioAccount || '',
      decisionStatus: input.decisionStatus || 'Pending',
      executionStatus: input.executionStatus || 'Not Executed',
      riskRating: input.riskRating || score.riskRating,
      expectedAlpha: input.expectedAlpha || '',
      benchmark: input.benchmark || '',
      investmentThesis: input.investmentThesis || '',
      supportingEvidence: input.supportingEvidence || '',
      assumptions: input.assumptions || '',
      risks: input.risks || '',
      cioNotes:
        (input.cioNotes || '') +
        ' | Engine Score: ' +
        score.overallScore +
        ' | Classification: ' +
        recommendationType
    });

    foInfo_(
      module,
      'Complete',
      'Recommendation created: ' + recommendationId
    );

    return {
      status: 'SUCCESS',
      recommendationId: recommendationId,
      eventId: result.eventId,
      score: score,
      recommendationType: recommendationType
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foValidateRecommendationInput_(input) {
  const issues = [];

  if (!input) {
    issues.push('Input object is required.');
    return {
      valid: false,
      issues: issues
    };
  }

  if (!input.ticker) {
    issues.push('Ticker is required.');
  }

  if (!input.company) {
    issues.push('Company is required.');
  }

  if (!input.recommendationAction) {
    issues.push('Recommendation Action is required.');
  }

  if (!input.investmentThesis) {
    issues.push('Investment thesis is required.');
  }

  return {
    valid: issues.length === 0,
    issues: issues
  };
}

function foGenerateRecommendationId_(ticker) {
  const cleanTicker = String(ticker || 'UNKNOWN')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');

  const ts = Utilities.formatDate(
    new Date(),
    FO_CONFIG.TIMEZONE,
    'yyyyMMdd-HHmmss'
  );

  return 'REC-' + cleanTicker + '-' + ts;
}

function foCalculateInitialRecommendationScore_(input) {
  const confidence = Number(input.confidenceScore || 0);
  const materiality = Number(input.materialityScore || 0);
  const conviction = Number(input.convictionScore || confidence || 0);
  const risk = String(input.riskRating || '').toLowerCase();

  let riskPenalty = 0;

  if (risk === 'high') riskPenalty = 15;
  if (risk === 'medium') riskPenalty = 8;
  if (risk === 'low') riskPenalty = 2;

  const rawScore =
    confidence * 0.4 +
    materiality * 0.3 +
    conviction * 0.3 -
    riskPenalty;

  const overallScore = Math.max(0, Math.min(100, Math.round(rawScore)));

  return {
    confidenceScore: confidence,
    materialityScore: materiality,
    convictionScore: conviction,
    riskRating: input.riskRating || 'Medium',
    riskPenalty: riskPenalty,
    overallScore: overallScore,
    suggestedAllocation: foSuggestAllocation_(overallScore, input.riskRating)
  };
}

function foSuggestAllocation_(score, riskRating) {
  const risk = String(riskRating || '').toLowerCase();

  if (score >= 90 && risk === 'low') return 5;
  if (score >= 85) return 4;
  if (score >= 75) return 3;
  if (score >= 65) return 2;
  if (score >= 50) return 1;

  return 0;
}

function foAssignPriority_(score) {
  const overallScore = Number(score.overallScore || 0);

  if (overallScore >= 85) return 'High';
  if (overallScore >= 65) return 'Medium';
  return 'Low';
}

function foClassifyRecommendationType_(input, score) {
  const action = String(input.recommendationAction || '').toLowerCase();
  const overallScore = Number(score.overallScore || 0);

  if (action === 'buy' && overallScore >= 85) return 'High Conviction Buy';
  if (action === 'buy' && overallScore >= 70) return 'Buy Candidate';
  if (action === 'hold') return 'Hold / Monitor';
  if (action === 'sell') return 'Sell / Exit Review';
  if (action === 'watch') return 'Watchlist Candidate';

  return 'General Recommendation';
}

function foRunRecommendationEngineSmokeTest() {
  const module = 'RecommendationEngine';

  try {
    foInfo_(module, 'Start', 'Recommendation Engine smoke test started.');

    const result = foCreateRecommendation({
      ticker: 'SAMPLE',
      company: 'Sample Company',
      assetClass: 'Equity',
      sector: 'Technology',
      investmentTheme: 'Automation Smoke Test',
      recommendationAction: 'Watch',
      recommendationStatus: 'Proposed',
      confidenceScore: 70,
      materialityScore: 60,
      convictionScore: 65,
      riskRating: 'Medium',
      timeHorizon: '12 months',
      benchmark: 'QQC',
      investmentThesis: 'Smoke test recommendation created by Recommendation Engine.',
      supportingEvidence: 'Generated as part of Wave 1C.3 implementation.',
      assumptions: 'No investment decision. Test record only.',
      risks: 'Test data should not be used for real investment decisions.',
      cioNotes: 'Wave 1C.3 smoke test.'
    });

    foInfo_(
      module,
      'Complete',
      'Recommendation Engine smoke test completed: ' + result.recommendationId
    );

    return result;

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}