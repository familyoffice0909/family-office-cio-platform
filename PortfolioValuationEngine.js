/************************************************************
 * PortfolioValuationEngine.gs
 * Sprint v3.2.5 — Portfolio Valuation Evidence and Reconciliation
 ************************************************************/

function foRunPortfolioValuation() {
  const module = 'PortfolioValuationEngine';

  try {
    foInfo_(module, 'Start', 'Portfolio valuation started.');

    const dashboard = foDashboard_();
    const portfolioSheet = dashboard.getSheetByName(FO_SHEETS.PORTFOLIO_MASTER);

    if (!portfolioSheet) {
      throw new Error('Portfolio Master sheet not found.');
    }

    const values = portfolioSheet.getDataRange().getValues();
    const headers = values[0].map(String);

    const result = foCalculatePortfolioValuation_(portfolioSheet, values, headers);
    foWritePortfolioValuationSummary_(dashboard, result);

    foInfo_(module, 'Complete', 'Portfolio valuation completed.');

    return result;

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foCalculatePortfolioValuation_(portfolioSheet, values, headers) {
  const tickerIndex = headers.indexOf('Ticker');
  const quantityIndex = headers.indexOf('Quantity');
  const priceIndex = headers.indexOf('Current Price');
  const marketValueIndex = headers.indexOf('Market Value');
  const costBasisIndex = headers.indexOf('Cost Basis');
  const accountIndex = headers.indexOf('Account');
  const priceTimestampIndex = headers.indexOf('Price Timestamp');
  const priceSourceIndex = headers.indexOf('Price Source');
  const priceStatusIndex = headers.indexOf('Price Status');
  const priceBasisIndex = headers.indexOf('Price Basis');
  const valuationStatusIndex = headers.indexOf('Valuation Status');
  const marketValueBasisIndex = headers.indexOf('Market Value Basis');

  if (
    tickerIndex < 0 ||
    quantityIndex < 0 ||
    priceIndex < 0 ||
    marketValueIndex < 0 ||
    costBasisIndex < 0
  ) {
    throw new Error('Portfolio Master valuation schema is incomplete.');
  }

  let totalMarketValue = 0;
  let totalCostBasis = 0;
  let valuedPositions = 0;
  let missingPriceCount = 0;
  let totalActivePositions = 0;
  let positionsWithCostBasis = 0;

  const holdingEvidence = [];
  const accountEvidenceMap = {};

  for (let r = 1; r < values.length; r++) {
    const ticker = String(values[r][tickerIndex] || '')
      .trim()
      .toUpperCase();

    if (!ticker) continue;

    const account =
      accountIndex >= 0
        ? String(values[r][accountIndex] || '').trim().toUpperCase()
        : '';

    const quantity = foSafeNumber_(values[r][quantityIndex]);
    const price = foSafeNumber_(values[r][priceIndex]);
    const costBasis =
      costBasisIndex >= 0
        ? foSafeNumber_(values[r][costBasisIndex])
        : 0;

    if (foIsExcludedValuationRow_(account, ticker, quantity, price)) continue;
    if (quantity <= 0) continue;

    totalActivePositions++;

    const priceTimestamp =
      priceTimestampIndex >= 0
        ? values[r][priceTimestampIndex]
        : '';

    const priceSource =
      priceSourceIndex >= 0
        ? values[r][priceSourceIndex]
        : '';

    const priceStatus =
      priceStatusIndex >= 0
        ? values[r][priceStatusIndex]
        : '';

    const priceBasis =
      priceBasisIndex >= 0
        ? values[r][priceBasisIndex]
        : '';

    const valuationStatus =
      valuationStatusIndex >= 0
        ? values[r][valuationStatusIndex]
        : '';

    const marketValueBasis =
      marketValueBasisIndex >= 0
        ? values[r][marketValueBasisIndex]
        : '';

    totalCostBasis += costBasis;

    if (costBasis > 0) {
      positionsWithCostBasis++;
    }

    const persistedMarketValue =
      foSafeNumber_(values[r][marketValueIndex]);

    let marketValue = 0;
    let hasUsableValuation = false;

    if (
      marketValueBasis === 'PERSISTED_FALLBACK' &&
      persistedMarketValue > 0
    ) {
      marketValue = persistedMarketValue;
      hasUsableValuation = true;
    } else if (price > 0) {
      marketValue = quantity * price;
      hasUsableValuation = true;

      portfolioSheet
        .getRange(r + 1, marketValueIndex + 1)
        .setValue(marketValue);
    }

    if (!hasUsableValuation) {
      missingPriceCount++;

      portfolioSheet
        .getRange(r + 1, marketValueIndex + 1)
        .clearContent();

      continue;
    }

    totalMarketValue += marketValue;
    valuedPositions++;

    holdingEvidence.push({
      account: account,
      ticker: ticker,
      quantity: quantity,
      marketValue: marketValue,
      costBasis: costBasis,
      priceTimestamp: priceTimestamp,
      priceSource: priceSource,
      priceStatus: priceStatus,
      priceBasis: priceBasis,
      valuationStatus: valuationStatus,
      marketValueBasis: marketValueBasis
    });

    if (!accountEvidenceMap[account]) {
      accountEvidenceMap[account] = {
        account: account,
        marketValue: 0,
        costBasis: 0,
        positions: 0
      };
    }

    accountEvidenceMap[account].marketValue += marketValue;
    accountEvidenceMap[account].costBasis += costBasis;
    accountEvidenceMap[account].positions++;
  }

  const unrealizedGainLoss = totalMarketValue - totalCostBasis;

  const unrealizedGainLossPct =
    totalCostBasis > 0
      ? unrealizedGainLoss / totalCostBasis
      : 0;

  const accountEvidence =
    Object.keys(accountEvidenceMap).map(function(key) {
      return accountEvidenceMap[key];
    });

  const priceCoveragePct =
    totalActivePositions > 0
      ? valuedPositions / totalActivePositions
      : 0;

  const costBasisCoveragePct =
    totalActivePositions > 0
      ? positionsWithCostBasis / totalActivePositions
      : 0;

  const reconciledAccountMarketValue =
    accountEvidence.reduce(function(total, item) {
      return total + foSafeNumber_(item.marketValue);
    }, 0);

  const reconciliationVariance =
    totalMarketValue - reconciledAccountMarketValue;

  const reconciliationStatus =
    Math.abs(reconciliationVariance) < 0.01
      ? 'RECONCILED'
      : 'REVIEW_REQUIRED';

  const certificationStatus =
    missingPriceCount === 0 &&
    reconciliationStatus === 'RECONCILED'
      ? 'CERTIFIED'
      : 'PARTIALLY_CERTIFIED';

  return {
    status: 'SUCCESS',
    totalMarketValue: totalMarketValue,
    totalCostBasis: totalCostBasis,
    unrealizedGainLoss: unrealizedGainLoss,
    unrealizedGainLossPct: unrealizedGainLossPct,
    valuedPositions: valuedPositions,
    missingPriceCount: missingPriceCount,
    totalActivePositions: totalActivePositions,
    priceCoveragePct: priceCoveragePct,
    costBasisCoveragePct: costBasisCoveragePct,
    reconciliationVariance: reconciliationVariance,
    reconciliationStatus: reconciliationStatus,
    certificationStatus: certificationStatus,
    holdingEvidence: holdingEvidence,
    accountEvidence: accountEvidence
  };
}

function foSafeNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;

  const cleaned = String(value)
    .replace(/\$/g, '')
    .replace(/,/g, '')
    .replace(/%/g, '')
    .trim();

  const number = Number(cleaned);

  return isNaN(number) ? 0 : number;
}

function foIsExcludedValuationRow_(account, ticker, quantity, price) {
  const excludedAccounts = [
    '',
    'N/A',
    'NA',
    'PENDING',
    'REFERENCE',
    'LIBRARY',
    'WATCHLIST',
    'WATCH LIST',
    'TEMPLATE'
  ];

  if (
    excludedAccounts.indexOf(account) >= 0 &&
    quantity <= 0 &&
    price <= 0
  ) {
    return true;
  }

  return false;
}

function foWritePortfolioValuationSummary_(dashboard, result) {
  const sheet = foEnsureSheet_(dashboard, 'Portfolio Valuation Summary', [
    'Timestamp',
    'Metric',
    'Value',
    'Platform Version',
    'Baseline'
  ]);

  sheet.clearContents();

  sheet.getRange(1, 1, 1, 5).setValues([[
    'Timestamp',
    'Metric',
    'Value',
    'Platform Version',
    'Baseline'
  ]]);

  const timestamp = new Date();

  const rows = [
    [timestamp, 'Total Market Value', result.totalMarketValue, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Total Cost Basis', result.totalCostBasis, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Unrealized Gain/Loss', result.unrealizedGainLoss, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Unrealized Gain/Loss %', result.unrealizedGainLossPct, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Valued Positions', result.valuedPositions, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Missing Price Count', result.missingPriceCount, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Total Active Positions', result.totalActivePositions, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Price Coverage %', result.priceCoveragePct, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Cost Basis Coverage %', result.costBasisCoveragePct, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Reconciliation Variance', result.reconciliationVariance, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Reconciliation Status', result.reconciliationStatus, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Certification Status', result.certificationStatus, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE]
  ];

  sheet.getRange(2, 1, rows.length, 5).setValues(rows);
}

function foRunPortfolioValuationSmokeTest() {
  const module = 'PortfolioValuationEngine';

  try {
    foInfo_(module, 'Start', 'Portfolio Valuation smoke test started.');

    const result = foRunPortfolioValuation();

    foInfo_(module, 'Complete', 'Portfolio Valuation smoke test completed.');

    return result;

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}
