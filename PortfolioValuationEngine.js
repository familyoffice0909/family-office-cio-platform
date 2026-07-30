/************************************************************
 * PortfolioValuationEngine.gs
 * Sprint v3.2.5 — Portfolio Valuation Evidence and Reconciliation
 * Sprint v3.2.6 — Comparable Valuation and Reporting Integrity
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

    // Convert the runtime response into an Execution API-safe object.
    const runtimeResult = JSON.parse(
      JSON.stringify(result, function (key, value) {
        if (value instanceof Date) {
          return value.toISOString();
        }
        return value;
      })
    );

    return runtimeResult;
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
    tickerIndex < 0 || quantityIndex < 0 || priceIndex < 0 ||
    marketValueIndex < 0 || costBasisIndex < 0
  ) {
    throw new Error('Portfolio Master valuation schema is incomplete.');
  }

  let totalMarketValue = 0;
  let totalCostBasis = 0;
  let comparableCostBasis = 0;
  let valuedPositions = 0;
  let missingPriceCount = 0;
  let totalActivePositions = 0;
  let positionsWithCostBasis = 0;
  let latestPriceTimestamp = null;

  const holdingEvidence = [];
  const accountEvidenceMap = {};
  const missingPriceTickers = [];
  const missingCostBasisTickers = [];
  const priceBasisSet = {};
  const valuationTimestamp = new Date();

  for (let r = 1; r < values.length; r++) {
    const ticker = String(values[r][tickerIndex] || '').trim().toUpperCase();
    if (!ticker) continue;

    const account = accountIndex >= 0
      ? String(values[r][accountIndex] || '').trim().toUpperCase()
      : '';
    const quantity = foSafeNumber_(values[r][quantityIndex]);
    const price = foSafeNumber_(values[r][priceIndex]);
    const costBasis = foSafeNumber_(values[r][costBasisIndex]);

    if (foIsExcludedValuationRow_(account, ticker, quantity, price)) continue;
    if (quantity <= 0) continue;

    totalActivePositions++;
    totalCostBasis += costBasis;

    if (costBasis > 0) {
      positionsWithCostBasis++;
    } else {
      missingCostBasisTickers.push(ticker);
    }

    const priceTimestamp = priceTimestampIndex >= 0 ? values[r][priceTimestampIndex] : '';
    const priceSource = priceSourceIndex >= 0 ? values[r][priceSourceIndex] : '';
    const priceStatus = priceStatusIndex >= 0 ? values[r][priceStatusIndex] : '';
    const priceBasis = priceBasisIndex >= 0 ? values[r][priceBasisIndex] : '';
    const valuationStatus = valuationStatusIndex >= 0 ? values[r][valuationStatusIndex] : '';
    const marketValueBasis = marketValueBasisIndex >= 0 ? values[r][marketValueBasisIndex] : '';
    const persistedMarketValue = foSafeNumber_(values[r][marketValueIndex]);

    let marketValue = 0;
    let hasUsableValuation = false;

    if (String(marketValueBasis).toUpperCase() === 'PERSISTED_FALLBACK' && persistedMarketValue > 0) {
      marketValue = persistedMarketValue;
      hasUsableValuation = true;
    } else if (price > 0) {
      marketValue = quantity * price;
      hasUsableValuation = true;
      portfolioSheet.getRange(r + 1, marketValueIndex + 1).setValue(marketValue);
    }

    if (!hasUsableValuation) {
      missingPriceCount++;
      missingPriceTickers.push(ticker);
      portfolioSheet.getRange(r + 1, marketValueIndex + 1).clearContent();
      continue;
    }

    totalMarketValue += marketValue;
    comparableCostBasis += costBasis;
    valuedPositions++;

    const normalizedPriceBasis = foNormalizePriceBasis_(priceBasis, marketValueBasis);
    priceBasisSet[normalizedPriceBasis] = true;

    const parsedTimestamp = foParseValuationTimestamp_(priceTimestamp);
    if (parsedTimestamp && (!latestPriceTimestamp || parsedTimestamp > latestPriceTimestamp)) {
      latestPriceTimestamp = parsedTimestamp;
    }

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
        positions: 0,
        cashIncluded: false
      };
    }

    accountEvidenceMap[account].marketValue += marketValue;
    accountEvidenceMap[account].costBasis += costBasis;
    accountEvidenceMap[account].positions++;
    if (foIsCashTicker_(ticker)) accountEvidenceMap[account].cashIncluded = true;
  }

  const priceCoveragePct = totalActivePositions > 0
    ? valuedPositions / totalActivePositions
    : 0;
  const costBasisCoveragePct = totalActivePositions > 0
    ? positionsWithCostBasis / totalActivePositions
    : 0;
  const fullPortfolioReturnEligible =
    totalActivePositions > 0 && missingPriceCount === 0 && priceCoveragePct === 1;

  const comparableUnrealizedGainLoss = totalMarketValue - comparableCostBasis;
  const comparableUnrealizedGainLossPct = comparableCostBasis > 0
    ? comparableUnrealizedGainLoss / comparableCostBasis
    : null;
  const unrealizedGainLoss = fullPortfolioReturnEligible
    ? totalMarketValue - totalCostBasis
    : null;
  const unrealizedGainLossPct = fullPortfolioReturnEligible && totalCostBasis > 0
    ? unrealizedGainLoss / totalCostBasis
    : null;

  const accountEvidence = Object.keys(accountEvidenceMap).map(function(key) {
    return accountEvidenceMap[key];
  });
  const reconciledAccountMarketValue = accountEvidence.reduce(function(total, item) {
    return total + foSafeNumber_(item.marketValue);
  }, 0);
  const reconciliationVariance = totalMarketValue - reconciledAccountMarketValue;
  const reconciliationStatus = Math.abs(reconciliationVariance) < 0.01
    ? 'RECONCILED'
    : 'REVIEW_REQUIRED';
  const valuationCompletenessStatus = totalActivePositions === 0
    ? 'UNAVAILABLE'
    : (fullPortfolioReturnEligible ? 'COMPLETE' : 'PARTIAL');
  const certificationStatus =
    valuationCompletenessStatus === 'COMPLETE' && reconciliationStatus === 'RECONCILED'
      ? 'CERTIFIED'
      : 'PARTIALLY_CERTIFIED';

  return {
    status: 'SUCCESS',
    valuationTimestamp: valuationTimestamp,
    latestPriceTimestamp: latestPriceTimestamp || '',
    portfolioPriceBasis: foAggregatePriceBasis_(priceBasisSet),
    valuationCompletenessStatus: valuationCompletenessStatus,
    totalMarketValue: totalMarketValue,
    totalCostBasis: totalCostBasis,
    comparableCostBasis: comparableCostBasis,
    unrealizedGainLoss: unrealizedGainLoss,
    unrealizedGainLossPct: unrealizedGainLossPct,
    comparableUnrealizedGainLoss: comparableUnrealizedGainLoss,
    comparableUnrealizedGainLossPct: comparableUnrealizedGainLossPct,
    fullPortfolioReturnEligible: fullPortfolioReturnEligible,
    valuedPositions: valuedPositions,
    missingPriceCount: missingPriceCount,
    missingPriceTickers: missingPriceTickers,
    missingCostBasisCount: missingCostBasisTickers.length,
    missingCostBasisTickers: missingCostBasisTickers,
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

function foNormalizePriceBasis_(priceBasis, marketValueBasis) {
  const basis = String(priceBasis || '').trim().toUpperCase();
  const marketBasis = String(marketValueBasis || '').trim().toUpperCase();
  if (marketBasis === 'PERSISTED_FALLBACK') return 'PERSISTED_FALLBACK';
  if (basis === 'LIVE' || basis === 'DELAYED' || basis === 'PRIOR_CLOSE' ||
      basis === 'PERSISTED_FALLBACK' || basis === 'ESTIMATED') return basis;
  return 'NOT_AVAILABLE';
}

function foAggregatePriceBasis_(basisSet) {
  const bases = Object.keys(basisSet || {});
  if (!bases.length) return 'NOT_AVAILABLE';
  if (bases.length === 1) return bases[0];
  return 'MIXED';
}

function foParseValuationTimestamp_(value) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function foIsCashTicker_(ticker) {
  const value = String(ticker || '').trim().toUpperCase();
  return value === 'CASH' || value === 'CAD' || value === 'USD' ||
    value === 'CAD.CASH' || value === 'USD.CASH';
}

function foSafeNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  const cleaned = String(value).replace(/\$/g, '').replace(/,/g, '').replace(/%/g, '').trim();
  const number = Number(cleaned);
  return isNaN(number) ? 0 : number;
}

function foIsExcludedValuationRow_(account, ticker, quantity, price) {
  const excludedAccounts = ['', 'N/A', 'NA', 'PENDING', 'REFERENCE', 'LIBRARY', 'WATCHLIST', 'WATCH LIST', 'TEMPLATE'];
  return excludedAccounts.indexOf(account) >= 0 && quantity <= 0 && price <= 0;
}

function foWritePortfolioValuationSummary_(dashboard, result) {
  const sheet = foEnsureSheet_(dashboard, 'Portfolio Valuation Summary', [
    'Timestamp', 'Metric', 'Value', 'Platform Version', 'Baseline'
  ]);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 5).setValues([[
    'Timestamp', 'Metric', 'Value', 'Platform Version', 'Baseline'
  ]]);

  const timestamp = result.valuationTimestamp || new Date();
  const rows = [
    [timestamp, 'Valuation Timestamp', timestamp, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Latest Price Timestamp', result.latestPriceTimestamp || 'NOT AVAILABLE', FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Price Basis', result.portfolioPriceBasis, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Valuation Completeness Status', result.valuationCompletenessStatus, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Valued-Position Market Value', result.totalMarketValue, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Total Market Value', result.totalMarketValue, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Total Cost Basis', result.totalCostBasis, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Comparable Cost Basis', result.comparableCostBasis, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Unrealized Gain/Loss', result.fullPortfolioReturnEligible ? result.unrealizedGainLoss : 'SUPPRESSED', FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Unrealized Gain/Loss %', result.fullPortfolioReturnEligible ? result.unrealizedGainLossPct : 'SUPPRESSED', FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Comparable Unrealized Gain/Loss', result.comparableUnrealizedGainLoss, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Comparable Unrealized Gain/Loss %', result.comparableUnrealizedGainLossPct === null ? 'NOT AVAILABLE' : result.comparableUnrealizedGainLossPct, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Full Portfolio Return Eligible', result.fullPortfolioReturnEligible ? 'YES' : 'NO', FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Valued Positions', result.valuedPositions, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Total Active Positions', result.totalActivePositions, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Missing Price Count', result.missingPriceCount, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Missing Price Tickers', result.missingPriceTickers.join(', ') || 'NONE', FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Missing Cost Basis Count', result.missingCostBasisCount, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Missing Cost Basis Tickers', result.missingCostBasisTickers.join(', ') || 'NONE', FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Price Coverage %', result.priceCoveragePct, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Cost Basis Coverage %', result.costBasisCoveragePct, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Reconciliation Variance', result.reconciliationVariance, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Reconciliation Status', result.reconciliationStatus, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [timestamp, 'Certification Status', result.certificationStatus, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE]
  ];

  result.accountEvidence.forEach(function(account) {
    const label = account.account || 'UNSPECIFIED';
    rows.push([timestamp, label + ' Market Value', account.marketValue, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE]);
    rows.push([timestamp, label + ' Cost Basis', account.costBasis, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE]);
    if (label.indexOf('IBKR') >= 0 || label.indexOf('INTERACTIVE') >= 0) {
      rows.push([timestamp, 'IBKR Cash Included', account.cashIncluded ? 'YES' : 'NO', FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE]);
    }
  });

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
