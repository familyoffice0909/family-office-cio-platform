/************************************************************
 * PortfolioExposureAttributionEngine.gs
 * Wave 2.3.2 — Exposure & Attribution Engine
 ************************************************************/

function foRunPortfolioExposureAttribution() {
  const module = 'PortfolioExposureAttributionEngine';

  try {
    foInfo_(module, 'Start', 'Portfolio exposure attribution started.');

    const dashboard = foDashboard_();
    const sheet = dashboard.getSheetByName('Portfolio Performance Positions');

    if (!sheet) {
      throw new Error('Portfolio Performance Positions sheet not found. Run Portfolio Performance first.');
    }

    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(String);

    const positions = foBuildExposurePositions_(values, headers);
    const totalMarketValue = positions.reduce(function(sum, p) {
      return sum + p.marketValue;
    }, 0);

    const accountExposure = foGroupExposure_(positions, 'account', totalMarketValue);
    const sectorExposure = foGroupExposure_(positions, 'sector', totalMarketValue);
    const assetClassExposure = foGroupExposure_(positions, 'assetClass', totalMarketValue);
    const currencyExposure = foGroupExposure_(positions, 'currency', totalMarketValue);

    const concentration = foBuildConcentrationMetrics_(positions, totalMarketValue);

    foWriteExposureSheet_(dashboard, 'Account Exposure', accountExposure);
    foWriteExposureSheet_(dashboard, 'Sector Exposure', sectorExposure);
    foWriteExposureSheet_(dashboard, 'Asset Class Exposure', assetClassExposure);
    foWriteExposureSheet_(dashboard, 'Currency Exposure', currencyExposure);
    foWriteConcentrationSheet_(dashboard, concentration);

    foInfo_(module, 'Complete', 'Portfolio exposure attribution completed.');

    return {
      status: 'SUCCESS',
      positions: positions.length,
      totalMarketValue: totalMarketValue,
      accountGroups: accountExposure.length,
      sectorGroups: sectorExposure.length,
      assetClassGroups: assetClassExposure.length,
      currencyGroups: currencyExposure.length
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foBuildExposurePositions_(values, headers) {
  const positions = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const ticker = String(foGetVal_(row, headers, 'Ticker') || '').trim().toUpperCase();
    if (!ticker) continue;

    const account = String(foGetVal_(row, headers, 'Account') || 'Unknown').trim() || 'Unknown';
    const sector = String(foGetVal_(row, headers, 'Sector') || 'Unknown').trim() || 'Unknown';
    const assetClass = String(foGetVal_(row, headers, 'Asset Class') || 'Unknown').trim() || 'Unknown';
    const marketValue = foExposureNumber_(foGetVal_(row, headers, 'Market Value'));
    const costBasis = foExposureNumber_(foGetVal_(row, headers, 'Cost Basis'));

    if (marketValue <= 0) continue;

    positions.push({
      ticker: ticker,
      account: account,
      sector: sector,
      assetClass: assetClass,
      currency: foInferExposureCurrency_(ticker),
      marketValue: marketValue,
      costBasis: costBasis,
      gainLoss: marketValue - costBasis
    });
  }

  return positions;
}

function foGroupExposure_(positions, field, totalMarketValue) {
  const groups = {};

  positions.forEach(function(p) {
    const key = p[field] || 'Unknown';

    if (!groups[key]) {
      groups[key] = {
        group: key,
        marketValue: 0,
        costBasis: 0,
        gainLoss: 0,
        positionCount: 0,
        tickers: []
      };
    }

    groups[key].marketValue += p.marketValue;
    groups[key].costBasis += p.costBasis;
    groups[key].gainLoss += p.gainLoss;
    groups[key].positionCount++;
    groups[key].tickers.push(p.ticker);
  });

  return Object.keys(groups)
    .map(function(key) {
      const g = groups[key];

      return {
        group: g.group,
        marketValue: g.marketValue,
        costBasis: g.costBasis,
        gainLoss: g.gainLoss,
        returnPct: g.costBasis > 0 ? g.gainLoss / g.costBasis : 0,
        portfolioWeight: totalMarketValue > 0 ? g.marketValue / totalMarketValue : 0,
        positionCount: g.positionCount,
        tickers: g.tickers.join(', ')
      };
    })
    .sort(function(a, b) {
      return b.marketValue - a.marketValue;
    });
}

function foBuildConcentrationMetrics_(positions, totalMarketValue) {
  const sorted = positions.slice().sort(function(a, b) {
    return b.marketValue - a.marketValue;
  });

  const top1 = sorted[0] || null;
  const top3 = sorted.slice(0, 3);
  const top5 = sorted.slice(0, 5);

  const top3Value = top3.reduce(function(sum, p) {
    return sum + p.marketValue;
  }, 0);

  const top5Value = top5.reduce(function(sum, p) {
    return sum + p.marketValue;
  }, 0);

  const top1Weight = totalMarketValue > 0 && top1 ? top1.marketValue / totalMarketValue : 0;
  const top3Weight = totalMarketValue > 0 ? top3Value / totalMarketValue : 0;
  const top5Weight = totalMarketValue > 0 ? top5Value / totalMarketValue : 0;

  let concentrationRisk = 'LOW';

  if (top1Weight >= 0.35 || top3Weight >= 0.65) {
    concentrationRisk = 'HIGH';
  } else if (top1Weight >= 0.25 || top3Weight >= 0.5) {
    concentrationRisk = 'MEDIUM';
  }

  return {
    totalMarketValue: totalMarketValue,
    largestPosition: top1 ? top1.ticker : '',
    largestPositionValue: top1 ? top1.marketValue : 0,
    largestPositionWeight: top1Weight,
    top3Weight: top3Weight,
    top5Weight: top5Weight,
    concentrationRisk: concentrationRisk,
    top3Tickers: top3.map(function(p) { return p.ticker; }).join(', '),
    top5Tickers: top5.map(function(p) { return p.ticker; }).join(', ')
  };
}

function foWriteExposureSheet_(dashboard, sheetName, exposureRows) {
  const sheet = foEnsureSheet_(dashboard, sheetName, [
    'Timestamp',
    'Group',
    'Market Value',
    'Cost Basis',
    'Gain/Loss',
    'Return %',
    'Portfolio Weight',
    'Position Count',
    'Tickers',
    'Platform Version',
    'Baseline'
  ]);

  sheet.clearContents();

  sheet.getRange(1, 1, 1, 11).setValues([[
    'Timestamp',
    'Group',
    'Market Value',
    'Cost Basis',
    'Gain/Loss',
    'Return %',
    'Portfolio Weight',
    'Position Count',
    'Tickers',
    'Platform Version',
    'Baseline'
  ]]);

  const rows = exposureRows.map(function(g) {
    return [
      new Date(),
      g.group,
      g.marketValue,
      g.costBasis,
      g.gainLoss,
      g.returnPct,
      g.portfolioWeight,
      g.positionCount,
      g.tickers,
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 11).setValues(rows);
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 11);
}

function foWriteConcentrationSheet_(dashboard, c) {
  const sheet = foEnsureSheet_(dashboard, 'Portfolio Concentration Summary', [
    'Timestamp',
    'Metric',
    'Value',
    'Notes',
    'Platform Version',
    'Baseline'
  ]);

  sheet.clearContents();

  sheet.getRange(1, 1, 1, 6).setValues([[
    'Timestamp',
    'Metric',
    'Value',
    'Notes',
    'Platform Version',
    'Baseline'
  ]]);

  const rows = [
    [new Date(), 'Total Market Value', c.totalMarketValue, '', FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [new Date(), 'Largest Position', c.largestPositionValue, c.largestPosition, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [new Date(), 'Largest Position Weight', c.largestPositionWeight, c.largestPosition, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [new Date(), 'Top 3 Weight', c.top3Weight, c.top3Tickers, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [new Date(), 'Top 5 Weight', c.top5Weight, c.top5Tickers, FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE],
    [new Date(), 'Concentration Risk', c.concentrationRisk, 'Based on largest position and top 3 concentration.', FO_CONFIG.PLATFORM_VERSION, FO_CONFIG.BASELINE]
  ];

  sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 6);
}

function foInferExposureCurrency_(ticker) {
  const t = String(ticker || '').trim().toUpperCase();

  if (['QBTS', 'RGTI', 'MU', 'AVGO', 'QCOM', 'NVDA', 'META', 'PLTR'].indexOf(t) >= 0) {
    return 'USD';
  }

  return 'CAD';
}

function foExposureNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;

  const cleaned = String(value)
    .replace(/\$/g, '')
    .replace(/,/g, '')
    .replace(/%/g, '')
    .trim();

  const n = Number(cleaned);

  return isNaN(n) ? 0 : n;
}

function foRunPortfolioExposureAttributionSmokeTest() {
  const module = 'PortfolioExposureAttributionEngine';

  try {
    foInfo_(module, 'Start', 'Portfolio Exposure Attribution smoke test started.');

    const result = foRunPortfolioExposureAttribution();

    foInfo_(module, 'Complete', 'Portfolio Exposure Attribution smoke test completed.');

    return result;

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}