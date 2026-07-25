/**
 * PortfolioStateService.js
 * Sprint v3.0.2 — Phase D correction
 *
 * Rebuilds Portfolio State deterministically from Portfolio Master using the
 * governed 22-column Portfolio State schema. Existing legacy headers and rows
 * are replaced to prevent positional shifts, #VALUE!, and #DIV/0! errors.
 */

function foRebuildPortfolioState() {
  const module = 'PortfolioStateService';
  const stateHeaders = [
    'Timestamp',
    'Account',
    'Ticker',
    'Name',
    'Vehicle',
    'Market Access Type',
    'Quantity',
    'Native Currency',
    'Native Price',
    'FX Rate',
    'Market Value CAD',
    'Cost Basis CAD',
    'Unrealized P&L CAD',
    'Unrealized P&L %',
    'Asset Class',
    'Theme',
    'Sector',
    'Target Weight',
    'Current Weight',
    'Drift',
    'Status',
    'Notes'
  ];

  try {
    foInfo_(module, 'Start', 'Portfolio State rebuild started.');

    const dashboard = foDashboard_();
    const source = dashboard.getSheetByName(FO_SHEETS.PORTFOLIO_MASTER);
    if (!source) throw new Error('Portfolio Master worksheet not found.');

    let state = dashboard.getSheetByName(FO_SHEETS.PORTFOLIO_STATE || 'Portfolio State');
    if (!state) state = dashboard.insertSheet(FO_SHEETS.PORTFOLIO_STATE || 'Portfolio State');

    // Enforce one authoritative schema. clearContents removes stale formulas and
    // rows but preserves worksheet formatting and protections.
    state.clearContents();
    state.getRange(1, 1, 1, stateHeaders.length).setValues([stateHeaders]);
    state.setFrozenRows(1);

    const values = source.getDataRange().getValues();
    if (values.length < 2) {
      foWarn_(module, 'No Data', 'Portfolio Master has no active data rows.');
      return { status: 'NO_DATA', rowsWritten: 0, totalMarketValueCAD: 0 };
    }

    const headers = values[0].map(function(header) {
      return String(header || '').trim();
    });

    const records = [];
    let totalMarketValueCAD = 0;

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const ticker = String(foStateValue_(row, headers, ['Ticker', 'Symbol']) || '')
        .trim()
        .toUpperCase();
      if (!ticker) continue;

      const account = String(foStateValue_(row, headers, ['Account']) || '').trim();
      const quantity = foStateNumber_(foStateValue_(row, headers, ['Quantity', 'Shares', 'Units']));
      if (!account || quantity <= 0) continue;

      const nativePrice = foStateNumberOrBlank_(
        foStateValue_(row, headers, ['Current Price', 'Native Price', 'Price'])
      );
      const fxRateRaw = foStateNumberOrBlank_(foStateValue_(row, headers, ['FX Rate']));
      const nativeCurrency = String(
        foStateValue_(row, headers, ['Native Currency', 'Currency']) || 'CAD'
      ).trim().toUpperCase();
      const fxRate = fxRateRaw !== '' ? fxRateRaw : nativeCurrency === 'USD' ? 1 : 1;

      const explicitMarketValue = foStateNumberOrBlank_(
        foStateValue_(row, headers, ['Market Value CAD', 'Market Value'])
      );
      const marketValueCAD = explicitMarketValue !== ''
        ? explicitMarketValue
        : nativePrice !== ''
          ? quantity * nativePrice * fxRate
          : '';
      const costBasisCAD = foStateNumberOrBlank_(
        foStateValue_(row, headers, ['Cost Basis CAD', 'Cost Basis', 'Book Value'])
      );
      const unrealizedCAD = marketValueCAD !== '' && costBasisCAD !== ''
        ? marketValueCAD - costBasisCAD
        : '';
      const unrealizedPct = unrealizedCAD !== '' && costBasisCAD !== 0
        ? unrealizedCAD / costBasisCAD
        : '';
      const targetWeight = foStateNumberOrBlank_(
        foStateValue_(row, headers, ['Target Weight'])
      );

      if (marketValueCAD !== '') totalMarketValueCAD += marketValueCAD;

      records.push({
        Timestamp: new Date(),
        Account: account,
        Ticker: ticker,
        Name: foStateValue_(row, headers, ['Company', 'Company / Fund', 'Name', 'Security']) || '',
        Vehicle: foStateValue_(row, headers, ['Preferred Vehicle', 'Vehicle']) || account,
        'Market Access Type': foStateValue_(row, headers, ['Market Access Type']) || '',
        Quantity: quantity,
        'Native Currency': nativeCurrency,
        'Native Price': nativePrice,
        'FX Rate': fxRate,
        'Market Value CAD': marketValueCAD,
        'Cost Basis CAD': costBasisCAD,
        'Unrealized P&L CAD': unrealizedCAD,
        'Unrealized P&L %': unrealizedPct,
        'Asset Class': foStateValue_(row, headers, ['Asset Class']) || '',
        Theme: foStateValue_(row, headers, ['Investment Theme', 'Theme']) || '',
        Sector: foStateValue_(row, headers, ['Sector']) || '',
        'Target Weight': targetWeight,
        'Current Weight': '',
        Drift: '',
        Status: 'Active',
        Notes: 'Generated from synchronized Portfolio Master by PortfolioStateService'
      });
    }

    records.forEach(function(record) {
      const currentWeight = totalMarketValueCAD > 0 && record['Market Value CAD'] !== ''
        ? record['Market Value CAD'] / totalMarketValueCAD
        : '';
      record['Current Weight'] = currentWeight;
      record.Drift = currentWeight !== '' && record['Target Weight'] !== ''
        ? currentWeight - record['Target Weight']
        : '';
    });

    const output = records.map(function(record) {
      const row = stateHeaders.map(function(header) {
        const value = record[header];
        return value === undefined || value === null ? '' : value;
      });
      if (row.length !== stateHeaders.length) {
        throw new Error(
          'Portfolio State schema mismatch: expected ' + stateHeaders.length +
          ' columns, received ' + row.length
        );
      }
      return row;
    });

    if (output.length > 0) {
      state.getRange(2, 1, output.length, stateHeaders.length).setValues(output);
      state.getRange(2, 7, output.length, 1).setNumberFormat('0.########');
      state.getRange(2, 9, output.length, 5).setNumberFormat('$#,##0.00;-$#,##0.00');
      state.getRange(2, 14, output.length, 1).setNumberFormat('0.00%');
      state.getRange(2, 18, output.length, 3).setNumberFormat('0.00%');
    }

    foInfo_(
      module,
      'Complete',
      output.length + ' Portfolio State rows rebuilt. Total market value CAD: ' + totalMarketValueCAD
    );

    return {
      status: 'SUCCESS',
      rowsWritten: output.length,
      totalMarketValueCAD: totalMarketValueCAD
    };
  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foStateValue_(row, headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const wanted = String(aliases[i]).trim().toUpperCase();
    for (let h = 0; h < headers.length; h++) {
      if (String(headers[h]).trim().toUpperCase() === wanted) return row[h];
    }
  }
  return '';
}

function foStateNumber_(value) {
  const parsed = foStateNumberOrBlank_(value);
  return parsed === '' ? 0 : parsed;
}

function foStateNumberOrBlank_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number') return isFinite(value) ? value : '';
  const normalized = String(value)
    .replace(/C\$/gi, '')
    .replace(/US\$/gi, '')
    .replace(/\$/g, '')
    .replace(/,/g, '')
    .replace(/%/g, '')
    .trim();
  if (!normalized) return '';
  const parsed = Number(normalized);
  return isNaN(parsed) ? '' : parsed;
}
