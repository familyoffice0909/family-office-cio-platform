const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');

function read(name) {
  return fs.readFileSync(path.join(root, name), 'utf8');
}

describe('Sprint v3.0.2 portfolio synchronization integration', () => {
  test('registers synchronization and Portfolio State producers', () => {
    const source = read('ModuleRegistry.js');
    expect(source).toContain(
      'PORTFOLIO_SYNCHRONIZATION: foRunPortfolioDataSynchronization'
    );
    expect(source).toContain('PORTFOLIO_STATE: foRebuildPortfolioState');
  });

  test('runs synchronization before market data and valuation', () => {
    const source = read('AutonomousCioOrchestrator.js');
    const sync = source.indexOf("'Portfolio Data Synchronization'");
    const marketData = source.indexOf("'Market Data Refresh'");
    const valuation = source.indexOf("'Portfolio Valuation'");
    const state = source.indexOf("'Portfolio State Rebuild'");
    const integrity = source.indexOf("'Portfolio Data Integrity'");

    expect(sync).toBeGreaterThan(-1);
    expect(sync).toBeLessThan(marketData);
    expect(marketData).toBeLessThan(valuation);
    expect(valuation).toBeLessThan(state);
    expect(state).toBeLessThan(integrity);
  });

  test('rebuilds Portfolio Master from governed operational holdings sheets', () => {
    const source = read('PortfolioDataSynchronizationService.js');
    expect(source).toContain("sheetName: 'TFSA Holdings'");
    expect(source).toContain("sheetName: 'LIRA Holdings'");
    expect(source).toContain("sheetName: 'Interactive Brokers'");
    expect(source).toContain('foReplacePortfolioMasterRows_');
    expect(source).not.toContain("sheetName: 'Watchlists'");
  });

  test('fails closed when required source columns cannot be resolved', () => {
    const source = read('PortfolioDataSynchronizationService.js');
    expect(source).toContain(
      "' must contain ticker, quantity and cost-basis columns. Resolved: '"
    );
  });

  test('counts active cost basis even when current price is unavailable', () => {
    const source = read('PortfolioValuationEngine.js');
    const costBasis = source.indexOf(
      'totalCostBasis += foSafeNumber_(costBasis);'
    );
    const missingPrice = source.indexOf('if (price <= 0)');

    expect(costBasis).toBeGreaterThan(-1);
    expect(missingPrice).toBeGreaterThan(-1);
    expect(costBasis).toBeLessThan(missingPrice);
  });

  test('updates the governed release metadata', () => {
    const config = read('Config.js');
    const packageJson = JSON.parse(read('package.json'));
    expect(config).toContain("PLATFORM_VERSION: 'v3.0.2'");
    expect(packageJson.version).toBe('3.0.2');
  });
});
