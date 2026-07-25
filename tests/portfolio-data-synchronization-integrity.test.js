const fs = require('fs');
const path = require('path');
const root = path.resolve(__dirname, '..');
const read = name => fs.readFileSync(path.join(root, name), 'utf8');

describe('Sprint v3.0.2 Phase D portfolio integrity corrections', () => {
  test('synchronization fails closed before replacing Portfolio Master', () => {
    const source = read('PortfolioDataSynchronizationService.js');
    const assertion = source.indexOf('foAssertRequiredPortfolioSources_');
    const replacement = source.indexOf('foReplacePortfolioMasterRows_');
    expect(assertion).toBeGreaterThan(-1);
    expect(replacement).toBeGreaterThan(assertion);
    expect(source).toContain('Portfolio Master was not modified.');
  });

  test('TFSA and LIRA are required operational sources', () => {
    const source = read('PortfolioDataSynchronizationService.js');
    expect(source).toContain("sheetName: 'TFSA Holdings', account: 'TFSA', required: true");
    expect(source).toContain("sheetName: 'LIRA Holdings', account: 'LIRA', required: true");
  });

  test('Portfolio State enforces the governed 22-column schema', () => {
    const source = read('PortfolioStateService.js');
    expect(source).toContain("'Market Value CAD'");
    expect(source).toContain("'Cost Basis CAD'");
    expect(source).toContain("'Current Weight'");
    expect(source).toContain('state.clearContents()');
    expect(source).toContain('stateHeaders.map');
    expect(source).toContain('Portfolio State schema mismatch');
  });

  test('provides a one-time reconciled holdings bootstrap', () => {
    const source = read('PortfolioHoldingsBootstrapV302.js');
    expect(source).toContain('function foBootstrapReconciledHoldingsV302()');
    expect(source).toContain("['QQC', 'Invesco NASDAQ 100 Index ETF CAD Hedged', 555, 27578.82");
    expect(source).toContain("['QNC', 'Quantum eMotion', 7697, 13112.45");
    expect(source).toContain("['ONE', '01 Quantum', 8698, 10307.80");
  });

  test('backup folders are excluded from clasp deployment', () => {
    const source = read('.claspignore');
    expect(source).toContain('.sprint-*-backup/**');
  });
});
