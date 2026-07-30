const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'PortfolioValuationEngine.js'), 'utf8');
const context = vm.createContext({ console, Date });
vm.runInContext(source, context, { filename: 'PortfolioValuationEngine.js' });

function sheetStub() {
  return {
    getRange() {
      return { setValue() {}, clearContent() {} };
    }
  };
}

const headers = [
  'Ticker', 'Account', 'Quantity', 'Current Price', 'Market Value', 'Cost Basis',
  'Price Timestamp', 'Price Source', 'Price Status', 'Price Basis',
  'Valuation Status', 'Market Value Basis'
];

test('partial pricing suppresses full-portfolio return and calculates comparable return', () => {
  const values = [headers,
    ['AAA', 'TFSA', 10, 20, '', 100, new Date('2026-07-30T12:00:00Z'), 'Feed', 'OK', 'LIVE', 'VALUED', 'CURRENT_PRICE'],
    ['BBB', 'LIRA', 5, '', '', 80, '', '', 'MISSING', '', 'UNVALUED', '']
  ];
  const result = context.foCalculatePortfolioValuation_(sheetStub(), values, headers);
  expect(result.totalMarketValue).toBe(200);
  expect(result.totalCostBasis).toBe(180);
  expect(result.comparableCostBasis).toBe(100);
  expect(result.comparableUnrealizedGainLoss).toBe(100);
  expect(result.comparableUnrealizedGainLossPct).toBe(1);
  expect(result.fullPortfolioReturnEligible).toBe(false);
  expect(result.unrealizedGainLoss).toBeNull();
  expect(result.unrealizedGainLossPct).toBeNull();
  expect(result.valuationCompletenessStatus).toBe('PARTIAL');
  expect(Array.from(result.missingPriceTickers)).toEqual(['BBB']);
});

test('complete pricing enables full-portfolio return', () => {
  const values = [headers,
    ['AAA', 'TFSA', 10, 20, '', 100, new Date('2026-07-30T12:00:00Z'), 'Feed', 'OK', 'LIVE', 'VALUED', 'CURRENT_PRICE'],
    ['BBB', 'LIRA', 5, 30, '', 80, new Date('2026-07-30T12:05:00Z'), 'Feed', 'OK', 'DELAYED', 'VALUED', 'CURRENT_PRICE']
  ];
  const result = context.foCalculatePortfolioValuation_(sheetStub(), values, headers);
  expect(result.totalMarketValue).toBe(350);
  expect(result.totalCostBasis).toBe(180);
  expect(result.fullPortfolioReturnEligible).toBe(true);
  expect(result.unrealizedGainLoss).toBe(170);
  expect(result.unrealizedGainLossPct).toBeCloseTo(170 / 180);
  expect(result.valuationCompletenessStatus).toBe('COMPLETE');
  expect(result.portfolioPriceBasis).toBe('MIXED');
});

test('account evidence and cash inclusion are preserved', () => {
  const values = [headers,
    ['CASH', 'IBKR', 1000, 1, '', 1000, new Date('2026-07-30T12:00:00Z'), 'Ledger', 'OK', 'LIVE', 'VALUED', 'CURRENT_PRICE']
  ];
  const result = context.foCalculatePortfolioValuation_(sheetStub(), values, headers);
  expect(result.accountEvidence).toHaveLength(1);
  expect(result.accountEvidence[0].account).toBe('IBKR');
  expect(result.accountEvidence[0].marketValue).toBe(1000);
  expect(result.accountEvidence[0].cashIncluded).toBe(true);
});
