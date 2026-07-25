/**
 * PortfolioHoldingsBootstrapV302.js
 * One-time governed bootstrap for the reconciled v3.0.2 operational holdings.
 *
 * Run foBootstrapReconciledHoldingsV302() once before synchronization.
 * It replaces the data rows in TFSA Holdings and LIRA Holdings with the
 * approved reconciliation baseline. It does not run from the orchestrator.
 */

function foBootstrapReconciledHoldingsV302() {
  const module = 'PortfolioHoldingsBootstrapV302';
  const dashboard = foDashboard_();
  const headers = [
    'Ticker',
    'Company',
    'Quantity',
    'Cost Basis CAD',
    'Current Price',
    'Market Value CAD',
    'Native Currency',
    'Asset Class',
    'Theme',
    'Sector',
    'Status',
    'Notes'
  ];

  const tfsaRows = [
    ['QQC', 'Invesco NASDAQ 100 Index ETF CAD Hedged', 555, 27578.82, '', '', 'CAD', 'Equity ETF', 'AI / Nasdaq', 'Technology', 'Active', 'Reconciled baseline v3.0.2']
  ];

  const liraRows = [
    ['QNC', 'Quantum eMotion', 7697, 13112.45, '', '', 'CAD', 'Equity', 'Quantum', 'Technology', 'Active', 'Reconciled baseline v3.0.2'],
    ['BNS', 'Bank of Nova Scotia', 105, 9906.16, '', '', 'CAD', 'Equity', 'Canadian Banks', 'Financials', 'Active', 'Reconciled baseline v3.0.2'],
    ['TD', 'Toronto-Dominion Bank', 47, 4902.94, '', '', 'CAD', 'Equity', 'Canadian Banks', 'Financials', 'Active', 'Reconciled baseline v3.0.2'],
    ['ABX', 'Barrick Mining', 155, 10731.40, '', '', 'CAD', 'Equity', 'Gold', 'Materials', 'Active', 'Reconciled baseline v3.0.2'],
    ['ONE', '01 Quantum', 8698, 10307.80, '', '', 'CAD', 'Equity', 'Quantum', 'Technology', 'Active', 'Reconciled baseline v3.0.2'],
    ['QQC', 'Invesco NASDAQ 100 Index ETF CAD Hedged', 10, 506.75, '', '', 'CAD', 'Equity ETF', 'AI / Nasdaq', 'Technology', 'Active', 'Reconciled baseline v3.0.2']
  ];

  foWriteReconciledHoldingsV302_(dashboard, 'TFSA Holdings', headers, tfsaRows);
  foWriteReconciledHoldingsV302_(dashboard, 'LIRA Holdings', headers, liraRows);

  const tfsaCost = tfsaRows.reduce(function(sum, row) { return sum + Number(row[3]); }, 0);
  const liraCost = liraRows.reduce(function(sum, row) { return sum + Number(row[3]); }, 0);

  foInfo_(module, 'Complete', 'TFSA and LIRA operational holdings bootstrapped.');
  return {
    status: 'SUCCESS',
    tfsaRows: tfsaRows.length,
    tfsaCostBasisCAD: tfsaCost,
    liraRows: liraRows.length,
    liraCostBasisCAD: liraCost
  };
}

function foWriteReconciledHoldingsV302_(dashboard, sheetName, headers, rows) {
  let sheet = dashboard.getSheetByName(sheetName);
  if (!sheet) sheet = dashboard.insertSheet(sheetName);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet.getRange(2, 3, Math.max(rows.length, 1), 1).setNumberFormat('0.########');
  sheet.getRange(2, 4, Math.max(rows.length, 1), 3).setNumberFormat('$#,##0.00;-$#,##0.00');
}
