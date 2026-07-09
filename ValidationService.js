/************************************************************
 * ValidationService.gs
 * Wave 1C.2 — Data Validation Automation
 ************************************************************/

function foRunDataValidation() {
  const module = 'ValidationService';

  try {
    foInfo_(module, 'Start', 'Data validation started.');

    const dashboard = foDashboard_();
    const results = [];

    foValidateRequiredFields_(
      dashboard,
      FO_SHEETS.PORTFOLIO_MASTER,
      ['Position ID', 'Ticker', 'Account', 'Quantity', 'Current Price'],
      results
    );

    foValidateRequiredFields_(
      dashboard,
      FO_SHEETS.RECOMMENDATION_LEDGER,
      ['Recommendation ID', 'Ticker', 'Recommendation Action', 'Recommendation Status'],
      results
    );

    foValidateRequiredFields_(
      dashboard,
      FO_SHEETS.RECOMMENDATION_PERFORMANCE,
      ['Recommendation ID', 'Ticker'],
      results
    );

    foValidateRequiredFields_(
      dashboard,
      FO_SHEETS.PORTFOLIO_ATTRIBUTION,
      ['Attribution ID', 'Ticker'],
      results
    );

    foValidateRequiredFields_(
      dashboard,
      FO_SHEETS.KNOWLEDGE_BASE,
      ['Knowledge ID', 'Knowledge Category', 'Lesson Learned', 'Knowledge Status'],
      results
    );

    const outputSheet = foEnsureSheet_(dashboard, 'Data Validation Results', [
      'Timestamp',
      'Severity',
      'Worksheet',
      'Row',
      'Field',
      'Issue',
      'Status',
      'Platform Version',
      'Baseline'
    ]);

    if (outputSheet.getLastRow() > 1) {
      outputSheet.getRange(2, 1, outputSheet.getLastRow() - 1, 9).clearContent();
    }

    if (results.length > 0) {
      outputSheet.getRange(2, 1, results.length, 9).setValues(results);
    }

    const status = results.length === 0 ? 'PASS' : 'PASS WITH OBSERVATIONS';

    foInfo_(module, 'Complete', 'Data validation completed. Findings: ' + results.length);

    return {
      status: status,
      findings: results.length
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foValidateRequiredFields_(spreadsheet, sheetName, requiredFields, results) {
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    results.push([
      new Date(),
      'CRITICAL',
      sheetName,
      '',
      '',
      'Worksheet missing.',
      'Open',
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE
    ]);
    return;
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    results.push([
      new Date(),
      'LOW',
      sheetName,
      '',
      '',
      'No data rows found.',
      'Open',
      FO_CONFIG.PLATFORM_VERSION,
      FO_CONFIG.BASELINE
    ]);
    return;
  }

  const headers = values[0].map(String);

  requiredFields.forEach(function(field) {
    if (headers.indexOf(field) < 0) {
      results.push([
        new Date(),
        'HIGH',
        sheetName,
        '',
        field,
        'Required field missing from schema.',
        'Open',
        FO_CONFIG.PLATFORM_VERSION,
        FO_CONFIG.BASELINE
      ]);
    }
  });

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const rowHasData = row.some(function(cell) {
      return cell !== '' && cell !== null;
    });

    if (!rowHasData) continue;

    requiredFields.forEach(function(field) {
      const index = headers.indexOf(field);
      if (index < 0) return;

      const value = row[index];

      if (value === '' || value === null || value === undefined) {
        results.push([
          new Date(),
          'MEDIUM',
          sheetName,
          r + 1,
          field,
          'Required value is blank.',
          'Open',
          FO_CONFIG.PLATFORM_VERSION,
          FO_CONFIG.BASELINE
        ]);
      }
    });
  }
}