function foRunLedgerIntegrityCheck() {
  const module = 'IntegrityService';

  try {
    foInfo_(module, 'Start', 'Ledger integrity check started.');

    const ledger = foLedger_();
    const rec = ledger.getSheetByName('Recommendations');

    let status = 'PASS';
    const issues = [];

    if (!rec) {
      status = 'FAIL';
      issues.push('Recommendations tab missing.');
    } else {
      const data = rec.getDataRange().getValues();
      const headers = data[0] || [];

      if (headers.indexOf('Recommendation ID') < 0) {
        status = 'FAIL';
        issues.push('Recommendation ID column missing.');
      }

      if (headers.indexOf('Event ID') < 0) {
        status = 'FAIL';
        issues.push('Event ID column missing.');
      }

      if (headers.indexOf('Planned Review Trigger') < 0) {
        status = 'FAIL';
        issues.push('Planned Review Trigger column missing.');
      }

      if (data.length <= 1) {
        issues.push('No recommendation events recorded yet.');
      }
    }

    const sheet = foEnsureSheet_(ledger, 'Data Integrity', [
      'Timestamp',
      'Check Type',
      'Status',
      'Details',
      'Resolved?'
    ]);

    sheet.appendRow([
      new Date(),
      'Ledger Integrity Check',
      status,
      issues.join(' | ') || 'No structural issues detected.',
      status === 'PASS' ? 'Yes' : 'No'
    ]);

    foInfo_(module, 'Complete', 'Ledger integrity check status: ' + status);

    return {
      status: status,
      issues: issues
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foRunPlatformIntegrityCheck() {
  const module = 'IntegrityService';

  try {
    foInfo_(module, 'Start', 'Platform integrity check started.');

    const health = foRunPlatformHealthCheck();
    const ledger = foRunLedgerIntegrityCheck();

    const status =
      health.status === 'PASS' && ledger.status === 'PASS'
        ? 'PASS'
        : 'REVIEW';

    foInfo_(module, 'Complete', 'Platform integrity check completed: ' + status);

    return {
      status: status,
      health: health,
      ledger: ledger
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}