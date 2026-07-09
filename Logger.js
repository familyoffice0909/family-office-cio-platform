function foLog_(level, module, action, message) {
  const sheet = foEnsureSheet_(foDashboard_(), FO_SHEETS.AUTOMATION_LOG, [
    'Timestamp',
    'Level',
    'Module',
    'Action',
    'Message',
    'Platform Version',
    'Engine Version',
    'Baseline',
    'User'
  ]);

  sheet.appendRow([
    foTimestamp_(),
    level,
    module,
    action,
    message,
    FO_CONFIG.PLATFORM_VERSION,
    FO_CONFIG.ENGINE_VERSION,
    FO_CONFIG.BASELINE,
    foGetActiveUser_()
  ]);
}

function foInfo_(module, action, message) {
  foLog_('INFO', module, action, message);
}

function foWarn_(module, action, message) {
  foLog_('WARNING', module, action, message);
}

function foError_(module, action, error) {
  const message = error && error.stack ? error.stack : String(error);
  foLog_('ERROR', module, action, message);
}