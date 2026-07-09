function foBootstrap() {
  const module = 'Bootstrap';

  try {
    foInfo_(module, 'Start', 'Bootstrap started.');

    foEnsureSheet_(foDashboard_(), FO_SHEETS.AUTOMATION_LOG, [
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

    foEnsureSheet_(foDashboard_(), FO_SHEETS.PLATFORM_HEALTH, [
      'Timestamp',
      'Check',
      'Status',
      'Details',
      'Platform Version',
      'Baseline'
    ]);

    const health = foRunPlatformHealthCheck();

    foInfo_(
      module,
      'Complete',
      'Bootstrap completed. Health status: ' + health.status
    );

    return {
      status: 'SUCCESS',
      version: foGetVersion(),
      health: health
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}