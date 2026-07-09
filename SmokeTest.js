function foRunModularSmokeTest() {
  const module = 'SmokeTest';

  try {
    foInfo_(module, 'Start', 'Modular smoke test started.');

    const bootstrap = foBootstrap();
    const health = foRunPlatformHealthCheck();
    const integrity = foRunPlatformIntegrityCheck();

    foAppendSampleRecommendation();

    foArchiveSmokeTestReport();

    foInfo_(module, 'Complete', 'Modular smoke test completed.');

    return {
      status: 'SUCCESS',
      bootstrap: bootstrap.status,
      health: health.status,
      integrity: integrity.status
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}