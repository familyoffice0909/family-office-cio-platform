/************************************************************
 * AutonomousCioOrchestrator.gs
 * Wave 2.2.2
 ************************************************************/

function foRunAutonomousCioOrchestrator() {

  const module = 'AutonomousCioOrchestrator';

  const runId = foNowId_('CIO-RUN');

  const startedAt = new Date();

  try {

    foInfo_(module,'Start','Autonomous CIO started.');

    const steps = [];

    // Platform

    steps.push(
      foRunOrchestratorStep_(runId,'Platform Health',foGetModule('HEALTH'))
    );

    steps.push(
      foRunOrchestratorStep_(runId,'Platform Integrity',foGetModule('INTEGRITY'))
    );

    steps.push(
      foRunOrchestratorStep_(runId,'Data Validation',foGetModule('VALIDATION'))
    );

    // Market

    steps.push(
      foRunOrchestratorStep_(runId,'Market Data Refresh',foGetModule('MARKET_DATA'))
    );

    // Portfolio

    steps.push(
      foRunOrchestratorStep_(runId,'Portfolio Valuation',foGetModule('VALUATION'))
    );

    steps.push(
      foRunOrchestratorStep_(runId,'Portfolio Data Integrity',foGetModule('PORTFOLIO_DATA_INTEGRITY'))
    );

    steps.push(
      foRunOrchestratorStep_(runId,'Portfolio Snapshot',foGetModule('PORTFOLIO'))
    );

    // Intelligence

    steps.push(
      foRunOrchestratorStep_(runId,'Market Intelligence',foGetModule('MARKET'))
    );

    steps.push(
      foRunOrchestratorStep_(runId,'CIO Decision Engine',foGetModule('CIO'))
    );

    // Reporting

    steps.push(
      foRunOrchestratorStep_(runId,'Executive Report',foGetModule('REPORT'))
    );

    steps.push(
      foRunOrchestratorStep_(runId,'Executive Dashboard',foGetModule('DASHBOARD'))
    );

    const summary =
      foBuildOrchestratorSummary_(steps,startedAt);

    foWriteOrchestratorRunLog_(
      runId,
      startedAt,
      summary,
      steps
    );

    foInfo_(
      module,
      'Complete',
      'Autonomous CIO completed.'
    );

    return {

      status:summary.status,

      runId:runId,

      successfulSteps:summary.successfulSteps,

      failedSteps:summary.failedSteps,

      durationSeconds:summary.durationSeconds

    };

  }

  catch(error){

    foError_(module,'Failure',error);

    throw error;

  }

}

function foRunAutonomousCioOrchestratorSmokeTest(){

    return foRunAutonomousCioOrchestrator();

}