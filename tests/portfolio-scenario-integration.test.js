'use strict';

const fs = require('fs');
const path = require('path');

const packageRoot = path.resolve(__dirname, '..');
const scenarioSource = fs.readFileSync(
  path.join(packageRoot, 'PortfolioScenarioIntelligence.js'),
  'utf8'
);
const moduleRegistrySource = fs.readFileSync(
  path.join(packageRoot, 'ModuleRegistry.js'),
  'utf8'
);
const orchestratorSource = fs.readFileSync(
  path.join(packageRoot, 'AutonomousCioOrchestrator.js'),
  'utf8'
);
const reportingSource = fs.readFileSync(
  path.join(packageRoot, 'ExecutiveReportingEngine.js'),
  'utf8'
);
const configSource = fs.readFileSync(
  path.join(packageRoot, 'Config.js'),
  'utf8'
);

describe('Sprint 3.1.0 platform integration contracts', () => {
  test('configuration declares scenario worksheets', () => {
    expect(configSource).toContain("PORTFOLIO_SCENARIOS");
    expect(configSource).toContain("PORTFOLIO_SCENARIO_SUMMARY");
  });

  test('module registry exposes setup and execution modules', () => {
    expect(moduleRegistrySource)
      .toContain('PORTFOLIO_SCENARIO_SETUP');
    expect(moduleRegistrySource)
      .toContain('PORTFOLIO_SCENARIO: foRunPortfolioScenarioIntelligence');
  });

  test('orchestrator runs scenario setup before health and comparison after optimization', () => {
    const setupIndex = orchestratorSource.indexOf(
      'Portfolio Scenario Architecture'
    );
    const healthIndex = orchestratorSource.indexOf(
      'Platform Health Check'
    );


   const optimizationIndex = orchestratorSource.indexOf(
     "foGetModule('PORTFOLIO_OPTIMIZATION')"
    );

   const scenarioIndex = orchestratorSource.indexOf(
    "foGetModule('PORTFOLIO_SCENARIO')"
   );

    expect(setupIndex).toBeGreaterThanOrEqual(0);
    expect(setupIndex).toBeLessThan(healthIndex);
    expect(scenarioIndex).toBeGreaterThan(optimizationIndex);
  });

  test('executive reporting surfaces preferred scenario intelligence', () => {
    expect(reportingSource)
      .toContain('foAppendPortfolioScenarioExecutiveRows_');
    expect(reportingSource).toContain('Preferred Scenario');
    expect(reportingSource).toContain('Scenario Recommendation');
  });

  test('scenario module is explicitly non-predictive', () => {
    expect(scenarioSource).toContain('not predictive');
    expect(scenarioSource).toContain('does not estimate investment returns');
  });
});
