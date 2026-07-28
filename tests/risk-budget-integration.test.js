'use strict';
const fs = require('fs');
const path = require('path');
const root = path.resolve(__dirname, '..');
const read = (name) => fs.readFileSync(path.join(root, name), 'utf8');
const config = read('Config.js');
const registry = read('ModuleRegistry.js');
const orchestrator = read('AutonomousCioOrchestrator.js');
const schemas = read('WorksheetSchemaRegistry.js');
const reporting = read('ExecutiveReportingEngine.js');
describe('Sprint 3.2.1 integration contracts', () => {
  test('configuration declares both governed worksheets', () => {
    expect(config).toContain("RISK_BUDGET_ASSESSMENT: 'Risk Budget Assessment'");
    expect(config).toContain("RISK_BUDGET_SUMMARY: 'Risk Budget Summary'");
  });
  test('worksheet registry registers both exact contracts', () => {
    expect(schemas).toContain('FO_SHEETS.RISK_BUDGET_ASSESSMENT');
    expect(schemas).toContain('FO_SHEETS.RISK_BUDGET_SUMMARY');
    expect(schemas).toContain("'Risk Budget Utilization'");
    expect(schemas).toContain("schemaVersion: '1.1'");
    expect(schemas).toContain("'Primary Blocker'");
    expect(schemas).toContain("'Executive Summary'");
  });
  test('module registry exposes setup and execution modules', () => {
    expect(registry).toContain('RISK_BUDGET_SETUP: foSetupRiskBudgetIntelligence');
    expect(registry).toContain('RISK_BUDGET: foRunRiskBudgetIntelligence');
  });
  test('setup runs before health and execution runs after scenario', () => {
    const setup = orchestrator.indexOf("foGetModule('RISK_BUDGET_SETUP')");
    const health = orchestrator.indexOf("foGetModule('HEALTH')");
    const scenario = orchestrator.indexOf("foGetModule('PORTFOLIO_SCENARIO')");
    const riskBudget = orchestrator.indexOf("foGetModule('RISK_BUDGET')");
    expect(setup).toBeGreaterThanOrEqual(0); expect(setup).toBeLessThan(health);
    expect(riskBudget).toBeGreaterThan(scenario);
  });
  test('executive reporting surfaces status utilization breaches and directive', () => {
    expect(reporting).toContain('foAppendRiskBudgetExecutiveRows_');
    expect(reporting).toContain('Portfolio Budget Utilization');
    expect(reporting).toContain('Breach Count');
    expect(reporting).toContain('Primary Blocker');
    expect(reporting).toContain('Blocked Positions');
    expect(reporting).toContain('Executive Directive');
  });
});
