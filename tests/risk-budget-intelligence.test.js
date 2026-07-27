'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'RiskBudgetIntelligence.js'), 'utf8');
const context = vm.createContext({console});
vm.runInContext(source, context, {filename: 'RiskBudgetIntelligence.js'});
function allocation(overrides) { return Object.assign({
  scenarioRank: 1, scenarioId: 'SCN-OPTIMIZED', scenarioName: 'Optimized',
  ticker: 'QQC', account: 'TFSA', deploymentDecision: 'DEPLOY NOW',
  currentWeight: 0.05, proposedIncrementalWeight: 0.04, proposedTargetWeight: 0.09,
  maximumPositionWeight: 0.15, upstreamConstraintStatus: 'PASS',
  upstreamConstraintReason: 'NONE', riskDisciplineScore: 80,
  constraintComplianceScore: 100
}, overrides || {}); }
const summary = {portfolioRiskLevel:'MEDIUM', riskDisciplineScore:80, constraintComplianceScore:100, upstreamBreachCount:0};
describe('Sprint 3.2.0 Risk Budget Intelligence', () => {
  test('keeps a compliant allocation within budget', () => {
    const result = context.foEvaluateRiskBudget_([allocation()], summary);
    expect(result.assessments[0].budgetStatus).toBe('WITHIN BUDGET');
    expect(result.summary.overallStatus).toBe('WITHIN BUDGET');
    expect(result.summary.breachCount).toBe(0);
  });
  test('flags target weight above governed maximum as breach', () => {
    const result = context.foEvaluateRiskBudget_([allocation({proposedTargetWeight:0.16})], summary);
    expect(result.assessments[0].budgetStatus).toBe('BREACH');
    expect(result.summary.overallStatus).toBe('BREACH');
  });
  test('preserves upstream constraint authority', () => {
    const result = context.foEvaluateRiskBudget_([allocation({upstreamConstraintStatus:'BLOCKED', upstreamConstraintReason:'Upstream limit'})], summary);
    expect(result.assessments[0].budgetStatus).toBe('BREACH');
    expect(result.assessments[0].breachReason).toBe('Upstream limit');
  });
  test('marks allocations at ninety percent utilization as constrained', () => {
    const result = context.foEvaluateRiskBudget_([allocation({proposedTargetWeight:0.135})], summary);
    expect(result.assessments[0].budgetStatus).toBe('CONSTRAINED');
  });
  test('uses Scenario Rank governed header', () => { expect(source).toContain("'Scenario Rank'"); expect(source).not.toContain("headers, 'Rank'"); });
  test('is explicitly non-predictive', () => { expect(source).toContain('does not forecast returns'); expect(source).toContain('not predictive'); });
});
