'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(
  path.join(root, 'PortfolioScenarioIntelligence.js'),
  'utf8'
);
const context = vm.createContext({console});
vm.runInContext(source, context, {
  filename: 'PortfolioScenarioIntelligence.js'
});

function candidate(overrides) {
  return Object.assign({
    rank: 1,
    ticker: 'QQC',
    account: 'TFSA',
    deploymentDecision: 'DEPLOY NOW',
    deploymentScore: 90,
    currentPortfolioWeight: 0.05,
    optimizedIncrementalWeight: 0.05,
    optimizedTargetWeight: 0.10,
    maximumPositionWeight: 0.15,
    constraintStatus: 'PASS',
    constraintReason: 'NONE'
  }, overrides || {});
}

function normalContext(overrides) {
  const base = {
    risk: {
      riskScore: 35,
      diversificationScore: 60,
      largestPositionWeight: 0.18,
      topFiveWeight: 0.60,
      stressTestScore: 45,
      overallRisk: 'MEDIUM',
      critical: false
    },
    stress: {
      enabledScenarioCount: 6,
      stressPressureScore: 55,
      highestSeverity: 'CRITICAL',
      scenarioNames: []
    }
  };
  return Object.assign(base, overrides || {});
}

describe('Sprint 3.1.0 portfolio scenario intelligence', () => {
  test('builds five deterministic scenarios', () => {
    const result = context.foBuildPortfolioScenarios_(
      [
        candidate(),
        candidate({
          rank: 2,
          ticker: 'NVDA',
          currentPortfolioWeight: 0.08,
          optimizedIncrementalWeight: 0.03,
          optimizedTargetWeight: 0.11,
          deploymentScore: 85
        }),
        candidate({
          rank: 3,
          ticker: 'TD',
          currentPortfolioWeight: 0.03,
          optimizedIncrementalWeight: 0.02,
          optimizedTargetWeight: 0.05,
          deploymentScore: 75
        })
      ],
      normalContext()
    );

    expect(result.scenarios).toHaveLength(5);
    expect(result.scenarios[0].rank).toBe(1);
    expect(result.scenarios.filter((item) => item.preferred))
      .toHaveLength(1);
  });

  test('never allocates above the upstream optimized increment', () => {
    const result = context.foBuildPortfolioScenarios_(
      [candidate()],
      normalContext()
    );

    result.scenarios.forEach((scenario) => {
      scenario.allocations.forEach((allocation) => {
        expect(allocation.proposedIncrementalWeight)
          .toBeLessThanOrEqual(
            allocation.optimizedIncrementalWeight
          );
        expect(allocation.proposedTargetWeight)
          .toBeLessThanOrEqual(allocation.maximumPositionWeight);
      });
    });
  });

  test('prefers defensive allocation when risk is critical', () => {
    const critical = normalContext({
      risk: {
        riskScore: 90,
        diversificationScore: 45,
        largestPositionWeight: 0.30,
        topFiveWeight: 0.88,
        stressTestScore: 90,
        overallRisk: 'CRITICAL',
        critical: true
      },
      stress: {
        enabledScenarioCount: 6,
        stressPressureScore: 88,
        highestSeverity: 'CRITICAL',
        scenarioNames: []
      }
    });
    const result = context.foBuildPortfolioScenarios_(
      [candidate()],
      critical
    );

    expect(result.scenarios[0].scenarioId).toBe('SCN-DEFENSIVE');
    expect(result.scenarios[0].totalIncrementalWeight).toBe(0);
  });

  test('focused scenario limits deployment to at most two candidates', () => {
    const result = context.foBuildPortfolioScenarios_(
      [
        candidate({ticker: 'A'}),
        candidate({rank: 2, ticker: 'B'}),
        candidate({rank: 3, ticker: 'C'})
      ],
      normalContext()
    );
    const focused = result.scenarios.find(
      (scenario) => scenario.scenarioId === 'SCN-FOCUSED'
    );

    expect(focused.fundedCandidateCount).toBeLessThanOrEqual(2);
  });

  test('output contract contains governed scenario fields', () => {
    expect(source).toContain("'Scenario Score'");
    expect(source).toContain("'Preferred'");
    expect(source).toContain("'Executive Recommendation'");
    expect(source).toContain("'Scenario Rationale'");
  });
});
