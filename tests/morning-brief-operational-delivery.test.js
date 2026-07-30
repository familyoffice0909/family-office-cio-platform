'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const read = (file) => fs.readFileSync(path.join(root, file), 'utf8');

describe('v3.2.5 Morning Brief operational delivery', () => {
  test('uses v3.2.5 as the single configured version authority', () => {
    const config = read('Config.js');
    const packageJson = JSON.parse(read('package.json'));
    const packageLock = JSON.parse(read('package-lock.json'));

    expect(config).toContain("PLATFORM_VERSION: 'v3.2.5'");
    expect(config).toContain("ENGINE_VERSION: 'v3.2.5'");
    expect(config).toContain(
      "RELEASE_NAME: 'Morning Brief Operational Delivery & Version Authority Correction'"
    );
    expect(packageJson.version).toBe('3.2.5');
    expect(packageLock.version).toBe('3.2.5');
    expect(packageLock.packages[''].version).toBe('3.2.5');
  });

  test('report, dashboard, archive, and orchestration derive version from FO_CONFIG', () => {
    const reporting = read('ExecutiveReportingEngine.js');
    const dashboard = read('ExecutiveDashboardEngine.js');
    const persistence = read('ExecutiveReportPersistenceService.js');

    expect(reporting).toContain('FO_CONFIG.PLATFORM_VERSION');
    expect(reporting).toContain("foEnsureSheet_(dashboard, 'Executive Report Archive'");
    expect(dashboard).toContain(
      "['Platform Version', FO_CONFIG.PLATFORM_VERSION"
    );
    expect(persistence).toContain(
      'foAppendExecutiveReportOrchestrationV324_'
    );
    expect(persistence).toContain('FO_CONFIG.PLATFORM_VERSION');

    [
      'Config.js',
      'ExecutiveReportingEngine.js',
      'ExecutiveDashboardEngine.js',
      'ExecutiveReportPersistenceService.js',
      'SchedulerService.js'
    ].forEach((file) => {
      expect(read(file)).not.toContain("'v3.2.1'");
    });
  });

  test('writes reconciliable success evidence to the Ledger Orchestration Log', () => {
    const appendRow = jest.fn();
    const orchestrationSheet = { appendRow };
    const ledger = {};

    const context = vm.createContext({
      console,
      Date,
      FO_CONFIG: {
        PLATFORM_VERSION: 'v3.2.5'
      },
      FO_SHEETS: {
        ORCHESTRATION_LOG: 'Orchestration Log'
      },
      foLedger_: jest.fn(() => ledger),
      foEnsureSheet_: jest.fn(() => orchestrationSheet)
    });

    vm.runInContext(read('ExecutiveReportPersistenceService.js'), context);

    const runId = context.foAppendExecutiveReportOrchestrationV324_(
      'EXEC-RPT-TEST',
      'DIRECT',
      'SNAPSHOT-TEST'
    );

    expect(runId).toBe('EXEC-RPT-TEST');
    expect(context.foEnsureSheet_).toHaveBeenCalledWith(
      ledger,
      'Orchestration Log',
      [
        'Timestamp',
        'Run ID',
        'Channel',
        'Action',
        'Status',
        'Message',
        'Version'
      ]
    );

    const row = appendRow.mock.calls[0][0];
    expect(row[1]).toBe('EXEC-RPT-TEST');
    expect(row[2]).toBe('Morning Brief');
    expect(row[3]).toBe('Generate and persist executive report');
    expect(row[4]).toBe('SUCCESS');
    expect(row[5]).toContain('Snapshot ID: SNAPSHOT-TEST');
    expect(row[6]).toBe('v3.2.5');
  });

  test('runs the production-ready wrapper on weekdays and skips weekends', () => {
    function createScheduler(day) {
      const context = vm.createContext({
        console,
        Date,
        Utilities: {
          formatDate: jest.fn(() => String(day))
        },
        foInfo_: jest.fn(),
        foError_: jest.fn(),
        foRunExecutiveReportProductionReady: jest.fn(() => ({
          status: 'SUCCESS'
        })),
        ScriptApp: {
          getProjectTriggers: jest.fn(() => [])
        }
      });

      vm.runInContext(read('SchedulerService.js'), context);
      return context;
    }

    const weekday = createScheduler(3);
    const weekend = createScheduler(6);

    expect(weekday.foRunScheduledMorningBrief()).toEqual({
      status: 'SUCCESS'
    });
    expect(weekday.foRunExecutiveReportProductionReady)
      .toHaveBeenCalledTimes(1);

    expect(weekend.foRunScheduledMorningBrief()).toEqual({
      status: 'SKIPPED',
      reason: 'WEEKEND'
    });
    expect(weekend.foRunExecutiveReportProductionReady)
      .not.toHaveBeenCalled();
  });

  test('installs the governed production-ready Morning Brief handler', () => {
    const scheduler = read('SchedulerService.js');

    expect(scheduler).toContain(
      "FO_MORNING_BRIEF_TRIGGER_HANDLER_ = 'foRunScheduledMorningBrief'"
    );
    expect(scheduler).toContain('foRunExecutiveReportProductionReady()');
    expect(scheduler).toContain('.atHour(7)');
    expect(scheduler).toContain('.nearMinute(45)');
    expect(scheduler).toContain(
      ".inTimezone(FO_MORNING_BRIEF_TIMEZONE_)"
    );
  });
});
