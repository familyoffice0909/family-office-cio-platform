const FO_MORNING_BRIEF_TRIGGER_HANDLER_ = 'foRunScheduledMorningBrief';
const FO_MORNING_BRIEF_TIMEZONE_ = 'America/Toronto';

function foListTriggers() {
  const module = 'SchedulerService';

  try {
    foInfo_(module, 'Start', 'Listing triggers.');

    const triggers = ScriptApp.getProjectTriggers();

    const result = triggers.map(function(trigger) {
      return {
        handlerFunction: trigger.getHandlerFunction(),
        eventType: trigger.getEventType(),
        source: trigger.getTriggerSource()
      };
    });

    foInfo_(module, 'Complete', result.length + ' triggers found.');

    return result;

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foDeleteAllTriggers() {
  const module = 'SchedulerService';

  try {
    foInfo_(module, 'Start', 'Deleting all triggers.');

    const triggers = ScriptApp.getProjectTriggers();

    triggers.forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
    });

    foInfo_(module, 'Complete', triggers.length + ' triggers deleted.');

    return {
      status: 'SUCCESS',
      deleted: triggers.length
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foRunScheduledMorningBrief() {
  const module = 'SchedulerService';
  const weekday = Number(
    Utilities.formatDate(
      new Date(),
      FO_MORNING_BRIEF_TIMEZONE_,
      'u'
    )
  );

  if (weekday > 5) {
    foInfo_(module, 'Morning Brief', 'Scheduled Morning Brief skipped on weekend.');
    return {
      status: 'SKIPPED',
      reason: 'WEEKEND'
    };
  }

  return foRunExecutiveReportProductionReady();
}

function foInstallMorningBriefTrigger() {
  const module = 'SchedulerService';

  try {
    const existing = ScriptApp.getProjectTriggers().filter(function(trigger) {
      return trigger.getHandlerFunction() ===
        FO_MORNING_BRIEF_TRIGGER_HANDLER_;
    });

    existing.forEach(function(trigger) {
      ScriptApp.deleteTrigger(trigger);
    });

    const trigger = ScriptApp
      .newTrigger(FO_MORNING_BRIEF_TRIGGER_HANDLER_)
      .timeBased()
      .atHour(7)
      .nearMinute(45)
      .everyDays(1)
      .inTimezone(FO_MORNING_BRIEF_TIMEZONE_)
      .create();

    foInfo_(
      module,
      'Morning Brief Trigger',
      'Installed weekday Morning Brief trigger near 07:45 America/Toronto.'
    );

    return {
      status: 'SUCCESS',
      handlerFunction: FO_MORNING_BRIEF_TRIGGER_HANDLER_,
      timezone: FO_MORNING_BRIEF_TIMEZONE_,
      schedule: 'Daily near 07:45; weekdays enforced at runtime.',
      triggerId:
        trigger && typeof trigger.getUniqueId === 'function'
          ? trigger.getUniqueId()
          : ''
    };
  } catch (error) {
    foError_(module, 'Morning Brief Trigger Failure', error);
    throw error;
  }
}

