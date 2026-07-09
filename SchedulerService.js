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