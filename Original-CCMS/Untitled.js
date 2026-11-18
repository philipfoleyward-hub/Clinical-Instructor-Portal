/**
 * CCMS read-only Professional Competency submission deployment diagnostic.
 * Install this file alone even if the V2 candidate is not yet present.
 * Run ONLY qaPcSubmissionWiringReadOnly().
 * It never modifies a Form, trigger, property, Grade Records, Drive file or email.
 * Trigger list reflects only triggers visible to the executing account.
 */
function qaPcSubmissionWiringReadOnly() {
  const formId = '1ObYHlSpNZlazoyLlr2eLQ0GNLes1yz_-sh_EanGgR4g';
  const properties = PropertiesService.getScriptProperties();
  const isCandidateInstalled =
    typeof pcV2_processSubmission_ === 'function' &&
    typeof pcV2_formSnapshot_ === 'function' &&
    typeof pcV2_preparePdfEmailAttachment_ === 'function';
  const handlersPresent = typeof professionalCompetency_onFormSubmit === 'function';
  const triggers = ScriptApp.getProjectTriggers()
    .filter(function(t){return t.getHandlerFunction() === 'professionalCompetency_onFormSubmit';})
    .map(function(t){
      let sourceId = null;
      try {sourceId = t.getTriggerSourceId();} catch (_) {}
      return {event:String(t.getEventType()),source:String(t.getTriggerSource()),sourceId:sourceId};
    });
  const formTriggers = triggers.filter(function(t) {
    return t.event === 'ON_FORM_SUBMIT' && t.source === 'FORMS' && t.sourceId === formId;
  });
  const recipients = typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.NOTIFICATIONS && APP_CONFIG.NOTIFICATIONS.COORDINATOR_RECIPIENTS;
  const missing = [];
  if(!handlersPresent) missing.push('No professionalCompetency_onFormSubmit definition in the current script project');
  if(!isCandidateInstalled) missing.push('V2 PDF-attachment candidate is not fully installed');
  if(properties.getProperty('CCMS_PC_V2_ENABLED') !== 'true') missing.push('CCMS_PC_V2_ENABLED is not true (candidate safely disabled)');
  if(formTriggers.length !== 1) missing.push('No unique Form-bound ON_FORM_SUBMIT trigger visible to this account for the authorized Form');
  if(typeof scheduling_getStudentScheduleForDate !== 'function') missing.push('Date-scoped scheduling resolver expected by candidate unavailable');
  if(typeof dashboardDocuments_getStudent_ !== 'function' ||
     typeof dashboardDocuments_findExistingStudentFolder_ !== 'function')
    missing.push('Existing student Drive folder resolver expected by candidate unavailable');
  if(typeof gradeRecords_add !== 'function' || typeof gradeRecords_list !== 'function')
    missing.push('Existing Grade Records service expected by candidate unavailable');
  if(!Array.isArray(recipients) || !recipients.length)
    missing.push('Coordinator recipients absent from configured notification list');
  const result = {
    handlerPresent:handlersPresent,
    candidateWithPdfAttachmentInstalled:isCandidateInstalled,
    v2Enabled:properties.getProperty('CCMS_PC_V2_ENABLED') === 'true',
    passThresholdProperty:properties.getProperty('CCMS_PC_PASS_MIN_PERCENT') || null,
    visibleHandlerTriggers:triggers,
    visibleAuthorizedFormTriggerCount:formTriggers.length,
    coordinatorRecipientCount:Array.isArray(recipients)?recipients.length:0,
    findings:missing,
    triggerVisibilityNote:'Apps Script displays only triggers visible to the executing account; another user may own a trigger.',
    noWrites:true
  };
  console.log('CCMS_PC_WIRING_DIAGNOSTIC='+JSON.stringify(result));
  return result;
}
