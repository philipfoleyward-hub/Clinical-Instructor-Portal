/**
 * 06_ProfessionalCompetencyForm.gs — COMPLETE REPLACEMENT CANDIDATE 2026-11-05
 * IMPORTANT: Not deployed. Reconcile against current Apps Script HEAD before installation.
 * Owns ONLY event orchestration; no Form rebuilding or new data tables.
 * Dependencies: existing GradeRecordsService, SchedulingService, Student repository,
 * 04_DashboardDocumentsService (existing-folder resolver), and three PC_V2 modules.
 * Install an onFormSubmit trigger on the Form itself (not a spreadsheet trigger).
 * Existing handler name is intentionally retained; remove old DEFINITION, never duplicate it.
 */
const PC_V2_FORM_ID = '1ObYHlSpNZlazoyLlr2eLQ0GNLes1yz_-sh_EanGgR4g';
const PC_V2_TEMPLATES = {
  'Clinical 1-2':'1VwJDexR_nuBOmZdvCt541P4kOTggEEPAHqQY2BXMDCo',
  'Clinical 3-5':'1byjBcUaPDUtXixNpB9ue3FM5nJQxFKYIvUU1DP0R_Dc'
};
const PC_V2_SOURCE = 'Professional Competency Google Form';

function professionalCompetency_onFormSubmit(e) {
  const enabled = PropertiesService.getScriptProperties().getProperty('CCMS_PC_V2_ENABLED');
  if (enabled !== 'true') {
    console.log('CCMS_PC_V2_DISABLED: candidate submit handler is not authorized. No write/email.');
    return {success:false,disabled:true};
  }
  if (!e || !e.response || !e.source || e.source.getId() !== PC_V2_FORM_ID) {
    throw new Error('A Form-bound onFormSubmit event from the authorized Professional Competency Form is required.');
  }
  const responseId = pcV2_text_(e.response.getId());
  if (!responseId) throw new Error('Form response ID is unavailable; cannot guarantee idempotence.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return pcV2_processSubmission_(e.response, responseId); }
  finally { lock.releaseLock(); }
}

function pcV2_processSubmission_(response, responseId) {
  const properties = PropertiesService.getScriptProperties();
  const form = FormApp.openById(PC_V2_FORM_ID);
  pcV2_formSnapshot_(form); // live Form must exactly match the approved audited criteria and choices
  const answers = pcV2_responseAnswers_(response);
  const studentValue = pcV2_text_(answers.byTitle['Student']);
  const evaluatorValue = pcV2_text_(answers.byTitle['Evaluator']);
  if (!studentValue || !evaluatorValue) throw new Error('Student or evaluator missing in Form response.');
  const dateRaw = answers.byTitle['Evaluation Date'];
  const date = dateRaw instanceof Date ? dateRaw : new Date(dateRaw);
  if (!dateRaw || isNaN(date.getTime())) throw new Error('Evaluation Date missing/invalid.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const student = pcV2_resolveStudent_(ss, studentValue);
  const evaluator = pcV2_resolveEvaluator_(ss, evaluatorValue);
  const term = clinicalTerm_getForDate(date);
  const family = pcV2_familyForTerm_(term);
  if (pcV2_familyForTerm_(student.clinicalTerm) !== family) {
    throw new Error('Student term ('+student.clinicalTerm+') and evaluation-date term ('+term+') require different rubrics. Resolve before grading.');
  }
  const score = pcV2_calculate_(family,answers);
  const result = pcV2_passResult_(score.percentGrade,properties);
  const site = pcV2_resolveSite_(student.studentId,date);
  const recipientList = pcV2_recipients_();
  const folder = pcV2_existingEvaluationFolder_(student);
  pcV2_validateTemplate_(PC_V2_TEMPLATES[family]);
  const strengths = pcV2_text_(answers.byTitle['The student demonstrated the following positive strengths:']);
  const improvements = pcV2_text_(answers.byTitle['The student demonstrates the following areas in need of improvement:']);
  const period = pcV2_text_(answers.byTitle['Evaluation Period']) || 'Not captured in Form';
  const sourceId = 'PC-'+responseId;
  const input = {responseId:responseId, sourceId:sourceId, student:student,
    evaluator:evaluator.name,evaluatorId:evaluator.id, date:date, clinicalTerm:term,
    site:site, period:period, score:score, result:result,
    strengths:strengths, improvements:improvements, folder:folder, recipients:recipientList};
  const fingerprint = pcV2_fingerprint_(input);
  const stateKey = 'CCMS_PC_V2_'+responseId;
  const state = JSON.parse(properties.getProperty(stateKey) || '{}');
  if (state.fingerprint && state.fingerprint !== fingerprint) {
    throw new Error('The submitted Form response was edited after initial processing. A controlled supersession is required; automatic overwrite refused.');
  }
  state.fingerprint = fingerprint;
  // This checkpoint is intentionally stored before any write (fail-safe reentry).
  pcV2_checkpoint_(properties,stateKey,state);

  const document = pcV2_buildDocument_(input,state.docId||'');
  state.docId = document.docId;
  pcV2_checkpoint_(properties,stateKey,state);
  const pdf = pcV2_buildPdf_(input,document,state.pdfId||'');
  state.pdfId = pdf.pdfId;
  pcV2_checkpoint_(properties,stateKey,state);

  // Do not regenerate an already recorded evaluation or silently reuse a mismatched grade.
  const records = gradeRecords_list(student.studentId,'',true).filter(function(r) {
    return pcV2_text_(r.sourceId) === sourceId && pcV2_text_(r.gradeCategory) === 'Professional Competency';
  });
  if (records.length > 1) throw new Error('Duplicate Grade Records with Source ID '+sourceId+' require coordinator reconciliation.');
  let grade = records[0] || null;
  if (grade) {
    if (Math.abs(Number(grade.rawScore)-score.rawScore)>0.001 ||
        Math.abs(Number(grade.percentGrade)-score.percentGrade)>0.001 ||
        pcV2_text_(grade.studentId)!==student.studentId ||
        (pcV2_text_(grade.sourceFileId) && grade.sourceFileId!==pdf.pdfId)) {
      throw new Error('Existing Grade Record differs from Form calculation or document. Historical record preserved; manual reconciliation required.');
    }
  } else {
    grade = gradeRecords_add({
      studentId:student.studentId,student:student.student,clinicalTerm:term,
      gradeCategory:'Professional Competency',
      assignmentName:'Professional Competency — '+period,
      evaluationPeriod:period,date:date,instructor:evaluator.name,
      rawScore:score.rawScore,percentGrade:score.percentGrade,maximumScore:'',
      deliveryMethod:'Digital',source:PC_V2_SOURCE,sourceId:sourceId,
      sourceDocument:pdf.url,sourceFileId:pdf.pdfId,
      notes:['Result: '+result,'Clinical site: '+site.name,'Site ID: '+site.id,
        'Evaluator ID: '+evaluator.id,'Form rubric: '+family,
        'Applicable weight: '+score.applicableWeight,
        'Criterion points: '+score.totalCriterionScore,
        'Positive strengths: '+strengths,'Areas for improvement: '+improvements,
        'Form response ID: '+responseId].join('\n')
    });
  }
  state.gradeRecordId = grade.gradeRecordId;
  pcV2_checkpoint_(properties,stateKey,state);

  // Prepare the actual student-specific PDF attachment before any mail send checkpoint.
  // If preparation fails, retries can safely resume without a stuck SENDING status.
  if (!state.mailStatus) {
    const pdfAttachment = pcV2_preparePdfEmailAttachment_(input,pdf.pdfId);
    state.mailStatus='SENDING';
    pcV2_checkpoint_(properties,stateKey,state);
    pcV2_sendCoordinatorEmail_(input,grade,pdf.url,document.url,pdfAttachment);
    state.mailStatus='SENT';
    pcV2_checkpoint_(properties,stateKey,state);
  } else if (state.mailStatus==='SENDING') {
    console.error('CCMS PC email delivery uncertain for '+sourceId+'. No automatic resend.');
  }
  console.log('CCMS_PC_V2_RESULT '+JSON.stringify({sourceId:sourceId,
    studentId:student.studentId, gradeRecordId:grade.gradeRecordId,
    percentGrade:score.percentGrade,result:result,
    docId:document.docId,pdfId:pdf.pdfId,mailStatus:state.mailStatus}));
  return {success:true,sourceId:sourceId,studentId:student.studentId,
    gradeRecordId:grade.gradeRecordId,grade:score.percentGrade,result:result,
    documentUrl:document.url,pdfUrl:pdf.url,mailStatus:state.mailStatus};
}
function pcV2_checkpoint_(properties,key,state) {properties.setProperty(key,JSON.stringify(state));}
function pcV2_fingerprint_(input) {
  const s = JSON.stringify({studentId:input.student.studentId,evaluatorId:input.evaluatorId,
    date:input.date.toISOString(),siteId:input.site.id,clinicalTerm:input.clinicalTerm,
    items:input.score.details.map(function(x){return [x.id,x.rating,x.weight];}),
    strengths:input.strengths,improvements:input.improvements});
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s);
  return Utilities.base64EncodeWebSafe(bytes);
}
function pcV2_resolveStudent_(ss, selected) {
  const data=appSheetToObjects_(appGetSheet_(ss,APP_CONFIG.SHEETS.STUDENTS));
  const match=data.filter(function(x) {
    return appIsTrue_(x.Active) && pcV2_normalizedText_(x.Student).toLowerCase()===pcV2_normalizedText_(selected).toLowerCase();
  });
  const ids={};match.forEach(function(r){ids[pcV2_text_(r['Student ID'])]=r;});
  const matches=Object.keys(ids).filter(Boolean);
  if(matches.length!==1)throw new Error('Student selection must resolve to exactly one active stable Student ID: '+selected);
  const r=ids[matches[0]];
  return {studentId:matches[0],student:pcV2_text_(r.Student),cohort:pcV2_text_(r.Cohort),clinicalTerm:pcV2_text_(r['Clinical Term'])};
}
function pcV2_resolveEvaluator_(ss,selected) {
  const data=appSheetToObjects_(appGetSheet_(ss,APP_CONFIG.SHEETS.EVALUATORS));
  const matches=data.filter(function(x) {return appIsTrue_(x.Active) &&
    pcV2_normalizedText_(x.Evaluator).toLowerCase()===pcV2_normalizedText_(selected).toLowerCase();});
  if(matches.length!==1 || !pcV2_text_(matches[0]['Evaluator ID']))
    throw new Error('Evaluator selection must resolve to exactly one active Evaluator ID: '+selected);
  return {id:pcV2_text_(matches[0]['Evaluator ID']),name:pcV2_text_(matches[0].Evaluator)};
}
function pcV2_resolveSite_(studentId,date) {
  if(typeof scheduling_getStudentScheduleForDate!=='function')
    throw new Error('SchedulingService date-scoped student-site resolution unavailable. Site cannot be guessed from current student profile.');
  const scheduled=scheduling_getStudentScheduleForDate(studentId,date);
  if(!scheduled || !scheduled.siteId || !scheduled.site)
    throw new Error('No verified dated clinical-site assignment for '+studentId+'. Resolve assignment before submission.');
  return {id:pcV2_text_(scheduled.siteId),name:pcV2_text_(scheduled.site)};
}
function pcV2_recipients_() {
  const list=APP_CONFIG.NOTIFICATIONS && APP_CONFIG.NOTIFICATIONS.COORDINATOR_RECIPIENTS;
  if(!Array.isArray(list)||!list.length || !list.every(function(x){return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(pcV2_text_(x));}))
    throw new Error('No approved CCMS coordinator notification recipients configured.');
  return list.map(pcV2_text_);
}

/** Safeguard: older setup/sync code rebuilt the Form. These methods cannot mutate it now. */
function setupProfessionalCompetencyForm() { throw new Error('Disabled: existing live Professional Competency Form must never be rebuilt.'); }
function syncProfessionalCompetencyForm() { throw new Error('Disabled: use read-only Form QA instead.'); }
/** This does NOT install/remove triggers. Inspect existing handler triggers before enabling. */
function pcV2_readOnlyDeploymentStatus() {
 const triggers=ScriptApp.getProjectTriggers().filter(function(x){return x.getHandlerFunction()==='professionalCompetency_onFormSubmit';});
 return {enabled:PropertiesService.getScriptProperties().getProperty('CCMS_PC_V2_ENABLED')==='true',
   formId:PC_V2_FORM_ID,triggers:triggers.map(function(x){return {handler:x.getHandlerFunction(),event:String(x.getEventType()),source:String(x.getTriggerSource())};}),
   templateIds:PC_V2_TEMPLATES};
}
