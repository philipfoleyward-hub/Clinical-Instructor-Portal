/**
 * 06_ProfessionalCompetencyPortal.gs — CI Portal Professional Competency.
 * Gateway actions pc_schema / pc_preview / pc_submit (see CIPortalTest001 doPost).
 * Authority: PC_V2_RUBRICS (locked Form item IDs, wording, weights),
 * pcV2_calculate_ / pcV2_passResult_ (approved grading), the approved
 * tokenized Sheet templates and coordinator email. Identity is resolved only
 * from stable IDs in the CCMS reference tables; nothing is taken from names.
 */
const PC_PORTAL_EXAM_ID = 'PROFESSIONAL-COMPETENCY';
const PC_PORTAL_SOURCE = 'CI Portal';

function pcPortal_identity_(data) {
  if (!data || typeof data !== 'object') throw Error('Professional Competency data is required');
  const claimed = ccmsPortal_clean_(data.authenticatedEvaluatorId);
  if (!claimed || claimed !== ccmsPortal_clean_(data.evaluatorId)) throw Error('Authenticated evaluator mismatch');
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const student = ccmsPortal_reference_(book, APP_CONFIG.SHEETS.STUDENTS, 'Student ID', data.studentId, ['Student', 'Active']);
  if (!ccmsPortal_true_(student.Active)) throw Error('Student is not active');
  const site = ccmsPortal_reference_(book, APP_CONFIG.SHEETS.CLINICAL_SITES, 'Clinical Site ID', data.siteId, ['Clinical Site', 'Active']);
  if (!ccmsPortal_true_(site.Active)) throw Error('Clinical site is not active');
  const evaluator = ccmsPortal_reference_(book, APP_CONFIG.SHEETS.EVALUATORS, 'Evaluator ID', data.evaluatorId,
    ['Evaluator', 'Evaluator Type', 'Active', 'Portal Access']);
  if (!ccmsPortal_true_(evaluator.Active) || !ccmsPortal_true_(evaluator['Portal Access']) ||
      ccmsPortal_norm_(evaluator['Evaluator Type']) !== 'college ci')
    throw Error('Evaluator is not an active, authorized College CI');
  const date = ccmsPortal_date_(data.date);
  const term = pcV2_text_(clinicalTerm_getForDate(date));
  const family = pcV2_familyForTerm_(term);
  const studentTerm = pcV2_text_(student['Clinical Term']);
  if (studentTerm && pcV2_familyForTerm_(studentTerm) !== family)
    throw Error('Student term (' + studentTerm + ') and evaluation-date term (' + term +
      ') require different Professional Competency versions. Resolve before grading.');
  return {
    student: {studentId: ccmsPortal_clean_(student['Student ID']), student: ccmsPortal_clean_(student.Student),
      cohort: ccmsPortal_clean_(student.Cohort || ''), clinicalTerm: studentTerm || term},
    site: {id: ccmsPortal_clean_(site['Clinical Site ID']), name: ccmsPortal_clean_(site['Clinical Site'])},
    evaluator: {id: ccmsPortal_clean_(evaluator['Evaluator ID']), name: ccmsPortal_clean_(evaluator.Evaluator)},
    date: date, clinicalTerm: term, family: family
  };
}

function pcPortal_schema_(data) {
  const id = pcPortal_identity_(data);
  return {success: true, family: id.family,
    versionLabel: id.family === 'Clinical 1-2' ? 'Clinical Practicum 1 & 2' : 'Clinical Practicum 3 - 5',
    clinicalTerm: id.clinicalTerm, student: id.student.student, studentId: id.student.studentId,
    site: id.site.name, evaluator: id.evaluator.name,
    criteria: pcV2_expectedRubric_(id.family).map(function(x) {return {id: String(x.id), title: x.title, weight: x.weight};})};
}

/** Portal answers are keyed by the stable rubric (Form) item ID. */
function pcPortal_score_(family, data) {
  const given = data.answers && typeof data.answers === 'object' ? data.answers : {};
  const byId = {};
  pcV2_expectedRubric_(family).forEach(function(item) {byId[String(item.id)] = pcV2_text_(given[String(item.id)]);});
  const score = pcV2_calculate_(family, {byId: byId});
  const result = pcV2_passResult_(score.percentGrade, PropertiesService.getScriptProperties());
  return {score: score, result: result};
}

function pcPortal_preview_(data) {
  const id = pcPortal_identity_(data);
  const graded = pcPortal_score_(id.family, data);
  return {success: true, grade: {family: id.family, rawScore: graded.score.rawScore,
    percentGrade: graded.score.percentGrade, result: graded.result,
    applicableWeight: graded.score.applicableWeight, totalCriterionScore: graded.score.totalCriterionScore,
    criterionCount: graded.score.criterionCount}};
}

function pcPortal_submit_(data) {
  if (!ciLink_productionEnabled_()) throw Error('Production evaluation submission is disabled');
  const submissionId = ciLink_draftId_(data.submissionId);
  const id = pcPortal_identity_(data);
  const review = data.uiState && data.uiState.studentReview;
  if (!review || !['pending', 'reviewed'].includes(String(review.status)))
    throw Error('Student Review status is required before submission');
  ['strengths', 'improvements', 'discussion'].forEach(function(k) {
    if (typeof review[k] !== 'string' || review[k].length > 8000) throw Error('Invalid Student Review feedback');
  });
  if (review.status === 'reviewed' && !Number.isFinite(Date.parse(review.reviewedAt)))
    throw Error('Student Review timestamp is required');
  const graded = pcPortal_score_(id.family, data);
  const period = pcV2_text_(data.period).slice(0, 120) || 'Not specified';
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const folder = pcV2_existingEvaluationFolder_(id.student);
    pcV2_validateTemplate_(PC_V2_TEMPLATES[id.family]);
    const input = {responseId: submissionId, sourceId: 'PC-' + submissionId, student: id.student,
      evaluator: id.evaluator.name, evaluatorId: id.evaluator.id, date: id.date, clinicalTerm: id.clinicalTerm,
      site: id.site, period: period, score: graded.score, result: graded.result,
      strengths: pcV2_text_(review.strengths), improvements: pcV2_text_(review.improvements),
      discussion: pcV2_text_(review.discussion), reviewStatus: String(review.status),
      folder: folder, recipients: pcV2_recipients_(), channel: PC_PORTAL_SOURCE};
    const out = pcV2_commit_(input, {source: PC_PORTAL_SOURCE,
      originLine: 'CI Portal submission ID: ' + submissionId,
      extraNotes: ['Evaluation period: ' + period,
        'Student review: ' + (review.status === 'reviewed' ? 'Reviewed with student ' + review.reviewedAt : 'Pending'),
        'Review discussion: ' + pcV2_text_(review.discussion)]});
    return {success: true, duplicate: false, submissionId: submissionId, sourceId: out.sourceId,
      grade: {rawScore: out.rawScore, percentGrade: out.grade, result: out.result, family: id.family},
      document: {docUrl: out.documentUrl, pdfUrl: out.pdfUrl}, gradeRecordId: out.gradeRecordId,
      mailStatus: out.mailStatus};
  } finally {lock.releaseLock();}
}
