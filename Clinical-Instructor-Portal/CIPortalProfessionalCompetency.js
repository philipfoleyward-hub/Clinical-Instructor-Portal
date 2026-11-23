/**
 * Clinical Instructor Portal — Professional Competency Evaluation.
 * Same lifecycle as the competency evaluation: authorized open, private
 * evaluator-owned draft, official server grade preview, finalize & submit.
 * The rubric version (Clinical Practicum 1 & 2 vs 3 - 5), grading, document
 * and email all come from the original CCMS through the owner-token gateway.
 */
const CI_PC_EXAM_ID = 'PROFESSIONAL-COMPETENCY';

function ciPcTrusted_(token, request) {
  const data = ciPortalEvaluationTrustedPayload_(token, Object.assign({}, request, {evaluationType: 'Competency'}));
  data.examinationId = CI_PC_EXAM_ID;
  data.evaluationType = 'Professional Competency';
  data.period = String(request && request.period || '').trim().slice(0, 120);
  return data;
}
function ciPcSchema_(data) {
  const key = 'CI_PC_SCHEMA_' + [data.studentId, data.siteId, data.date, data.evaluatorId].join('|');
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const schema = ciPortalEvaluationTransport_('pc_schema', data);
  if (!schema.family || !Array.isArray(schema.criteria) || !schema.criteria.length)
    throw new Error('Professional Competency rubric was not returned by CCMS.');
  try {cache.put(key, JSON.stringify(schema), 600);} catch (e) {}
  return schema;
}
function ciR1PcInitialize(token, request) {
  const data = ciPcTrusted_(token, request);
  const resuming = Boolean(data.submissionId);
  data.submissionId = data.submissionId || 'CI-PORTAL-' + Utilities.getUuid();
  const schema = ciPcSchema_(data);
  let initialDraft = {success: true, notCreated: resuming};
  if (!resuming) {
    try {initialDraft = ciDraftSave_(data, {answers: {}, uiState: null, currentSection: 0, reviewPending: true});}
    catch (e) {initialDraft = {success: false, error: String(e.message || e)};}
  }
  return {success: true, submissionId: data.submissionId, schema: schema, initialDraft: initialDraft,
    html: HtmlService.createHtmlOutputFromFile('CI_Portal_Professional_Competency_Inner').getContent()};
}
function ciPcDraftPayload_(request) {
  const uiState = request.uiState || null;
  if (uiState) uiState.period = String(request.period || '').slice(0, 120);
  return {answers: request.answers || {}, uiState: uiState,
    currentSection: request.currentSection, reviewPending: request.reviewPending};
}
function ciR1PcSaveDraft(token, request) {
  const data = ciPcTrusted_(token, request);
  if (!data.submissionId) throw new Error('Draft ID required');
  return ciDraftSave_(data, ciPcDraftPayload_(request));
}
function ciR1PcLoadDraft(token, request) {
  const data = ciPcTrusted_(token, request);
  if (!data.submissionId) throw new Error('Draft ID required');
  return ciDraftLoad_(data);
}
function ciR1PcPreview(token, request) {
  const data = ciPcTrusted_(token, request);
  return ciPortalEvaluationTransport_('pc_preview', data);
}
function ciR1PcSubmit(token, request) {
  const data = ciPcTrusted_(token, request);
  if (!data.submissionId) throw new Error('Stable submission ID required');
  const mode = ciPortalEvaluationMode_();
  if (mode !== 'live' || !ciPortalEvaluationHasConnection_())
    throw new Error('Official submission is disabled. Draft preserved; no evaluation was submitted.');
  if (!data.answers || !Object.keys(data.answers).length) throw new Error('Submission has no ratings');
  const review = request.uiState && request.uiState.studentReview;
  if (!review || !['pending', 'reviewed'].includes(String(review.status))) throw new Error('Student Review & Feedback status is required');
  for (const prop of ['strengths', 'improvements', 'discussion'])
    if (typeof review[prop] !== 'string' || review[prop].length > 8000) throw new Error('Student review text is invalid');
  if (review.status === 'reviewed' && (!review.reviewedAt || !Number.isFinite(Date.parse(review.reviewedAt))))
    throw new Error('Reviewed with Student requires a recorded time');
  data.uiState = request.uiState;
  ciDraftSave_(data, ciPcDraftPayload_(request)); // preserve final answers before commit
  return ciPortalEvaluationTransport_('pc_submit', data);
}
