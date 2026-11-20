/**
 * CCMS Portal live TEST-001 connection — original CCMS project only.
 * 2026-10-16. This is an additive adapter to the EXISTING native services.
 *
 * Needs the user's installed, GridItem-corrected 04_PortalEvaluationAdapter,
 * ORIGINAL 03_DigitalEvaluationIntake helper digitalEvaluation_buildTrackingRecord_,
 * original grading engine, clinicalTerm, DriveService, TrackingService,
 * document and requirements services. It does not replace them or form triggers.
 *
 * Wire to existing standalone portal's CIPortalEvaluationTransport.js.
 * Caller authentication: this web app is deployed "Execute as: Me" with
 * "Who has access: Only myself". Google rejects any request that does not
 * carry the owner's own OAuth token, which the Portal (also running as the
 * owner) sends server-to-server. No shared secret is used.
 * No student other than TEST-001 can commit through this version, ever.
 * Installing this file does not enable a request or commit.
 */

function ciLink_prop_(name) {
  return String(PropertiesService.getScriptProperties().getProperty(name) || '').trim();
}
function ciLink_enabled_() {
  if (ciLink_prop_('CI_PORTAL_GATEWAY_ENABLED') !== 'true')
    throw Error('CCMS evaluation service is not enabled');
  // Defence in depth on top of Google's "Only myself" web-app access.
  const caller = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (caller && owner && caller !== owner) throw Error('Unauthorized gateway caller');
}
function ciLink_signature_(msg, key) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(msg, key))
    .replace(/=+$/, '');
}
function ciLink_equal_(a, b) {
  a = String(a || ''); b = String(b || '');
  let diff = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++)
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}
function ciLink_verify_(message) {
  ciLink_enabled_();
  if (!message || typeof message !== 'object' || Array.isArray(message))
    throw Error('Malformed signed request');
  const action = String(message.action || '');
  const actions = ['schema', 'preview', 'submit', 'draft_save', 'draft_load', 'draft_stop', 'draft_list'];
  if (!actions.includes(action)) throw Error('Unknown gateway action');
  const ms = Number(message.timestamp), nonce = String(message.nonce || '');
  if (!Number.isInteger(ms) || Math.abs(Date.now() - ms) > 120000 ||
      !/^[a-f\d-]{36}$/i.test(nonce)) throw Error('Expired or malformed gateway request');
  if (typeof message.payloadJson !== 'string' || message.payloadJson.length > 260000)
    throw Error('Invalid gateway payload');
  // Replay lock prevents concurrent use of the same signed nonce.
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const cache = CacheService.getScriptCache(), tag = 'CI_PORTAL_NONCE_' + nonce;
    if (cache.get(tag)) throw Error('Gateway request already used');
    cache.put(tag, '1', 180);
  } finally { lock.releaseLock(); }
  let data;
  try { data = JSON.parse(message.payloadJson); }
  catch (e) { throw Error('Invalid JSON payload'); }
  if (!data || typeof data !== 'object' || Array.isArray(data)) throw Error('Invalid payload shape');
  return {action: action, data: data};
}
function ciLink_identity_(data) {
  // Identity is re-resolved from authoritative CCMS reference tables.
  if (typeof ccmsPortal_resolveIdentity_ !== 'function')
    throw Error('Native grading adapter not installed');
  const identity = ccmsPortal_resolveIdentity_(data);
  if (String(data.evaluatorId || '') !== String(data.authenticatedEvaluatorId || '') ||
      identity.evaluatorId !== String(data.authenticatedEvaluatorId || ''))
    throw Error('Evaluator session identity mismatch');
  return identity;
}
function ciLink_testOnly_(data) {
  if (ciLink_prop_('CI_PORTAL_CCMS_SUBMIT_MODE') !== 'test_only')
    throw Error('CCMS submission is not enabled for TEST-001');
  if (String(data.studentId || '') !== 'TEST-001')
    throw Error('This connection can submit TEST-001 only');
  const inbox = ciLink_prop_('CI_PORTAL_QA_COORDINATOR_EMAIL').toLowerCase();
  const configured = APP_CONFIG.NOTIFICATIONS.COORDINATOR_RECIPIENTS;
  const recipients = (Array.isArray(configured) ? configured : String(configured || '').split(/[;,]/))
    .map(x => String(x || '').trim().toLowerCase()).filter(Boolean);
  if (!inbox || recipients.length !== 1 || recipients[0] !== inbox)
    throw Error('Approved test-only coordinator mailbox is not isolated');
  return inbox;
}
function ciLink_draftFolder_() {
  // Dedicated working-drafts folder is required; NOT a separate staging project.
  const id = ciLink_prop_('CI_PORTAL_DRAFTS_FOLDER_ID') ||
    String(APP_CONFIG.DRIVE.CI_PORTAL_DRAFTS_FOLDER_ID || '').trim();
  if (!/^[A-Za-z0-9_-]{20,}$/.test(id))
    throw Error('Configure CI_PORTAL_DRAFTS_FOLDER_ID in original CCMS Script Properties');
  const root = String(APP_CONFIG.DRIVE.CLINICAL_EVALUATIONS_ROOT_ID || ciLink_prop_('CI_PORTAL_OUTPUT_ROOT_ID') || '').trim();
  if (id === root) throw Error('Draft folder must be separate from final evaluations root');
  return DriveApp.getFolderById(id);
}
function ciLink_draftId_(sid) {
  const id = String(sid || '');
  if (!/^CI-PORTAL-[a-f\d-]{36}$/i.test(id)) throw Error('Invalid evaluation draft ID');
  return id;
}
function ciLink_draftFile_(id) {
  const files = ciLink_draftFolder_().getFilesByName(ciLink_draftId_(id) + '.json');
  if (!files.hasNext()) return null;
  const file = files.next();
  if (files.hasNext()) throw Error('Duplicate draft IDs require coordinator reconciliation');
  return file;
}
function ciLink_draftOwned_(saved, identity) {
  if (saved.studentId !== identity.studentId || saved.siteId !== identity.siteId ||
      saved.evaluatorId !== identity.evaluatorId || saved.examinationId !== identity.examinationId)
    throw Error('Evaluation draft access denied');
  const date = Utilities.formatDate(identity.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (saved.date !== date) throw Error('Evaluation draft date mismatch');
}
function ciLink_duplicate_(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING);
  return clinicalTracking_submissionExists_(sheet, ciLink_draftId_(id));
}
function ciLink_draftSave_(data) {
  const identity = ciLink_identity_(data), id = ciLink_draftId_(data.submissionId);
  if (ciLink_duplicate_(id)) throw Error('Evaluation is already finalized');
  const folder = ciLink_draftFolder_(), old = ciLink_draftFile_(id);
  if (old) ciLink_draftOwned_(JSON.parse(old.getBlob().getDataAsString('UTF-8')), identity);
  const record = {
    version: 1, submissionId: id, studentId: identity.studentId, siteId: identity.siteId,
    evaluatorId: identity.evaluatorId, examinationId: identity.examinationId,
    date: Utilities.formatDate(identity.date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    evaluationType: 'Competency', answers: data.answers || {}, uiState: data.uiState || null,
    currentSection: Math.max(0, Math.min(8, Number(data.currentSection) || 0)),
    reviewPending: !!data.reviewPending, savedAt: new Date().toISOString()
  };
  const json = JSON.stringify(record);
  if (json.length > 220000) throw Error('Draft is too large');
  if (old) old.setContent(json);
  else folder.createFile(id + '.json', json, MimeType.PLAIN_TEXT);
  return {success: true, submissionId: id, savedAt: record.savedAt};
}
function ciLink_draftLoad_(data) {
  const identity = ciLink_identity_(data), file = ciLink_draftFile_(data.submissionId);
  if (!file) return {success: true, found: false};
  const saved = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
  ciLink_draftOwned_(saved, identity);
  return {success: true, found: true, draft: saved};
}
function ciLink_draftList_(data) {
  const evaluatorId = String(data.authenticatedEvaluatorId || '');
  // A signed request by itself isn't a user authorization. Verify active
  // evaluator role and Portal Access against the original source workbook.
  const row = ccmsPortal_reference_(SpreadsheetApp.getActiveSpreadsheet(),
    APP_CONFIG.SHEETS.EVALUATORS, 'Evaluator ID', evaluatorId,
    ['Evaluator Type', 'Active', 'Portal Access']);
  if (!ccmsPortal_true_(row.Active) || !ccmsPortal_true_(row['Portal Access']) ||
      ccmsPortal_norm_(row['Evaluator Type']) !== 'college ci') throw Error('Draft access denied');
  const drafts = [], files = ciLink_draftFolder_().getFiles();
  let count = 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tracking = appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING);
  while (files.hasNext()) {
    if (++count > 500) throw Error('Draft list limit exceeded');
    const f = files.next();
    if (!/^CI-PORTAL-[a-f\d-]{36}\.json$/i.test(f.getName())) continue;
    let v; try { v = JSON.parse(f.getBlob().getDataAsString('UTF-8')); } catch (e) { continue; }
    if (v.evaluatorId !== evaluatorId || clinicalTracking_submissionExists_(tracking, v.submissionId)) continue;
    drafts.push({submissionId: v.submissionId, studentId: v.studentId,
      siteId: v.siteId, examinationId: v.examinationId, date: v.date,
      currentSection: v.currentSection || 0, savedAt: v.savedAt || '', reviewPending: !!v.reviewPending});
  }
  return {success: true, drafts: drafts.sort((a, b) => String(b.savedAt).localeCompare(String(a.savedAt))).slice(0, 100)};
}
function ciLink_commitTest001_(data) {
  // Recheck release at the closest point to irreversible effects.
  const mailbox = ciLink_testOnly_(data);
  const id = ciLink_draftId_(data.submissionId);
  const identity = ciLink_identity_(data);
  if (identity.studentId !== 'TEST-001') throw Error('Only TEST-001 is allowed');
  const review = data.uiState && data.uiState.studentReview;
  if (!review || !['pending', 'reviewed'].includes(String(review.status)))
    throw Error('Student Review status is required before Finalize');
  for (const k of ['strengths', 'improvements', 'discussion'])
    if (typeof review[k] !== 'string' || review[k].length > 8000)
      throw Error('Invalid student review comments');
  if (review.status === 'reviewed' && !Number.isFinite(Date.parse(review.reviewedAt)))
    throw Error('Reviewed status requires a valid timestamp');
  // Prevent test artifacts from landing in a similarly named but WRONG folder.
  const expectedRoot = '1-Mia4Zq7OI1NLUfgwUC5okl584FAIg_H';
  if (driveService_getClinicalEvaluationsRoot_().getId() !== expectedRoot)
    throw Error('Official evaluation Drive root is not the approved cohort/student root');
  if (typeof gradeClinicalEvaluationFromPortalPayload !== 'function')
    throw Error('Native grading adapter missing');
  const outcome = gradeClinicalEvaluationFromPortalPayload(data);
  if (outcome.identity.studentId !== identity.studentId ||
      outcome.identity.evaluatorId !== identity.evaluatorId) throw Error('Grade identity mismatch');

  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let doc, track, record;
  const key = 'CI_TEST001_DOC_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id)
  ).replace(/=+$/, '');
  try {
    if (ciLink_duplicate_(id))
      return {success: true, duplicate: true, submissionId: id,
        message: 'TEST-001 was already committed; no repeat documents, ledger entry or email'};
    const props = PropertiesService.getScriptProperties();
    try {doc = JSON.parse(props.getProperty(key) || 'null');} catch (e) {doc = null;}
    if (!doc || doc.submissionId !== id || !doc.docUrl || !doc.pdfUrl) {
      // Official approved document generator and DriveService remain owners.
      doc = evaluationDocument_generate(outcome.grade);
      if (!doc || !doc.success || !doc.docUrl || !doc.pdfUrl)
        throw Error('Official Doc/PDF output not confirmed; no Tracking append');
      props.setProperty(key, JSON.stringify({submissionId: id, docUrl: doc.docUrl,
        pdfUrl: doc.pdfUrl, docId: doc.docId || '', pdfId: doc.pdfId || ''}));
    }
    const attributed = clinicalTerm_getForDate(outcome.grade.evaluation.date);
    record = digitalEvaluation_buildTrackingRecord_({
      submissionId: id, grade: outcome.grade, attributedClinicalTerm: attributed,
      documentResult: doc, evaluatorEmail: '', editResponseUrl: '', formResponse: null
    });
    Object.assign(record, {
      'Student ID': identity.studentId, 'Clinical Site ID': identity.siteId,
      'Evaluator ID': identity.evaluatorId, 'Evaluator Type': identity.evaluatorType,
      'Examination ID': identity.examinationId, 'Entered By': 'CI Portal — TEST-001',
      'Source': 'CI Portal', 'Submission Method': 'CI Portal'
    });
    track = clinicalTracking_addRecord(record);
    if (!track || !track.success) throw Error('Original ClinicalTrackingService failed to confirm write');
    if (track.duplicate) return {success: true, duplicate: true, submissionId: id};
    props.deleteProperty(key);
    record = Object.assign({}, track.record || record, {'PDF Document': doc.pdfUrl,
      'Evaluator Email': ''});
  } finally { lock.releaseLock(); }

  let req, notice;
  try {req = rebuildStudentRequirementStatus();}
  catch (err) {req = {success: false, error: err.message};}
  // Only this approved institutional inbox is contacted; do NOT call
  // optional evaluator receipt, and visibly label QA subject.
  try {
    MailApp.sendEmail({to: mailbox,
      subject: '[TEST-001 QA ONLY] Clinical Competency Evaluation Submitted',
      body: 'This is a TEST submission, not a real student evaluation.\n' +
        'Submission ID: ' + id + '\nResult: ' + String(outcome.grade.result || '') +
        '\nFinal Score: ' + String(outcome.grade.finalScore) +
        '\nGoogle Doc: ' + doc.docUrl + '\nPDF: ' + doc.pdfUrl});
    notice = {success: true, recipient: mailbox, testOnly: true};
  } catch (err) {notice = {success: false, error: err.message};}
  return {success: true, duplicate: false, testOnly: true, submissionId: id,
    document: {docId: doc.docId || '', docUrl: doc.docUrl || '',
      pdfId: doc.pdfId || '', pdfUrl: doc.pdfUrl || ''},
    requirements: req, notifications: notice,
    grade: {finalScore: outcome.grade.finalScore, result: outcome.grade.result}};
}
/** Production release is separately authorized on the ORIGINAL CCMS server. */
function ciLink_productionEnabled_() {
  return ciLink_prop_('CI_PORTAL_CCMS_SUBMIT_MODE') === 'live' &&
    ciLink_prop_('CI_PORTAL_PRODUCTION_APPROVED') === 'true';
}
function ciLink_commitProduction_(data) {
  if (!ciLink_productionEnabled_()) throw Error('Production evaluation submission is disabled');
  const id = ciLink_draftId_(data.submissionId);
  const identity = ciLink_identity_(data);
  const review = data.uiState && data.uiState.studentReview;
  if (!review || !['pending','reviewed'].includes(String(review.status)))
    throw Error('Student Review is required before submission');
  for (const k of ['strengths','improvements','discussion'])
    if (typeof review[k] !== 'string' || review[k].length > 8000)
      throw Error('Invalid Student Review feedback');
  if (review.status === 'reviewed' && !Number.isFinite(Date.parse(review.reviewedAt)))
    throw Error('Student Review timestamp is required');
  if (driveService_getClinicalEvaluationsRoot_().getId() !== '1-Mia4Zq7OI1NLUfgwUC5okl584FAIg_H')
    throw Error('Evaluation filing root differs from approved authority');
  const outcome = gradeClinicalEvaluationFromPortalPayload(data);
  if (outcome.identity.studentId !== identity.studentId ||
      outcome.identity.evaluatorId !== identity.evaluatorId ||
      outcome.identity.examinationId !== identity.examinationId)
    throw Error('Official grading identity mismatch');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let doc, reviewDoc, record, track;
  const key = 'CI_PROD_DOC_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id)).replace(/=+$/, '');
  try {
    const props = PropertiesService.getScriptProperties();
    // Receipt survives a retry so a resubmission never duplicates documents.
    let receipt;
    try {receipt = JSON.parse(props.getProperty(key) || 'null');} catch (e) {receipt = null;}
    if (!receipt || receipt.submissionId !== id)
      receipt = {submissionId:id, submittedAt:new Date().toISOString(), doc:null, reviewDoc:null};
    if (ciLink_duplicate_(id)) return {success:true,duplicate:true,submissionId:id,
      document:receipt.doc||null, reviewDocument:receipt.reviewDoc||null,
      message:'Already recorded; no second document or notification was sent'};
    doc = receipt.doc;
    if (!doc || !doc.docId || !doc.pdfId || !doc.docUrl || !doc.pdfUrl) {
      const generated = evaluationDocument_generate(outcome.grade);
      if (!generated || !generated.success || !generated.docUrl || !generated.pdfUrl ||
          !generated.docId || !generated.pdfId)
        throw Error('Evaluation document generation did not complete');
      doc = {docId:generated.docId, docUrl:generated.docUrl, pdfId:generated.pdfId, pdfUrl:generated.pdfUrl};
      receipt.doc = doc;
      props.setProperty(key, JSON.stringify(receipt));
    }
    reviewDoc = receipt.reviewDoc;
    if (!reviewDoc || !reviewDoc.docUrl || !reviewDoc.pdfUrl) {
      const generatedReview = ccmsStudentReview_generate_({
        identity: identity, submissionId: id, submittedAt: receipt.submittedAt,
        review: review, officialDocument: doc});
      if (!generatedReview || !generatedReview.success || !generatedReview.docUrl || !generatedReview.pdfUrl)
        throw Error('Student Review document generation did not complete');
      reviewDoc = {docId:generatedReview.docId, docUrl:generatedReview.docUrl,
        pdfId:generatedReview.pdfId, pdfUrl:generatedReview.pdfUrl};
      receipt.reviewDoc = reviewDoc;
      props.setProperty(key, JSON.stringify(receipt));
    }
    const term = clinicalTerm_getForDate(outcome.grade.evaluation.date);
    // Evaluator identity and address come from the authoritative Evaluators row.
    const erow = ccmsPortal_reference_(SpreadsheetApp.getActiveSpreadsheet(),
      APP_CONFIG.SHEETS.EVALUATORS,'Evaluator ID',identity.evaluatorId,['Evaluator','Evaluator Type','Active','Portal Access']);
    const email = String(erow['Evaluator Email'] || erow.Email || '').trim();
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))
      throw Error('Evaluator email in reference table is invalid');
    record = digitalEvaluation_buildTrackingRecord_({submissionId:id,grade:outcome.grade,
      attributedClinicalTerm:term,documentResult:doc,evaluatorEmail:email,
      editResponseUrl:'',formResponse:null});
    Object.assign(record,{'Student ID':identity.studentId,'Clinical Site ID':identity.siteId,
      'Evaluator ID':identity.evaluatorId,'Evaluator Type':identity.evaluatorType,
      'Examination ID':identity.examinationId,'Entered By':'CI Portal',
      'Source':'CI Portal','Submission Method':'CI Portal',
      'Evaluator Email':email,'PDF Document':doc.pdfUrl});
    track = clinicalTracking_addRecord(record);
    if (!track || !track.success) throw Error('Clinical Tracking did not confirm the record');
    if (track.duplicate) return {success:true,duplicate:true,submissionId:id};
    props.deleteProperty(key);
    record = Object.assign({},track.record||record,{'Evaluator Email':email,
      'PDF Document':doc.pdfUrl,'Source Document':doc.docUrl});
  } finally {lock.releaseLock();}
  let requirements, notifications;
  // Recalculate only this student's requirement rows (full rebuild as fallback).
  try {requirements=typeof reqUpdateStudent_==='function'
    ? reqUpdateStudent_(SpreadsheetApp.getActiveSpreadsheet(), identity.studentId)
    : rebuildStudentRequirementStatus();}
  catch(e){requirements={success:false,error:String(e.message||e)};}
  // Reuse the exact existing coordinator/evaluator email templates and routing.
  try {notifications=notification_sendEvaluationNotificationsSafely(record);}
  catch(e){notifications={coordinator:{success:false,error:String(e.message||e)},
    evaluator:{success:false,error:String(e.message||e)}};}
  return {success:true,duplicate:false,submissionId:id,
    grade:{finalScore:outcome.grade.finalScore,result:outcome.grade.result},
    document:{docId:doc.docId||'',docUrl:doc.docUrl,pdfId:doc.pdfId||'',pdfUrl:doc.pdfUrl},
    reviewDocument:reviewDoc,
    tracking:{success:true},requirements:requirements,notifications:notifications};
}
/** Original CCMS web app POST, no standalone doGet replacement. */
function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents || '';
    if (!raw || raw.length > 360000) throw Error('Missing or oversize gateway request');
    const verified = ciLink_verify_(JSON.parse(raw));
    const a = verified.action, data = verified.data;
    let result;
    if (a === 'schema') {
      ciLink_identity_(data);
      result = {success: true, formItems: ccmsPortal_schema_()};
    } else if (a === 'preview') {
      ciLink_identity_(data);
      result = {success: true, grade: gradeClinicalEvaluationFromPortalPayload(data).grade};
    } else if (a === 'draft_save' || a === 'draft_stop') {
      result = ciLink_draftSave_(data);
    } else if (a === 'draft_load') {
      result = ciLink_draftLoad_(data);
    } else if (a === 'draft_list') {
      result = ciLink_draftList_(data);
    } else {
      result = ciLink_productionEnabled_() ? ciLink_commitProduction_(data) : ciLink_commitTest001_(data);
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('CCMS TEST-001 gateway: ' + String(err.message || err));
    return ContentService.createTextOutput(JSON.stringify({success: false,
      error: String(err.message || 'Request rejected')}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
