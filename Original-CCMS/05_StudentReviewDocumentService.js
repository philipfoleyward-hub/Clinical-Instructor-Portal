/**
 * CCMS ORIGINAL PROJECT ONLY — Student Review & Submission Record document.
 * CANDIDATE; do not install before verifying current original CCMS HEAD.
 *
 * This is distinct from the official academic V5.2 generator. Does not grade,
 * create folders, append Tracking, send mail, or change the Form workflow.
 *
 * Authority: ccmsPortal_resolveIdentity_ in original CCMS; immutable submission
 * ID; server persisted submission time; review state from authenticated draft.
 * Document filing is delegated to the established DriveService API.
 */
const CCMS_STUDENT_REVIEW_TEMPLATE_ID =
  '1jwwDvKrYiaWnba8wHFP8RZAHDgTJsScCy0Ddr1MMRKw';

function ccmsStudentReview_text_(value, maxLength, field) {
  if (value !== undefined && value !== null && typeof value !== 'string') {
    throw new Error(field + ' must be text');
  }
  const text = String(value || '').trim();
  if (text.length > maxLength) throw new Error(field + ' exceeds ' + maxLength + ' characters');
  return text || 'None recorded';
}

function ccmsStudentReview_clock_(date) {
  if (!(date instanceof Date) || !Number.isFinite(date.getTime())) {
    throw new Error('Invalid record timestamp');
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
}

/** Pure server-side template mapping; never trust client-supplied display names. */
function ccmsStudentReview_tokenValues_(request) {
  request = request || {};
  const identity = request.identity || {};
  const studentId = String(identity.studentId || '').trim();
  const evaluatorId = String(identity.evaluatorId || '').trim();
  const submissionId = String(request.submissionId || '').trim();
  if (!studentId || !evaluatorId || !identity.student || !identity.evaluator ||
      !identity.site || !identity.examination) {
    throw new Error('Student review requires complete server-verified identity');
  }
  if (!/^CI-PORTAL-[a-f0-9-]{36}$/i.test(submissionId)) {
    throw new Error('Student review requires a valid immutable Portal submission ID');
  }
  const submittedAt = request.submittedAt instanceof Date
    ? request.submittedAt : new Date(request.submittedAt || NaN);
  if (!Number.isFinite(submittedAt.getTime())) {
    throw new Error('Student review requires a persisted server submission time');
  }
  const review = request.review || {};
  const status = String(review.status || '').trim().toLowerCase();
  if (!['pending', 'reviewed'].includes(status)) {
    throw new Error('Student review requires pending or reviewed status');
  }
  let reviewedAt = 'Not reviewed';
  if (status === 'reviewed') {
    const date = new Date(review.reviewedAt || NaN);
    if (!Number.isFinite(date.getTime())) {
      throw new Error('Reviewed status requires a valid review timestamp');
    }
    if (date.getTime() > submittedAt.getTime() + 120000) {
      throw new Error('Review timestamp cannot be later than submission');
    }
    reviewedAt = ccmsStudentReview_clock_(date);
  }
  const date = identity.date instanceof Date ? identity.date : new Date(identity.date || NaN);
  if (!Number.isFinite(date.getTime())) throw new Error('Evaluation date is invalid');

  return {
    '{{STUDENT}}': String(identity.student).trim(),
    '{{CLINICAL_SITE}}': String(identity.site).trim(),
    '{{EVALUATOR}}': String(identity.evaluator).trim(),
    '{{EXAMINATION}}': String(identity.examination).trim(),
    '{{DATE}}': Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM d, yyyy'),
    '{{STUDENT_STRENGTHS}}': ccmsStudentReview_text_(review.strengths, 8000, 'Strengths'),
    '{{AREAS_FOR_IMPROVEMENT}}': ccmsStudentReview_text_(review.improvements, 8000, 'Areas for improvement'),
    '{{STUDENT_REVIEW_NOTES}}': ccmsStudentReview_text_(review.discussion, 8000, 'Discussion notes'),
    '{{STUDENT_REVIEW_STATUS}}': status === 'reviewed' ? 'Reviewed with Student' : 'Review Pending',
    '{{STUDENT_REVIEW_DATE}}': reviewedAt,
    '{{STUDENT_ID}}': studentId,
    '{{EVALUATOR_ID}}': evaluatorId,
    '{{SUBMISSION_ID}}': submissionId,
    '{{SUBMITTED_AT}}': ccmsStudentReview_clock_(submittedAt)
  };
}

function ccmsStudentReview_tokensIn_(body) {
  return Array.from(new Set(body.getText().match(/\{\{[^{}]+\}\}/g) || [])).sort();
}
function ccmsStudentReview_replaceTokens_(body, tokenValues) {
  const before = ccmsStudentReview_tokensIn_(body);
  const expected = Object.keys(tokenValues).sort();
  const missing = expected.filter(x => before.indexOf(x) === -1);
  const unexpected = before.filter(x => expected.indexOf(x) === -1);
  if (missing.length || unexpected.length) {
    throw new Error('Review template contract mismatch. Missing: ' +
      missing.join(', ') + ' Unexpected: ' + unexpected.join(', '));
  }
  for (const token of expected) {
    const regex = token.replace(/([\\^$.*+?()[\]{}|])/g, '\\$1');
    body.replaceText(regex, String(tokenValues[token]));
  }
  const unresolved = ccmsStudentReview_tokensIn_(body);
  if (unresolved.length) {
    throw new Error('Unresolved student review tokens: ' + unresolved.join(', '));
  }
}

/**
 * Generate the companion beside an ALREADY SERVER-GENERATED official V5.2.
 * `officialDocId` and `officialPdfId` MUST come from the original CCMS
 * generator in the SAME authorized submission transaction, never the browser.
 * NEW DriveService companion helpers verify sibling filing; no folder names
 * or student names are accepted as filing selectors here.
 *
 * Caller is responsible for persisting this returned receipt BEFORE writing
 * Clinical Tracking, and using the same submissionId on retry.
 */
function ccmsStudentReview_generate_(request) {
  request = request || {};
  const values = ccmsStudentReview_tokenValues_(request);
  const official = request.officialDocument || {};
  if (!official.docId || !official.pdfId) {
    throw new Error('Generate the official V5.2 Doc and PDF before the review record');
  }
  if (typeof driveService_copySubmissionCompanion_ !== 'function' ||
      typeof driveService_saveSubmissionCompanionPdf_ !== 'function') {
    throw new Error('Required original CCMS DriveService companion methods are unavailable');
  }
  const template = DriveApp.getFileById(CCMS_STUDENT_REVIEW_TEMPLATE_ID);
  if (template.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('Student review template is not a native Google Doc');
  }
  const safeId = request.submissionId.replace(/[^A-Za-z0-9-]/g, '');
  const fileName = 'Student Review and Submission Record - ' + safeId;
  const copy = driveService_copySubmissionCompanion_({
    templateId: CCMS_STUDENT_REVIEW_TEMPLATE_ID,
    fileName: fileName,
    officialDocId: official.docId,
    officialPdfId: official.pdfId
  });
  if (!copy || !copy.fileId) throw new Error('DriveService did not return companion Doc ID');
  const doc = DocumentApp.openById(copy.fileId);
  const body = doc.getBody();
  ccmsStudentReview_replaceTokens_(body, values);
  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(copy.fileId).getAs(MimeType.PDF)
    .setName(fileName + '.pdf');
  const pdf = driveService_saveSubmissionCompanionPdf_({
    blob: pdfBlob,
    fileName: fileName + '.pdf',
    officialDocId: official.docId,
    officialPdfId: official.pdfId
  });
  if (!pdf || !pdf.fileId) throw new Error('DriveService did not return companion PDF ID');
  return {
    success: true,
    submissionId: request.submissionId,
    docId: copy.fileId,
    docUrl: copy.fileUrl,
    pdfId: pdf.fileId,
    pdfUrl: pdf.fileUrl
  };
}

/** Read-only real Apps Script template check; makes no document copies. */
function ccmsStudentReview_qaTemplateReadOnly() {
  const file = DriveApp.getFileById(CCMS_STUDENT_REVIEW_TEMPLATE_ID);
  if (file.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('Student review template is not a native Google Doc');
  }
  const body = DocumentApp.openById(CCMS_STUDENT_REVIEW_TEMPLATE_ID).getBody();
  const expected = Object.keys(ccmsStudentReview_tokenValues_({
    identity: {studentId: 'TEST-001', student: 'Test Student', site: 'Test Site',
      evaluatorId: 'EV-TEST', evaluator: 'Test Evaluator', examination: 'Ankle', date: new Date()},
    submissionId: 'CI-PORTAL-123e4567-e89b-12d3-a456-426614174000',
    submittedAt: new Date(),
    review: {status: 'pending'}
  })).sort();
  const found = ccmsStudentReview_tokensIn_(body);
  const missing = expected.filter(token => found.indexOf(token) < 0);
  const unexpected = found.filter(token => expected.indexOf(token) < 0);
  return {
    success: missing.length === 0 && unexpected.length === 0,
    templateId: CCMS_STUDENT_REVIEW_TEMPLATE_ID,
    expectedTokens: expected.length,
    foundTokens: found.length,
    missingTokens: missing,
    unexpectedTokens: unexpected,
    tableCount: body.getTables().length,
    didWrite: false
  };
}