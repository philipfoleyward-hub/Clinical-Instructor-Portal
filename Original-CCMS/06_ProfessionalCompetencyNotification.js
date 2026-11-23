/**
 * 06_ProfessionalCompetencyNotification.gs — CANDIDATE
 * Coordinator-only email for a completed Professional Competency submission.
 * Uses APP_CONFIG.NOTIFICATIONS.COORDINATOR_RECIPIENTS; no new recipient store.
 * No email is sent from read-only QA or template validation.
 */
/** Rows added only for CI Portal submissions (rubric version, period, review). */
function pcV2_emailExtraRows_(input) {
  if (!input || input.channel !== 'CI Portal') return [];
  const rows = [['Rubric version', input.score.family === 'Clinical 1-2' ?
      'Clinical Practicum 1 & 2' : 'Clinical Practicum 3 - 5'],
    ['Evaluation period', input.period],
    ['Student review', input.reviewStatus === 'reviewed' ? 'Reviewed with student' : 'Review pending'],
    ['Submitted via', 'Clinical Instructor Portal']];
  if (input.discussion) rows.push(['Review discussion', input.discussion]);
  return rows;
}
function pcV2_emailPayload_(input,grade,pdfUrl,sheetUrl) {
  const when=Utilities.formatDate(input.date,Session.getScriptTimeZone(),'MMM d, yyyy');
  const subject='Professional Competency — '+input.student.student+' — '+input.result;
  const lines=[
    'A Professional Competency evaluation has been submitted and graded.',
    '',
    'Student: '+input.student.student+' ('+input.student.studentId+')',
    'Clinical term: '+input.clinicalTerm,
    'Evaluator: '+input.evaluator,
    'Clinical site: '+input.site.name,
    'Evaluation date: '+when,
    'Weighted score: '+input.score.rawScore.toFixed(2)+' / 3.00',
    'Grade: '+input.score.percentGrade.toFixed(2)+'%',
    'Result: '+input.result
  ].concat(pcV2_emailExtraRows_(input).map(function(r){return r[0]+': '+r[1];})).concat([
    '',
    'Attached: completed Professional Competency evaluation (PDF)',
    'Completed evaluation (PDF): '+pdfUrl,
    'Editable evaluation sheet: '+sheetUrl,
    'Submission ID: '+input.sourceId,
    'Grade Record ID: '+grade.gradeRecordId
  ]);
  const plain=lines.join('\n');
  function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
  const html='<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.55;max-width:650px">'+
    '<p>A <strong>Professional Competency</strong> evaluation has been submitted and graded.</p>'+
    '<table style="border-collapse:collapse">'+
    [['Student',input.student.student+' ('+input.student.studentId+')'],
     ['Clinical term',input.clinicalTerm],['Evaluator',input.evaluator],
     ['Clinical site',input.site.name],['Evaluation date',when],
     ['Weighted score',input.score.rawScore.toFixed(2)+' / 3.00'],
     ['Grade',input.score.percentGrade.toFixed(2)+'%'],['Result',input.result]
    ].concat(pcV2_emailExtraRows_(input)).map(function(r){return '<tr><th style="padding:4px 16px 4px 0;text-align:left">'+esc(r[0])+'</th><td>'+esc(r[1])+'</td></tr>';}).join('')+
    '</table><p><strong>Attached:</strong> Completed Professional Competency evaluation (PDF)</p>'+
    '<p><a href="'+esc(pdfUrl)+'">View completed evaluation (PDF)</a> &nbsp;|&nbsp;'+
    '<a href="'+esc(sheetUrl)+'">Open grading sheet</a></p>'+
    '<p style="font-size:12px;color:#555">Submission ID: '+esc(input.sourceId)+
    '<br>Grade Record ID: '+esc(grade.gradeRecordId)+'</p></div>';
  return {subject:subject,body:plain,htmlBody:html};
}
/**
 * Validate and materialize the already-filed PDF before moving mailStatus to SENDING.
 * That preserves a safe retry when an attachment cannot be prepared.
 */
function pcV2_preparePdfEmailAttachment_(input, pdfId) {
  if (!pdfId || !input || !input.folder || !input.folder.folderId)
    throw new Error('Missing PDF or authorized student folder for email attachment.');
  const file = DriveApp.getFileById(pdfId);
  if (file.getMimeType() !== MimeType.PDF)
    throw new Error('Coordinator attachment is not a PDF; email halted.');
  pcV2_verifyParent_(file, input.folder.folderId);
  const size = file.getSize();
  if (!size || size > 20 * 1024 * 1024)
    throw new Error('Completed evaluation PDF is empty or too large for safe email attachment ('+size+' bytes).');
  const blob = file.getBlob().setName(file.getName());
  if (blob.getContentType() !== 'application/pdf')
    throw new Error('Completed evaluation attachment content type is not PDF.');
  return blob;
}

function pcV2_sendCoordinatorEmail_(input,grade,pdfUrl,sheetUrl,pdfAttachment) {
  if (!pdfAttachment || typeof pdfAttachment.getContentType !== 'function' ||
      pdfAttachment.getContentType() !== 'application/pdf')
    throw new Error('PDF attachment is required. Coordinator email not sent.');
  const message=pcV2_emailPayload_(input,grade,pdfUrl,sheetUrl);
  MailApp.sendEmail({
    to:input.recipients.join(','),
    subject:message.subject,
    body:message.body,
    htmlBody:message.htmlBody,
    attachments:[pdfAttachment],
    name:(APP_CONFIG.NOTIFICATIONS && APP_CONFIG.NOTIFICATIONS.SENDER_NAME)||'Clinical Evaluation System'
  });
  return {success:true,recipients:input.recipients,subject:message.subject,attachedPdfName:pdfAttachment.getName()};
}
