/**
 * 07_CIReflections.gs — Coordinator side of the Clinical Instructor Summary
 * Reflection (the CI "survey response").
 *
 * Flow (coordinator decision 2026-11-24):
 *   1. The coordinator shares each instructor's student-feedback summary for a
 *      semester (pasted text and/or a PDF) → the instructor is notified.
 *   2. The instructor completes the three sections of the approved template in
 *      the CI Portal (My CI Workspace → Survey Responses) and submits; the Portal
 *      fills the Google Doc template, saves the Doc + PDF in a private folder and
 *      notifies the coordinator.
 *   3. The coordinator accepts or returns it with notes.
 * Shared sheet: "CI Reflections" (also written by the CI Portal).
 */
const CP_REFL = {
  SHEET: 'CI Reflections',
  HEADERS: ['Reflection ID', 'Evaluator ID', 'Evaluator', 'Term', 'Summary Text', 'Summary File ID', 'Shared At', 'Shared By',
    'Observations', 'Learning Experiences', 'Goals', 'Status', 'Saved At', 'Submitted At', 'Doc ID', 'PDF ID',
    'Coordinator Note', 'Reviewed At', 'Reviewed By', 'Version'],
  FOLDER: 'CCMS — CI Reflections (restricted)'
};
function cpRefl_rows_() {return cpMile_rows_(CP_REFL.SHEET, CP_REFL.HEADERS);}
function cpRefl_write_(row) {cpMile_write_(CP_REFL.SHEET, CP_REFL.HEADERS, row);}
function cpRefl_append_(obj) {cpMile_sheet_(CP_REFL.SHEET, CP_REFL.HEADERS).appendRow(CP_REFL.HEADERS.map(function(h) {return obj[h] == null ? '' : obj[h];}));}
function cpRefl_folder_(sub) {
  const root = DriveApp.getRootFolder();
  const it = root.getFoldersByName(CP_REFL.FOLDER);
  const base = it.hasNext() ? it.next() : root.createFolder(CP_REFL.FOLDER);
  if (!sub) return base;
  const s = base.getFoldersByName(sub);
  return s.hasNext() ? s.next() : base.createFolder(sub);
}
function cpRefl_out_(r) {
  const f = function(id, kind) {return id ? (kind === 'doc' ? 'https://docs.google.com/document/d/' + id + '/edit' : 'https://drive.google.com/file/d/' + id + '/view') : '';};
  return {reflectionId: cpLink_t_(r['Reflection ID']), evaluatorId: cpLink_t_(r['Evaluator ID']), evaluator: cpLink_t_(r['Evaluator']),
    term: cpLink_t_(r['Term']), summaryText: cpLink_t_(r['Summary Text']), summaryUrl: f(cpLink_t_(r['Summary File ID'])),
    sharedAt: cpLink_iso_(r['Shared At']), status: cpLink_t_(r['Status']) || 'Requested',
    observations: cpLink_t_(r['Observations']), learning: cpLink_t_(r['Learning Experiences']), goals: cpLink_t_(r['Goals']),
    submittedAt: cpLink_iso_(r['Submitted At']), docUrl: f(cpLink_t_(r['Doc ID']), 'doc'), pdfUrl: f(cpLink_t_(r['PDF ID'])),
    coordinatorNote: cpLink_t_(r['Coordinator Note']), reviewedAt: cpLink_iso_(r['Reviewed At']), version: cpMile_n_(r['Version']) || 0};
}
/** All reflections, newest first, grouped by evaluator (used by the Clinical Instructors page). */
function cpRefl_byEvaluator_() {
  const out = {};
  cpRefl_rows_().forEach(function(r) {
    const o = cpRefl_out_(r);
    (out[o.evaluatorId] = out[o.evaluatorId] || []).push(o);
  });
  Object.keys(out).forEach(function(k) {out[k].sort(function(a, b) {return (b.sharedAt || '').localeCompare(a.sharedAt || '');});});
  return out;
}

/** Share (or update) an instructor's student-feedback summary for a semester and ask for the reflection. */
function dashboardV2ReflectionShare(request) {
  const r = request || {};
  const id = cpLink_t_(r.evaluatorId);
  const term = cpMile_termRange_(r.term).term;
  const text = cpLink_requireText_(r.summaryText, 'Summary', 20000);
  const ev = cpHours_evaluators_(SpreadsheetApp.getActiveSpreadsheet()).filter(function(e) {return e.evaluatorId === id;})[0];
  if (!ev) throw new Error('Instructor not found (college clinical instructors only).');
  let fileId = '';
  if (r.pdfBase64) {
    const bytes = Utilities.base64Decode(String(r.pdfBase64));
    if (bytes.length > 10 * 1024 * 1024) throw new Error('The PDF is larger than 10 MB.');
    if (String.fromCharCode.apply(null, bytes.slice(0, 5)) !== '%PDF-') throw new Error('The attachment must be a PDF.');
    fileId = cpRefl_folder_(term).createFile(Utilities.newBlob(bytes, 'application/pdf', ev.evaluator + ' — ' + term + ' student feedback summary.pdf')).getId();
  }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let row;
  try {
    row = cpRefl_rows_().filter(function(x) {return cpLink_t_(x['Evaluator ID']) === id && cpLink_t_(x['Term']) === term;})[0];
    if (row && cpLink_t_(row['Status']) === 'Accepted') throw new Error('This reflection was already accepted.');
    if (!text && !fileId && !(row && (cpLink_t_(row['Summary Text']) || cpLink_t_(row['Summary File ID'])))) throw new Error('Paste the feedback summary or attach it as a PDF.');
    if (row) {
      if (text) row['Summary Text'] = text;
      if (fileId) row['Summary File ID'] = fileId;
      row['Shared At'] = new Date(); row['Shared By'] = cpLink_user_();
      cpRefl_write_(row);
    } else {
      row = {'Reflection ID': 'REF-' + cpMile_termRange_(term).term.replace(/^(\w)\w+ \d\d(\d\d)$/, '$1$2') + '-' + id,
        'Evaluator ID': id, 'Evaluator': ev.evaluator, 'Term': term, 'Summary Text': text, 'Summary File ID': fileId,
        'Shared At': new Date(), 'Shared By': cpLink_user_(), 'Status': 'Requested', 'Version': 0};
      cpRefl_append_(row);
    }
  } finally {lock.releaseLock();}
  const due = cpLink_requireText_(r.due, 'Due', 100);
  cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': id, 'Recipient': ev.evaluator, 'Type': 'Reflection Requested',
    'Item Kind': 'reflection', 'Item ID': cpLink_t_(row['Reflection ID']),
    'Title': 'Student feedback summary — ' + term + ': please complete your reflection',
    'Message': (due ? 'Please complete by ' + due + '.\n' : '') + 'Open My CI Workspace → Survey Responses.', 'From': cpLink_user_()});
  let emailed = false;
  if (r.email !== false) {
    try {
      emailed = cpHours_email_(id, 'Your ' + term + ' student feedback summary and reflection',
        ['Hello ' + ev.evaluator + ',', '', 'Your student feedback summary for ' + term + ' is ready in the Clinical Instructor Portal.',
          'Please review it and complete the Clinical Instructor Summary Reflection (My CI Workspace → Survey Responses)' + (due ? ' by ' + due : '') + '.']);
    } catch (e) {console.warn('Reflection e-mail failed: ' + (e.message || e));}
  }
  return {success: true, reflectionId: cpLink_t_(row['Reflection ID']), emailed: emailed};
}

function dashboardV2ReflectionDecide(request) {
  const r = request || {};
  const decision = r.decision === 'Accepted' ? 'Accepted' : r.decision === 'Returned' ? 'Returned' : '';
  if (!decision) throw new Error('Choose accept or return.');
  const notes = cpLink_requireText_(r.notes, 'Note', 2000);
  if (decision === 'Returned' && !notes) throw new Error('Add a note so the instructor knows what to change.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let row;
  try {
    row = cpRefl_rows_().filter(function(x) {return cpLink_t_(x['Reflection ID']) === cpLink_t_(r.reflectionId);})[0];
    if (!row) throw new Error('Reflection not found.');
    if (cpLink_t_(row['Status']) !== 'Submitted') throw new Error('Only submitted reflections can be reviewed.');
    row['Status'] = decision; row['Coordinator Note'] = notes; row['Reviewed At'] = new Date(); row['Reviewed By'] = cpLink_user_();
    cpRefl_write_(row);
  } finally {lock.releaseLock();}
  const id = cpLink_t_(row['Evaluator ID']), term = cpLink_t_(row['Term']);
  cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': id, 'Recipient': cpLink_t_(row['Evaluator']), 'Type': 'Reflection ' + decision,
    'Item Kind': 'reflection', 'Item ID': cpLink_t_(row['Reflection ID']),
    'Title': 'Reflection ' + (decision === 'Accepted' ? 'accepted' : 'returned to you') + ' — ' + term, 'Message': notes, 'From': cpLink_user_()});
  cpMile_markInboxRead_(cpLink_t_(row['Reflection ID']));
  try {
    const sh = cpLink_sheet_();
    cpLink_rows_().forEach(function(n) {
      if (cpLink_t_(n['Recipient Type']) === 'Coordinator' && cpLink_t_(n['Status']) === 'Unread' && cpLink_t_(n['Item ID']) === cpLink_t_(row['Reflection ID'])) {
        sh.getRange(n._row, CP_LINK.HEADERS.indexOf('Status') + 1, 1, 2).setValues([['Read', new Date()]]);
      }
    });
  } catch (e) {}
  let emailed = false;
  if (r.email !== false) {
    try {
      emailed = cpHours_email_(id, 'Your ' + term + ' reflection was ' + (decision === 'Accepted' ? 'accepted' : 'returned for changes'),
        (decision === 'Accepted' ? ['Thank you. Your ' + term + ' Clinical Instructor Summary Reflection was accepted.']
          : ['Your ' + term + ' reflection was returned for changes. Please update and resubmit it in the Clinical Instructor Portal (My CI Workspace → Survey Responses).'])
          .concat(notes ? ['', 'Coordinator note: ' + notes] : []));
    } catch (e) {console.warn('Reflection e-mail failed: ' + (e.message || e));}
  }
  return {success: true, emailed: emailed};
}
