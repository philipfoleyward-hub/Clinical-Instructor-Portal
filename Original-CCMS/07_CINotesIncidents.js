/**
 * 07_CINotesIncidents.gs — Coordinator side of CI Notes / College Clinical
 * Instructor Reports and Incident Reports (Form 1B Instructor).
 *
 * Coordinator decisions 2026-11-24:
 *   - CI Notes: instructors add dated notes as they go in the CI Portal; at mid-
 *     and end of semester they submit the College Clinical Instructor Report,
 *     which compiles the notes into the approved template (Doc + PDF). The
 *     coordinator accepts or returns it here. Returning unlocks the notes so the
 *     instructor can fix them and resubmit.
 *   - Incident Report: the Portal fills Form 1B, raises an urgent alert in the
 *     coordinator bell and on the home page, and e-mails the coordinator and the
 *     department chair. The coordinator marks it reviewed here (with a note).
 * Shared sheets (also written by the CI Portal): "CI Notes", "CI Reports",
 * "Incident Reports".
 */
const CP_DOCS = {
  NOTES: 'CI Notes',
  NOTE_HEADERS: ['Note ID', 'Evaluator ID', 'Evaluator', 'Kind', 'Student ID', 'Student', 'Clinical Term', 'Date', 'Site', 'Text',
    'Created At', 'Updated At', 'Report ID', 'Status'],
  REPORTS: 'CI Reports',
  REPORT_HEADERS: ['Report ID', 'Evaluator ID', 'Evaluator', 'Term', 'Period', 'Clinical Practicum', 'Sites', 'From', 'To',
    'Student Notes', 'Status', 'Submitted At', 'Doc ID', 'PDF ID', 'Coordinator Note', 'Reviewed At', 'Reviewed By', 'Version'],
  INCIDENTS: 'Incident Reports',
  INCIDENT_HEADERS: ['Incident ID', 'Evaluator ID', 'Evaluator', 'Student ID', 'Student', 'Clinical Term', 'Incident Date', 'Incident Time',
    'Site', 'Location', 'Additional Persons', 'Witnesses', 'Description', 'Resulting Action', 'Status', 'Submitted At', 'Doc ID', 'PDF ID',
    'Coordinator Note', 'Reviewed At', 'Reviewed By', 'Emailed']
};
function cpDocs_url_(id, kind) {
  id = cpLink_t_(id);
  return id ? (kind === 'doc' ? 'https://docs.google.com/document/d/' + id + '/edit' : 'https://drive.google.com/file/d/' + id + '/view') : '';
}
function cpDocs_day_(v) {return cpMile_iso_(v) || (/^\d{4}-\d{2}-\d{2}/.test(cpLink_t_(v)) ? cpLink_t_(v).slice(0, 10) : '');}
function cpDocs_reportOut_(r) {
  return {reportId: cpLink_t_(r['Report ID']), evaluatorId: cpLink_t_(r['Evaluator ID']), evaluator: cpLink_t_(r['Evaluator']),
    term: cpLink_t_(r['Term']), period: cpLink_t_(r['Period']), practicum: cpLink_t_(r['Clinical Practicum']), sites: cpLink_t_(r['Sites']),
    from: cpDocs_day_(r['From']), to: cpDocs_day_(r['To']), studentNotes: cpMile_n_(r['Student Notes']) || 0,
    status: cpLink_t_(r['Status']) || 'Submitted', submittedAt: cpLink_iso_(r['Submitted At']),
    docUrl: cpDocs_url_(r['Doc ID'], 'doc'), pdfUrl: cpDocs_url_(r['PDF ID']), coordinatorNote: cpLink_t_(r['Coordinator Note']),
    reviewedAt: cpLink_iso_(r['Reviewed At']), version: cpMile_n_(r['Version']) || 1};
}
function cpDocs_incidentOut_(r) {
  return {incidentId: cpLink_t_(r['Incident ID']), evaluatorId: cpLink_t_(r['Evaluator ID']), evaluator: cpLink_t_(r['Evaluator']),
    studentId: cpLink_t_(r['Student ID']), student: cpLink_t_(r['Student']), clinicalTerm: cpLink_t_(r['Clinical Term']),
    date: cpDocs_day_(r['Incident Date']), time: cpLink_t_(r['Incident Time']), site: cpLink_t_(r['Site']), location: cpLink_t_(r['Location']),
    persons: cpLink_t_(r['Additional Persons']), witnesses: cpLink_t_(r['Witnesses']), description: cpLink_t_(r['Description']),
    action: cpLink_t_(r['Resulting Action']), status: cpLink_t_(r['Status']) || 'Submitted', submittedAt: cpLink_iso_(r['Submitted At']),
    docUrl: cpDocs_url_(r['Doc ID'], 'doc'), pdfUrl: cpDocs_url_(r['PDF ID']), coordinatorNote: cpLink_t_(r['Coordinator Note']),
    reviewedAt: cpLink_iso_(r['Reviewed At']), emailed: cpLink_t_(r['Emailed'])};
}
/** Incidents not yet reviewed, newest first (home-page banner). */
function dashboardV2OpenIncidents() {
  return {success: true, incidents: cpMile_rows_(CP_DOCS.INCIDENTS, CP_DOCS.INCIDENT_HEADERS).map(cpDocs_incidentOut_)
    .filter(function(x) {return x.status !== 'Reviewed';})
    .sort(function(a, b) {return (b.submittedAt || '').localeCompare(a.submittedAt || '');})
    .map(function(x) {return {incidentId: x.incidentId, student: x.student, site: x.site, date: x.date, evaluator: x.evaluator};})};
}
/** CI Notes & Incidents page: reports and notes for a semester, plus all incidents. */
function dashboardV2CiDocs(request) {
  const q = request || {};
  const terms = cpMile_terms_();
  const range = cpMile_termRange_(/^(Spring|Summer|Fall) \d{4}$/.test(cpLink_t_(q.term)) ? q.term : terms[0]);
  const reports = cpMile_rows_(CP_DOCS.REPORTS, CP_DOCS.REPORT_HEADERS).map(cpDocs_reportOut_)
    .filter(function(r) {return r.term === range.term;})
    .sort(function(a, b) {return (b.submittedAt || '').localeCompare(a.submittedAt || '');});
  const notes = cpMile_rows_(CP_DOCS.NOTES, CP_DOCS.NOTE_HEADERS)
    .filter(function(n) {const d = cpDocs_day_(n['Date']); return cpLink_t_(n['Status']) !== 'Deleted' && d >= range.start && d <= range.end;})
    .map(function(n) {return {noteId: cpLink_t_(n['Note ID']), evaluatorId: cpLink_t_(n['Evaluator ID']), evaluator: cpLink_t_(n['Evaluator']),
      kind: cpLink_t_(n['Kind']), studentId: cpLink_t_(n['Student ID']), student: cpLink_t_(n['Student']), date: cpDocs_day_(n['Date']),
      site: cpLink_t_(n['Site']), text: cpLink_t_(n['Text']), reportId: cpLink_t_(n['Report ID'])};})
    .sort(function(a, b) {return b.date.localeCompare(a.date) || a.evaluator.localeCompare(b.evaluator);});
  const incidents = cpMile_rows_(CP_DOCS.INCIDENTS, CP_DOCS.INCIDENT_HEADERS).map(cpDocs_incidentOut_)
    .sort(function(a, b) {return (a.status === 'Reviewed') - (b.status === 'Reviewed') || (b.submittedAt || '').localeCompare(a.submittedAt || '');});
  return {success: true, terms: terms, term: range.term, reports: reports, notes: notes, incidents: incidents};
}
function cpDocs_markRead_(itemId) {
  try {
    const sh = cpLink_sheet_();
    cpLink_rows_().forEach(function(n) {
      if (cpLink_t_(n['Recipient Type']) === 'Coordinator' && cpLink_t_(n['Status']) === 'Unread' && cpLink_t_(n['Item ID']) === itemId) {
        sh.getRange(n._row, CP_LINK.HEADERS.indexOf('Status') + 1, 1, 2).setValues([['Read', new Date()]]);
      }
    });
  } catch (e) {}
}
/** Accept or return a College Clinical Instructor Report. Returning unlocks its notes in the Portal. */
function dashboardV2CiReportDecide(request) {
  const r = request || {};
  const decision = r.decision === 'Accepted' ? 'Accepted' : r.decision === 'Returned' ? 'Returned' : '';
  if (!decision) throw new Error('Choose accept or return.');
  const notes = cpLink_requireText_(r.notes, 'Note', 2000);
  if (decision === 'Returned' && !notes) throw new Error('Add a note so the instructor knows what to change.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let row;
  try {
    row = cpMile_rows_(CP_DOCS.REPORTS, CP_DOCS.REPORT_HEADERS).filter(function(x) {return cpLink_t_(x['Report ID']) === cpLink_t_(r.reportId);})[0];
    if (!row) throw new Error('Report not found.');
    if (cpLink_t_(row['Status']) !== 'Submitted') throw new Error('Only submitted reports can be reviewed.');
    row['Status'] = decision; row['Coordinator Note'] = notes; row['Reviewed At'] = new Date(); row['Reviewed By'] = cpLink_user_();
    cpMile_write_(CP_DOCS.REPORTS, CP_DOCS.REPORT_HEADERS, row);
  } finally {lock.releaseLock();}
  const id = cpLink_t_(row['Evaluator ID']), label = cpLink_t_(row['Term']) + ' ' + cpLink_t_(row['Period']).toLowerCase() + ' CI report';
  cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': id, 'Recipient': cpLink_t_(row['Evaluator']), 'Type': 'CI Report ' + decision,
    'Item Kind': 'ci-report', 'Item ID': cpLink_t_(row['Report ID']),
    'Title': (decision === 'Accepted' ? 'CI report accepted — ' : 'CI report returned to you — ') + label,
    'Message': (decision === 'Returned' ? 'Update your notes in My CI Workspace → CI Notes & Reports and resubmit.\n' : '') + notes, 'From': cpLink_user_()});
  cpDocs_markRead_(cpLink_t_(row['Report ID']));
  let emailed = false;
  if (r.email !== false) {
    try {
      emailed = cpHours_email_(id, 'Your ' + label + ' was ' + (decision === 'Accepted' ? 'accepted' : 'returned for changes'),
        (decision === 'Accepted' ? ['Thank you. Your ' + label + ' was accepted.']
          : ['Your ' + label + ' was returned for changes. Please update your notes and resubmit it in the Clinical Instructor Portal (My CI Workspace → CI Notes & Reports).'])
          .concat(notes ? ['', 'Coordinator note: ' + notes] : []));
    } catch (e) {console.warn('CI report e-mail failed: ' + (e.message || e));}
  }
  return {success: true, emailed: emailed};
}
/** Mark an incident report reviewed, with the coordinator's follow-up note. */
function dashboardV2IncidentReview(request) {
  const r = request || {};
  const notes = cpLink_requireText_(r.notes, 'Follow-up note', 4000);
  if (!notes) throw new Error('Add a short follow-up note (what was done).');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let row;
  try {
    row = cpMile_rows_(CP_DOCS.INCIDENTS, CP_DOCS.INCIDENT_HEADERS).filter(function(x) {return cpLink_t_(x['Incident ID']) === cpLink_t_(r.incidentId);})[0];
    if (!row) throw new Error('Incident report not found.');
    if (cpLink_t_(row['Status']) === 'Reviewed') throw new Error('This incident was already marked reviewed.');
    row['Status'] = 'Reviewed'; row['Coordinator Note'] = notes; row['Reviewed At'] = new Date(); row['Reviewed By'] = cpLink_user_();
    cpMile_write_(CP_DOCS.INCIDENTS, CP_DOCS.INCIDENT_HEADERS, row);
  } finally {lock.releaseLock();}
  const id = cpLink_t_(row['Evaluator ID']);
  if (r.notifyCi !== false) {
    cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': id, 'Recipient': cpLink_t_(row['Evaluator']), 'Type': 'Incident Reviewed',
      'Item Kind': 'incident', 'Item ID': cpLink_t_(row['Incident ID']), 'Student ID': cpLink_t_(row['Student ID']), 'Student': cpLink_t_(row['Student']),
      'Title': 'Incident report reviewed — ' + cpLink_t_(row['Student']) + ' · ' + cpDocs_day_(row['Incident Date']), 'Message': notes, 'From': cpLink_user_()});
  }
  cpDocs_markRead_(cpLink_t_(row['Incident ID']));
  return {success: true};
}
