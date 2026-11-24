/**
 * Clinical Instructor Portal — CI Notes / College Clinical Instructor Report,
 * and Incident Reports (Form 1B Instructor).
 *
 * CI Notes (coordinator decision 2026-11-24): instructors add dated notes as they
 * go — a comment on a student, a positioning lab / lecture activity, or a
 * communication with a hospital clinical instructor. At mid-semester and end of
 * semester they submit the College Clinical Instructor Report: the approved
 * template is copied, filled with every note in the chosen period, saved as a
 * Google Doc + PDF in a private folder, and the coordinator is notified.
 *
 * Incident Report: the instructor completes Form 1B; the approved template is
 * filled (Doc + PDF), the coordinator gets an urgent dashboard alert and an
 * e-mail, and the department chair is e-mailed.
 *
 * Shared CCMS sheets: "CI Notes", "CI Reports", "Incident Reports", "Program Settings".
 */
const CI_DOCS = {
  NOTES: 'CI Notes',
  NOTE_HEADERS: ['Note ID', 'Evaluator ID', 'Evaluator', 'Kind', 'Student ID', 'Student', 'Clinical Term', 'Date', 'Site', 'Text',
    'Created At', 'Updated At', 'Report ID', 'Status'],
  REPORTS: 'CI Reports',
  REPORT_HEADERS: ['Report ID', 'Evaluator ID', 'Evaluator', 'Term', 'Period', 'Clinical Practicum', 'Sites', 'From', 'To',
    'Student Notes', 'Status', 'Submitted At', 'Doc ID', 'PDF ID', 'Coordinator Note', 'Reviewed At', 'Reviewed By', 'Version'],
  INCIDENTS: 'Incident Reports',
  INCIDENT_HEADERS: ['Incident ID', 'Evaluator ID', 'Evaluator', 'Student ID', 'Student', 'Clinical Term', 'Incident Date', 'Incident Time',
    'Site', 'Location', 'Additional Persons', 'Witnesses', 'Description', 'Resulting Action', 'Status', 'Submitted At', 'Doc ID', 'PDF ID',
    'Coordinator Note', 'Reviewed At', 'Reviewed By', 'Emailed'],
  SETTINGS: 'Program Settings',
  // Approved templates (Word files in Drive). Override with script properties if replaced.
  REPORT_TEMPLATE: '1lGmclCoI8yp5FhiY3SpDNSjRk0_uONTH',   // CI Report Form Template Clinical 1-5 1_27_26.docx
  INCIDENT_TEMPLATE: '1yHa58KL6GtmPlRTe8zcDi0r5m1_zKKUD', // Incident Report Form 1B Instructor.docx
  FOLDER: 'CCMS — CI Notes, Reports & Incidents (restricted)'
};
const ciDocT_ = v => v == null ? '' : String(v).trim();
const ciDocIso_ = v => v instanceof Date ? v.toISOString() : ciDocT_(v);
const ciDocDay_ = v => ciDataDate_(v) || (/^\d{4}-\d{2}-\d{2}/.test(ciDocT_(v)) ? ciDocT_(v).slice(0, 10) : '');
function ciDocProp_(key, dflt) {return ciDocT_(PropertiesService.getScriptProperties().getProperty(key)) || dflt;}
function ciDocFolder_(path) {
  const root = DriveApp.getRootFolder();
  const it = root.getFoldersByName(CI_DOCS.FOLDER);
  let f = it.hasNext() ? it.next() : root.createFolder(CI_DOCS.FOLDER);
  (path || []).forEach(n => {const s = f.getFoldersByName(n); f = s.hasNext() ? s.next() : f.createFolder(n);});
  return f;
}
/** Copy an approved Word template into the folder as a Google Doc (Drive converts it). */
function ciDocFromWordTemplate_(templateId, name, folder) {
  const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(templateId) + '/copy?supportsAllDrives=true&fields=id', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
    payload: JSON.stringify({name: name, mimeType: 'application/vnd.google-apps.document', parents: [folder.getId()]})});
  if (res.getResponseCode() !== 200) throw new Error('Could not copy the approved template (' + res.getResponseCode() + ').');
  return JSON.parse(res.getContentText()).id;
}
function ciDocPdf_(docId, folder, name) {
  return folder.createFile(DriveApp.getFileById(docId).getBlob().getAs('application/pdf').setName(name + '.pdf')).getId();
}
function ciDocStudents_() {
  return ciDataRows_(ciDataBook_(), 'Students', ['Student ID', 'Student', 'Active'])
    .filter(s => ciDataTrue_(s.Active))
    .map(s => ({studentId: ciDocT_(s['Student ID']), student: ciDocT_(s.Student), site: ciDocT_(s['Clinical Site']),
      clinicalTerm: ciDocT_(s['Clinical Term']), cohort: ciDocT_(s.Cohort)}))
    .filter(s => s.studentId).sort((a, b) => a.student.localeCompare(b.student));
}
function ciDocSettings_() {
  const sh = ciMileSheet_(CI_DOCS.SETTINGS, ['Setting', 'Value']);
  const out = {}, rows = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach((r, i) => {if (ciDocT_(r[0])) {out[ciDocT_(r[0])] = ciDocT_(r[1]); rows[ciDocT_(r[0])] = i + 2;}});
  const defaults = {'Chair Name': 'Nicole Murray', 'Chair E-mail': '', 'Coordinator E-mail': ''};
  Object.keys(defaults).forEach(k => {if (!(k in out)) {sh.appendRow([k, defaults[k]]); out[k] = defaults[k]; rows[k] = sh.getLastRow();}});
  // Chair e-mail: filled once from the Evaluators sheet (same surname) when left blank.
  if (!out['Chair E-mail'] && out['Chair Name']) {
    const last = out['Chair Name'].split(/\s+/).pop().toLowerCase(), first = out['Chair Name'].charAt(0).toLowerCase();
    const ev = ciDataRows_(ciDataBook_(), 'Evaluators', ['Evaluator ID', 'Evaluator']).find(e => {
      const n = ciDocT_(e.Evaluator).toLowerCase(); return n.split(/\s+/).pop() === last && n.charAt(0) === first;});
    const em = ev ? ciDocT_(ev['E-mail']) : '';
    if (em) {sh.getRange(rows['Chair E-mail'], 2).setValue(em); out['Chair E-mail'] = em;}
  }
  if (!out['Coordinator E-mail']) {
    try {const me = Session.getEffectiveUser().getEmail(); if (me) {sh.getRange(rows['Coordinator E-mail'], 2).setValue(me); out['Coordinator E-mail'] = me;}} catch (e) {}
  }
  return out;
}
function ciDocNotify_(ev, type, kind, itemId, title, message) {
  try {
    const n = {'Notification ID': 'NT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 14).toUpperCase(), 'Created At': new Date(),
      'Recipient Type': 'Coordinator', 'Recipient': 'Clinical Coordinator', 'Type': type, 'Item Kind': kind, 'Item ID': itemId,
      'Title': title, 'Message': message, 'From': ev.evaluatorName + ' (' + ev.evaluatorId + ')', 'Status': 'Unread'};
    ciNtSheet_().appendRow(CI_NT_HEADERS.map(h => n[h] == null ? '' : n[h]));
  } catch (e) {console.warn('Coordinator notification failed: ' + (e.message || e));}
}
function ciDocRequireCollege_(ev) {
  if (ciDocT_(ev.evaluatorType).toLowerCase() !== 'college ci') throw new Error('CI Notes and Reports are for college clinical instructors.');
}

/* ============================ CI NOTES ============================ */
function ciDocNoteOut_(n) {
  return {noteId: ciDocT_(n['Note ID']), kind: ciDocT_(n.Kind), studentId: ciDocT_(n['Student ID']), student: ciDocT_(n.Student),
    clinicalTerm: ciDocT_(n['Clinical Term']), date: ciDocDay_(n.Date), site: ciDocT_(n.Site), text: ciDocT_(n.Text),
    reportId: ciDocT_(n['Report ID']), updatedAt: ciDocIso_(n['Updated At'] || n['Created At'])};
}
function ciDocReportOut_(r) {
  return {reportId: ciDocT_(r['Report ID']), term: ciDocT_(r.Term), period: ciDocT_(r.Period), practicum: ciDocT_(r['Clinical Practicum']),
    sites: ciDocT_(r.Sites), from: ciDocDay_(r.From), to: ciDocDay_(r.To), studentNotes: Number(r['Student Notes']) || 0,
    status: ciDocT_(r.Status), submittedAt: ciDocIso_(r['Submitted At']), hasPdf: !!ciDocT_(r['PDF ID']),
    coordinatorNote: ciDocT_(r['Coordinator Note']), version: Number(r.Version) || 1};
}
function ciR1NotesInit(token, term) {
  const ev = ciR1Require_(token);
  ciDocRequireCollege_(ev);
  const terms = ciMileTerms_();
  const range = ciMileTermRange_(term && /^(Spring|Summer|Fall) \d{4}$/.test(term) ? term : terms[0]);
  const notes = ciMileRows_(CI_DOCS.NOTES, CI_DOCS.NOTE_HEADERS)
    .filter(n => ciDocT_(n['Evaluator ID']) === ev.evaluatorId && ciDocT_(n.Status) !== 'Deleted')
    .map(ciDocNoteOut_).filter(n => n.date >= range.start && n.date <= range.end)
    .sort((a, b) => b.date.localeCompare(a.date) || b.updatedAt.localeCompare(a.updatedAt));
  const reports = ciMileRows_(CI_DOCS.REPORTS, CI_DOCS.REPORT_HEADERS)
    .filter(r => ciDocT_(r['Evaluator ID']) === ev.evaluatorId && ciDocT_(r.Term) === range.term).map(ciDocReportOut_);
  const locked = {};
  reports.forEach(r => {if (r.status !== 'Returned') locked[r.reportId] = true;});
  notes.forEach(n => {n.locked = !!(n.reportId && locked[n.reportId]);});
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return {success: true, evaluator: ev.evaluatorName, terms: terms, term: range.term, start: range.start, end: range.end, today: today,
    students: ciDocStudents_(), notes: notes, reports: reports};
}
function ciR1NoteSave(token, request) {
  const ev = ciR1Require_(token);
  ciDocRequireCollege_(ev);
  const q = request || {};
  const kind = ['Student', 'Activity', 'Mentoring'].indexOf(q.kind) >= 0 ? q.kind : '';
  if (!kind) throw new Error('Choose the kind of note.');
  const date = ciDocT_(q.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Choose the date.');
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (date > today) throw new Error('The date cannot be in the future.');
  const text = String(q.text == null ? '' : q.text).replace(/\r\n?/g, '\n').trim().slice(0, 6000);
  if (text.length < 3) throw new Error('Write the note first.');
  let student = null;
  if (kind === 'Student') {
    student = ciDocStudents_().find(s => s.studentId === ciDocT_(q.studentId));
    if (!student) throw new Error('Choose the student.');
  }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = ciMileSheet_(CI_DOCS.NOTES, CI_DOCS.NOTE_HEADERS);
    if (q.noteId) {
      const n = ciMileRows_(CI_DOCS.NOTES, CI_DOCS.NOTE_HEADERS).find(x => ciDocT_(x['Note ID']) === ciDocT_(q.noteId) && ciDocT_(x['Evaluator ID']) === ev.evaluatorId);
      if (!n || ciDocT_(n.Status) === 'Deleted') throw new Error('Note not found.');
      ciDocAssertEditable_(n);
      Object.assign(n, {'Kind': kind, 'Student ID': student ? student.studentId : '', 'Student': student ? student.student : '',
        'Clinical Term': student ? student.clinicalTerm : '', 'Date': date, 'Site': ciDocT_(q.site) || (student ? student.site : ''), 'Text': text, 'Updated At': new Date()});
      ciMileWrite_(CI_DOCS.NOTES, CI_DOCS.NOTE_HEADERS, n._row, n);
      return {success: true, noteId: ciDocT_(n['Note ID'])};
    }
    const id = 'CIN-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
    const row = {'Note ID': id, 'Evaluator ID': ev.evaluatorId, 'Evaluator': ev.evaluatorName, 'Kind': kind,
      'Student ID': student ? student.studentId : '', 'Student': student ? student.student : '', 'Clinical Term': student ? student.clinicalTerm : '',
      'Date': date, 'Site': ciDocT_(q.site) || (student ? student.site : ''), 'Text': text, 'Created At': new Date(), 'Status': 'Active'};
    sh.appendRow(CI_DOCS.NOTE_HEADERS.map(h => row[h] == null ? '' : row[h]));
    return {success: true, noteId: id};
  } finally {lock.releaseLock();}
}
function ciDocAssertEditable_(n) {
  const rid = ciDocT_(n['Report ID']);
  if (!rid) return;
  const r = ciMileRows_(CI_DOCS.REPORTS, CI_DOCS.REPORT_HEADERS).find(x => ciDocT_(x['Report ID']) === rid);
  if (r && ciDocT_(r.Status) !== 'Returned') throw new Error('This note is part of a submitted report and cannot be changed.');
}
function ciR1NoteDelete(token, noteId) {
  const ev = ciR1Require_(token);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const n = ciMileRows_(CI_DOCS.NOTES, CI_DOCS.NOTE_HEADERS).find(x => ciDocT_(x['Note ID']) === ciDocT_(noteId) && ciDocT_(x['Evaluator ID']) === ev.evaluatorId);
    if (!n) throw new Error('Note not found.');
    ciDocAssertEditable_(n);
    n.Status = 'Deleted'; n['Updated At'] = new Date();
    ciMileWrite_(CI_DOCS.NOTES, CI_DOCS.NOTE_HEADERS, n._row, n);
    return {success: true};
  } finally {lock.releaseLock();}
}

/** Compile the College Clinical Instructor Report for a period and submit it. */
function ciR1CiReportSubmit(token, request) {
  const ev = ciR1Require_(token);
  ciDocRequireCollege_(ev);
  const q = request || {};
  const range = ciMileTermRange_(q.term);
  const period = q.period === 'End of semester' ? 'End of semester' : q.period === 'Mid-semester' ? 'Mid-semester' : '';
  if (!period) throw new Error('Choose mid-semester or end of semester.');
  const practicum = /^Clinical [1-5]$/.test(ciDocT_(q.practicum)) ? ciDocT_(q.practicum) : '';
  if (!practicum) throw new Error('Choose the clinical practicum.');
  const from = ciDocT_(q.from), to = ciDocT_(q.to);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) throw new Error('Choose the period dates.');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const reports = ciMileRows_(CI_DOCS.REPORTS, CI_DOCS.REPORT_HEADERS)
      .filter(r => ciDocT_(r['Evaluator ID']) === ev.evaluatorId && ciDocT_(r.Term) === range.term && ciDocT_(r.Period) === period);
    const existing = reports[reports.length - 1] || null;
    if (existing && ciDocT_(existing.Status) !== 'Returned') throw new Error('Your ' + period.toLowerCase() + ' report for ' + range.term + ' was already submitted.');
    const otherLocked = {};
    ciMileRows_(CI_DOCS.REPORTS, CI_DOCS.REPORT_HEADERS).forEach(r => {
      if (ciDocT_(r['Evaluator ID']) === ev.evaluatorId && ciDocT_(r.Status) !== 'Returned' && (!existing || r._row !== existing._row)) otherLocked[ciDocT_(r['Report ID'])] = true;});
    const all = ciMileRows_(CI_DOCS.NOTES, CI_DOCS.NOTE_HEADERS).filter(n => ciDocT_(n['Evaluator ID']) === ev.evaluatorId && ciDocT_(n.Status) !== 'Deleted');
    const inPeriod = all.filter(n => {const d = ciDocDay_(n.Date); return d >= from && d <= to && !otherLocked[ciDocT_(n['Report ID'])];})
      .sort((a, b) => ciDocDay_(a.Date).localeCompare(ciDocDay_(b.Date)) || ciDocT_(a.Student).localeCompare(ciDocT_(b.Student)));
    const studentNotes = inPeriod.filter(n => ciDocT_(n.Kind) === 'Student');
    if (!studentNotes.length) throw new Error('There are no student notes between ' + from + ' and ' + to + '.');
    const activities = inPeriod.filter(n => ciDocT_(n.Kind) === 'Activity'), mentoring = inPeriod.filter(n => ciDocT_(n.Kind) === 'Mentoring');
    const sites = Array.from(new Set(inPeriod.map(n => ciDocT_(n.Site)).filter(Boolean))).join(', ');
    const id = existing ? ciDocT_(existing['Report ID']) : 'CIR-' + ciMileTermRange_(range.term).code + '-' + ev.evaluatorId + '-' + (period === 'Mid-semester' ? 'MID' : 'END');
    const version = existing ? (Number(existing.Version) || 1) + 1 : 1;
    const folder = ciDocFolder_(['CI Reports', range.term]);
    const name = ev.evaluatorName + ' — ' + practicum + ' ' + range.term + ' ' + period + ' CI Report' + (version > 1 ? ' (v' + version + ')' : '');
    const docId = ciDocFromWordTemplate_(ciDocProp_('CI_REPORT_TEMPLATE_ID', CI_DOCS.REPORT_TEMPLATE), name, folder);
    ciDocFillReport_(docId, {instructor: ev.evaluatorName, sites: sites, practicum: practicum, term: range.term, period: period, from: from, to: to,
      activities: activities, mentoring: mentoring, notes: studentNotes});
    const pdfId = ciDocPdf_(docId, folder, name);
    const row = existing || {};
    Object.assign(row, {'Report ID': id, 'Evaluator ID': ev.evaluatorId, 'Evaluator': ev.evaluatorName, 'Term': range.term, 'Period': period,
      'Clinical Practicum': practicum, 'Sites': sites, 'From': from, 'To': to, 'Student Notes': studentNotes.length, 'Status': 'Submitted',
      'Submitted At': new Date(), 'Doc ID': docId, 'PDF ID': pdfId, 'Version': version});
    if (existing) ciMileWrite_(CI_DOCS.REPORTS, CI_DOCS.REPORT_HEADERS, existing._row, row);
    else ciMileAppend_(CI_DOCS.REPORTS, CI_DOCS.REPORT_HEADERS, row);
    const sh = ciMileSheet_(CI_DOCS.NOTES, CI_DOCS.NOTE_HEADERS), col = CI_DOCS.NOTE_HEADERS.indexOf('Report ID') + 1;
    inPeriod.forEach(n => sh.getRange(n._row, col).setValue(id));
    ciDocNotify_(ev, 'CI Report Submitted', 'ci-report', id, (version > 1 ? 'CI report resubmitted — ' : 'CI report submitted — ') + ev.evaluatorName + ' · ' + practicum + ' · ' + period,
      range.term + ' · ' + from + ' to ' + to + ' · ' + studentNotes.length + ' student note' + (studentNotes.length === 1 ? '' : 's') +
      ' · ' + activities.length + ' activit' + (activities.length === 1 ? 'y' : 'ies') + ' · ' + mentoring.length + ' hospital CI contact' + (mentoring.length === 1 ? '' : 's'));
    return {success: true, reportId: id, version: version, studentNotes: studentNotes.length};
  } finally {lock.releaseLock();}
}
function ciDocFillReport_(docId, d) {
  const doc = DocumentApp.openById(docId), body = doc.getBody();
  const plain = {};
  plain[DocumentApp.Attribute.ITALIC] = false; plain[DocumentApp.Attribute.BOLD] = false; plain[DocumentApp.Attribute.UNDERLINE] = false;
  const paraIdx = needle => {
    for (let i = 0; i < body.getNumChildren(); i++) {
      const ch = body.getChild(i);
      if (ch.getType() === DocumentApp.ElementType.PARAGRAPH && ch.asParagraph().getText().indexOf(needle) >= 0) return i;
    }
    return -1;
  };
  const fmt = iso => {const p = iso.split('-'); return Number(p[1]) + '/' + Number(p[2]) + '/' + p[0].slice(2);};
  let i = paraIdx('Clinical Site:');
  if (i >= 0) body.getChild(i).asParagraph().setText('Clinical Site: ' + (d.sites || '—') + '\t\tInstructor: ' + d.instructor);
  i = paraIdx('Clinical Practicum (');
  if (i >= 0) body.getChild(i).asParagraph().setText('Clinical Practicum (' + d.practicum.toLowerCase() + '), (' + d.term.replace(' ', ' & ') + ') — ' + d.period + ' report, ' + fmt(d.from) + ' – ' + fmt(d.to));
  // Mentoring first (lower in the document) so the activities index stays valid.
  const insertList = (anchor, items) => {
    const at = paraIdx(anchor);
    const lines = items.length ? items.map(n => fmt(ciDocDay_(n.Date)) + ' — ' + ciDocT_(n.Text).replace(/\s*\n\s*/g, ' ')) : ['None recorded for this period.'];
    if (at < 0) {lines.forEach(l => body.appendParagraph(l).setAttributes(plain)); return;}
    lines.forEach((l, k) => body.insertParagraph(at + 1 + k, '• ' + l).setAttributes(plain));
  };
  insertList('Communication & Mentoring of Hospital Clinical Instructors', d.mentoring);
  insertList('Examples: Development of Technique Book', d.activities);
  // Student comments table: header row + one row per note.
  const tables = body.getTables();
  const table = tables.find(t => t.getNumRows() && /STUDENT NAME/i.test(t.getRow(0).getText())) || tables[0];
  if (table) {
    const rows = d.notes.map(n => [ciDocT_(n.Student), fmt(ciDocDay_(n.Date)), ciDocT_(n.Text)]);
    let r = 1;
    rows.forEach(vals => {
      let tr;
      if (r < table.getNumRows()) tr = table.getRow(r);
      else {tr = table.appendTableRow(); for (let c = 0; c < 3; c++) tr.appendTableCell('');}
      for (let c = 0; c < Math.min(3, tr.getNumCells()); c++) tr.getCell(c).setText(vals[c]).setAttributes(plain);
      r++;
    });
    while (table.getNumRows() > r && table.getNumRows() > 2) table.removeRow(table.getNumRows() - 1);
  }
  doc.saveAndClose();
}
function ciR1CiReportPdf(token, reportId) {
  const ev = ciR1Require_(token);
  const r = ciMileRows_(CI_DOCS.REPORTS, CI_DOCS.REPORT_HEADERS).find(x => ciDocT_(x['Report ID']) === ciDocT_(reportId) && ciDocT_(x['Evaluator ID']) === ev.evaluatorId);
  if (!r || !ciDocT_(r['PDF ID'])) throw new Error('Report not found.');
  return {success: true, name: ciDocT_(r['Report ID']) + '.pdf', base64: Utilities.base64Encode(DriveApp.getFileById(ciDocT_(r['PDF ID'])).getBlob().getBytes())};
}

/* ============================ INCIDENTS ============================ */
function ciR1IncidentInit(token) {
  const ev = ciR1Require_(token);
  const mine = ciMileRows_(CI_DOCS.INCIDENTS, CI_DOCS.INCIDENT_HEADERS).filter(r => ciDocT_(r['Evaluator ID']) === ev.evaluatorId)
    .map(r => ({incidentId: ciDocT_(r['Incident ID']), student: ciDocT_(r.Student), date: ciDocDay_(r['Incident Date']), site: ciDocT_(r.Site),
      status: ciDocT_(r.Status), submittedAt: ciDocIso_(r['Submitted At']), coordinatorNote: ciDocT_(r['Coordinator Note']), hasPdf: !!ciDocT_(r['PDF ID'])}))
    .sort((a, b) => (b.submittedAt || '').localeCompare(a.submittedAt || ''));
  return {success: true, evaluator: ev.evaluatorName, students: ciDocStudents_(), incidents: mine,
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')};
}
function ciR1IncidentSubmit(token, request) {
  const ev = ciR1Require_(token);
  const q = request || {};
  const t = (v, max) => String(v == null ? '' : v).replace(/\r\n?/g, '\n').trim().slice(0, max || 500);
  const student = ciDocStudents_().find(s => s.studentId === ciDocT_(q.studentId));
  if (!student) throw new Error('Choose the student.');
  const date = ciDocT_(q.date), today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date > today) throw new Error('Enter the date of the incident.');
  const time = t(q.time, 20);
  if (!time) throw new Error('Enter the time of the incident.');
  const description = t(q.description, 10000), action = t(q.action, 6000);
  if (description.length < 20) throw new Error('Describe the incident (a few sentences).');
  const rec = {'Incident ID': 'INC-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase(),
    'Evaluator ID': ev.evaluatorId, 'Evaluator': ev.evaluatorName, 'Student ID': student.studentId, 'Student': student.student,
    'Clinical Term': student.clinicalTerm, 'Incident Date': date, 'Incident Time': time, 'Site': t(q.site, 120) || student.site,
    'Location': t(q.location, 300), 'Additional Persons': t(q.persons, 500), 'Witnesses': t(q.witnesses, 500),
    'Description': description, 'Resulting Action': action, 'Status': 'Submitted', 'Submitted At': new Date()};
  const folder = ciDocFolder_(['Incident Reports', student.student]);
  const name = student.student + ' — Incident Report ' + date + ' (' + ev.evaluatorName + ')';
  const docId = ciDocFromWordTemplate_(ciDocProp_('INCIDENT_TEMPLATE_ID', CI_DOCS.INCIDENT_TEMPLATE), name, folder);
  ciDocFillIncident_(docId, rec);
  rec['Doc ID'] = docId; rec['PDF ID'] = ciDocPdf_(docId, folder, name);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {ciMileAppend_(CI_DOCS.INCIDENTS, CI_DOCS.INCIDENT_HEADERS, rec);} finally {lock.releaseLock();}
  ciDocNotify_(ev, 'Incident Report', 'incident', rec['Incident ID'], 'INCIDENT REPORT — ' + student.student + ' · ' + rec.Site + ' · ' + date,
    description.slice(0, 400) + (description.length > 400 ? '…' : ''));
  // E-mail the coordinator and the department chair with the PDF attached.
  const sent = [];
  try {
    const s = ciDocSettings_();
    const to = Array.from(new Set([s['Coordinator E-mail'], s['Chair E-mail']].filter(x => /@/.test(x || ''))));
    if (to.length) {
      const pdf = DriveApp.getFileById(rec['PDF ID']).getBlob();
      MailApp.sendEmail({to: to.join(','), subject: 'Incident report: ' + student.student + ' — ' + rec.Site + ' — ' + date,
        body: ['An incident report was submitted in the Clinical Instructor Portal.', '',
          'Student: ' + student.student, 'Instructor: ' + ev.evaluatorName, 'Date/time: ' + date + ' ' + time, 'Site: ' + rec.Site + (rec.Location ? ' · ' + rec.Location : ''),
          '', 'Description:', description, '', action ? 'Resulting action / consequence:\n' + action : '', '',
          'The completed Form 1B (PDF) is attached. Review it in the Coordinator Dashboard → Reports & Paperwork → CI Notes & Incidents.'].join('\n'),
        attachments: [pdf], name: 'Clinical Evaluation System'});
      sent.push.apply(sent, to);
    }
  } catch (e) {console.warn('Incident e-mail failed: ' + (e.message || e));}
  if (sent.length) {
    const r = ciMileRows_(CI_DOCS.INCIDENTS, CI_DOCS.INCIDENT_HEADERS).find(x => ciDocT_(x['Incident ID']) === rec['Incident ID']);
    if (r) {r.Emailed = sent.length + ' recipient' + (sent.length === 1 ? '' : 's') + ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd HH:mm'); ciMileWrite_(CI_DOCS.INCIDENTS, CI_DOCS.INCIDENT_HEADERS, r._row, r);}
  }
  return {success: true, incidentId: rec['Incident ID'], emailed: sent.length};
}
function ciDocFillIncident_(docId, rec) {
  const doc = DocumentApp.openById(docId), body = doc.getBody();
  const fmt = iso => {const p = iso.split('-'); return Number(p[1]) + '/' + Number(p[2]) + '/' + p[0];};
  const values = [
    [/^STUDENT:/i, rec.Student], [/^DATE OF INCIDENT/i, fmt(rec['Incident Date'])], [/^TIME OF INCIDENT/i, rec['Incident Time']],
    [/^SPECIFIC AREA/i, [rec.Site, rec.Location].filter(Boolean).join(' — ')], [/^ADDITIONAL PERSON/i, rec['Additional Persons'] || 'None'],
    [/^WITNESSES/i, rec.Witnesses || 'None'], [/^INSTRUCTOR.?S NAME/i, rec.Evaluator]];
  const big = [[/^INCIDENT DESCRIPTION/i, rec.Description], [/^RESULTING ACTION/i, rec['Resulting Action'] || '']];
  const table = body.getTables()[0];
  if (!table) throw new Error('The incident template layout was not recognised.');
  const span = c => {try {return c.getColSpan();} catch (e) {return 1;}};
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const label = row.getCell(c).getText().trim();
      const v = values.find(x => x[0].test(label));
      if (v) {
        let filled = false;
        for (let k = c + 1; k < row.getNumCells(); k++) {
          const cell = row.getCell(k);
          if (span(cell) === 0) continue;
          if (!cell.getText().trim()) {cell.setText(v[1]); filled = true; break;}
          break;
        }
        if (!filled) row.getCell(c).setText(label + ' ' + v[1]);
        continue;
      }
      const b = big.find(x => x[0].test(label));
      if (b && b[1]) {
        const next = r + 1 < table.getNumRows() ? table.getRow(r + 1) : null;
        if (next && next.getNumCells() && !next.getText().trim()) next.getCell(0).setText(b[1]);
        else row.getCell(c).appendParagraph(b[1]);
      }
    }
  }
  // Instructor's date (the cell after the last "DATE:" on the instructor row).
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (!/INSTRUCTOR.?S NAME/i.test(row.getText())) continue;
    for (let c = 0; c < row.getNumCells(); c++) {
      if (/^DATE:?$/i.test(row.getCell(c).getText().trim()) && c + 1 < row.getNumCells() && !row.getCell(c + 1).getText().trim()) {
        row.getCell(c + 1).setText(fmt(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')));
      }
    }
  }
  doc.saveAndClose();
}
function ciR1IncidentPdf(token, incidentId) {
  const ev = ciR1Require_(token);
  const r = ciMileRows_(CI_DOCS.INCIDENTS, CI_DOCS.INCIDENT_HEADERS).find(x => ciDocT_(x['Incident ID']) === ciDocT_(incidentId) && ciDocT_(x['Evaluator ID']) === ev.evaluatorId);
  if (!r || !ciDocT_(r['PDF ID'])) throw new Error('Incident report not found.');
  return {success: true, name: ciDocT_(r['Incident ID']) + '.pdf', base64: Utilities.base64Encode(DriveApp.getFileById(ciDocT_(r['PDF ID'])).getBlob().getBytes())};
}

/** Run once from the editor: grants the e-mail permission and builds SAMPLE documents (no records) to check the templates. */
function ciNotesIncidentsAuthorize() {
  const folder = ciDocFolder_(['Samples']);
  const rid = ciDocFromWordTemplate_(ciDocProp_('CI_REPORT_TEMPLATE_ID', CI_DOCS.REPORT_TEMPLATE), 'SAMPLE — CI Report layout check', folder);
  ciDocFillReport_(rid, {instructor: 'Sample Instructor', sites: 'Sample Site', practicum: 'Clinical 3', term: 'Fall 2026', period: 'Mid-semester',
    from: '2026-09-01', to: '2026-10-15', activities: [{Date: '2026-09-10', Text: 'Positioning lab: knee series'}],
    mentoring: [{Date: '2026-09-12', Text: 'Met hospital CI about competency expectations'}],
    notes: [{Student: 'Sample Student A', Date: '2026-09-03', Text: 'Sample comment one.'}, {Student: 'Sample Student B', Date: '2026-09-04', Text: 'Sample comment two.'}]});
  const iid = ciDocFromWordTemplate_(ciDocProp_('INCIDENT_TEMPLATE_ID', CI_DOCS.INCIDENT_TEMPLATE), 'SAMPLE — Incident Report layout check', folder);
  ciDocFillIncident_(iid, {Student: 'Sample Student', 'Incident Date': '2026-11-24', 'Incident Time': '10:15 AM', Site: 'Sample Site', Location: 'X-ray room 2',
    'Additional Persons': 'None', Witnesses: 'Sample Technologist', Description: 'Sample description of the event.', 'Resulting Action': 'Sample resulting action.', Evaluator: 'Sample Instructor'});
  const r = ciDocPdf_(rid, folder, 'SAMPLE — CI Report layout check'), i = ciDocPdf_(iid, folder, 'SAMPLE — Incident Report layout check');
  const s = ciDocSettings_();
  console.log('Samples: report PDF ' + r + ' · incident PDF ' + i + ' · chair e-mail ' + (s['Chair E-mail'] ? 'found' : 'missing') + ' · e-mail quota ' + MailApp.getRemainingDailyQuota());
  return [r, i];
}
