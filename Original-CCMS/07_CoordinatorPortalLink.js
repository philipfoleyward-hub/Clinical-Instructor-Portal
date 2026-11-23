/**
 * 07_CoordinatorPortalLink.gs — Coordinator Dashboard ⇄ Clinical Instructor Portal.
 *
 * One shared, append-only log ("Portal Notifications") carries everything the
 * coordinator sends an instructor (accepted / returned with notes) and every
 * instructor reply. Clinical Tracking review columns stay owned by
 * 04_ClinicalReviewService (flag, notes, reviewed); this file only orchestrates
 * them and records the message. Professional Competency records (Grade
 * Records) have no review columns, so their review state is derived from the
 * latest coordinator decision in this log.
 */
const CP_LINK = {
  SHEET: 'Portal Notifications',
  HEADERS: ['Notification ID', 'Created At', 'Recipient Type', 'Recipient ID', 'Recipient',
    'Type', 'Item Kind', 'Item ID', 'Student ID', 'Student', 'Title', 'Message', 'From',
    'Status', 'Read At', 'Response', 'Responded At']
};

function cpLink_sheet_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CP_LINK.SHEET);
  if (!sh) {
    sh = ss.insertSheet(CP_LINK.SHEET);
    sh.getRange(1, 1, 1, CP_LINK.HEADERS.length).setValues([CP_LINK.HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function cpLink_rows_(ss) {
  const sh = cpLink_sheet_(ss), last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(1, 1, last, CP_LINK.HEADERS.length).getValues();
  const head = values[0].map(String);
  return values.slice(1).map(function(r, i) {
    const o = {_row: i + 2};
    head.forEach(function(h, j) {o[h] = r[j];});
    return o;
  });
}
function cpLink_t_(v) {return v == null ? '' : String(v).trim();}
function cpLink_iso_(v) {return v instanceof Date ? v.toISOString() : cpLink_t_(v);}
function cpLink_add_(n) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = cpLink_sheet_();
    const id = 'NT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 14).toUpperCase();
    const row = {'Notification ID': id, 'Created At': new Date(), 'Status': 'Unread'};
    Object.keys(n).forEach(function(k) {row[k] = n[k];});
    sh.appendRow(CP_LINK.HEADERS.map(function(h) {return row[h] == null ? '' : row[h];}));
    return id;
  } finally {lock.releaseLock();}
}
function cpLink_user_() {
  let email = '';
  try {email = String(Session.getActiveUser().getEmail() || '');} catch (e) {}
  return email || 'Coordinator';
}

/* ------------------------------------------------------------------ */
/* Review items: Clinical Tracking evaluations + Professional Competency */
/* ------------------------------------------------------------------ */
function cpLink_pcItems_(ss, log) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const recs = appSheetToObjects_(appGetSheet_(ss, gradeRecords_sheetName_()));
  const decisions = {};
  (log || cpLink_rows_(ss)).forEach(function(n) {
    if (cpLink_t_(n['Item Kind']) !== 'professional' || cpLink_t_(n['Recipient Type']) !== 'CI') return;
    const t = cpLink_t_(n['Type']);
    if (t === 'Accepted' || t === 'Returned') decisions[cpLink_t_(n['Item ID'])] = n;
  });
  return recs.filter(function(r) {
    return appNormalize_(r['Grade Category']) === 'professional competency' && appIsTrue_(r['Active']) &&
      /^PC-/.test(cpLink_t_(r['Source ID']));
  }).map(function(r) {
    const notes = cpLink_t_(r['Notes']);
    const m = /Evaluator ID:\s*([^\n]+)/.exec(notes), site = /Clinical site:\s*([^\n]+)/.exec(notes),
      res = /Result:\s*([^\n]+)/.exec(notes), rub = /(?:Form )?rubric:\s*([^\n]+)/i.exec(notes);
    const id = cpLink_t_(r['Grade Record ID']), d = decisions[id];
    const date = r['Date'] instanceof Date ? r['Date'] : new Date(r['Date']);
    return {kind: 'professional', itemId: id, submissionId: cpLink_t_(r['Source ID']),
      studentId: cpLink_t_(r['Student ID']), student: cpLink_t_(r['Student']),
      clinicalTerm: cpLink_t_(r['Clinical Term']), examination: 'Professional Competency',
      anatomicalArea: rub ? (rub[1].trim() === 'Clinical 1-2' ? 'Clinical Practicum 1 & 2' : 'Clinical Practicum 3 - 5') : '',
      evaluationType: 'Professional Competency', clinicalSite: site ? site[1].trim() : '',
      evaluator: cpLink_t_(r['Instructor']), evaluatorId: m ? m[1].trim() : '',
      date: isNaN(date) ? '' : Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      dateValue: isNaN(date) ? 0 : date.getTime(),
      gradeDisplay: r['Percent Grade'] === '' ? '' : Number(r['Percent Grade']).toFixed(2) + '%',
      result: res ? res[1].trim() : '', sourceDocument: cpLink_t_(r['Source Document']),
      source: cpLink_t_(r['Source']),
      coordinatorReviewStatus: d && cpLink_t_(d['Type']) === 'Accepted' ? 'Reviewed' : 'Pending',
      flagged: !!(d && cpLink_t_(d['Type']) === 'Returned'),
      coordinatorNotes: d ? cpLink_t_(d['Message']) : ''};
  }).sort(function(a, b) {return b.dateValue - a.dateValue;});
}
function dashboardV2ProfessionalCompetencies() {
  return cpLink_pcItems_(SpreadsheetApp.getActiveSpreadsheet());
}
function cpLink_trackingRecord_(submissionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING));
  const rec = rows.filter(function(r) {return cpLink_t_(r['Submission ID']) === cpLink_t_(submissionId);})[0];
  if (!rec) throw new Error('Evaluation not found: ' + submissionId);
  return rec;
}

/* ------------------------------------------------------------------ */
/* Dashboard V2 Home                                                   */
/* ------------------------------------------------------------------ */
function dashboardV2GetHome() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const email = cpLink_user_();
  let name = '';
  try {
    const ev = appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.EVALUATORS)).filter(function(e) {
      return appNormalize_(e['E-mail']) === appNormalize_(email);
    })[0];
    if (ev) name = cpLink_t_(ev['Evaluator']);
  } catch (e) {}
  if (!name && email.indexOf('@') > 0) name = email.split('@')[0];
  let term = '';
  try {term = cpLink_t_(clinicalTerm_getForDate(new Date()));} catch (e) {}
  const log = cpLink_rows_(ss);
  const tracking = getReviewQueue('all').slice(0, 60).map(function(x) {
    return {kind: 'evaluation', itemId: x.submissionId, submissionId: x.submissionId, studentId: x.studentId,
      student: x.student, examination: x.examination, evaluationType: x.evaluationType,
      clinicalSite: x.clinicalSite, evaluator: x.evaluator, date: x.date, dateValue: x.dateValue,
      gradeDisplay: x.gradeDisplay, result: x.result, sourceDocument: x.sourceDocument,
      coordinatorReviewStatus: x.coordinatorReviewStatus || 'Pending', flagged: !!x.flagged};
  });
  const pcs = cpLink_pcItems_(ss, log);
  const replies = log.filter(function(n) {
    return cpLink_t_(n['Recipient Type']) === 'Coordinator';
  }).map(cpLink_noteOut_).reverse().slice(0, 30);
  return {user: {name: name, email: email}, clinicalTerm: term,
    recent: tracking.concat(pcs.slice(0, 40)).sort(function(a, b) {return b.dateValue - a.dateValue;}).slice(0, 60),
    pc: {pending: pcs.filter(function(p) {return p.coordinatorReviewStatus !== 'Reviewed' && !p.flagged;}).length,
      returned: pcs.filter(function(p) {return p.flagged;}).length},
    inbox: replies, inboxUnread: replies.filter(function(r) {return r.status === 'Unread';}).length};
}
function cpLink_noteOut_(n) {
  return {id: cpLink_t_(n['Notification ID']), createdAt: cpLink_iso_(n['Created At']),
    recipientType: cpLink_t_(n['Recipient Type']), recipientId: cpLink_t_(n['Recipient ID']),
    recipient: cpLink_t_(n['Recipient']), type: cpLink_t_(n['Type']), itemKind: cpLink_t_(n['Item Kind']),
    itemId: cpLink_t_(n['Item ID']), studentId: cpLink_t_(n['Student ID']), student: cpLink_t_(n['Student']),
    title: cpLink_t_(n['Title']), message: cpLink_t_(n['Message']), from: cpLink_t_(n['From']),
    status: cpLink_t_(n['Status']), readAt: cpLink_iso_(n['Read At']), response: cpLink_t_(n['Response']),
    respondedAt: cpLink_iso_(n['Responded At'])};
}

/* ------------------------------------------------------------------ */
/* Review detail + coordinator decisions                              */
/* ------------------------------------------------------------------ */
function dashboardV2ReviewDetail(kind, itemId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = cpLink_rows_(ss);
  let item;
  if (kind === 'professional') {
    item = cpLink_pcItems_(ss, log).filter(function(p) {return p.itemId === itemId;})[0];
    if (!item) throw new Error('Professional Competency record not found: ' + itemId);
  } else {
    const r = cpLink_trackingRecord_(itemId);
    const state = clinicalReview_getState(itemId);
    const g = gradeEngine_formatEvaluationGrade(r['Final Score']);
    item = {kind: 'evaluation', itemId: itemId, submissionId: itemId, studentId: cpLink_t_(r['Student ID']),
      student: cpLink_t_(r['Student']), clinicalTerm: cpLink_t_(r['Clinical Term']),
      examination: cpLink_t_(r['Examination']), evaluationType: cpLink_t_(r['Evaluation Type']),
      clinicalSite: cpLink_t_(r['Clinical Site']), evaluator: cpLink_t_(r['Evaluator']),
      evaluatorId: cpLink_t_(r['Evaluator ID']), date: cpLink_iso_(r['Date']).slice(0, 10),
      gradeDisplay: g.combinedText, result: cpLink_t_(r['Result']),
      sourceDocument: cpLink_t_(r['Source Document']), submissionMethod: cpLink_t_(r['Submission Method']),
      coordinatorReviewStatus: cpLink_t_(r['Coordinator Review Status']) || 'Pending',
      flagged: clinicalReview_isFlaggedValue_(r['Coordinator Flag']),
      coordinatorNotes: cpLink_t_(r['Coordinator Notes']), reviewedBy: cpLink_t_(r['Reviewed By']),
      state: state};
  }
  item.thread = log.filter(function(n) {return cpLink_t_(n['Item ID']) === itemId;}).map(cpLink_noteOut_);
  return item;
}
function cpLink_requireText_(v, label, max) {
  const s = cpLink_t_(v);
  if (s.length > (max || 4000)) throw new Error(label + ' is too long.');
  return s;
}
/** Return to the CI: flag + notes on the record, message to the instructor. */
function dashboardV2SendBack(request) {
  const kind = request && request.kind === 'professional' ? 'professional' : 'evaluation';
  const itemId = cpLink_t_(request && request.itemId);
  const notes = cpLink_requireText_(request && request.notes, 'Notes');
  if (!itemId) throw new Error('Item ID required.');
  if (!notes) throw new Error('Write a note so the instructor knows what to review.');
  const item = dashboardV2ReviewDetail(kind, itemId);
  if (!item.evaluatorId) throw new Error('This record has no Evaluator ID, so it cannot be routed to an instructor.');
  if (kind === 'evaluation') {
    clinicalReview_saveNotes(itemId, notes);
    clinicalReview_setFlag(itemId, true);
  }
  const id = cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': item.evaluatorId, 'Recipient': item.evaluator,
    'Type': 'Returned', 'Item Kind': kind, 'Item ID': itemId, 'Student ID': item.studentId, 'Student': item.student,
    'Title': (kind === 'professional' ? 'Professional Competency' : item.examination + ' ' + (item.evaluationType || '')) +
      ' returned for review', 'Message': notes, 'From': cpLink_user_()});
  return {success: true, notificationId: id};
}
/** Accept: mark reviewed (and clear a prior flag), tell the instructor. */
function dashboardV2Accept(request) {
  const kind = request && request.kind === 'professional' ? 'professional' : 'evaluation';
  const itemId = cpLink_t_(request && request.itemId);
  const note = cpLink_requireText_(request && request.notes, 'Note');
  if (!itemId) throw new Error('Item ID required.');
  const item = dashboardV2ReviewDetail(kind, itemId);
  if (kind === 'evaluation') {
    if (note) clinicalReview_saveNotes(itemId, note);
    if (item.flagged) clinicalReview_setFlag(itemId, false);
    clinicalReview_markReviewed(itemId);
  }
  let id = '';
  if (item.evaluatorId) id = cpLink_add_({'Recipient Type': 'CI', 'Recipient ID': item.evaluatorId,
    'Recipient': item.evaluator, 'Type': 'Accepted', 'Item Kind': kind, 'Item ID': itemId,
    'Student ID': item.studentId, 'Student': item.student,
    'Title': (kind === 'professional' ? 'Professional Competency' : item.examination + ' ' + (item.evaluationType || '')) + ' accepted',
    'Message': note, 'From': cpLink_user_()});
  return {success: true, notificationId: id};
}
function dashboardV2SaveNotes(request) {
  const itemId = cpLink_t_(request && request.itemId);
  if (!itemId || (request && request.kind) === 'professional') throw new Error('Notes are saved on evaluations only.');
  return clinicalReview_saveNotes(itemId, cpLink_requireText_(request.notes, 'Notes'));
}
function dashboardV2MarkInboxRead(notificationId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = cpLink_sheet_(ss);
  const row = cpLink_rows_(ss).filter(function(n) {
    return cpLink_t_(n['Notification ID']) === cpLink_t_(notificationId) && cpLink_t_(n['Recipient Type']) === 'Coordinator';
  })[0];
  if (!row) throw new Error('Message not found.');
  sh.getRange(row._row, CP_LINK.HEADERS.indexOf('Status') + 1, 1, 2).setValues([['Read', new Date()]]);
  return {success: true};
}
/** Links for System & Setup (reference tabs in this workbook). */
function dashboardV2SystemLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = ss.getUrl();
  return {workbook: url, sheets: ss.getSheets().map(function(s) {
    return {name: s.getName(), url: url + '#gid=' + s.getSheetId()};
  })};
}
