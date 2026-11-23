/**
 * Clinical Instructor Portal — coordinator notifications and My Submissions.
 * Reads the shared "Portal Notifications" log in the CCMS workbook (written by
 * the Coordinator Dashboard: Accepted / Returned with notes) and lets the
 * signed-in instructor mark items read and reply. Replies are appended for the
 * coordinator. Instructors only ever see rows addressed to their stable
 * Evaluator ID. Clinical Tracking / Grade Records are read, never written.
 */
const CI_NT_SHEET = 'Portal Notifications';
const CI_NT_HEADERS = ['Notification ID', 'Created At', 'Recipient Type', 'Recipient ID', 'Recipient',
  'Type', 'Item Kind', 'Item ID', 'Student ID', 'Student', 'Title', 'Message', 'From',
  'Status', 'Read At', 'Response', 'Responded At'];

function ciNtSheet_() {
  const book = ciDataBook_();
  let sh = book.getSheetByName(CI_NT_SHEET);
  if (!sh) {
    sh = book.insertSheet(CI_NT_SHEET);
    sh.getRange(1, 1, 1, CI_NT_HEADERS.length).setValues([CI_NT_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function ciNtRows_() {
  const sh = ciNtSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(1, 1, last, CI_NT_HEADERS.length).getValues();
  const head = values[0].map(String);
  return values.slice(1).map((r, i) => {
    const o = {_row: i + 2};
    head.forEach((h, j) => {o[h] = r[j];});
    return o;
  });
}
const ciNtT_ = v => v == null ? '' : String(v).trim();
const ciNtIso_ = v => v instanceof Date ? v.toISOString() : ciNtT_(v);
function ciNtOut_(n) {
  return {id: ciNtT_(n['Notification ID']), createdAt: ciNtIso_(n['Created At']), type: ciNtT_(n['Type']),
    itemKind: ciNtT_(n['Item Kind']), itemId: ciNtT_(n['Item ID']), studentId: ciNtT_(n['Student ID']),
    student: ciNtT_(n['Student']), title: ciNtT_(n['Title']), message: ciNtT_(n['Message']),
    from: ciNtT_(n['From']), status: ciNtT_(n['Status']) || 'Unread', readAt: ciNtIso_(n['Read At']),
    response: ciNtT_(n['Response']), respondedAt: ciNtIso_(n['Responded At'])};
}
function ciNtMine_(ev) {
  return ciNtRows_().filter(n => ciNtT_(n['Recipient Type']) === 'CI' && ciNtT_(n['Recipient ID']) === ev.evaluatorId);
}
function ciR1Notifications(token) {
  const ev = ciR1Require_(token);
  const mine = ciNtMine_(ev).map(ciNtOut_).reverse();
  return {success: true, unread: mine.filter(n => n.status === 'Unread').length, notifications: mine.slice(0, 60)};
}
function ciNtOwned_(ev, id) {
  const row = ciNtMine_(ev).find(n => ciNtT_(n['Notification ID']) === ciNtT_(id));
  if (!row) throw new Error('Notification not found');
  return row;
}
function ciR1NotificationRead(token, id) {
  const ev = ciR1Require_(token);
  const row = ciNtOwned_(ev, id);
  if (ciNtT_(row.Status) !== 'Unread') return {success: true};
  ciNtSheet_().getRange(row._row, CI_NT_HEADERS.indexOf('Status') + 1, 1, 2).setValues([['Read', new Date()]]);
  return {success: true};
}
function ciR1NotificationReply(token, id, text) {
  const ev = ciR1Require_(token);
  const reply = String(text || '').trim();
  if (!reply) throw new Error('Write a reply first');
  if (reply.length > 4000) throw new Error('Reply is too long');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const row = ciNtOwned_(ev, id), sh = ciNtSheet_();
    const at = new Date();
    sh.getRange(row._row, CI_NT_HEADERS.indexOf('Status') + 1, 1, 4).setValues([['Read', row['Read At'] || at, reply, at]]);
    const out = {'Notification ID': 'NT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 14).toUpperCase(),
      'Created At': at, 'Recipient Type': 'Coordinator', 'Recipient ID': '', 'Recipient': 'Clinical Coordinator',
      'Type': 'CI Reply', 'Item Kind': row['Item Kind'], 'Item ID': row['Item ID'], 'Student ID': row['Student ID'],
      'Student': row['Student'], 'Title': 'Reply from ' + ev.evaluatorName + ' — ' + ciNtT_(row['Title']),
      'Message': reply, 'From': ev.evaluatorName + ' (' + ev.evaluatorId + ')', 'Status': 'Unread'};
    sh.appendRow(CI_NT_HEADERS.map(h => out[h] == null ? '' : out[h]));
    return {success: true};
  } finally {lock.releaseLock();}
}
/** Everything this instructor has submitted, with the coordinator's review state. */
function ciR1MySubmissions(token) {
  const ev = ciR1Require_(token);
  const book = ciDataBook_();
  const decisions = {};
  ciNtMine_(ev).forEach(n => {
    const t = ciNtT_(n['Type']);
    if (t === 'Accepted' || t === 'Returned') decisions[ciNtT_(n['Item ID'])] = ciNtOut_(n);
  });
  const fmt = v => {const d = v instanceof Date ? v : new Date(v); return isNaN(d) ? ciNtT_(v) : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');};
  const tracking = ciDataRows_(book, 'Clinical Tracking', ['Submission ID', 'Evaluator ID'])
    .filter(r => ciNtT_(r['Evaluator ID']) === ev.evaluatorId && ciDataNorm_(r['Submission Method']) !== 'historical import')
    .map(r => ({kind: 'evaluation', itemId: ciNtT_(r['Submission ID']), studentId: ciNtT_(r['Student ID']),
      student: ciNtT_(r.Student), title: ciNtT_(r.Examination), evaluationType: ciNtT_(r['Evaluation Type']),
      site: ciNtT_(r['Clinical Site']), date: fmt(r.Date), result: ciNtT_(r.Result),
      grade: r['Final Score'] === '' ? '' : String(r['Final Score']),
      status: ciDataTrue_(r['Coordinator Flag']) || /^(yes|flagged)$/i.test(ciNtT_(r['Coordinator Flag'])) ? 'Returned'
        : ciDataNorm_(r['Coordinator Review Status']) === 'reviewed' ? 'Accepted' : 'Submitted',
      notes: ciNtT_(r['Coordinator Notes']), document: ciNtT_(r['Source Document']),
      method: ciNtT_(r['Submission Method'])}));
  let pcs = [];
  try {
    pcs = ciDataRows_(book, 'Grade Records', ['Grade Record ID', 'Grade Category'])
      .filter(r => ciDataNorm_(r['Grade Category']) === 'professional competency' && ciDataTrue_(r.Active) &&
        /^PC-/.test(ciNtT_(r['Source ID'])) && new RegExp('Evaluator ID:\\s*' + ev.evaluatorId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*(\\n|$)').test(ciNtT_(r.Notes)))
      .map(r => {
        const id = ciNtT_(r['Grade Record ID']), d = decisions[id];
        const res = /Result:\s*([^\n]+)/.exec(ciNtT_(r.Notes)), site = /Clinical site:\s*([^\n]+)/.exec(ciNtT_(r.Notes));
        return {kind: 'professional', itemId: id, studentId: ciNtT_(r['Student ID']), student: ciNtT_(r.Student),
          title: 'Professional Competency', evaluationType: ciNtT_(r['Clinical Term']), site: site ? site[1].trim() : '',
          date: fmt(r.Date), result: res ? res[1].trim() : '',
          grade: r['Percent Grade'] === '' ? '' : Number(r['Percent Grade']).toFixed(2) + '%',
          status: d ? (d.type === 'Returned' ? 'Returned' : 'Accepted') : 'Submitted', notes: d ? d.message : '',
          document: ciNtT_(r['Source Document']), method: ciNtT_(r.Source)};
      });
  } catch (e) {console.warn('PC submissions unavailable: ' + (e.message || e));}
  const all = tracking.concat(pcs).sort((a, b) => String(b.date).localeCompare(String(a.date)));
  return {success: true, submissions: all.slice(0, 200), count: all.length};
}
