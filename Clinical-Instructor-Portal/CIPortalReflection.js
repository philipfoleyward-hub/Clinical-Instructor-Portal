/**
 * Clinical Instructor Portal — Survey Responses (Clinical Instructor Summary
 * Reflection Form).
 *
 * The coordinator shares each instructor's student-feedback summary per
 * semester (CCMS Coordinator Dashboard → Clinical Instructors). Here the
 * instructor reads it, completes the three sections of the approved template,
 * saves drafts, and submits. On submit the approved Google Doc template is
 * copied and filled (name, semester, the three answers), a PDF is saved beside
 * it in a private Drive folder, and the coordinator is notified to accept or
 * return it. Shared sheet: "CI Reflections" in the CCMS workbook.
 */
const CI_REFL = {
  SHEET: 'CI Reflections',
  HEADERS: ['Reflection ID', 'Evaluator ID', 'Evaluator', 'Term', 'Summary Text', 'Summary File ID', 'Shared At', 'Shared By',
    'Observations', 'Learning Experiences', 'Goals', 'Status', 'Saved At', 'Submitted At', 'Doc ID', 'PDF ID',
    'Coordinator Note', 'Reviewed At', 'Reviewed By', 'Version'],
  TEMPLATE_ID: '1WAJO5KTslzQcO812hL-G_cQa0syqb3-Z0zRg19FQSzg',
  FOLDER: 'CCMS — CI Reflections (restricted)',
  // Where each answer goes: inserted after the template paragraph containing this text.
  ANCHORS: {goals: 'Strengthening student support', learning: 'Did you try anything new', observations: 'What stood out to you'}
};
const ciReflT_ = v => v == null ? '' : String(v).trim();
const ciReflIso_ = v => v instanceof Date ? v.toISOString() : ciReflT_(v);
function ciReflTemplateId_() {
  return ciReflT_(PropertiesService.getScriptProperties().getProperty('CI_REFLECTION_TEMPLATE_ID')) || CI_REFL.TEMPLATE_ID;
}
function ciReflMine_(ev) {
  return ciMileRows_(CI_REFL.SHEET, CI_REFL.HEADERS).filter(r => ciReflT_(r['Evaluator ID']) === ev.evaluatorId);
}
function ciReflOwned_(ev, id) {
  const r = ciReflMine_(ev).find(x => ciReflT_(x['Reflection ID']) === ciReflT_(id));
  if (!r) throw new Error('Reflection not found.');
  return r;
}
function ciReflOut_(r) {
  return {reflectionId: ciReflT_(r['Reflection ID']), term: ciReflT_(r.Term), status: ciReflT_(r.Status) || 'Requested',
    summaryText: ciReflT_(r['Summary Text']), hasSummaryPdf: !!ciReflT_(r['Summary File ID']), sharedAt: ciReflIso_(r['Shared At']),
    observations: ciReflT_(r.Observations), learning: ciReflT_(r['Learning Experiences']), goals: ciReflT_(r.Goals),
    savedAt: ciReflIso_(r['Saved At']), submittedAt: ciReflIso_(r['Submitted At']), hasPdf: !!ciReflT_(r['PDF ID']),
    coordinatorNote: ciReflT_(r['Coordinator Note']), reviewedAt: ciReflIso_(r['Reviewed At']), version: Number(r.Version) || 0};
}

function ciR1ReflectionList(token) {
  const ev = ciR1Require_(token);
  const list = ciReflMine_(ev).map(ciReflOut_).sort((a, b) => (b.sharedAt || '').localeCompare(a.sharedAt || ''));
  return {success: true, evaluator: ev.evaluatorName, reflections: list};
}

/** PDF download for the instructor: the shared summary, or their submitted reflection. */
function ciR1ReflectionPdf(token, reflectionId, which) {
  const ev = ciR1Require_(token);
  const r = ciReflOwned_(ev, reflectionId);
  const id = ciReflT_(which === 'reflection' ? r['PDF ID'] : r['Summary File ID']);
  if (!id) throw new Error('No PDF is available.');
  const blob = DriveApp.getFileById(id).getBlob();
  return {success: true, name: (which === 'reflection' ? 'Summary Reflection ' : 'Student feedback summary ') + ciReflT_(r.Term) + '.pdf',
    base64: Utilities.base64Encode(blob.getBytes())};
}

/** Save a draft, or submit (fills the approved template, saves Doc + PDF, notifies the coordinator). */
function ciR1ReflectionSave(token, request) {
  const ev = ciR1Require_(token);
  const q = request || {};
  const clean = v => String(v == null ? '' : v).replace(/\r\n?/g, '\n').replace(/[ \t]+\n/g, '\n').trim().slice(0, 8000);
  const ans = {observations: clean(q.observations), learning: clean(q.learning), goals: clean(q.goals)};
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const r = ciReflOwned_(ev, q.reflectionId);
    const st = ciReflT_(r.Status) || 'Requested';
    if (st === 'Submitted') throw new Error('This reflection is waiting for the coordinator. It can be changed if it is returned to you.');
    if (st === 'Accepted') throw new Error('This reflection was already accepted.');
    r.Observations = ans.observations; r['Learning Experiences'] = ans.learning; r.Goals = ans.goals; r['Saved At'] = new Date();
    if (!q.submit) {
      ciMileWrite_(CI_REFL.SHEET, CI_REFL.HEADERS, r._row, r);
      return {success: true, saved: true};
    }
    const missing = [['observations', 'Observations and Insights'], ['learning', 'Learning Experiences'], ['goals', 'Goals for the future']]
      .filter(x => ans[x[0]].length < 20).map(x => x[1]);
    if (missing.length) throw new Error('Please complete: ' + missing.join(', ') + ' (a few sentences each).');
    const version = (Number(r.Version) || 0) + 1;
    const term = ciReflT_(r.Term);
    const folder = ciReflFolder_(term);
    const name = ev.evaluatorName + ' — ' + term + ' CI Summary Reflection' + (version > 1 ? ' (v' + version + ')' : '');
    const docFile = DriveApp.getFileById(ciReflTemplateId_()).makeCopy(name, folder);
    ciReflFill_(docFile.getId(), ev.evaluatorName, term, ans);
    const pdf = folder.createFile(docFile.getBlob().getAs('application/pdf').setName(name + '.pdf'));
    r.Status = 'Submitted'; r['Submitted At'] = new Date(); r['Doc ID'] = docFile.getId(); r['PDF ID'] = pdf.getId(); r.Version = version;
    ciMileWrite_(CI_REFL.SHEET, CI_REFL.HEADERS, r._row, r);
    try {
      const n = {'Notification ID': 'NT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 14).toUpperCase(), 'Created At': new Date(),
        'Recipient Type': 'Coordinator', 'Recipient': 'Clinical Coordinator', 'Type': 'Reflection Submitted', 'Item Kind': 'reflection',
        'Item ID': ciReflT_(r['Reflection ID']), 'Title': (version > 1 ? 'Reflection resubmitted — ' : 'Reflection submitted — ') + ev.evaluatorName + ' · ' + term,
        'Message': ans.observations.slice(0, 300) + (ans.observations.length > 300 ? '…' : ''), 'From': ev.evaluatorName + ' (' + ev.evaluatorId + ')', 'Status': 'Unread'};
      ciNtSheet_().appendRow(CI_NT_HEADERS.map(h => n[h] == null ? '' : n[h]));
    } catch (e) {console.warn('Coordinator reflection notification failed: ' + (e.message || e));}
    return {success: true, submitted: true, version: version};
  } finally {lock.releaseLock();}
}

function ciReflFolder_(term) {
  const root = DriveApp.getRootFolder();
  const it = root.getFoldersByName(CI_REFL.FOLDER);
  const base = it.hasNext() ? it.next() : root.createFolder(CI_REFL.FOLDER);
  const s = base.getFoldersByName(term);
  return s.hasNext() ? s.next() : base.createFolder(term);
}
/** Fill the copied template: name and semester cells, then each answer after its section prompt. */
function ciReflFill_(docId, name, term, ans) {
  const doc = DocumentApp.openById(docId), body = doc.getBody();
  const tables = body.getTables();
  if (tables.length && tables[0].getNumRows() >= 2) {
    tables[0].getCell(0, 1).setText(name);
    tables[0].getCell(1, 1).setText(term);
  }
  const plain = {};
  plain[DocumentApp.Attribute.ITALIC] = false; plain[DocumentApp.Attribute.BOLD] = false;
  plain[DocumentApp.Attribute.UNDERLINE] = false; plain[DocumentApp.Attribute.FONT_SIZE] = 12;
  ['goals', 'learning', 'observations'].forEach(key => {   // bottom-up so earlier indexes stay valid
    let idx = -1;
    for (let i = 0; i < body.getNumChildren(); i++) {
      const ch = body.getChild(i);
      if (ch.getType() === DocumentApp.ElementType.PARAGRAPH && ch.asParagraph().getText().indexOf(CI_REFL.ANCHORS[key]) >= 0) {idx = i; break;}
    }
    const lines = ans[key].split('\n');
    if (idx < 0) {lines.forEach(l => body.appendParagraph(l).setAttributes(plain)); return;}
    lines.forEach((l, k) => body.insertParagraph(idx + 1 + k, l).setAttributes(plain));
  });
  doc.saveAndClose();
}
