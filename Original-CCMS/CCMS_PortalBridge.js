/** ORIGINAL CCMS PROJECT ONLY. Read-only library API. */
function ciBridgeGetEvaluators(workbookId) {
  const ss = SpreadsheetApp.openById(ciBridgeWorkbookId_(workbookId));
  const sh = ss.getSheetByName('Evaluators');
  if (!sh) throw new Error('Evaluators sheet missing');
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(String);
  ['Evaluator ID','Evaluator','Evaluator Type','Active','Portal Access'].forEach(h => {
    if (!headers.includes(h)) throw new Error('Missing Evaluators header: ' + h);
  });
  return values.map(row => Object.fromEntries(headers.map((h,i)=>[h,row[i]])))
    .filter(r => ciBridgeTrue_(r.Active) && ciBridgeTrue_(r['Portal Access']))
    .map(r => ({evaluatorId:String(r['Evaluator ID']||'').trim(),evaluatorName:String(r.Evaluator||'').trim(),evaluatorType:String(r['Evaluator Type']||'').trim()}))
    .filter(r => r.evaluatorId && r.evaluatorName)
    .sort((a,b)=>a.evaluatorName.localeCompare(b.evaluatorName));
}
function ciBridgeGetHome(workbookId,evaluatorId,isoDate) {
  const evaluator=ciBridgeGetEvaluators(workbookId).find(e=>e.evaluatorId===String(evaluatorId||'').trim());
  if (!evaluator) throw new Error('Instructor access unavailable');
  const match=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(isoDate||''));
  if(!match) throw new Error('Date must be YYYY-MM-DD');
  const date=new Date(+match[1],+match[2]-1,+match[3]);
  if(date.getFullYear()!==+match[1]||date.getMonth()!==+match[2]-1||date.getDate()!==+match[3]) throw new Error('Invalid date');
  // Fail closed rather than silently reading the wrong workbook.
  // The existing scheduling service uses getActiveSpreadsheet internally.
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active || active.getId() !== ciBridgeWorkbookId_(workbookId)) {
    throw new Error('Scheduling workbook context unavailable or mismatched. CCMS scheduling service requires explicit workbook injection before standalone Home can be enabled.');
  }
  const workspace=scheduling_getCiDayWorkspace(evaluator.evaluatorId,date);
  return {success:true,evaluator,date:isoDate,assigned:workspace.assigned===true,assignments:workspace.assignments||[]};
}
function ciBridgeWorkbookId_(id) {
  const v=String(id||'').trim();
  if(!/^[A-Za-z0-9_-]{20,}$/.test(v)) throw new Error('Invalid workbook ID');
  return v;
}
function ciBridgeTrue_(v) {return v===true||['TRUE','YES','1'].includes(String(v||'').trim().toUpperCase());}

/** Read-only active-student selector from authoritative CCMS Students sheet.
 *  Uses the established scheduling deduplication rule; no fabricated records.
 *  This is a program-wide selector, NOT the instructor's assigned roster.
 */
function ciBridgeGetActiveStudents(workbookId) {
  const ss = SpreadsheetApp.openById(ciBridgeWorkbookId_(workbookId));
  const students = scheduling_getUniqueActiveStudents_(ss);
  return {
    success: true,
    students: students.map(function(s) {
      return {
        studentId: String(s.studentId || ''),
        student: String(s.student || ''),
        cohort: String(s.cohort || ''),
        clinicalTerm: '' // not supplied by this authoritative helper
      };
    }).filter(function(s) { return !!s.studentId; })
  };
}
