/**
 * CCMS Clinical Instructor Portal → canonical Google Form item response map.
 * Executed ONLY in the original CCMS Apps Script project, against its bound
 * authoritative workbook and real Digital Clinical Evaluation Form.
 *
 * This script does NOT grade, file, notify, or append records. It reuses
 * grading_gradeMappedEvaluation_ in 04_ClinicalEvaluationGradingEngine.
 * Hospital-CI self-registration remains explicitly blocked until there is
 * an approved authentication/audit identity contract.
 */
function ccmsPortal_clean_(value) { return String(value === null || value === undefined ? '' : value).trim(); }
function ccmsPortal_norm_(value) { return ccmsPortal_clean_(value).toLowerCase().replace(/\s+/g,' '); }
function ccmsPortal_true_(value) { return ['true','yes','y','1','active'].includes(ccmsPortal_norm_(value)); }
function ccmsPortal_reference_(ss, sheetName, idField, id, fields) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Missing authoritative reference sheet: ' + sheetName);
  const vals = sheet.getDataRange().getValues();
  const headers = (vals.shift() || []).map(ccmsPortal_clean_);
  for (const field of [idField].concat(fields || [])) {
    if (!headers.includes(field)) throw new Error('Missing '+field+' header in '+sheetName);
  }
  const index = headers.indexOf(idField);
  const matches = vals.filter(r=>ccmsPortal_norm_(r[index])===ccmsPortal_norm_(id));
  if (matches.length!==1) throw new Error('Reference ID unavailable or ambiguous: '+sheetName);
  return Object.fromEntries(headers.map((h,j)=>[h,matches[0][j]]));
}
function ccmsPortal_date_(str) {
  const raw=ccmsPortal_clean_(str);
  const m=raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Evaluation date must be YYYY-MM-DD');
  const date=new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),12,0,0);
  if (date.getFullYear()!==Number(m[1])||date.getMonth()!==Number(m[2])-1||date.getDate()!==Number(m[3])) throw new Error('Invalid evaluation date');
  return date;
}
function ccmsPortal_resolveIdentity_(data) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Original CCMS workbook context unavailable');
  if (ccmsPortal_clean_(data.evaluationType)!=='Competency') throw new Error('Only the approved Competency instrument is supported by this adapter');
  const student=ccmsPortal_reference_(ss,APP_CONFIG.SHEETS.STUDENTS,'Student ID',data.studentId,['Student','Active']);
  if (!ccmsPortal_true_(student.Active)) throw new Error('Student is not active');
  const site=ccmsPortal_reference_(ss,APP_CONFIG.SHEETS.CLINICAL_SITES,'Clinical Site ID',data.siteId,['Clinical Site','Active']);
  if (!ccmsPortal_true_(site.Active)) throw new Error('Clinical site is not active');
  const evaluator=ccmsPortal_reference_(ss,APP_CONFIG.SHEETS.EVALUATORS,'Evaluator ID',data.evaluatorId,['Evaluator','Evaluator Type','Active','Portal Access']);
  if (!ccmsPortal_true_(evaluator.Active)||!ccmsPortal_true_(evaluator['Portal Access'])) throw new Error('Evaluator does not have active Portal Access');
  if (ccmsPortal_norm_(evaluator['Evaluator Type'])!=='college ci') throw new Error('Hospital CI portal evaluation identity is not yet authorized for this integration');
  const examination=ccmsPortal_reference_(ss,APP_CONFIG.SHEETS.EXAMINATIONS,'Examination ID',data.examinationId,['Examination','Active','College CI Allowed']);
  if (!ccmsPortal_true_(examination.Active)||!ccmsPortal_true_(examination['College CI Allowed'])) throw new Error('Examination not permitted for College CI');
  return {
    studentId:ccmsPortal_clean_(student['Student ID']), student:ccmsPortal_clean_(student.Student),
    siteId:ccmsPortal_clean_(site['Clinical Site ID']),site:ccmsPortal_clean_(site['Clinical Site']),
    evaluatorId:ccmsPortal_clean_(evaluator['Evaluator ID']),evaluator:ccmsPortal_clean_(evaluator.Evaluator),
    evaluatorType:ccmsPortal_clean_(evaluator['Evaluator Type']), examinationId:ccmsPortal_clean_(examination['Examination ID']),
    examination:ccmsPortal_clean_(examination.Examination), date:ccmsPortal_date_(data.date)
  };
}
function ccmsPortal_schema_() {
  const form=FormApp.openById(APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION);
  return form.getItems().map(item=>{
    const type=String(item.getType());
    const row={itemId:String(item.getId()),title:String(item.getTitle()),type:type};
    if(type===String(FormApp.ItemType.GRID)) row.rows=item.asGridItem().getRows();
    if(typeof item.asMultipleChoiceItem==='function' && type===String(FormApp.ItemType.MULTIPLE_CHOICE)) row.choices=item.asMultipleChoiceItem().getChoices().map(c=>String(c.getValue()));
    if(typeof item.asCheckboxItem==='function' && type===String(FormApp.ItemType.CHECKBOX)) row.choices=item.asCheckboxItem().getChoices().map(c=>String(c.getValue()));
    return row;
  });
}
/**
 * Server ignores identity answers supplied by browser and overwrites them
 * from authoritative Student/Site/Evaluator/Examination reference rows.
 * Non-identity answers must use exact Form item IDs from ccmsPortal_schema_.
 */
function ccmsPortal_responseMap_(form, data, identity) {
  const items=form.getItems(),byId=new Map(items.map(i=>[String(i.getId()),i]));
  const received=data.answers||{};
  if (!received||Array.isArray(received)||typeof received!=='object') throw new Error('Evaluation answers must be an item-ID map');
  if (Object.keys(received).length>500) throw new Error('Unexpected evaluation payload size');
  const map=Object.create(null);
  const metadata={
    'student':identity.student,'clinical site':identity.site,'evaluator':identity.evaluator,
    'date':identity.date,'evaluation date':identity.date,'examination':identity.examination,
    'evaluation type':'Competency'
  };
  const protectedTitles=new Set(['student','clinical site','evaluator','date','evaluation date','examination','evaluation type']);
  const types=FormApp.ItemType;
  Object.keys(received).forEach(id=>{
    const item=byId.get(String(id));
    if (!item) throw new Error('Unrecognized evaluation item ID: '+id);
    const title=ccmsPortal_norm_(item.getTitle()),raw=received[id];
    if (protectedTitles.has(title)) throw new Error('Evaluation identity fields are server-controlled');
    if (item.getType()===types.PAGE_BREAK||item.getType()===types.SECTION_HEADER) throw new Error('Unsupported answer for section header');
    if (item.getType()===types.GRID) {
      const rows=item.asGridItem().getRows();
      if (!Array.isArray(raw)||raw.length!==rows.length) throw new Error('Criterion count mismatch: '+item.getTitle());
      map[id]=raw.map(v=>{
        const answer=ccmsPortal_clean_(v);
        if (!['0','1','2','3','N/A'].includes(answer)) throw new Error('Invalid criterion rating for '+item.getTitle());
        return answer;
      });
    } else if (Array.isArray(raw)) {
      if (raw.length>25||raw.some(v=>typeof v!=='string'||v.length>350)) throw new Error('Invalid multi-select response');
      const choices=item.getType()===types.CHECKBOX&&typeof item.asCheckboxItem==='function'
        ? item.asCheckboxItem().getChoices().map(c=>String(c.getValue())) : null;
      if (choices&&raw.some(v=>!choices.includes(v))) throw new Error('Response not in official form choices: '+item.getTitle());
      map[id]=raw.slice();
    } else {
      if (typeof raw!=='string'&&typeof raw!=='number') throw new Error('Invalid item response type');
      const val=ccmsPortal_clean_(raw);
      if (val.length>3000) throw new Error('Response too long for '+item.getTitle());
      const t=item.getType();
      if ((t===types.MULTIPLE_CHOICE||t===types.LIST)&&typeof item.asMultipleChoiceItem==='function'&&t===types.MULTIPLE_CHOICE) {
        const choices=item.asMultipleChoiceItem().getChoices().map(c=>String(c.getValue()));
        if (!choices.includes(val)) throw new Error('Response not in official form choices: '+item.getTitle());
      }
      map[id]=val;
    }
  });
  let identityCount={student:0,site:0,evaluator:0,exam:0,date:0};
  items.forEach(item=>{
    const title=ccmsPortal_norm_(item.getTitle());
    if (Object.prototype.hasOwnProperty.call(metadata,title)) {
      map[String(item.getId())]=metadata[title];
      if(title==='student')identityCount.student++;
      if(title==='clinical site')identityCount.site++;
      if(title==='evaluator')identityCount.evaluator++;
      if(title==='examination')identityCount.exam++;
      if(title==='date'||title==='evaluation date')identityCount.date++;
    }
  });
  if(Object.values(identityCount).some(n=>n===0)) throw new Error('Official Form identity schema has changed; portal submission paused');
  const failure=grading_readAutomaticFailure_(form,map);
  if (failure.occurred&&(!failure.reasons.length||!failure.comments)) throw new Error('Automatic Failure requires an official reason and evaluator explanation');
  if (!failure.occurred) {
    const perfTitles=new Set(['exam preparation','patient care','radiation protection','equipment manipulation','positioning skills','portable / contrast']);
    items.forEach(item=>{
      if(item.getType()===types.GRID&&perfTitles.has(ccmsPortal_norm_(item.getTitle()))&&!Object.prototype.hasOwnProperty.call(map,String(item.getId()))) {
        throw new Error('Complete all performance criterion sections or select Auto-Fail');
      }
    });
  }
  return map;
}
function gradeClinicalEvaluationFromPortalPayload(data) {
  const identity=ccmsPortal_resolveIdentity_(data);
  const form=FormApp.openById(APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION);
  const map=ccmsPortal_responseMap_(form,data,identity);
  const grade=grading_gradeMappedEvaluation_(form,map,{
    responseId:ccmsPortal_clean_(data.submissionId),submittedAt:new Date()
  });
  return {grade:grade,identity:identity};
}
