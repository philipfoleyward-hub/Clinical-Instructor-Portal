/**
 * CCMS TEST-001 Grade QA — READ ONLY.
 * Install as a new server-side file in the ORIGINAL CCMS project, not the portal.
 * Exercises actual CCMS reference tables, Google Form and existing native grade engine.
 * No portal grading adapter is required; this tests original grading only, NOT portal submission parity.
 * No Form submit, spreadsheet append, Drive write, email, deployment or side effects.
 *
 * REQUIRED nonsecret Script Properties on ORIGINAL CCMS:
 *  CI_PORTAL_QA_EVALUATOR_ID: an active authorized College CI evaluator ID
 *  CI_PORTAL_QA_SITE_ID: an active site ID, e.g. TEST-001's test site
 *  CI_PORTAL_QA_EXAMINATION_ID: an active College-CI-allowed exam ID
 * OPTIONAL CI_PORTAL_QA_DATE: YYYY-MM-DD; otherwise today in script time zone
 *
 * This is scoring validation only, not proof of real CI rotation authorization.
 */
function ccmsPortalQaTest001GradePreviewReadOnly() {
  const props=PropertiesService.getScriptProperties();
  function required(key){const v=String(props.getProperty(key)||'').trim();if(!v)throw Error('READ-ONLY QA BLOCKED: configure '+key+' in ORIGINAL CCMS Script Properties');return v;}
  const evaluatorId=required('CI_PORTAL_QA_EVALUATOR_ID');
  const siteId=required('CI_PORTAL_QA_SITE_ID');
  const examinationId=required('CI_PORTAL_QA_EXAMINATION_ID');
  const date=String(props.getProperty('CI_PORTAL_QA_DATE')||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd')).trim();
  // The Form, student, evaluator, site, examination and weights are retrieved
  // directly from the real authoritative sources. The answers are synthetic.
  const form=FormApp.openById(APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION);
  const items=form.getItems();
  const gridNames=new Set(['exam preparation','patient care','radiation protection','equipment manipulation','positioning skills','portable / contrast']);
  const types=FormApp.ItemType;
  const gridItems=items.filter(item=>item.getType()===types.GRID&&gridNames.has(String(item.getTitle()).trim().toLowerCase()));
  if(gridItems.length!==6)throw Error('READ-ONLY QA BLOCKED: expected six authoritative performance grids; found '+gridItems.length);
  const firstView=items.findIndex(item=>item.getType()===types.PAGE_BREAK&&/^image evaluation\s*-\s*view 1$/i.test(String(item.getTitle()).trim()));
  if(firstView<0)throw Error('READ-ONLY QA BLOCKED: official Image Evaluation - View 1 not found');
  let end=items.length;
  for(let i=firstView+1;i<items.length;i++)if(items[i].getType()===types.PAGE_BREAK){end=i;break;}
  const imageParts=new Set(['positioning','equipment','radiation protection','exposure']);
  const imageItems=items.slice(firstView+1,end).filter(item=>imageParts.has(String(item.getTitle()).trim().toLowerCase()));
  const projection=items.slice(firstView+1,end).find(item=>String(item.getTitle()).trim().toLowerCase()==='view/projection');
  if(imageItems.length!==4||!projection)throw Error('READ-ONLY QA BLOCKED: first official image view should have four scoring parts and a projection field');
  function highestGradeChoice(item,grade){
    const type=item.getType();
    if(type===types.MULTIPLE_CHOICE){
      const options=item.asMultipleChoiceItem().getChoices().map(choice=>String(choice.getValue()));
      const exact=options.find(x=>x===String(grade));
      const partial=options.find(x=>new RegExp('^\\s*'+grade+'(?:\\b|\\s|\\s*[-=:;])').test(x));
      if(!exact&&!partial)throw Error('READ-ONLY QA BLOCKED: image choice for grade '+grade+' unavailable: '+item.getTitle());
      return exact||partial;
    }
    if(type===types.LIST){
      const options=item.asListItem().getChoices().map(choice=>String(choice.getValue()));
      const chosen=options.find(x=>x===String(grade)||new RegExp('^\\s*'+grade+'\\b').test(x));
      if(!chosen)throw Error('READ-ONLY QA BLOCKED: image choice for grade '+grade+' unavailable: '+item.getTitle());
      return chosen;
    }
    if(type===types.TEXT||type===types.PARAGRAPH_TEXT)return String(grade);
    throw Error('READ-ONLY QA BLOCKED: unexpected image item type for '+item.getTitle());
  }
  function buildSyntheticAnswers(rating){
    const answers={};
    for(const item of gridItems)answers[String(item.getId())]=item.asGridItem().getRows().map(()=>String(rating));
    answers[String(projection.getId())]='TEST-001 QA (synthetic; no radiograph)';
    for(const item of imageItems)answers[String(item.getId())]=highestGradeChoice(item,3);
    return answers;
  }
  const base={studentId:'TEST-001',evaluatorId,authenticatedEvaluatorId:evaluatorId,siteId,examinationId,date,evaluationType:'Competency'};
  function evaluate(label,rating){
    const payload=Object.assign({},base,{submissionId:'CI-PORTAL-'+Utilities.getUuid(),answers:buildSyntheticAnswers(rating)});
    // Read-only compatibility: use ORIGINAL CCMS's existing native FormResponse
    // grading entrypoint. This object is in memory only: NO createResponse(),
    // submit(), Form writes, document generation, or Tracking mutations.
    // Do not change the authoritative grading engine just to run QA.
    if (typeof gradeClinicalEvaluation !== 'function') {
      throw Error('READ-ONLY QA BLOCKED: original gradeClinicalEvaluation function is unavailable');
    }
    const mappedItems = items.filter(item => Object.prototype.hasOwnProperty.call(payload.answers, String(item.getId())));
    const syntheticResponse = {
      getItemResponses() {
        return mappedItems.map(item => ({
          getItem() { return item; },
          getResponse() { return payload.answers[String(item.getId())]; }
        }));
      },
      getId() { return payload.submissionId; },
      getTimestamp() { return new Date(); }
    };
    const result=gradeClinicalEvaluation(syntheticResponse);
    const summary={scenario:label,performanceScore:result.performance&&result.performance.score,imageScore:result.image&&result.image.score,finalScore:result.finalScore,result:result.result,autoFail:!!(result.automaticFailure&&result.automaticFailure.occurred),imageViewsUsed:result.image&&result.image.views&&result.image.views.length};
    if(!Number.isFinite(summary.performanceScore)||!Number.isFinite(summary.imageScore)||!Number.isFinite(summary.finalScore))throw Error('READ-ONLY QA BLOCKED: incomplete official scoring in '+label+'; '+JSON.stringify(summary));
    if(label==='all_3_should_PASS'&&summary.result!=='PASS')throw Error('READ-ONLY QA BLOCKED: high-score synthetic scenario not PASS: '+JSON.stringify(summary));
    if(label==='all_1_should_FAIL'&&summary.result!=='FAIL')throw Error('READ-ONLY QA BLOCKED: low-score synthetic scenario not FAIL: '+JSON.stringify(summary));
    return summary;
  }
  const summary={readOnly:true,studentId:'TEST-001',formGrids:gridItems.length,imageComponents:imageItems.length,scenarios:[evaluate('all_3_should_PASS',3),evaluate('all_1_should_FAIL',1)],writeOperations:0,notes:'This preview exercised the existing CCMS gradeClinicalEvaluation(FormResponse) engine using in-memory FormResponse-shaped answers and official Form/weights. It did NOT validate the portal adapter, authenticate a CI, or submit anything.'};
  console.log('TEST-001 READ-ONLY GRADE QA: '+JSON.stringify(summary));
  return summary;
}

/** Read only: help the coordinator find allowed IDs for the three QA Script Properties.
 * IDs only; this function does not change role/site/roster assignments.
 */
function ccmsPortalQaListAllowedIdsReadOnly(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!ss)throw Error('Original CCMS workbook is unavailable');
  function rows(sheetName,needed){
    const sh=ss.getSheetByName(sheetName);if(!sh)throw Error('Worksheet missing: '+sheetName);
    const vals=sh.getDataRange().getValues();const heads=(vals.shift()||[]).map(x=>String(x).trim());
    needed.forEach(field=>{if(!heads.includes(field))throw Error(sheetName+' missing '+field)});
    return vals.map(row=>Object.fromEntries(heads.map((h,i)=>[h,row[i]])));
  }
  const yes=v=>['true','yes','1','active','y'].includes(String(v||'').toLowerCase().trim());
  const evaluatorIds=rows(APP_CONFIG.SHEETS.EVALUATORS,['Evaluator ID','Evaluator Type','Active','Portal Access'])
    .filter(r=>yes(r.Active)&&yes(r['Portal Access'])&&String(r['Evaluator Type']).trim().toLowerCase()==='college ci')
    .map(r=>String(r['Evaluator ID']).trim()).filter(Boolean).slice(0,30);
  const siteIds=rows(APP_CONFIG.SHEETS.CLINICAL_SITES,['Clinical Site ID','Active'])
    .filter(r=>yes(r.Active)).map(r=>String(r['Clinical Site ID']).trim()).filter(Boolean).slice(0,30);
  const examinationIds=rows(APP_CONFIG.SHEETS.EXAMINATIONS,['Examination ID','Active','College CI Allowed'])
    .filter(r=>yes(r.Active)&&yes(r['College CI Allowed']))
    .map(r=>String(r['Examination ID']).trim()).filter(Boolean).slice(0,30);
  const out={readOnly:true,evaluatorIds,siteIds,examinationIds,note:'ID lists are candidates, not proof that this evaluator has an authorized rotation for this student. Set QA properties only for approved TEST-001 context.'};
  console.log('TEST-001 QA OPTIONS: '+JSON.stringify(out));return out;
}
