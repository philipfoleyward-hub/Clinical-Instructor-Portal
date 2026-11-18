/**
 * READ-ONLY TEST-001 adapter parity — ORIGINAL CCMS project only.
 * Requires your previous CI_PORTAL_QA_EVALUATOR_ID/SITE_ID/EXAMINATION_ID
 * Script Properties and the new 04_PortalEvaluationAdapter.js.
 *
 * Invokes the REAL production gradeClinicalEvaluation() twice per scenario:
 * (1) trusted native FormResponse-shaped data, and (2) validated portal map.
 * Compares grade fields. NO .submit(), spreadsheet append, Drive or email.
 */
function ccmsPortalQaNativeAdapterParityReadOnly() {
  const p=PropertiesService.getScriptProperties();
  function required(k){const v=String(p.getProperty(k)||'').trim();if(!v)throw Error('READ-ONLY QA BLOCKED: missing Script Property '+k);return v;}
  const evaluatorId=required('CI_PORTAL_QA_EVALUATOR_ID');
  const siteId=required('CI_PORTAL_QA_SITE_ID');
  const examinationId=required('CI_PORTAL_QA_EXAMINATION_ID');
  const date=String(p.getProperty('CI_PORTAL_QA_DATE')||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'));
  const form=FormApp.openById(APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION);
  const items=form.getItems(),types=FormApp.ItemType;
  const names=['exam preparation','patient care','radiation protection','equipment manipulation','positioning skills','portable / contrast'];
  const grids=items.filter(i=>i.getType()===types.GRID&&names.includes(String(i.getTitle()).trim().toLowerCase()));
  if(grids.length!==6)throw Error('READ-ONLY QA BLOCKED: six authoritative performance grids required');
  const start=items.findIndex(i=>i.getType()===types.PAGE_BREAK&&/^image evaluation\s*-\s*view 1$/i.test(i.getTitle()));
  if(start<0)throw Error('READ-ONLY QA BLOCKED: Image Evaluation - View 1 not found');
  let end=items.length;for(let i=start+1;i<items.length;i++)if(items[i].getType()===types.PAGE_BREAK){end=i;break;}
  const components=new Set(['positioning','equipment','radiation protection','exposure']);
  const view=items.slice(start+1,end);
  const projection=view.find(i=>String(i.getTitle()).trim().toLowerCase()==='view/projection');
  const imageParts=view.filter(i=>components.has(String(i.getTitle()).trim().toLowerCase()));
  if(!projection||imageParts.length!==4)throw Error('READ-ONLY QA BLOCKED: official View 1 scoring fields changed');
  function choose(item,grade) {
    const type=item.getType();
    if(type===types.TEXT||type===types.PARAGRAPH_TEXT)return String(grade);
    if(type!==types.MULTIPLE_CHOICE&&type!==types.LIST)throw Error('READ-ONLY QA BLOCKED: image score field type changed: '+item.getTitle());
    const choices=ccmsPortal_choices_(item)||[];
    const val=choices.find(x=>x===String(grade))||choices.find(x=>new RegExp('^\\s*'+grade+'(?:\\b|\\s|\\s*[-=:;])').test(x));
    if(!val)throw Error('READ-ONLY QA BLOCKED: official image score choice unavailable: '+item.getTitle());
    return val;
  }
  function answers(score){
    const map={};
    for(const grid of grids) {
      const choices=ccmsPortal_choices_(grid)||[];
      if(!choices.includes(String(score)))throw Error('READ-ONLY QA BLOCKED: rating '+score+' absent from '+grid.getTitle());
      map[String(grid.getId())]=grid.asGridItem().getRows().map(()=>String(score));
    }
    map[String(projection.getId())]='TEST-001 synthetic image (no radiograph)';
    for(const i of imageParts)map[String(i.getId())]=choose(i,3);
    return map;
  }
  function single(label,score,expected){
    const data={studentId:'TEST-001',siteId,evaluatorId,authenticatedEvaluatorId:evaluatorId,
      examinationId,date,evaluationType:'Competency',submissionId:'CI-PORTAL-'+Utilities.getUuid(),answers:answers(score)};
    const portal=gradeClinicalEvaluationFromPortalPayload(data);
    // Read-only native Form path with EXACT SAME authoritative identity and item responses.
    const responseMap=ccmsPortal_responseMap_(form,data,portal.identity);
    const native=gradeClinicalEvaluation(ccmsPortal_memoryFormResponse_(form,responseMap,data.submissionId));
    function scoreSummary(g){return {
      result:String(g.result),finalScore:Number(g.finalScore),
      performanceScore:Number(g.performance&&g.performance.score),
      imageScore:Number(g.image&&g.image.score),
      autoFail:Boolean(g.automaticFailure&&g.automaticFailure.occurred),
      student:String(g.evaluation&&g.evaluation.student||''),
      evaluator:String(g.evaluation&&g.evaluation.evaluator||''),
      examination:String(g.evaluation&&g.evaluation.examination||'')
    };}
    const a=scoreSummary(portal.grade),b=scoreSummary(native);
    if(JSON.stringify(a)!==JSON.stringify(b))throw Error('READ-ONLY QA BLOCKED: portal/native grade parity mismatch in '+label);
    if(a.result!==expected||!Number.isFinite(a.finalScore)||!Number.isFinite(a.imageScore)||!Number.isFinite(a.performanceScore))
      throw Error('READ-ONLY QA BLOCKED: unexpected authoritative grade '+label+' '+JSON.stringify(a));
    return {scenario:label,grade:a,parity:true};
  }
  const report={readOnly:true,studentId:'TEST-001',scenarios:[single('all_3_should_PASS',3,'PASS'),single('all_1_should_FAIL',1,'FAIL')],
    writeOperations:0,notes:'REAL original CCMS engine via validated portal item map AND in-memory native form parity. No signed gateway, Form submit, Tracking, document, Drive or email was invoked.'};
  console.log('TEST-001 PORTAL ADAPTER NATIVE PARITY: '+JSON.stringify(report));return report;
}
/** Inspect source dependencies without causing any write or network submission. */
function ccmsPortalQaSubmissionDependenciesReadOnly() {
  const checks={
    portalNativeAdapter:typeof gradeClinicalEvaluationFromPortalPayload==='function',
    officialGrading:typeof gradeClinicalEvaluation==='function',
    originalDigitalIntake:typeof digitalEvaluation_submitPortal_==='function',
    trackingWriter:typeof clinicalTracking_addRecord==='function',
    documentGenerator:typeof evaluationDocument_generate==='function',
    requirementRefresh:typeof rebuildStudentRequirementStatus==='function',
    notificationSender:typeof notification_sendEvaluationNotificationsSafely==='function',
    signedGateway:typeof ccmsPortalGateway_verify_==='function',
    gatewayDoPost:typeof doPost==='function'
  };
  const property=PropertiesService.getScriptProperties(),mode=String(property.getProperty('CI_PORTAL_CCMS_SUBMIT_MODE')||'disabled');
  const result={readOnly:true,checks,allFunctionsPresent:Object.values(checks).every(Boolean),submitMode:mode,
    note:'Presence != installation verification. Verify all domain service wiring, signed gateway permissions, Form mapping, 317 approved Doc tokens, authorized Drive root, duplicate recovery, and coordinator/evaluator emails before a TEST-001 commit. No action performed.'};
  console.log('CCMS PORTAL SUBMISSION DEPENDENCIES: '+JSON.stringify(result));return result;
}
