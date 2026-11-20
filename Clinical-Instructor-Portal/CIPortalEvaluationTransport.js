/**
 * Server-side-only transport for approved future CCMS portal evaluation calls.
 * Never place credentials, evaluator names, or student records into HTML.
 * Requires the CI_PORTAL_CCMS_GATEWAY_URL Script Property. Authentication is the
 * project owner's own OAuth token (this Portal web app executes as the owner);
 * the CCMS web app is deployed "Only myself", so Google admits only that token.
 */
/**
 * Single read-only configuration inspection reused by requests and release status.
 * This never reveals or logs the secret, and does not test authentication by
 * submitting anything. Institutional and standard Google web-app URLs are valid.
 */
function ciPortalEvaluationGatewayConfig_() {
  const props=PropertiesService.getScriptProperties();
  const url=String(props.getProperty('CI_PORTAL_CCMS_GATEWAY_URL')||'').trim();
  const issues=[];
  if(!url) issues.push('CI_PORTAL_CCMS_GATEWAY_URL is missing in the Clinical Instructor Portal Script Properties');
  else if(!/^https:\/\/script\.google\.com\/(?:a\/macros\/[a-z0-9.-]+\/|macros\/)s\/[A-Za-z0-9_-]+\/exec\/?$/.test(url))
    issues.push('CI_PORTAL_CCMS_GATEWAY_URL must be the original CCMS Web app /exec URL (standard or institutional format)');
  return {url:url,issues:issues};
}

function ciPortalEvaluationTransport_(action,verifiedPayload) {
  const gateway=ciPortalEvaluationGatewayConfig_();
  if(gateway.issues.length)
    throw new Error('Portal connection configuration: '+gateway.issues.join('; ')+'. No evaluation was submitted.');
  const timestamp=Date.now(),nonce=Utilities.getUuid(),payloadJson=JSON.stringify(verifiedPayload);
  let res;
  try {
    res=UrlFetchApp.fetch(gateway.url,{
      method:'post',contentType:'application/json',payload:JSON.stringify({timestamp,nonce,action,payloadJson}),
      headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},
      muteHttpExceptions:true,followRedirects:true
    });
  } catch(err) {
    throw new Error('Unable to reach the original CCMS web app: '+String(err.message||err));
  }
  const code=res.getResponseCode();
  if(code!==200)
    throw new Error('Original CCMS web app returned HTTP '+code+'. Verify gateway deployment version and Web app access permissions. No evaluation was submitted.');
  let value;
  try{value=JSON.parse(res.getContentText());}
  catch(e){throw new Error('Original CCMS gateway did not return JSON. Verify the /exec deployment, Web app access and doPost handler. No evaluation was submitted.');}
  if(!value||value.success!==true)
    throw new Error(String(value&&value.error||'Original CCMS gateway rejected the request'));
  return value;
}
function ciPortalEvaluationTrustedPayload_(token,request) {
  // Retain the original portal's session and site authorization checks.
  // ciHWAccess_ verifies this session once; do not call ciR1Require_ twice.
  const day=String(request.date||'');
  const access=ciHWAccess_(token,request.siteId,day);
  const evaluator=access.evaluator;
  if(String(evaluator.evaluatorType||'').trim().toLowerCase()!=='college ci')throw new Error('Hospital instructor evaluation access is not yet configured');
  if(!ciHWGetStudent_(access,request.studentId))throw new Error('Student access unavailable');
  const id=String(evaluator.evaluatorId||'');
  if(!id)throw new Error('Missing authenticated evaluator ID');
  const result={
    studentId:String(request.studentId||''),siteId:String(access.site.siteId),
    evaluatorId:id,authenticatedEvaluatorId:id,
    examinationId:String(request.examinationId||''),date:day,
    evaluationType:'Competency',submissionId:String(request.submissionId||''),
    answers:request.answers||{}
  };
  if(!/^CI-PORTAL-[a-f0-9-]{36}$/i.test(result.submissionId)&&result.submissionId)throw new Error('Invalid submission ID');
  return result;
}
/**
 * Open the approved evaluation using the EXISTING official Form schema.
 * Form display must not depend on the separate grading/submission gateway.
 * The exact Form ID is from the original CCMS 01_Config.gs, not a new form.
 * No academic grade, draft, or student record is written by this function.
 */
function ciR1EvaluationFormSchema(token,request) {
  ciPortalEvaluationTrustedPayload_(token,request); // Existing evaluator / site / student checks.
  const result=ciPortalLoadOfficialFormSchema_(false);
  return {success:true,formItems:result.formItems,schemaLoadedAt:result.loadedAt};
}
/**
 * PREVIEW ONLY: grade the authenticated Portal's official Form item answers
 * inside the Portal project, against the original CCMS workbook rules.
 * This does not depend on the cross-project submission gateway.
 * It cannot submit, create a FormResponse, alter Clinical Tracking, file a
 * document, or send email. Submission remains governed separately below.
 */
function ciR1EvaluationPreview(token,request) {
  const data=ciPortalEvaluationTrustedPayload_(token,request);
  // Preview uses the actual current evaluation controls rather than an
  // incomplete Form-ID export from the browser. Grading weights and penalties
  // still come solely from the authoritative CCMS workbook.
  if(request&&request.uiState)ciPortalCheckViewSelections_(request.uiState);
  const payload=request && request.uiState
    ? ciPortalGradeMapUIState_(data,request.uiState)
    : ciPortalGradeMapOfficialAnswers_(data);
  const grade=ciClinicalGradeEvaluation(payload);
  // CCMS percentage mapping, not the incorrect score/3 shortcut.
  // An Auto-Fail keeps the officially recorded numeric grade 0.0; section
  // scores still describe what was completed before the failure.
  grade.convertedPercent=grade.automaticFailure.occurred?null:
    ciPortalAcademic_scoreToPercent_(grade.finalScore);
  if(grade.performance) {
    grade.performance.convertedPercent=ciPortalAcademic_scoreToPercent_(grade.performance.score);
    for(const sec of (grade.performance.sections||[]))
      sec.convertedPercent=ciPortalAcademic_scoreToPercent_(sec.score);
  }
  if(grade.image)grade.image.convertedPercent=ciPortalAcademic_scoreToPercent_(grade.image.score);
  // User-facing sections 6 and 7 are a *single* official weighted group.
  // For the approved review layout, expose read-only subtotals from the exact
  // same already-weighted rows. These do not alter the authoritative grade.
  const combined=(grade.performance?.sections||[]).find(x=>String(x.name).trim().toLowerCase()==='portable / contrast');
  grade.uiSectionScores=[];
  if(combined){
    const slices=[['Portable / Mobile Radiography',(combined.rows||[]).slice(0,7)],
                  ['Contrast',(combined.rows||[]).slice(7,16)]];
    for(const [name,rows] of slices){
      const applicable=rows.filter(r=>r.applicable===true);
      const points=applicable.reduce((sum,r)=>sum+Number(r.points||0),0);
      const weight=applicable.reduce((sum,r)=>sum+Number(r.weight||0),0);
      const score=weight>0?points/weight:null;
      grade.uiSectionScores.push({name,score,
        convertedPercent:ciPortalAcademic_scoreToPercent_(score),displayOnly:true,
        points,weight});
    }
  }
  if(grade.finalScore===null||!Number.isFinite(Number(grade.finalScore)))
    throw new Error('The official grader could not produce a numeric final grade. Performance score: '+(grade.performance?.score===null?'not applicable':grade.performance?.score)+', image score: '+(grade.image?.score===null?'no scored image':grade.image?.score)+'. Confirm applicable criteria and at least one completed image view.');
  if(!['PASS','FAIL'].includes(grade.result))throw new Error('Invalid academic grading result');
  return {success:true,grade:grade,previewOnly:true,submitted:false};
}
/** Disabled by default. QA mode accepts only the explicitly approved test student.
 * Live mode needs two separate affirmative properties, never a client flag. */
function ciPortalEvaluationMode_(){
 const p=PropertiesService.getScriptProperties();
 const mode=String(p.getProperty('CI_PORTAL_EVALUATION_SUBMIT_MODE')||'disabled');
 if(mode==='test_only')return 'test_only';
 if(mode==='live'&&p.getProperty('CI_PORTAL_EVALUATION_LIVE_APPROVED')==='true')return 'live';
 return 'disabled';
}
function ciPortalEvaluationHasConnection_(){
 return ciPortalEvaluationGatewayConfig_().issues.length===0;
}
function ciR1EvaluationSubmit(token,request) {
 const data=ciPortalEvaluationTrustedPayload_(token,request);
 if(!data.submissionId)throw new Error('Stable submission ID required');
 const mode=ciPortalEvaluationMode_();
 if(mode==='disabled'||!ciPortalEvaluationHasConnection_())throw new Error('Official submission is disabled. Draft preserved; no evaluation was submitted.');
 if(mode==='test_only'&&data.studentId!=='TEST-001')throw new Error('Test-only release: student is not the authorized test identity');
 if(!data.answers||!Object.keys(data.answers).length)throw new Error('Submission has no official Form item responses');
 const review=request.uiState&&request.uiState.studentReview;
 if(!review||!['pending','reviewed'].includes(String(review.status)))throw new Error('Student Review & Feedback status is required');
 for(const prop of ['strengths','improvements','discussion'])if(typeof review[prop]!=='string'||review[prop].length>8000)throw new Error('Student review text is invalid');
 if(review.status==='reviewed'&&(!review.reviewedAt||!Number.isFinite(Date.parse(review.reviewedAt))))throw new Error('Reviewed with Student requires a recorded time');
 data.uiState=request.uiState;
 data.currentSection=request.currentSection||0;
 data.reviewPending=review.status==='pending';
 // Preserve the exact review and final answers on the server FIRST. The
 // coordinator may later approve longitudinal promotion; this is not
 // itself an independent official student-development judgment.
 ciDraftSave_(data,data); // Preserve evaluator-owned draft before final submission.
 return ciPortalEvaluationTransport_('submit',data);
}


/** Evaluator-owned draft lifecycle in the Portal's private Drive repository.
 * Final submission alone uses the separately controlled original CCMS gateway. */
function ciR1EvaluationSaveDraft(token,request){
  const data=ciPortalEvaluationTrustedPayload_(token,request);
  if(!data.submissionId)throw new Error('Draft ID required');
  data.currentSection=request.currentSection;data.reviewPending=request.reviewPending;data.uiState=request.uiState||null;
  return ciDraftSave_(data,data);
}
function ciR1EvaluationLoadDraft(token,request){
  const data=ciPortalEvaluationTrustedPayload_(token,request);
  if(!data.submissionId)throw new Error('Draft ID required');
  return ciDraftLoad_(data);
}
function ciR1EvaluationStop(token,request){
  const data=ciPortalEvaluationTrustedPayload_(token,request);
  if(!data.submissionId)throw new Error('Draft ID required');
  data.currentSection=request.currentSection;data.reviewPending=request.reviewPending;data.uiState=request.uiState||null;
  return ciDraftSave_(data,data);
}
/**
 * Conservative student-aware exam selector. Stable examination IDs are
 * compared ONLY with explicit stable requirement/examination IDs. No name
 * matching, reconstructed prerequisites, or silent enrollment decisions.
 * Unmatched items are labelled REVIEW, never assumed eligible or completed.
 */
function ciR1EvaluationExamOptions(token,siteId,date,studentId) {
  const access=ciHWAccess_(token,siteId,date);
  if(!ciHWGetStudent_(access,studentId))throw new Error('Student is not available');
  const rows=ciDataRows_(access.book,'Examinations',['Examination ID','Examination','Active','College CI Allowed']);
  const requirements=ciDataRows_(access.book,'Student Requirement Status',['Student ID','Status Type','Requirement Scope','Requirement ID','Remaining Count','Requirement Timing']);
  const studentRequirements=requirements.filter(r=>ciDataNorm_(r['Student ID'])===ciDataNorm_(studentId) &&
    ciDataNorm_(r['Status Type'])==='detail' && ciDataNorm_(r['Requirement Scope'])==='program');
  const requirementById=new Map();
  for(const r of studentRequirements){
    const id=String(r['Examination ID']||r['Requirement ID']||'').trim();
    if(!id)continue;
    const key=ciDataNorm_(id);
    if(!requirementById.has(key))requirementById.set(key,[]);
    requirementById.get(key).push(r);
  }
  const result=rows.filter(r=>ciDataTrue_(r.Active)&&ciDataTrue_(r['College CI Allowed']))
    .map(r=>{
      const examinationId=String(r['Examination ID']||'').trim(),examination=String(r.Examination||'').trim();
      const matched=requirementById.get(ciDataNorm_(examinationId))||[];
      let eligibility='review',reason='Eligibility has not been verified against a student-specific requirement.';
      let type='',remaining=null;
      if(matched.length===1){
        const status=ciDataNorm_(matched[0].Status),timing=ciDataNorm_(matched[0]['Requirement Timing']);
        const count=Number(matched[0]['Remaining Count']);
        type=String(matched[0]['Requirement Type']||'');
        if(Number.isFinite(count) && count>=0){
          remaining=count;
          if(count===0){eligibility='completed';reason='Requirement already completed; repeat requires instructor judgment.';}
          else if(/prerequis|review|blocked|not eligible/.test(timing+' '+status)){eligibility='prerequisite';reason='Prerequisite or coordinator review required.';}
          else if(/future|not yet|later/.test(timing+' '+status)){eligibility='future';reason='Future requirement, not currently eligible.';}
          else if(/eligible now|^now$|^eligible$|currently eligible/.test(timing+' '+status)){
            eligibility='eligible';reason='Eligible now according to Student Requirement Status.';
          }else{reason='Student requirement matched, but eligibility status needs review.';}
        }
      }else if(matched.length>1){reason='Multiple student requirement rows match this exam; coordinator review required.';}
      return {examinationId,examination,eligibility,reason,type,remaining};
    }).filter(r=>r.examinationId&&r.examination)
    .sort((a,b)=>({eligible:0,review:1,prerequisite:2,future:3,completed:4}[a.eligibility]-
                        {eligible:0,review:1,prerequisite:2,future:3,completed:4}[b.eligibility])||a.examination.localeCompare(b.examination));
  return {success:true,examinations:result,eligibilitySource:'Student Requirement Status; exact stable-ID matches only',
    note:'Previously completed, future, prerequisite and unverified exams remain selectable for documented instructor judgment. No academic eligibility is silently overridden.'};
}
/**
 * One server-side evaluation bootstrap, reducing multiple sequential
 * browser/server round trips. Every request is independently authorized.
 * An initial private draft is created before opening a new editor; a failed
 * draft write is returned visibly, never misrepresented as saved.
 */
function ciR1EvaluationInitialize(token,request){
  const verified=ciPortalEvaluationTrustedPayload_(token,request);
  const book=ciDataBook_();
  const authorized=ciDataRows_(book,'Examinations',['Examination ID','Active','College CI Allowed']).some(r=>
    ciDataNorm_(r['Examination ID'])===ciDataNorm_(verified.examinationId) && ciDataTrue_(r.Active)&&ciDataTrue_(r['College CI Allowed']));
  if(!authorized)throw new Error('Examination is no longer authorized for this clinical instructor.');
  verified.submissionId=verified.submissionId||'CI-PORTAL-'+Utilities.getUuid();
  const schema=ciPortalLoadOfficialFormSchema_(false);
  const views=ciPortalViewCatalog_();
  const html=HtmlService.createHtmlOutputFromFile('CI_Portal_Competency_Approved_Inner').getContent();
  let initialDraft={success:true,notCreated:Boolean(request.submissionId)};
  if(!request.submissionId){
    try{initialDraft=ciDraftSave_(verified,{answers:{},uiState:null,currentSection:0,reviewPending:true});}
    catch(e){initialDraft={success:false,error:String(e.message||e)};}
  }
  return {success:true,submissionId:verified.submissionId,formItems:schema.formItems,
    schemaLoadedAt:schema.loadedAt,views,html,initialDraft};
}
function ciR1NewEvaluationId(token,siteId,date,studentId){
  const access=ciHWAccess_(token,siteId,date);
  if(!ciHWGetStudent_(access,studentId))throw new Error('Student is not available');
  return {success:true,submissionId:'CI-PORTAL-'+Utilities.getUuid()};
}
function ciR1ApprovedEvaluationInnerHtml(token){
  ciR1Require_(token);
  return HtmlService.createHtmlOutputFromFile('CI_Portal_Competency_Approved_Inner').getContent();
}

function ciR1EvaluationDraftList(token){
 const evaluator=ciR1Require_(token);
 if(String(evaluator.evaluatorType||'').toLowerCase()!=='college ci')throw new Error('Hospital CI evaluation access not yet enabled');
 return ciDraftList_(evaluator.evaluatorId);
}
function ciR1EvaluationReleaseStatus(token,studentId){
 ciR1Require_(token);
 const mode=ciPortalEvaluationMode_();
 const permitted=mode==='live'||(mode==='test_only'&&String(studentId||'')==='TEST-001');
 if(!permitted)return {success:true,submitEnabled:false,reason:'Official submission has not been released. Draft saving and grading do not require the gateway.'};
 const connected=ciPortalEvaluationHasConnection_();
 return {success:true,submitEnabled:connected,
   reason:connected?'':'Official submission requires a separately authorized CCMS connection.'};
}

