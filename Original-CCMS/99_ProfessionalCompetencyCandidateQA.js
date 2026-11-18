/**
 * READ-ONLY QA. Do not call the Form-submit handler manually.
 * qaPcV2PureGradeQA() performs NO Forms, spreadsheet, Drive, grades, or email writes.
 * qaPcV2ReadOnlyPreflight() reads the Form, templates, site service, and properties only.
 */
function qaPcV2PureGradeQA() {
  const tests=[];
  Object.keys(PC_V2_RUBRICS).forEach(function(family){
    const rubric=PC_V2_RUBRICS[family];
    const sum=rubric.reduce(function(s,x){return s+x.weight;},0);
    const expected=family==='Clinical 1-2'?{count:19,weight:43}:{count:26,weight:62};
    if(rubric.length!==expected.count||sum!==expected.weight)throw new Error('Rubric map count/weight QA failed for '+family);
    [
      {label:'all 3',value:'3',raw:3,percent:100},
      {label:'all 2',value:'2',raw:2,percent:75},
      {label:'all 1',value:'1',raw:1,percent:50},
      {label:'all 0',value:'0',raw:0,percent:25}
    ].forEach(function(test){
      const byId={};rubric.forEach(function(x){byId[x.id]=test.value;});
      const s=pcV2_calculate_(family,{byId:byId});
      if(s.rawScore!==test.raw||s.percentGrade!==test.percent)throw new Error('Score QA failed: '+family+'/'+test.label);
      tests.push({family:family,label:test.label,score:s.rawScore,percent:s.percentGrade});
    });
    const nA={};rubric.forEach(function(x){nA[x.id]='N/A';});
    nA[rubric[0].id]='2';
    const subset=pcV2_calculate_(family,{byId:nA});
    if(subset.applicableWeight!==rubric[0].weight || subset.percentGrade!==75)throw new Error('N/A QA failed');
    tests.push({family:family,label:'N/A weight exclusion',applicableWeight:subset.applicableWeight,percent:subset.percentGrade});
    let rejected=false;
    try {pcV2_calculate_(family,{byId:Object.assign({},nA,{[rubric[0].id]:'N/A'})});}catch(e){rejected=true;}
    if(!rejected)throw new Error('All-N/A QA failed');
  });
  if(pcV2_round2_((2.75+1)*25)!==93.75)throw new Error('Approved percentage conversion QA failed');
  const qaProperties={getProperty:function(){return '';}};
  if(pcV2_passResult_(75,qaProperties)!=='Pass' || pcV2_passResult_(74.75,qaProperties)!=='Fail' ||
      pcV2_passResult_(100,qaProperties)!=='Pass' || pcV2_passResult_(25,qaProperties)!=='Fail')
    throw new Error('Approved 75% pass/fail boundary QA failed');
  let rejectedConflict=false;
  try{pcV2_passResult_(75,{getProperty:function(){return '70';}});}catch(_){rejectedConflict=true;}
  if(!rejectedConflict)throw new Error('Conflicting legacy pass-threshold property was not rejected');
  const emailPreview=pcV2_emailPayload_({student:{student:'TEST Student',studentId:'TEST-001'},
    evaluator:'TEST Evaluator',clinicalTerm:'Clinical 3',site:{name:'TEST SITE'},
    date:new Date(2026,10,5),score:{rawScore:2.75,percentGrade:93.75},
    result:'Pass',sourceId:'PC-TEST'},
    {gradeRecordId:'GR-TEST'},'https://example.invalid/report.pdf','https://example.invalid/report');
  ['Student: TEST Student','Evaluator: TEST Evaluator','Clinical site: TEST SITE','Result: Pass','Attached: completed Professional Competency evaluation (PDF)'].forEach(function(part){
    if(emailPreview.body.indexOf(part)===-1)throw new Error('Email preview missing '+part);
  });
  return {success:true,tests:tests,emailPreview:emailPreview.body,
    approvedPassMinimumPercent:75,warning:'Email preview is an illustrative fixture; boundary QA verifies 75% as approved cutoff.'};
}
function qaPcV2ReadOnlyPreflight() {
  const form=FormApp.openById(PC_V2_FORM_ID);
  const criteria=pcV2_formSnapshot_(form); // checks all 45 exact Form IDs, wording, weights and options
  Object.keys(PC_V2_TEMPLATES).forEach(function(k){pcV2_validateTemplate_(PC_V2_TEMPLATES[k]);});
  const props=PropertiesService.getScriptProperties();
  const threshold=props.getProperty('CCMS_PC_PASS_MIN_PERCENT');
  return {success:true,formId:PC_V2_FORM_ID,
    rubricCounts:Object.fromEntries(Object.keys(criteria).map(function(k){return [k,criteria[k].length];})),
    templateIds:PC_V2_TEMPLATES,
    approvedPassMinimumPercent:75,
    passThresholdLegacyProperty:threshold||null,
    thresholdPropertyConflict:!!(threshold && Number(threshold)!==75),
    enabled:props.getProperty('CCMS_PC_V2_ENABLED')==='true',
    deployment:pcV2_readOnlyDeploymentStatus(),
    warning:'Preflight performs no submission, grade write, PDF generation or email. Four Form choice sets may block preflight pending validation.'};
}

/** Read ONLY the four unusual item choice sets if full Form preflight fails. */
function qaPcV2ChoiceExceptionsReadOnly() {
  const form=FormApp.openById(PC_V2_FORM_ID);
  const ids=[100711765,1218665472,1375968503,1946107895];
  const details=ids.map(function(id){
    const item=form.getItemById(id);
    return {id:id,title:item.getTitle(),weightHelp:item.getHelpText(),
      isRequired:item.asListItem().isRequired(),
      choices:item.asListItem().getChoices().map(function(c){return c.getValue();})};
  });
  console.log('CCMS_PC_CHOICE_EXCEPTIONS=' + JSON.stringify(details));
  return details;
}
