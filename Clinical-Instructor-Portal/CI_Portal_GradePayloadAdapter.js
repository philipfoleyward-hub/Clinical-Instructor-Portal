/**
 * CCMS Portal — Official Form answer-to-grader projection (read-only).
 * The Portal's current UI sends real Google Form item IDs in answers.
 * This adapter resolves those same item IDs and row labels directly from the
 * official Google Form; it does not invent scoring groups, section weights,
 * image penalties, form schema, grade values, or a second academic policy.
 * NEVER writes a FormResponse, grade, Drive record, or email.
 */
function ciPortalGradeMapOfficialAnswers_(trusted) {
  if (!trusted||typeof trusted!=='object'||!trusted.answers||typeof trusted.answers!=='object'||Array.isArray(trusted.answers)) {
    throw new Error('Official grade preview requires evaluation answers.');
  }
  const formId=String(PropertiesService.getScriptProperties().getProperty('CI_PORTAL_OFFICIAL_EVALUATION_FORM_ID')||
    '1TGm4ciMz3YkKv902gON-FmnofNUORqMmKySVA5pPrsY').trim();
  // Must match the official Form used by ciR1EvaluationFormSchema.
  const items=ciPortalGradingFormItems_();
  const answer=trusted.answers;
  const id=i=>String(i.getId());
  const clean=v=>String(v===null||v===undefined?'':v).trim();
  const norm=v=>clean(v).toLowerCase().replace(/\s+/g,' ');
  const value=i=>answer[id(i)];
  const expectedGrid=['exam preparation','patient care','radiation protection',
    'equipment manipulation','positioning skills','portable / contrast'];
  const grids=items.filter(i=>i.getType()===FormApp.ItemType.GRID&&expectedGrid.includes(norm(i.getTitle())));
  if(grids.length!==expectedGrid.length||new Set(grids.map(g=>norm(g.getTitle()))).size!==expectedGrid.length)
    throw new Error('Official competency Form performance grids changed. No grade calculated.');
  const performance={sections:grids.map(grid=>{
    const rows=grid.asGridItem().getRows();
    const supplied=value(grid);
    if(supplied!==undefined&&!Array.isArray(supplied))throw new Error('Invalid scoring grid response: '+grid.getTitle());
    if(supplied&&supplied.length!==rows.length)throw new Error('Incomplete official scoring grid: '+grid.getTitle());
    const ratings=rows.map((_,i)=>supplied&&supplied[i]!==undefined?clean(supplied[i]):'');
    for(const rating of ratings)if(rating&&!['0','1','2','3','n/a'].includes(norm(rating)))
      throw new Error('Invalid official clinical criterion rating');
    return {name:String(grid.getTitle()),rows:rows.map((criterion,i)=>({criterion:String(criterion),rating:ratings[i]}))};
  })};
  const automaticFailure={occurred:false,reasons:[],comments:''};
  const images={views:[]};
  let view=null;
  const componentKeys={positioning:'positioning',equipment:'equipment',
    'radiation protection':'radiationProtection',exposure:'exposure'};
  const permittedIds=new Set(items.map(id));
  for(const key of Object.keys(answer))if(!permittedIds.has(String(key)))
    throw new Error('The evaluation contains an unrecognized Form item ID. Reload the official Form schema.');
  for(const item of items){
    const type=item.getType();
    const title=norm(item.getTitle());
    const raw=value(item);
    if(type===FormApp.ItemType.PAGE_BREAK){
      view=null;
      const match=String(item.getTitle()).match(/^Image Evaluation\s*-\s*View\s+(\d+)$/i);
      if(match){
        view={projection:'',components:{},exposureIndex:'',marker:'',comments:''};
        images.views[Number(match[1])-1]=view;
      }
      continue;
    }
    if(title==='did an automatic failure occur?'||
        title==='did an automatic-failure condition occur during this evaluation?'){
      automaticFailure.occurred=automaticFailure.occurred||norm(raw)==='yes';
    } else if(title==='automatic failure reason(s)'){
      const list=(Array.isArray(raw)?raw:raw?[raw]:[]).map(clean).filter(Boolean);
      const valid=item.getType()===FormApp.ItemType.CHECKBOX?
        item.asCheckboxItem().getChoices().map(c=>clean(c.getValue())):[];
      if(list.some(s=>!valid.includes(s)))throw new Error('Unrecognized official automatic failure reason');
      automaticFailure.reasons=list;
    } else if(title==='automatic failure comments'){
      automaticFailure.comments=clean(raw);
    }
    if(!view)continue;
    if(title==='view/projection')view.projection=clean(raw);
    else if(Object.prototype.hasOwnProperty.call(componentKeys,title)){
      if(raw!==undefined&&raw!==null&&clean(raw)!==''){
        if(!/^[0-3]([\s=:;\-].*)?$|^n\/a(\s.*)?$/i.test(clean(raw)))throw new Error('Invalid image score');
        view.components[componentKeys[title]]=clean(raw);
      }
    }else if(title.includes('exposure index')||title.includes('s number'))view.exposureIndex=clean(raw);
    else if(title==='marker placement'||title==='marker evaluation')view.marker=clean(raw);
    else if(/^view \d+ comments$/.test(title)||title==='comments'||title==='comment')view.comments=clean(raw);
  }
  // Preserve the actual submitted criterion details on Auto-Fail even though
  // the handbook overrides the final *recorded* score to 0.0.
  const result={
    evaluation:{studentId:clean(trusted.studentId),clinicalSiteId:clean(trusted.siteId),
      evaluatorId:clean(trusted.evaluatorId),examinationId:clean(trusted.examinationId),
      evaluationType:'Competency',date:clean(trusted.date)},
    automaticFailure,performance,image:{views:images.views.filter(Boolean)}
  };
  if(automaticFailure.occurred&&(!automaticFailure.reasons.length||!automaticFailure.comments))
    throw new Error('Automatic Failure requires an official reason and evaluator explanation.');
  return result;
}

/**
 * Grade-preview projection from the existing approved evaluation's current
 * scores and image selections. Verifies the official Form row sequence and
 * authorized Auto-Fail choices before invoking the existing grading engine.
 * Only a read-only preview uses this path; official submission retains its
 * separate canonical Form-ID mapping and release gates.
 */
function ciPortalGradeMapUIState_(trusted, state) {
  if (!state || typeof state !== 'object' || Array.isArray(state))
    throw new Error('Evaluation score state was not supplied. Reopen the evaluation and try again.');
  const uiScores=state.scores,uiImages=state.images;
  if (!uiScores || typeof uiScores!=='object' || Array.isArray(uiScores))
    throw new Error('Evaluation criterion scores are unavailable.');
  if (!Array.isArray(uiImages))throw new Error('Evaluation image views are unavailable.');
  const formId=String(PropertiesService.getScriptProperties().getProperty('CI_PORTAL_OFFICIAL_EVALUATION_FORM_ID')||
    '1TGm4ciMz3YkKv902gON-FmnofNUORqMmKySVA5pPrsY').trim();
  const items=ciPortalGradingFormItems_();
  const clean=v=>String(v===null||v===undefined?'':v).trim();
  const norm=v=>clean(v).toLowerCase().replace(/\s+/g,' ');
  const rubric=['Exam Preparation','Patient Care','Radiation Protection','Equipment Manipulation','Positioning Skills','Portable / Contrast'];
  const found=items.filter(i=>i.getType()===FormApp.ItemType.GRID && rubric.some(t=>norm(t)===norm(i.getTitle())));
  if(found.length!==6||new Set(found.map(g=>norm(g.getTitle()))).size!==6)
    throw new Error('Official Form performance rubric changed. Grading halted; verify the authoritative Form.');
  const failure=state.autoFail||{};
  const af={occurred:failure.occurred===true,reasons:Array.isArray(failure.reasons)?failure.reasons.map(clean).filter(Boolean):[],comments:clean(failure.comments)};
  const officialFailureChoices=items.find(i=>norm(i.getTitle())==='automatic failure reason(s)');
  if(af.occurred){
    if(!officialFailureChoices || officialFailureChoices.getType()!==FormApp.ItemType.CHECKBOX)
      throw new Error('Official automatic-failure reasons are unavailable.');
    const allowed=officialFailureChoices.asCheckboxItem().getChoices().map(c=>clean(c.getValue()));
    if(!af.reasons.length||af.reasons.some(x=>!allowed.includes(x))||!af.comments)
      throw new Error('Record an approved Auto-Fail reason and explanation.');
  }
  const getRating=(k,desc)=>{
    const v=uiScores[k];
    if(v===undefined || v===null || v===''){
      if(af.occurred)return '';
      throw new Error('Incomplete performance scoring: '+desc+'.');
    }
    const x=clean(v);
    if(!['0','1','2','3','N/A'].includes(x))throw new Error('Invalid clinical criterion rating at '+desc);
    return x;
  };
  const performance={sections:rubric.map((name,group)=>{
    const grid=found.find(g=>norm(g.getTitle())===norm(name));
    const rows=grid.asGridItem().getRows();
    const half=group===5?uiScores:null;
    if(group===5 && rows.length!==16)throw new Error('Combined Portable / Contrast rubric is no longer 16 criteria.');
    return {name:String(grid.getTitle()), rows:rows.map((criterion,i)=>{
      const section=group===5?(i<7?5:6):group;
      const index=group===5?(i<7?i:i-7):i;
      return {criterion:String(criterion),rating:getRating(section+'-'+index,name+' criterion '+(i+1))};
    })};
  })};
  const viewSlots=items.filter(i=>i.getType()===FormApp.ItemType.PAGE_BREAK &&
    /^Image Evaluation\s*-\s*View\s+\d+$/i.test(clean(i.getTitle()))).length;
  if(viewSlots===0)throw new Error('Official image-view rubric is unavailable.');
  if(uiImages.length>viewSlots)throw new Error('Evaluation contains more image views than the official Form.');
  const views=[];
  const ratingKeys=['positioning','equipment','radiationProtection','exposure'];
  for(let i=0;i<uiImages.length;i++){
    const v=uiImages[i]||{};
    const selected=Boolean(clean(v.name)||Object.keys(v.scores||{}).length||v.present||v.side||v.legible||v.ei||v.comment);
    if(!selected)continue;
    const projection=clean(v.name);
    if(!projection && !af.occurred)throw new Error('Image view '+(i+1)+' needs a View/Projection name.');
    const components={};
    for(let j=0;j<ratingKeys.length;j++){
      let rating=(v.scores||{})[j];
      if(rating===undefined||rating===null||rating===''){
        if(!af.occurred)throw new Error('Image view '+(i+1)+' is missing '+ratingKeys[j]+' score.');
      }else{
        rating=clean(rating);
        if(!['0','1','2','3','N/A'].includes(rating))throw new Error('Invalid image score in view '+(i+1));
        components[ratingKeys[j]]=rating;
      }
    }
    if(!af.occurred && (!v.present||!v.side||!v.legible))
      throw new Error('Image view '+(i+1)+' needs all three marker selections.');
    const marker=v.present==='Missing'?'Missing Marker':v.side==='Wrong'?'Wrong Marker':v.legible==='Illegible'?'Illegible Marker':'Correct';
    if(!['','Present','Missing'].includes(clean(v.present))||
       !['','Correct','Wrong'].includes(clean(v.side))||
       !['','Legible','Illegible'].includes(clean(v.legible)))
      throw new Error('Image marker selection is invalid in view '+(i+1));
    views.push({projection,components,exposureIndex:clean(v.ei),marker,comments:clean(v.comment)});
  }
  if(!af.occurred && views.length===0)throw new Error('Choose and score at least one image view in Image Evaluation.');
  return {
    evaluation:{studentId:clean(trusted.studentId),clinicalSiteId:clean(trusted.siteId),
      evaluatorId:clean(trusted.evaluatorId),examinationId:clean(trusted.examinationId),
      evaluationType:'Competency',date:clean(trusted.date),comments:clean(state.sectionComments&&state.sectionComments[8])},
    automaticFailure:af,performance,image:{views}
  };
}
