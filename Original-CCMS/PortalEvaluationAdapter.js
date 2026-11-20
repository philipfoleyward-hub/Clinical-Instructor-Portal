/**
 * CCMS Portal → native Digital Clinical Evaluation Form bridge.
 * Install in the ORIGINAL CCMS Apps Script project ONLY.
 * Authoritative engine: existing gradeClinicalEvaluation(FormResponse).
 *
 * PREVIEW IS READ-ONLY: no FormResponse submit(), Sheets/Drive writes or MailApp.
 * This file never creates a gradebook entry. Only the existing authoritative
 * digitalEvaluation_submitPortal_ intake is allowed to commit a graded result.
 *
 * Do NOT also install the previous 04_PortalEvaluationAdapter candidate. Both
 * expose the same integration names; the old one requires grading_gradeMappedEvaluation_
 * (not present in the production engine successfully tested on 2026-10-16).
 */
function ccmsPortal_clean_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}
function ccmsPortal_norm_(value) {
  return ccmsPortal_clean_(value).toLowerCase().replace(/\s+/g, ' ');
}
function ccmsPortal_true_(value) {
  return ['true','yes','y','1','active'].includes(ccmsPortal_norm_(value));
}
function ccmsPortal_reference_(book, sheetName, idColumn, id, requiredColumns) {
  const sheet=book.getSheetByName(sheetName);
  if (!sheet) throw new Error('Authoritative sheet unavailable: '+sheetName);
  const values=sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Authoritative sheet has no headers: '+sheetName);
  const headers=values[0].map(ccmsPortal_clean_);
  [idColumn].concat(requiredColumns||[]).forEach(h=>{
    if (!headers.includes(h)) throw new Error(sheetName+' missing required header: '+h);
  });
  const col=headers.indexOf(idColumn),target=ccmsPortal_norm_(id);
  if(!target)throw new Error('A '+idColumn+' is required');
  const matches=values.slice(1).filter(row=>ccmsPortal_norm_(row[col])===target);
  if(matches.length!==1)throw new Error('Missing or ambiguous reference in '+sheetName+' for '+idColumn);
  return Object.fromEntries(headers.map((h,i)=>[h,matches[0][i]]));
}
function ccmsPortal_date_(raw) {
  const text=ccmsPortal_clean_(raw),m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if(!m)throw new Error('Evaluation date must be YYYY-MM-DD');
  const date=new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),12,0,0);
  if(date.getFullYear()!==Number(m[1])||date.getMonth()!==Number(m[2])-1||date.getDate()!==Number(m[3]))
    throw new Error('Evaluation date is invalid');
  return date;
}
/** Resolves ALL academic identity text from trusted CCMS reference rows. */
function ccmsPortal_resolveIdentity_(data) {
  if(!data||typeof data!=='object'||Array.isArray(data))throw new Error('Evaluation data is required');
  if(ccmsPortal_norm_(data.evaluationType)!=='competency')throw new Error('Only Competency evaluations are authorized');
  const claimed=ccmsPortal_clean_(data.authenticatedEvaluatorId);
  if(!claimed||claimed!==ccmsPortal_clean_(data.evaluatorId))throw new Error('Authenticated evaluator mismatch');
  const book=SpreadsheetApp.getActiveSpreadsheet();
  if(!book)throw new Error('Original CCMS workbook context unavailable');
  const student=ccmsPortal_reference_(book,APP_CONFIG.SHEETS.STUDENTS,'Student ID',data.studentId,['Student','Active']);
  if(!ccmsPortal_true_(student.Active))throw new Error('Student is not active');
  const site=ccmsPortal_reference_(book,APP_CONFIG.SHEETS.CLINICAL_SITES,'Clinical Site ID',data.siteId,['Clinical Site','Active']);
  if(!ccmsPortal_true_(site.Active))throw new Error('Clinical site is not active');
  const evaluator=ccmsPortal_reference_(book,APP_CONFIG.SHEETS.EVALUATORS,'Evaluator ID',data.evaluatorId,
    ['Evaluator','Evaluator Type','Active','Portal Access']);
  if(!ccmsPortal_true_(evaluator.Active)||!ccmsPortal_true_(evaluator['Portal Access'])||
      ccmsPortal_norm_(evaluator['Evaluator Type'])!=='college ci')
    throw new Error('Evaluator is not an active, authorized College CI');
  const examination=ccmsPortal_reference_(book,APP_CONFIG.SHEETS.EXAMINATIONS,'Examination ID',data.examinationId,
    ['Examination','Active','College CI Allowed']);
  if(!ccmsPortal_true_(examination.Active)||!ccmsPortal_true_(examination['College CI Allowed']))
    throw new Error('Examination not available to College CI');
  return {
    studentId:ccmsPortal_clean_(student['Student ID']),student:ccmsPortal_clean_(student.Student),
    siteId:ccmsPortal_clean_(site['Clinical Site ID']),site:ccmsPortal_clean_(site['Clinical Site']),
    evaluatorId:ccmsPortal_clean_(evaluator['Evaluator ID']),evaluator:ccmsPortal_clean_(evaluator.Evaluator),
    evaluatorType:ccmsPortal_clean_(evaluator['Evaluator Type']),
    examinationId:ccmsPortal_clean_(examination['Examination ID']),examination:ccmsPortal_clean_(examination.Examination),
    examinationCategory:ccmsPortal_clean_(examination['Examination Category']||examination.Category||''),
    date:ccmsPortal_date_(data.date)
  };
}
function ccmsPortal_choices_(item) {
  const type = item.getType();
  const types = FormApp.ItemType;

  if (type === types.GRID) {
    return item.asGridItem().getColumns().map(String);
  }

  let choiceItem;

  if (type === types.CHECKBOX) {
    choiceItem = item.asCheckboxItem();
  } else if (type === types.MULTIPLE_CHOICE) {
    choiceItem = item.asMultipleChoiceItem();
  } else if (type === types.LIST) {
    choiceItem = item.asListItem();
  } else {
    return null;
  }

  return choiceItem.getChoices().map(choice =>
    String(choice.getValue())
  );
}
/** Returns the real Google Form item IDs/rows/choices for the locked UI. */
function ccmsPortal_schema_() {
  const form=FormApp.openById(APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION),t=FormApp.ItemType;
  return form.getItems().map(item=>{
    const type=String(item.getType());
    const result={itemId:String(item.getId()),title:String(item.getTitle()),type};
    if(item.getType()===t.GRID)result.rows=item.asGridItem().getRows();
    const choices=ccmsPortal_choices_(item);
    if(choices)result.choices=choices;
    return result;
  });
}
function ccmsPortal_responseMap_(form,data,identity) {
  const items=form.getItems(),types=FormApp.ItemType;
  const byId=new Map(items.map(i=>[String(i.getId()),i]));
  const answers=data.answers;
  if(!answers||typeof answers!=='object'||Array.isArray(answers))throw new Error('Official Form item-ID answer map required');
  const keys=Object.keys(answers);
  if(keys.length>500||JSON.stringify(answers).length>180000)throw new Error('Portal answer payload too large');
  const map=Object.create(null),protectedTitles=new Set([
    'student','clinical site','evaluator','date','evaluation date','examination','evaluation type'
  ]);
  for(const id of keys) {
    if(!/^\d+$/.test(id)||!byId.has(id))throw new Error('Unknown Form item ID: '+id);
    const item=byId.get(id),title=ccmsPortal_norm_(item.getTitle()),type=item.getType(),raw=answers[id];
    if(protectedTitles.has(title))throw new Error('Identity answer is managed by CCMS: '+title);
    if(type===types.PAGE_BREAK||type===types.SECTION_HEADER)throw new Error('Page and section headers cannot receive answers');
    if(type===types.GRID) {
      const rows=item.asGridItem().getRows(),allowed=ccmsPortal_choices_(item)||[];
      if(!Array.isArray(raw)||raw.length!==rows.length)throw new Error('Criterion row count mismatch: '+item.getTitle());
      map[id]=raw.map(value=>{
        const score=ccmsPortal_clean_(value);
        if(!['0','1','2','3','N/A'].includes(score)||!allowed.includes(score))
          throw new Error('Invalid Form criterion choice: '+item.getTitle());
        return score;
      });
    } else if(type===types.CHECKBOX) {
      const choices=ccmsPortal_choices_(item)||[];
      if(!Array.isArray(raw)||raw.length>25||raw.some(v=>typeof v!=='string'||v.length>350||!choices.includes(v)))
        throw new Error('Invalid checkbox answer: '+item.getTitle());
      map[id]=raw.slice();
    } else if(type===types.MULTIPLE_CHOICE||type===types.LIST) {
      if(typeof raw!=='string'||!ccmsPortal_choices_(item).includes(raw))throw new Error('Invalid official Form choice: '+item.getTitle());
      map[id]=raw;
    } else if(type===types.TEXT||type===types.PARAGRAPH_TEXT) {
      if(typeof raw!=='string'||raw.length>8000)throw new Error('Invalid text response: '+item.getTitle());
      map[id]=raw;
    } else {
      // No silent coercion for dates, uploads, scales or unsupported Form item types.
      throw new Error('Unsupported Form item type for '+item.getTitle()+'; requires schema review');
    }
  }
  const identityByTitle={
    'student':identity.student,'clinical site':identity.site,'evaluator':identity.evaluator,
    'examination':identity.examination,'evaluation type':'Competency',
    'date':identity.date,'evaluation date':identity.date
  };
  const found={student:0,site:0,evaluator:0,exam:0,date:0};
  // The official Form selects the exam through "Examination Category" plus one
  // "Examination - <Category>" dropdown per category (the grading engine reads
  // any title starting with "examination" other than the category). Answer the
  // single category dropdown that lists this exam, and its category.
  const examItems=items.filter(item=>{
    const t=ccmsPortal_norm_(item.getTitle());
    return t!=='examination'&&t.startsWith('examination')&&t!=='examination category'&&
      [types.LIST,types.MULTIPLE_CHOICE].includes(item.getType());
  });
  if(examItems.length){
    let hits=examItems.filter(item=>(ccmsPortal_choices_(item)||[]).includes(String(identity.examination)));
    const wanted=ccmsPortal_norm_(identity.examinationCategory||'');
    if(hits.length>1&&wanted)hits=hits.filter(item=>ccmsPortal_norm_(item.getTitle()).endsWith(wanted));
    if(hits.length!==1)throw new Error('Examination "'+identity.examination+'" must appear in exactly one official Form category list (found '+hits.length+')');
    const examItem=hits[0];
    for(const other of examItems)delete map[String(other.getId())];
    map[String(examItem.getId())]=identity.examination;
    found.exam++;
    const categoryItem=items.find(item=>ccmsPortal_norm_(item.getTitle())==='examination category'&&
      [types.LIST,types.MULTIPLE_CHOICE].includes(item.getType()));
    if(categoryItem){
      const examTitle=ccmsPortal_norm_(examItem.getTitle());
      const category=(ccmsPortal_choices_(categoryItem)||[]).filter(c=>examTitle.endsWith(ccmsPortal_norm_(c)))
        .sort((a,b)=>b.length-a.length)[0];
      if(category)map[String(categoryItem.getId())]=category;
    }
  }
  for(const item of items) {
    const title=ccmsPortal_norm_(item.getTitle()),id=String(item.getId());
    if(!Object.prototype.hasOwnProperty.call(identityByTitle,title))continue;
    const value=identityByTitle[title],type=item.getType();
    if([types.LIST,types.MULTIPLE_CHOICE].includes(type)) {
      const options=ccmsPortal_choices_(item)||[];
      if(!options.includes(String(value)))throw new Error('CCMS reference does not match the official Form choice: '+title);
    }
    map[id]=value;
    if(title==='student')found.student++;
    if(title==='clinical site')found.site++;
    if(title==='evaluator')found.evaluator++;
    if(title==='examination')found.exam++;
    if(title==='date'||title==='evaluation date')found.date++;
  }
  if(Object.values(found).some(v=>v!==1))throw new Error('Official Form identity field count changed: '+JSON.stringify(found));
  const automaticFailure=typeof grading_readAutomaticFailure_==='function'
    ?grading_readAutomaticFailure_(form,map):{occurred:false,reasons:[],comments:''};
  if(automaticFailure.occurred&&(!automaticFailure.reasons.length||!automaticFailure.comments))
    throw new Error('Auto-Fail requires an official reason and explanation');
  if(!automaticFailure.occurred) {
    const expected=new Set(['exam preparation','patient care','radiation protection',
      'equipment manipulation','positioning skills','portable / contrast']);
    let count=0;
    for(const item of items) {
      if(item.getType()!==types.GRID||!expected.has(ccmsPortal_norm_(item.getTitle())))continue;
      count++;
      if(!Object.prototype.hasOwnProperty.call(map,String(item.getId())))
        throw new Error('Complete the official '+item.getTitle()+' section before grading');
    }
    if(count!==6)throw new Error('Official Form performance section contract changed: expected six grids, found '+count);
  }
  // Image check is based on authoritative Form page titles, not hardcoded
  // radiograph projections. Any begun view must have all four scored parts.
  let usedViews=0;
  for(let i=0;i<items.length;i++) {
    const item=items[i];
    if(item.getType()!==types.PAGE_BREAK||!/^Image Evaluation\s*-\s*View\s+\d+$/i.test(ccmsPortal_clean_(item.getTitle())))continue;
    const group=[];
    for(let j=i+1;j<items.length&&items[j].getType()!==types.PAGE_BREAK;j++)group.push(items[j]);
    const answered=group.filter(part=>{
      const v=map[String(part.getId())];return v!==undefined&&v!==null&&v!==''&&(!Array.isArray(v)||v.length);
    });
    if(!answered.length)continue;
    usedViews++;
    const projection=group.find(p=>ccmsPortal_norm_(p.getTitle())==='view/projection');
    if(!projection||!ccmsPortal_clean_(map[String(projection.getId())]))throw new Error('Started image view requires a projection');
    for(const partName of ['positioning','equipment','radiation protection','exposure']) {
      const matching=group.filter(p=>ccmsPortal_norm_(p.getTitle())===partName);
      if(matching.length!==1||!ccmsPortal_clean_(map[String(matching[0].getId())]))
        throw new Error('Complete all four scored components for a started image view: '+partName);
    }
  }
  if(usedViews===0)throw new Error('At least one scored image view is required for the official grade');
  return map;
}
/** A FormResponse-shaped in-memory object. It NEVER calls createResponse or submit. */
function ccmsPortal_memoryFormResponse_(form,responseMap,submissionId) {
  const itemResponses=form.getItems().filter(item=>
    Object.prototype.hasOwnProperty.call(responseMap,String(item.getId()))).map(item=>({
      getItem:function(){return item;},
      getResponse:function(){return responseMap[String(item.getId())];}
    }));
  return {
    getId:function(){return ccmsPortal_clean_(submissionId)||'CI-PORTAL-READ-ONLY-PREVIEW';},
    getTimestamp:function(){return new Date();},
    getItemResponses:function(){return itemResponses;}
  };
}
/**
 * Stable contract used by the existing signed CCMS gateway and DigitalIntake.
 * The only grading call is to the production gradeClinicalEvaluation() function.
 */
function gradeClinicalEvaluationFromPortalPayload(data) {
  const identity=ccmsPortal_resolveIdentity_(data);
  const form=FormApp.openById(APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION);
  const map=ccmsPortal_responseMap_(form,data,identity);
  if(typeof gradeClinicalEvaluation!=='function')throw new Error('Authoritative native grading function unavailable');
  const grade=gradeClinicalEvaluation(ccmsPortal_memoryFormResponse_(form,map,data.submissionId));
  if(!grade||!Number.isFinite(Number(grade.finalScore))||!['PASS','FAIL'].includes(String(grade.result)))
    throw new Error('Authoritative grading engine returned an incomplete result');
  // Verify the actual engine saw the server-resolved identity, not a client alias.
  const evalInfo=grade.evaluation||{};
  if(ccmsPortal_norm_(evalInfo.student)!==ccmsPortal_norm_(identity.student)||
      ccmsPortal_norm_(evalInfo.clinicalSite)!==ccmsPortal_norm_(identity.site)||
      ccmsPortal_norm_(evalInfo.evaluator)!==ccmsPortal_norm_(identity.evaluator)||
      ccmsPortal_norm_(evalInfo.evaluationType)!=='competency'||
      ccmsPortal_norm_(evalInfo.examination)!==ccmsPortal_norm_(identity.examination))
    throw new Error('Grade identity does not agree with authoritative CCMS references');
  return {grade:grade,identity:identity};
}
