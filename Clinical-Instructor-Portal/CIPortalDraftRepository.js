/**
 * Clinical Instructor Portal — evaluator-owned, unsent competency drafts.
 * This is a distinct pre-submission lifecycle. It NEVER files final clinical
 * records, modifies Clinical Tracking, generates final documents, or emails.
 *
 * App-authorized access: only ciR1Evaluation* callers may invoke these private
 * helpers; those callers reauthorize an actual signed-in evaluator, clinical
 * site and stable student ID before reaching this repository.
 *
 * The folder ID is the durable routing key, not its human-readable name.
 * Its first creation uses the web app's deploying Google account. The default
 * folder is private to that account. CI_PORTAL_DRAFT_FOLDER_ID can instead be
 * pre-configured to an approved private folder without creating a new one.
 * No shared gateway key or original-CCMS web app is needed for draft operations.
 */
function ciDraftFolder_() {
  const props=PropertiesService.getScriptProperties();
  const key='CI_PORTAL_DRAFT_FOLDER_ID';
  const configured=String(props.getProperty(key)||'').trim();
  if(configured){
    if(!/^[A-Za-z0-9_-]{20,}$/.test(configured))throw new Error(key+' must contain a valid Drive folder ID');
    return DriveApp.getFolderById(configured);
  }
  // Use the same script-wide lock for first creation, avoiding competing folders.
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    const again=String(props.getProperty(key)||'').trim();
    if(again)return DriveApp.getFolderById(again);
    const folder=DriveApp.createFolder('CCMS — Clinical Instructor Portal — In Progress');
    props.setProperty(key,folder.getId());
    return folder;
  }finally{lock.releaseLock();}
}
function ciDraftId_(value){
  const id=String(value||'');
  if(!/^CI-PORTAL-[a-f0-9-]{36}$/i.test(id))throw new Error('Invalid stable evaluation draft ID');
  return id;
}
function ciDraftFind_(folder,id){
  const files=folder.getFilesByName(ciDraftId_(id)+'.json');
  if(!files.hasNext())return null;
  const file=files.next();
  if(files.hasNext())throw new Error('Duplicate draft IDs found; coordinator reconciliation required');
  return file;
}
function ciDraftParsed_(file){
  const contents=file.getBlob().getDataAsString('UTF-8');
  if(contents.length>250000)throw new Error('Evaluation draft is unexpectedly large');
  const data=JSON.parse(contents);
  if(!data||data.version!==1||ciDraftId_(data.submissionId)!==file.getName().slice(0,-5))
    throw new Error('Draft identity or version mismatch');
  return data;
}
function ciDraftOwned_(data,identity){
  if(String(data.evaluatorId)!==String(identity.evaluatorId)||
     String(data.studentId)!==String(identity.studentId)||
     String(data.siteId)!==String(identity.siteId)||
     String(data.examinationId)!==String(identity.examinationId)||
     String(data.date)!==String(identity.date))
    throw new Error('This draft does not belong to the selected evaluator/student/exam/site/date');
}
function ciDraftSave_(identity,payload){
  const id=ciDraftId_(identity.submissionId),folder=ciDraftFolder_();
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const existing=ciDraftFind_(folder,id);
    if(existing)ciDraftOwned_(ciDraftParsed_(existing),identity);
    const data={
      version:1,submissionId:id,studentId:String(identity.studentId),
      siteId:String(identity.siteId),evaluatorId:String(identity.evaluatorId),
      examinationId:String(identity.examinationId),date:String(identity.date),
      evaluationType:'Competency',answers:payload.answers||{},
      uiState:payload.uiState||null,
      currentSection:Math.max(0,Math.min(8,Number(payload.currentSection)||0)),
      reviewPending:Boolean(payload.reviewPending),savedAt:new Date().toISOString()
    };
    const content=JSON.stringify(data);
    if(content.length>250000)throw new Error('Evaluation draft exceeds 250 KB; contact the coordinator');
    if(existing)existing.setContent(content);
    else folder.createFile(id+'.json',content,MimeType.PLAIN_TEXT);
    return {success:true,submissionId:id,savedAt:data.savedAt};
  }finally{lock.releaseLock();}
}
function ciDraftLoad_(identity){
  const file=ciDraftFind_(ciDraftFolder_(),identity.submissionId);
  if(!file)return {success:true,found:false};
  const data=ciDraftParsed_(file);ciDraftOwned_(data,identity);
  return {success:true,found:true,draft:data};
}
function ciDraftList_(evaluatorId){
  const folder=ciDraftFolder_(),files=folder.getFiles();
  const result=[];let scanned=0;
  while(files.hasNext()){
    if(++scanned>1000)throw new Error('Draft repository too large for unpaged listing');
    const file=files.next();
    if(!/^CI-PORTAL-[a-f0-9-]{36}\.json$/i.test(file.getName()))continue;
    let data;
    try{data=ciDraftParsed_(file);}catch(err){continue;}
    if(String(data.evaluatorId)!==String(evaluatorId))continue;
    result.push({submissionId:data.submissionId,studentId:data.studentId,
      siteId:data.siteId,examinationId:data.examinationId,date:data.date,
      currentSection:data.currentSection||0,savedAt:data.savedAt||'',
      reviewPending:Boolean(data.reviewPending)});
  }
  // Resolve display fields from existing authoritative reference tables, NOT
  // mutable names used as relational keys. Failure should not hide drafts.
  try{
    const book=ciDataBook_();
    const students=ciDataRows_(book,'Students',['Student ID','Student']);
    const exams=ciDataRows_(book,'Examinations',['Examination ID','Examination']);
    const studentMap=new Map(students.map(r=>[String(r['Student ID']),String(r.Student||'')]));
    const examMap=new Map(exams.map(r=>[String(r['Examination ID']),String(r.Examination||'')]));
    result.forEach(d=>{d.student=studentMap.get(d.studentId)||'';d.examination=examMap.get(d.examinationId)||'';});
  }catch(err){console.warn('Draft list display-name lookup unavailable: '+String(err.message||err));}
  result.sort((a,b)=>String(b.savedAt).localeCompare(String(a.savedAt)));
  return {success:true,drafts:result.slice(0,100),count:result.length};
}
