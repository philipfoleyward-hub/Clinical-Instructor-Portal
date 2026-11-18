/** CI PORTAL: live read models for site switch, any-active-student temporary selection,
 * program requirements, and own-evaluator history. Candidate, not production deployed.
 * No academic records or official assignments are written here.
 */
function ciHWAccess_(token,siteId,date) {
  const evaluator=ciR1Require_(token);
  const day=String(date||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'));
  if(!/^\d{4}-\d{2}-\d{2}$/.test(day)||ciDataDate_(new Date(day+'T12:00:00'))!==day) throw new Error('Invalid clinical date');
  const book=ciDataBook_();
  const sites=ciHWSiteOptionsFor_(evaluator,day,book);
  const site=sites.find(s=>ciDataNorm_(s.siteId)===ciDataNorm_(siteId));
  if(!site)throw new Error('Clinical site access unavailable.');
  return {evaluator:evaluator,book:book,site:site,date:day};
}
function ciHWSiteOptionsFor_(evaluator,date,book) {
  const reference=ciDataRows_(book,'Clinical Sites',['Clinical Site ID','Clinical Site','Active']);
  const all=reference.filter(r=>ciDataTrue_(r.Active)).map(r=>({siteId:String(r['Clinical Site ID']||'').trim(),site:String(r['Clinical Site']||'').trim()})).filter(s=>s.siteId&&s.site);
  if(ciDataNorm_(evaluator.evaluatorType)==='college ci')return all;
  // Existing named Hospital-CI accounts stay fail-closed until their distinct
  // self-entered-name/facility authentication policy has been approved.
  const dated=ciDataRows_(book,'CI Assignments',['Evaluator ID','Date','Site ID','Status'])
   .filter(r=>ciDataNorm_(r['Evaluator ID'])===ciDataNorm_(evaluator.evaluatorId)&&ciDataDate_(r.Date)===date&&ciDataNorm_(r.Status)==='current');
  const allowed=new Set(dated.map(r=>ciDataNorm_(r['Site ID'])));
  return all.filter(s=>allowed.has(ciDataNorm_(s.siteId)));
}
function ciHWAllActiveStudents_(book){
  const students=ciDataRows_(book,'Students',['Student ID','Student','Active','Cohort','Clinical Term']);
  const result=new Map();
  students.forEach(r=>{
    const id=String(r['Student ID']||'').trim(),key=ciDataNorm_(id);
    if(!id||!ciDataTrue_(r.Active)||result.has(key))return;
    result.set(key,{studentId:id,student:String(r.Student||'').trim(),cohort:String(r.Cohort||'').trim(),clinicalTerm:String(r['Clinical Term']||'').trim(),temporary:true,selectionSource:'Manual session selection'});
  });
  return [...result.values()].sort((a,b)=>a.student.localeCompare(b.student));
}
function ciHWGetStudent_(access,studentId){
  return ciHWAllActiveStudents_(access.book).find(s=>ciDataNorm_(s.studentId)===ciDataNorm_(studentId))||null;
}
function ciR1SiteOptions(token,date){
  const evaluator=ciR1Require_(token);
  const day=String(date||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'));
  const book=ciDataBook_();
  return {success:true,date:day,sites:ciHWSiteOptionsFor_(evaluator,day,book)};
}
function ciR1SiteRoster(token,siteId,date){
  const a=ciHWAccess_(token,siteId,date);
  const placements=ciDataRows_(a.book,'Student Site Assignments',['Student ID','Site ID','Start Date','End Date','Clinical Term']);
  const students=new Map(ciHWAllActiveStudents_(a.book).map(s=>[ciDataNorm_(s.studentId),s]));
  const rotations=ciDataRows_(a.book,'Student Room Rotations',['Student ID','Assignment ID','Start Date','End Date','Rotation Area']);
  const roster=new Map();
  placements.forEach(p=>{
    if(ciDataNorm_(p['Site ID'])!==ciDataNorm_(a.site.siteId)||!ciDataWithin_(a.date,p['Start Date'],p['End Date'])||['cancelled','canceled','superseded','inactive'].includes(ciDataNorm_(p.Status)))return;
    const id=ciDataNorm_(p['Student ID']),student=students.get(id);if(!student)return;
    if(roster.has(id))throw new Error('Conflicting current student placement for '+student.studentId);
    const matching=rotations.filter(r=>ciDataNorm_(r['Student ID'])===id&&ciDataNorm_(r['Assignment ID'])===ciDataNorm_(p['Assignment ID'])&&['','published'].includes(ciDataNorm_(r.Status))&&ciDataWithin_(a.date,r['Start Date'],r['End Date']));
    if(matching.length>1)throw new Error('Conflicting student rotations for '+student.studentId);
    roster.set(id,Object.assign({},student,{temporary:false,selectionSource:'Student Site Assignments',siteId:a.site.siteId,clinicalTerm:String(p['Clinical Term']||student.clinicalTerm),rotationArea:matching.length?String(matching[0]['Rotation Area']||''):'',scheduleStatus:matching.length?'Published room rotation':'Room rotation not assigned'}));
  });
  return {success:true,date:a.date,siteId:a.site.siteId,site:a.site.site,students:[...roster.values()].sort((x,y)=>x.student.localeCompare(y.student))};
}
function ciR1StudentSearch(token,siteId,date,query){
  const a=ciHWAccess_(token,siteId,date);
  const q=ciDataNorm_(query);
  const all=ciHWAllActiveStudents_(a.book).filter(s=>!q||ciDataNorm_(s.student).includes(q)||ciDataNorm_(s.studentId).includes(q));
  return {success:true,siteId:a.site.siteId,students:all.slice(0,100).map(s=>Object.assign({},s,{siteId:a.site.siteId})),total:all.length,scope:'All active CCMS students. Temporary selection does not change site placement.'};
}
function ciR1ValidateTemp(token,siteId,date,studentIds){
  const a=ciHWAccess_(token,siteId,date);
  const allowed=new Map(ciHWAllActiveStudents_(a.book).map(s=>[ciDataNorm_(s.studentId),s]));
  const ids=Array.isArray(studentIds)?studentIds.slice(0,100):[];
  return {success:true,students:[...new Set(ids.map(ciDataNorm_))].map(id=>allowed.get(id)).filter(Boolean).map(s=>Object.assign({},s,{siteId:a.site.siteId}))};
}
function ciHWRequirements_(access,studentId){
  const id=ciDataNorm_(studentId);
  const rows=ciDataRows_(access.book,'Student Requirement Status',['Student ID','Status Type','Requirement Scope','Requirement ID','Requirement Name','Requirement Timing','Remaining Count','Status']);
  const mine=rows.filter(r=>ciDataNorm_(r['Student ID'])===id);
  const details=mine.filter(r=>ciDataNorm_(r['Status Type'])==='detail'&&ciDataNorm_(r['Requirement Scope'])==='program')
    .filter(r=>Number(r['Remaining Count'])>0)
    .map(r=>({requirementId:String(r['Requirement ID']||''),name:String(r['Requirement Name']||''),type:String(r['Requirement Type']||''),group:String(r['Requirement Group']||''),timing:String(r['Requirement Timing']||''),remaining:Number(r['Remaining Count'])||0,status:String(r.Status||'')}));
  // Only two program-wide top-level rules are additive; group rules overlap.
  const top=mine.filter(r=>ciDataNorm_(r['Status Type'])==='summary'&&ciDataNorm_(r['Requirement Scope'])==='program'&&!String(r['Requirement Group']||'').trim()&&['PR-001','PR-002'].includes(String(r['Requirement ID'])));
  const ready=top.length===2&&top.every(r=>Number.isFinite(Number(r['Required Count']))&&Number.isFinite(Number(r['Completed Count'])));
  const progress=ready?{required:top.reduce((n,r)=>n+Number(r['Required Count']),0),completed:top.reduce((n,r)=>n+Number(r['Completed Count']),0),source:'Student Requirement Status, top-level program rules PR-001 and PR-002'}:null;
  return {opportunities:details,programProgress:progress,asOf:'Authoritative Student Requirement Status sheet (last rebuild time not established)'};
}
function ciR1StudentRequirements(token,siteId,date,studentId){
  const a=ciHWAccess_(token,siteId,date);
  const student=ciHWGetStudent_(a,studentId);if(!student)throw new Error('Active student not found');
  return Object.assign({success:true,studentId:student.studentId},ciHWRequirements_(a,student.studentId));
}
function ciR1StudentProfile(token,siteId,date,studentId){
  const access=ciHWAccess_(token,siteId,date),student=ciHWGetStudent_(access,studentId);
  if(!student)throw new Error('Active student not found');
  const sheet=access.book.getSheetByName('Clinical Tracking');
  let evaluations=[],historyNote='No evaluations attributed to this signed-in evaluator ID were found.';
  if(sheet){
    const values=sheet.getDataRange().getValues(),headers=(values[0]||[]).map(x=>String(x||'').trim());
    if(headers.includes('Student ID')&&headers.includes('Evaluator ID')){
      const i=Object.fromEntries(headers.map((h,j)=>[h,j]));
      evaluations=values.slice(1).filter(r=>ciDataNorm_(r[i['Student ID']])===ciDataNorm_(student.studentId)&&ciDataNorm_(r[i['Evaluator ID']])===ciDataNorm_(access.evaluator.evaluatorId)).map(r=>({
        submissionId:String(r[i['Submission ID']]||''),examination:String(r[i.Examination]||''),evaluationType:String(r[i['Evaluation Type']]||''),clinicalTerm:String(r[i['Clinical Term']]||''),site:String(r[i['Clinical Site']]||''),evaluator:String(r[i.Evaluator]||''),date:ciDataDate_(r[i.Date])||String(r[i.Date]||''),result:String(r[i.Result]||''),reviewStatus:String(r[i['Coordinator Review Status']]||''),revisionAvailable:false
      })).sort((a,b)=>String(b.date).localeCompare(String(a.date))).slice(0,100);
    } else historyNote='Historical evaluator-ID attribution unavailable; name-based lookup cannot authorize history.';
  }else historyNote='Clinical Tracking sheet unavailable.';
  return Object.assign({success:true,student:student,evaluations:evaluations,historyNote:historyNote,revisionStatus:'Read-only until source-specific audited revision writer is established.'},ciHWRequirements_(access,student.studentId));
}
