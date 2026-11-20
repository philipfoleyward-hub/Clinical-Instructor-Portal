/** Read-only direct CCMS workbook access. Preserves stable IDs and published scheduling rules. */
// One workbook handle and one read per sheet per request; small reference
// sheets are also shared across requests for 2 minutes (read-only data).
let ciDataBookMemo_=null;const ciDataRowsMemo_=Object.create(null);
const CI_DATA_CACHEABLE_=new Set(['Clinical Sites','Students','Evaluators','Examinations','Clinical Calendar','CI Assignments']);
function ciDataBook_(){return ciDataBookMemo_||(ciDataBookMemo_=SpreadsheetApp.openById(ciPortalWorkbookId_()));}
function ciDataRows_(book,name,required){
 const key='ciRows2:'+book.getId()+':'+name;
 let entry=ciDataRowsMemo_[key];
 if(!entry&&CI_DATA_CACHEABLE_.has(name)){
  try{const hit=CacheService.getScriptCache().get(key);
   if(hit)entry=JSON.parse(hit,(k,v)=>v&&typeof v==='object'&&v.$d!==undefined?new Date(v.$d):v);}catch(e){entry=null;}
 }
 if(!entry){
  const sh=book.getSheetByName(name);if(!sh)throw new Error('Missing CCMS sheet: '+name);
  const values=sh.getDataRange().getValues();if(!values.length)throw new Error('Empty CCMS sheet: '+name);
  const headers=values.shift().map(v=>String(v).trim());
  entry={headers,rows:values.filter(r=>r.some(v=>v!==''&&v!==null))};
  if(CI_DATA_CACHEABLE_.has(name)){
   try{const text=JSON.stringify(entry,function(k,v){return this[k] instanceof Date?{$d:this[k].getTime()}:v;});
    if(text.length<95000)CacheService.getScriptCache().put(key,text,120);}catch(e){}
  }
 }
 ciDataRowsMemo_[key]=entry;
 (required||[]).forEach(h=>{if(!entry.headers.includes(h))throw new Error('Missing '+name+' header: '+h);});
 return entry.rows.map(r=>Object.fromEntries(entry.headers.map((h,i)=>[h,r[i]])));
}
function ciDataNorm_(v){return String(v==null?'':v).trim().toLowerCase();}
function ciDataTrue_(v){return v===true||['true','yes','1'].includes(ciDataNorm_(v));}
function ciDataCohort_(v){const m=String(v||'').match(/\b20\d{2}\b/);return m?m[0]:ciDataNorm_(v);}
function ciDataDate_(v){
 if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
 const s=String(v||'').trim();return /^\d{4}-\d{2}-\d{2}$/.test(s)?s:'';
}
function ciDataWithin_(date,start,end){const a=ciDataDate_(start),b=ciDataDate_(end);return !!a&&!!b&&a<=date&&date<=b;}
function ciDataEvaluators_(){
 const rows=ciDataRows_(ciDataBook_(),'Evaluators',['Evaluator ID','Evaluator','Evaluator Type','Active','Portal Access']);
 return rows.filter(r=>ciDataTrue_(r.Active)&&ciDataTrue_(r['Portal Access']))
  .map(r=>({evaluatorId:String(r['Evaluator ID']||'').trim(),evaluatorName:String(r.Evaluator||'').trim(),evaluatorType:String(r['Evaluator Type']||'').trim()}))
  .filter(r=>r.evaluatorId&&r.evaluatorName).sort((a,b)=>a.evaluatorName.localeCompare(b.evaluatorName));
}
function ciDataHome_(evaluator,date){
 if(!/^\d{4}-\d{2}-\d{2}$/.test(date)||ciDataDate_(new Date(date+'T12:00:00'))!==date)throw new Error('Invalid schedule date');
 const book=ciDataBook_();
 const ci=ciDataRows_(book,'CI Assignments',['CI Assignment ID','Evaluator ID','Date','Site ID','Status'])
  .filter(r=>ciDataNorm_(r['Evaluator ID'])===ciDataNorm_(evaluator.evaluatorId)&&ciDataDate_(r.Date)===date&&ciDataNorm_(r.Status)==='current');
 if(!ci.length)return {success:true,evaluator,date,assigned:false,assignments:[]};
 const sites=ciDataRows_(book,'Clinical Sites',['Clinical Site ID']);
 const students=ciDataRows_(book,'Students',['Student ID','Student','Cohort','Active']);
 const calendar=ciDataRows_(book,'Clinical Calendar',['Date','Clinical Term']);
 const placements=ciDataRows_(book,'Student Site Assignments',['Student ID','Site ID','Start Date','End Date','Clinical Term']);
 const rotations=ciDataRows_(book,'Student Room Rotations',['Student ID','Assignment ID','Start Date','End Date','Rotation Area']);
 const unique=new Map();students.filter(s=>ciDataTrue_(s.Active)).forEach(s=>{const id=ciDataNorm_(s['Student ID']);if(id&&!unique.has(id))unique.set(id,s);});
 const scheduledBySite=new Map();
 for(const s of unique.values()){
  const id=String(s['Student ID']).trim(),cohort=String(s.Cohort||'').trim();
  const cal=calendar.filter(c=>(!('Active' in c)||ciDataTrue_(c.Active))&&ciDataDate_(c.Date)===date&&(!String(c.Cohort||'').trim()||ciDataCohort_(c.Cohort)===ciDataCohort_(cohort)));
  if(cal.length>1)throw new Error('Clinical Calendar conflict for student '+id);
  if(!cal.length)continue;
  const ps=placements.filter(p=>ciDataNorm_(p['Student ID'])===ciDataNorm_(id)&&!['cancelled','canceled','superseded','inactive'].includes(ciDataNorm_(p.Status))&&ciDataWithin_(date,p['Start Date'],p['End Date']));
  if(ps.length>1)throw new Error('Multiple active site assignments for student '+id);
  if(!ps.length)continue;
  const p=ps[0];if(ciDataNorm_(p['Clinical Term'])!==ciDataNorm_(cal[0]['Clinical Term']))throw new Error('Clinical Term mismatch for student '+id);
  const siteId=String(p['Site ID']||'').trim();
  if(!sites.some(x=>ciDataNorm_(x['Clinical Site ID'])===ciDataNorm_(siteId)))throw new Error('Unknown site ID '+siteId+' for student '+id);
  const rr=rotations.filter(r=>ciDataNorm_(r['Student ID'])===ciDataNorm_(id)&&ciDataNorm_(r['Assignment ID'])===ciDataNorm_(p['Assignment ID'])&&['','published'].includes(ciDataNorm_(r.Status))&&ciDataWithin_(date,r['Start Date'],r['End Date']));
  if(rr.length>1)throw new Error('Multiple room rotations for student '+id);
  const student={studentId:id,student:String(s.Student||'').trim(),cohort,clinicalTerm:String(cal[0]['Clinical Term']||'').trim(),siteId,rotationArea:rr.length?String(rr[0]['Rotation Area']||'').trim():'',scheduleStatus:rr.length?'Scheduled':'Rotation Not Assigned'};
  const key=ciDataNorm_(siteId);if(!scheduledBySite.has(key))scheduledBySite.set(key,[]);scheduledBySite.get(key).push(student);
 }
 const assignments=ci.map(a=>{
  const siteId=String(a['Site ID']||'').trim();const site=sites.find(s=>ciDataNorm_(s['Clinical Site ID'])===ciDataNorm_(siteId));
  if(!site)throw new Error('Unknown site ID '+siteId+' in CI Assignments');
  return {ciAssignmentId:String(a['CI Assignment ID']||'').trim(),evaluatorId:evaluator.evaluatorId,date,siteId,
   site:String(site['Clinical Site']||site.Site||site['Clinical Site Name']||'').trim(),
   startTime:ciDataTime_(a['Start Time']),endTime:ciDataTime_(a['End Time']),assignmentType:String(a['Assignment Type']||''),status:String(a.Status||''),notes:String(a.Notes||''),
   students:(scheduledBySite.get(ciDataNorm_(siteId))||[]).slice().sort((x,y)=>x.student.localeCompare(y.student))};
 }).sort((a,b)=>a.startTime.localeCompare(b.startTime));
 return {success:true,evaluator,date,assigned:true,assignments};
}
function ciDataTime_(v){return v instanceof Date?Utilities.formatDate(v,Session.getScriptTimeZone(),'HH:mm'):String(v==null?'':v).trim();}
