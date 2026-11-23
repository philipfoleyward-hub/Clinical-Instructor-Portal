/**
 * 06_ProfessionalCompetencyDocument.gs — CANDIDATE, no production changes.
 * Copy the user's two verified tokenized Google Sheets templates, fill by Form item ID,
 * export only the report tab as PDF, and file both in the established student folder.
 * No student/clinical entity folders are created and no names are used as relational IDs.
 */
function pcV2_existingEvaluationFolder_(student) {
  if(typeof dashboardDocuments_findExistingStudentFolder_!=='function' ||
     typeof dashboardDocuments_getStudent_!=='function')
    throw new Error('Authoritative existing student folder resolver is missing.');
  // The resolver takes a record grounded by Student ID; if its legacy name-based
  // lookup resolves ambiguously it throws rather than making another folder.
  const canonical=dashboardDocuments_getStudent_(student.studentId);
  if(pcV2_text_(canonical.studentId)!==student.studentId)
    throw new Error('Student folder resolver did not retain the stable Student ID.');
  const studentFolder=dashboardDocuments_findExistingStudentFolder_(canonical);
  if(!studentFolder)throw new Error('Existing student folder unavailable. No folder created.');
  const found=studentFolder.getFoldersByName('Evaluations');
  let folder=null;
  while(found.hasNext()){
    if(folder)throw new Error('Duplicate student Evaluations folders; filing refused.');
    folder=found.next();
  }
  if(!folder)throw new Error('Existing student Evaluations subfolder not found. No folder created.');
  return {folder:folder,folderId:folder.getId(),url:folder.getUrl()};
}
function pcV2_validateTemplate_(id) {
  if(!id)throw new Error('Missing approved Professional Competency template ID.');
  const file=DriveApp.getFileById(id);
  if(file.getMimeType()!==MimeType.GOOGLE_SHEETS)throw new Error('Approved template is not a native Google Sheet.');
  const sheet=SpreadsheetApp.openById(id).getSheetByName('Sheet1');
  if(!sheet)throw new Error('Approved Professional Competency template missing Sheet1.');
}
function pcV2_safeName_(s) {
  return pcV2_text_(s).replace(/[\\/<>:"|?*\x00-\x1f]/g,'-').replace(/\s+/g,' ').slice(0,180);
}
function pcV2_documentName_(input) {
  const time=Utilities.formatDate(input.date,Session.getScriptTimeZone(),'yyyy-MM-dd');
  return pcV2_safeName_('PC '+input.student.studentId+' '+input.clinicalTerm+' '+time+' '+input.sourceId);
}
function pcV2_findExactFile_(folder,name,mimeType) {
  const it=folder.getFilesByName(name);
  let file=null;
  while(it.hasNext()){
    const candidate=it.next();
    if(candidate.getMimeType()!==mimeType)continue;
    if(file)throw new Error('Duplicate generated files named '+name+'; manual reconciliation required.');
    file=candidate;
  }
  return file;
}
function pcV2_verifyParent_(file,folderId) {
  const parents=file.getParents();
  let valid=false;
  while(parents.hasNext())if(parents.next().getId()===folderId)valid=true;
  if(!valid)throw new Error('Generated file is not within the authorized student evaluation folder.');
}
function pcV2_buildDocument_(input,priorId) {
  // Approved file name: "<Student> - Professional Competency Evaluation - MM-dd-yyyy".
  const fileName=pcV2_safeName_([
    input.student.student,
    'Professional Competency Evaluation',
    Utilities.formatDate(input.date,Session.getScriptTimeZone(),'MM-dd-yyyy')
  ].join(' - '));
  // Retry-safe: reuse ONLY the copy this submission already checkpointed.
  // Never reuse by name: two evaluations on one day share the same name.
  const file=priorId
    ? DriveApp.getFileById(priorId)
    : DriveApp.getFileById(PC_V2_TEMPLATES[input.score.family]).makeCopy(fileName,input.folder.folder);
  pcV2_verifyParent_(file,input.folder.folderId);
  const book=SpreadsheetApp.openById(file.getId());
  const sheet=book.getSheetByName('Sheet1');
  if(!sheet)throw new Error('Copied Professional Competency sheet is missing Sheet1.');
  const rows=input.score.family==='Clinical 1-2'
    ? input.score.details.map(function(_,i){return i<9?13+i:15+i;})
    : input.score.details.map(function(_,i){return i<9?13+i:i<20?15+i:16+i;});
  if(rows.length!==input.score.criterionCount)throw new Error('Template/rubric row count mismatch.');
  const token=sheet.getRange(rows[0],7).getDisplayValue();
  if(token.indexOf('{{PC_')!==0 && token!==input.score.details[0].rating)
    throw new Error('Copied template does not contain expected stable Form-ID rating token.');

  const head={
    A8:'Student: '+input.student.student,F8:'Evaluator: '+input.evaluator,
    A9:'Student ID: '+input.student.studentId,F9:'Clinical Site: '+input.site.name,
    A10:'Date: '+Utilities.formatDate(input.date,Session.getScriptTimeZone(),'M/d/yyyy'),
    D10:'Period: '+input.period,G10:'Term: '+input.clinicalTerm,
    A11:'Submission ID: '+input.sourceId
  };
  Object.keys(head).forEach(function(a1){sheet.getRange(a1).setValue(head[a1]);});
  input.score.details.forEach(function(item,i){
    const r=rows[i];
    // Titles and help-text weights come from the verified live Form snapshot.
    sheet.getRange(r,1).setValue((i+1)+'.  '+item.text);
    sheet.getRange(r,7,1,2).setValues([[item.rating==='N/A'?'N/A':Number(item.rating),item.weight]]);
  });
  if(input.score.family==='Clinical 1-2'){
    sheet.getRange('A41').setValue(input.strengths);
    sheet.getRange('A44').setValue(input.improvements);
  }else{
    sheet.getRange('A49').setValue(input.strengths);
    sheet.getRange('A52').setValue(input.improvements);
  }
  SpreadsheetApp.flush();
  const totalRow=input.score.family==='Clinical 1-2'?34:42;
  const gradeRow=input.score.family==='Clinical 1-2'?36:44;
  const output=Number(sheet.getRange(gradeRow,7).getValue())*100;
  const applied=Number(sheet.getRange(totalRow,10).getValue());
  if(!isFinite(output)||Math.abs(output-input.score.percentGrade)>0.011 ||
     Math.abs(applied-input.score.applicableWeight)>0.001){
    throw new Error('Generated template formula disagrees with authoritative calculated grade. Filing halted.');
  }
  return {docId:file.getId(),url:file.getUrl(),name:fileName,
    sheetGid:sheet.getSheetId(),folderId:input.folder.folderId};
}
function pcV2_buildPdf_(input,document,priorId) {
  const name=document.name+'.pdf';
  const existing=priorId?DriveApp.getFileById(priorId):null;
  if(existing){pcV2_verifyParent_(existing,input.folder.folderId);
    return {pdfId:existing.getId(),url:existing.getUrl()};}
  const exportUrl='https://docs.google.com/spreadsheets/d/'+encodeURIComponent(document.docId)+
    '/export?format=pdf&gid='+encodeURIComponent(document.sheetGid)+
    '&size=letter&portrait=true&fitw=true&sheetnames=false&printtitle=false&gridlines=false&fzr=false';
  const res=UrlFetchApp.fetch(exportUrl,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},
    muteHttpExceptions:true});
  if(res.getResponseCode()!==200 || !/pdf/i.test(res.getHeaders()['Content-Type']||res.getHeaders()['content-type']||''))
    throw new Error('Professional Competency PDF export failed with status '+res.getResponseCode()+'. Grade record not created.');
  const blob=res.getBlob().setName(name);
  const file=input.folder.folder.createFile(blob);
  pcV2_verifyParent_(file,input.folder.folderId);
  return {pdfId:file.getId(),url:file.getUrl()};
}
