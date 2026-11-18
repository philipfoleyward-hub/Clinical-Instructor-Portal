/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * DRIVE SERVICE
 * 05_DriveService.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Single authoritative service for Google Drive filing.
 *
 * THIS FILE OWNS:
 *
 *   - Clinical Evaluations root folder lookup
 *   - Cohort folder creation / lookup
 *   - Student folder creation / lookup
 *   - Student subfolder creation / lookup
 *   - Moving uploaded evaluation files
 *   - Creating standardized file locations
 *   - Standardized Drive-safe names
 *
 * THIS FILE DOES NOT:
 *
 *   - Grade evaluations
 *   - Write Clinical Tracking
 *   - Calculate requirements
 *   - Generate document CONTENT
 *   - Send emails
 *
 * IMPORTANT
 * ---------
 * Other scripts should not independently recreate student-folder logic.
 *
 * =====================================================================
 */


/**
 * =====================================================================
 * GET / CREATE STUDENT FOLDER
 * =====================================================================
 *
 * Intended structure:
 *
 * Clinical Evaluations
 *     └── Class of 2027
 *          └── Isabella Steinberg
 *
 * Input:
 *
 * {
 *   student: 'Isabella Steinberg',
 *   cohort: '2027'
 * }
 *
 * Returns:
 *
 * {
 *   rootFolder,
 *   cohortFolder,
 *   studentFolder,
 *   studentFolderId,
 *   studentFolderUrl
 * }
 */
function driveService_getStudentFolder(
  options
) {

  options =
    options || {};


  const studentName =
    driveService_clean_(
      options.student
    );


  const cohort =
    driveService_clean_(
      options.cohort
    );


  if (!studentName) {

    throw new Error(
      'DriveService requires a student name.'
    );
  }


  if (!cohort) {

    throw new Error(
      'DriveService requires a cohort.'
    );
  }


  // ===================================================================
  // ROOT
  // ===================================================================

  const rootFolder =
    driveService_getClinicalEvaluationsRoot_();


  // ===================================================================
  // COHORT
  // ===================================================================

  const cohortFolderName =
    driveService_cohortFolderName_(
      cohort
    );


  const cohortFolder =
    driveService_getOrCreateChildFolder_(
      rootFolder,
      cohortFolderName
    );


  // ===================================================================
  // STUDENT
  // ===================================================================

  const safeStudentName =
    appSafeFileText_(
      studentName
    );


  const studentFolder =
    driveService_getOrCreateChildFolder_(
      cohortFolder,
      safeStudentName
    );


  return {

    rootFolder:
      rootFolder,

    cohortFolder:
      cohortFolder,

    studentFolder:
      studentFolder,

    studentFolderId:
      studentFolder.getId(),

    studentFolderUrl:
      studentFolder.getUrl()

  };
}


/**
 * =====================================================================
 * GET / CREATE STUDENT SUBFOLDER
 * =====================================================================
 *
 * Example:
 *
 * Clinical Evaluations
 *   └── Class of 2027
 *        └── Isabella Steinberg
 *             └── Evaluations
 *
 * or:
 *
 *             └── Uploaded Evaluations
 *
 * or:
 *
 *             └── Generated Documents
 */
function driveService_getStudentSubfolder(
  options
) {

  options =
    options || {};


  const subfolderName =
    driveService_clean_(
      options.subfolder
    );


  if (!subfolderName) {

    throw new Error(
      'DriveService requires a subfolder name.'
    );
  }


  const studentFolders =
    driveService_getStudentFolder(
      options
    );


  const subfolder =
    driveService_getOrCreateChildFolder_(
      studentFolders.studentFolder,
      subfolderName
    );


  return {

    rootFolder:
      studentFolders.rootFolder,

    cohortFolder:
      studentFolders.cohortFolder,

    studentFolder:
      studentFolders.studentFolder,

    subfolder:
      subfolder,

    folderId:
      subfolder.getId(),

    folderUrl:
      subfolder.getUrl()

  };
}


/**
 * =====================================================================
 * GET STANDARD EVALUATION FOLDER
 * =====================================================================
 *
 * Digital evaluations and uploaded evaluations should ultimately live
 * under the student's record rather than in unrelated output folders.
 */
function driveService_getEvaluationFolder(
  options
) {

  return driveService_getStudentSubfolder({

    student:
      options.student,

    cohort:
      options.cohort,

    subfolder:
      'Evaluations'

  });
}


/**
 * =====================================================================
 * GET GENERATED DOCUMENTS FOLDER
 * =====================================================================
 *
 * Semester Requirement sheets, Program Requirement sheets, etc.
 */
function driveService_getGeneratedDocumentsFolder(
  options
) {

  return driveService_getStudentSubfolder({

    student:
      options.student,

    cohort:
      options.cohort,

    subfolder:
      'Generated Documents'

  });
}


/**
 * =====================================================================
 * MOVE EXISTING FILE INTO STUDENT EVALUATIONS
 * =====================================================================
 *
 * Used mainly for QR/upload submissions.
 *
 * Input:
 *
 * {
 *   fileId,
 *   student,
 *   cohort,
 *   newFileName
 * }
 */
function driveService_fileUploadedEvaluation(
  options
) {

  options =
    options || {};


  const fileId =
    driveService_clean_(
      options.fileId
    );


  if (!fileId) {

    throw new Error(
      'DriveService requires a file ID.'
    );
  }


  const folders =
    driveService_getEvaluationFolder(
      options
    );


  const file =
    DriveApp.getFileById(
      fileId
    );


  if (
    options.newFileName
  ) {

    file.setName(
      appSafeFileText_(
        options.newFileName
      )
    );
  }


  /*
   * moveTo() makes the student's Evaluations folder
   * the permanent location.
   */
  file.moveTo(
    folders.subfolder
  );


  return {

    success:
      true,

    fileId:
      file.getId(),

    fileName:
      file.getName(),

    fileUrl:
      file.getUrl(),

    folderId:
      folders.folderId,

    folderUrl:
      folders.folderUrl

  };
}


/**
 * =====================================================================
 * COPY TEMPLATE TO STUDENT FOLDER
 * =====================================================================
 *
 * Useful for:
 *
 *   - Digital evaluation document generator
 *   - Semester Requirement generator
 *   - Program Requirement generator
 *
 * This service only COPIES the file.
 *
 * The document generator remains responsible for populating its content.
 */
function driveService_copyTemplateToStudentFolder(
  options
) {

  options =
    options || {};


  const templateId =
    driveService_clean_(
      options.templateId
    );


  const fileName =
    appSafeFileText_(
      options.fileName ||
      'Generated Document'
    );


  if (!templateId) {

    throw new Error(
      'DriveService requires a template ID.'
    );
  }


  let destination;


  if (
    options.destination ===
    'evaluations'
  ) {

    destination =
      driveService_getEvaluationFolder(
        options
      );

  } else {

    destination =
      driveService_getGeneratedDocumentsFolder(
        options
      );
  }


  const template =
    DriveApp.getFileById(
      templateId
    );


  const copy =
    template.makeCopy(
      fileName,
      destination.subfolder
    );


  return {

    success:
      true,

    file:
      copy,

    fileId:
      copy.getId(),

    fileName:
      copy.getName(),

    fileUrl:
      copy.getUrl(),

    folderId:
      destination.folderId,

    folderUrl:
      destination.folderUrl

  };
}


/**
 * =====================================================================
 * SAVE PDF BLOB TO STUDENT FOLDER
 * =====================================================================
 *
 * Input:
 *
 * {
 *   blob,
 *   student,
 *   cohort,
 *   fileName,
 *   destination: 'evaluations' | 'documents'
 * }
 */
function driveService_saveBlob(
  options
) {

  options =
    options || {};


  if (!options.blob) {

    throw new Error(
      'DriveService requires a Blob.'
    );
  }


  const fileName =
    appSafeFileText_(
      options.fileName ||
      'Generated File'
    );


  let destination;


  if (
    options.destination ===
    'evaluations'
  ) {

    destination =
      driveService_getEvaluationFolder(
        options
      );

  } else {

    destination =
      driveService_getGeneratedDocumentsFolder(
        options
      );
  }


  const blob =
    options.blob.copyBlob();


  blob.setName(
    fileName
  );


  const file =
    destination
      .subfolder
      .createFile(
        blob
      );


  return {

    success:
      true,

    fileId:
      file.getId(),

    fileName:
      file.getName(),

    fileUrl:
      file.getUrl(),

    folderId:
      destination.folderId,

    folderUrl:
      destination.folderUrl

  };
}


/**
 * =====================================================================
 * BUILD STANDARD EVALUATION FILE NAME
 * =====================================================================
 *
 * Example:
 *
 * 2026-07-16 - Isabella Steinberg - Orbits - Competency
 */
function driveService_buildEvaluationFileName(
  options
) {

  options =
    options || {};


  const date =
    appDateOnly_(
      options.date
    );


  const dateText =
    date

      ? Utilities.formatDate(
          date,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        )

      : Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        );


  const student =
    appSafeFileText_(
      options.student ||
      'Unknown Student'
    );


  const examination =
    appSafeFileText_(
      options.examination ||
      'Examination'
    );


  const evaluationType =
    appSafeFileText_(
      options.evaluationType ||
      'Evaluation'
    );


  return (
    dateText +
    ' - ' +
    student +
    ' - ' +
    examination +
    ' - ' +
    evaluationType
  );
}


/**
 * =====================================================================
 * GET CLINICAL EVALUATIONS ROOT
 * =====================================================================
 *
 * Preferred:
 *
 * APP_CONFIG.DRIVE.CLINICAL_EVALUATIONS_ROOT_ID
 *
 * Transitional fallback:
 *
 * unique folder named "Clinical Evaluations"
 *
 * Once we identify the permanent folder ID, the name-based fallback
 * can eventually be removed.
 */
function driveService_getClinicalEvaluationsRoot_() {

  const configuredId =
    driveService_clean_(
      APP_CONFIG.DRIVE
        .CLINICAL_EVALUATIONS_ROOT_ID
    );


  if (configuredId) {

    try {

      return DriveApp
        .getFolderById(
          configuredId
        );

    } catch (error) {

      throw new Error(
        'Configured Clinical Evaluations root folder ID is invalid or inaccessible: ' +
        configuredId
      );
    }
  }


  const folderName =
    APP_CONFIG.DRIVE
      .CLINICAL_EVALUATIONS_ROOT_NAME ||
    'Clinical Evaluations';


  const folders =
    DriveApp.getFoldersByName(
      folderName
    );


  if (
    !folders.hasNext()
  ) {

    throw new Error(
      'Could not find Drive folder "' +
      folderName +
      '". Add its folder ID to APP_CONFIG.DRIVE.CLINICAL_EVALUATIONS_ROOT_ID.'
    );
  }


  const folder =
    folders.next();


  if (
    folders.hasNext()
  ) {

    throw new Error(
      'More than one Drive folder is named "' +
      folderName +
      '". Add the correct folder ID to APP_CONFIG so DriveService does not have to guess.'
    );
  }


  return folder;
}


/**
 * =====================================================================
 * GET OR CREATE DIRECT CHILD FOLDER
 * =====================================================================
 */
function driveService_getOrCreateChildFolder_(
  parentFolder,
  childName
) {

  const safeName =
    driveService_clean_(
      childName
    );


  if (!safeName) {

    throw new Error(
      'Cannot create a Drive folder without a name.'
    );
  }


  const folders =
    parentFolder
      .getFoldersByName(
        safeName
      );


  if (
    folders.hasNext()
  ) {

    const folder =
      folders.next();


    /*
     * Multiple folders with the same name under the same parent
     * indicate a filing problem.
     *
     * Do not silently choose one.
     */
    if (
      folders.hasNext()
    ) {

      throw new Error(
        'More than one folder named "' +
        safeName +
        '" exists inside "' +
        parentFolder.getName() +
        '".'
      );
    }


    return folder;
  }


  if (
    APP_CONFIG.SETTINGS
      .CREATE_MISSING_STUDENT_FOLDERS !==
      true
  ) {

    throw new Error(
      'Drive folder does not exist: "' +
      safeName +
      '".'
    );
  }


  console.log(
    'Creating Drive folder: ' +
    safeName
  );


  return parentFolder
    .createFolder(
      safeName
    );
}


/**
 * =====================================================================
 * COHORT → FOLDER NAME
 * =====================================================================
 *
 * Accepts:
 *
 *   2027
 *   Class of 2027
 *   Class Of 2027
 *
 * Returns:
 *
 *   Class of 2027
 */
function driveService_cohortFolderName_(
  cohort
) {

  const text =
    driveService_clean_(
      cohort
    );


  const match =
    text.match(
      /\d{4}/
    );


  if (!match) {

    throw new Error(
      'Could not determine cohort year from: ' +
      text
    );
  }


  return (
    'Class of ' +
    match[0]
  );
}


/**
 * =====================================================================
 * CLEAN
 * =====================================================================
 */
function driveService_clean_(
  value
) {

  return String(
    value === null ||
    value === undefined
      ? ''
      : value
  ).trim();
}


/**
 * =====================================================================
 * TEST STUDENT FOLDER RESOLUTION
 * =====================================================================
 *
 * SAFE-ish TEST:
 *
 * This may CREATE missing cohort/student folders because that is the
 * intended production behavior.
 *
 * Fill in a real TEST student/cohort before running.
 */
function testDriveServiceStudentFolder() {

  const TEST_STUDENT =
    '';

  const TEST_COHORT =
    '';


  if (
    !TEST_STUDENT ||
    !TEST_COHORT
  ) {

    throw new Error(
      'Enter a test student and cohort before running testDriveServiceStudentFolder().'
    );
  }


  const result =
    driveService_getStudentFolder({

      student:
        TEST_STUDENT,

      cohort:
        TEST_COHORT

    });


  console.log(
    'Student folder: ' +
    result.studentFolderUrl
  );


  return {

    studentFolderId:
      result.studentFolderId,

    studentFolderUrl:
      result.studentFolderUrl

  };
}
/**
 * Approved CCMS DriveService remains the sole Drive storage owner.
 * CI drafts are private text JSON under an EXPLICITLY CONFIGURED,
 * separate draft folder. It is never the final Evaluations output folder.
 * The user has not approved a specific draft folder ID: this guard is
 * intentional; configure APP_CONFIG.DRIVE.CI_PORTAL_DRAFTS_FOLDER_ID.
 */
function driveService_ciDraftFolder_() {
  const id=String(APP_CONFIG.DRIVE.CI_PORTAL_DRAFTS_FOLDER_ID||'').trim();
  if(!/^[A-Za-z0-9_-]{20,}$/.test(id))throw new Error('CI Portal drafts folder not approved/configured in 01_Config.gs');
  return DriveApp.getFolderById(id);
}
function driveService_ciDraftFile_(submissionId) {
  const id=String(submissionId||'');
  if(!/^CI-PORTAL-[a-f0-9-]{36}$/i.test(id))throw new Error('Invalid CI Portal draft ID');
  const iterator=driveService_ciDraftFolder_().getFilesByName(id+'.json');
  if(!iterator.hasNext())return null;
  const file=iterator.next();
  if(iterator.hasNext())throw new Error('Duplicate CI Portal draft files; manual reconciliation required');
  return file;
}
function driveService_ciDraftSave_(identity,payload) {
  const folder=driveService_ciDraftFolder_(),sid=String(payload.submissionId||'');
  const prior=driveService_ciDraftFile_(sid);
  if(prior){
    const former=JSON.parse(prior.getBlob().getDataAsString('UTF-8'));
    if(former.evaluatorId!==identity.evaluatorId||former.studentId!==identity.studentId)throw new Error('Draft ownership mismatch');
  }
  const data={
    version:1,submissionId:sid,studentId:identity.studentId,siteId:identity.siteId,
    evaluatorId:identity.evaluatorId,examinationId:identity.examinationId,
    date:String(payload.date),evaluationType:'Competency',
    answers:payload.answers||{},uiState:payload.uiState||null,currentSection:Math.max(0,Math.min(8,Number(payload.currentSection)||0)),
    reviewPending:Boolean(payload.reviewPending),savedAt:new Date().toISOString()
  };
  const contents=JSON.stringify(data);
  if(contents.length>250000)throw new Error('Evaluation draft exceeds storage limit');
  const file=prior||folder.createFile(sid+'.json',contents,MimeType.PLAIN_TEXT);
  if(prior)file.setContent(contents);
  return {success:true,submissionId:sid,savedAt:data.savedAt};
}
function driveService_ciDraftLoad_(identity,sid) {
  const file=driveService_ciDraftFile_(sid);
  if(!file)return {success:true,found:false};
  const data=JSON.parse(file.getBlob().getDataAsString('UTF-8'));
  if(data.studentId!==identity.studentId||data.evaluatorId!==identity.evaluatorId)throw new Error('Draft access denied');
  return {success:true,found:true,draft:data};
}
function driveService_ciDraftDelete_(identity,sid) {
  const file=driveService_ciDraftFile_(sid);
  if(!file)return {success:true,deleted:false};
  const data=JSON.parse(file.getBlob().getDataAsString('UTF-8'));
  if(data.studentId!==identity.studentId||data.evaluatorId!==identity.evaluatorId)throw new Error('Draft access denied');
  file.setTrashed(true);
  return {success:true,deleted:true};
}
/** Current evaluator's unfinished evaluation drafts only; exact ID retained. */
function driveService_ciDraftList_(evaluatorId) {
  const folder=driveService_ciDraftFolder_();
  const files=folder.getFiles();const result=[];let scanned=0;
  const sheet=appGetSheet_(SpreadsheetApp.getActiveSpreadsheet(),APP_CONFIG.SHEETS.CLINICAL_TRACKING);
  while(files.hasNext()){
    if(++scanned>500)throw new Error('Too many drafts for unpaged listing; administrator reconciliation needed');
    const f=files.next();
    if(!/^CI-PORTAL-[a-f0-9-]{36}\.json$/i.test(f.getName()))continue;
    let data;try{data=JSON.parse(f.getBlob().getDataAsString('UTF-8'))}catch(e){continue}
    if(data.evaluatorId!==evaluatorId)continue;
    if(clinicalTracking_submissionExists_(sheet,data.submissionId))continue;
    result.push({submissionId:data.submissionId,studentId:data.studentId,siteId:data.siteId,
      examinationId:data.examinationId,date:data.date,currentSection:data.currentSection||0,
      savedAt:data.savedAt||'',reviewPending:!!data.reviewPending});
  }
  return result.sort((a,b)=>String(b.savedAt).localeCompare(String(a.savedAt))).slice(0,100);
}
