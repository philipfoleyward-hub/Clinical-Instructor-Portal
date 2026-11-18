/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * TESTS / DIAGNOSTICS
 * 99_Tests.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Central location for MANUAL testing and system diagnostics.
 *
 * THIS FILE MAY:
 *
 *   - Verify configuration
 *   - Verify sheets
 *   - Verify Forms
 *   - Verify templates
 *   - Test grading calculations
 *   - Test student lookups
 *   - Test term attribution
 *   - Test grade summaries
 *
 * THIS FILE SHOULD NOT:
 *
 *   - Contain production business logic
 *   - Be called by production workflows
 *   - Become a dependency of another script
 *
 * IMPORTANT
 * ---------
 * Tests are divided into:
 *
 *   SAFE / READ-ONLY
 *
 * and:
 *
 *   MUTATING TESTS
 *
 * Mutating tests may create:
 *
 *   - Drive folders
 *   - Google Docs
 *   - PDFs
 *   - Clinical Tracking records
 *   - Emails
 *
 * They are clearly labeled and should only be run deliberately.
 * =====================================================================
 */


/**
 * =====================================================================
 * TEST SETTINGS
 * =====================================================================
 *
 * Enter a real TEST student here when needed.
 *
 * Leave blank until we are ready to test student-specific functions.
 */
const TEST_CONFIG = {

  STUDENT_ID:
    '',

  STUDENT_NAME:
    '',

  CLINICAL_TERM:
    'Clinical 2',

  COHORT:
    '',

  TEST_EMAIL:
    ''

};


/**
 * =====================================================================
 * MASTER SAFE SYSTEM CHECK
 * =====================================================================
 *
 * READ ONLY.
 *
 * Runs the major non-destructive infrastructure checks.
 */
function testSystemHealth() {

  console.log(
    '=========================================='
  );

  console.log(
    'CLINICAL COMPETENCY TRACKER - SYSTEM TEST'
  );

  console.log(
    '=========================================='
  );


  const results = {


    config:
      testSystemConfiguration(),


    sheets:
      testMasterWorkbookStructure(),


    forms:
      testConfiguredForms(),


    templates:
      testConfiguredTemplates()

  };


  console.log(
    '=========================================='
  );

  console.log(
    'SYSTEM HEALTH TEST COMPLETE'
  );

  console.log(
    JSON.stringify(
      results,
      null,
      2
    )
  );

  console.log(
    '=========================================='
  );


  return results;
}


/**
 * =====================================================================
 * CONFIGURATION TEST
 * =====================================================================
 *
 * READ ONLY.
 */
function testSystemConfiguration() {

  if (
    typeof APP_CONFIG ===
    'undefined'
  ) {

    throw new Error(
      'APP_CONFIG was not found. Check 01_Config.gs.'
    );
  }


  const requiredSheets = [

    'STUDENTS',
    'CLINICAL_SITES',
    'EVALUATORS',
    'EXAMINATIONS',
    'SEMESTER_RULES',
    'PROGRAM_RULES',
    'SCORING_WEIGHTS',
    'IMAGE_SCORING_RULES',
    'CLINICAL_TRACKING',
    'STUDENT_REQUIREMENT_STATUS',
    'HISTORICAL_IMPORT'

  ];


  const missingSheets =
    requiredSheets.filter(key =>

      !APP_CONFIG.SHEETS[key]

    );


  if (
    missingSheets.length > 0
  ) {

    throw new Error(
      'APP_CONFIG.SHEETS is missing: ' +
      missingSheets.join(', ')
    );
  }


  if (
    !APP_CONFIG.FORMS
      .DIGITAL_CLINICAL_EVALUATION
  ) {

    throw new Error(
      'Digital Clinical Evaluation Form ID is missing.'
    );
  }


  if (
    !APP_CONFIG.FORMS
      .QR_SUBMISSION
  ) {

    throw new Error(
      'QR Submission Form ID is missing.'
    );
  }


  console.log(
    'PASS - Configuration structure is valid.'
  );


  return {

    success:
      true,

    message:
      'Configuration structure is valid.'

  };
}


/**
 * =====================================================================
 * MASTER WORKBOOK STRUCTURE
 * =====================================================================
 *
 * READ ONLY.
 *
 * Confirms that configured authoritative sheets exist.
 */
function testMasterWorkbookStructure() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const required = [

    APP_CONFIG.SHEETS.STUDENTS,
    APP_CONFIG.SHEETS.CLINICAL_SITES,
    APP_CONFIG.SHEETS.EVALUATORS,
    APP_CONFIG.SHEETS.EXAMINATIONS,
    APP_CONFIG.SHEETS.SEMESTER_RULES,
    APP_CONFIG.SHEETS.PROGRAM_RULES,
    APP_CONFIG.SHEETS.SCORING_WEIGHTS,
    APP_CONFIG.SHEETS.IMAGE_SCORING_RULES,
    APP_CONFIG.SHEETS.CLINICAL_TRACKING,
    APP_CONFIG.SHEETS.STUDENT_REQUIREMENT_STATUS

  ];


  const optional = [

    APP_CONFIG.SHEETS.ANATOMICAL_AREAS,
    APP_CONFIG.SHEETS.HISTORICAL_IMPORT,
    APP_CONFIG.SHEETS.FORM_RESPONSES_1,
    APP_CONFIG.SHEETS.FORM_RESPONSES_2

  ];


  const found = [];
  const missing = [];
  const optionalFound = [];
  const optionalMissing = [];


  required.forEach(
    sheetName => {

      if (
        ss.getSheetByName(
          sheetName
        )
      ) {

        found.push(
          sheetName
        );

      } else {

        missing.push(
          sheetName
        );
      }

    }
  );


  optional
    .filter(Boolean)
    .forEach(
      sheetName => {

        if (
          ss.getSheetByName(
            sheetName
          )
        ) {

          optionalFound.push(
            sheetName
          );

        } else {

          optionalMissing.push(
            sheetName
          );
        }

      }
    );


  console.log(
    'Required sheets found: ' +
    found.length
  );


  console.log(
    'Required sheets missing: ' +
    missing.length
  );


  if (
    missing.length > 0
  ) {

    console.log(
      'MISSING: ' +
      missing.join(', ')
    );


    throw new Error(
      'Required workbook sheet(s) missing: ' +
      missing.join(', ')
    );
  }


  return {

    success:
      true,

    requiredFound:
      found,

    optionalFound:
      optionalFound,

    optionalMissing:
      optionalMissing

  };
}


/**
 * =====================================================================
 * CONFIGURED GOOGLE FORMS
 * =====================================================================
 *
 * READ ONLY.
 *
 * Confirms both configured Form IDs are accessible.
 */
function testConfiguredForms() {

  const results = {};


  // -------------------------------------------------------------------
  // DIGITAL EVALUATION
  // -------------------------------------------------------------------

  try {

    const digitalForm =
      FormApp.openById(
        APP_CONFIG.FORMS
          .DIGITAL_CLINICAL_EVALUATION
      );


    results.digitalEvaluation = {

      success:
        true,

      name:
        digitalForm.getTitle(),

      responseCount:
        digitalForm
          .getResponses()
          .length

    };


  } catch (error) {

    results.digitalEvaluation = {

      success:
        false,

      error:
        error.message

    };
  }


  // -------------------------------------------------------------------
  // QR FORM
  // -------------------------------------------------------------------

  try {

    const qrForm =
      FormApp.openById(
        APP_CONFIG.FORMS
          .QR_SUBMISSION
      );


    results.qrSubmission = {

      success:
        true,

      name:
        qrForm.getTitle(),

      responseCount:
        qrForm
          .getResponses()
          .length

    };


  } catch (error) {

    results.qrSubmission = {

      success:
        false,

      error:
        error.message

    };
  }


  console.log(
    JSON.stringify(
      results,
      null,
      2
    )
  );


  return results;
}


/**
 * =====================================================================
 * CONFIGURED DOCUMENT TEMPLATES
 * =====================================================================
 *
 * READ ONLY.
 */
function testConfiguredTemplates() {

  const results = {};


  // -------------------------------------------------------------------
  // COMPETENCY EVALUATION TEMPLATE
  // -------------------------------------------------------------------

  results.competencyEvaluation =
    testTemplateId_(

      APP_CONFIG
        .TEMPLATES
        .CLINICAL_COMPETENCY_EVALUATION

    );


  // -------------------------------------------------------------------
  // SEMESTER REQUIREMENT TEMPLATES
  // -------------------------------------------------------------------

  results.semesterRequirements = {};


  const semesterTemplates =
    APP_CONFIG
      .TEMPLATES
      .SEMESTER_REQUIREMENTS ||
    {};


  Object
    .keys(
      semesterTemplates
    )
    .forEach(key => {

      const id =
        semesterTemplates[key];


      if (!id) {

        results
          .semesterRequirements[
            key
          ] = {

            configured:
              false

          };


        return;
      }


      results
        .semesterRequirements[
          key
        ] =
          testTemplateId_(
            id
          );

    });


  // -------------------------------------------------------------------
  // PROGRAM REQUIREMENTS TEMPLATE
  // -------------------------------------------------------------------

  const programTemplate =
    APP_CONFIG
      .TEMPLATES
      .PROGRAM_REQUIREMENTS;


  results.programRequirements =

    programTemplate

      ? testTemplateId_(
          programTemplate
        )

      : {

          configured:
            false

        };


  console.log(
    JSON.stringify(
      results,
      null,
      2
    )
  );


  return results;
}


/**
 * =====================================================================
 * TEMPLATE HELPER
 * =====================================================================
 */
function testTemplateId_(
  templateId
) {

  if (!templateId) {

    return {

      configured:
        false

    };
  }


  try {

    const file =
      DriveApp.getFileById(
        templateId
      );


    return {

      configured:
        true,

      success:
        true,

      name:
        file.getName(),

      mimeType:
        file.getMimeType(),

      nativeGoogleDoc:

        file.getMimeType() ===
        'application/vnd.google-apps.document'

    };


  } catch (error) {

    return {

      configured:
        true,

      success:
        false,

      error:
        error.message

    };
  }
}


/**
 * =====================================================================
 * SHARED UTILITIES
 * =====================================================================
 *
 * READ ONLY.
 */
function testSharedUtils() {

  const results = {


    normalize:
      appNormalize_(
        '  Clinical   2 '
      ),


    trueYes:
      appIsTrue_(
        'Yes'
      ),


    trueBoolean:
      appIsTrue_(
        true
      ),


    falseNo:
      appIsTrue_(
        'No'
      ),


    safeFile:
      appSafeFileText_(
        'Test: Student / File?'
      ),


    clinicalDate:
      appDateOnly_(
        new Date()
      )

  };


  console.log(
    JSON.stringify(
      results,
      null,
      2
    )
  );


  if (
    results.normalize !==
    'clinical 2'
  ) {

    throw new Error(
      'appNormalize_() failed.'
    );
  }


  if (
    results.trueYes !==
    true
  ) {

    throw new Error(
      'appIsTrue_() failed for Yes.'
    );
  }


  if (
    results.falseNo !==
    false
  ) {

    throw new Error(
      'appIsTrue_() failed for No.'
    );
  }


  console.log(
    'PASS - SharedUtils basic tests passed.'
  );


  return results;
}


/**
 * =====================================================================
 * DIGITAL EVALUATION GRADING
 * =====================================================================
 *
 * READ ONLY.
 *
 * Grades the newest Digital Clinical Evaluation Form response.
 *
 * DOES NOT:
 *
 *   - create documents
 *   - write tracking
 *   - send email
 */
function testDigitalEvaluationGrading() {

  const form =
    FormApp.openById(
      APP_CONFIG.FORMS
        .DIGITAL_CLINICAL_EVALUATION
    );


  const responses =
    form.getResponses();


  if (
    responses.length === 0
  ) {

    throw new Error(
      'Digital Clinical Evaluation Form has no responses.'
    );
  }


  const response =
    responses[
      responses.length - 1
    ];


  const result =
    gradeClinicalEvaluation(
      response
    );


  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;
}


/**
 * =====================================================================
 * CLINICAL TERM ATTRIBUTION TEST
 * =====================================================================
 *
 * READ ONLY.
 *
 * Enter a date below.
 */
function testClinicalTermAttribution() {

  const TEST_DATE =
    new Date();


  const clinicalTerm =
    clinicalTerm_getForDate(
      TEST_DATE
    );


  console.log(
    'Test Date: ' +
    TEST_DATE
  );


  console.log(
    'Attributed Clinical Term: ' +
    clinicalTerm
  );


  return {

    date:
      TEST_DATE,

    clinicalTerm:
      clinicalTerm

  };
}


/**
 * =====================================================================
 * STUDENT LOOKUP TEST
 * =====================================================================
 *
 * READ ONLY.
 */
function testStudentLookup() {

  const studentId =
    TEST_CONFIG.STUDENT_ID;


  if (!studentId) {

    throw new Error(
      'Set TEST_CONFIG.STUDENT_ID before running testStudentLookup().'
    );
  }


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const students =
    appSheetToObjects_(

      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENTS
      )

    );


  const matches =
    students.filter(row =>

      appNormalize_(
        row['Student ID']
      ) ===
      appNormalize_(
        studentId
      )

    );


  console.log(
    JSON.stringify(
      matches,
      null,
      2
    )
  );


  if (
    matches.length === 0
  ) {

    throw new Error(
      'No Students rows found for: ' +
      studentId
    );
  }


  return matches;
}


/**
 * =====================================================================
 * STUDENT GRADE SUMMARY TEST
 * =====================================================================
 *
 * READ ONLY.
 */
function testStudentGradeSummary() {

  const studentId =
    TEST_CONFIG.STUDENT_ID;


  if (!studentId) {

    throw new Error(
      'Set TEST_CONFIG.STUDENT_ID before running testStudentGradeSummary().'
    );
  }


  const result =
    gradeEngine_getStudentSummary(
      studentId
    );


  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;
}


/**
 * =====================================================================
 * REQUIREMENTS STATUS TEST
 * =====================================================================
 *
 * READ ONLY.
 *
 * Does NOT rebuild requirements.
 *
 * It simply reads the currently generated Student Requirement Status
 * rows for the configured test student.
 */
function testStudentRequirementStatus() {

  const studentId =
    TEST_CONFIG.STUDENT_ID;


  if (!studentId) {

    throw new Error(
      'Set TEST_CONFIG.STUDENT_ID before running testStudentRequirementStatus().'
    );
  }


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const rows =
    appSheetToObjects_(

      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS
          .STUDENT_REQUIREMENT_STATUS
      )

    );


  const studentRows =
    rows.filter(row =>

      appNormalize_(
        row['Student ID']
      ) ===
      appNormalize_(
        studentId
      )

    );


  console.log(
    'Requirement rows found: ' +
    studentRows.length
  );


  console.log(
    JSON.stringify(
      studentRows,
      null,
      2
    )
  );


  return studentRows;
}


/**
 * =====================================================================
 * CLINICAL TRACKING READ TEST
 * =====================================================================
 *
 * READ ONLY.
 */
function testClinicalTrackingForStudent() {

  const studentId =
    TEST_CONFIG.STUDENT_ID;


  if (!studentId) {

    throw new Error(
      'Set TEST_CONFIG.STUDENT_ID before running testClinicalTrackingForStudent().'
    );
  }


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const tracking =
    appSheetToObjects_(

      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS
          .CLINICAL_TRACKING
      )

    );


  const records =
    tracking.filter(row =>

      appNormalize_(
        row['Student ID']
      ) ===
      appNormalize_(
        studentId
      )

    );


  console.log(
    'Clinical Tracking records: ' +
    records.length
  );


  console.log(
    JSON.stringify(
      records,
      null,
      2
    )
  );


  return records;
}


/**
 * =====================================================================
 * DRIVE ROOT TEST
 * =====================================================================
 *
 * READ ONLY.
 *
 * Finds the Clinical Evaluations root.
 * Does not create student folders.
 */
function testDriveRoot() {

  const folder =
    driveService_getClinicalEvaluationsRoot_();


  const result = {

    id:
      folder.getId(),

    name:
      folder.getName(),

    url:
      folder.getUrl()

  };


  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;
}


/**
 * =====================================================================
 * ================================================================
 * MUTATING TESTS BELOW
 * ================================================================
 *
 * EVERYTHING BELOW THIS POINT MAY CREATE OR MODIFY REAL DATA.
 *
 * RUN DELIBERATELY.
 * =====================================================================
 */


/**
 * =====================================================================
 * MUTATING TEST
 * DRIVE STUDENT FOLDER
 * =====================================================================
 *
 * MAY CREATE:
 *
 * Clinical Evaluations
 *   └── Class of XXXX
 *       └── Student Name
 */
function MUTATING_testCreateStudentFolder() {

  const student =
    TEST_CONFIG.STUDENT_NAME;


  const cohort =
    TEST_CONFIG.COHORT;


  if (
    !student ||
    !cohort
  ) {

    throw new Error(
      'Set TEST_CONFIG.STUDENT_NAME and TEST_CONFIG.COHORT first.'
    );
  }


  const result =
    driveService_getStudentFolder({

      student:
        student,

      cohort:
        cohort

    });


  console.log(
    result.studentFolderUrl
  );


  return {

    folderId:
      result.studentFolderId,

    folderUrl:
      result.studentFolderUrl

  };
}


/**
 * =====================================================================
 * MUTATING TEST
 * GENERATE LATEST EVALUATION DOCUMENT
 * =====================================================================
 *
 * CREATES:
 *
 *   - Google Doc
 *   - PDF
 *
 * in the student's Evaluations folder.
 */
function MUTATING_testEvaluationDocument() {

  const form =
    FormApp.openById(
      APP_CONFIG.FORMS
        .DIGITAL_CLINICAL_EVALUATION
    );


  const responses =
    form.getResponses();


  if (
    responses.length === 0
  ) {

    throw new Error(
      'Digital Clinical Evaluation Form has no responses.'
    );
  }


  const response =
    responses[
      responses.length - 1
    ];


  const grade =
    gradeClinicalEvaluation(
      response
    );


  return evaluationDocument_generate(
    grade
  );
}


/**
 * =====================================================================
 * MUTATING TEST
 * GENERATE SEMESTER REQUIREMENT DOCUMENT
 * =====================================================================
 *
 * CREATES:
 *
 * a real Semester Requirement Google Doc.
 */
function MUTATING_testSemesterRequirementDocument() {

  const studentId =
    TEST_CONFIG.STUDENT_ID;


  const clinicalTerm =
    TEST_CONFIG.CLINICAL_TERM;


  if (!studentId) {

    throw new Error(
      'Set TEST_CONFIG.STUDENT_ID first.'
    );
  }


  return generateSemesterRequirementSheet(
    studentId,
    clinicalTerm
  );
}


/**
 * =====================================================================
 * MUTATING TEST
 * REBUILD ATTRIBUTED CLINICAL TERMS
 * =====================================================================
 *
 * MODIFIES:
 *
 * Clinical Tracking -> Attributed Clinical Term
 */
function MUTATING_testRebuildClinicalTermAttribution() {

  return rebuildAttributedClinicalTerms();
}


/**
 * =====================================================================
 * MUTATING TEST
 * REBUILD REQUIREMENTS
 * =====================================================================
 *
 * MODIFIES:
 *
 * Student Requirement Status
 */
function MUTATING_testRebuildRequirements() {

  return rebuildStudentRequirementStatus();
}


/**
 * =====================================================================
 * MUTATING TEST
 * SEND TEST NOTIFICATION
 * =====================================================================
 *
 * SENDS REAL EMAIL.
 *
 * Uses TEST_CONFIG.TEST_EMAIL rather than coordinator recipients.
 *
 * This intentionally does NOT use the production notification function
 * because we do not want a test email going to all coordinators.
 */
function MUTATING_testEmailSystem() {

  const email =
    String(
      TEST_CONFIG.TEST_EMAIL || ''
    ).trim();


  if (!email) {

    throw new Error(
      'Set TEST_CONFIG.TEST_EMAIL before running MUTATING_testEmailSystem().'
    );
  }


  MailApp.sendEmail({

    to:
      email,

    subject:
      'Clinical Evaluation System - Test Email',

    body:
      'This is a test of the Clinical Evaluation System notification service.',

    htmlBody:

      '<div style="' +
        'font-family:Arial,sans-serif;' +
        'font-size:14px;' +
      '">' +

        '<h2>' +
          'Clinical Evaluation System Test' +
        '</h2>' +

        '<p>' +
          'The notification system is able to send email successfully.' +
        '</p>' +

      '</div>',

    name:
      APP_CONFIG.NOTIFICATIONS
        .SENDER_NAME ||
      'Clinical Evaluation System'

  });


  console.log(
    'Test email sent to: ' +
    email
  );


  return {

    success:
      true,

    email:
      email

  };
}

/**
 * Controlled visual-QA generation for Isabella requirement documents.
 * Creates only Clinical 1, Clinical 2, and Program Requirements outputs.
 */
function qaGenerateIsabellaRequirementDocuments() {
  const studentId = 'DEPT # 22';

  return {
    clinical1: generateSemesterRequirementSheet(studentId, 'Clinical 1'),
    clinical2: generateSemesterRequirementSheet(studentId, 'Clinical 2'),
    program: generateProgramRequirementSheet(studentId)
  };
}


/**
 * Controlled visual-QA regeneration for Isabella C1 and C2 only.
 */
function qaGenerateIsabellaSemesterRequirements() {
  const studentId = 'DEPT # 22';

  return {
    clinical1: generateSemesterRequirementSheet(studentId, 'Clinical 1'),
    clinical2: generateSemesterRequirementSheet(studentId, 'Clinical 2')
  };
}


/**
 * Controlled visual-QA regeneration for Isabella C2 only.
 */
function qaGenerateIsabellaClinical2SemesterRequirement() {
  return generateSemesterRequirementSheet('DEPT # 22', 'Clinical 2');
}


/**
 * Controlled TEST-001 visual QA for remaining Semester Requirement families.
 * Creates documents only; does not write Clinical Tracking or Grade Records.
 */
function qaGenerateTest001Clinical345SemesterRequirements() {
  const studentId = 'TEST-001';
  return {
    clinical3: generateSemesterRequirementSheet(studentId, 'Clinical 3'),
    clinical4: generateSemesterRequirementSheet(studentId, 'Clinical 4'),
    clinical5: generateSemesterRequirementSheet(studentId, 'Clinical 5')
  };
}
