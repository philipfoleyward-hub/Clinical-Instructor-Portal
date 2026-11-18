/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * QR / UPLOADED EVALUATION INTAKE
 * 03_QRSubmissionIntake.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Processes submitted QR / uploaded clinical evaluations.
 *
 * FLOW
 * ----
 * QR Form
 *   -> read response
 *   -> resolve student
 *   -> determine attributed clinical term
 *   -> move uploaded file
 *   -> write Clinical Tracking
 *   -> rebuild requirements
 *   -> send notifications
 *
 * IMPORTANT
 * ---------
 * QR submissions are already completed evaluations.
 *
 * This script DOES NOT:
 *   - re-grade the evaluation
 *   - generate a new evaluation document
 *
 * =====================================================================
 */


/**
 * =====================================================================
 * PRODUCTION FORM SUBMIT HANDLER
 * =====================================================================
 */
function qrSubmission_onFormSubmit(e) {

  if (!e || !e.response) {
    throw new Error(
      'QR Submission Intake requires a Google Forms submit event.'
    );
  }


  const response =
    e.response;


  const responseId =
    String(
      response.getId() || ''
    ).trim();


  if (!responseId) {
    throw new Error(
      'QR submission has no Google Forms response ID.'
    );
  }


  const submissionId =
    'QR-' + responseId;


  console.log(
    'QR intake started: ' +
    submissionId
  );


  // ===================================================================
  // READ ANSWERS
  // ===================================================================

  const answers =
    qrSubmission_getAnswers_(response);


  const student =
    qrSubmission_text_(
      answers['student']
    );


  const clinicalSite =
    qrSubmission_text_(
      answers['clinical site']
    );


  const examination =
    qrSubmission_text_(
      answers['examination']
    );


  const evaluationType =
    qrSubmission_text_(
      answers['evaluation type']
    );


  const evaluationDate =
    qrSubmission_parseDate_(
      answers['evaluation date']
    );


  const result =
    qrSubmission_parseResult_(
      answers['results']
    );


  const finalScore =
    qrSubmission_parseScore_(
      answers['final score']
    );


  if (!student) {
    throw new Error(
      'QR submission is missing Student.'
    );
  }


  if (!clinicalSite) {
    throw new Error(
      'QR submission is missing Clinical Site.'
    );
  }


  if (!examination) {
    throw new Error(
      'QR submission is missing Examination.'
    );
  }


  if (!evaluationType) {
    throw new Error(
      'QR submission is missing Evaluation Type.'
    );
  }


  if (!evaluationDate) {
    throw new Error(
      'QR submission is missing a valid Evaluation Date.'
    );
  }


  if (!result) {
    throw new Error(
      'QR submission is missing a valid Result.'
    );
  }


  if (finalScore === null) {
    throw new Error(
      'QR submission is missing a valid Final Score.'
    );
  }


  // ===================================================================
  // EVALUATOR
  // ===================================================================

  const collegeEvaluator =
    qrSubmission_text_(
      answers['evaluators']
    );


  const siteInstructor =
    qrSubmission_text_(
      answers['clinical site instructor']
    );


  const evaluatorRoute =
    qrSubmission_text_(
      answers['evaluator']
    );


  const evaluator =
    siteInstructor ||
    collegeEvaluator;


  if (!evaluator) {
    throw new Error(
      'QR submission is missing an evaluator.'
    );
  }


  // ===================================================================
  // EMAIL
  // ===================================================================

  const evaluatorEmail =
    qrSubmission_getEmail_(
      response,
      answers
    );


  // ===================================================================
  // STUDENT LOOKUP
  // ===================================================================

  const studentInfo =
    qrSubmission_lookupStudent_(
      student
    );


  if (!studentInfo) {
    throw new Error(
      'Student not found in Students sheet: ' +
      student
    );
  }


  if (!studentInfo.cohort) {
    throw new Error(
      'Student has no Cohort assigned: ' +
      student
    );
  }


  // ===================================================================
  // ATTRIBUTED TERM
  // ===================================================================

  const attributedClinicalTerm =
    clinicalTerm_getForDate(
      evaluationDate
    );


  // ===================================================================
  // UPLOAD
  // ===================================================================

  const uploadValue =
    answers[
      'upload a photo or pdf of the completed evaluation'
    ];


  const fileIds =
    qrSubmission_getFileIds_(
      uploadValue
    );


  if (fileIds.length === 0) {
    throw new Error(
      'QR submission has no uploaded evaluation file.'
    );
  }


  const baseName =
    driveService_buildEvaluationFileName({

      date:
        evaluationDate,

      student:
        student,

      examination:
        examination,

      evaluationType:
        evaluationType

    });


  const filedFiles =
    [];


  fileIds.forEach(
    function(fileId, index) {

      const originalFile =
        DriveApp.getFileById(
          fileId
        );


      const extension =
        qrSubmission_getExtension_(
          originalFile.getName()
        );


      let newName =
        baseName;


      if (fileIds.length > 1) {
        newName +=
          ' - ' +
          (index + 1);
      }


      if (extension) {
        newName +=
          '.' +
          extension;
      }


      const filed =
        driveService_fileUploadedEvaluation({

          fileId:
            fileId,

          student:
            student,

          cohort:
            studentInfo.cohort,

          newFileName:
            newName

        });


      filedFiles.push(
        filed
      );
    }
  );


  const primaryFile =
    filedFiles[0];


  // ===================================================================
  // NORMALIZED CLINICAL TRACKING RECORD
  // ===================================================================

  const record = {

    'Submission ID':
      submissionId,

    'Student ID':
      studentInfo.studentId,

    'Student':
      student,

    'Clinical Term':
      studentInfo.clinicalTerm,

    'Attributed Clinical Term':
      attributedClinicalTerm,

    'Clinical Site':
      clinicalSite,

    'Evaluator':
      evaluator,

    'Evaluator Type':
      '',

    'Evaluation Type':
      evaluationType,

    'Examination':
      examination,

    'Anatomical Area':
      '',

    'Evaluation Date':
      evaluationDate,

    'Performance Score':
      '',

    'Image Score':
      '',

    'Final Score':
      finalScore,

    'Result':
      result,

    'Automatic Failure':
      '',

    'Source':
      'QR Submission',

    'Source Document':
      primaryFile.fileUrl,

    'Submission Method':
      'QR / Uploaded Evaluation',

    'Entered By':
      evaluatorEmail ||
      evaluator,

    'Coordinator Review Status':
      '',

    'Coordinator Flag':
      '',

    'Evaluator Email':
      evaluatorEmail,

    'Evaluator Route':
      evaluatorRoute,

    'Google Form Response ID':
      responseId

  };


  // ===================================================================
  // CLINICAL TRACKING
  // ===================================================================

  const trackingResult =
    clinicalTracking_addRecord(
      record
    );


  if (
    !trackingResult ||
    !trackingResult.success
  ) {

    throw new Error(
      'Clinical Tracking Service did not report success.'
    );
  }


  if (trackingResult.duplicate) {

    console.log(
      'Duplicate QR submission detected. ' +
      'Skipping requirements and notifications.'
    );


    return {

      success:
        true,

      duplicate:
        true,

      submissionId:
        submissionId,

      sourceDocument:
        primaryFile.fileUrl

    };
  }


  // ===================================================================
  // AUTHORITATIVE RECORD FOR EMAIL
  // ===================================================================

  const notificationRecord =
    Object.assign(
      {},
      trackingResult.record || record,
      {
        'Evaluator Email':
          evaluatorEmail,

        'Source Document':
          primaryFile.fileUrl
      }
    );


  // ===================================================================
  // REQUIREMENTS
  // ===================================================================

  let requirementsResult;


  try {

    requirementsResult =
      rebuildStudentRequirementStatus();

  } catch (error) {

    console.log(
      'REQUIREMENTS ERROR: ' +
      error.message
    );


    requirementsResult = {

      success:
        false,

      error:
        error.message

    };
  }


  // ===================================================================
  // NOTIFICATIONS
  // ===================================================================

  let notificationResult;


  try {

    notificationResult =
      notification_sendEvaluationNotificationsSafely(
        notificationRecord
      );

  } catch (error) {

    console.log(
      'NOTIFICATION ERROR: ' +
      error.message
    );


    notificationResult = {

      success:
        false,

      error:
        error.message

    };
  }


  console.log(
    'QR intake complete: ' +
    submissionId
  );


  return {

    success:
      true,

    duplicate:
      false,

    submissionId:
      submissionId,

    files:
      filedFiles,

    requirements:
      requirementsResult,

    notifications:
      notificationResult

  };
}


/**
 * =====================================================================
 * READ FORM ANSWERS
 * =====================================================================
 */
function qrSubmission_getAnswers_(
  response
) {

  const answers = {};


  response
    .getItemResponses()
    .forEach(
      function(itemResponse) {

        const title =
          appNormalize_(
            itemResponse
              .getItem()
              .getTitle()
          );


        answers[title] =
          itemResponse
            .getResponse();
      }
    );


  return answers;
}


/**
 * =====================================================================
 * STUDENT LOOKUP
 * =====================================================================
 */
function qrSubmission_lookupStudent_(
  studentName
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENTS
    );


  const rows =
    appSheetToObjects_(
      sheet
    );


  const target =
    appNormalize_(
      studentName
    );


  const matches =
    rows.filter(
      function(row) {

        return (
          appNormalize_(
            row['Student']
          ) ===
          target
        );
      }
    );


  if (matches.length === 0) {
    return null;
  }


  const active =
    matches.filter(
      function(row) {

        return appIsTrue_(
          row['Active']
        );
      }
    );


  const candidates =
    active.length
      ? active
      : matches;


  candidates.sort(
    function(a, b) {

      return (
        qrSubmission_clinicalNumber_(
          b['Clinical Term']
        ) -
        qrSubmission_clinicalNumber_(
          a['Clinical Term']
        )
      );
    }
  );


  const row =
    candidates[0];


  return {

    studentId:
      qrSubmission_text_(
        row['Student ID']
      ),

    student:
      qrSubmission_text_(
        row['Student']
      ),

    clinicalTerm:
      qrSubmission_text_(
        row['Clinical Term']
      ),

    cohort:
      qrSubmission_text_(
        row['Cohort']
      )

  };
}


/**
 * =====================================================================
 * EMAIL
 * =====================================================================
 */
function qrSubmission_getEmail_(
  response,
  answers
) {

  try {

    const respondentEmail =
      response.getRespondentEmail
        ? response.getRespondentEmail()
        : '';


    if (
      qrSubmission_isEmail_(
        respondentEmail
      )
    ) {

      return respondentEmail;
    }

  } catch (error) {

    console.log(
      'Could not read respondent email: ' +
      error.message
    );
  }


  const fallback =
    qrSubmission_text_(
      answers['email address']
    );


  return qrSubmission_isEmail_(
    fallback
  )
    ? fallback
    : '';
}


/**
 * =====================================================================
 * FILE IDS
 * =====================================================================
 */
function qrSubmission_getFileIds_(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {

    return [];
  }


  if (Array.isArray(value)) {

    return value
      .map(
        function(item) {

          return qrSubmission_fileId_(
            item
          );
        }
      )
      .filter(Boolean);
  }


  const one =
    qrSubmission_fileId_(
      value
    );


  return one
    ? [one]
    : [];
}


/**
 * =====================================================================
 * ONE FILE ID
 * =====================================================================
 */
function qrSubmission_fileId_(
  value
) {

  const text =
    qrSubmission_text_(
      value
    );


  if (!text) {
    return '';
  }


  if (!text.includes('/')) {
    return text;
  }


  let match =
    text.match(
      /\/d\/([A-Za-z0-9_-]+)/
    );


  if (
    match &&
    match[1]
  ) {

    return match[1];
  }


  match =
    text.match(
      /[?&]id=([A-Za-z0-9_-]+)/
    );


  return (
    match &&
    match[1]
  )
    ? match[1]
    : '';
}


/**
 * =====================================================================
 * SCORE
 * =====================================================================
 */
function qrSubmission_parseScore_(
  value
) {

  const text =
    qrSubmission_text_(
      value
    );


  if (!text) {
    return null;
  }


  const match =
    text.match(
      /-?\d+(?:\.\d+)?/
    );


  if (!match) {
    return null;
  }


  const score =
    Number(
      match[0]
    );


  if (
    !Number.isFinite(score)
  ) {

    return null;
  }


  if (
    score < 0 ||
    score > 3
  ) {

    throw new Error(
      'QR Final Score must be between 0 and 3. Received: ' +
      score
    );
  }


  return Math.round(
    score * 100
  ) / 100;
}


/**
 * =====================================================================
 * RESULT
 * =====================================================================
 */
function qrSubmission_parseResult_(
  value
) {

  const text =
    appNormalize_(
      value
    );


  if (
    text === 'pass' ||
    text === 'passed'
  ) {

    return 'PASS';
  }


  if (
    text === 'fail' ||
    text === 'failed'
  ) {

    return 'FAIL';
  }


  return '';
}


/**
 * =====================================================================
 * DATE
 * =====================================================================
 */
function qrSubmission_parseDate_(
  value
) {

  if (!value) {
    return null;
  }


  if (
    value instanceof Date &&
    !isNaN(
      value.getTime()
    )
  ) {

    return value;
  }


  const parsed =
    new Date(
      value
    );


  return isNaN(
    parsed.getTime()
  )
    ? null
    : parsed;
}


/**
 * =====================================================================
 * FILE EXTENSION
 * =====================================================================
 */
function qrSubmission_getExtension_(
  fileName
) {

  const match =
    qrSubmission_text_(
      fileName
    ).match(
      /\.([^.]+)$/
    );


  return match
    ? match[1]
    : '';
}


/**
 * =====================================================================
 * CLINICAL NUMBER
 * =====================================================================
 */
function qrSubmission_clinicalNumber_(
  value
) {

  const match =
    String(
      value || ''
    ).match(
      /clinical\s*(\d+)/i
    );


  return match
    ? Number(
        match[1]
      )
    : 0;
}


/**
 * =====================================================================
 * TEXT
 * =====================================================================
 */
function qrSubmission_text_(
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
 * EMAIL VALIDATION
 * =====================================================================
 */
function qrSubmission_isEmail_(
  value
) {

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    .test(
      qrSubmission_text_(
        value
      )
    );
}


/**
 * =====================================================================
 * INSTALL FORM TRIGGER
 * =====================================================================
 *
 * RUN MANUALLY ONCE.
 */
function qrSubmission_installTrigger() {

  const handler =
    'qrSubmission_onFormSubmit';


  const formId =
    APP_CONFIG
      .FORMS
      .QR_SUBMISSION;


  if (!formId) {

    throw new Error(
      'APP_CONFIG.FORMS.QR_SUBMISSION is not configured.'
    );
  }


  const form =
    FormApp.openById(
      formId
    );


  const triggers =
    ScriptApp.getProjectTriggers();


  let removed =
    0;


  triggers.forEach(
    function(trigger) {

      if (
        trigger.getHandlerFunction() ===
        handler
      ) {

        ScriptApp.deleteTrigger(
          trigger
        );

        removed++;
      }
    }
  );


  const trigger =
    ScriptApp
      .newTrigger(
        handler
      )
      .forForm(
        form
      )
      .onFormSubmit()
      .create();


  console.log(
    'QR trigger installed.'
  );


  console.log(
    'Form: ' +
    form.getTitle()
  );


  console.log(
    'Previous QR triggers removed: ' +
    removed
  );


  return {

    success:
      true,

    form:
      form.getTitle(),

    removed:
      removed,

    triggerId:
      trigger.getUniqueId()

  };
}


/**
 * =====================================================================
 * REMOVE FORM TRIGGER
 * =====================================================================
 */
function qrSubmission_removeTriggers() {

  const handler =
    'qrSubmission_onFormSubmit';


  let removed =
    0;


  ScriptApp
    .getProjectTriggers()
    .forEach(
      function(trigger) {

        if (
          trigger.getHandlerFunction() ===
          handler
        ) {

          ScriptApp.deleteTrigger(
            trigger
          );

          removed++;
        }
      }
    );


  return {

    success:
      true,

    removed:
      removed

  };
}