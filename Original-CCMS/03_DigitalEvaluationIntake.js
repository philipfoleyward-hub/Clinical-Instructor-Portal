/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * DIGITAL EVALUATION INTAKE
 * 03_DigitalEvaluationIntake.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Single production entry point for a submitted Digital Clinical
 * Evaluation Google Form.
 *
 * THIS FILE IS THE ORCHESTRATOR.
 *
 * It does NOT independently:
 *
 *   - calculate evaluation grades
 *   - calculate semester rules
 *   - write directly to Clinical Tracking
 *   - create Drive folder structures
 *   - build evaluation documents
 *   - calculate academic requirements
 *   - compose notification emails
 *
 * Instead it coordinates the authoritative services that own those jobs.
 *
 * =====================================================================
 *
 * PRODUCTION FLOW
 * ---------------
 *
 * Google Form submission
 *
 *      ↓
 *
 * digitalEvaluation_onFormSubmit(e)
 *
 *      ↓
 *
 * gradeClinicalEvaluation(formResponse)
 *
 *      ↓
 *
 * clinicalTerm_getForDate(evaluationDate)
 *
 *      ↓
 *
 * evaluationDocument_generate(grade)
 *
 *      ↓
 *
 * clinicalTracking_addRecord(record)
 *
 *      ↓
 *
 * rebuildStudentRequirementStatus()
 *
 *      ↓
 *
 * notification_sendEvaluationNotificationsSafely(record)
 *
 *      ↓
 *
 * Dashboard automatically sees updated authoritative data.
 *
 * =====================================================================
 */


/**
 * =====================================================================
 * PRODUCTION FORM-SUBMIT ENTRY POINT
 * =====================================================================
 *
 * IMPORTANT:
 *
 * This function MUST receive a GOOGLE FORM form-submit event.
 *
 * It expects:
 *
 *   e.response
 *
 * which is a FormResponse object.
 *
 * Do NOT attach this as:
 *
 *   From spreadsheet -> On form submit
 *
 * Use digitalEvaluation_installTrigger() below instead.
 */
function digitalEvaluation_onFormSubmit(
  e
) {

  const startedAt =
    new Date();


  console.log(
    '=================================================='
  );

  console.log(
    'DIGITAL EVALUATION INTAKE START'
  );

  console.log(
    'Started: ' +
    startedAt
  );

  console.log(
    '=================================================='
  );


  // ===================================================================
  // VALIDATE EVENT
  // ===================================================================

  if (
    !e ||
    !e.response
  ) {

    throw new Error(
      'Digital Evaluation Intake requires a Google Forms ' +
      'form-submit event containing e.response. ' +
      'Use digitalEvaluation_installTrigger() to install the correct trigger.'
    );
  }


  const formResponse =
    e.response;


  // ===================================================================
  // RESPONSE ID / SUBMISSION ID
  //
  // Google Forms response ID becomes the stable identity of this
  // evaluation.
  //
  // Later revision support can use the same identity to UPDATE the
  // existing Clinical Tracking record instead of creating a duplicate.
  // ===================================================================

  const responseId =
    formResponse.getId
      ? String(
          formResponse.getId() || ''
        ).trim()
      : '';


  if (!responseId) {

    throw new Error(
      'The submitted Google Form response does not have a response ID.'
    );
  }


  const submissionId =
    digitalEvaluation_buildSubmissionId_(
      responseId
    );


  console.log(
    'Submission ID: ' +
    submissionId
  );


  // ===================================================================
  // 1. GRADE THE EVALUATION
  // ===================================================================

  console.log(
    'STEP 1 - Grading evaluation...'
  );


  const grade =
    gradeClinicalEvaluation(
      formResponse
    );


  if (!grade) {

    throw new Error(
      'Clinical Evaluation Grading Engine returned no result.'
    );
  }


  console.log(
    'Grade complete.'
  );

  console.log(
    'Final Score: ' +
    grade.finalScore
  );

  console.log(
    'Result: ' +
    grade.result
  );


  // ===================================================================
  // EVALUATION
  // ===================================================================

  const evaluation =
    grade.evaluation || {};


  if (
    !String(
      evaluation.student || ''
    ).trim()
  ) {

    throw new Error(
      'The graded evaluation does not contain a student name.'
    );
  }


  if (
    !evaluation.date
  ) {

    throw new Error(
      'The graded evaluation does not contain an evaluation date.'
    );
  }


  // ===================================================================
  // 2. CLINICAL TERM ATTRIBUTION
  // ===================================================================

  console.log(
    'STEP 2 - Determining attributed clinical term...'
  );


  const attributedClinicalTerm =
    clinicalTerm_getForDate(
      evaluation.date
    );


  console.log(
    'Attributed Clinical Term: ' +
    attributedClinicalTerm
  );


  // ===================================================================
  // RESPONSE METADATA
  // ===================================================================

  const evaluatorEmail =
    digitalEvaluation_getEvaluatorEmail_(
      formResponse
    );


  const editResponseUrl =
    digitalEvaluation_getEditResponseUrl_(
      formResponse
    );


  // ===================================================================
  // 3. GENERATE OFFICIAL DOCUMENT + PDF
  //
  // Generate before Clinical Tracking so Source Document is available
  // when the factual record is written.
  // ===================================================================

  console.log(
    'STEP 3 - Generating evaluation document...'
  );


  const documentResult =
    evaluationDocument_generate(
      grade
    );


  if (
    !documentResult ||
    !documentResult.success
  ) {

    throw new Error(
      'Evaluation Document Generator did not report success.'
    );
  }


  console.log(
    'Evaluation document created.'
  );

  console.log(
    'Google Doc: ' +
    (
      documentResult.docUrl ||
      ''
    )
  );

  console.log(
    'PDF: ' +
    (
      documentResult.pdfUrl ||
      ''
    )
  );


  // ===================================================================
  // 4. BUILD NORMALIZED CLINICAL TRACKING RECORD
  // ===================================================================

  const record =
    digitalEvaluation_buildTrackingRecord_({

      submissionId:
        submissionId,

      grade:
        grade,

      attributedClinicalTerm:
        attributedClinicalTerm,

      documentResult:
        documentResult,

      evaluatorEmail:
        evaluatorEmail,

      editResponseUrl:
        editResponseUrl,

      formResponse:
        formResponse

    });


  // ===================================================================
  // 5. WRITE CLINICAL TRACKING
  // ===================================================================

  console.log(
    'STEP 4 - Writing Clinical Tracking...'
  );


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


  if (
    trackingResult.duplicate
  ) {

    /*
     * At present, duplicate means this response has already been
     * processed.
     *
     * Later:
     *
     * Form revision
     *      ↓
     * same Submission ID
     *      ↓
     * Clinical Tracking UPDATE service
     *
     * For now we safely stop to prevent duplicate requirements /
     * duplicate notifications.
     */
    console.log(
      'Submission already exists in Clinical Tracking. ' +
      'Skipping duplicate downstream processing.'
    );


    return {

      success:
        true,

      duplicate:
        true,

      submissionId:
        submissionId,

      responseId:
        responseId,

      grade:
        grade,

      document:
        {

          docUrl:
            documentResult.docUrl || '',

          pdfUrl:
            documentResult.pdfUrl || ''

        }

    };
  }


  console.log(
    'Clinical Tracking record created.'
  );


  // ===================================================================
  // USE AUTHORITATIVE RECORD RETURNED BY TRACKING SERVICE
  //
  // It may have resolved:
  //
  //   Student ID
  //   Clinical Term
  //   Evaluator Type
  //   Submission Method
  //   Review status
  //
  // Preserve additional intake-only metadata as well.
  // ===================================================================

  const authoritativeRecord =
    Object.assign(

      {},

      trackingResult.record || record,

      {

        'Evaluator Email':
          evaluatorEmail,

        'Edit Response URL':
          editResponseUrl,

        'PDF Document':
          documentResult.pdfUrl || '',

        'Google Form Response ID':
          responseId

      }

    );


  // ===================================================================
  // 6. REBUILD REQUIREMENTS
  //
  // RequirementsEngine reads Clinical Tracking and rebuilds the derived
  // Student Requirement Status table.
  //
  // A Requirements error is IMPORTANT and should be visible in the
  // execution log, but the evaluation itself is already safely recorded.
  // ===================================================================

  console.log(
    'STEP 5 - Rebuilding student requirement status...'
  );


  let requirementsResult =
    null;


  try {

    requirementsResult =
      rebuildStudentRequirementStatus();


    console.log(
      'Requirements rebuild complete.'
    );

  } catch (error) {

    console.log(
      'REQUIREMENTS ENGINE ERROR: ' +
      error.message
    );


    /*
     * Do not erase an already-successful factual evaluation because a
     * derived summary failed.
     *
     * Reconciliation can be rerun after the configuration problem is
     * corrected.
     */
    requirementsResult = {

      success:
        false,

      error:
        error.message

    };
  }


  // ===================================================================
  // 7. NOTIFICATIONS
  //
  // Notification service already protects the core submission from
  // email failures.
  // ===================================================================

  console.log(
    'STEP 6 - Sending notifications...'
  );


  let notificationResult =
    null;


  try {

    notificationResult =
      notification_sendEvaluationNotificationsSafely(
        authoritativeRecord
      );


    console.log(
      'Notification workflow complete.'
    );

  } catch (error) {

    /*
     * Extra safety.
     *
     * The Notification Service already catches its own mail errors,
     * but intake should still never invalidate the evaluation because
     * of notification problems.
     */
    console.log(
      'NOTIFICATION WORKFLOW ERROR: ' +
      error.message
    );


    notificationResult = {

      success:
        false,

      error:
        error.message

    };
  }


  // ===================================================================
  // COMPLETE
  // ===================================================================

  const completedAt =
    new Date();


  console.log(
    '=================================================='
  );

  console.log(
    'DIGITAL EVALUATION INTAKE COMPLETE'
  );

  console.log(
    'Submission ID: ' +
    submissionId
  );

  console.log(
    'Student: ' +
    (
      authoritativeRecord[
        'Student'
      ] || ''
    )
  );

  console.log(
    'Final Score: ' +
    (
      authoritativeRecord[
        'Final Score'
      ]
    )
  );

  console.log(
    'Result: ' +
    (
      authoritativeRecord[
        'Result'
      ] || ''
    )
  );

  console.log(
    'Elapsed ms: ' +
    (
      completedAt.getTime() -
      startedAt.getTime()
    )
  );

  console.log(
    '=================================================='
  );


  return {

    success:
      true,

    duplicate:
      false,

    submissionId:
      submissionId,

    responseId:
      responseId,

    record:
      authoritativeRecord,

    document:
      {

        docId:
          documentResult.docId || '',

        docUrl:
          documentResult.docUrl || '',

        pdfId:
          documentResult.pdfId || '',

        pdfUrl:
          documentResult.pdfUrl || ''

      },

    requirements:
      requirementsResult,

    notifications:
      notificationResult

  };
}


/**
 * =====================================================================
 * BUILD NORMALIZED TRACKING RECORD
 * =====================================================================
 */
function digitalEvaluation_buildTrackingRecord_(
  args
) {

  const grade =
    args.grade || {};


  const evaluation =
    grade.evaluation || {};


  const automaticFailure =
    grade.automaticFailure || {};


  const performance =
    grade.performance || {};


  const image =
    grade.image || {};


  // ===================================================================
  // CURRENT STUDENT SNAPSHOT
  //
  // ClinicalTrackingService can resolve this itself, but supplying
  // Student ID and current term here makes the source record explicit
  // and reduces ambiguity when a student has multiple assignment rows.
  // ===================================================================

  const studentSnapshot =
    digitalEvaluation_lookupStudent_(

      evaluation.student

    );


  return {

    // -----------------------------------------------------------------
    // Identity
    // -----------------------------------------------------------------

    'Submission ID':
      args.submissionId,

    'Student ID':
      studentSnapshot
        ? studentSnapshot.studentId
        : '',

    'Student':
      evaluation.student || '',


    // -----------------------------------------------------------------
    // Semester / location
    // -----------------------------------------------------------------

    'Clinical Term':
      studentSnapshot
        ? studentSnapshot.clinicalTerm
        : '',

    'Attributed Clinical Term':
      args.attributedClinicalTerm || '',

    'Clinical Site':
      evaluation.clinicalSite || '',


    // -----------------------------------------------------------------
    // Evaluator
    // -----------------------------------------------------------------

    'Evaluator':
      evaluation.evaluator || '',

    /*
     * ClinicalTrackingService resolves Evaluator Type from Evaluators
     * when possible.
     */
    'Evaluator Type':
      '',

    'Evaluator Email':
      args.evaluatorEmail || '',


    // -----------------------------------------------------------------
    // Evaluation
    // -----------------------------------------------------------------

    'Evaluation Type':
      evaluation.evaluationType || '',

    'Examination':
      evaluation.examination || '',

    /*
     * Current Digital Grading Engine does not separately return an
     * Anatomical Area field.
     *
     * Leave blank rather than inventing one.
     */
    'Anatomical Area':
      '',

    'Evaluation Date':
      evaluation.date || '',


    // -----------------------------------------------------------------
    // Scores
    // -----------------------------------------------------------------

    'Performance Score':
      performance.score === null ||
      performance.score === undefined

        ? ''

        : grading_round2_(
            performance.score
          ),

    'Image Score':
      image.score === null ||
      image.score === undefined

        ? ''

        : grading_round2_(
            image.score
          ),

    'Final Score':
      grade.finalScore === null ||
      grade.finalScore === undefined

        ? ''

        : grade.finalScore,

    'Result':
      grade.result || '',

    'Automatic Failure':
      automaticFailure.occurred
        ? 'Yes'
        : 'No',


    // -----------------------------------------------------------------
    // Source / audit
    // -----------------------------------------------------------------

    'Source':
      APP_CONFIG.SOURCES &&
      APP_CONFIG.SOURCES
        .DIGITAL_EVALUATION

        ? APP_CONFIG.SOURCES
            .DIGITAL_EVALUATION

        : 'Digital Evaluation',

    'Submission Method':
      'Digital Evaluation',

    /*
     * Google Doc is the live completed evaluation link shown to
     * coordinators/dashboard.
     *
     * PDF remains available separately for evaluator download/printing.
     */
    'Source Document':
      args.documentResult.docUrl || '',

    'PDF Document':
      args.documentResult.pdfUrl || '',

    'Entered By':
      'Automated Digital Evaluation Intake',

    'Google Form Response ID':
      args.formResponse &&
      args.formResponse.getId

        ? args.formResponse.getId()

        : '',

    'Edit Response URL':
      args.editResponseUrl || '',


    // -----------------------------------------------------------------
    // Coordinator review
    // -----------------------------------------------------------------

    /*
     * Let ClinicalTrackingService apply the configured default review
     * status.
     */
    'Coordinator Review Status':
      '',

    'Coordinator Flag':
      ''

  };
}


/**
 * =====================================================================
 * STUDENT SNAPSHOT LOOKUP
 * =====================================================================
 *
 * Used only to capture:
 *
 *   Student ID
 *   current Clinical Term
 *
 * The factual Clinical Site remains the one submitted on the evaluation.
 */
function digitalEvaluation_lookupStudent_(
  studentName
) {

  if (!studentName) {

    return null;
  }


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


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
    rows.filter(row =>

      appNormalize_(
        row['Student']
      ) === target

    );


  if (
    matches.length === 0
  ) {

    return null;
  }


  /*
   * Prefer active rows.
   */
  const active =
    matches.filter(row =>

      appIsTrue_(
        row['Active']
      )

    );


  const candidates =
    active.length
      ? active
      : matches;


  /*
   * If the student has multiple site-assignment rows, choose the most
   * advanced Clinical Term for the snapshot.
   */
  candidates.sort(
    (a, b) =>

      digitalEvaluation_clinicalNumber_(
        b['Clinical Term']
      ) -

      digitalEvaluation_clinicalNumber_(
        a['Clinical Term']
      )
  );


  const row =
    candidates[0];


  return {

    studentId:
      String(
        row['Student ID'] || ''
      ).trim(),

    student:
      String(
        row['Student'] || ''
      ).trim(),

    clinicalTerm:
      String(
        row['Clinical Term'] || ''
      ).trim(),

    cohort:
      String(
        row['Cohort'] || ''
      ).trim()

  };
}


/**
 * =====================================================================
 * EVALUATOR EMAIL
 * =====================================================================
 *
 * Preferred:
 *
 *   Google Forms collected respondent email
 *
 * Fallback:
 *
 *   answered form field named one of:
 *
 *     Evaluator Email
 *     Email Address
 *     Email
 */
function digitalEvaluation_getEvaluatorEmail_(
  formResponse
) {

  // ===================================================================
  // COLLECT EMAIL ADDRESSES
  // ===================================================================

  try {

    if (
      formResponse.getRespondentEmail
    ) {

      const respondentEmail =
        String(
          formResponse.getRespondentEmail() ||
          ''
        ).trim();


      if (
        digitalEvaluation_isLikelyEmail_(
          respondentEmail
        )
      ) {

        return respondentEmail;
      }
    }

  } catch (error) {

    console.log(
      'Could not read respondent email: ' +
      error.message
    );
  }


  // ===================================================================
  // FORM QUESTION FALLBACK
  // ===================================================================

  const answers =
    digitalEvaluation_getAnswersByTitle_(
      formResponse
    );


  const candidateTitles = [

    'evaluator email',
    'evaluator email address',
    'email address',
    'email'

  ];


  for (
    let i = 0;
    i < candidateTitles.length;
    i++
  ) {

    const value =
      String(
        answers[
          candidateTitles[i]
        ] || ''
      ).trim();


    if (
      digitalEvaluation_isLikelyEmail_(
        value
      )
    ) {

      return value;
    }
  }


  return '';
}


/**
 * =====================================================================
 * EDIT RESPONSE URL
 * =====================================================================
 *
 * This is intentionally captured NOW even though revision processing
 * will be activated later.
 *
 * That gives us the foundation for the chair's requested:
 *
 *     Edit / Revise Evaluation
 *
 * button without changing the intake identity model later.
 */
function digitalEvaluation_getEditResponseUrl_(
  formResponse
) {

  try {

    if (
      formResponse.getEditResponseUrl
    ) {

      return String(
        formResponse
          .getEditResponseUrl() ||
        ''
      ).trim();
    }

  } catch (error) {

    console.log(
      'Edit response URL unavailable: ' +
      error.message
    );
  }


  return '';
}


/**
 * =====================================================================
 * ANSWERS BY NORMALIZED TITLE
 * =====================================================================
 */
function digitalEvaluation_getAnswersByTitle_(
  formResponse
) {

  const answers = {};


  formResponse
    .getItemResponses()
    .forEach(itemResponse => {

      const item =
        itemResponse.getItem();


      const title =
        appNormalize_(
          item.getTitle()
        );


      if (!title) {

        return;
      }


      const response =
        itemResponse.getResponse();


      if (
        response === null ||
        response === undefined ||
        response === ''
      ) {

        return;
      }


      answers[title] =
        Array.isArray(response)

          ? response.join(', ')

          : response;

    });


  return answers;
}


/**
 * =====================================================================
 * SUBMISSION ID
 * =====================================================================
 *
 * Stable response identity.
 *
 * Example:
 *
 * DIGITAL-2_ABCD1234...
 *
 * Do not use timestamps alone; timestamps are not a safe unique key.
 */
function digitalEvaluation_buildSubmissionId_(
  responseId
) {

  const cleaned =
    String(
      responseId || ''
    )
      .trim();


  if (!cleaned) {

    throw new Error(
      'Cannot build Digital Evaluation Submission ID without response ID.'
    );
  }


  return (
    'DIGITAL-' +
    cleaned
  );
}


/**
 * =====================================================================
 * CLINICAL NUMBER
 * =====================================================================
 */
function digitalEvaluation_clinicalNumber_(
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
 * EMAIL CHECK
 * =====================================================================
 */
function digitalEvaluation_isLikelyEmail_(
  value
) {

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    .test(
      String(
        value || ''
      ).trim()
    );
}


/**
 * =====================================================================
 * INSTALL DIGITAL EVALUATION FORM TRIGGER
 * =====================================================================
 *
 * RUN THIS MANUALLY ONCE.
 *
 * It creates the correct trigger:
 *
 *   Digital Clinical Evaluation Google Form
 *        ↓
 *   On form submit
 *        ↓
 *   digitalEvaluation_onFormSubmit(e)
 *
 * This gives the intake a real FormResponse in e.response.
 *
 * The installer also removes duplicate triggers for THIS handler so we
 * do not process one submission more than once.
 */
function digitalEvaluation_installTrigger() {

  const handlerName =
    'digitalEvaluation_onFormSubmit';


  const formId =
    APP_CONFIG.FORMS
      .DIGITAL_CLINICAL_EVALUATION;


  if (!formId) {

    throw new Error(
      'APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION is not configured.'
    );
  }


  const form =
    FormApp.openById(
      formId
    );


  // ===================================================================
  // REMOVE OLD COPIES OF THIS SPECIFIC HANDLER
  // ===================================================================

  const triggers =
    ScriptApp.getProjectTriggers();


  let removed =
    0;


  triggers.forEach(trigger => {

    if (
      trigger.getHandlerFunction() ===
      handlerName
    ) {

      ScriptApp.deleteTrigger(
        trigger
      );


      removed++;
    }

  });


  // ===================================================================
  // CREATE ONE CORRECT FORM TRIGGER
  // ===================================================================

  const trigger =
    ScriptApp
      .newTrigger(
        handlerName
      )
      .forForm(
        form
      )
      .onFormSubmit()
      .create();


  console.log(
    '=========================================='
  );

  console.log(
    'DIGITAL EVALUATION TRIGGER INSTALLED'
  );

  console.log(
    'Form: ' +
    form.getTitle()
  );

  console.log(
    'Form ID: ' +
    formId
  );

  console.log(
    'Handler: ' +
    handlerName
  );

  console.log(
    'Old handler triggers removed: ' +
    removed
  );

  console.log(
    'Trigger ID: ' +
    trigger.getUniqueId()
  );

  console.log(
    '=========================================='
  );


  return {

    success:
      true,

    form:
      form.getTitle(),

    formId:
      formId,

    handler:
      handlerName,

    removedOldTriggers:
      removed,

    triggerId:
      trigger.getUniqueId()

  };
}


/**
 * =====================================================================
 * REMOVE DIGITAL EVALUATION TRIGGER
 * =====================================================================
 *
 * Maintenance utility only.
 */
function digitalEvaluation_removeTriggers() {

  const handlerName =
    'digitalEvaluation_onFormSubmit';


  const triggers =
    ScriptApp.getProjectTriggers();


  let removed =
    0;


  triggers.forEach(trigger => {

    if (
      trigger.getHandlerFunction() ===
      handlerName
    ) {

      ScriptApp.deleteTrigger(
        trigger
      );


      removed++;
    }

  });


  console.log(
    'Removed ' +
    removed +
    ' Digital Evaluation intake trigger(s).'
  );


  return {

    success:
      true,

    removed:
      removed

  };
}


/**
 * =====================================================================
 * TEST — PROCESS MOST RECENT DIGITAL EVALUATION
 * =====================================================================
 *
 * WARNING:
 *
 * THIS IS NOT READ ONLY.
 *
 * It runs the production intake pipeline on the most recent form
 * response.
 *
 * Because Clinical Tracking has Submission ID duplicate protection,
 * rerunning the same response will not create a second tracking row.
 *
 * HOWEVER:
 *
 * Document generation occurs before duplicate detection, so using this
 * repeatedly can create duplicate test Doc/PDF files.
 *
 * Use this only when deliberately testing.
 */
function testDigitalEvaluationIntakeLatest() {

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


  const event = {

    response:
      response

  };


  return digitalEvaluation_onFormSubmit(
    event
  );
}