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
/** Native Google Forms production trigger — preserved contract. */
function digitalEvaluation_onFormSubmit(e) {
  if (!e || !e.response || typeof e.response.getItemResponses !== 'function') {
    throw new Error('Digital Evaluation Intake requires a native FormResponse event. Use digitalEvaluation_installTrigger().');
  }
  const formResponse=e.response;
  const responseId=String(formResponse.getId ? formResponse.getId() || '' : '').trim();
  if (!responseId) throw new Error('Google Form response has no response ID');
  return digitalEvaluation_commitGraded_({
    submissionId:digitalEvaluation_buildSubmissionId_(responseId), responseId:responseId,
    grade:gradeClinicalEvaluation(formResponse), formResponse:formResponse,
    evaluatorEmail:digitalEvaluation_getEvaluatorEmail_(formResponse),
    editResponseUrl:digitalEvaluation_getEditResponseUrl_(formResponse)
  });
}

/**
 * Portal ingress (invoked ONLY by verified server-to-server gateway).
 * The portal does not grade or own other domain side effects; the original
 * intake orchestrates both ingress types through digitalEvaluation_commitGraded_.
 * A portal identity is ALWAYS re-resolved from authoritative reference rows.
 */
function digitalEvaluation_submitPortal_(data) {
  if (!data || !/^CI-PORTAL-[a-f0-9-]{36}$/i.test(String(data.submissionId||''))) {
    throw new Error('Invalid stable portal submission ID');
  }
  const outcome=gradeClinicalEvaluationFromPortalPayload(data);
  if (String(outcome.identity.evaluatorId)!==String(data.authenticatedEvaluatorId||'')) {
    throw new Error('Portal identity mismatch');
  }
  return digitalEvaluation_commitGraded_({
    submissionId:String(data.submissionId), responseId:'', grade:outcome.grade,
    identity:outcome.identity, source:'CI Portal'
  });
}

/**
 * SINGLE AUTHORITATIVE SIDE-EFFECT ORCHESTRATOR.
 * Build factual Doc/PDF, append one Clinical Tracking record, rebuild
 * requirements, notify coordinator through existing NotificationService.
 * Deliberately retains native Form response metadata for legacy submissions.
 * A script lock protects the submission-ID check + document + ledger append.
 */
function digitalEvaluation_commitGraded_(args) {
  const start=new Date();
  const grade=args.grade;
  if (!grade || !grade.evaluation || !String(grade.evaluation.student||'').trim() || !grade.evaluation.date) {
    throw new Error('Graded evaluation identity/date missing');
  }
  if (!String(args.submissionId||'').trim()) throw new Error('Submission ID required');
  const submissionId=String(args.submissionId);
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  let docResult,trackingResult;
  let duplicate=false;
  let authoritativeRecord=null;
  const journalKey='CCMS_DOCUMENT_PENDING_'+Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,submissionId)
  ).replace(/=+$/,'');
  try {
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Original CCMS workbook context unavailable');
    const tracking=appGetSheet_(ss,APP_CONFIG.SHEETS.CLINICAL_TRACKING);
    // DUPLICATE GUARD BEFORE DOCUMENT GENERATION — native Form and portal.
    if (clinicalTracking_submissionExists_(tracking,submissionId)) {
      duplicate=true;
      console.log('Submission already committed: '+submissionId);
    } else {
      const clinicalTerm=clinicalTerm_getForDate(grade.evaluation.date);
      const props=PropertiesService.getScriptProperties();
      let stored=null;
      try {stored=JSON.parse(props.getProperty(journalKey)||'null');} catch(e){stored=null;}
      if (stored && stored.submissionId===submissionId && stored.docUrl && stored.pdfUrl) {
        docResult=stored;
        console.log('Resuming document artifacts from previous interrupted submission');
      } else {
        docResult=evaluationDocument_generate(grade);
        if (!docResult || !docResult.success || !docResult.docUrl || !docResult.pdfUrl) {
          throw new Error('Authoritative document generator did not return verified Doc/PDF URLs');
        }
        // A retry after this save can reuse the generated artifacts rather
        // than produce duplicate files. A crash DURING document generation
        // remains a known recovery case requiring Drive verification.
        props.setProperty(journalKey,JSON.stringify(Object.assign({submissionId:submissionId},docResult)));
      }
      const record=digitalEvaluation_buildTrackingRecord_({
        submissionId:submissionId,grade:grade,attributedClinicalTerm:clinicalTerm,
        documentResult:docResult,evaluatorEmail:args.evaluatorEmail||'',
        editResponseUrl:args.editResponseUrl||'',formResponse:args.formResponse||null
      });
      if (args.identity) {
        const id=args.identity;
        // ORIGINAL identity is resolved by the CCMS adapter; never trust
        // student/evaluator/site text supplied directly by browser.
        record['Student ID']=id.studentId;
        record['Clinical Site ID']=id.siteId;
        record['Evaluator ID']=id.evaluatorId;
        record['Evaluator Type']=id.evaluatorType;
        record['Examination ID']=id.examinationId;
        record['Entered By']='CI Portal — '+id.evaluator+' ('+id.evaluatorId+')';
        record['Submission Method']='CI Portal';
        record['Source']='CI Portal';
      }
      trackingResult=clinicalTracking_addRecord(record);
      if (!trackingResult || !trackingResult.success) {
        throw new Error('Clinical Tracking append not confirmed');
      }
      duplicate=Boolean(trackingResult.duplicate);
      if (!duplicate) {
        authoritativeRecord=Object.assign({},trackingResult.record||record,{
          'Evaluator Email':args.evaluatorEmail||'',
          'Edit Response URL':args.editResponseUrl||'',
          'PDF Document':docResult.pdfUrl||'',
          'Google Form Response ID':args.responseId||''
        });
        props.deleteProperty(journalKey);
      }
    }
  } finally {
    lock.releaseLock();
  }
  if (duplicate) return {
    success:true,duplicate:true,submissionId:submissionId,responseId:args.responseId||'',
    message:'Existing Clinical Tracking record retained; no second document or notification generated.'
  };
  let requirementsResult,notificationResult;
  try { requirementsResult=rebuildStudentRequirementStatus(); }
  catch(e){ console.log('Requirements reconciliation needed: '+e.message); requirementsResult={success:false,error:e.message}; }
  try { notificationResult=notification_sendEvaluationNotificationsSafely(authoritativeRecord); }
  catch(e){ console.log('Coordinator notification reconciliation needed: '+e.message); notificationResult={success:false,error:e.message}; }
  console.log('CCMS evaluation '+submissionId+' completed in '+(new Date()-start)+'ms');
  return {
    success:true,duplicate:false,submissionId:submissionId,responseId:args.responseId||'',
    record:authoritativeRecord,
    document:{docId:docResult.docId||'',docUrl:docResult.docUrl||'',pdfId:docResult.pdfId||'',pdfUrl:docResult.pdfUrl||''},
    requirements:requirementsResult,notifications:notificationResult
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