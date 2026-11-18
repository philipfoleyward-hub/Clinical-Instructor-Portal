/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * NOTIFICATION SERVICE
 * 05_NotificationService.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Single authoritative service for automated system notifications.
 *
 * THIS FILE OWNS:
 *
 *   - New clinical evaluation notifications
 *   - Failure alerts
 *   - Coordinator review links
 *   - Links to completed evaluation documents
 *   - Evaluator confirmation emails
 *   - Future workflow notifications
 *
 * THIS FILE DOES NOT:
 *
 *   - Process form submissions
 *   - Grade evaluations
 *   - Write Clinical Tracking
 *   - Move Drive files
 *   - Calculate requirements
 *
 * IMPORTANT
 * ---------
 * Notification failure should NEVER invalidate an otherwise
 * successful clinical evaluation submission.
 *
 * Intake scripts should call this service inside try/catch.
 * =====================================================================
 */


/**
 * =====================================================================
 * NEW CLINICAL EVALUATION
 * =====================================================================
 *
 * Main notification sent to program personnel after:
 *
 *   1. Submission is processed
 *   2. File is successfully stored
 *   3. Clinical Tracking record is successfully created
 *
 * Expected normalized record:
 *
 * {
 *   'Submission ID': '',
 *   'Student ID': '',
 *   'Student': '',
 *   'Clinical Term': '',
 *   'Clinical Site': '',
 *   'Evaluator': '',
 *   'Evaluator Type': '',
 *   'Evaluation Type': '',
 *   'Examination': '',
 *   'Evaluation Date': '',
 *   'Performance Score': '',
 *   'Image Score': '',
 *   'Final Score': '',
 *   'Result': '',
 *   'Automatic Failure': '',
 *   'Source Document': ''
 * }
 */
function notification_sendNewEvaluation(
  record
) {

  if (!record) {

    throw new Error(
      'NotificationService requires an evaluation record.'
    );
  }


  const recipients =
    notification_getCoordinatorRecipients_();


  if (
    recipients.length === 0
  ) {

    console.log(
      'No coordinator notification recipients are configured.'
    );


    return {

      success:
        false,

      skipped:
        true,

      reason:
        'No recipients configured'

    };
  }


  // ===================================================================
  // VALUES
  // ===================================================================

  const student =
    notification_firstValue_(
      record,
      [
        'Student',
        'Student Name'
      ]
    );


  const studentId =
    notification_firstValue_(
      record,
      [
        'Student ID'
      ]
    );


  const clinicalTerm =
    notification_firstValue_(
      record,
      [
        'Clinical Term'
      ]
    );


  const clinicalSite =
    notification_firstValue_(
      record,
      [
        'Clinical Site'
      ]
    );


  const evaluator =
    notification_firstValue_(
      record,
      [
        'Evaluator'
      ]
    );


  const evaluatorType =
    notification_firstValue_(
      record,
      [
        'Evaluator Type'
      ]
    );


  const evaluationType =
    notification_firstValue_(
      record,
      [
        'Evaluation Type'
      ]
    );


  const examination =
    notification_firstValue_(
      record,
      [
        'Examination'
      ]
    );


  const result =
    String(
      notification_firstValue_(
        record,
        [
          'Result'
        ]
      ) || ''
    )
      .trim()
      .toUpperCase();


  const performanceScore =
    notification_firstValuePreserveZero_(
      record,
      [
        'Performance Score'
      ]
    );


  const imageScore =
    notification_firstValuePreserveZero_(
      record,
      [
        'Image Score'
      ]
    );


  const finalScore =
    notification_firstValuePreserveZero_(
      record,
      [
        'Final Score'
      ]
    );


  const automaticFailure =
    notification_firstValue_(
      record,
      [
        'Automatic Failure'
      ]
    );


  const submissionId =
    notification_firstValue_(
      record,
      [
        'Submission ID',
        'Record ID',
        'Record/ID'
      ]
    );


  const sourceDocument =
    notification_firstValue_(
      record,
      [
        'Source Document',
        'File URL'
      ]
    );


  const evaluationDate =
    notification_firstValue_(
      record,
      [
        'Evaluation Date',
        'Date'
      ]
    );


  const evaluationDateText =
    notification_formatDate_(
      evaluationDate
    );


  // ===================================================================
  // SUBJECT
  // ===================================================================

  let subject;


  if (
    result === 'FAIL'
  ) {

    subject =
      'Clinical Evaluation | FAIL - REVIEW REQUIRED | ' +
      student +
      ' | ' +
      examination;

  } else {

    subject =
      'Clinical Evaluation | ' +
      (
        result ||
        'Submitted'
      ) +
      ' | ' +
      student +
      ' | ' +
      examination;
  }


  // ===================================================================
  // BUTTONS
  // ===================================================================

  let buttonsHtml =
    '';


  let buttonsPlain =
    '';


  // -------------------------------------------------------------------
  // COMPLETED EVALUATION
  // -------------------------------------------------------------------

  if (
    sourceDocument
  ) {

    buttonsHtml +=
      notification_button_(
        'View Completed Evaluation',
        sourceDocument
      );


    buttonsPlain +=
      '\nView Completed Evaluation:\n' +
      sourceDocument +
      '\n';
  }


  // -------------------------------------------------------------------
  // DASHBOARD / REVIEW
  // -------------------------------------------------------------------

  const dashboardUrl =
    notification_buildDashboardUrl_(
      studentId,
      submissionId
    );


  if (
    dashboardUrl
  ) {

    buttonsHtml +=
      notification_button_(
        'Review New Competency',
        dashboardUrl
      );


    buttonsPlain +=
      '\nReview New Competency:\n' +
      dashboardUrl +
      '\n';
  }


  // ===================================================================
  // FAILURE WARNING
  // ===================================================================

  let warningHtml =
    '';


  let warningPlain =
    '';


  if (
    result === 'FAIL'
  ) {

    warningHtml =
      '<div style="' +
        'margin:18px 0;' +
        'padding:12px 14px;' +
        'background:#fce8e6;' +
        'border-left:5px solid #b3261e;' +
        'font-weight:bold;' +
        'color:#8c1d18;' +
      '">' +

        'Attention Required: ' +
        'This evaluation was submitted as FAIL. ' +
        'Please review the evaluation in the Clinical Dashboard.' +

      '</div>';


    warningPlain =
      '\nATTENTION REQUIRED: ' +
      'This evaluation was submitted as FAIL. ' +
      'Please review the evaluation.\n';
  }


  // ===================================================================
  // AUTOMATIC FAILURE WARNING
  // ===================================================================

  let autoFailureHtml =
    '';


  let autoFailurePlain =
    '';


  if (
    notification_isYes_(
      automaticFailure
    )
  ) {

    autoFailureHtml =
      '<div style="' +
        'margin:18px 0;' +
        'padding:12px 14px;' +
        'background:#fef7e0;' +
        'border-left:5px solid #f9ab00;' +
        'font-weight:bold;' +
      '">' +

        'Automatic Failure condition recorded.' +

      '</div>';


    autoFailurePlain =
      '\nAutomatic Failure condition recorded.\n';
  }


  // ===================================================================
  // RESULT DISPLAY
  // ===================================================================

  const resultHtml =
    notification_resultBadge_(
      result
    );


  // ===================================================================
  // HTML BODY
  // ===================================================================

  const htmlBody =

    '<div style="' +
      'font-family:Arial,Helvetica,sans-serif;' +
      'font-size:14px;' +
      'color:#202124;' +
      'max-width:680px;' +
    '">' +

      '<h2 style="' +
        'color:#1f4e79;' +
        'margin-bottom:6px;' +
      '">' +
        'Clinical Evaluation Submitted' +
      '</h2>' +

      '<p style="margin-top:0;">' +
        'A clinical evaluation has been processed, ' +
        'filed in the student&apos;s Clinical Evaluations folder, ' +
        'and added to Clinical Tracking.' +
      '</p>' +

      warningHtml +

      autoFailureHtml +

      '<table style="' +
        'border-collapse:collapse;' +
        'width:100%;' +
        'margin-top:18px;' +
      '">' +

        notification_row_(
          'Student',
          student
        ) +

        notification_row_(
          'Student ID',
          studentId
        ) +

        notification_row_(
          'Clinical Term',
          clinicalTerm
        ) +

        notification_row_(
          'Clinical Site',
          clinicalSite
        ) +

        notification_row_(
          'Evaluator',
          evaluator +
          (
            evaluatorType
              ? ' - ' +
                evaluatorType
              : ''
          )
        ) +

        notification_row_(
          'Evaluation Date',
          evaluationDateText
        ) +

        notification_row_(
          'Evaluation Type',
          evaluationType
        ) +

        notification_row_(
          'Examination',
          examination
        ) +

        notification_optionalRow_(
          'Performance Score',
          performanceScore
        ) +

        notification_optionalRow_(
          'Image Score',
          imageScore
        ) +

        notification_optionalRow_(
          'Final Score',
          finalScore
        ) +

        '<tr>' +

          '<td style="' +
            'padding:8px;' +
            'border-bottom:1px solid #dddddd;' +
            'font-weight:bold;' +
            'width:180px;' +
          '">' +
            'Result' +
          '</td>' +

          '<td style="' +
            'padding:8px;' +
            'border-bottom:1px solid #dddddd;' +
          '">' +

            resultHtml +

          '</td>' +

        '</tr>' +

        notification_row_(
          'Submission ID',
          submissionId
        ) +

      '</table>' +

      buttonsHtml +

      '<p style="' +
        'font-size:12px;' +
        'color:#666666;' +
        'margin-top:25px;' +
      '">' +
        'This is an automated notification from the ' +
        'Radiologic Technology Clinical Evaluation System.' +
      '</p>' +

    '</div>';


  // ===================================================================
  // PLAIN TEXT FALLBACK
  // ===================================================================

  let plainBody =

    'A clinical evaluation has been processed.\n\n' +

    'Student: ' +
    student +
    '\n' +

    'Student ID: ' +
    studentId +
    '\n' +

    'Clinical Term: ' +
    clinicalTerm +
    '\n' +

    'Clinical Site: ' +
    clinicalSite +
    '\n' +

    'Evaluator: ' +
    evaluator +
    (
      evaluatorType
        ? ' - ' +
          evaluatorType
        : ''
    ) +
    '\n' +

    'Evaluation Date: ' +
    evaluationDateText +
    '\n' +

    'Evaluation Type: ' +
    evaluationType +
    '\n' +

    'Examination: ' +
    examination +
    '\n';


  if (
    performanceScore !== ''
  ) {

    plainBody +=
      'Performance Score: ' +
      performanceScore +
      '\n';
  }


  if (
    imageScore !== ''
  ) {

    plainBody +=
      'Image Score: ' +
      imageScore +
      '\n';
  }


  if (
    finalScore !== ''
  ) {

    plainBody +=
      'Final Score: ' +
      finalScore +
      '\n';
  }


  plainBody +=

    'Result: ' +
    result +
    '\n' +

    'Submission ID: ' +
    submissionId +
    '\n' +

    warningPlain +

    autoFailurePlain +

    buttonsPlain;


  // ===================================================================
  // SEND
  // ===================================================================

  MailApp.sendEmail({

    to:
      recipients.join(','),

    subject:
      subject,

    body:
      plainBody,

    htmlBody:
      htmlBody,

    name:
      APP_CONFIG.NOTIFICATIONS
        .SENDER_NAME ||
      'Clinical Evaluation System'

  });


  console.log(
    'Clinical evaluation notification sent to: ' +
    recipients.join(', ')
  );


  return {

    success:
      true,

    recipients:
      recipients,

    subject:
      subject

  };
}


/**
 * =====================================================================
 * EVALUATOR CONFIRMATION
 * =====================================================================
 *
 * This is deliberately its own function because coordinator
 * notifications and evaluator confirmations are different workflows.
 *
 * We can later add:
 *
 *   - Edit Response button
 *   - Download PDF button
 *   - Revision instructions
 *
 * without changing the coordinator notification.
 */
function notification_sendEvaluatorConfirmation(
  record
) {

  if (!record) {

    throw new Error(
      'NotificationService requires an evaluation record.'
    );
  }


  const evaluatorEmail =
    notification_firstValue_(
      record,
      [
        'Evaluator Email'
      ]
    );


  if (!evaluatorEmail) {

    return {

      success:
        false,

      skipped:
        true,

      reason:
        'No evaluator email available'

    };
  }


  const student =
    notification_firstValue_(
      record,
      [
        'Student',
        'Student Name'
      ]
    );


  const examination =
    notification_firstValue_(
      record,
      [
        'Examination'
      ]
    );


  const result =
    notification_firstValue_(
      record,
      [
        'Result'
      ]
    );


  const sourceDocument =
    notification_firstValue_(
      record,
      [
        'Source Document',
        'File URL'
      ]
    );


  const subject =
    'Clinical Evaluation Received | ' +
    student +
    ' | ' +
    examination;


  let htmlBody =

    '<div style="' +
      'font-family:Arial,Helvetica,sans-serif;' +
      'font-size:14px;' +
      'color:#202124;' +
      'max-width:650px;' +
    '">' +

      '<h2 style="color:#1f4e79;">' +
        'Clinical Evaluation Received' +
      '</h2>' +

      '<p>' +
        'Your clinical evaluation submission has been received successfully.' +
      '</p>' +

      '<table style="' +
        'border-collapse:collapse;' +
        'width:100%;' +
      '">' +

        notification_row_(
          'Student',
          student
        ) +

        notification_row_(
          'Examination',
          examination
        ) +

        notification_row_(
          'Result',
          result
        ) +

      '</table>';


  if (
    sourceDocument
  ) {

    htmlBody +=
      notification_button_(
        'View Completed Evaluation',
        sourceDocument
      );
  }


  htmlBody +=

      '<p style="' +
        'font-size:12px;' +
        'color:#666666;' +
        'margin-top:25px;' +
      '">' +
        'This is an automated confirmation from the ' +
        'Radiologic Technology Clinical Evaluation System.' +
      '</p>' +

    '</div>';


  let plainBody =

    'Your clinical evaluation submission has been received successfully.\n\n' +

    'Student: ' +
    student +
    '\n' +

    'Examination: ' +
    examination +
    '\n' +

    'Result: ' +
    result +
    '\n';


  if (
    sourceDocument
  ) {

    plainBody +=
      '\nView Completed Evaluation:\n' +
      sourceDocument;
  }


  MailApp.sendEmail({

    to:
      evaluatorEmail,

    subject:
      subject,

    body:
      plainBody,

    htmlBody:
      htmlBody,

    name:
      APP_CONFIG.NOTIFICATIONS
        .SENDER_NAME ||
      'Clinical Evaluation System'

  });


  return {

    success:
      true,

    recipient:
      evaluatorEmail

  };
}


/**
 * =====================================================================
 * SAFE NOTIFICATION WRAPPER
 * =====================================================================
 *
 * This is the function our intake workflows should normally use.
 *
 * Notification problems are logged but do NOT throw back into the
 * submission workflow.
 */
function notification_sendEvaluationNotificationsSafely(
  record
) {

  const result = {

    coordinator:
      null,

    evaluator:
      null

  };


  try {

    result.coordinator =
      notification_sendNewEvaluation(
        record
      );

  } catch (error) {

    console.log(
      'COORDINATOR NOTIFICATION ERROR: ' +
      error.message
    );


    result.coordinator = {

      success:
        false,

      error:
        error.message

    };
  }


  try {

    result.evaluator =
      notification_sendEvaluatorConfirmation(
        record
      );

  } catch (error) {

    console.log(
      'EVALUATOR NOTIFICATION ERROR: ' +
      error.message
    );


    result.evaluator = {

      success:
        false,

      error:
        error.message

    };
  }


  return result;
}


/**
 * =====================================================================
 * COORDINATOR RECIPIENTS
 * =====================================================================
 */
function notification_getCoordinatorRecipients_() {

  const configured =
    APP_CONFIG.NOTIFICATIONS &&
    Array.isArray(
      APP_CONFIG.NOTIFICATIONS
        .COORDINATOR_RECIPIENTS
    )

      ? APP_CONFIG.NOTIFICATIONS
          .COORDINATOR_RECIPIENTS

      : [];


  return configured

    .map(value =>
      String(
        value || ''
      ).trim()
    )

    .filter(Boolean);

}


/**
 * =====================================================================
 * DASHBOARD URL
 * =====================================================================
 *
 * We deliberately do not invent the Web App URL.
 *
 * When APP_CONFIG.NOTIFICATIONS.DASHBOARD_URL is populated,
 * this automatically starts generating Review buttons.
 *
 * Parameters are included so DashboardUI can eventually open directly
 * to the correct student/submission.
 */
function notification_buildDashboardUrl_(
  studentId,
  submissionId
) {

  const baseUrl =
    APP_CONFIG.NOTIFICATIONS
      ? String(
          APP_CONFIG.NOTIFICATIONS
            .DASHBOARD_URL ||
          ''
        ).trim()
      : '';


  if (!baseUrl) {
    return '';
  }


  const params = [];


  if (
    studentId
  ) {

    params.push(
      'studentId=' +
      encodeURIComponent(
        studentId
      )
    );
  }


  if (
    submissionId
  ) {

    params.push(
      'submissionId=' +
      encodeURIComponent(
        submissionId
      )
    );
  }


  if (
    params.length === 0
  ) {

    return baseUrl;
  }


  return (
    baseUrl +
    (
      baseUrl.includes('?')
        ? '&'
        : '?'
    ) +
    params.join('&')
  );
}


/**
 * =====================================================================
 * HTML TABLE ROW
 * =====================================================================
 */
function notification_row_(
  label,
  value
) {

  return (

    '<tr>' +

      '<td style="' +
        'padding:8px;' +
        'border-bottom:1px solid #dddddd;' +
        'font-weight:bold;' +
        'width:180px;' +
      '">' +

        notification_escape_(
          label
        ) +

      '</td>' +

      '<td style="' +
        'padding:8px;' +
        'border-bottom:1px solid #dddddd;' +
      '">' +

        notification_escape_(
          value
        ) +

      '</td>' +

    '</tr>'

  );
}


/**
 * =====================================================================
 * OPTIONAL HTML ROW
 * =====================================================================
 */
function notification_optionalRow_(
  label,
  value
) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {

    return '';
  }


  return notification_row_(
    label,
    value
  );
}


/**
 * =====================================================================
 * BUTTON
 * =====================================================================
 */
function notification_button_(
  label,
  url
) {

  if (!url) {
    return '';
  }


  return (

    '<p style="' +
      'margin:22px 10px 22px 0;' +
      'display:inline-block;' +
    '">' +

      '<a href="' +
        notification_escape_(
          url
        ) +
      '" style="' +

        'display:inline-block;' +
        'padding:10px 18px;' +
        'background:#1f4e79;' +
        'color:#ffffff;' +
        'text-decoration:none;' +
        'font-weight:bold;' +
        'border-radius:4px;' +

      '">' +

        notification_escape_(
          label
        ) +

      '</a>' +

    '</p>'

  );
}


/**
 * =====================================================================
 * RESULT BADGE
 * =====================================================================
 */
function notification_resultBadge_(
  result
) {

  const normalized =
    String(
      result || ''
    )
      .trim()
      .toUpperCase();


  let background =
    '#eeeeee';


  let color =
    '#202124';


  if (
    normalized === 'PASS'
  ) {

    background =
      '#e6f4ea';

    color =
      '#188038';
  }


  if (
    normalized === 'FAIL'
  ) {

    background =
      '#fce8e6';

    color =
      '#b3261e';
  }


  return (

    '<span style="' +

      'display:inline-block;' +
      'padding:4px 10px;' +

      'background:' +
      background +
      ';' +

      'color:' +
      color +
      ';' +

      'font-weight:bold;' +
      'border-radius:4px;' +

    '">' +

      notification_escape_(
        normalized
      ) +

    '</span>'

  );
}


/**
 * =====================================================================
 * DATE DISPLAY
 * =====================================================================
 */
function notification_formatDate_(
  value
) {

  const date =
    appDateOnly_(
      value
    );


  if (!date) {

    return String(
      value || ''
    ).trim();
  }


  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'MMMM d, yyyy'
  );
}


/**
 * =====================================================================
 * FIRST NON-BLANK VALUE
 * =====================================================================
 */
function notification_firstValue_(
  record,
  keys
) {

  for (
    let i = 0;
    i < keys.length;
    i++
  ) {

    const key =
      keys[i];


    if (
      !Object.prototype
        .hasOwnProperty
        .call(
          record,
          key
        )
    ) {

      continue;
    }


    const value =
      record[key];


    if (
      value !== null &&
      value !== undefined &&
      String(value).trim() !== ''
    ) {

      return value;
    }
  }


  return '';
}


/**
 * =====================================================================
 * FIRST VALUE — PRESERVE ZERO
 * =====================================================================
 */
function notification_firstValuePreserveZero_(
  record,
  keys
) {

  for (
    let i = 0;
    i < keys.length;
    i++
  ) {

    const key =
      keys[i];


    if (
      !Object.prototype
        .hasOwnProperty
        .call(
          record,
          key
        )
    ) {

      continue;
    }


    const value =
      record[key];


    if (
      value === 0
    ) {
      return 0;
    }


    if (
      value !== null &&
      value !== undefined &&
      String(value).trim() !== ''
    ) {

      return value;
    }
  }


  return '';
}


/**
 * =====================================================================
 * YES / TRUE
 * =====================================================================
 */
function notification_isYes_(
  value
) {

  const normalized =
    appNormalize_(
      value
    );


  return (
    normalized === 'yes' ||
    normalized === 'true' ||
    normalized === 'y' ||
    normalized === '1'
  );
}


/**
 * =====================================================================
 * HTML ESCAPE
 * =====================================================================
 */
function notification_escape_(
  value
) {

  return String(
    value === null ||
    value === undefined
      ? ''
      : value
  )
    .replace(
      /&/g,
      '&amp;'
    )
    .replace(
      /</g,
      '&lt;'
    )
    .replace(
      />/g,
      '&gt;'
    )
    .replace(
      /"/g,
      '&quot;'
    )
    .replace(
      /'/g,
      '&#39;'
    );
}
