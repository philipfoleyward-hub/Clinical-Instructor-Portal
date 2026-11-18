/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * CLINICAL REVIEW SERVICE
 * 04_ClinicalReviewService.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Owns coordinator review workflow for evaluations already stored in
 * Clinical Tracking.
 *
 * THIS SERVICE MAY CHANGE ONLY:
 *
 *   Coordinator Review Status
 *   Coordinator Flag
 *   Coordinator Notes
 *   Reviewed By
 *   Review Date
 *
 * THIS SERVICE DOES NOT:
 *
 *   - change grades
 *   - change PASS / FAIL
 *   - change student information
 *   - change evaluation information
 *   - change requirements
 *   - generate documents
 *   - send notifications
 *
 * Submission ID is the authoritative evaluation identifier.
 *
 * =====================================================================
 */


/**
 * =====================================================================
 * CONFIGURATION
 * =====================================================================
 */

const CLINICAL_REVIEW = {

  STATUS_PENDING:
    'Pending',

  STATUS_REVIEWED:
    'Reviewed',

  FLAG_YES:
    'Yes',

  FLAG_NO:
    'No'

};


/**
 * =====================================================================
 * PUBLIC: MARK EVALUATION REVIEWED
 * =====================================================================
 *
 * Sets:
 *
 *   Coordinator Review Status = Reviewed
 *   Reviewed By               = current user when available
 *   Review Date               = current date/time
 *
 * Does NOT automatically remove a flag.
 */
function clinicalReview_markReviewed(
  submissionId
) {

  const id =
    clinicalReview_clean_(
      submissionId
    );


  if (!id) {

    throw new Error(
      'clinicalReview_markReviewed requires a Submission ID.'
    );
  }


  const context =
    clinicalReview_findRecord_(
      id
    );


  const reviewedBy =
    clinicalReview_getCurrentUser_();


  const reviewDate =
    new Date();


  clinicalReview_setCell_(
    context,
    'Coordinator Review Status',
    CLINICAL_REVIEW.STATUS_REVIEWED
  );


  clinicalReview_setCell_(
    context,
    'Reviewed By',
    reviewedBy
  );


  clinicalReview_setCell_(
    context,
    'Review Date',
    reviewDate
  );


  SpreadsheetApp.flush();


  console.log(
    'Marked reviewed: ' +
    id
  );


  return {

    success:
      true,

    submissionId:
      id,

    status:
      CLINICAL_REVIEW.STATUS_REVIEWED,

    reviewedBy:
      reviewedBy,

    reviewDate:
      reviewDate,

    flagged:
      clinicalReview_isFlaggedValue_(
        clinicalReview_getCellValue_(
          context,
          'Coordinator Flag'
        )
      )

  };
}


/**
 * =====================================================================
 * PUBLIC: SET FLAG
 * =====================================================================
 *
 * flagged = true
 *     Coordinator Flag = Yes
 *
 * flagged = false
 *     Coordinator Flag = No
 *
 * Flagging does NOT alter review status.
 */
function clinicalReview_setFlag(
  submissionId,
  flagged
) {

  const id =
    clinicalReview_clean_(
      submissionId
    );


  if (!id) {

    throw new Error(
      'clinicalReview_setFlag requires a Submission ID.'
    );
  }


  const context =
    clinicalReview_findRecord_(
      id
    );


  const shouldFlag =
    clinicalReview_toBoolean_(
      flagged
    );


  const value =
    shouldFlag
      ? CLINICAL_REVIEW.FLAG_YES
      : CLINICAL_REVIEW.FLAG_NO;


  clinicalReview_setCell_(
    context,
    'Coordinator Flag',
    value
  );


  SpreadsheetApp.flush();


  console.log(
    (
      shouldFlag
        ? 'Flagged: '
        : 'Unflagged: '
    ) +
    id
  );


  return {

    success:
      true,

    submissionId:
      id,

    flagged:
      shouldFlag,

    coordinatorFlag:
      value

  };
}


/**
 * =====================================================================
 * PUBLIC: GET REVIEW STATE
 * =====================================================================
 *
 * Read-only.
 *
 * Useful to Dashboard.gs when it needs the current state of one
 * evaluation.
 */
function clinicalReview_getState(
  submissionId
) {

  const id =
    clinicalReview_clean_(
      submissionId
    );


  if (!id) {

    throw new Error(
      'clinicalReview_getState requires a Submission ID.'
    );
  }


  const context =
    clinicalReview_findRecord_(
      id
    );


  const status =
    clinicalReview_clean_(
      clinicalReview_getCellValue_(
        context,
        'Coordinator Review Status'
      )
    );


  const flagValue =
    clinicalReview_getCellValue_(
      context,
      'Coordinator Flag'
    );


  return {

    success:
      true,

    submissionId:
      id,

    status:
      status ||
      CLINICAL_REVIEW.STATUS_PENDING,

    flagged:
      clinicalReview_isFlaggedValue_(
        flagValue
      ),

    notes:
      clinicalReview_clean_(
        clinicalReview_getCellValue_(
          context,
          'Coordinator Notes'
        )
      ),

    reviewedBy:
      clinicalReview_clean_(
        clinicalReview_getCellValue_(
          context,
          'Reviewed By'
        )
      ),

    reviewDate:
      clinicalReview_getCellValue_(
        context,
        'Review Date'
      )

  };
}


/**
 * =====================================================================
 * PUBLIC: SAVE COORDINATOR NOTES
 * =====================================================================
 *
 * We are NOT putting this into the Dashboard yet.
 *
 * The column already exists, so this service owns it now and the
 * Dashboard can use this later without another architectural change.
 *
 * Saving a note does NOT mark an evaluation reviewed.
 */
function clinicalReview_saveNotes(
  submissionId,
  notes
) {

  const id =
    clinicalReview_clean_(
      submissionId
    );


  if (!id) {

    throw new Error(
      'clinicalReview_saveNotes requires a Submission ID.'
    );
  }


  const context =
    clinicalReview_findRecord_(
      id
    );


  clinicalReview_setCell_(
    context,
    'Coordinator Notes',
    clinicalReview_clean_(
      notes
    )
  );


  SpreadsheetApp.flush();


  return {

    success:
      true,

    submissionId:
      id,

    notes:
      clinicalReview_clean_(
        notes
      )

  };
}


/**
 * =====================================================================
 * INTERNAL: FIND CLINICAL TRACKING RECORD
 * =====================================================================
 *
 * Submission ID must resolve to exactly one Clinical Tracking row.
 *
 * We intentionally refuse to guess if duplicate IDs exist.
 */
function clinicalReview_findRecord_(
  submissionId
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.CLINICAL_TRACKING
    );


  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {

    throw new Error(
      'Clinical Tracking contains no evaluation records.'
    );
  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getValues()[0]
      .map(
        function(header) {

          return clinicalReview_clean_(
            header
          );
        }
      );


  clinicalReview_validateHeaders_(
    headers
  );


  const submissionColumn =
    headers.indexOf(
      'Submission ID'
    );


  const values =
    sheet
      .getRange(
        2,
        submissionColumn + 1,
        lastRow - 1,
        1
      )
      .getValues();


  const matches = [];


  values.forEach(
    function(row, index) {

      if (
        clinicalReview_clean_(
          row[0]
        ) ===
        submissionId
      ) {

        matches.push(
          index + 2
        );
      }
    }
  );


  if (
    matches.length === 0
  ) {

    throw new Error(
      'Submission ID not found in Clinical Tracking: ' +
      submissionId
    );
  }


  if (
    matches.length > 1
  ) {

    throw new Error(
      'Duplicate Submission ID found in Clinical Tracking: ' +
      submissionId +
      '. Review service will not modify an ambiguous record.'
    );
  }


  return {

    sheet:
      sheet,

    headers:
      headers,

    row:
      matches[0]

  };
}


/**
 * =====================================================================
 * INTERNAL: VALIDATE REQUIRED HEADERS
 * =====================================================================
 */
function clinicalReview_validateHeaders_(
  headers
) {

  const required = [

    'Submission ID',
    'Coordinator Review Status',
    'Coordinator Flag',
    'Coordinator Notes',
    'Reviewed By',
    'Review Date'

  ];


  const missing =
    required.filter(
      function(header) {

        return (
          headers.indexOf(
            header
          ) === -1
        );
      }
    );


  if (
    missing.length > 0
  ) {

    throw new Error(
      'Clinical Tracking is missing required review column(s): ' +
      missing.join(', ')
    );
  }
}


/**
 * =====================================================================
 * INTERNAL: WRITE ONE REVIEW CELL
 * =====================================================================
 */
function clinicalReview_setCell_(
  context,
  header,
  value
) {

  const columnIndex =
    context.headers.indexOf(
      header
    );


  if (
    columnIndex === -1
  ) {

    throw new Error(
      'Clinical Tracking column not found: ' +
      header
    );
  }


  context.sheet
    .getRange(
      context.row,
      columnIndex + 1
    )
    .setValue(
      value
    );
}


/**
 * =====================================================================
 * INTERNAL: READ ONE REVIEW CELL
 * =====================================================================
 */
function clinicalReview_getCellValue_(
  context,
  header
) {

  const columnIndex =
    context.headers.indexOf(
      header
    );


  if (
    columnIndex === -1
  ) {

    throw new Error(
      'Clinical Tracking column not found: ' +
      header
    );
  }


  return context.sheet
    .getRange(
      context.row,
      columnIndex + 1
    )
    .getValue();
}


/**
 * =====================================================================
 * INTERNAL: CURRENT USER
 * =====================================================================
 *
 * Google Workspace sometimes allows ActiveUser email and sometimes
 * intentionally returns blank depending on deployment/security context.
 *
 * Blank is acceptable. Review Date and Review Status are still recorded.
 */
function clinicalReview_getCurrentUser_() {

  try {

    const email =
      Session
        .getActiveUser()
        .getEmail();


    if (email) {

      return String(
        email
      ).trim();
    }

  } catch (error) {

    console.log(
      'Active reviewer email unavailable: ' +
      error.message
    );
  }


  return '';
}


/**
 * =====================================================================
 * INTERNAL: FLAG VALUE
 * =====================================================================
 */
function clinicalReview_isFlaggedValue_(
  value
) {

  const normalized =
    String(
      value === null ||
      value === undefined
        ? ''
        : value
    )
      .trim()
      .toLowerCase();


  return (
    normalized === 'yes' ||
    normalized === 'true' ||
    normalized === '1' ||
    normalized === 'flagged'
  );
}


/**
 * =====================================================================
 * INTERNAL: BOOLEAN
 * =====================================================================
 */
function clinicalReview_toBoolean_(
  value
) {

  if (
    value === true ||
    value === 1
  ) {

    return true;
  }


  const normalized =
    String(
      value === null ||
      value === undefined
        ? ''
        : value
    )
      .trim()
      .toLowerCase();


  return (
    normalized === 'true' ||
    normalized === 'yes' ||
    normalized === '1' ||
    normalized === 'flagged'
  );
}


/**
 * =====================================================================
 * INTERNAL: CLEAN TEXT
 * =====================================================================
 */
function clinicalReview_clean_(
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
 * TEST HELPERS
 * =====================================================================
 *
 * These deliberately DO NOT automatically change anything.
 *
 * Put a Submission ID below temporarily when testing.
 */


/**
 * READ-ONLY TEST
 *
 * Safe to run after inserting a real Submission ID.
 */
function testClinicalReview_getState() {

  const submissionId =
    'QR-2_ABaOnud-k46nCSZj4BHl_F4J0hTmagZWjlRgDT4JGjCAkZb9V1RNWj4q4TYl1LNSeIbC1hc';


  if (
    submissionId ===
    'PUT-SUBMISSION-ID-HERE'
  ) {

    throw new Error(
      'Insert a real Submission ID in testClinicalReview_getState first.'
    );
  }


  const result =
    clinicalReview_getState(
      submissionId
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
 * FLAG TEST
 *
 * WARNING: writes Coordinator Flag = Yes.
 */
function testClinicalReview_flag() {

  const submissionId =
    'QR-2_ABaOnud-k46nCSZj4BHl_F4J0hTmagZWjlRgDT4JGjCAkZb9V1RNWj4q4TYl1LNSeIbC1hc';


  if (
    submissionId ===
    'PUT-SUBMISSION-ID-HERE'
  ) {

    throw new Error(
      'Insert a real Submission ID in testClinicalReview_flag first.'
    );
  }


  return clinicalReview_setFlag(
    submissionId,
    true
  );
}


/**
 * UNFLAG TEST
 *
 * WARNING: writes Coordinator Flag = No.
 */
function testClinicalReview_unflag() {

  const submissionId =
    'PUT-SUBMISSION-ID-HERE';


  if (
    submissionId ===
    'PUT-SUBMISSION-ID-HERE'
  ) {

    throw new Error(
      'Insert a real Submission ID in testClinicalReview_unflag first.'
    );
  }


  return clinicalReview_setFlag(
    submissionId,
    false
  );
}


/**
 * MARK REVIEWED TEST
 *
 * WARNING:
 *
 * Writes:
 *   Coordinator Review Status = Reviewed
 *   Reviewed By = current user if available
 *   Review Date = now
 */
function testClinicalReview_markReviewed() {

  const submissionId =
    'QR-2_ABaOnud-k46nCSZj4BHl_F4J0hTmagZWjlRgDT4JGjCAkZb9V1RNWj4q4TYl1LNSeIbC1hc';


  if (
    submissionId ===
    'PUT-SUBMISSION-ID-HERE'
  ) {

    throw new Error(
      'Insert a real Submission ID in testClinicalReview_markReviewed first.'
    );
  }


  return clinicalReview_markReviewed(
    submissionId
  );
}