/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * QR / UPLOADED EVALUATION FORM SYNC
 * 02_QRSubmissionFormSync.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Keeps the QR / uploaded Clinical Evaluation Submission Form
 * synchronized with authoritative data in the Master Workbook.
 *
 * THIS FILE OWNS:
 *
 *   - Student dropdown
 *   - Clinical Site dropdown
 *   - College CI / evaluator dropdown
 *   - Examination dropdown
 *
 * THIS FILE DOES NOT:
 *
 *   - Process submitted evaluations
 *   - Move uploaded files
 *   - Write to Clinical Tracking
 *   - Calculate grades
 *   - Calculate requirements
 *   - Generate documents
 *   - Send emails
 *
 * IMPORTANT
 * ---------
 * Current form design uses one Student dropdown containing all
 * ACTIVE students.
 *
 * Student-by-site filtering is NOT implemented here yet.
 * If we decide the QR form needs that behavior, it will be added
 * deliberately to this same synchronization subsystem.
 * =====================================================================
 */


/**
 * =====================================================================
 * MAIN SYNC
 * =====================================================================
 *
 * Run this when:
 *
 *   - a student is added/activated/deactivated
 *   - a clinical site is added/activated/deactivated
 *   - a College CI is added/activated/deactivated
 *   - an examination is added/activated/deactivated
 */
function syncQRSubmissionForm() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const form =
    FormApp.openById(
      APP_CONFIG.FORMS
        .QR_SUBMISSION
    );


  console.log(
    'Starting QR Submission Form sync...'
  );


  // ===================================================================
  // LOAD AUTHORITATIVE MASTER DATA
  // ===================================================================

  const students =
    qrFormSync_getActiveStudents_(ss);


  const clinicalSites =
    qrFormSync_getActiveClinicalSites_(ss);


  const evaluators =
    qrFormSync_getCollegeCIs_(ss);


  const examinations =
    qrFormSync_getActiveExaminations_(ss);


  // ===================================================================
  // SAFETY
  // ===================================================================

  if (students.length === 0) {

    throw new Error(
      'No active students were found.'
    );
  }


  if (clinicalSites.length === 0) {

    throw new Error(
      'No active clinical sites were found.'
    );
  }


  if (evaluators.length === 0) {

    throw new Error(
      'No active College CIs were found.'
    );
  }


  if (examinations.length === 0) {

    throw new Error(
      'No active examinations were found.'
    );
  }


  // ===================================================================
  // UPDATE FORM
  // ===================================================================

  qrFormSync_setDropdown_(
    form,
    'Student',
    students
  );


  qrFormSync_setDropdown_(
    form,
    'Clinical Site',
    clinicalSites
  );


  /*
   * The existing form question is titled "Evaluators".
   *
   * We preserve the existing question title rather than changing
   * the Form structure during this cleanup.
   */
  qrFormSync_setDropdown_(
    form,
    'Evaluators',
    evaluators
  );


  qrFormSync_setDropdown_(
    form,
    'Examination',
    examinations
  );


  // ===================================================================
  // RESULTS
  // ===================================================================

  console.log(
    '================================'
  );

  console.log(
    'QR SUBMISSION FORM SYNC COMPLETE'
  );

  console.log(
    'Students: ' +
    students.length
  );

  console.log(
    'Clinical Sites: ' +
    clinicalSites.length
  );

  console.log(
    'College CIs: ' +
    evaluators.length
  );

  console.log(
    'Examinations: ' +
    examinations.length
  );

  console.log(
    '================================'
  );


  return {

    success: true,

    students:
      students.length,

    clinicalSites:
      clinicalSites.length,

    evaluators:
      evaluators.length,

    examinations:
      examinations.length

  };
}


/**
 * =====================================================================
 * ACTIVE STUDENTS
 * =====================================================================
 *
 * Authoritative source:
 *
 *   Students
 *
 * Required:
 *
 *   Student
 *   Active
 *
 * NOTE:
 * Clinical Site is intentionally NOT used for filtering here yet.
 */
function qrFormSync_getActiveStudents_(ss) {

  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENTS
    );


  const rows =
    appSheetToObjects_(sheet);


  const students =
    rows

      .filter(row =>

        appIsTrue_(
          row['Active']
        )

      )

      .map(row =>

        String(
          row['Student'] || ''
        ).trim()

      )

      .filter(Boolean);


  return qrFormSync_uniqueSort_(
    students
  );
}


/**
 * =====================================================================
 * ACTIVE CLINICAL SITES
 * =====================================================================
 *
 * Authoritative source:
 *
 *   Clinical Sites
 *
 * Accepts site name column:
 *
 *   Clinical Site
 *   Site
 *   Clinical Site Name
 *
 * If Active exists, it is honored.
 * If Active does not exist, listed sites are included.
 */
function qrFormSync_getActiveClinicalSites_(ss) {

  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS
        .CLINICAL_SITES
    );


  const rows =
    appSheetToObjects_(sheet);


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(header =>
        String(
          header || ''
        ).trim()
      );


  const hasActive =
    headers.includes(
      'Active'
    );


  const sites =
    rows

      .filter(row => {

        if (
          hasActive &&
          !appIsTrue_(
            row['Active']
          )
        ) {

          return false;
        }


        return true;
      })

      .map(row =>

        String(
          row['Clinical Site'] ||
          row['Site'] ||
          row['Clinical Site Name'] ||
          ''
        ).trim()

      )

      .filter(Boolean);


  return qrFormSync_uniqueSort_(
    sites
  );
}


/**
 * =====================================================================
 * COLLEGE CI / EVALUATOR LIST
 * =====================================================================
 *
 * Authoritative source:
 *
 *   Evaluators
 *
 * Current QR form behavior:
 *
 *   College CIs appear in the dropdown.
 *   Hospital CIs / Clinical Site Instructors are NOT added here
 *   because that route allows them to be entered manually.
 */
function qrFormSync_getCollegeCIs_(ss) {

  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.EVALUATORS
    );


  const rows =
    appSheetToObjects_(sheet);


  const evaluators =
    rows

      .filter(row => {

        if (
          !appIsTrue_(
            row['Active']
          )
        ) {

          return false;
        }


        const evaluator =
          String(
            row['Evaluator'] || ''
          ).trim();


        if (!evaluator) {
          return false;
        }


        const evaluatorType =
          appNormalize_(
            row['Evaluator Type']
          );


        /*
         * Preserve existing behavior:
         *
         * Anything explicitly identified as Hospital CI or
         * Clinical Site Instructor is excluded from this dropdown.
         */
        if (
          evaluatorType.includes(
            'hospital'
          ) ||
          evaluatorType.includes(
            'clinical site'
          )
        ) {

          return false;
        }


        return true;
      })

      .map(row =>

        String(
          row['Evaluator'] || ''
        ).trim()

      );


  return qrFormSync_uniqueSort_(
    evaluators
  );
}


/**
 * =====================================================================
 * ACTIVE EXAMINATIONS
 * =====================================================================
 *
 * Authoritative source:
 *
 *   Examinations
 *
 * Required:
 *
 *   Examination
 *
 * Active is honored when present.
 */
function qrFormSync_getActiveExaminations_(ss) {

  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS
        .EXAMINATIONS
    );


  const rows =
    appSheetToObjects_(sheet);


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(header =>
        String(
          header || ''
        ).trim()
      );


  const hasActive =
    headers.includes(
      'Active'
    );


  const examinations =
    rows

      .filter(row => {

        if (
          hasActive &&
          !appIsTrue_(
            row['Active']
          )
        ) {

          return false;
        }


        return Boolean(
          String(
            row['Examination'] || ''
          ).trim()
        );
      })

      .map(row =>

        String(
          row['Examination'] || ''
        ).trim()

      );


  return qrFormSync_uniqueSort_(
    examinations
  );
}


/**
 * =====================================================================
 * UPDATE ONE DROPDOWN
 * =====================================================================
 */
function qrFormSync_setDropdown_(
  form,
  title,
  choices
) {

  if (
    !Array.isArray(choices) ||
    choices.length === 0
  ) {

    throw new Error(
      'No choices available for QR Form question: "' +
      title +
      '".'
    );
  }


  const target =
    appNormalize_(title);


  const items =
    form.getItems(
      FormApp.ItemType.LIST
    );


  for (
    let i = 0;
    i < items.length;
    i++
  ) {

    const item =
      items[i]
        .asListItem();


    if (
      appNormalize_(
        item.getTitle()
      ) === target
    ) {

      item.setChoiceValues(
        choices
      );


      console.log(
        'Updated "' +
        title +
        '" with ' +
        choices.length +
        ' choices.'
      );


      return;
    }
  }


  throw new Error(
    'Could not find QR Form dropdown titled "' +
    title +
    '".'
  );
}


/**
 * =====================================================================
 * UNIQUE + ALPHABETICAL SORT
 * =====================================================================
 */
function qrFormSync_uniqueSort_(
  values
) {

  const unique = {};


  (values || [])
    .forEach(value => {

      const text =
        String(
          value || ''
        ).trim();


      if (!text) {
        return;
      }


      const key =
        appNormalize_(text);


      if (!unique[key]) {

        unique[key] =
          text;
      }
    });


  return Object
    .values(unique)
    .sort(
      (a, b) =>
        a.localeCompare(b)
    );
}
