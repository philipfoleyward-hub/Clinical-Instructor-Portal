/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * Diagnostic timing instrumentation added for dashboard generation tracing.
 * CLINICAL GRADES DOCUMENT GENERATOR
 * 05_ClinicalGradesDocumentGenerator.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Creates a student-specific copy of the official Clinical Grades
 * Google Sheet and populates all currently authoritative grade data.
 *
 * POPULATES
 * ---------
 * - Student name / student number
 * - Clinical 1 through Clinical 5 competency evaluations
 * - Instructor, percent grade, and date
 * - ISO evaluations
 * - Failed competency evaluations
 * - Competency average (numeric passing + numeric failed competency grades)
 * - ISO average
 *
 * Professional Competency detail rows come from active Grade Records.
 * Category and semester averages come only from gradeEngine_getStudentSummary().
 *
 * DEPENDENCIES
 * ------------
 * - APP_CONFIG
 * - appGetSheet_
 * - appSheetToObjects_
 * - appNormalize_
 * - appIsTrue_
 * - appDateOnly_
 * - appDateWithin_
 * - appDateValue_
 * - driveService_copyTemplateToStudentFolder
 * - gradeEngine_formatEvaluationGrade
 *
 * =====================================================================
 */


/**
 * =====================================================================
 * TEMPLATE CONFIGURATION
 * =====================================================================
 */
const CLINICAL_GRADES_TEMPLATE_ID =
  '1OpzLEf4G-OmBLaAnL9W_UIGBHX5P308JoRMr2QOJCWM';


/**
 * Read-only validation of the configured Clinical Grades master template.
 * Opens the master directly and verifies all five PC average tokens and all
 * fifty PC detail tokens without creating a student document.
 */
function clinicalGrades_preflightTemplate() {

  const templateFile =
    DriveApp.getFileById(
      CLINICAL_GRADES_TEMPLATE_ID
    );

  if (templateFile.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    throw new Error('Clinical Grades template must be a native Google Sheet.');
  }

  const spreadsheet =
    SpreadsheetApp.openById(
      CLINICAL_GRADES_TEMPLATE_ID
    );

  const sheet =
    clinicalGrades_getMainSheet_(
      spreadsheet
    );

  const averageTokensFound = [];
  const detailTokensFound = [];
  const missingTokens = [];

  const tokenExists = token =>
    Boolean(
      sheet
        .createTextFinder(token)
        .matchEntireCell(true)
        .matchCase(true)
        .useRegularExpression(false)
        .findNext()
    );

  for (let clinicalNumber = 1; clinicalNumber <= 5; clinicalNumber++) {
    const averageToken =
      '{{C' + clinicalNumber + '_PROF_AVG}}';

    if (tokenExists(averageToken)) {
      averageTokensFound.push(averageToken);
    }

    for (let index = 1; index <= 5; index++) {
      [
        '{{C' + clinicalNumber + 'PC' + index + 'I}}',
        '{{C' + clinicalNumber + 'PC' + index + 'G}}'
      ].forEach(token => {
        if (tokenExists(token)) {
          detailTokensFound.push(token);
        } else {
          missingTokens.push(token);
        }
      });
    }
  }

  const result = {
    success: missingTokens.length === 0,
    readOnly: true,
    templateId: CLINICAL_GRADES_TEMPLATE_ID,
    templateName: templateFile.getName(),
    sheetName: sheet.getName(),
    expectedAverageTokenCount: 0,
    optionalAverageTokenCount: 5,
    averageTokenCount: averageTokensFound.length,
    averageTokensFound: averageTokensFound,
    expectedDetailTokenCount: 50,
    detailTokenCount: detailTokensFound.length,
    detailTokensFound: detailTokensFound,
    missingTokens: missingTokens,
    documentsCreated: 0,
    recordsWrittenOrChanged: 0
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


/**
 * =====================================================================
 * PUBLIC GENERATOR
 * =====================================================================
 */
function generateClinicalGradesSheet(
  studentId,
  diagnosticRequestId
) {

  diagnosticRequestId =
    String(
      diagnosticRequestId ||
      'DIRECT-RUN'
    ).trim();

  const gradesDiagnosticStartedAt_ =
    Date.now();

  const gradesTrace_ =
    function(stage, details) {

      console.log(
        '[Clinical Grades][' +
        diagnosticRequestId +
        '] +' +
        (Date.now() - gradesDiagnosticStartedAt_) +
        'ms] ' +
        stage +
        (details ? ' | ' + details : '')
      );
    };

  gradesTrace_(
    'START Grades generation',
    'Student ID: ' + studentId
  );

  if (!studentId) {
    throw new Error(
      'Student ID is required.'
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

  const tracking =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.CLINICAL_TRACKING
      )
    );

  const semesterRules =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.SEMESTER_RULES
      )
    );

  gradesTrace_(
    'Source grade data loaded',
    'Students: ' + students.length +
    ', Tracking: ' + tracking.length +
    ', Semester Rules: ' + semesterRules.length
  );


  // ===================================================================
  // STUDENT
  // ===================================================================

  const studentRows =
    students.filter(row =>
      appNormalize_(
        row['Student ID']
      ) ===
      appNormalize_(
        studentId
      )
    );

  if (studentRows.length === 0) {
    throw new Error(
      'Student not found: ' +
      studentId
    );
  }

  const student =
    clinicalGrades_selectStudentRow_(
      studentRows
    );

  const studentName =
    String(
      student['Student'] || ''
    ).trim();

  const cohort =
    String(
      student['Cohort'] || ''
    ).trim();

  if (!studentName) {
    throw new Error(
      'Student name is missing for Student ID: ' +
      studentId
    );
  }

  if (!cohort) {
    throw new Error(
      'Student cohort is missing for: ' +
      studentName
    );
  }

  const displayStudentId =
    clinicalGrades_formatStudentId_(
      studentId
    );

  gradesTrace_(
    'Student resolved',
    studentName + ' | Cohort: ' + cohort
  );

  // Authoritative read-only semester/category calculation. Record creation,
  // evaluation scoring, and Grade Records writes occur upstream and are not
  // part of this document consumer.
  const gradeSummary =
    gradeEngine_getStudentSummary(
      studentId
    );

  const gradeTerms = {};

  (gradeSummary.terms || [])
    .forEach(term => {
      gradeTerms[
        appNormalize_(
          term.clinicalTerm
        )
      ] = term;
    });

  gradesTrace_(
    'Authoritative Grade Engine summary loaded',
    'Terms: ' + Object.keys(gradeTerms).length
  );


  // ===================================================================
  // STUDENT TRACKING HISTORY
  // ===================================================================

  const records =
    tracking.filter(record =>
      clinicalGrades_sameStudent_(
        record,
        studentId,
        studentName
      )
    );

  const supplementalRecords =
    gradeEngine_getSupplementalGradeRecords_(
      ss
    )
      .filter(record =>
        gradeEngine_sameStudent_(
          record,
          studentId,
          studentName
        )
      );

  gradesTrace_(
    'Student source grade data selected',
    'Matching tracking records: ' + records.length
  );


  // ===================================================================
  // VERIFY TEMPLATE
  // ===================================================================

  const template =
    DriveApp.getFileById(
      CLINICAL_GRADES_TEMPLATE_ID
    );

  if (
    template.getMimeType() !==
    MimeType.GOOGLE_SHEETS
  ) {
    throw new Error(
      'Clinical Grades template must be a native Google Sheet.'
    );
  }

  gradesTrace_(
    'Grades template opened',
    'Template ID: ' + CLINICAL_GRADES_TEMPLATE_ID
  );


  // ===================================================================
  // CREATE STUDENT COPY
  // ===================================================================

  const outputName =
    studentName +
    ' - Clinical Grades';

  const copyResult =
    driveService_copyTemplateToStudentFolder({

      templateId:
        CLINICAL_GRADES_TEMPLATE_ID,

      fileName:
        outputName,

      student:
        studentName,

      cohort:
        cohort,

      destination:
        'documents'

    });

  gradesTrace_(
    'Student Drive folder resolved and Grades template copied',
    'Generated file ID: ' + copyResult.fileId
  );


  // ===================================================================
  // OPEN GENERATED SHEET
  // ===================================================================

  const generated =
    clinicalGrades_openSpreadsheetWithRetry_(
      copyResult.fileId
    );

  const sheet =
    clinicalGrades_getMainSheet_(
      generated
    );

  gradesTrace_(
    'Generated Grades spreadsheet opened',
    'Sheet: ' + sheet.getName()
  );


  // ===================================================================
  // CAPTURE TOKEN LOCATIONS BEFORE REPLACEMENT
  // ===================================================================

  const tokenLayout =
    clinicalGrades_scanTemplate_(
      sheet
    );

  clinicalGrades_assertProfessionalCompetencyTokens_(
    sheet
  );

  gradesTrace_(
    'Token payload/layout built',
    'Clinical term layouts: ' + Object.keys(tokenLayout).length
  );


  // ===================================================================
  // STUDENT HEADER
  // ===================================================================

  gradesTrace_(
    'Token replacement started',
    'Student header and Clinical 1-5 records'
  );

  clinicalGrades_replaceEmbeddedToken_(
    sheet,
    '{{STUDENT_NAME}}',
    studentName
  );

  clinicalGrades_replaceEmbeddedToken_(
    sheet,
    '{{STUDENT_NO}}',
    displayStudentId
  );


  // ===================================================================
  // CLINICAL 1 - 5
  // ===================================================================

  const counts = {};

  for (
    let clinicalNumber = 1;
    clinicalNumber <= 5;
    clinicalNumber++
  ) {

    const clinicalTerm =
      'Clinical ' +
      clinicalNumber;

    gradesTrace_(
      'Token replacement term started',
      clinicalTerm
    );

    /*
     * Detailed source rows use the same effective-term rule as the Grade
     * Engine. Historical rows therefore remain in their reviewed/stored term;
     * current rows retain normal Attributed Clinical Term behavior.
     */
    const termRecords =
      records
        .filter(record =>
          appNormalize_(
            gradeEngine_effectiveClinicalTerm_(
              record
            )
          ) ===
          appNormalize_(
            clinicalTerm
          )
        )
        .sort(
          clinicalGrades_recordDateSort_
        );

    const passingCompetencies =
      termRecords.filter(
        clinicalGrades_isPassingCompetency_
      );

    const failedCompetencies =
      termRecords.filter(
        clinicalGrades_isFailedCompetency_
      );

    const isoRecords =
      termRecords.filter(record =>
        appNormalize_(
          record['Evaluation Type']
        ) ===
        'iso'
      );

    const termProfessionalRecords =
      gradeEngine_filterGradeRecordsByCategory_(
        supplementalRecords.filter(record =>
          appNormalize_(
            record['Clinical Term']
          ) ===
          appNormalize_(
            clinicalTerm
          )
        ),
        'professional'
      )
        .sort(
          clinicalGrades_recordDateSort_
        );

    const layout =
      tokenLayout[
        clinicalNumber
      ];

    clinicalGrades_populateCompetencies_(
      sheet,
      clinicalNumber,
      passingCompetencies,
      layout.competency
    );

    clinicalGrades_populateIso_(
      sheet,
      clinicalNumber,
      isoRecords,
      layout.iso
    );

    clinicalGrades_populateFailures_(
      sheet,
      clinicalNumber,
      failedCompetencies,
      layout.failure
    );

    const professionalCapacity =
      clinicalGrades_populateProfessionalCompetencies_(
        sheet,
        clinicalNumber,
        termProfessionalRecords,
        gradeTerms[
          appNormalize_(
            clinicalTerm
          )
        ] || null
      );

    counts[
      'clinical' +
      clinicalNumber
    ] = {

      competencies:
        passingCompetencies.length,

      iso:
        isoRecords.length,

      failures:
        failedCompetencies.length,

      professionalCompetencies:
        termProfessionalRecords.length,

      competencyOverflow:
        Math.max(
          passingCompetencies.length -
          layout.competency,
          0
        ),

      isoOverflow:
        Math.max(
          isoRecords.length -
          layout.iso,
          0
        ),

      failureOverflow:
        Math.max(
          failedCompetencies.length -
          layout.failure,
          0
        ),

      professionalCompetencyOverflow:
        Math.max(
          termProfessionalRecords.length -
          professionalCapacity,
          0
        )

    };

    gradesTrace_(
      'Token replacement term completed',
      clinicalTerm +
      ' | Competencies: ' + passingCompetencies.length +
      ', ISO: ' + isoRecords.length +
      ', Failures: ' + failedCompetencies.length +
      ', Professional Competencies: ' + termProfessionalRecords.length
    );
  }

  // Preserve the institutional fixed summary layout while also publishing
  // every legitimate source record on an unbounded detail sheet.
  clinicalGrades_writeCompleteDetailSheet_(
    generated,
    studentName,
    displayStudentId,
    records,
    supplementalRecords
  );


  // ===================================================================
  // CLEAR ANY UNUSED TOKENS
  // ===================================================================

  clinicalGrades_clearAllRemainingTokens_(
    sheet
  );

  gradesTrace_(
    'Token replacement completed',
    'Unused tokens cleared'
  );


  // ===================================================================
  // AUTHORITATIVE SEMESTER / CATEGORY VALUES
  //
  // Detailed rows above remain source listings. Every category and semester
  // total below comes from gradeEngine_getStudentSummary().
  // ===================================================================

  clinicalGrades_writeAuthoritativeSummaries_(
    sheet,
    gradeTerms
  );


  // ===================================================================
  // DATE FORMATTING
  // ===================================================================

  clinicalGrades_formatDates_(
    sheet,
    tokenLayout
  );


  SpreadsheetApp.flush();

  gradesTrace_(
    'Document saved/closed',
    'SpreadsheetApp.flush completed; Google Sheets has no explicit close operation'
  );

  gradesTrace_(
    'Generated file resolved',
    copyResult.fileId + ' | ' + copyResult.fileUrl
  );


  // ===================================================================
  // LOG / RETURN
  // ===================================================================

  console.log(
    '================================'
  );

  console.log(
    'CLINICAL GRADES SHEET CREATED'
  );

  console.log(
    'Student: ' +
    studentName
  );

  console.log(
    'Spreadsheet: ' +
    copyResult.fileUrl
  );

  console.log(
    JSON.stringify(
      counts,
      null,
      2
    )
  );

  console.log(
    '================================'
  );

  gradesTrace_(
    'Result returned to dashboard wrapper',
    outputName
  );

  return {

    success:
      true,

    name:
      outputName,

    url:
      copyResult.fileUrl,

    spreadsheetId:
      copyResult.fileId,

    counts:
      counts,

    gradeSummary:
      gradeSummary

  };
}


/**
 * =====================================================================
 * POPULATE COMPETENCIES
 * =====================================================================
 */
function clinicalGrades_populateCompetencies_(
  sheet,
  clinicalNumber,
  records,
  capacity
) {

  for (
    let i = 1;
    i <= capacity;
    i++
  ) {

    const record =
      records[
        i - 1
      ] || null;

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'E' +
        i +
        '}}',
      record
        ? clinicalGrades_displayExam_(
            record
          )
        : ''
    );

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'I' +
        i +
        '}}',
      record
        ? clinicalGrades_displayEvaluator_(
            record
          )
        : ''
    );

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'G' +
        i +
        '}}',
      record
        ? clinicalGrades_percentGrade_(
            record
          )
        : ''
    );

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'D' +
        i +
        '}}',
      record
        ? clinicalGrades_date_(
            record['Date']
          )
        : ''
    );
  }
}


/**
 * =====================================================================
 * POPULATE ISO
 * =====================================================================
 */
function clinicalGrades_populateIso_(
  sheet,
  clinicalNumber,
  records,
  capacity
) {

  for (
    let i = 1;
    i <= capacity;
    i++
  ) {

    const record =
      records[
        i - 1
      ] || null;

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'ISO' +
        i +
        'E}}',
      record
        ? clinicalGrades_displayExam_(
            record
          )
        : ''
    );

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'ISO' +
        i +
        'I}}',
      record
        ? clinicalGrades_displayEvaluator_(
            record
          )
        : ''
    );

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'ISO' +
        i +
        'G}}',
      record
        ? clinicalGrades_percentGrade_(
            record
          )
        : ''
    );

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'ISO' +
        i +
        'D}}',
      record
        ? clinicalGrades_date_(
            record['Date']
          )
        : ''
    );
  }
}


/**
 * =====================================================================
 * POPULATE FAILED COMPETENCIES
 * =====================================================================
 */
function clinicalGrades_populateFailures_(
  sheet,
  clinicalNumber,
  records,
  capacity
) {

  for (
    let i = 1;
    i <= capacity;
    i++
  ) {

    const record =
      records[
        i - 1
      ] || null;

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'F' +
        i +
        'E}}',
      record
        ? clinicalGrades_displayExam_(
            record
          )
        : ''
    );

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'F' +
        i +
        'I}}',
      record
        ? clinicalGrades_displayEvaluator_(
            record
          )
        : ''
    );

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'F' +
        i +
        'G}}',
      record
        ? clinicalGrades_percentGrade_(
            record
          )
        : ''
    );

    clinicalGrades_setExactToken_(
      sheet,
      '{{C' +
        clinicalNumber +
        'F' +
        i +
        'D}}',
      record
        ? clinicalGrades_date_(
            record['Date']
          )
        : ''
    );
  }
}


/**
 * =====================================================================
 * TEMPLATE SCAN
 * =====================================================================
 */
function clinicalGrades_scanTemplate_(
  sheet
) {

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  const result = {};

  for (
    let c = 1;
    c <= 5;
    c++
  ) {

    result[c] = {

      competency: 0,
      iso: 0,
      failure: 0,

      competencyGradeCells: [],
      isoGradeCells: [],
      failedGradeCells: [],
      dateCells: []

    };
  }

  for (
    let r = 0;
    r < values.length;
    r++
  ) {

    for (
      let col = 0;
      col < values[r].length;
      col++
    ) {

      const text =
        String(
          values[r][col] || ''
        ).trim();

      if (!text) {
        continue;
      }

      const a1 =
        sheet
          .getRange(
            r + 1,
            col + 1
          )
          .getA1Notation();

      let match;


      // ---------------------------------------------------------------
      // COMPETENCY: {{C1E1}} {{C1I1}} {{C1G1}} {{C1D1}}
      // ---------------------------------------------------------------

      match =
        text.match(
          /^\{\{C([1-5])([EIGD])(\d+)\}\}$/
        );

      if (match) {

        const clinical =
          Number(
            match[1]
          );

        const field =
          match[2];

        const index =
          Number(
            match[3]
          );

        result[clinical].competency =
          Math.max(
            result[clinical].competency,
            index
          );

        if (field === 'G') {
          result[clinical]
            .competencyGradeCells
            .push(
              a1
            );
        }

        if (field === 'D') {
          result[clinical]
            .dateCells
            .push(
              a1
            );
        }

        continue;
      }


      // ---------------------------------------------------------------
      // ISO: {{C1ISO1E}} {{C1ISO1I}} {{C1ISO1G}} {{C1ISO1D}}
      // ---------------------------------------------------------------

      match =
        text.match(
          /^\{\{C([1-5])ISO(\d+)([EIGD])\}\}$/
        );

      if (match) {

        const clinical =
          Number(
            match[1]
          );

        const index =
          Number(
            match[2]
          );

        const field =
          match[3];

        result[clinical].iso =
          Math.max(
            result[clinical].iso,
            index
          );

        if (field === 'G') {
          result[clinical]
            .isoGradeCells
            .push(
              a1
            );
        }

        if (field === 'D') {
          result[clinical]
            .dateCells
            .push(
              a1
            );
        }

        continue;
      }


      // ---------------------------------------------------------------
      // FAILED: {{C1F1E}} {{C1F1I}} {{C1F1G}} {{C1F1D}}
      // ---------------------------------------------------------------

      match =
        text.match(
          /^\{\{C([1-5])F(\d+)([EIGD])\}\}$/
        );

      if (match) {

        const clinical =
          Number(
            match[1]
          );

        const index =
          Number(
            match[2]
          );

        const field =
          match[3];

        result[clinical].failure =
          Math.max(
            result[clinical].failure,
            index
          );

        if (field === 'G') {
          result[clinical]
            .failedGradeCells
            .push(
              a1
            );
        }

        if (field === 'D') {
          result[clinical]
            .dateCells
            .push(
              a1
            );
        }
      }
    }
  }

  return result;
}


/**
 * =====================================================================
 * WRITE AVERAGE FORMULAS
 * =====================================================================
 */
function clinicalGrades_writeAuthoritativeSummaries_(
  sheet,
  gradeTerms
) {

  const lastRow =
    sheet.getLastRow();

  const clinicalStarts =
    {};

  for (
    let row = 1;
    row <= lastRow;
    row++
  ) {

    const value =
      String(
        sheet
          .getRange(
            row,
            1
          )
          .getDisplayValue() ||
        ''
      ).trim();

    const match =
      value.match(
        /^Clinical\s+([1-5])$/i
      );

    if (match) {
      clinicalStarts[
        Number(
          match[1]
        )
      ] = row;
    }
  }

  for (
    let clinical = 1;
    clinical <= 5;
    clinical++
  ) {

    const start =
      clinicalStarts[
        clinical
      ];

    if (!start) {
      continue;
    }

    const next =
      clinicalStarts[
        clinical + 1
      ] ||
      (
        lastRow +
        1
      );

    const end =
      next - 1;

    const term =
      gradeTerms[
        appNormalize_(
          'Clinical ' + clinical
        )
      ] || null;

    const categories =
      term && term.categories
        ? term.categories
        : {};

    const categoryValue = function(key) {
      const category = categories[key];
      return category && category.available
        ? category.percentAverage
        : '';
    };

    for (
      let row = start;
      row <= end;
      row++
    ) {

      const label =
        String(
          sheet
            .getRange(
              row,
              3
            )
            .getDisplayValue() ||
          ''
        )
          .trim()
          .toLowerCase();

      if (
        label ===
        'average'
      ) {

        const target =
          sheet.getRange(
            row,
            4
          );

        clinicalGrades_setSummaryValue_(
          target,
          categoryValue('competency')
        );

        break;
      }
    }


    // ---------------------------------------------------------------
    // ISO AVERAGE
    // ---------------------------------------------------------------

    for (
      let row = start;
      row <= end;
      row++
    ) {

      const label =
        String(
          sheet
            .getRange(
              row,
              7
            )
            .getDisplayValue() ||
          ''
        )
          .trim()
          .toLowerCase();

      if (
        label ===
        'iso average'
      ) {

        const target =
          sheet.getRange(
            row,
            9
          );

        clinicalGrades_setSummaryValue_(
          target,
          categoryValue('iso')
        );

        break;
      }
    }

    const summaryRows = [
      { labels: ['professional comps  average'], value: categoryValue('professional') },
      { labels: ['clinical quizzes average 10%'], value: categoryValue('quizzes') },
      { labels: ['clinical test 25%', 'clinical test 30%'], value: categoryValue('clinicalTest') },
      { labels: ['case presentation', 'case presentation 5%', 'case presentation 15%'], value: categoryValue('casePresentation') },
      { labels: ['course average'], value: term ? term.averageBeforeDeductions : '' },
      { labels: ['final ave.'], value: term ? term.recordedAverage : '' }
    ];

    summaryRows.forEach(item => {
      for (let row = start; row <= end; row++) {
        const gLabel = String(sheet.getRange(row, 7).getDisplayValue() || '')
          .trim().toLowerCase();
        const hLabel = String(sheet.getRange(row, 8).getDisplayValue() || '')
          .trim().toLowerCase();
        if (item.labels.indexOf(gLabel) === -1 && item.labels.indexOf(hLabel) === -1) {
          continue;
        }
        clinicalGrades_setSummaryValue_(sheet.getRange(row, 9), item.value);
        break;
      }
    });

    const deductionItems = term && term.deductions && Array.isArray(term.deductions.items)
      ? term.deductions.items
      : [];
    const pointRows = [];
    for (let row = start; row <= end; row++) {
      const label = String(sheet.getRange(row, 7).getDisplayValue() || '')
        .trim().toLowerCase();
      if (label === 'points off:') pointRows.push(row);
    }
    pointRows.forEach((row, index) => {
      const item = deductionItems[index] || null;
      const target = sheet.getRange(row, 9);
      if (!item) {
        target.clearContent();
      } else {
        target.setValue(Number(item.points));
        target.setNumberFormat('0.00');
      }
    });
  }
}


/**
 * Populates the existing Professional Competency token slots in the official
 * template. The category average token is filled only from the authoritative
 * gradeEngine_getStudentSummary() term result; this function never calculates
 * an average independently.
 */
function clinicalGrades_populateProfessionalCompetencies_(
  sheet,
  clinicalNumber,
  records,
  termSummary
) {

  const capacity = 5;

  for (let index = 0; index < capacity; index++) {
    const record = records[index] || null;
    const instructor = record
      ? String(record['Instructor'] || record['Entered By'] || '').trim()
      : '';
    const percent = record
      ? gradeEngine_numeric_(record['Percent Grade'])
      : null;

    const tokenBase =
      '{{C' + clinicalNumber + 'PC' + (index + 1);

    clinicalGrades_setExactToken_(
      sheet,
      tokenBase + 'I}}',
      instructor
    );

    clinicalGrades_setExactToken_(
      sheet,
      tokenBase + 'G}}',
      percent === null
        ? ''
        : (
            percent >= 0 && percent <= 1
              ? percent * 100
              : percent
          )
    );
  }

  const professional =
    termSummary &&
    termSummary.categories
      ? termSummary.categories.professional
      : null;

  clinicalGrades_setExactToken_(
    sheet,
    '{{C' + clinicalNumber + '_PROF_AVG}}',
    professional && professional.available
      ? professional.percentAverage
      : ''
  );

  return capacity;
}


/**
 * Refuse document generation when the master template is missing any of the
 * established Professional Competency token slots. This prevents values from
 * being written into inferred coordinates or silently discarded.
 */
function clinicalGrades_assertProfessionalCompetencyTokens_(
  sheet
) {

  const missing = [];

  for (let clinicalNumber = 1; clinicalNumber <= 5; clinicalNumber++) {
    for (let index = 1; index <= 5; index++) {
      [
        '{{C' + clinicalNumber + 'PC' + index + 'I}}',
        '{{C' + clinicalNumber + 'PC' + index + 'G}}'
      ].forEach(token => {
        const found = sheet
          .createTextFinder(token)
          .matchEntireCell(true)
          .matchCase(true)
          .useRegularExpression(false)
          .findNext();
        if (!found) missing.push(token);
      });
    }

    // The authoritative summary writer already owns the Professional
    // Competency average. Requiring an average token duplicates that contract.
  }

  if (missing.length) {
    throw new Error(
      'Clinical Grades template is missing Professional Competency tokens: ' +
      missing.join(', ')
    );
  }

  return true;
}


function clinicalGrades_writeCompleteDetailSheet_(
  spreadsheet,
  studentName,
  displayStudentId,
  trackingRecords,
  supplementalRecords
) {
  const sheetName = 'Complete Detail';
  let detail = spreadsheet.getSheetByName(sheetName);

  if (!detail) {
    const reusableBlank = spreadsheet.getSheets().find(candidate =>
      candidate.getName() !== 'Clinical Grades' &&
      candidate.getLastRow() <= 1 &&
      candidate.getLastColumn() <= 1
    );
    detail = reusableBlank || spreadsheet.insertSheet();
    detail.setName(sheetName);
  }

  detail.clear();
  detail.setHiddenGridlines(true);

  const headers = [
    'Clinical Term', 'Category', 'Date', 'Evaluation / Assignment',
    'Evaluator', 'Raw Final Score', 'Selected Grade', 'Grade Source',
    'Result', 'Record ID'
  ];
  const rows = [];

  (trackingRecords || []).forEach(record => {
    const type = String(record['Evaluation Type'] || '').trim();
    const normalizedType = appNormalize_(type);
    if (['competency', 'iso', 'sso', 'simulation'].indexOf(normalizedType) === -1) return;

    const resolved = gradeEngine_resolveTrackingGrade_(record);
    const result = String(record['Result'] || '').trim();
    rows.push([
      gradeEngine_effectiveClinicalTerm_(record), type || 'Clinical Tracking',
      appDateOnly_(record['Date']) || '', String(record['Examination'] || '').trim(),
      String(record['Evaluator'] || '').trim(), gradeEngine_numeric_(record['Final Score']),
      resolved.available && resolved.percent !== null ? resolved.percent : result,
      resolved.available ? resolved.source : (result ? 'Result' : ''), result,
      String(record['Submission ID'] || '').trim()
    ]);
  });

  (supplementalRecords || []).forEach(record => {
    const percent = gradeEngine_numeric_(record['Percent Grade']);
    rows.push([
      String(record['Clinical Term'] || '').trim(),
      String(record['Grade Category'] || '').trim(),
      appDateOnly_(record['Evaluation Date'] || record['Date']) || '',
      String(record['Assignment / Evaluation Name'] || '').trim(),
      String(record['Instructor'] || record['Entered By'] || '').trim(),
      gradeEngine_numeric_(record['Raw Score']),
      percent === null ? '' : (percent >= 0 && percent <= 1 ? percent * 100 : percent),
      'Grade Records Percent Grade', '',
      String(record['Grade Record ID'] || record['Record ID'] || '').trim()
    ]);
  });

  rows.sort((a, b) => {
    const termDifference = clinicalGrades_clinicalNumber_(a[0]) - clinicalGrades_clinicalNumber_(b[0]);
    if (termDifference) return termDifference;
    const aDate = a[2] instanceof Date ? a[2].getTime() : 0;
    const bDate = b[2] instanceof Date ? b[2].getTime() : 0;
    return aDate - bDate;
  });

  detail.getRange('A1:J1').merge();
  detail.getRange('A1').setValue(studentName + ' (' + displayStudentId + ') — Complete Clinical Grade Detail')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1f4e78').setFontColor('#ffffff');
  detail.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight('bold')
    .setBackground('#d9eaf7').setHorizontalAlignment('center');

  if (rows.length) {
    detail.getRange(3, 1, rows.length, headers.length).setValues(rows);
    detail.getRange(3, 3, rows.length, 1).setNumberFormat('M/d/yy');
    detail.getRange(3, 6, rows.length, 2).setNumberFormat('0.00');
    detail.getRange(3, 1, rows.length, headers.length).setVerticalAlignment('middle').setWrap(true);
  }

  detail.setFrozenRows(2);
  detail.setColumnWidths(1, 2, 115); detail.setColumnWidth(3, 85);
  detail.setColumnWidth(4, 220); detail.setColumnWidth(5, 150);
  detail.setColumnWidths(6, 2, 105); detail.setColumnWidth(8, 170);
  detail.setColumnWidth(9, 85); detail.setColumnWidth(10, 150);
  detail.setRowHeight(1, 30); detail.setRowHeight(2, 32);
  const filter = detail.getFilter();
  if (filter) filter.remove();
  detail.getRange(2, 1, Math.max(rows.length + 1, 2), headers.length).createFilter();

  return { rowCount: rows.length, sheetName: sheetName };
}


function clinicalGrades_setSummaryValue_(
  range,
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    range.clearContent();
    return;
  }

  range.setValue(
    Number(value)
  );

  range.setNumberFormat(
    '0.00'
  );
}


/**
 * =====================================================================
 * PERCENT GRADE
 * =====================================================================
 *
 * Numeric evaluation grades are converted using the existing Grade
 * Engine so the grades sheet receives the same percentage representation
 * used elsewhere in the dashboard.
 *
 * Pass/fail-only records remain text.
 */
function clinicalGrades_percentGrade_(
  record
) {

  const resolved =
    gradeEngine_resolveTrackingGrade_(
      record
    );

  if (
    resolved.available &&
    resolved.percent !== null
  ) {
    return resolved.percent;
  }

  if (!resolved.available) {

    const result =
      String(
        record['Result'] || ''
      ).trim();

    return result
      ? clinicalGrades_titleCase_(
          result
        )
      : '';
  }

  const result =
    String(
      record['Result'] || ''
    ).trim();

  return result
    ? clinicalGrades_titleCase_(
        result
      )
    : '';
}


/**
 * =====================================================================
 * PASS / FAIL HELPERS
 * =====================================================================
 */
function clinicalGrades_isPassingCompetency_(
  record
) {

  return (

    appNormalize_(
      record['Evaluation Type']
    ) ===
      'competency' &&

    appNormalize_(
      record['Result']
    ) ===
      'pass' &&

    !appIsTrue_(
      record['Automatic Failure']
    )

  );
}


function clinicalGrades_isFailedCompetency_(
  record
) {

  return (

    appNormalize_(
      record['Evaluation Type']
    ) ===
      'competency' &&

    (
      appNormalize_(
        record['Result']
      ) ===
        'fail' ||
      appIsTrue_(
        record['Automatic Failure']
      )
    )

  );
}


/**
 * =====================================================================
 * STUDENT HELPERS
 * =====================================================================
 */
function clinicalGrades_selectStudentRow_(
  rows
) {

  const active =
    rows.filter(row =>
      appIsTrue_(
        row['Active']
      )
    );

  const candidates =
    active.length
      ? active
      : rows;

  return candidates
    .slice()
    .sort(
      (a, b) =>
        clinicalGrades_clinicalNumber_(
          b['Clinical Term']
        ) -
        clinicalGrades_clinicalNumber_(
          a['Clinical Term']
        )
    )[0];
}


function clinicalGrades_sameStudent_(
  record,
  studentId,
  studentName
) {

  const targetId =
    appNormalize_(
      studentId
    );

  const recordId =
    appNormalize_(
      record['Student ID']
    );

  if (
    targetId &&
    recordId
  ) {

    return (
      targetId ===
      recordId
    );
  }

  return (
    appNormalize_(
      record['Student']
    ) ===
    appNormalize_(
      studentName
    )
  );
}


/**
 * =====================================================================
 * SEMESTER RULE
 * =====================================================================
 */
function clinicalGrades_findSemesterRule_(
  semesterRules,
  clinicalTerm
) {

  const active =
    semesterRules.find(row =>

      appNormalize_(
        row['Clinical Term']
      ) ===
      appNormalize_(
        clinicalTerm
      ) &&

      appIsTrue_(
        row['Active']
      )

    );

  if (active) {
    return active;
  }

  return (
    semesterRules.find(row =>
      appNormalize_(
        row['Clinical Term']
      ) ===
      appNormalize_(
        clinicalTerm
      )
    ) ||
    null
  );
}


/**
 * =====================================================================
 * TEMPLATE TOKEN HELPERS
 * =====================================================================
 */
function clinicalGrades_replaceEmbeddedToken_(
  sheet,
  token,
  replacement
) {

  sheet
    .createTextFinder(
      token
    )
    .matchCase(
      true
    )
    .useRegularExpression(
      false
    )
    .replaceAllWith(
      String(
        replacement || ''
      )
    );
}


function clinicalGrades_setExactToken_(
  sheet,
  token,
  value
) {

  const matches =
    sheet
      .createTextFinder(
        token
      )
      .matchEntireCell(
        true
      )
      .matchCase(
        true
      )
      .useRegularExpression(
        false
      )
      .findAll();

  matches.forEach(range => {

    if (
      value instanceof Date
    ) {

      range.setValue(
        value
      );

      return;
    }

    if (
      typeof value ===
        'number' &&
      Number.isFinite(
        value
      )
    ) {

      range.setValue(
        value
      );

      return;
    }

    range.setValue(
      value === null ||
      value === undefined
        ? ''
        : String(
            value
          )
    );
  });
}


function clinicalGrades_clearTermTokens_(
  sheet,
  clinicalNumber
) {

  sheet
    .createTextFinder(
      '\\{\\{C' +
      clinicalNumber +
      '[^}]+\\}\\}'
    )
    .useRegularExpression(
      true
    )
    .replaceAllWith(
      ''
    );
}


function clinicalGrades_clearAllRemainingTokens_(
  sheet
) {

  sheet
    .createTextFinder(
      '\\{\\{[^}]+\\}\\}'
    )
    .useRegularExpression(
      true
    )
    .replaceAllWith(
      ''
    );
}


/**
 * =====================================================================
 * DATE FORMAT
 * =====================================================================
 */
function clinicalGrades_formatDates_(
  sheet,
  tokenLayout
) {

  for (
    let clinical = 1;
    clinical <= 5;
    clinical++
  ) {

    (
      tokenLayout[
        clinical
      ].dateCells || []
    )
      .forEach(a1 =>

        sheet
          .getRange(
            a1
          )
          .setNumberFormat(
            'M/d/yy'
          )

      );
  }
}


/**
 * =====================================================================
 * DISPLAY HELPERS
 * =====================================================================
 */
function clinicalGrades_displayExam_(
  record
) {

  return String(
    record['Examination'] || ''
  ).trim();
}


function clinicalGrades_displayEvaluator_(
  record
) {

  return String(
    record['Evaluator'] || ''
  ).trim();
}


function clinicalGrades_date_(
  value
) {

  return (
    appDateOnly_(
      value
    ) ||
    ''
  );
}


function clinicalGrades_recordDateSort_(
  a,
  b
) {

  return (
    appDateValue_(
      a['Date']
    ) -
    appDateValue_(
      b['Date']
    )
  );
}


function clinicalGrades_formatStudentId_(
  studentId
) {

  const value =
    String(
      studentId || ''
    ).trim();

  const match =
    value.match(
      /#\s*(\d+)/i
    );

  return match
    ? '#' + match[1]
    : value;
}


function clinicalGrades_clinicalNumber_(
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


function clinicalGrades_titleCase_(
  value
) {

  const text =
    String(
      value || ''
    )
      .trim()
      .toLowerCase();

  if (!text) {
    return '';
  }

  return (
    text.charAt(0)
      .toUpperCase() +
    text.slice(1)
  );
}


/**
 * =====================================================================
 * GENERATED SPREADSHEET HELPERS
 * =====================================================================
 */
function clinicalGrades_getMainSheet_(
  spreadsheet
) {

  return (
    spreadsheet.getSheetByName(
      'Clinical Grades'
    ) ||
    spreadsheet.getSheetByName(
      'Sheet1'
    ) ||
    spreadsheet.getSheets()[0]
  );
}


function clinicalGrades_openSpreadsheetWithRetry_(
  spreadsheetId
) {

  const maxAttempts =
    6;

  let lastError =
    null;

  for (
    let attempt = 1;
    attempt <= maxAttempts;
    attempt++
  ) {

    try {

      return SpreadsheetApp.openById(
        spreadsheetId
      );

    } catch (error) {

      lastError =
        error;

      Utilities.sleep(
        1000 *
        attempt
      );
    }
  }

  throw new Error(
    'Could not open the generated Clinical Grades spreadsheet after ' +
    maxAttempts +
    ' attempts. ' +
    (
      lastError
        ? lastError.message
        : ''
    )
  );
}


/**
 * =====================================================================
 * MANUAL TEST
 * =====================================================================
 *
 * Run THIS function from the Apps Script Run menu.
 *
 * It will create a real Clinical Grades sheet inside the student's
 * Generated Documents folder.
 */
function testClinicalGradesSheet() {

  const TEST_STUDENT_ID =
    'DEPT # 22';

  const result =
    generateClinicalGradesSheet(
      TEST_STUDENT_ID
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
