/**Program Requirements
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * PROGRAM REQUIREMENT DOCUMENT GENERATOR
 * 05_ProgramRequirementDocumentGenerator.gs
 * =====================================================================
 *
 * CURRENT-COHORT VERSION
 *
 * PURPOSE
 * -------
 * Generates the cumulative December 2025/current-cohort Program
 * Requirements document for one student.
 *
 * AUTHORITATIVE SOURCES
 * ---------------------
 * - Clinical Tracking = factual evaluation history
 * - Examinations      = authoritative examination definitions / IDs
 * - 04_RequirementsEngine.gs = authoritative pass/fail + uniqueness logic
 *
 * TEMPLATE CONTRACT
 * -----------------
 * - P01 through P72 map directly to EX-001 through EX-072.
 * - No fuzzy text matching is permitted.
 * - General Patient Care PC1-PC10 is intentionally left blank until an
 *   authoritative data source exists.
 * - Failed competency and ISO history use the existing Clinical Tracking
 *   logic and presentation.
 * =====================================================================
 */

const PROGRAM_DOCUMENT_TEMPLATE_ID =
  '1Pa31xYCjEjwODytBHHw1UK29bHhbqgQ8ZWVNYi094Y0';


/**
 * =====================================================================
 * MAIN PUBLIC FUNCTION
 * =====================================================================
 */
function generateProgramRequirementSheet(
  studentId
) {

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


  const examinations =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.EXAMINATIONS
      )
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
    programDocument_selectStudentRow_(
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
      'Student has no Cohort assigned: ' +
      studentName
    );
  }


  const displayStudentId =
    programDocument_formatStudentId_(
      studentId
    );


  // ===================================================================
  // STUDENT TRACKING HISTORY
  // ===================================================================

  const records =
    tracking.filter(record =>
      programDocument_sameStudent_(
        record,
        studentId,
        studentName
      )
    );


  const passingCompetencies =
    reqGetUniquePassingCompetencies_(
      records.filter(
        reqIsPassingCompetency_
      )
    )
      .sort(
        programDocument_recordDateSort_
      );


  const failedRecords =
    programDocument_uniqueFailedRecords_(
      records
        .filter(
          reqIsFailedCompetency_
        )
        .sort(
          programDocument_recordDateSort_
        )
    );


  const isoRecords =
    records
      .filter(record =>
        appNormalize_(
          record['Evaluation Type']
        ) ===
        'iso'
      )
      .sort(
        programDocument_recordDateSort_
      );


  // ===================================================================
  // VERIFY CURRENT-COHORT EXAMINATION CONTRACT
  //
  // P01 -> EX-001
  // ...
  // P72 -> EX-072
  //
  // EX-073 is intentionally NOT part of this current-cohort template.
  // ===================================================================

  const examById = {};


  examinations.forEach(exam => {

    const examId =
      String(
        exam['Examination ID'] || ''
      ).trim();


    if (examId) {
      examById[
        appNormalize_(
          examId
        )
      ] = exam;
    }

  });


  const expectedExamIds = {};


  for (
    let i = 1;
    i <= 72;
    i++
  ) {

    const examId =
      programDocument_examIdForProgramNumber_(
        i
      );


    expectedExamIds[
      appNormalize_(
        examId
      )
    ] = true;


    const exam =
      examById[
        appNormalize_(
          examId
        )
      ];


    if (!exam) {
      throw new Error(
        'Current-cohort Program Requirements expects ' +
        examId +
        ', but that Examination ID was not found.'
      );
    }


    if (
      Object.prototype
        .hasOwnProperty
        .call(exam, 'Active') &&
      !appIsTrue_(
        exam['Active']
      )
    ) {
      throw new Error(
        'Current-cohort Program Requirements expects ' +
        examId +
        ', but that examination is inactive.'
      );
    }

  }


  // ===================================================================
  // INDEX PASSING COMPETENCIES BY EXAMINATION ID
  //
  // Because Clinical Tracking now carries stable Examination IDs, the
  // template never needs to infer identity from wording.
  //
  // If multiple passing records share one Examination ID, the earliest
  // passing record is used to fill the single cumulative program row.
  // ===================================================================

  const passingByExamId = {};


  passingCompetencies.forEach(record => {

    const examId =
      appNormalize_(
        record['Examination ID']
      );


    if (
      !examId ||
      !expectedExamIds[examId] ||
      passingByExamId[examId]
    ) {
      return;
    }


    passingByExamId[examId] =
      record;

  });


  // ===================================================================
  // VERIFY TEMPLATE
  // ===================================================================

  const template =
    DriveApp.getFileById(
      PROGRAM_DOCUMENT_TEMPLATE_ID
    );


  if (
    template.getMimeType() !==
    'application/vnd.google-apps.document'
  ) {
    throw new Error(
      'Program Requirement template is not a native Google Doc.'
    );
  }


  // ===================================================================
  // COPY TO STUDENT GENERATED DOCUMENTS FOLDER
  // ===================================================================

  const outputName =
    studentName +
    ' - Program Requirements';


  const copyResult =
    driveService_copyTemplateToStudentFolder({

      templateId:
        PROGRAM_DOCUMENT_TEMPLATE_ID,

      fileName:
        outputName,

      student:
        studentName,

      cohort:
        cohort,

      destination:
        'documents'

    });


  const document =
    programDocument_openWithRetry_(
      copyResult.fileId
    );


  const body =
    document.getBody();


  // ===================================================================
  // BASIC TOKENS
  // ===================================================================

  programDocument_replaceToken_(
    body,
    '{{STUDENT_NAME}}',
    studentName
  );


  programDocument_replaceToken_(
    body,
    '{{STUDENT_NO}}',
    displayStudentId
  );


  // ===================================================================
  // PROGRAM REQUIREMENT ROWS P01-P72
  // ===================================================================

  let programRowsPopulated = 0;


  for (
    let i = 1;
    i <= 72;
    i++
  ) {

    const prefix =
      'P' +
      String(i)
        .padStart(
          2,
          '0'
        );


    const examId =
      programDocument_examIdForProgramNumber_(
        i
      );


    const record =
      passingByExamId[
        appNormalize_(
          examId
        )
      ] ||
      null;


    if (record) {
      programRowsPopulated++;
    }


    programDocument_replaceToken_(
      body,
      '{{' + prefix + 'I}}',
      record
        ? programDocument_displayEvaluator_(
            record
          )
        : ''
    );


    programDocument_replaceToken_(
      body,
      '{{' + prefix + 'G}}',
      record
        ? programDocument_displayScore_(
            record
          )
        : ''
    );


    programDocument_replaceToken_(
      body,
      '{{' + prefix + 'D}}',
      record
        ? programDocument_displayDate_(
            record['Date']
          )
        : ''
    );

  }


  // ===================================================================
  // GENERAL PATIENT CARE
  //
  // All ten rows exist in the current-cohort master. There is not yet an
  // authoritative system data source for these items, so intentionally
  // leave every student-specific token blank.
  // ===================================================================

  for (
    let i = 1;
    i <= 10;
    i++
  ) {

    programDocument_replaceToken_(
      body,
      '{{PC' + i + 'D}}',
      ''
    );


    programDocument_replaceToken_(
      body,
      '{{PC' + i + 'V}}',
      ''
    );

  }


  // ===================================================================
  // FAILED COMPETENCY HISTORY
  // ===================================================================

  for (
    let i = 1;
    i <= 13;
    i++
  ) {

    const record =
      failedRecords[
        i - 1
      ] ||
      null;


    programDocument_replaceToken_(
      body,
      '{{F' + i + 'E}}',
      record
        ? programDocument_displayExam_(
            record
          )
        : ''
    );


    programDocument_replaceToken_(
      body,
      '{{F' + i + 'I}}',
      record
        ? programDocument_displayEvaluator_(
            record
          )
        : ''
    );


    programDocument_replaceToken_(
      body,
      '{{F' + i + 'D}}',
      record
        ? programDocument_displayDate_(
            record['Date']
          )
        : ''
    );


    programDocument_replaceToken_(
      body,
      '{{F' + i + 'R}}',
      record &&
      programDocument_remediationComplete_(
        records,
        record
      )
        ? 'Yes'
        : ''
    );

  }


  // ===================================================================
  // ISO HISTORY
  // ===================================================================

  for (
    let i = 1;
    i <= 13;
    i++
  ) {

    const record =
      isoRecords[
        i - 1
      ] ||
      null;


    programDocument_replaceToken_(
      body,
      '{{ISO' + i + 'E}}',
      record
        ? programDocument_displayExam_(
            record
          )
        : ''
    );


    programDocument_replaceToken_(
      body,
      '{{ISO' + i + 'I}}',
      record
        ? programDocument_displayEvaluator_(
            record
          )
        : ''
    );


    programDocument_replaceToken_(
      body,
      '{{ISO' + i + 'G}}',
      record
        ? programDocument_displayScore_(
            record
          )
        : ''
    );


    programDocument_replaceToken_(
      body,
      '{{ISO' + i + 'D}}',
      record
        ? programDocument_displayDate_(
            record['Date']
          )
        : ''
    );

  }


  // ===================================================================
  // UNRESOLVED PASSING RECORDS
  //
  // A passing competency is unresolved only when it lacks a usable
  // current-cohort Examination ID. It is never guessed into another row.
  // ===================================================================

  const unresolvedPassing =
    passingCompetencies
      .filter(record => {

        const examId =
          appNormalize_(
            record['Examination ID']
          );


        return (
          !examId ||
          !expectedExamIds[examId]
        );

      })
      .map(record => ({
        submissionId:
          String(
            record['Submission ID'] || ''
          ).trim(),

        examinationId:
          String(
            record['Examination ID'] || ''
          ).trim(),

        examination:
          String(
            record['Examination'] || ''
          ).trim()
      }));


  programDocument_forceBlackHeaderTextWhite_(
    body
  );


  document.saveAndClose();


  // ===================================================================
  // LOG / RESULT
  // ===================================================================

  console.log(
    '================================'
  );

  console.log(
    'PROGRAM REQUIREMENT SHEET CREATED'
  );

  console.log(
    'Student: ' +
    studentName
  );

  console.log(
    'Document: ' +
    copyResult.fileUrl
  );

  console.log(
    'Program rows populated: ' +
    programRowsPopulated +
    ' / 72'
  );


  if (
    unresolvedPassing.length > 0
  ) {

    console.log(
      'UNRESOLVED PASSING RECORDS: ' +
      JSON.stringify(
        unresolvedPassing
      )
    );

  }


  console.log(
    '================================'
  );


  return {

    success:
      true,

    name:
      outputName,

    url:
      copyResult.fileUrl,

    documentId:
      copyResult.fileId,

    counts: {

      passingCompetencies:
        passingCompetencies.length,

      programRowsPopulated:
        programRowsPopulated,

      iso:
        isoRecords.length,

      failures:
        failedRecords.length,

      unresolvedPassingRecords:
        unresolvedPassing.length

    },

    unresolvedPassingExams:
      unresolvedPassing

  };
}


/**
 * =====================================================================
 * PROGRAM NUMBER -> EXAMINATION ID
 * =====================================================================
 */
function programDocument_examIdForProgramNumber_(
  number
) {

  return (
    'EX-' +
    String(number)
      .padStart(
        3,
        '0'
      )
  );
}


/**
 * =====================================================================
 * UNIQUE FAILED COMPETENCY HISTORY
 *
 * The Program Requirements form has one row per failed examination/body
 * part, not one row per failed attempt. Keep the earliest failed attempt
 * for display and let the remediation helper inspect all attempts.
 * =====================================================================
 */
function programDocument_uniqueFailedRecords_(
  failedRecords
) {

  const seen = {};
  const unique = [];


  (failedRecords || []).forEach(record => {

    const key =
      programDocument_remediationKey_(
        record
      );


    if (
      !key ||
      seen[key]
    ) {
      return;
    }


    seen[key] = true;
    unique.push(record);

  });


  return unique;
}


/**
 * =====================================================================
 * REMEDIATION COMPLETE FOR FAILURE KEY
 * =====================================================================
 */
function programDocument_remediationComplete_(
  allRecords,
  failureRecord
) {

  const key =
    programDocument_remediationKey_(
      failureRecord
    );


  if (!key) {
    return false;
  }


  return allRecords.some(record =>
    programDocument_remediationKey_(
      record
    ) ===
    key &&
    appIsTrue_(
      record[
        'Remediation Complete'
      ]
    )
  );
}


function programDocument_remediationKey_(
  record
) {

  const exam =
    programDocument_normalize_(
      record[
        'Examination'
      ]
    );


  const area =
    programDocument_normalize_(
      record[
        'Anatomical Area'
      ]
    );


  if (!exam) {
    return '';
  }


  return area
    ? exam + '|' + area
    : exam;
}


/**
 * =====================================================================
 * SELECT CURRENT/HIGHEST STUDENT ROW
 * =====================================================================
 */
function programDocument_selectStudentRow_(
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
        programDocument_clinicalNumber_(
          b['Clinical Term']
        ) -
        programDocument_clinicalNumber_(
          a['Clinical Term']
        )
    )[0];
}


/**
 * =====================================================================
 * SAME STUDENT
 * =====================================================================
 */
function programDocument_sameStudent_(
  record,
  studentId,
  studentName
) {

  const targetId =
    programDocument_normalize_(
      studentId
    );


  const recordId =
    programDocument_normalize_(
      record[
        'Student ID'
      ]
    );


  if (
    targetId &&
    recordId
  ) {
    return targetId === recordId;
  }


  return (
    programDocument_normalize_(
      record[
        'Student'
      ]
    ) ===
    programDocument_normalize_(
      studentName
    )
  );
}


/**
 * =====================================================================
 * DISPLAY HELPERS
 * =====================================================================
 */
function programDocument_displayExam_(
  record
) {

  return String(
    record[
      'Examination'
    ] || ''
  ).trim();
}


function programDocument_displayEvaluator_(
  record
) {

  return String(
    record[
      'Evaluator'
    ] || ''
  ).trim();
}


/**
 * =====================================================================
 * SCORE DISPLAY
 *
 * Priority:
 * 1. True structured raw/final score.
 * 2. Historical awarded percentage when no raw score exists.
 * 3. True Pass/Fail/Simulation fallback.
 * =====================================================================
 */
function programDocument_displayScore_(
  record
) {

  record = record || {};


  // 1. True structured raw/final competency score.
  const rawScoreFields = [
    'Final Score',
    'Historical Raw Score',
    'Original Raw Score',
    'Reconstructed Raw Score'
  ];


  for (
    let i = 0;
    i < rawScoreFields.length;
    i++
  ) {

    const field =
      rawScoreFields[i];


    if (
      !Object.prototype
        .hasOwnProperty
        .call(
          record,
          field
        )
    ) {
      continue;
    }


    const rawScore =
      programDocument_numericScore_(
        record[field]
      );


    if (rawScore !== null) {
      return rawScore.toFixed(2);
    }

  }


  // 2. Historical awarded percentage.
  //    This is used when no true 0-3 raw score exists.
  //    It is NOT reverse-converted into a fabricated raw score.
  const historicalPercent =
    programDocument_numericScore_(
      record['Historical Percent Grade']
    );


  if (historicalPercent !== null) {

    const displayPercent =
      Number.isInteger(
        historicalPercent
      )
        ? String(
            historicalPercent
          )
        : String(
            Math.round(
              historicalPercent * 100
            ) / 100
          );


    return (
      displayPercent +
      '%'
    );
  }


  // 3. True Pass/Fail/Simulation fallback.
  const evaluationType =
    programDocument_normalize_(
      record['Evaluation Type']
    );


  const result =
    String(
      record['Result'] || ''
    ).trim();


  const normalizedResult =
    programDocument_normalize_(
      result
    );


  if (
    evaluationType === 'simulation' ||
    normalizedResult === 'simulation' ||
    normalizedResult === 'sim'
  ) {
    return 'Simulation';
  }


  if (normalizedResult === 'pass') {
    return 'Pass';
  }


  if (normalizedResult === 'fail') {
    return 'Fail';
  }


  return '';
}



/**
 * =====================================================================
 * NUMERIC SCORE HELPER
 * =====================================================================
 */
function programDocument_numericScore_(
  value
) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return null;
  }


  const numeric =
    Number(
      value
    );


  return Number.isFinite(
    numeric
  )
    ? numeric
    : null;
}


/**
 * =====================================================================
 * DATE DISPLAY HELPER
 * =====================================================================
 */
function programDocument_displayDate_(
  value
) {

  const date =
    appDateOnly_(
      value
    );


  if (!date) {
    return '';
  }


  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'M/d/yy'
  );
}


/**
 * =====================================================================
 * TOKEN REPLACEMENT / OUTPUT FORMATTING
 *
 * NOTE:
 * The visible font size of placeholder text in the MASTER does not
 * control the final generated value. This function explicitly formats
 * every value inserted into the generated student copy.
 * =====================================================================
 */
function programDocument_replaceToken_(
  body,
  token,
  value
) {

  const replacement =
    String(
      value === null ||
      value === undefined
        ? ''
        : value
    );


  const pattern =
    appEscapeRegex_(
      token
    );


  const isStudentName =
    token ===
    '{{STUDENT_NAME}}';


  const isStudentId =
    token ===
    '{{STUDENT_NO}}';


  const isGrade =
    /\{\{(?:P\d{2}G|ISO\d+G)\}\}/
      .test(
        token
      );


  let fontSize =
    isStudentName
      ? 16
      : isStudentId
        ? 14
        : 10;


  if (
    replacement.length > 28
  ) {
    fontSize =
      Math.min(
        fontSize,
        8
      );
  } else if (
    replacement.length > 20
  ) {
    fontSize =
      Math.min(
        fontSize,
        9
      );
  }


  let replaced = false;
  let found = null;


  while (
    (
      found =
        body.findText(
          pattern
        )
    )
  ) {

    replaced = true;


    const text =
      found
        .getElement()
        .asText();


    const start =
      found.getStartOffset();


    const end =
      found.getEndOffsetInclusive();


    text.deleteText(
      start,
      end
    );


    if (!replacement) {
      continue;
    }


    text.insertText(
      start,
      replacement
    );


    const insertedEnd =
      start +
      replacement.length -
      1;


    text.setFontSize(
      start,
      insertedEnd,
      fontSize
    );


    if (
      isStudentName ||
      isStudentId ||
      isGrade
    ) {
      text.setBold(
        start,
        insertedEnd,
        true
      );
    }

  }


  return replaced;
}


/**
 * =====================================================================
 * BLACK TABLE HEADER TEXT
 *
 * Preserve the template's existing fill, borders, sizing, and wording.
 * Any table cell whose background is black is treated as a header cell;
 * its existing text is forced to white and bold.
 * =====================================================================
 */
function programDocument_forceBlackHeaderTextWhite_(
  body
) {

  const tables =
    body.getTables();


  tables.forEach(table => {

    for (
      let rowIndex = 0;
      rowIndex < table.getNumRows();
      rowIndex++
    ) {

      const row =
        table.getRow(
          rowIndex
        );


      for (
        let cellIndex = 0;
        cellIndex < row.getNumCells();
        cellIndex++
      ) {

        const cell =
          row.getCell(
            cellIndex
          );


        const background =
          String(
            cell.getBackgroundColor() || ''
          )
            .trim()
            .toLowerCase();


        if (
          background !== '#000000' &&
          background !== '#000'
        ) {
          continue;
        }


        const text =
          cell.editAsText();


        const value =
          text.getText();


        if (!value) {
          continue;
        }


        text.setForegroundColor(
          0,
          value.length - 1,
          '#ffffff'
        );


        text.setBold(
          0,
          value.length - 1,
          true
        );

      }

    }

  });
}


/**
 * =====================================================================
 * DATE SORT
 * =====================================================================
 */
function programDocument_recordDateSort_(
  a,
  b
) {

  return (
    programDocument_dateValue_(
      a['Date']
    ) -
    programDocument_dateValue_(
      b['Date']
    )
  );
}


function programDocument_dateValue_(
  value
) {

  const date =
    appDateOnly_(
      value
    );


  return date
    ? date.getTime()
    : 0;
}


/**
 * =====================================================================
 * STUDENT ID DISPLAY
 * =====================================================================
 */
function programDocument_formatStudentId_(
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


/**
 * =====================================================================
 * CLINICAL NUMBER
 * =====================================================================
 */
function programDocument_clinicalNumber_(
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
 * NORMALIZE
 * =====================================================================
 */
function programDocument_normalize_(
  value
) {

  return String(
    value || ''
  )
    .trim()
    .toLowerCase()
    .replace(
      /\s+/g,
      ' '
    );
}


/**
 * =====================================================================
 * OPEN GENERATED DOCUMENT WITH RETRY
 * =====================================================================
 */
function programDocument_openWithRetry_(
  documentId
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

      return DocumentApp.openById(
        documentId
      );

    } catch (error) {

      lastError =
        error;


      Utilities.sleep(
        1000 * attempt
      );

    }

  }


  throw new Error(
    'Could not open the generated Program Requirement document after ' +
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
 * Running this WILL create a generated Program Requirements document.
 */
function testProgramRequirementDocument() {

  const TEST_STUDENT_ID =
    'TEST-001';


  const result =
    generateProgramRequirementSheet(
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
/**
 * =====================================================================
 * DIAGNOSTIC — PROGRAM REQUIREMENTS DOCUMENT ROUTING
 * =====================================================================
 *
 * READ ONLY.
 *
 * Purpose:
 * Trace why the dashboard may be opening an older Program Requirements
 * document instead of the newly generated copy.
 *
 * Makes NO changes.
 */
function diagnoseProgramRequirementRouting() {

  const STUDENT_ID = 'TEST-001';

  // The unexpected document currently being surfaced.
  const SUSPICIOUS_DOCUMENT_ID =
    '1NRgtRMeB8oOvYlrzXRHbvO-NsXar1xnKflV09_okSEk';


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const students =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENTS
      )
    );


  const studentRows =
    students.filter(row =>
      appNormalize_(
        row['Student ID']
      ) ===
      appNormalize_(
        STUDENT_ID
      )
    );


  if (!studentRows.length) {
    throw new Error(
      'Student not found: ' +
      STUDENT_ID
    );
  }


  const student =
    programDocument_selectStudentRow_(
      studentRows
    );


  const studentName =
    String(
      student['Student'] || ''
    ).trim();


  const expectedOutputName =
    studentName +
    ' - Program Requirements';


  console.log(
    '================================================'
  );

  console.log(
    'PROGRAM REQUIREMENTS ROUTING DIAGNOSTIC'
  );

  console.log(
    '================================================'
  );


  // ===================================================================
  // 1. WHAT TEMPLATE DOES THIS GENERATOR THINK IT USES?
  // ===================================================================

  console.log('');
  console.log(
    '--- GENERATOR TEMPLATE ---'
  );


  console.log(
    'PROGRAM_DOCUMENT_TEMPLATE_ID: ' +
    PROGRAM_DOCUMENT_TEMPLATE_ID
  );


  try {

    const generatorTemplate =
      DriveApp.getFileById(
        PROGRAM_DOCUMENT_TEMPLATE_ID
      );


    console.log(
      'Generator template name: ' +
      generatorTemplate.getName()
    );


    console.log(
      'Generator template URL: ' +
      generatorTemplate.getUrl()
    );

  } catch (error) {

    console.log(
      'ERROR opening generator template: ' +
      error.message
    );

  }


  // ===================================================================
  // 2. WHAT TEMPLATE DOES APP_CONFIG THINK IT USES?
  // ===================================================================

  console.log('');
  console.log(
    '--- APP_CONFIG TEMPLATE ---'
  );


  let configTemplateId =
    '';


  try {

    configTemplateId =
      String(
        APP_CONFIG &&
        APP_CONFIG.TEMPLATES &&
        APP_CONFIG.TEMPLATES.PROGRAM_REQUIREMENTS
          ? APP_CONFIG.TEMPLATES.PROGRAM_REQUIREMENTS
          : ''
      ).trim();

  } catch (error) {

    configTemplateId =
      '';

  }


  console.log(
    'APP_CONFIG.TEMPLATES.PROGRAM_REQUIREMENTS: ' +
    (
      configTemplateId ||
      '[NOT CONFIGURED]'
    )
  );


  if (configTemplateId) {

    try {

      const configTemplate =
        DriveApp.getFileById(
          configTemplateId
        );


      console.log(
        'APP_CONFIG template name: ' +
        configTemplate.getName()
      );


      console.log(
        'APP_CONFIG template URL: ' +
        configTemplate.getUrl()
      );

    } catch (error) {

      console.log(
        'ERROR opening APP_CONFIG template: ' +
        error.message
      );

    }

  }


  // ===================================================================
  // 3. INSPECT THE SUSPICIOUS DOCUMENT DIRECTLY
  // ===================================================================

  console.log('');
  console.log(
    '--- SUSPICIOUS DOCUMENT ---'
  );


  diagnoseProgramRequirementFile_(
    SUSPICIOUS_DOCUMENT_ID
  );


  // ===================================================================
  // 4. FIND EVERY FILE WITH THE EXPECTED GENERATED NAME
  // ===================================================================

  console.log('');
  console.log(
    '--- ALL FILES NAMED "' +
    expectedOutputName +
    '" ---'
  );


  const matchingFiles =
    DriveApp.getFilesByName(
      expectedOutputName
    );


  let matchCount =
    0;


  while (
    matchingFiles.hasNext()
  ) {

    matchCount++;


    const file =
      matchingFiles.next();


    console.log('');
    console.log(
      'MATCH #' +
      matchCount
    );


    diagnoseProgramRequirementFile_(
      file.getId()
    );

  }


  console.log('');
  console.log(
    'Total same-named files found: ' +
    matchCount
  );


  // ===================================================================
  // 5. KNOWN NEW TEST DOCUMENT
  // ===================================================================

  console.log('');
  console.log(
    '--- KNOWN NEW TEST DOCUMENT ---'
  );


  const KNOWN_NEW_DOCUMENT_ID =
    '1VP7VuRBdWnXmEKbZUqhvCxXl1dbHjZTj8cVKyiIppSQ';


  diagnoseProgramRequirementFile_(
    KNOWN_NEW_DOCUMENT_ID
  );


  console.log('');
  console.log(
    '================================================'
  );

  console.log(
    'END ROUTING DIAGNOSTIC'
  );

  console.log(
    '================================================'
  );

}


/**
 * =====================================================================
 * FILE INSPECTION HELPER
 * =====================================================================
 */
function diagnoseProgramRequirementFile_(
  fileId
) {

  try {

    const file =
      DriveApp.getFileById(
        fileId
      );


    console.log(
      'ID: ' +
      file.getId()
    );


    console.log(
      'Name: ' +
      file.getName()
    );


    console.log(
      'MIME: ' +
      file.getMimeType()
    );


    console.log(
      'Created: ' +
      file.getDateCreated()
    );


    console.log(
      'Last updated: ' +
      file.getLastUpdated()
    );


    console.log(
      'URL: ' +
      file.getUrl()
    );


    // ===============================================================
    // PARENT FOLDERS
    // ===============================================================

    const parents =
      file.getParents();


    const parentInfo =
      [];


    while (
      parents.hasNext()
    ) {

      const parent =
        parents.next();


      parentInfo.push(
        parent.getName() +
        ' [' +
        parent.getId() +
        ']'
      );

    }


    console.log(
      'Parent folder(s): ' +
      (
        parentInfo.length
          ? parentInfo.join(' | ')
          : '[NONE]'
      )
    );


    // ===============================================================
    // FIRST PART OF DOCUMENT BODY
    // ===============================================================

    if (
      file.getMimeType() ===
      MimeType.GOOGLE_DOCS
    ) {

      try {

        const doc =
          DocumentApp.openById(
            fileId
          );


        const text =
          doc
            .getBody()
            .getText()
            .replace(
              /\s+/g,
              ' '
            )
            .trim();


        console.log(
          'Document opening text: ' +
          text.substring(
            0,
            350
          )
        );

      } catch (error) {

        console.log(
          'Could not inspect document body: ' +
          error.message
        );

      }

    }

  } catch (error) {

    console.log(
      'ERROR inspecting file ID ' +
      fileId +
      ': ' +
      error.message
    );

  }

}