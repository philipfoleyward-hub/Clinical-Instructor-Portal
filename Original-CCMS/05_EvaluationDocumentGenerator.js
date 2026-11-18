/**
 * =====================================================================
 * CCMS - CLINICAL COMPETENCY EVALUATION DOCUMENT GENERATOR
 * 05_EvaluationDocumentGenerator_V5_2_PERCENT_PAGINATION_FINAL_CANDIDATE.gs
 * =====================================================================
 *
 * Candidate generator for:
 *   Clinical_Competency_Evaluation_TEMPLATE_V5_2_CANDIDATE
 *
 * Template ID:
 *   1w2nHGKLtUynCfP-mcEb3LS9gXas-_BKgpkrcwvj1cTc
 *
 * AUTHORITY
 * ---------
 * - QA-passed 04_ClinicalEvaluationGradingEngine.gs
 * - August 2026 Program Handbook
 * - Established CCMS grading rules
 * - Existing GradeEngine 0-3 -> percentage conversion is display authority
 *   (known scale inaccuracies intentionally NOT revised in this workstream)
 * - Pagination repair preserves V5.2 table geometry/font sizes, removes
 *   generator-created blank spacer paragraphs between image views, and
 *   minimally tightens dynamic View/AF/Final heading paragraph spacing.
 *
 * SCOPE
 * -----
 * - Generates the official Clinical Competency Evaluation Doc + PDF.
 * - Does NOT grade.
 * - Does NOT write Clinical Tracking.
 * - Does NOT modify the Form.
 * - Does NOT modify DriveService.
 *
 * QA
 * --
 * qaGenerateComprehensiveTest001ClinicalCompetencyDocument()
 * creates ONE synthetic TEST-001 AF evaluation exercising:
 * all six performance sections, comments, N/A, three image views,
 * image N/A, Zero-in-View, Missing/Wrong/Illegible marker penalties,
 * Exposure Index/S#, AF reason/comments, Final Score 0.00, FAIL,
 * and overall evaluator comments.
 * =====================================================================
 */

const EVALUATION_DOCUMENT_TEMPLATE_V5_2_ID =
  '1w2nHGKLtUynCfP-mcEb3LS9gXas-_BKgpkrcwvj1cTc';


function evaluationDocument_generate(grade) {

  if (!grade) {
    throw new Error(
      'EvaluationDocumentGenerator requires a grading result.'
    );
  }

  const evaluation = grade.evaluation || {};
  const performance = grade.performance || {};
  const image = grade.image || {};
  const automaticFailure = grade.automaticFailure || {
    occurred: false,
    reasons: [],
    comments: ''
  };

  const studentName =
    String(evaluation.student || '').trim();

  if (!studentName) {
    throw new Error(
      'Evaluation document cannot be generated without a student name.'
    );
  }

  const qaOutput = Boolean(grade.qaOutput);

  const studentRecord = qaOutput
    ? {
        studentId: 'TEST-001',
        student: studentName,
        cohort: 'TEST'
      }
    : evaluationDocument_lookupStudent_(studentName);

  if (!studentRecord) {
    throw new Error(
      'Student was not found on Students tab: ' + studentName
    );
  }

  const cohort =
    String(studentRecord.cohort || '').trim();

  if (!cohort) {
    throw new Error(
      'Student has no Cohort assigned on Students tab: ' +
      studentName
    );
  }

  const templateId =
    EVALUATION_DOCUMENT_TEMPLATE_V5_2_ID;

  const masterFile =
    DriveApp.getFileById(templateId);

  if (
    masterFile.getMimeType() !==
    'application/vnd.google-apps.document'
  ) {
    throw new Error(
      'Clinical Competency Evaluation V5.2 template is not a native Google Doc.'
    );
  }

  const fileName =
    driveService_buildEvaluationFileName({
      date: evaluation.date,
      student: studentName,
      examination: evaluation.examination,
      evaluationType: evaluation.evaluationType
    });

  console.log(
    'Generating evaluation document: ' + fileName
  );

  const documentFile = qaOutput
    ? evaluationDocument_copyTemplateForQa_(
        masterFile,
        fileName
      )
    : driveService_copyTemplateToStudentFolder({
        templateId: templateId,
        fileName: fileName,
        student: studentName,
        cohort: cohort,
        destination: 'evaluations'
      });

  const doc =
    evaluationDocument_openWithRetry_(
      documentFile.fileId
    );

  const body = doc.getBody();

  let tables = body.getTables();

  /*
   * V5.2 pristine table contract:
   *  0 Info
   *  1 Top score summary
   *  2 Exam Preparation
   *  3 Patient Care
   *  4 Radiation Protection
   *  5 Equipment Manipulation
   *  6 Positioning Skills
   *  7 Portable / Contrast
   *  8 View 1 prototype
   *  9 View 2 prototype
   * 10 View 3 prototype
   * 11 Image Summary
   * 12 Automatic Failure Details
   * 13 Final Evaluation Summary
   */
  if (tables.length < 14) {
    throw new Error(
      'V5.2 template contract failure. Expected at least 14 tables; found ' +
      tables.length + '.'
    );
  }

  // -------------------------------------------------------------------
  // TOP EVALUATION INFORMATION
  // -------------------------------------------------------------------

  const infoTable = tables[0];

  evaluationDocument_setCell_(
    infoTable, 0, 1, evaluation.student || ''
  );
  evaluationDocument_setCell_(
    infoTable, 0, 3, evaluation.clinicalSite || ''
  );
  evaluationDocument_setCell_(
    infoTable, 1, 1, evaluation.evaluator || ''
  );
  evaluationDocument_setCell_(
    infoTable, 1, 3,
    evaluationDocument_formatDate_(evaluation.date)
  );
  evaluationDocument_setCell_(
    infoTable, 2, 1,
    evaluation.examination || 'Not captured'
  );
  evaluationDocument_setCell_(
    infoTable, 2, 3,
    evaluation.evaluationType || ''
  );
  evaluationDocument_setCell_(
    infoTable, 3, 1,
    automaticFailure.occurred ? 'YES' : 'No'
  );
  evaluationDocument_setCell_(
    infoTable, 3, 3,
    String(grade.result || '').toUpperCase()
  );
  evaluationDocument_styleResultCell_(
    infoTable.getRow(3).getCell(3),
    grade.result
  );

  // -------------------------------------------------------------------
  // TOP SCORE SUMMARY
  // -------------------------------------------------------------------

  const topScoreTable = tables[1];

  evaluationDocument_setCell_(
    topScoreTable, 1, 0,
    evaluationDocument_gradeDisplay_(performance.score)
  );
  evaluationDocument_setCell_(
    topScoreTable, 1, 1,
    evaluationDocument_gradeDisplay_(image.score)
  );
  evaluationDocument_setCell_(
    topScoreTable, 1, 2,
    evaluationDocument_gradeDisplay_(grade.finalScore)
  );
  evaluationDocument_setCell_(
    topScoreTable, 1, 3,
    String(grade.result || '').toUpperCase()
  );
  evaluationDocument_styleResultCell_(
    topScoreTable.getRow(1).getCell(3),
    grade.result
  );

  // -------------------------------------------------------------------
  // PERFORMANCE SECTIONS
  // -------------------------------------------------------------------

  const performanceTables = [
    tables[2],
    tables[3],
    tables[4],
    tables[5],
    tables[6],
    tables[7]
  ];

  const sections =
    performance.sections || [];

  if (
    sections.length !==
    performanceTables.length
  ) {
    throw new Error(
      'Expected 6 performance sections but grading engine returned ' +
      sections.length + '.'
    );
  }

  const performanceCommentTokens = {
    'exam preparation':
      '{{EXAM_PREPARATION_COMMENTS}}',
    'patient care':
      '{{PATIENT_CARE_COMMENTS}}',
    'radiation protection':
      '{{RADIATION_PROTECTION_COMMENTS}}',
    'equipment manipulation':
      '{{EQUIPMENT_MANIPULATION_COMMENTS}}',
    'positioning skills':
      '{{POSITIONING_SKILLS_COMMENTS}}',
    'portable / contrast':
      '{{PORTABLE_CONTRAST_COMMENTS}}'
  };

  for (
    let sectionIndex = 0;
    sectionIndex < sections.length;
    sectionIndex++
  ) {
    const section = sections[sectionIndex];

    evaluationDocument_fillPerformanceTable_(
      performanceTables[sectionIndex],
      section
    );

    evaluationDocument_updatePerformanceHeading_(
      body,
      section
    );

    const commentToken =
      performanceCommentTokens[
        appNormalize_(section.name)
      ];

    if (commentToken) {
      evaluationDocument_replaceToken_(
        body,
        commentToken,
        section.comments || 'None'
      );
    }
  }

  /*
   * V5.2 intentionally keeps Portable / Contrast on every evaluation.
   * This is now a compact, stable 4-page template and avoids fragile
   * document surgery while preserving all six authoritative sections.
   */

  evaluationDocument_replaceParagraphStarting_(
    body,
    'Performance Score Calculation:',
    'Performance Score Calculation: ' +
      performance.totalPoints +
      ' weighted points / ' +
      performance.totalWeight +
      ' applicable weight = ' +
      evaluationDocument_scoreText_(
        performance.score
      )
  );

  // -------------------------------------------------------------------
  // IMAGE VIEW PROTOTYPES
  // -------------------------------------------------------------------

  tables = body.getTables();

  const sampleViewTable = tables[8];
  const sampleViewTableCopy =
    sampleViewTable.copy();

  const sampleViewTitle =
    evaluationDocument_findParagraphStarting_(
      body,
      'View 1 -'
    );

  if (!sampleViewTitle) {
    throw new Error(
      'Could not locate View 1 title in V5.2 template.'
    );
  }

  const sampleTitleCopy =
    sampleViewTitle.copy();

  /*
   * Save the Image Summary table reference BEFORE removing prototypes.
   */
  let imageSummaryTable = tables[11];

  evaluationDocument_removeViewBlock_(
    body,
    'View 3 -',
    tables[10]
  );
  evaluationDocument_removeViewBlock_(
    body,
    'View 2 -',
    tables[9]
  );
  evaluationDocument_removeViewBlock_(
    body,
    'View 1 -',
    tables[8]
  );

  let insertionIndex =
    body.getChildIndex(imageSummaryTable);

  const views = image.views || [];

  if (views.length < 1 || views.length > 8) {
    throw new Error(
      'Expected 1-8 completed Image Evaluation views. Found: ' +
      views.length
    );
  }

  // -------------------------------------------------------------------
  // INSERT ACTUAL IMAGE VIEWS
  // -------------------------------------------------------------------

  for (
    let viewIndex = 0;
    viewIndex < views.length;
    viewIndex++
  ) {
    const view = views[viewIndex];

    /*
     * Natural pagination is intentional in V5.2.
     * We do NOT insert a hard page break after every two views.
     * This allows Google Docs/PDF to respond to actual comment length.
     */

    const titleParagraph =
      sampleTitleCopy
        .copy()
        .asParagraph();

    titleParagraph.setText(
      'View ' +
      (viewIndex + 1) +
      ' - ' +
      (
        view.projection ||
        'Not specified'
      )
    );

    titleParagraph
      .setBold(true)
      .setFontSize(11)
      .setSpacingBefore(0)
      .setSpacingAfter(2);

    body.insertParagraph(
      insertionIndex,
      titleParagraph
    );

    insertionIndex++;

    const viewTable =
      sampleViewTableCopy
        .copy()
        .asTable();

    body.insertTable(
      insertionIndex,
      viewTable
    );

    evaluationDocument_fillImageViewTable_(
      viewTable,
      view
    );

    insertionIndex++;

    /*
     * No spacer paragraph after generated view tables.
     * V5.2 already provides adequate visual separation through its
     * headings/tables. Removing these three synthetic blank lines keeps
     * the Final Summary + Evaluator Comments together on page 4 without
     * shrinking fonts or altering table geometry.
     */
  }

  // -------------------------------------------------------------------
  // IMAGE SUMMARY - V5.2 is a 4 x 4 table
  // -------------------------------------------------------------------

  imageSummaryTable =
    evaluationDocument_findTableByFirstCell_(
      body,
      'Views Evaluated'
    );

  if (!imageSummaryTable) {
    throw new Error(
      'Could not locate Image Scoring Summary table.'
    );
  }

  evaluationDocument_setCell_(
    imageSummaryTable, 1, 0,
    String(views.length)
  );
  evaluationDocument_setCell_(
    imageSummaryTable, 1, 1,
    String(image.grossPoints || 0)
  );
  evaluationDocument_setCell_(
    imageSummaryTable, 1, 2,
    String(image.netPoints || 0) +
      ' / ' +
      String(image.sectionsScored || 0)
  );
  evaluationDocument_setCell_(
    imageSummaryTable, 1, 3,
    evaluationDocument_gradeDisplay_(image.score)
  );

  evaluationDocument_setCell_(
    imageSummaryTable, 3, 0,
    String(image.zeroPenalty || 0)
  );
  evaluationDocument_setCell_(
    imageSummaryTable, 3, 1,
    String(image.missingMarkerPenalty || 0)
  );
  evaluationDocument_setCell_(
    imageSummaryTable, 3, 2,
    String(image.wrongMarkerPenalty || 0)
  );
  evaluationDocument_setCell_(
    imageSummaryTable, 3, 3,
    String(image.illegibleMarkerPenalty || 0)
  );

  // -------------------------------------------------------------------
  // AUTOMATIC FAILURE DETAILS
  // -------------------------------------------------------------------

  const afTable =
    evaluationDocument_findTableByFirstCell_(
      body,
      'Automatic Failure'
    );

  if (!afTable) {
    throw new Error(
      'Could not locate Automatic Failure Details table.'
    );
  }

  if (automaticFailure.occurred) {

    evaluationDocument_setCell_(
      afTable, 0, 1, 'YES'
    );

    evaluationDocument_setCell_(
      afTable, 1, 1,
      automaticFailure.reasons &&
      automaticFailure.reasons.length
        ? automaticFailure.reasons.join('; ')
        : 'Not specified'
    );

    evaluationDocument_setCell_(
      afTable, 2, 1,
      automaticFailure.comments || 'None'
    );

  } else {

    evaluationDocument_removeAutomaticFailureBlock_(
      body,
      afTable
    );
  }

  // -------------------------------------------------------------------
  // FINAL CALCULATION LINE
  // -------------------------------------------------------------------

  const finalCalculation =
    automaticFailure.occurred
      ? (
          'Performance Score: ' +
          evaluationDocument_gradeDisplay_(
            performance.score
          ) +
          '     |     Image Score: ' +
          evaluationDocument_gradeDisplay_(
            image.score
          ) +
          '     |     Final Score: ' +
          evaluationDocument_gradeDisplay_(0) +
          ' - Automatic Failure Override'
        )
      : (
          'Performance: ' +
          performance.totalPoints +
          ' / ' +
          performance.totalWeight +
          ' = ' +
          evaluationDocument_gradeDisplay_(
            performance.score
          ) +
          '     |     Image: ' +
          image.netPoints +
          ' / ' +
          image.sectionsScored +
          ' = ' +
          evaluationDocument_gradeDisplay_(
            image.score
          ) +
          '     |     Final: (' +
          evaluationDocument_scoreText_(
            performance.score
          ) +
          ' + ' +
          evaluationDocument_scoreText_(
            image.score
          ) +
          ') / 2 = ' +
          evaluationDocument_gradeDisplay_(
            grade.finalScore
          )
        );

  if (
    !evaluationDocument_replaceToken_(
      body,
      '{{FINAL_CALCULATION_LINE}}',
      finalCalculation
    )
  ) {
    /*
     * Backward-safe fallback if the converted Google Doc changed the
     * placeholder paragraph during import.
     */
    evaluationDocument_replaceParagraphStarting_(
      body,
      'Performance:',
      finalCalculation
    );
  }

  // -------------------------------------------------------------------
  // FINAL EVALUATION SUMMARY
  // -------------------------------------------------------------------

  const finalTable =
    evaluationDocument_findFinalEvaluationTable_(
      body
    );

  if (!finalTable) {
    throw new Error(
      'Could not locate Final Evaluation Summary table.'
    );
  }

  const performancePass =
    performance.score !== null &&
    performance.score !== undefined &&
    Number(performance.score) >= 2.0;

  const imagePass =
    image.score !== null &&
    image.score !== undefined &&
    Number(image.score) >= 2.0;

  evaluationDocument_setCell_(
    finalTable, 0, 1,
    evaluationDocument_gradeDisplay_(
      performance.score
    )
  );
  evaluationDocument_setCell_(
    finalTable, 0, 2,
    performancePass ? 'YES' : 'NO'
  );

  evaluationDocument_setCell_(
    finalTable, 1, 1,
    evaluationDocument_gradeDisplay_(
      image.score
    )
  );
  evaluationDocument_setCell_(
    finalTable, 1, 2,
    imagePass ? 'YES' : 'NO'
  );

  evaluationDocument_setCell_(
    finalTable, 2, 1,
    evaluationDocument_gradeDisplay_(
      grade.finalScore
    )
  );

  evaluationDocument_setCell_(
    finalTable, 2, 2,
    automaticFailure.occurred
      ? 'Automatic Failure Override'
      : '(Performance + Image) / 2'
  );

  evaluationDocument_setCell_(
    finalTable, 3, 1,
    String(grade.result || '').toUpperCase()
  );

  evaluationDocument_styleResultCell_(
    finalTable.getRow(3).getCell(1),
    grade.result
  );

  // -------------------------------------------------------------------
  // OVERALL EVALUATOR COMMENTS
  // -------------------------------------------------------------------

  evaluationDocument_replaceParagraphAfterHeading_(
    body,
    'Evaluator Comments',
    evaluation.comments ||
      'No additional comments.'
  );

  // -------------------------------------------------------------------
  // PDF PAGINATION MICRO-ADJUSTMENT
  // -------------------------------------------------------------------
  /*
   * Recover a small amount of vertical space on the dynamic final page
   * without changing font sizes, tables, margins, or document content.
   */
  evaluationDocument_tightenHeadingSpacing_(
    body,
    'AUTOMATIC FAILURE DETAILS',
    2,
    2
  );

  evaluationDocument_tightenHeadingSpacing_(
    body,
    'Final Evaluation Summary',
    2,
    2
  );

  // -------------------------------------------------------------------
  // UNRESOLVED TOKEN QA
  // -------------------------------------------------------------------

  const unresolvedTokens =
    evaluationDocument_findUnresolvedTokens_(body);

  if (unresolvedTokens.length) {
    throw new Error(
      'Generated evaluation still contains unresolved template tokens: ' +
      unresolvedTokens.join(', ')
    );
  }

  // -------------------------------------------------------------------
  // ELECTRONIC RECORD NOTE
  // -------------------------------------------------------------------

  const electronicNoteText =
    'Electronically submitted evaluation. ' +
    'Evaluator identity and submission timestamp are retained ' +
    'with the electronic record.';

  let electronicNote =
    evaluationDocument_findParagraphStarting_(
      body,
      electronicNoteText
    );

  if (!electronicNote) {
    electronicNote =
      body.appendParagraph(
        electronicNoteText
      );
  }

  electronicNote
    .setAlignment(
      DocumentApp.HorizontalAlignment.CENTER
    )
    .setItalic(true)
    .setFontSize(8)
    .setForegroundColor('#666666');

  /*
   * Synthetic QA value gate.
   * This verifies the populated Google Doc body before PDF export.
   */
  if (qaOutput) {
    evaluationDocument_assertComprehensiveQaBody_(
      body,
      grade
    );
  }

  doc.saveAndClose();

  Utilities.sleep(1500);

  const finishedDoc =
    DriveApp.getFileById(
      documentFile.fileId
    );

  const pdfBlob =
    finishedDoc
      .getAs(MimeType.PDF)
      .setName(fileName + '.pdf');

  const pdfFile = qaOutput
    ? evaluationDocument_saveBlobForQa_(
        pdfBlob,
        fileName + '.pdf'
      )
    : driveService_saveBlob({
        blob: pdfBlob,
        student: studentName,
        cohort: cohort,
        destination: 'evaluations',
        fileName: fileName + '.pdf'
      });

  const result = {
    success: true,
    fileName: fileName,
    cohort: cohort,
    docId: documentFile.fileId,
    docUrl: documentFile.fileUrl,
    pdfId: pdfFile.fileId,
    pdfUrl: pdfFile.fileUrl,
    pdfBlob: pdfBlob,
    grade: grade
  };

  console.log(
    JSON.stringify(
      {
        success: true,
        fileName: fileName,
        docUrl: documentFile.fileUrl,
        pdfUrl: pdfFile.fileUrl,
        automaticFailure:
          automaticFailure.occurred,
        finalScore: grade.finalScore,
        result: grade.result
      },
      null,
      2
    )
  );

  return result;
}


/**
 * =====================================================================
 * STUDENT LOOKUP
 * =====================================================================
 */
function evaluationDocument_lookupStudent_(
  studentName
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const studentsSheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENTS
    );

  const students =
    appSheetToObjects_(
      studentsSheet
    );

  const target =
    appNormalize_(studentName);

  const matches =
    students.filter(row =>
      appNormalize_(
        row['Student']
      ) === target
    );

  if (!matches.length) {
    return null;
  }

  const active =
    matches.find(row =>
      appIsTrue_(row['Active'])
    );

  const row =
    active || matches[0];

  return {
    studentId:
      String(
        row['Student ID'] || ''
      ).trim(),
    student:
      String(
        row['Student'] || ''
      ).trim(),
    cohort:
      String(
        row['Cohort'] || ''
      ).trim(),
    clinicalTerm:
      String(
        row['Clinical Term'] || ''
      ).trim()
  };
}


/**
 * =====================================================================
 * PERFORMANCE TABLE
 * =====================================================================
 */
function evaluationDocument_fillPerformanceTable_(
  table,
  section
) {

  const rows =
    section.rows || [];

  const expectedRows =
    table.getNumRows() - 1;

  if (expectedRows !== rows.length) {
    throw new Error(
      'Performance table row count mismatch for "' +
      section.name +
      '". Template rows=' +
      expectedRows +
      ', grading rows=' +
      rows.length +
      '.'
    );
  }

  for (
    let i = 0;
    i < rows.length;
    i++
  ) {

    const row = rows[i];
    const tableRow = i + 1;

    evaluationDocument_setCell_(
      table,
      tableRow,
      1,
      row.applicable
        ? String(row.ratingText)
        : 'N/A'
    );

    evaluationDocument_setCell_(
      table,
      tableRow,
      2,
      row.applicable
        ? String(row.weight)
        : 'N/A'
    );

    evaluationDocument_setCell_(
      table,
      tableRow,
      3,
      row.applicable
        ? String(row.points)
        : 'N/A'
    );
  }
}


/**
 * =====================================================================
 * PERFORMANCE SECTION HEADING
 * =====================================================================
 */
function evaluationDocument_updatePerformanceHeading_(
  body,
  section
) {

  const scoreText =
    evaluationDocument_scoreText_(
      section.score
    );

  const points =
    section.points || 0;

  const weight =
    section.weight || 0;

  const target =
    String(section.name || '')
      .toUpperCase();

  for (
    let i = 0;
    i < body.getNumChildren();
    i++
  ) {

    const child =
      body.getChild(i);

    if (
      child.getType() !==
      DocumentApp.ElementType.PARAGRAPH
    ) {
      continue;
    }

    const paragraph =
      child.asParagraph();

    const text =
      paragraph.getText().trim();

    if (
      text.toUpperCase().indexOf(
        target
      ) !== 0
    ) {
      continue;
    }

    paragraph.setText(
      target +
      '    SECTION SCORE ' +
      scoreText
    );

    paragraph
      .setBold(true)
      .setFontSize(11)
      .setForegroundColor('#1F4E79');

    /*
     * The points/weight line immediately follows the heading in V5.2.
     */
    if (
      i + 1 < body.getNumChildren()
    ) {

      const next =
        body.getChild(i + 1);

      if (
        next.getType() ===
        DocumentApp.ElementType.PARAGRAPH
      ) {
        next
          .asParagraph()
          .setText(
            points +
            ' weighted points / ' +
            weight +
            ' applicable weight'
          );
      }
    }

    return true;
  }

  console.log(
    'WARNING: Performance heading not found: ' +
    section.name
  );

  return false;
}


/**
 * =====================================================================
 * IMAGE VIEW TABLE - V5.2
 * =====================================================================
 */
function evaluationDocument_fillImageViewTable_(
  table,
  view
) {

  const components =
    view.components || [];

  const componentMap = {};

  components.forEach(
    component => {
      componentMap[
        appNormalize_(component.name)
      ] = component;
    }
  );

  evaluationDocument_fillImageComponentRow_(
    table,
    1,
    componentMap['positioning']
  );

  evaluationDocument_fillImageComponentRow_(
    table,
    2,
    componentMap[
      'equipment / central ray'
    ]
  );

  evaluationDocument_fillImageComponentRow_(
    table,
    3,
    componentMap[
      'radiation protection / artifacts'
    ]
  );

  evaluationDocument_fillImageComponentRow_(
    table,
    4,
    componentMap['exposure']
  );

  evaluationDocument_setCell_(
    table,
    5,
    1,
    view.exposureIndex ||
      'Not entered'
  );

  evaluationDocument_setCell_(
    table,
    5,
    2,
    'Marker: ' +
      (
        view.marker ||
        'Not answered'
      )
  );

  /*
   * In the imported V5.2 Google Doc the View Comments row has
   * three cells. Put the comment in cell 1 and leave cell 2 blank.
   */
  evaluationDocument_setCell_(
    table,
    6,
    1,
    view.comments || 'None'
  );

  evaluationDocument_setCell_(
    table,
    6,
    2,
    ''
  );

  /*
   * V5.2 intentionally uses one compact calculation string so all four
   * penalty types remain visible and cannot drift into the wrong column.
   */
  const calculationText =
    'Criterion points ' +
    String(view.grossPoints || 0) +
    ' | Zero-in-View ' +
    String(view.zeroPenalty || 0) +
    ' | Missing marker ' +
    String(view.missingMarkerPenalty || 0) +
    ' | Wrong marker ' +
    String(view.wrongMarkerPenalty || 0) +
    ' | Illegible/poor marker ' +
    String(view.illegibleMarkerPenalty || 0) +
    ' | Net ' +
    String(view.netPoints || 0) +
    ' / ' +
    String(view.sectionsScored || 0) +
    ' | Raw view score ' +
    evaluationDocument_scoreText_(
      view.imageScore
    );

  evaluationDocument_setCell_(
    table,
    7,
    1,
    calculationText
  );

  if (
    table.getRow(7).getNumCells() > 2
  ) {
    evaluationDocument_setCell_(
      table,
      7,
      2,
      ''
    );
  }
}


/**
 * =====================================================================
 * IMAGE COMPONENT ROW
 * =====================================================================
 */
function evaluationDocument_fillImageComponentRow_(
  table,
  rowIndex,
  component
) {

  if (!component) {

    evaluationDocument_setCell_(
      table,
      rowIndex,
      1,
      'N/A'
    );

    evaluationDocument_setCell_(
      table,
      rowIndex,
      2,
      'Not Applicable'
    );

    return;
  }

  if (!component.applicable) {

    evaluationDocument_setCell_(
      table,
      rowIndex,
      1,
      'N/A'
    );

    evaluationDocument_setCell_(
      table,
      rowIndex,
      2,
      component.explanation ||
        'Not Applicable'
    );

    return;
  }

  evaluationDocument_setCell_(
    table,
    rowIndex,
    1,
    String(component.score)
  );

  evaluationDocument_setCell_(
    table,
    rowIndex,
    2,
    component.explanation || ''
  );
}


/**
 * =====================================================================
 * AUTOMATIC FAILURE BLOCK REMOVAL
 * =====================================================================
 *
 * V5.2 contains a designed AF heading + table.
 * If AF did not occur, remove BOTH so normal evaluations do not contain
 * an empty Automatic Failure section.
 */
function evaluationDocument_removeAutomaticFailureBlock_(
  body,
  afTable
) {

  if (
    afTable &&
    afTable.getParent()
  ) {
    body.removeChild(afTable);
  }

  const heading =
    evaluationDocument_findParagraphStarting_(
      body,
      'AUTOMATIC FAILURE DETAILS'
    );

  if (heading) {
    body.removeChild(heading);
  }
}


/**
 * =====================================================================
 * REMOVE SAMPLE IMAGE VIEW
 * =====================================================================
 */
function evaluationDocument_removeViewBlock_(
  body,
  titleStart,
  table
) {

  if (
    table &&
    table.getParent()
  ) {
    body.removeChild(table);
  }

  const paragraph =
    evaluationDocument_findParagraphStarting_(
      body,
      titleStart
    );

  if (paragraph) {
    body.removeChild(paragraph);
  }
}


/**
 * =====================================================================
 * PAGINATION - TARGETED HEADING SPACING
 * =====================================================================
 *
 * PDF-only micro-adjustment. Preserves font size, table geometry,
 * margins, content, and the V5.2 visual design.
 */
function evaluationDocument_tightenHeadingSpacing_(
  body,
  headingText,
  before,
  after
) {

  const paragraph =
    evaluationDocument_findParagraphStarting_(
      body,
      headingText
    );

  if (!paragraph) {
    console.log(
      'WARNING: Could not tighten heading spacing: ' +
      headingText
    );
    return false;
  }

  paragraph
    .setSpacingBefore(before)
    .setSpacingAfter(after);

  return true;
}


/**
 * =====================================================================
 * PARAGRAPH HELPERS
 * =====================================================================
 */
function evaluationDocument_findParagraphStarting_(
  body,
  startText
) {

  for (
    let i = 0;
    i < body.getNumChildren();
    i++
  ) {

    const child =
      body.getChild(i);

    if (
      child.getType() !==
      DocumentApp.ElementType.PARAGRAPH
    ) {
      continue;
    }

    const paragraph =
      child.asParagraph();

    const text =
      paragraph.getText().trim();

    if (
      text.indexOf(startText) === 0
    ) {
      return paragraph;
    }
  }

  return null;
}


function evaluationDocument_replaceParagraphStarting_(
  body,
  startText,
  newText
) {

  const paragraph =
    evaluationDocument_findParagraphStarting_(
      body,
      startText
    );

  if (!paragraph) {
    console.log(
      'WARNING: Paragraph not found: ' +
      startText
    );
    return false;
  }

  paragraph.setText(newText);
  return true;
}


function evaluationDocument_replaceParagraphAfterHeading_(
  body,
  headingText,
  newText
) {

  for (
    let i = 0;
    i < body.getNumChildren() - 1;
    i++
  ) {

    const child =
      body.getChild(i);

    if (
      child.getType() !==
      DocumentApp.ElementType.PARAGRAPH
    ) {
      continue;
    }

    if (
      child
        .asParagraph()
        .getText()
        .trim() !==
      headingText
    ) {
      continue;
    }

    const next =
      body.getChild(i + 1);

    if (
      next.getType() ===
      DocumentApp.ElementType.PARAGRAPH
    ) {

      next
        .asParagraph()
        .setText(newText);

      return true;
    }
  }

  console.log(
    'WARNING: Could not populate paragraph after heading: ' +
    headingText
  );

  return false;
}


/**
 * Replace a literal {{TOKEN}} anywhere in the document body.
 */
function evaluationDocument_replaceToken_(
  body,
  token,
  value
) {

  const escaped =
    token.replace(
      /([\\^$.*+?()[\]{}|])/g,
      '\\$1'
    );

  const found =
    body.findText(escaped);

  if (!found) {
    console.log(
      'WARNING: Template token not found: ' +
      token
    );
    return false;
  }

  body.replaceText(
    escaped,
    String(
      value === null ||
      value === undefined
        ? ''
        : value
    )
  );

  return true;
}


/**
 * =====================================================================
 * TABLE HELPERS
 * =====================================================================
 */
function evaluationDocument_findTableByFirstCell_(
  body,
  firstCellText
) {

  const tables =
    body.getTables();

  for (
    let i = 0;
    i < tables.length;
    i++
  ) {

    const table =
      tables[i];

    if (
      table.getNumRows() < 1 ||
      table.getRow(0).getNumCells() < 1
    ) {
      continue;
    }

    const text =
      table
        .getRow(0)
        .getCell(0)
        .getText()
        .trim();

    if (text === firstCellText) {
      return table;
    }
  }

  return null;
}


function evaluationDocument_findFinalEvaluationTable_(
  body
) {

  const tables =
    body.getTables();

  for (
    let i = 0;
    i < tables.length;
    i++
  ) {

    const table =
      tables[i];

    if (
      table.getNumRows() !== 4 ||
      table.getRow(0).getNumCells() !== 3
    ) {
      continue;
    }

    const firstCell =
      table
        .getRow(0)
        .getCell(0)
        .getText()
        .trim();

    if (
      firstCell ===
      'Performance Score'
    ) {
      return table;
    }
  }

  return null;
}


function evaluationDocument_setCell_(
  table,
  row,
  column,
  value
) {

  const cell =
    table
      .getRow(row)
      .getCell(column);

  const displayValue =
    String(
      value === null ||
      value === undefined
        ? ''
        : value
    );

  cell.setText(displayValue);

  if (displayValue) {

    const fontSize =
      displayValue.length > 120
        ? 7
        : displayValue.length > 90
          ? 8
          : displayValue.length > 45
            ? 9
            : 10;

    cell
      .editAsText()
      .setFontSize(fontSize);
  }
}


/**
 * =====================================================================
 * QA - UNRESOLVED TEMPLATE TOKENS
 * =====================================================================
 */
function evaluationDocument_findUnresolvedTokens_(
  body
) {

  const text =
    body.getText();

  const matches =
    text.match(
      /\{\{[^{}]+\}\}/g
    ) || [];

  return Array.from(
    new Set(matches)
  ).sort();
}


/**
 * =====================================================================
 * COMPREHENSIVE SYNTHETIC QA - GENERATED BODY ASSERTIONS
 * =====================================================================
 */
function evaluationDocument_assertComprehensiveQaBody_(
  body,
  grade
) {

  const text = body.getText();

  const required = [
    'TEST-001 Test Student',
    'QA EP COMMENT',
    'QA PC COMMENT',
    'QA RP COMMENT',
    'QA EM COMMENT',
    'QA PS COMMENT',
    'QA PORT COMMENT',
    'QA VIEW 1 COMMENT',
    'QA VIEW 2 COMMENT',
    'QA VIEW 3 COMMENT',
    'Wrong Marker',
    'Missing Marker',
    'Illegible Marker',
    'QA TEST - Student risked their own health or safety.',
    'QA AF COMMENT',
    'QA OVERALL EVALUATOR COMMENT',
    evaluationDocument_gradeDisplay_(grade.performance.score),
    evaluationDocument_gradeDisplay_(grade.image.score),
    evaluationDocument_gradeDisplay_(grade.finalScore),
    'Automatic Failure Override',
    'FAIL'
  ];

  const missing = required.filter(
    value => text.indexOf(String(value)) === -1
  );

  if (missing.length) {
    throw new Error(
      'Comprehensive document QA failed. Missing generated values: ' +
      missing.join(' | ')
    );
  }

  console.log(
    'Comprehensive generated-body QA passed: all required values present.'
  );
}


/**
 * =====================================================================
 * QA OUTPUT FOLDER
 * =====================================================================
 */
function evaluationDocument_getQaFolder_() {

  const name =
    'Generated Document QA - TEST-001';

  const folders =
    DriveApp.getFoldersByName(name);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(name);
}


function evaluationDocument_copyTemplateForQa_(
  masterFile,
  fileName
) {

  const copy =
    masterFile.makeCopy(
      fileName,
      evaluationDocument_getQaFolder_()
    );

  return {
    fileId: copy.getId(),
    fileUrl: copy.getUrl(),
    fileName: copy.getName()
  };
}


function evaluationDocument_saveBlobForQa_(
  blob,
  fileName
) {

  const file =
    evaluationDocument_getQaFolder_()
      .createFile(
        blob.setName(fileName)
      );

  return {
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    fileName: file.getName()
  };
}


/**
 * =====================================================================
 * RESULT CELL STYLE
 * =====================================================================
 */
function evaluationDocument_styleResultCell_(
  cell,
  result
) {

  const normalized =
    String(result || '')
      .trim()
      .toUpperCase();

  const text =
    cell.editAsText();

  if (normalized === 'PASS') {

    cell.setBackgroundColor(
      '#E2F0D9'
    );

    text
      .setForegroundColor('#006100')
      .setBold(true);

    return;
  }

  if (normalized === 'FAIL') {

    cell.setBackgroundColor(
      '#FCE8E6'
    );

    text
      .setForegroundColor('#B00020')
      .setBold(true);

    return;
  }

  text.setBold(true);
}


/**
 * =====================================================================
 * SCORE DISPLAY
 * =====================================================================
 */
/**
 * =====================================================================
 * SCORE + PERCENT DISPLAY
 * =====================================================================
 *
 * Percentage authority remains GradeEngine.gs.
 * Do NOT duplicate or reinterpret the 0-3 -> percent conversion here.
 *
 * Example: 2.00 • 75%
 */
function evaluationDocument_gradeDisplay_(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === '' ||
    !Number.isFinite(Number(value))
  ) {
    return 'N/A';
  }

  if (
    typeof gradeEngine_formatEvaluationGrade !==
    'function'
  ) {
    throw new Error(
      'GradeEngine display helper is unavailable. ' +
      'Expected gradeEngine_formatEvaluationGrade(rawScore).'
    );
  }

  const display =
    gradeEngine_formatEvaluationGrade(
      Number(value)
    );

  return (
    display &&
    display.combinedText
  )
    ? display.combinedText
    : evaluationDocument_scoreText_(value);
}


function evaluationDocument_scoreText_(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === '' ||
    !Number.isFinite(Number(value))
  ) {
    return 'N/A';
  }

  return Number(value).toFixed(2);
}


/**
 * =====================================================================
 * DATE DISPLAY
 * =====================================================================
 */
function evaluationDocument_formatDate_(
  value
) {

  const date =
    appDateOnly_(value);

  if (!date) {
    return String(value || '').trim();
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'MM/dd/yyyy'
  );
}


/**
 * =====================================================================
 * OPEN DOCUMENT WITH RETRY
 * =====================================================================
 */
function evaluationDocument_openWithRetry_(
  documentId
) {

  const maxAttempts = 6;
  let lastError = null;

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

      lastError = error;

      console.log(
        'Document open attempt ' +
        attempt +
        ' of ' +
        maxAttempts +
        ' failed.'
      );

      Utilities.sleep(
        1000 * attempt
      );
    }
  }

  throw new Error(
    'Could not open generated evaluation document after ' +
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
 * PURPOSE-BUILT COMPREHENSIVE SYNTHETIC QA
 * =====================================================================
 *
 * Creates EXACTLY ONE TEST-001 document + PDF.
 *
 * Exercises:
 * - all six Performance sections
 * - comments in every Performance section
 * - Performance N/A
 * - exactly three image views
 * - comment on every view
 * - Image N/A
 * - Zero-in-View
 * - Missing Marker
 * - Wrong Marker
 * - Illegible Marker
 * - Exposure Index / S Number
 * - underlying Performance and Image scores
 * - Automatic Failure = YES
 * - AF reason + comment
 * - Final Score = 0.00 • 25% using existing GradeEngine conversion
 * - Result = FAIL
 * - overall Evaluator Comment
 *
 * Does NOT read or submit the Google Form.
 * Does NOT write Clinical Tracking.
 */
function qaGenerateComprehensiveTest001ClinicalCompetencyDocument() {

  const sectionDefinitions = [
    {
      name: 'Exam Preparation',
      count: 4,
      comment:
        'QA EP COMMENT - verify Exam Preparation comments appear here.'
    },
    {
      name: 'Patient Care',
      count: 23,
      comment:
        'QA PC COMMENT - verify Patient Care comments remain attached to this section.'
    },
    {
      name: 'Radiation Protection',
      count: 8,
      comment:
        'QA RP COMMENT - verify Radiation Protection comments appear correctly.'
    },
    {
      name: 'Equipment Manipulation',
      count: 9,
      comment:
        'QA EM COMMENT - verify Equipment Manipulation comments appear correctly.'
    },
    {
      name: 'Positioning Skills',
      count: 8,
      comment:
        'QA PS COMMENT - verify Positioning Skills comments appear correctly.'
    },
    {
      name: 'Portable / Contrast',
      count: 16,
      comment:
        'QA PORT COMMENT - verify Portable / Contrast remains compact and its comments appear.'
    }
  ];

  const performanceSections =
    sectionDefinitions.map(
      (definition, sectionIndex) => {

        const rows = [];
        let points = 0;
        let weight = 0;

        for (
          let rowIndex = 0;
          rowIndex < definition.count;
          rowIndex++
        ) {

          /*
           * One deliberate Performance N/A:
           * Patient Care criterion 23.
           */
          const applicable =
            !(
              definition.name ===
                'Patient Care' &&
              rowIndex === 22
            );

          /*
           * Deliberately varied but mostly passing ratings.
           */
          const rating =
            rowIndex % 5 === 0
              ? 2
              : 3;

          const rowWeight =
            (rowIndex % 3) + 1;

          const rowPoints =
            applicable
              ? rating * rowWeight
              : 0;

          rows.push({
            applicable: applicable,
            ratingText:
              applicable
                ? String(rating)
                : 'N/A',
            weight: rowWeight,
            points: rowPoints
          });

          if (applicable) {
            points += rowPoints;
            weight += rowWeight;
          }
        }

        return {
          name: definition.name,
          rows: rows,
          points: points,
          weight: weight,
          score:
            weight
              ? Math.round(
                  points / weight * 100
                ) / 100
              : null,
          comments: definition.comment
        };
      }
    );

  const totalPoints =
    performanceSections.reduce(
      (sum, section) =>
        sum + section.points,
      0
    );

  const totalWeight =
    performanceSections.reduce(
      (sum, section) =>
        sum + section.weight,
      0
    );

  const performanceScore =
    Math.round(
      totalPoints /
      totalWeight *
      100
    ) / 100;

  function makeComponent(
    name,
    score,
    explanation,
    applicable
  ) {

    return {
      name: name,
      applicable:
        applicable !== false,
      score:
        applicable === false
          ? null
          : score,
      explanation: explanation
    };
  }

  /*
   * View 1:
   * - contains TWO zero criteria but must receive ONE Zero-in-View -5.
   * - Wrong Marker -5.
   */
  const view1 = {
    projection: 'AP',
    components: [
      makeComponent(
        'Positioning',
        0,
        'QA ZERO 1 - gross positioning error.',
        true
      ),
      makeComponent(
        'Equipment / Central Ray',
        0,
        'QA ZERO 2 - second zero in SAME view; Zero-in-View must still be only -5.',
        true
      ),
      makeComponent(
        'Radiation Protection / Artifacts',
        3,
        'QA satisfactory radiation protection.',
        true
      ),
      makeComponent(
        'Exposure',
        3,
        'QA satisfactory exposure.',
        true
      )
    ],
    exposureIndex: 'S# 210 / EI within QA target',
    marker: 'Wrong Marker',
    comments:
      'QA VIEW 1 COMMENT - two zeros + Wrong Marker. Verify only one Zero-in-View penalty.',
    grossPoints: 6,
    zeroPenalty: -5,
    missingMarkerPenalty: 0,
    wrongMarkerPenalty: -5,
    illegibleMarkerPenalty: 0,
    netPoints: -4,
    sectionsScored: 4,
    imageScore: -1.00
  };

  /*
   * View 2:
   * - Image N/A criterion.
   * - Missing Marker -5.
   */
  const view2 = {
    projection: 'Lateral',
    components: [
      makeComponent(
        'Positioning',
        3,
        'QA optimal positioning.',
        true
      ),
      makeComponent(
        'Equipment / Central Ray',
        3,
        'QA optimal CR.',
        true
      ),
      makeComponent(
        'Radiation Protection / Artifacts',
        2,
        'QA satisfactory.',
        true
      ),
      makeComponent(
        'Exposure',
        null,
        'QA IMAGE N/A - Exposure criterion intentionally not applicable.',
        false
      )
    ],
    exposureIndex: 'N/A - criterion not applicable',
    marker: 'Missing Marker',
    comments:
      'QA VIEW 2 COMMENT - Missing Marker plus Image N/A criterion.',
    grossPoints: 8,
    zeroPenalty: 0,
    missingMarkerPenalty: -5,
    wrongMarkerPenalty: 0,
    illegibleMarkerPenalty: 0,
    netPoints: 3,
    sectionsScored: 3,
    imageScore: 1.00
  };

  /*
   * View 3:
   * - Illegible Marker -3.
   */
  const view3 = {
    projection: 'Decubitus',
    components: [
      makeComponent(
        'Positioning',
        3,
        'QA optimal positioning.',
        true
      ),
      makeComponent(
        'Equipment / Central Ray',
        3,
        'QA optimal CR.',
        true
      ),
      makeComponent(
        'Radiation Protection / Artifacts',
        3,
        'QA optimal radiation protection.',
        true
      ),
      makeComponent(
        'Exposure',
        3,
        'QA optimal exposure.',
        true
      )
    ],
    exposureIndex: 'EI 325 / QA target',
    marker: 'Illegible Marker',
    comments:
      'QA VIEW 3 COMMENT - Illegible Marker penalty must be -3.',
    grossPoints: 12,
    zeroPenalty: 0,
    missingMarkerPenalty: 0,
    wrongMarkerPenalty: 0,
    illegibleMarkerPenalty: -3,
    netPoints: 9,
    sectionsScored: 4,
    imageScore: 2.25
  };

  const views = [
    view1,
    view2,
    view3
  ];

  const grossPoints =
    views.reduce(
      (sum, view) =>
        sum + view.grossPoints,
      0
    );

  const zeroPenalty =
    views.reduce(
      (sum, view) =>
        sum + view.zeroPenalty,
      0
    );

  const missingMarkerPenalty =
    views.reduce(
      (sum, view) =>
        sum + view.missingMarkerPenalty,
      0
    );

  const wrongMarkerPenalty =
    views.reduce(
      (sum, view) =>
        sum + view.wrongMarkerPenalty,
      0
    );

  const illegibleMarkerPenalty =
    views.reduce(
      (sum, view) =>
        sum + view.illegibleMarkerPenalty,
      0
    );

  const netPoints =
    views.reduce(
      (sum, view) =>
        sum + view.netPoints,
      0
    );

  const sectionsScored =
    views.reduce(
      (sum, view) =>
        sum + view.sectionsScored,
      0
    );

  /*
   * Keep the underlying raw Image score visible for documentation.
   * Production Grade Engine floors the stored Image score at 0.
   */
  const rawImageScore =
    netPoints / sectionsScored;

  const storedImageScore =
    Math.max(
      0,
      rawImageScore
    );

  const qaGrade = {
    qaOutput: true,

    responseId:
      'QA-SYNTHETIC-TEST-001-V5-2',

    evaluation: {
      student:
        'TEST-001 Test Student',
      clinicalSite:
        'TEST Clinical Site',
      evaluator:
        'TEST Evaluator - Synthetic QA',
      examination:
        'QA Comprehensive Three-View Examination',
      evaluationType:
        'Competency - Synthetic QA',
      date:
        new Date(2026, 8, 21),
      comments:
        'QA OVERALL EVALUATOR COMMENT - verify this appears in the final Evaluator Comments section.'
    },

    performance: {
      sections:
        performanceSections,
      totalPoints:
        totalPoints,
      totalWeight:
        totalWeight,
      score:
        performanceScore
    },

    image: {
      views:
        views,
      grossPoints:
        grossPoints,
      zeroPenalty:
        zeroPenalty,
      missingMarkerPenalty:
        missingMarkerPenalty,
      wrongMarkerPenalty:
        wrongMarkerPenalty,
      illegibleMarkerPenalty:
        illegibleMarkerPenalty,
      netPoints:
        netPoints,
      sectionsScored:
        sectionsScored,
      rawScore:
        rawImageScore,
      score:
        storedImageScore
    },

    automaticFailure: {
      occurred: true,
      reasons: [
        'QA TEST - Student risked their own health or safety.'
      ],
      comments:
        'QA AF COMMENT - preserve this Automatic Failure explanation in the official document.'
    },

    /*
     * Handbook §13.7:
     * Automatic Failure -> official competency grade 0.0.
     * Underlying Performance/Image scores remain visible as documentation.
     */
    finalScore: 0,

    result: 'FAIL'
  };

  const result =
    evaluationDocument_generate(
      qaGrade
    );

  console.log(
    '================================'
  );
  console.log(
    'COMPREHENSIVE TEST-001 DOCUMENT QA CREATED'
  );
  console.log(
    'Google Doc: ' +
    result.docUrl
  );
  console.log(
    'PDF: ' +
    result.pdfUrl
  );
  const expectedAfDisplay =
    evaluationDocument_gradeDisplay_(0);

  if (expectedAfDisplay !== '0.00 • 25%') {
    throw new Error(
      'QA percentage contract failure. Expected "0.00 • 25%" but GradeEngine returned "' +
      expectedAfDisplay +
      '".'
    );
  }

  console.log(
    'Expected AF Final Score: ' +
    expectedAfDisplay +
    ' - Automatic Failure Override'
  );
  console.log(
    'Expected Result: FAIL'
  );
  console.log(
    'Expected Image penalties: Zero=-5, Missing=-5, Wrong=-5, Illegible=-3'
  );
  console.log(
    '================================'
  );

  return result;
}


/**
 * =====================================================================
 * OPTIONAL FINAL END-TO-END TEST
 * =====================================================================
 *
 * DO NOT RUN until synthetic QA passes.
 *
 * Reads latest real Form response -> QA-passed Grade Engine ->
 * this V5.2 document generator.
 */
function testGenerateLatestEvaluationDocument() {

  const form =
    FormApp.openById(
      APP_CONFIG
        .FORMS
        .DIGITAL_CLINICAL_EVALUATION
    );

  const responses =
    form.getResponses();

  if (!responses.length) {
    throw new Error(
      'The Digital Clinical Evaluation Form has no responses.'
    );
  }

  const formResponse =
    responses[
      responses.length - 1
    ];

  const grade =
    gradeClinicalEvaluation(
      formResponse
    );

  const result =
    evaluationDocument_generate(
      grade
    );

  console.log(
    JSON.stringify(
      {
        success:
          result.success,
        fileName:
          result.fileName,
        docUrl:
          result.docUrl,
        pdfUrl:
          result.pdfUrl
      },
      null,
      2
    )
  );

  return result;
}
