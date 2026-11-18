/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * DASHBOARD SERVER
 * Dashboard.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Server / presentation layer for DashboardUI.html.
 *
 * Dashboard does NOT calculate academic rules itself.
 *
 * AUTHORITATIVE SOURCES
 * ---------------------
 *
 * Students
 *   -> identity / cohort / current semester
 *
 * Clinical Tracking
 *   -> evaluation history
 *
 * Student Requirement Status
 *   -> RequirementsEngine results
 *
 * GradeEngine
 *   -> percentage grades / partial grade status
 *
 * =====================================================================
 */



/**
 * =====================================================================
 * BOOTSTRAP
 * =====================================================================
 */
function getDashboardBootstrap() {

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
        APP_CONFIG.SHEETS
          .CLINICAL_TRACKING
      )

    );


  const activeStudents =
    dashGetUniqueActiveStudents_(
      students
    );


  return {

    students:
      activeStudents

        .map(
          dashStudentListItem_
        )

        .sort(
          (a, b) =>
            a.student.localeCompare(
              b.student
            )
        ),

    reviewStats:
      dashBuildReviewStats_(
        tracking
      ),

    documents:
      dashGetDocumentCapabilities_(),

    addNew:
      dashGetAddNewCapabilities_()

  };
}


/**
 * =====================================================================
 * STUDENT DASHBOARD
 * =====================================================================
 */
function getStudentDashboard(
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
        APP_CONFIG.SHEETS
          .CLINICAL_TRACKING
      )

    );


  const statuses =
    appSheetToObjects_(

      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS
          .STUDENT_REQUIREMENT_STATUS
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


  if (
    studentRows.length === 0
  ) {

    throw new Error(
      'Student not found: ' +
      studentId
    );
  }


  const student =
    dashGetMostAdvancedStudentRow_(
      studentRows
    );


  const studentName =
    String(
      student['Student'] || ''
    ).trim();


  const resolvedStudentId =
    String(
      student['Student ID'] ||
      studentId
    ).trim();


  // ===================================================================
  // STUDENT REQUIREMENT STATUS
  // ===================================================================

  const studentStatuses =
    statuses.filter(row =>

      appNormalize_(
        row['Student ID']
      ) ===
      appNormalize_(
        resolvedStudentId
      )

    );


  // ===================================================================
  // CLINICAL TRACKING
  // ===================================================================

  const studentTracking =
    tracking.filter(record =>

      dashSameStudent_(
        record,
        resolvedStudentId,
        studentName
      )

    );


  // ===================================================================
  // GRADES
  // ===================================================================

  let gradeSummary;


  try {

    gradeSummary =
      gradeEngine_getStudentSummary(
        resolvedStudentId
      );

  } catch (error) {

    console.log(
      'GRADE ENGINE ERROR: ' +
      error.message
    );


    gradeSummary =
      dashEmptyGradeSummary_(
        String(
          student['Clinical Term'] || ''
        ).trim(),
        error.message
      );
  }


  return {

    student:
      dashStudentDetail_(
        student
      ),

    semester:
      dashBuildSemesterSummary_(
        studentStatuses
      ),

    program:
      dashBuildProgramSummary_(
        studentStatuses
      ),

    requirements:
      dashBuildRequirementLists_(
        studentStatuses
      ),

    clinicalHistory:
      dashBuildClinicalHistoryFromService_(
        clinicalHistory_getStudent(
          resolvedStudentId
        )
      ),

    completed:
      dashBuildCompletedCompetencies_(
        studentTracking
      ),

    evaluations:
      dashBuildEvaluationHistory_(
        studentTracking
      ),

    remediation:
      dashBuildRemediationAlerts_(
        studentStatuses
      ),

    grade:
      dashBuildGradeDisplay_(
        gradeSummary
      ),

    documents:
      dashGetDocumentCapabilities_(),

    addNew:
      dashGetAddNewCapabilities_()

  };
}


/**
 * =====================================================================
 * CLINICAL REVIEW ANALYTICS
 * =====================================================================
 *
 * Read-only analytical/query endpoints.
 *
 * IMPORTANT:
 * ClinicalReviewAnalytics.gs exposes:
 *   clinicalReview_getOptions(dimension)
 *   clinicalReview_query(dimension, selectedValue)
 *
 * Dashboard.gs adapts those service results to the UI contract and applies
 * optional result filters. It does not write to Clinical Tracking and does
 * not replace GradeEngine calculations.
 */
function dashboardGetClinicalReviewOptions(reviewBy) {

  const options =
    clinicalReview_getOptions(
      reviewBy
    ) || [];


  return {
    options: options,
    cohortLabel:
      appNormalize_(reviewBy) === 'student'
        ? 'Class of ' + String(
            typeof CLINICAL_REVIEW_CURRENT_COHORT !== 'undefined'
              ? CLINICAL_REVIEW_CURRENT_COHORT
              : '2027'
          )
        : ''
  };
}


function dashboardRunClinicalReview(request) {

  request = request || {};

  const reviewBy =
    String(
      request.reviewBy || ''
    ).trim();

  const selectedValue =
    String(
      request.value || ''
    ).trim();

  if (!reviewBy) {
    throw new Error(
      'Clinical Review category is required.'
    );
  }

  if (!selectedValue) {
    throw new Error(
      'Clinical Review value is required.'
    );
  }


  const base =
    clinicalReview_query(
      reviewBy,
      selectedValue
    );


  const filters =
    dashNormalizeClinicalReviewFilters_(
      request.filters || {}
    );


  const allRows =
    (base.evaluations || [])
      .map(
        dashClinicalReviewRow_
      );


  const filterOptions = {
    clinicalTerms:
      dashUniqueSorted_(
        allRows.map(
          row => row.clinicalTerm
        )
      ),

    evaluationTypes:
      dashUniqueSorted_(
        allRows.map(
          row => row.evaluationType
        )
      )
  };


  const rows =
    allRows
      .filter(row =>
        dashClinicalReviewRowMatchesFilters_(
          row,
          filters
        )
      )
      .sort(
        dashClinicalReviewRowSort_
      );


  const summary =
    dashBuildClinicalReviewSummary_(
      rows
    );


  return {
    reviewBy: base.dimension || reviewBy,
    reviewByLabel:
      base.dimensionLabel ||
      dashClinicalReviewDimensionLabel_(
        reviewBy
      ),

    selectedValue:
      base.selectedValue ||
      selectedValue,

    selectedLabel:
      base.selectedLabel ||
      selectedValue,

    generatedAt:
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'MM/dd/yyyy h:mm a'
      ),

    filters: filters,
    filterOptions: filterOptions,
    summary: summary,
    rows: rows,

    dataCoverageNote:
      summary.totalEvaluations > 0 &&
      summary.numericCount < summary.totalEvaluations
        ? 'Historical evaluation records without numeric scores are included in the evaluation count and detail table but excluded from numeric score calculations.'
        : '',

    description:
      base.description || ''
  };
}


function dashboardGenerateClinicalReviewPdf(request) {

  const result =
    dashboardRunClinicalReview(
      request || {}
    );

  if (!result.rows.length) {
    throw new Error(
      'There are no matching Clinical Review evaluations to export.'
    );
  }


  const safeLabel =
    dashSafeFilenamePart_(
      result.selectedLabel ||
      result.selectedValue ||
      'Report'
    );

  const filename =
    'Clinical Review - ' +
    safeLabel +
    ' - ' +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    ) +
    '.pdf';


  const doc =
    DocumentApp.create(
      'TEMP - ' + filename.replace(/\.pdf$/i, '')
    );


  try {

    const body =
      doc.getBody();


    body
      .appendParagraph(
        'Clinical Review'
      )
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING1
      );


    body
      .appendParagraph(
        result.reviewByLabel +
        ': ' +
        result.selectedLabel
      )
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );


    const activeFilters = [];

    if (
      result.filters.clinicalTerm !== 'all'
    ) {
      activeFilters.push(
        'Clinical Term: ' +
        result.filters.clinicalTerm
      );
    }

    if (
      result.filters.evaluationType !== 'all'
    ) {
      activeFilters.push(
        'Evaluation Type: ' +
        result.filters.evaluationType
      );
    }


    if (activeFilters.length) {
      body.appendParagraph(
        activeFilters.join(' | ')
      );
    }


    body.appendParagraph(
      'Generated: ' +
      result.generatedAt
    );


    const s =
      result.summary;


    body.appendTable([
      [
        'Average Score',
        dashClinicalReviewPercentText_(s.averageScore),
        'Highest Score',
        dashClinicalReviewPercentText_(s.highestScore),
        'Lowest Score',
        dashClinicalReviewPercentText_(s.lowestScore)
      ],
      [
        'Total Evaluations',
        String(s.totalEvaluations),
        'Pass Rate',
        dashClinicalReviewPercentText_(s.passRate),
        'Numeric Coverage',
        dashClinicalReviewPercentText_(s.numericCoverage)
      ]
    ]);


    body.appendParagraph(
      'Passing: ' + s.passCount +
      ' | Failing: ' + s.failCount +
      ' | Automatic failures: ' + s.automaticFailureCount +
      ' | P/F-only records: ' + s.passFailOnlyCount
    );


    if (result.dataCoverageNote) {
      body.appendParagraph(
        result.dataCoverageNote
      );
    }


    body
      .appendParagraph(
        'Matching Evaluations — Highest to Lowest'
      )
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );


    const tableRows = [[
      'Student',
      'Examination',
      'Type',
      'Clinical Site',
      'Evaluator',
      'Term',
      'Date',
      'Score',
      'Result'
    ]];


    result.rows.forEach(row => {
      tableRows.push([
        row.student || '',
        row.examination || '',
        row.evaluationType || '',
        row.clinicalSite || '',
        row.evaluator || '',
        row.clinicalTerm || '',
        row.date || '',
        row.scoreDisplay || '',
        row.result || ''
      ]);
    });


    body.appendTable(
      tableRows
    );


    doc.saveAndClose();


    const file =
      DriveApp.getFileById(
        doc.getId()
      );


    const pdf =
      file
        .getBlob()
        .getAs(
          MimeType.PDF
        )
        .setName(
          filename
        );


    const response = {
      success: true,
      filename: filename,
      mimeType: MimeType.PDF,
      base64:
        Utilities.base64Encode(
          pdf.getBytes()
        )
    };


    file.setTrashed(
      true
    );


    return response;

  } catch (error) {

    try {
      doc.saveAndClose();
    } catch (closeError) {
      // Ignore cleanup error and preserve original exception.
    }


    try {
      DriveApp
        .getFileById(
          doc.getId()
        )
        .setTrashed(
          true
        );
    } catch (trashError) {
      // Ignore cleanup error and preserve original exception.
    }


    throw error;
  }
}


function dashNormalizeClinicalReviewFilters_(filters) {

  const clinicalTerm =
    String(
      filters.clinicalTerm ||
      'all'
    ).trim();

  const evaluationType =
    String(
      filters.evaluationType ||
      'all'
    ).trim();


  return {
    clinicalTerm:
      clinicalTerm || 'all',
    evaluationType:
      evaluationType || 'all'
  };
}


function dashClinicalReviewRow_(item) {

  return {
    submissionId:
      dashText_(
        item.submissionId
      ),

    studentId:
      dashText_(
        item.studentId
      ),

    student:
      dashText_(
        item.student
      ),

    examination:
      dashText_(
        item.examination
      ),

    anatomicalArea:
      dashText_(
        item.anatomicalArea
      ),

    evaluationType:
      dashText_(
        item.evaluationType
      ),

    clinicalSite:
      dashText_(
        item.clinicalSite
      ),

    evaluator:
      dashText_(
        item.evaluator
      ),

    evaluatorType:
      dashText_(
        item.evaluatorType
      ),

    clinicalTerm:
      dashText_(
        item.clinicalTerm
      ),

    attributedClinicalTerm:
      dashText_(
        item.attributedClinicalTerm
      ),

    date:
      dashText_(
        item.date
      ),

    dateValue:
      Number(
        item.dateValue
      ) || 0,

    percentScore:
      dashClinicalReviewNumberOrNull_(
        item.percentScore
      ),

    rawScore:
      dashClinicalReviewNumberOrNull_(
        item.rawScore
      ),

    scoreDisplay:
      dashText_(
        item.gradeDisplay
      ) ||
      '—',

    result:
      dashText_(
        item.result
      ),

    automaticFailure:
      Boolean(
        item.automaticFailure
      ),

    historical:
      Boolean(
        item.historical
      )
  };
}


function dashClinicalReviewRowMatchesFilters_(row, filters) {

  if (
    appNormalize_(filters.clinicalTerm) !== 'all' &&
    appNormalize_(row.clinicalTerm) !==
      appNormalize_(filters.clinicalTerm)
  ) {
    return false;
  }


  if (
    appNormalize_(filters.evaluationType) !== 'all' &&
    appNormalize_(row.evaluationType) !==
      appNormalize_(filters.evaluationType)
  ) {
    return false;
  }


  return true;
}


function dashClinicalReviewRowSort_(a, b) {

  const aNumeric =
    a.percentScore !== null;

  const bNumeric =
    b.percentScore !== null;


  if (
    aNumeric &&
    bNumeric &&
    a.percentScore !== b.percentScore
  ) {
    return b.percentScore - a.percentScore;
  }


  if (aNumeric && !bNumeric) {
    return -1;
  }

  if (!aNumeric && bNumeric) {
    return 1;
  }


  return b.dateValue - a.dateValue;
}


function dashBuildClinicalReviewSummary_(rows) {

  const numeric =
    rows.filter(row =>
      row.percentScore !== null
    );

  const scores =
    numeric.map(row =>
      row.percentScore
    );

  const passCount =
    rows.filter(row =>
      appNormalize_(row.result) === 'pass'
    ).length;

  const failCount =
    rows.filter(row =>
      appNormalize_(row.result) === 'fail'
    ).length;

  const decided =
    passCount + failCount;


  return {
    totalEvaluations:
      rows.length,

    numericCount:
      numeric.length,

    numericCoverage:
      rows.length
        ? dashRound2_(
            numeric.length /
            rows.length *
            100
          )
        : 0,

    averageScore:
      scores.length
        ? dashRound2_(
            scores.reduce(
              (sum, value) =>
                sum + value,
              0
            ) /
            scores.length
          )
        : null,

    highestScore:
      scores.length
        ? Math.max.apply(
            null,
            scores
          )
        : null,

    lowestScore:
      scores.length
        ? Math.min.apply(
            null,
            scores
          )
        : null,

    passRate:
      decided
        ? dashRound2_(
            passCount /
            decided *
            100
          )
        : null,

    passCount: passCount,
    failCount: failCount,

    automaticFailureCount:
      rows.filter(row =>
        row.automaticFailure
      ).length,

    passFailOnlyCount:
      rows.length - numeric.length,

    uniqueStudents:
      dashUniqueCount_(
        rows.map(row =>
          row.studentId ||
          row.student
        )
      ),

    uniqueEvaluators:
      dashUniqueCount_(
        rows.map(row =>
          row.evaluator
        )
      ),

    uniqueSites:
      dashUniqueCount_(
        rows.map(row =>
          row.clinicalSite
        )
      ),

    uniqueExaminations:
      dashUniqueCount_(
        rows.map(row =>
          row.examination
        )
      )
  };
}


function dashUniqueSorted_(values) {

  const seen = {};

  return (values || [])
    .map(value =>
      dashText_(
        value
      )
    )
    .filter(Boolean)
    .filter(value => {

      const key =
        appNormalize_(
          value
        );

      if (seen[key]) {
        return false;
      }

      seen[key] = true;
      return true;
    })
    .sort((a, b) =>
      a.localeCompare(
        b
      )
    );
}


function dashUniqueCount_(values) {

  const seen = {};

  (values || [])
    .forEach(value => {

      const key =
        appNormalize_(
          value
        );

      if (key) {
        seen[key] = true;
      }
    });


  return Object.keys(
    seen
  ).length;
}


function dashClinicalReviewDimensionLabel_(key) {

  const normalized =
    appNormalize_(key)
      .replace(/\s+/g, '');

  const labels = {
    evaluator: 'Evaluator',
    examination: 'Examination',
    exam: 'Examination',
    student: 'Student',
    clinicalsite: 'Clinical Site',
    site: 'Clinical Site'
  };


  return labels[normalized] ||
    dashText_(key);
}


function dashClinicalReviewPercentText_(value) {

  const number =
    dashClinicalReviewNumberOrNull_(
      value
    );

  if (number === null) {
    return '—';
  }


  return dashTrimNumber_(
    number
  ) + '%';
}


function dashClinicalReviewNumberOrNull_(value) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return null;
  }


  const number =
    Number(
      value
    );


  return Number.isFinite(
    number
  )
    ? number
    : null;
}


function dashRound2_(value) {

  const number =
    Number(
      value
    );

  if (!Number.isFinite(number)) {
    return 0;
  }

  return Math.round(
    (number + Number.EPSILON) * 100
  ) / 100;
}


function dashTrimNumber_(value) {

  const number =
    Number(
      value
    );

  if (!Number.isFinite(number)) {
    return '';
  }

  return number
    .toFixed(2)
    .replace(/\.?0+$/, '');
}


function dashSafeFilenamePart_(value) {

  const cleaned =
    String(
      value || ''
    )
      .replace(/[\\/:*?"<>|]+/g, '-')
      .replace(/\s+/g, ' ')
      .trim();


  return cleaned || 'Report';
}


/**
 * =====================================================================
 * REVIEW QUEUE
 * =====================================================================
 *
 * filter:
 *   pending
 *   flagged
 *   failed
 *   reviewedToday
 *   all
 */
function getReviewQueue(
  filter
) {

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

  const normalizedFilter =
    dashNormalizeReviewFilter_(
      filter
    );

  return (tracking || [])

    .filter(record =>
      dashReviewRecordMatchesFilter_(
        record,
        normalizedFilter
      )
    )

    .map(
      dashReviewQueueItem_
    )

    .sort(
      (a, b) =>
        b.dateValue -
        a.dateValue
    );
}


/**
 * =====================================================================
 * REVIEW ACTIONS
 * =====================================================================
 *
 * Dashboard is only a presentation/API layer.
 * All writes are delegated to 04_ClinicalReviewService.gs.
 */
function dashboardMarkEvaluationReviewed(
  submissionId
) {

  return clinicalReview_markReviewed(
    submissionId
  );
}


function dashboardSetEvaluationFlag(
  submissionId,
  flagged
) {

  return clinicalReview_setFlag(
    submissionId,
    flagged
  );
}


/**
 * =====================================================================
 * REVIEW QUEUE ITEM
 * =====================================================================
 */
function dashReviewQueueItem_(
  record
) {

  const grade =
    gradeEngine_formatEvaluationGrade(
      record['Final Score']
    );

  return {

    submissionId:
      dashText_(
        record['Submission ID']
      ),

    studentId:
      dashText_(
        record['Student ID']
      ),

    student:
      dashText_(
        record['Student']
      ),

    clinicalTerm:
      dashText_(
        record['Clinical Term']
      ),

    attributedClinicalTerm:
      dashText_(
        record[
          'Attributed Clinical Term'
        ]
      ),

    examination:
      dashText_(
        record['Examination']
      ),

    anatomicalArea:
      dashText_(
        record['Anatomical Area']
      ),

    evaluationType:
      dashText_(
        record['Evaluation Type']
      ),

    clinicalSite:
      dashText_(
        record['Clinical Site']
      ),

    evaluator:
      dashText_(
        record['Evaluator']
      ),

    evaluatorType:
      dashText_(
        record['Evaluator Type']
      ),

    result:
      dashText_(
        record['Result']
      ),

    rawScore:
      grade.raw,

    rawScoreText:
      grade.rawText,

    percentGrade:
      grade.percent,

    percentGradeText:
      grade.percentText,

    gradeDisplay:
      grade.combinedText,

    automaticFailure:
      appIsTrue_(
        record['Automatic Failure']
      ),

    sourceDocument:
      dashText_(
        record['Source Document']
      ),

    coordinatorFlag:
      dashText_(
        record['Coordinator Flag']
      ),

    flagged:
      dashIsFlagged_(
        record
      ),

    coordinatorReviewStatus:
      dashText_(
        record[
          'Coordinator Review Status'
        ]
      ),

    reviewedBy:
      dashText_(
        record['Reviewed By']
      ),

    reviewDate:
      dashDisplayDate_(
        record['Review Date']
      ),

    date:
      dashDisplayDate_(
        record['Date']
      ),

    dateValue:
      appDateValue_(
        record['Date']
      )

  };
}


/**
 * =====================================================================
 * REVIEW FILTER
 * =====================================================================
 */
function dashNormalizeReviewFilter_(
  filter
) {

  const value =
    appNormalize_(
      filter || 'pending'
    );

  const allowed = [
    'pending',
    'flagged',
    'failed',
    'reviewedtoday',
    'all'
  ];

  return allowed.includes(
    value
  )
    ? value
    : 'pending';
}


function dashReviewRecordMatchesFilter_(
  record,
  filter
) {

  if (
    appNormalize_(
      record['Submission Method']
    ) ===
    'historical import'
  ) {

    return false;
  }

  if (
    filter === 'all'
  ) {

    return true;
  }

  if (
    filter === 'flagged'
  ) {

    return dashIsFlagged_(
      record
    );
  }

  if (
    filter === 'failed'
  ) {

    return (
      appNormalize_(
        record['Result']
      ) ===
      'fail'
    );
  }

  if (
    filter === 'reviewedtoday'
  ) {

    return (
      appNormalize_(
        record[
          'Coordinator Review Status'
        ]
      ) ===
        'reviewed' &&

      dashDateKey_(
        record['Review Date']
      ) ===
        dashTodayKey_()
    );
  }

  return dashIsPendingReview_(
    record
  );
}


/**
 * =====================================================================
 * GRADE DISPLAY
 * =====================================================================
 *
 * This is the contract DashboardUI will use.
 *
 * Semester/program summary cards show PERCENT ONLY.
 *
 * The disclaimer explains whether the grade is partial.
 */
function dashBuildGradeDisplay_(
  gradeSummary
) {

  const semester =
    gradeSummary.semester || {};


  const program =
    gradeSummary.program || {};


  return {

    semesterAverage: {

      available:
        semester.recordedAverage !==
          null &&
        semester.recordedAverage !==
          undefined,

      label:
        semester.complete
          ? 'Current Semester Grade'
          : 'Current Recorded Grade Average',

      value:
        semester.recordedAverageText || '',

      raw:
        semester.recordedAverage,

      complete:
        Boolean(
          semester.complete
        ),

      partial:
        !semester.complete,

      availableWeight:
        semester.availableWeight || 0,

      includedCategories:
        semester.includedCategories || [],

      missingCategories:
        semester.missingCategories || [],

      note:
        semester.note || ''

    },


    programAverage: {

      available:
        Boolean(
          program.available
        ),

      label:
        program.complete
          ? 'Overall Program Average'
          : 'Overall Recorded Program Average',

      value:
        program.recordedAverageText || '',

      raw:
        program.recordedAverage,

      complete:
        Boolean(
          program.complete
        ),

      partial:
        !program.complete,

      semesterCount:
        program.semesterCount || 0,

      includedTerms:
        program.includedTerms || [],

      incompleteTerms:
        program.incompleteTerms || [],

      missingCategories:
        program.missingCategories || [],

      note:
        program.note || ''

    },


    /*
     * Full detail is available for future grade-detail UI.
     */
    semesterDetail:
      semester,

    programDetail:
      program,

    terms:
      gradeSummary.terms || []

  };
}


/**
 * =====================================================================
 * EVALUATION HISTORY
 * =====================================================================
 *
 * Individual grades contain BOTH representations:
 *
 *     2.74 • 90.5%
 */
function dashBuildEvaluationHistory_(
  tracking
) {

  return (tracking || [])

    .map(record => {

      const grade =
        gradeEngine_formatEvaluationGrade(
          record['Final Score']
        );


      return {

        submissionId:
          dashText_(
            record['Submission ID']
          ),

        clinicalTerm:
          dashText_(
            record['Clinical Term']
          ),

        attributedClinicalTerm:
          dashText_(
            record[
              'Attributed Clinical Term'
            ]
          ),

        evaluationType:
          dashText_(
            record['Evaluation Type']
          ),

        examination:
          dashText_(
            record['Examination']
          ),

        anatomicalArea:
          dashText_(
            record['Anatomical Area']
          ),

        evaluator:
          dashText_(
            record['Evaluator']
          ),

        evaluatorType:
          dashText_(
            record['Evaluator Type']
          ),

        clinicalSite:
          dashText_(
            record['Clinical Site']
          ),

        performanceScore:
          dashPreserveZero_(
            record['Performance Score']
          ),

        imageScore:
          dashPreserveZero_(
            record['Image Score']
          ),

        finalScore:
          dashPreserveZero_(
            record['Final Score']
          ),

        rawScore:
          grade.raw,

        rawScoreText:
          grade.rawText,

        percentGrade:
          grade.percent,

        percentGradeText:
          grade.percentText,

        gradeDisplay:
          grade.combinedText,

        result:
          dashText_(
            record['Result']
          ),

        automaticFailure:
          appIsTrue_(
            record['Automatic Failure']
          ),

        submissionMethod:
          dashText_(
            record['Submission Method']
          ),

        sourceDocument:
          dashText_(
            record['Source Document']
          ),

        coordinatorReviewStatus:
          dashText_(
            record[
              'Coordinator Review Status'
            ]
          ),

        coordinatorFlag:
          dashText_(
            record['Coordinator Flag']
          ),

        date:
          dashDisplayDate_(
            record['Date']
          ),

        dateValue:
          appDateValue_(
            record['Date']
          )

      };

    })

    .sort(
      (a, b) =>
        b.dateValue -
        a.dateValue
    );
}


/**
 * =====================================================================
 * COMPLETED COMPETENCIES
 * =====================================================================
 */
function dashBuildCompletedCompetencies_(
  tracking
) {

  const unique = {};


  (tracking || [])
    .forEach(record => {

      if (
        appNormalize_(
          record['Evaluation Type']
        ) !==
        'competency'
      ) {

        return;
      }


      if (
        appNormalize_(
          record['Result']
        ) !==
        'pass'
      ) {

        return;
      }


      if (
        appIsTrue_(
          record['Automatic Failure']
        )
      ) {

        return;
      }


      const examination =
        dashText_(
          record['Examination']
        );


      const anatomicalArea =
        dashText_(
          record['Anatomical Area']
        );


      if (!examination) {

        return;
      }


      const key =
        appNormalize_(

          examination +
          '|' +
          anatomicalArea

        );


      const grade =
        gradeEngine_formatEvaluationGrade(
          record['Final Score']
        );


      const item = {

        submissionId:
          dashText_(
            record['Submission ID']
          ),

        examination:
          examination,

        anatomicalArea:
          anatomicalArea,

        clinicalTerm:
          dashText_(
            record['Clinical Term']
          ),

        attributedClinicalTerm:
          dashText_(
            record[
              'Attributed Clinical Term'
            ]
          ),

        date:
          dashDisplayDate_(
            record['Date']
          ),

        dateValue:
          appDateValue_(
            record['Date']
          ),

        evaluator:
          dashText_(
            record['Evaluator']
          ),

        evaluatorType:
          dashText_(
            record['Evaluator Type']
          ),

        clinicalSite:
          dashText_(
            record['Clinical Site']
          ),

        finalScore:
          dashPreserveZero_(
            record['Final Score']
          ),

        rawScoreText:
          grade.rawText,

        percentGradeText:
          grade.percentText,

        gradeDisplay:
          grade.combinedText,

        sourceDocument:
          dashText_(
            record['Source Document']
          )

      };


      if (
        !unique[key] ||
        item.dateValue >
        unique[key].dateValue
      ) {

        unique[key] =
          item;
      }

    });


  return Object
    .values(
      unique
    )
    .sort(
      (a, b) =>
        b.dateValue -
        a.dateValue
    );
}


/**
 * =====================================================================
 * SEMESTER REQUIREMENT SUMMARY
 * =====================================================================
 */
function dashBuildSemesterSummary_(
  statuses
) {

  const result = {

    termStatus:
      '',

    termDates:
      '',

    semesterRequirement:
      0,

    effectiveProgress:
      0,

    stillNeeded:
      0,

    carryoverApplied:
      0,

    newThisTerm:
      0,

    additionalFuture:
      0,

    additionalFutureMaximum:
      '',

    isoCompleted:
      0,

    isoRequired:
      0,

    failuresThisTerm:
      0,

    hciUsed:
      0,

    hciLimit:
      0,

    hciRemaining:
      0,

    hciStatus:
      ''

  };


  (statuses || [])

    .filter(row =>

      appNormalize_(
        row['Status Type']
      ) ===
        'summary' &&

      appNormalize_(
        row['Requirement Scope']
      ) ===
        'semester'

    )

    .forEach(row => {

      const id =
        appNormalize_(
          row['Requirement ID']
        );


      if (
        id.endsWith(
          '-status'
        )
      ) {

        result.termStatus =
          dashText_(
            row['Status']
          );


        result.termDates =
          dashText_(
            row['Notes']
          );


        return;
      }


      if (
        id.endsWith(
          '-comp'
        )
      ) {

        result.semesterRequirement =
          dashNumberOrText_(
            row['Required Count']
          );


        result.effectiveProgress =
          Number(
            row['Completed Count']
          ) || 0;


        result.stillNeeded =
          Number(
            row['Remaining Count']
          ) || 0;


        return;
      }


      if (
        id.endsWith(
          '-carry'
        )
      ) {

        result.carryoverApplied =
          Number(
            row['Completed Count']
          ) || 0;


        return;
      }


      if (
        id.endsWith(
          '-new'
        )
      ) {

        result.newThisTerm =
          Number(
            row['Completed Count']
          ) || 0;


        return;
      }


      if (
        id.endsWith(
          '-additional'
        )
      ) {

        result.additionalFuture =
          Number(
            row['Completed Count']
          ) || 0;


        result.additionalFutureMaximum =
          row['Required Count'] || '';


        return;
      }


      if (
        id.endsWith(
          '-iso'
        )
      ) {

        result.isoRequired =
          Number(
            row['Required Count']
          ) || 0;


        result.isoCompleted =
          Number(
            row['Completed Count']
          ) || 0;


        return;
      }


      if (
        id.endsWith(
          '-fail'
        )
      ) {

        result.failuresThisTerm =
          Number(
            row['Completed Count']
          ) || 0;


        return;
      }


      if (
        id.endsWith(
          '-hci'
        )
      ) {

        result.hciLimit =
          Number(
            row['Required Count']
          ) || 0;


        result.hciUsed =
          Number(
            row['Completed Count']
          ) || 0;


        result.hciRemaining =
          Number(
            row['Remaining Count']
          ) || 0;


        result.hciStatus =
          dashText_(
            row['Status']
          );
      }

    });


  return result;
}


/**
 * =====================================================================
 * PROGRAM REQUIREMENT SUMMARY
 * =====================================================================
 */
function dashBuildProgramSummary_(
  statuses
) {

  const result = {

    mandatory:
      dashEmptyProgress_(),

    elective:
      dashEmptyProgress_(),

    fluoroscopy:
      dashEmptyProgress_(),

    head:
      dashEmptyProgress_()

  };


  (statuses || [])

    .filter(row =>

      appNormalize_(
        row['Status Type']
      ) ===
        'summary' &&

      appNormalize_(
        row['Requirement Scope']
      ) ===
        'program'

    )

    .forEach(row => {

      const type =
        appNormalize_(
          row['Requirement Type']
        );


      const group =
        appNormalize_(
          row['Requirement Group']
        );


      const progress = {

        required:
          Number(
            row['Required Count']
          ) || 0,

        completed:
          Number(
            row['Completed Count']
          ) || 0,

        remaining:
          Number(
            row['Remaining Count']
          ) || 0,

        status:
          dashText_(
            row['Status']
          )

      };


      if (
        type ===
        'mandatory'
      ) {

        result.mandatory =
          progress;


        return;
      }


      if (
        type ===
          'elective' &&
        group ===
          'fluoroscopy'
      ) {

        result.fluoroscopy =
          progress;


        return;
      }


      if (
        type ===
          'elective' &&
        group ===
          'head'
      ) {

        result.head =
          progress;


        return;
      }


      if (
        type ===
          'elective' &&
        !group
      ) {

        result.elective =
          progress;
      }

    });


  return result;
}


/**
 * =====================================================================
 * REQUIREMENT LISTS
 * =====================================================================
 */
function dashBuildRequirementLists_(
  statuses
) {

  const result = {

    eligibleNow:
      [],

    prerequisite:
      [],

    future:
      []

  };


  (statuses || [])

    .filter(row =>

      appNormalize_(
        row['Status Type']
      ) ===
        'detail' &&

      appNormalize_(
        row['Requirement Scope']
      ) ===
        'program'

    )

    .forEach(row => {

      const item = {

        requirementId:
          dashText_(
            row['Requirement ID']
          ),

        name:
          dashText_(
            row['Requirement Name']
          ),

        type:
          dashText_(
            row['Requirement Type']
          ),

        group:
          dashText_(
            row['Requirement Group']
          ),

        timing:
          dashText_(
            row['Requirement Timing']
          ),

        status:
          dashText_(
            row['Status']
          ),

        requiredCount:
          dashNumberOrText_(
            row['Required Count']
          ),

        completedCount:
          Number(
            row['Completed Count']
          ) || 0,

        remainingCount:
          Number(
            row['Remaining Count']
          ) || 0,

        notes:
          dashText_(
            row['Notes']
          )

      };


      const timing =
        appNormalize_(
          item.timing
        );


      if (
        timing ===
        'eligible now'
      ) {

        result.eligibleNow.push(
          item
        );


      } else if (
        timing ===
        'future requirement'
      ) {

        result.future.push(
          item
        );


      } else {

        result.prerequisite.push(
          item
        );
      }

    });


  return result;
}


/**
 * =====================================================================
 * CLINICAL HISTORY
 * =====================================================================
 *
 * StudentClinicalHistoryService.gs owns the read-only history construction.
 * Dashboard.gs adapts its response to the existing DashboardUI contract.
 */
function dashBuildClinicalHistoryFromService_(history) {

  if (
    !history ||
    !Array.isArray(history.terms)
  ) {
    return [];
  }


  return history.terms.map(term => {

    const grade =
      term.semesterAverage || {};

    const regularActivities =
      (term.evaluations || [])
        .filter(item =>
          appNormalize_(item.academicTerm) ===
          appNormalize_(item.occurredTerm)
        )
        .map(
          dashClinicalHistoryTrackingRow_
        );


    const academicRecords =
      (term.supplementalGrades || [])
        .map(
          dashClinicalHistoryGradeRecordRow_
        );


    const carryoverIntoTerm =
      (term.carryoverIntoTerm || [])
        .map(
          dashClinicalHistoryTrackingRow_
        );


    const bankedForFuture =
      (term.bankedForFuture || [])
        .map(
          dashClinicalHistoryTrackingRow_
        );


    const metrics =
      dashClinicalHistoryMetrics_(
        term.counts || {}
      );


    return {
      clinicalTerm:
        dashText_(
          term.clinicalTerm
        ),

      timing:
        '',

      status:
        dashText_(
          term.status
        ),

      grade: {
        available:
          Boolean(
            grade.available
          ),

        value:
          grade.available
            ? (
                dashText_(grade.display) ||
                (
                  grade.value !== null &&
                  grade.value !== undefined
                    ? dashTrimNumber_(grade.value) + '%'
                    : '—'
                )
              )
            : '—',

        rawValue:
          grade.value,

        complete:
          Boolean(
            grade.complete
          ),

        partial:
          Boolean(
            grade.partial
          ),

        availableWeight:
          Number(
            grade.availableWeight
          ) || 0,

        note:
          dashText_(
            grade.note
          )
      },

      metrics: metrics,

      notes:
        dashText_(
          grade.note
        ),

      activities:
        regularActivities,

      academicRecords:
        academicRecords,

      carryoverIntoTerm:
        carryoverIntoTerm,

      bankedForFuture:
        bankedForFuture,

      hasActivity:
        Boolean(
          regularActivities.length ||
          academicRecords.length ||
          carryoverIntoTerm.length ||
          bankedForFuture.length
        )
    };
  });
}


function dashClinicalHistoryTrackingRow_(item) {

  return {
    submissionId:
      dashText_(
        item.submissionId
      ),

    date:
      dashText_(
        item.date
      ),

    dateValue:
      Number(
        item.dateValue
      ) || 0,

    examination:
      dashText_(
        item.examination
      ),

    title:
      dashText_(
        item.examination
      ),

    type:
      dashText_(
        item.evaluationType
      ),

    evaluationType:
      dashText_(
        item.evaluationType
      ),

    evaluator:
      dashText_(
        item.evaluator
      ),

    clinicalSite:
      dashText_(
        item.clinicalSite
      ),

    scoreDisplay:
      dashText_(
        item.gradeDisplay
      ) ||
      '—',

    result:
      dashText_(
        item.result
      ),

    clinicalTerm:
      dashText_(
        item.academicTerm
      ),

    countsTowardTerm:
      dashText_(
        item.academicTerm
      ),

    occurredInTerm:
      dashText_(
        item.occurredTerm
      ),

    attributedClinicalTerm:
      dashText_(
        item.occurredTerm
      ),

    automaticFailure:
      Boolean(
        item.automaticFailure
      )
  };
}


function dashClinicalHistoryGradeRecordRow_(item) {

  let scoreDisplay = '—';

  if (
    item.percentGrade !== null &&
    item.percentGrade !== undefined &&
    item.percentGrade !== ''
  ) {
    scoreDisplay =
      dashTrimNumber_(
        item.percentGrade
      ) + '%';

  } else if (
    item.points !== null &&
    item.points !== undefined &&
    item.pointsPossible !== null &&
    item.pointsPossible !== undefined
  ) {
    scoreDisplay =
      dashTrimNumber_(
        item.points
      ) +
      ' / ' +
      dashTrimNumber_(
        item.pointsPossible
      );
  }


  return {
    gradeRecordId:
      dashText_(
        item.gradeRecordId
      ),

    date:
      dashText_(
        item.date
      ),

    dateValue:
      Number(
        item.dateValue
      ) || 0,

    examination:
      dashText_(
        item.assignmentName
      ),

    title:
      dashText_(
        item.assignmentName
      ),

    type:
      dashText_(
        item.gradeCategory
      ),

    evaluator:
      dashText_(
        item.instructor
      ),

    clinicalSite:
      '',

    scoreDisplay:
      scoreDisplay,

    result:
      '',

    clinicalTerm:
      dashText_(
        item.clinicalTerm
      ),

    notes:
      dashText_(
        item.notes
      )
  };
}


function dashClinicalHistoryMetrics_(counts) {

  const metrics = [];


  dashPushClinicalHistoryMetric_(
    metrics,
    'Evaluations',
    counts.evaluations
  );

  dashPushClinicalHistoryMetric_(
    metrics,
    'Competencies',
    counts.competencies
  );

  dashPushClinicalHistoryMetric_(
    metrics,
    'Passed Competencies',
    counts.passedCompetencies
  );

  dashPushClinicalHistoryMetric_(
    metrics,
    'Failed Competencies',
    counts.failedCompetencies
  );

  dashPushClinicalHistoryMetric_(
    metrics,
    'ISO',
    counts.iso
  );

  dashPushClinicalHistoryMetric_(
    metrics,
    'SSO',
    counts.sso
  );

  dashPushClinicalHistoryMetric_(
    metrics,
    'Simulations',
    counts.simulation
  );

  dashPushClinicalHistoryMetric_(
    metrics,
    'Academic / Grade Items',
    counts.supplementalGradeItems
  );


  return metrics;
}


function dashPushClinicalHistoryMetric_(metrics, label, value) {

  const number =
    Number(
      value
    ) || 0;

  if (number <= 0) {
    return;
  }


  metrics.push({
    label: label,
    value: number
  });
}


/**
 * =====================================================================
 * REMEDIATION
 * =====================================================================
 */
function dashBuildRemediationAlerts_(
  statuses
) {

  return (statuses || [])

    .filter(row =>

      appNormalize_(
        row['Status Type']
      ) ===
        'alert' &&

      appNormalize_(
        row['Requirement Type']
      ) ===
        'remediation'

    )

    .map(row => ({

      name:
        dashText_(
          row['Requirement Name']
        ),

      status:
        dashText_(
          row['Status']
        ),

      failureCount:
        Number(
          row['Completed Count']
        ) || 0,

      notes:
        dashText_(
          row['Notes']
        ),

      lastDate:
        dashDisplayDate_(
          row['Last Completed Date']
        )

    }));
}


/**
 * =====================================================================
 * REVIEW STATS
 * =====================================================================
 */
function dashBuildReviewStats_(
  tracking
) {

  let awaitingReview =
    0;

  let flagged =
    0;

  let failed =
    0;

  let reviewedToday =
    0;

  const today =
    dashTodayKey_();

  (tracking || [])
    .forEach(record => {

      if (
        appNormalize_(
          record['Submission Method']
        ) ===
        'historical import'
      ) {

        return;
      }

      if (
        dashIsPendingReview_(
          record
        )
      ) {

        awaitingReview++;
      }

      if (
        dashIsFlagged_(
          record
        )
      ) {

        flagged++;
      }

      if (
        appNormalize_(
          record['Result']
        ) ===
        'fail'
      ) {

        failed++;
      }

      if (
        appNormalize_(
          record[
            'Coordinator Review Status'
          ]
        ) ===
          'reviewed' &&

        dashDateKey_(
          record['Review Date']
        ) ===
          today
      ) {

        reviewedToday++;
      }

    });

  return {

    awaitingReview:
      awaitingReview,

    flagged:
      flagged,

    failed:
      failed,

    reviewedToday:
      reviewedToday

  };
}


/**
 * =====================================================================
 * PENDING REVIEW
 * =====================================================================
 */
function dashIsPendingReview_(
  record
) {

  if (
    appNormalize_(
      record['Submission Method']
    ) ===
    'historical import'
  ) {

    return false;
  }

  const status =
    appNormalize_(
      record[
        'Coordinator Review Status'
      ]
    );

  /*
   * New live records should normally be "Pending".
   * Blank is also treated as pending so an older live record cannot
   * silently disappear from the coordinator's queue.
   */
  if (!status) {

    return true;
  }

  return [

    'pending',
    'pending review',
    'awaiting review',
    'needs review',
    'submitted'

  ].includes(
    status
  );
}


/**
 * =====================================================================
 * FLAGGED
 * =====================================================================
 */
function dashIsFlagged_(
  record
) {

  return (
    appNormalize_(
      record['Coordinator Flag']
    ) ===
    'yes'
  );
}


/**
 * =====================================================================
 * TODAY KEY
 * =====================================================================
 */
function dashTodayKey_() {

  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


/**
 * =====================================================================
 * ADD NEW CAPABILITIES
 * =====================================================================
 *
 * No write functionality is being activated here yet.
 *
 * This simply tells the future DashboardUI what the architecture is
 * intended to support.
 */
function dashGetAddNewCapabilities_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const gradeRecordsSheet =
    ss.getSheetByName(
      (
        APP_CONFIG.SHEETS &&
        APP_CONFIG.SHEETS.GRADE_RECORDS
      )

        ? APP_CONFIG.SHEETS
            .GRADE_RECORDS

        : 'Grade Records'
    );


  return {

    gradeRecordsReady:
      Boolean(
        gradeRecordsSheet
      ),

    categories: [

      {
        key:
          'professional',

        label:
          'Professional Competency',

        entryMode:
          'manual-now-form-later'
      },

      {
        key:
          'clinicalTest',

        label:
          'Clinical Test',

        entryMode:
          'manual'
      },

      {
        key:
          'quizzes',

        label:
          'Clinical Quiz',

        entryMode:
          'manual'
      },

      {
        key:
          'casePresentation',

        label:
          'Case Presentation',

        entryMode:
          'manual'
      },

      {
        key:
          'deduction',

        label:
          'Point Deduction',

        entryMode:
          'manual'
      }

    ]

  };
}


/**
 * =====================================================================
 * DOCUMENT CAPABILITIES
 * =====================================================================
 */
function dashGetDocumentCapabilities_() {

  const semesterTemplates =
    (
      APP_CONFIG.TEMPLATES &&
      APP_CONFIG.TEMPLATES
        .SEMESTER_REQUIREMENTS
    )

      ? APP_CONFIG.TEMPLATES
          .SEMESTER_REQUIREMENTS

      : {};


  const termMap = {

    CLINICAL_1:
      'Clinical 1',

    CLINICAL_2:
      'Clinical 2',

    CLINICAL_3:
      'Clinical 3',

    CLINICAL_4:
      'Clinical 4',

    CLINICAL_5:
      'Clinical 5'

  };


  const semesterTerms =
    Object
      .keys(
        termMap
      )

      .filter(key =>
        Boolean(
          semesterTemplates[key]
        )
      )

      .map(key =>
        termMap[key]
      );


  const programConfigured =
    Boolean(
      APP_CONFIG.TEMPLATES &&
      APP_CONFIG.TEMPLATES
        .PROGRAM_REQUIREMENTS
    );


  return {

    semester: {

      available:
        semesterTerms.length > 0,

      terms:
        semesterTerms

    },

    program: {

      available:
        programConfigured,

      templateConfigured:
        programConfigured

    }

  };
}


/**
 * =====================================================================
 * ACTIVE STUDENTS
 * =====================================================================
 */
function dashGetUniqueActiveStudents_(
  students
) {

  const unique = {};


  (students || [])

    .filter(student =>
      appIsTrue_(
        student['Active']
      )
    )

    .forEach(student => {

      const studentId =
        dashText_(
          student['Student ID']
        );


      const studentName =
        dashText_(
          student['Student']
        );


      const key =
        studentId ||
        appNormalize_(
          studentName
        );


      if (!key) {

        return;
      }


      if (
        !unique[key]
      ) {

        unique[key] =
          student;


        return;
      }


      if (
        dashClinicalNumber_(
          student['Clinical Term']
        ) >
        dashClinicalNumber_(
          unique[key][
            'Clinical Term'
          ]
        )
      ) {

        unique[key] =
          student;
      }

    });


  return Object.values(
    unique
  );
}


/**
 * =====================================================================
 * MOST ADVANCED STUDENT ROW
 * =====================================================================
 */
function dashGetMostAdvancedStudentRow_(
  rows
) {

  const active =
    (rows || [])
      .filter(row =>
        appIsTrue_(
          row['Active']
        )
      );


  const candidates =
    active.length
      ? active
      : rows;


  return candidates.reduce(
    (best, current) => {

      if (!best) {

        return current;
      }


      return (

        dashClinicalNumber_(
          current['Clinical Term']
        ) >

        dashClinicalNumber_(
          best['Clinical Term']
        )

      )

        ? current
        : best;

    },
    null
  );
}


/**
 * =====================================================================
 * STUDENT LIST
 * =====================================================================
 */
function dashStudentListItem_(
  student
) {

  return {

    studentId:
      dashText_(
        student['Student ID']
      ),

    student:
      dashText_(
        student['Student']
      ),

    clinicalTerm:
      dashText_(
        student['Clinical Term']
      ),

    cohort:
      dashText_(
        student['Cohort']
      )

  };
}


/**
 * =====================================================================
 * STUDENT DETAIL
 * =====================================================================
 */
function dashStudentDetail_(
  student
) {

  return dashStudentListItem_(
    student
  );
}


/**
 * =====================================================================
 * SAME STUDENT
 * =====================================================================
 */
function dashSameStudent_(
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
 * EMPTY GRADE SUMMARY
 * =====================================================================
 */
function dashEmptyGradeSummary_(
  clinicalTerm,
  errorMessage
) {

  return {

    semester: {

      clinicalTerm:
        clinicalTerm,

      complete:
        false,

      recordedAverage:
        null,

      recordedAverageText:
        '',

      availableWeight:
        0,

      includedCategories:
        [],

      missingCategories:
        [],

      note:
        errorMessage
          ? 'Grade calculation unavailable: ' +
            errorMessage
          : 'No grade information available.'

    },

    program: {

      available:
        false,

      complete:
        false,

      recordedAverage:
        null,

      recordedAverageText:
        '',

      note:
        ''

    },

    terms:
      []

  };
}


/**
 * =====================================================================
 * EMPTY REQUIREMENT PROGRESS
 * =====================================================================
 */
function dashEmptyProgress_() {

  return {

    required:
      0,

    completed:
      0,

    remaining:
      0,

    status:
      ''

  };
}


/**
 * =====================================================================
 * TEXT
 * =====================================================================
 */
function dashText_(
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
 * NUMBER OR TEXT
 * =====================================================================
 */
function dashNumberOrText_(
  value
) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {

    return 0;
  }


  const number =
    Number(
      value
    );


  return Number.isFinite(
    number
  )
    ? number
    : String(
        value
      );
}


/**
 * =====================================================================
 * CLINICAL NUMBER
 * =====================================================================
 */
function dashClinicalNumber_(
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
 * DATE KEY
 * =====================================================================
 */
function dashDateKey_(
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
    'yyyy-MM-dd'
  );
}


/**
 * =====================================================================
 * DISPLAY DATE
 * =====================================================================
 */
function dashDisplayDate_(
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
    'M/d/yyyy'
  );
}


/**
 * =====================================================================
 * PRESERVE ZERO
 * =====================================================================
 */
function dashPreserveZero_(
  value
) {

  if (
    value === 0
  ) {

    return 0;
  }


  if (
    value === null ||
    value === undefined ||
    String(value).trim() === ''
  ) {

    return '';
  }


  return value;
}

/**
 * =====================================================================
 * COMMAND CENTER / SCHEDULING PRESENTATION ENDPOINTS
 * =====================================================================
 * Dashboard remains a presentation/API layer. Scheduling rules and writes
 * are delegated to SchedulingService.gs.
 */
function dashboardGetRoomRotationWorkspace(cohort, clinicalTerm) {
  const rows = scheduling_getRoomRotationWorkspace(cohort, clinicalTerm) || [];
  return rows.map(dashboardRoomRotationWorkspaceItem_);
}

function dashboardValidateRoomRotationDraft(input) {
  const record = scheduling_validateRoomRotationDraft(input || {});
  return dashboardRoomRotationDraftRecord_(record);
}

function dashboardSaveRoomRotationDraft(input) {
  const result = scheduling_saveRoomRotationDraft(input || {});
  return dashboardPlainObject_(result);
}

function dashboardRoomRotationWorkspaceItem_(item) {
  return {
    assignmentId: dashText_(item.assignmentId),
    studentId: dashText_(item.studentId),
    student: dashText_(item.student),
    cohort: dashText_(item.cohort),
    clinicalTerm: dashText_(item.clinicalTerm),
    siteId: dashText_(item.siteId),
    site: dashText_(item.site),
    startDate: dashDisplayDate_(item.startDate),
    endDate: dashDisplayDate_(item.endDate),
    assignmentType: dashText_(item.assignmentType),
    officialRotation: dashboardRoomRotationSheetRecord_(item.officialRotation),
    draftRotation: dashboardRoomRotationSheetRecord_(item.draftRotation),
    rotationAreas: (item.rotationAreas || []).map(String)
  };
}

function dashboardRoomRotationSheetRecord_(row) {
  if (!row) return null;
  return {
    rotationId: dashText_(row['Rotation ID']),
    assignmentId: dashText_(row['Assignment ID']),
    studentId: dashText_(row['Student ID']),
    siteId: dashText_(row['Site ID']),
    startDate: dashDisplayDate_(row['Start Date']),
    endDate: dashDisplayDate_(row['End Date']),
    rotationArea: dashText_(row['Rotation Area']),
    priorityNote: dashText_(row['Priority Note']),
    notes: dashText_(row['Notes']),
    status: dashText_(row['Status'])
  };
}

function dashboardRoomRotationDraftRecord_(row) {
  if (!row) return null;
  return {
    rotationId: dashText_(row.rotationId),
    assignmentId: dashText_(row.assignmentId),
    studentId: dashText_(row.studentId),
    siteId: dashText_(row.siteId),
    startDate: dashDisplayDate_(row.startDate),
    endDate: dashDisplayDate_(row.endDate),
    rotationArea: dashText_(row.rotationArea),
    priorityNote: dashText_(row.priorityNote),
    notes: dashText_(row.notes),
    status: dashText_(row.status)
  };
}

function dashboardPlainObject_(value) {
  if (value === null || value === undefined) return value;
  if (value instanceof Date) return dashDisplayDate_(value);
  if (Array.isArray(value)) return value.map(dashboardPlainObject_);
  if (typeof value === 'object') {
    const output = {};
    Object.keys(value).forEach(key => {
      output[key] = dashboardPlainObject_(value[key]);
    });
    return output;
  }
  return value;
}



/**
 * =====================================================================
 * COMMAND CENTER ATTENDANCE PRESENTATION ADAPTERS
 * =====================================================================
 * AttendanceService remains authoritative. These endpoints are read-only.
 */
function dashboardGetStudentAttendanceSummary(studentId, clinicalTerm, asOfDate) {
  if (typeof attendance_getStudentTermSummary !== 'function') {
    throw new Error('AttendanceService is not available.');
  }
  return attendance_getStudentTermSummary(studentId, clinicalTerm, asOfDate || new Date());
}

function dashboardGetCohortAttendanceSummary(cohort, clinicalTerm, asOfDate) {

  if (typeof attendance_getCohortTermSummary !== 'function') {
    throw new Error('AttendanceService is not available.');
  }

  var result = attendance_getCohortTermSummary(
    cohort,
    clinicalTerm,
    asOfDate || new Date()
  );

  return dashboardMakeBrowserSafe_(result);
}
function dashboardMakeBrowserSafe_(value) {

  if (value instanceof Date) {
    return isNaN(value.getTime())
      ? ''
      : value.toISOString();
  }

  if (Array.isArray(value)) {
    return value.map(function(item) {
      return dashboardMakeBrowserSafe_(item);
    });
  }

  if (value && typeof value === 'object') {
    var safe = {};

    Object.keys(value).forEach(function(key) {
      safe[key] = dashboardMakeBrowserSafe_(value[key]);
    });

    return safe;
  }

  return value;
}
function dashboardGetAttendanceActionItems(asOfDate) {
  if (typeof attendance_getStudentTermSummary !== 'function') {
    throw new Error('AttendanceService is not available.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const students = dashGetUniqueActiveStudents_(
    appSheetToObjects_(appGetSheet_(ss, APP_CONFIG.SHEETS.STUDENTS))
  );
  const actionRequired = [];
  const waiting = [];

  students.forEach(student => {
    const item = dashStudentListItem_(student);
    if (!item.studentId || !item.clinicalTerm) return;

    try {
      const summary = attendance_getStudentTermSummary(
        item.studentId,
        item.clinicalTerm,
        asOfDate || new Date()
      );
      const makeUp = summary.makeUp || {};
      const evening = summary.eveningRotation || {};
      const attendance = summary.attendance || {};
      const counts = summary.eventCounts || {};

      if (Number(makeUp.outstandingHours || 0) > 0) {
        actionRequired.push({
          type: 'MAKE_UP', studentId: item.studentId, student: item.student,
          clinicalTerm: item.clinicalTerm,
          message: makeUp.outstandingHours + ' make-up hour(s) outstanding.'
        });
      }
      if (evening.replacementRequired) {
        actionRequired.push({
          type: 'EVENING_REPLACEMENT', studentId: item.studentId, student: item.student,
          clinicalTerm: item.clinicalTerm,
          message: 'Evening rotation replacement remains outstanding.'
        });
      }
      if (attendance.maximumAbsenceReviewRequired) {
        actionRequired.push({
          type: 'ATTENDANCE_REVIEW', studentId: item.studentId, student: item.student,
          clinicalTerm: item.clinicalTerm,
          message: 'Attendance threshold requires coordinator review.'
        });
      }
      if (Number(counts.pending || 0) > 0) {
        waiting.push({
          type: 'PENDING_ATTENDANCE', studentId: item.studentId, student: item.student,
          clinicalTerm: item.clinicalTerm,
          message: counts.pending + ' attendance event(s) waiting for review.'
        });
      }
    } catch (err) {
      // One incomplete student/term should not prevent Home from loading.
    }
  });

  return { success: true, actionRequired: actionRequired, waiting: waiting };
}


/** Read-only semester planning summary. Scheduling/Attendance services remain authoritative. */
function dashboardGetSemesterPlanningSummary(cohort, clinicalTerm, asOfDate) {
  if (!cohort || !clinicalTerm) throw new Error('Cohort and Clinical Term are required.');

 var ss = SpreadsheetApp.openById(
  '1YAPUuZ4Es4dzOk1AGsUsc-QC7ECcocBFl4Zy1-IOx80'
);

  function rows_(name) {
  var sh = ss.getSheetByName(name);

  if (!sh) {
    throw new Error(
      'Required CCMS sheet not found: ' + name +
      ' in workbook: ' + ss.getName()
    );
  }

  if (sh.getLastRow() < 2) {
    return [];
  }

  var a = sh.getDataRange().getValues();
  var h = a.shift().map(function(x) {
    return String(x || '').trim();
  });

  return a.map(function(r) {
    var o = {};
    h.forEach(function(k, i) {
      if (k) o[k] = r[i];
    });
    return o;
  });
}
  function v_(o,names){for(var i=0;i<names.length;i++)if(Object.prototype.hasOwnProperty.call(o,names[i]))return o[names[i]];return '';}
  function eq_(a,b){return String(a==null?'':a).trim().toLowerCase()===String(b==null?'':b).trim().toLowerCase();}

  var assignments=rows_('Student Site Assignments').filter(function(r){
    return eq_(v_(r,['Cohort']),cohort) &&
           eq_(v_(r,['Clinical Term','ClinicalTerm']),clinicalTerm) &&
           !eq_(v_(r,['Status']),'Cancelled');
  });

  var aids={}, sids={}, sites={};
  assignments.forEach(function(r){
    var a=v_(r,['Assignment ID','AssignmentId','AssignmentID']);
    var s=v_(r,['Student ID','StudentId','StudentID']);
    var t=v_(r,['Clinical Site ID','Site ID','SiteId','SiteID']);
    if(a)aids[String(a).trim()]=true;if(s)sids[String(s).trim()]=true;if(t)sites[String(t).trim()]=true;
  });

  var rotations=rows_('Student Room Rotations').filter(function(r){
    var a=String(v_(r,['Assignment ID','AssignmentId','AssignmentID'])||'').trim();
    return !!aids[a] && !eq_(v_(r,['Status']),'Cancelled');
  });

  var assigned={}, drafts=0;
  rotations.forEach(function(r){
    var area=String(v_(r,['Rotation Area','Room / Rotation Area','Room/Rotation Area','Room Rotation Area'])||'').trim();
    if(!area)return;
    var key=String(v_(r,['Rotation ID','RotationId','RotationID'])||v_(r,['Assignment ID','AssignmentId','AssignmentID'])||'').trim();
    if(key)assigned[key]=true;
    if(eq_(v_(r,['Status']),'Draft'))drafts++;
  });

  var expected=assignments.length, done=Object.keys(assigned).length, remaining=Math.max(0,expected-done);
  return {
    success:true, cohort:cohort, clinicalTerm:clinicalTerm,
    studentCount:Object.keys(sids).length, siteCount:Object.keys(sites).length,
    siteAssignmentCount:expected,
    sitePlacementStatus:expected?'Available':'Not Started / No Assignments',
    roomRotationAssignedCount:done, roomRotationRemainingCount:remaining,
    roomRotationDraftCount:drafts,
    roomRotationStatus:expected?(remaining===0?'Complete':'In Progress'):'Not Started',
    sitePreferencesStatus:'Not Yet Connected',
    ciAvailabilityStatus:'Not Yet Connected',
    ciCoverageStatus:'Not Yet Connected',
    eveningRotationStatus:/^Clinical\s+[45]$/i.test(String(clinicalTerm))?'Not Yet Connected':'Not Applicable'
  };
}
function testDashboardSemesterPlanningSummary() {
  var result = dashboardGetSemesterPlanningSummary(
    'Class of 2027',
    'Clinical 3'
  );

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}