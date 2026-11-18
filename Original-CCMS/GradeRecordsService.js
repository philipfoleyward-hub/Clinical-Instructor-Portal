/**
 * =====================================================================
 * CLINICAL COMPETENCY TRACKER
 * GRADE RECORDS SERVICE
 * 04_GradeRecordsService.gs
 * =====================================================================
 *
 * PURPOSE
 * -------
 * Authoritative write/read service for grade-bearing records that do not
 * belong in Clinical Tracking:
 *
 *   - Professional Competency
 *   - Clinical Test
 *   - Clinical Quiz / Clinical 1 Journal Composite
 *   - Case Presentation
 *   - Point Deduction
 *
 * This file deliberately does NOT calculate semester weights. The existing
 * 04_GradeEngine.gs remains authoritative for category averages, weighted
 * semester grades, partial-grade status, and program averages.
 *
 * ONE-TIME SETUP
 * --------------
 * 1. Add this file to the Apps Script project.
 * 2. Add this entry inside APP_CONFIG.SHEETS in 01_Config.gs:
 *
 *      GRADE_RECORDS: 'Grade Records',
 *
 * 3. Run setupGradeRecords() once from the Apps Script editor.
 *
 * The Grade Engine already falls back to the literal sheet name
 * "Grade Records", so the service remains compatible during setup.
 * =====================================================================
 */


const GRADE_RECORDS_CONFIG = {

  sheetName:
    'Grade Records',

  headers: [
    'Grade Record ID',
    'Student ID',
    'Student',
    'Clinical Term',
    'Grade Category',
    'Assignment / Evaluation Name',
    'Evaluation Period',
    'Date',
    'Instructor',
    'Raw Score',
    'Maximum Score',
    'Percent Grade',
    'Points',
    'Delivery Method',
    'Source',
    'Source ID',
    'Source Document',
    'Source File ID',
    'Entered By',
    'Entered At',
    'Notes',
    'Active',
    'Supersedes Record ID'
  ],

  categories: {

    professional: {
      label: 'Professional Competency',
      multipleAllowed: true
    },

    clinicalTest: {
      label: 'Clinical Test',
      multipleAllowed: false
    },

    quizzes: {
      label: 'Clinical Quiz',
      multipleAllowed: true
    },

    casePresentation: {
      label: 'Case Presentation',
      multipleAllowed: false
    },

    deduction: {
      label: 'Point Deduction',
      multipleAllowed: true
    }

  },

  clinicalTerms: [
    'Clinical 1',
    'Clinical 2',
    'Clinical 3',
    'Clinical 4',
    'Clinical 5'
  ]

};


/**
 * =====================================================================
 * ONE-TIME SHEET SETUP
 * =====================================================================
 *
 * Creates Grade Records if it does not exist. If it already exists, missing
 * columns are appended without deleting or rearranging existing data.
 */
function setupGradeRecords() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    gradeRecords_getOrCreateSheet_(ss);

  gradeRecords_ensureHeaders_(sheet);
  gradeRecords_applySheetConfiguration_(sheet);

  SpreadsheetApp.flush();

  const result = {
    success: true,
    sheetName: sheet.getName(),
    sheetId: sheet.getSheetId(),
    headers: GRADE_RECORDS_CONFIG.headers.slice(),
    message: 'Grade Records is ready.'
  };

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}


/**
 * =====================================================================
 * DASHBOARD-SAFE PUBLIC API
 * =====================================================================
 */


/**
 * Adds a manual grade record and returns the refreshed grade summary.
 *
 * Payload example:
 *
 * {
 *   studentId: '1',
 *   clinicalTerm: 'Clinical 3',
 *   gradeCategory: 'Clinical Test',
 *   assignmentName: 'Clinical 3 Comprehensive Test',
 *   date: '2026-10-20',
 *   instructor: 'Professor Terracino',
 *   percentGrade: 84.3,
 *   notes: ''
 * }
 */
function dashboardAddGradeRecord(payload) {

  const record =
    gradeRecords_add(payload);

  let gradeSummary = null;

  if (
    typeof gradeEngine_getStudentSummary ===
      'function'
  ) {

    gradeSummary =
      gradeEngine_getStudentSummary(
        record.studentId
      );
  }

  return {
    success: true,
    message: 'Grade record saved.',
    record: record,
    gradeSummary: gradeSummary
  };
}


/**
 * Voids a record without deleting it and returns the refreshed grade summary.
 */
function dashboardVoidGradeRecord(
  gradeRecordId,
  reason
) {

  const record =
    gradeRecords_void(
      gradeRecordId,
      reason
    );

  let gradeSummary = null;

  if (
    typeof gradeEngine_getStudentSummary ===
      'function'
  ) {

    gradeSummary =
      gradeEngine_getStudentSummary(
        record.studentId
      );
  }

  return {
    success: true,
    message: 'Grade record voided.',
    record: record,
    gradeSummary: gradeSummary
  };
}


/**
 * Returns active grade records for one student and optional clinical term.
 */
function dashboardGetGradeRecords(
  studentId,
  clinicalTerm
) {

  return {
    success: true,
    records:
      gradeRecords_list(
        studentId,
        clinicalTerm,
        false
      )
  };
}


/**
 * Supplies stable dropdown values to the New Entry UI.
 */
function dashboardGetGradeRecordOptions() {

  const instructors =
    gradeRecords_getInstructorOptions_();

  return {
    categories: [
      { key: 'clinicalTest', label: 'Clinical Test', multipleAllowed: false },
      { key: 'quizzes', label: 'Clinical Quiz', multipleAllowed: true },
      { key: 'professional', label: 'Professional Competency', multipleAllowed: true },
      { key: 'competencyEvaluation', label: 'Competency Evaluation', multipleAllowed: true, recordSystem: 'Clinical Tracking' },
      { key: 'casePresentation', label: 'Case Presentation', multipleAllowed: false },
      { key: 'deduction', label: 'Grade Point Reduction', multipleAllowed: true }
    ],

    clinicalTerms:
      GRADE_RECORDS_CONFIG
        .clinicalTerms
        .slice(),

    instructors: instructors,

    competencyEvaluation:
      gradeRecords_getCompetencyEvaluationOptions_(),

    evaluationPeriods: [
      'First Half',
      'Second Half',
      'Full Semester',
      'Other'
    ],

    deliveryMethods: [
      'In Person',
      'Paper',
      'Brightspace / LMS',
      'Dashboard Manual Entry',
      'Historical Import',
      'Other'
    ]
  };
}


function gradeRecords_getInstructorOptions_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Evaluators');
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const rows = gradeRecords_getSheetObjects_(sheet);
  const names = rows
    .filter(row =>
      (!Object.prototype.hasOwnProperty.call(row, 'Active') ||
        gradeRecords_isTrue_(row['Active'])) &&
      gradeRecords_clean_(
        gradeRecords_first_(row['Evaluator'], row['Instructor'], row['Name'])
      )
    )
    .map(row => gradeRecords_clean_(
      gradeRecords_first_(row['Evaluator'], row['Instructor'], row['Name'])
    ));

  return Array.from(new Set(names)).sort();
}


/**
 * Reads the established Digital Clinical Evaluation form and scoring table.
 * The dashboard therefore follows the same choices and criteria as the form.
 */
function gradeRecords_getCompetencyEvaluationOptions_() {

  if (!APP_CONFIG.FORMS || !APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION) {
    throw new Error('The Digital Clinical Evaluation form is not configured.');
  }

  const form = FormApp.openById(APP_CONFIG.FORMS.DIGITAL_CLINICAL_EVALUATION);
  const items = form.getItems();
  let evaluationTypes = [];

  items.forEach(item => {
    if (gradeRecords_normalize_(item.getTitle()) !== 'evaluation type') return;
    const type = item.getType();
    if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
      evaluationTypes = item.asMultipleChoiceItem().getChoices().map(choice => choice.getValue());
    } else if (type === FormApp.ItemType.LIST) {
      evaluationTypes = item.asListItem().getChoices().map(choice => choice.getValue());
    }
  });

  if (!evaluationTypes.length) {
    throw new Error('No Evaluation Type choices were found in the Digital Clinical Evaluation form.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weights = grading_loadPerformanceWeights_(ss);
  const sections = [];

  weights.forEach(row => {
    let section = sections.find(item =>
      gradeRecords_normalize_(item.name) === gradeRecords_normalize_(row.section)
    );
    if (!section) {
      section = { name: row.section, criteria: [] };
      sections.push(section);
    }
    section.criteria.push({
      key: row.criterionKey,
      criterion: row.criterion,
      weight: row.weight
    });
  });

  return {
    evaluationTypes: evaluationTypes,
    performanceSections: sections,
    ratingChoices: ['3', '2', '1', '0', 'N/A'],
    markerChoices: ['Acceptable', 'Missing Marker', 'Illegible Marker'],
    maximumImageViews: 8
  };
}


/**
 * Unified New Entry boundary. Attachments are filed first so their plain
 * Drive identifiers can be stored with the authoritative structured record.
 */
function dashboardSaveNewEntry(payload) {

  const input = payload || {};
  let attachment = null;

  if (input.attachment && input.attachment.data) {
    attachment = gradeRecords_saveSupportingAttachment_(input);
    input.sourceDocument = attachment.fileUrl;
    input.sourceFileId = attachment.fileId;
  }

  try {
    if (gradeRecords_normalize_(input.gradeCategory) === 'competencyevaluation') {
      return gradeRecords_addManualCompetencyEvaluation_(input, attachment);
    }

    const result = dashboardAddGradeRecord(input);
    result.attachment = attachment;
    return result;

  } catch (error) {
    if (attachment) {
      throw new Error(
        'The supporting document was saved as "' + attachment.fileName +
        '", but the structured record was not saved. ' + error.message
      );
    }
    throw error;
  }
}


function gradeRecords_saveSupportingAttachment_(input) {

  if (typeof dashboardDocuments_resolveStudent_ !== 'function' ||
      typeof dashboardDocuments_findExistingStudentFolder_ !== 'function') {
    throw new Error('The established non-creating student-folder resolver is unavailable.');
  }

  const resolved = dashboardDocuments_resolveStudent_(input.studentId);
  const studentFolder = dashboardDocuments_findExistingStudentFolder_(resolved);
  if (!studentFolder) {
    throw new Error('The established student Drive folder could not be resolved. No folder was created.');
  }

  const attachment = input.attachment || {};
  const mimeType = gradeRecords_clean_(attachment.mimeType);
  const allowed = [
    'application/pdf',
    'image/jpeg',
    'image/png',
    'image/gif',
    'image/webp',
    'application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  ];
  if (allowed.indexOf(mimeType) === -1) {
    throw new Error('Supporting documents must be PDF, JPG, PNG, GIF, WebP, DOC, or DOCX.');
  }

  const supportingFolder = gradeRecords_getOrCreateKnownSubfolder_(
    studentFolder,
    'Supporting Documents'
  );
  const bytes = Utilities.base64Decode(String(attachment.data).replace(/^data:[^;]+;base64,/, ''));
  const safeName = gradeRecords_supportingFileName_(input, attachment);
  const blob = Utilities.newBlob(bytes, mimeType, safeName);
  const file = supportingFolder.createFile(blob);

  return {
    fileId: file.getId(),
    fileName: file.getName(),
    fileUrl: file.getUrl(),
    folderId: supportingFolder.getId(),
    folderName: supportingFolder.getName()
  };
}


function gradeRecords_supportingFileName_(input, attachment) {

  const originalName =
    gradeRecords_clean_(attachment.fileName) ||
    'Supporting Document';

  const extensionMatch =
    originalName.match(/(\.[A-Za-z0-9]{1,8})$/);

  const extension =
    extensionMatch
      ? extensionMatch[1].toLowerCase()
      : gradeRecords_extensionForMime_(attachment.mimeType);

  const categoryKey =
    gradeRecords_normalizeCategory_(
      gradeRecords_first_(
        input.gradeCategory,
        input.category,
        input['Grade Category']
      )
    );

  const entryType =
    categoryKey && GRADE_RECORDS_CONFIG.categories[categoryKey]
      ? GRADE_RECORDS_CONFIG.categories[categoryKey].label
      : gradeRecords_clean_(input.gradeCategory) || 'Supporting Document';

  const date =
    gradeRecords_parseDate_(
      gradeRecords_first_(input.date, input.evaluationDate)
    );

  const dateText =
    date
      ? Utilities.formatDate(
          date,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        )
      : 'Undated';

  const descriptor =
    gradeRecords_clean_(
      gradeRecords_first_(
        input.assignmentName,
        input.assignment,
        input.instructor,
        input.evaluator,
        originalName.replace(/\.[^.]+$/, '')
      )
    );

  const parts = [
    gradeRecords_clean_(input.clinicalTerm) || 'Unassigned Term',
    entryType,
    dateText,
    descriptor
  ].filter(Boolean);

  return appSafeFileText_(parts.join(' - ')).slice(0, 180) + extension;
}


function gradeRecords_extensionForMime_(mimeType) {

  const extensions = {
    'application/pdf': '.pdf',
    'image/jpeg': '.jpg',
    'image/png': '.png',
    'image/gif': '.gif',
    'image/webp': '.webp',
    'application/msword': '.doc',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx'
  };

  return extensions[gradeRecords_clean_(mimeType)] || '';
}


function gradeRecords_getOrCreateKnownSubfolder_(parent, name) {
  const matches = parent.getFoldersByName(name);
  if (matches.hasNext()) {
    const folder = matches.next();
    if (matches.hasNext()) {
      throw new Error('Multiple "' + name + '" folders exist for this student.');
    }
    return folder;
  }
  return parent.createFolder(name);
}


function gradeRecords_addManualCompetencyEvaluation_(input, attachment) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const student = gradeRecords_resolveStudent_(ss, input.studentId, input.student);
  const clinicalTerm = gradeRecords_validateClinicalTerm_(input.clinicalTerm);
  const evaluationType = gradeRecords_clean_(input.evaluationType);
  const examination = gradeRecords_clean_(input.examination);
  const date = gradeRecords_parseDate_(input.date);
  const evaluator = gradeRecords_clean_(input.instructor);

  if (!evaluationType || !examination || !date || !evaluator) {
    throw new Error('Evaluation Type, examination, date, and evaluator are required.');
  }

  const allowedTypes = gradeRecords_getCompetencyEvaluationOptions_().evaluationTypes;
  if (!allowedTypes.some(type => gradeRecords_normalize_(type) === gradeRecords_normalize_(evaluationType))) {
    throw new Error('Evaluation Type is not an active choice in the Digital Clinical Evaluation form.');
  }

  let grade;
  if (gradeRecords_normalize_(evaluationType) === 'competency') {
    grade = gradeRecords_gradeManualCompetency_(input, ss, student, date);
  } else {
    const finalScore = gradeRecords_number_(input.finalScore);
    const result = gradeRecords_clean_(input.result).toUpperCase();
    if (finalScore === null || result === '') {
      throw new Error('Final Score and Result are required for this evaluation type.');
    }
    grade = {
      evaluation: {
        student: student.student,
        clinicalSite: gradeRecords_clean_(input.clinicalSite),
        evaluator: evaluator,
        examination: examination,
        evaluationType: evaluationType,
        date: date,
        comments: gradeRecords_clean_(input.notes)
      },
      automaticFailure: { occurred: result === 'FAIL', reasons: [], comments: '' },
      performance: { sections: [], totalPoints: 0, totalWeight: 0, score: gradeRecords_number_(input.performanceScore) },
      image: { views: [], grossPoints: 0, zeroPenalty: 0, missingMarkerPenalty: 0, illegibleMarkerPenalty: 0, netPoints: 0, sectionsScored: 0, score: gradeRecords_number_(input.imageScore) },
      finalScore: finalScore,
      result: result
    };
  }

  let documentResult = null;
  if (gradeRecords_normalize_(evaluationType) === 'competency') {
    documentResult = evaluationDocument_generate(grade);
  }

  gradeRecords_ensureClinicalTrackingAttachmentHeaders_(ss);
  const submissionId = 'MANUAL-EVAL-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16).toUpperCase();
  const record = {
    'Submission ID': submissionId,
    'Student ID': student.studentId,
    'Student': student.student,
    'Clinical Term': clinicalTerm,
    'Attributed Clinical Term': clinicalTerm,
    'Clinical Site': grade.evaluation.clinicalSite,
    'Evaluator': evaluator,
    'Evaluator Type': gradeRecords_clean_(input.evaluatorType),
    'Evaluation Type': evaluationType,
    'Examination': examination,
    'Evaluation Date': date,
    'Performance Score': grade.performance.score === null ? '' : gradeRecords_round2_(grade.performance.score),
    'Image Score': grade.image.score === null ? '' : gradeRecords_round2_(grade.image.score),
    'Final Score': grade.finalScore,
    'Result': grade.result,
    'Automatic Failure': grade.automaticFailure.occurred ? 'Yes' : 'No',
    'Source': 'Dashboard Manual Entry',
    'Source Document': documentResult ? (documentResult.docUrl || '') : (attachment ? attachment.fileUrl : ''),
    'PDF Document': documentResult ? (documentResult.pdfUrl || '') : '',
    'Supporting Document': attachment ? attachment.fileUrl : '',
    'Supporting File ID': attachment ? attachment.fileId : '',
    'Submission Method': 'Dashboard Manual Entry',
    'Entered By': gradeRecords_currentUser_(),
    'Coordinator Review Status': 'Approved',
    'Coordinator Notes': gradeRecords_clean_(input.notes)
  };

  const tracking = clinicalTracking_addRecord(record);
  let associationError = '';
  if (attachment) {
    try {
      gradeRecords_setClinicalTrackingAttachment_(ss, submissionId, attachment);
    } catch (error) {
      associationError = error.message;
    }
  }
  let requirements = null;
  try { requirements = rebuildStudentRequirementStatus(); } catch (error) {
    requirements = { success: false, error: error.message };
  }

  return {
    success: !associationError,
    partial: Boolean(associationError),
    message: associationError
      ? 'The structured evaluation and supporting file were saved, but their association failed: ' + associationError
      : 'Competency Evaluation saved.',
    record: {
      submissionId: submissionId,
      studentId: student.studentId,
      student: student.student,
      clinicalTerm: clinicalTerm,
      evaluationType: evaluationType,
      finalScore: grade.finalScore,
      result: grade.result
    },
    tracking: tracking,
    document: documentResult,
    attachment: attachment,
    requirements: requirements
  };
}


function gradeRecords_gradeManualCompetency_(input, ss, student, date) {

  const weights = grading_loadPerformanceWeights_(ss);
  const ratings = input.performanceRatings || {};
  const sections = [];
  let totalPoints = 0;
  let totalWeight = 0;

  weights.forEach(rule => {
    let section = sections.find(item => gradeRecords_normalize_(item.name) === gradeRecords_normalize_(rule.section));
    if (!section) {
      section = { name: rule.section, points: 0, weight: 0, score: null, rows: [] };
      sections.push(section);
    }
    const raw = gradeRecords_clean_(ratings[rule.criterionKey]);
    const applicable = raw && gradeRecords_normalize_(raw) !== 'n/a';
    if (!raw) throw new Error('Every Competency performance criterion must be scored or marked N/A.');
    const rating = applicable ? Number(raw) : null;
    if (applicable && (!Number.isFinite(rating) || rating < 0 || rating > 3)) {
      throw new Error('Invalid performance rating for ' + rule.criterion + '.');
    }
    const points = applicable ? rating * rule.weight : 0;
    if (applicable) {
      section.points += points;
      section.weight += rule.weight;
      totalPoints += points;
      totalWeight += rule.weight;
    }
    section.rows.push({ criterion: rule.criterion, ratingText: raw, applicable: applicable, weight: rule.weight, points: points });
  });
  sections.forEach(section => { section.score = section.weight ? section.points / section.weight : null; });
  const performanceScore = totalWeight ? totalPoints / totalWeight : null;

  const penalties = grading_loadImagePenalties_(ss);
  const inputViews = input.imageViews || [];
  if (!inputViews.length) throw new Error('At least one image evaluation view is required.');
  let grossPoints = 0, zeroPenalty = 0, missingMarkerPenalty = 0, illegibleMarkerPenalty = 0, netPoints = 0, sectionsScored = 0;
  const views = inputViews.map((view, index) => {
    const components = [];
    let viewGross = 0, viewSections = 0, viewZero = 0;
    const componentNames = {
      positioning: 'Positioning',
      equipment: 'Equipment / Central Ray',
      radiationProtection: 'Radiation Protection / Artifacts',
      exposure: 'Exposure'
    };
    Object.keys(componentNames).forEach(key => {
      const raw = gradeRecords_clean_(view[key]);
      if (!raw) throw new Error('Every image component must be scored or marked N/A.');
      const applicable = gradeRecords_normalize_(raw) !== 'n/a';
      const score = applicable ? Number(raw) : null;
      if (applicable && (!Number.isFinite(score) || score < 0 || score > 3)) throw new Error('Invalid image score in View ' + (index + 1) + '.');
      if (applicable) { viewGross += score; viewSections++; if (score === 0) viewZero = penalties.zeroPenalty; }
      components.push({ name: componentNames[key], applicable: applicable, score: score, explanation: applicable ? 'Score ' + score : 'N/A' });
    });
    const marker = gradeRecords_normalize_(view.marker);
    const missing = marker === 'missing marker' ? penalties.missingMarkerPenalty : 0;
    const illegible = marker === 'illegible marker' ? penalties.illegibleMarkerPenalty : 0;
    const viewNet = viewGross + viewZero + missing + illegible;
    grossPoints += viewGross; zeroPenalty += viewZero; missingMarkerPenalty += missing; illegibleMarkerPenalty += illegible; netPoints += viewNet; sectionsScored += viewSections;
    return { viewNumber: index + 1, projection: gradeRecords_clean_(view.projection), components: components, exposureIndex: gradeRecords_clean_(view.exposureIndex), marker: gradeRecords_clean_(view.marker), comments: gradeRecords_clean_(view.comments), grossPoints: viewGross, zeroPenalty: viewZero, missingMarkerPenalty: missing, illegibleMarkerPenalty: illegible, netPoints: viewNet, sectionsScored: viewSections, imageScore: viewSections ? viewNet / viewSections : null };
  });
  const imageScore = sectionsScored ? netPoints / sectionsScored : null;
  const automaticFailure = Boolean(input.automaticFailure);
  const finalScore = performanceScore !== null && imageScore !== null ? gradeRecords_round2_((performanceScore + imageScore) / 2) : null;
  const result = automaticFailure || performanceScore < 2 || imageScore < 2 ? 'FAIL' : 'PASS';

  return {
    evaluation: { student: student.student, clinicalSite: gradeRecords_clean_(input.clinicalSite), evaluator: gradeRecords_clean_(input.instructor), examination: gradeRecords_clean_(input.examination), evaluationType: input.evaluationType, date: date, comments: gradeRecords_clean_(input.notes) },
    automaticFailure: { occurred: automaticFailure, reasons: input.automaticFailureReasons || [], comments: gradeRecords_clean_(input.automaticFailureComments) },
    performance: { sections: sections, totalPoints: totalPoints, totalWeight: totalWeight, score: performanceScore },
    image: { views: views, grossPoints: grossPoints, zeroPenalty: zeroPenalty, missingMarkerPenalty: missingMarkerPenalty, illegibleMarkerPenalty: illegibleMarkerPenalty, netPoints: netPoints, sectionsScored: sectionsScored, score: imageScore },
    finalScore: finalScore,
    result: result
  };
}


function gradeRecords_ensureClinicalTrackingAttachmentHeaders_(ss) {
  const sheet = ss.getSheetByName('Clinical Tracking');
  if (!sheet) throw new Error('Could not find Clinical Tracking tab.');
  const current = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(gradeRecords_clean_);
  ['Supporting Document', 'Supporting File ID'].forEach(header => {
    if (current.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      current.push(header);
    }
  });
}


function gradeRecords_setClinicalTrackingAttachment_(ss, submissionId, attachment) {
  const sheet = ss.getSheetByName('Clinical Tracking');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(gradeRecords_clean_);
  const idIndex = headers.indexOf('Submission ID');
  const urlIndex = headers.indexOf('Supporting Document');
  const fileIndex = headers.indexOf('Supporting File ID');
  if (idIndex === -1 || urlIndex === -1 || fileIndex === -1) {
    throw new Error('Clinical Tracking supporting-document columns are unavailable.');
  }
  const rowCount = sheet.getLastRow() - 1;
  if (rowCount < 1) throw new Error('Clinical Tracking contains no evaluation rows.');
  const values = sheet.getRange(2, idIndex + 1, rowCount, 1).getValues();
  for (let index = 0; index < values.length; index++) {
    if (gradeRecords_clean_(values[index][0]) === submissionId) {
      sheet.getRange(index + 2, urlIndex + 1).setValue(attachment.fileUrl);
      sheet.getRange(index + 2, fileIndex + 1).setValue(attachment.fileId);
      return;
    }
  }
  throw new Error('The evaluation was saved, but its supporting-document association could not be located.');
}


/**
 * =====================================================================
 * AUTHORITATIVE CREATE
 * =====================================================================
 */
function gradeRecords_add(payload) {

  const input =
    payload || {};

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    gradeRecords_getOrCreateSheet_(ss);

  gradeRecords_ensureHeaders_(sheet);

  const student =
    gradeRecords_resolveStudent_(
      ss,
      gradeRecords_first_(
        input.studentId,
        input['Student ID']
      ),
      gradeRecords_first_(
        input.student,
        input['Student']
      )
    );

  const clinicalTerm =
    gradeRecords_validateClinicalTerm_(
      gradeRecords_first_(
        input.clinicalTerm,
        input['Clinical Term']
      )
    );

  const categoryKey =
    gradeRecords_normalizeCategory_(
      gradeRecords_first_(
        input.gradeCategory,
        input.category,
        input['Grade Category']
      )
    );

  if (!categoryKey) {
    throw new Error(
      'A recognized Grade Category is required.'
    );
  }

  if (
    categoryKey === 'casePresentation' &&
    clinicalTerm !== 'Clinical 2' &&
    clinicalTerm !== 'Clinical 5'
  ) {
    throw new Error(
      'Case Presentation is grade-bearing only in Clinical 2 and Clinical 5.'
    );
  }

  const categoryLabel =
    GRADE_RECORDS_CONFIG
      .categories[categoryKey]
      .label;

  const evaluationDate =
    gradeRecords_parseDate_(
      gradeRecords_first_(
        input.date,
        input.evaluationDate,
        input['Date']
      )
    );

  if (!evaluationDate) {
    throw new Error(
      'A valid evaluation date is required.'
    );
  }

  const values =
    gradeRecords_calculateValues_(
      categoryKey,
      input
    );

  const assignmentName =
    gradeRecords_clean_(
      gradeRecords_first_(
        input.assignmentName,
        input.assignment,
        input['Assignment / Evaluation Name'],
        gradeRecords_defaultAssignmentName_(
          categoryKey,
          clinicalTerm
        )
      )
    );

  const lock =
    LockService.getDocumentLock();

  lock.waitLock(30000);

  try {

    const existing =
      gradeRecords_getSheetObjects_(sheet);

    const category =
      GRADE_RECORDS_CONFIG
        .categories[categoryKey];

    const replaceExisting =
      input.replaceExisting === undefined
        ? !category.multipleAllowed
        : gradeRecords_isTrue_(
            input.replaceExisting
          );

    let supersedesRecordId =
      gradeRecords_clean_(
        gradeRecords_first_(
          input.supersedesRecordId,
          input['Supersedes Record ID']
        )
      );

    if (replaceExisting) {

      const superseded =
        gradeRecords_deactivateMatching_(
          sheet,
          existing,
          student.studentId,
          clinicalTerm,
          categoryKey,
          'Superseded by a newer official record.'
        );

      if (
        !supersedesRecordId &&
        superseded.length > 0
      ) {

        supersedesRecordId =
          superseded[0];
      }
    }

    const now =
      new Date();

    const recordId =
      'GR-' +
      Utilities
        .getUuid()
        .replace(/-/g, '')
        .slice(0, 16)
        .toUpperCase();

    const record = {
      'Grade Record ID': recordId,
      'Student ID': student.studentId,
      'Student': student.student,
      'Clinical Term': clinicalTerm,
      'Grade Category': categoryLabel,
      'Assignment / Evaluation Name': assignmentName,
      'Evaluation Period':
        gradeRecords_clean_(
          gradeRecords_first_(
            input.evaluationPeriod,
            input['Evaluation Period']
          )
        ),
      'Date': evaluationDate,
      'Instructor':
        gradeRecords_clean_(
          gradeRecords_first_(
            input.instructor,
            input.evaluator,
            input['Instructor']
          )
        ),
      'Raw Score': values.rawScore,
      'Maximum Score': values.maximumScore,
      'Percent Grade': values.percentGrade,
      'Points': values.points,
      'Delivery Method':
        gradeRecords_clean_(
          gradeRecords_first_(
            input.deliveryMethod,
            input['Delivery Method']
          )
        ),
      'Source':
        gradeRecords_clean_(
          gradeRecords_first_(
            input.source,
            input['Source'],
            'Dashboard Manual Entry'
          )
        ),
      'Source ID':
        gradeRecords_clean_(
          gradeRecords_first_(
            input.sourceId,
            input['Source ID']
          )
        ),
      'Source Document':
        gradeRecords_clean_(
          gradeRecords_first_(input.sourceDocument, input['Source Document'])
        ),
      'Source File ID':
        gradeRecords_clean_(
          gradeRecords_first_(input.sourceFileId, input['Source File ID'])
        ),
      'Entered By':
        gradeRecords_currentUser_(),
      'Entered At': now,
      'Notes':
        gradeRecords_clean_(
          gradeRecords_first_(
            input.notes,
            input['Notes']
          )
        ),
      'Active': true,
      'Supersedes Record ID':
        supersedesRecordId
    };

    appAppendByHeaders_(
      sheet,
      record
    );

    SpreadsheetApp.flush();

    return gradeRecords_publicRecord_(
      record
    );

  } finally {

    lock.releaseLock();
  }
}


/**
 * =====================================================================
 * VOID / CORRECTION
 * =====================================================================
 *
 * Records are never deleted. A correction should normally be entered as a
 * new record that supersedes the old one. This function handles a true void.
 */
function gradeRecords_void(
  gradeRecordId,
  reason
) {

  const id =
    gradeRecords_clean_(gradeRecordId);

  if (!id) {
    throw new Error(
      'Grade Record ID is required.'
    );
  }

  const voidReason =
    gradeRecords_clean_(reason);

  if (!voidReason) {
    throw new Error(
      'A void reason is required.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    gradeRecords_getSheet_(ss);

  const lock =
    LockService.getDocumentLock();

  lock.waitLock(30000);

  try {

    const found =
      gradeRecords_findRowById_(
        sheet,
        id
      );

    if (!found) {
      throw new Error(
        'Grade Record not found: ' + id
      );
    }

    const headers =
      gradeRecords_headerMap_(sheet);

    gradeRecords_setCell_(
      sheet,
      found.rowNumber,
      headers,
      'Active',
      false
    );

    const oldNotes =
      gradeRecords_clean_(
        found.record['Notes']
      );

    const auditNote =
      '[VOID ' +
      gradeRecords_formatDateTime_(
        new Date()
      ) +
      ' by ' +
      gradeRecords_currentUser_() +
      '] ' +
      voidReason;

    gradeRecords_setCell_(
      sheet,
      found.rowNumber,
      headers,
      'Notes',
      oldNotes
        ? oldNotes + '\n' + auditNote
        : auditNote
    );

    SpreadsheetApp.flush();

    found.record['Active'] = false;
    found.record['Notes'] =
      oldNotes
        ? oldNotes + '\n' + auditNote
        : auditNote;

    return gradeRecords_publicRecord_(
      found.record
    );

  } finally {

    lock.releaseLock();
  }
}


/**
 * =====================================================================
 * LIST / QUERY
 * =====================================================================
 */
function gradeRecords_list(
  studentId,
  clinicalTerm,
  includeInactive
) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      gradeRecords_sheetName_()
    );

  if (!sheet) {
    return [];
  }

  const targetId =
    gradeRecords_normalize_(studentId);

  const targetTerm =
    gradeRecords_normalize_(clinicalTerm);

  return gradeRecords_getSheetObjects_(sheet)
    .filter(record => {

      if (
        targetId &&
        gradeRecords_normalize_(
          record['Student ID']
        ) !== targetId
      ) {
        return false;
      }

      if (
        targetTerm &&
        gradeRecords_normalize_(
          record['Clinical Term']
        ) !== targetTerm
      ) {
        return false;
      }

      if (
        !includeInactive &&
        !gradeRecords_isTrue_(
          record['Active']
        )
      ) {
        return false;
      }

      return true;
    })
    .sort((a, b) =>
      gradeRecords_dateValue_(
        b['Date']
      ) -
      gradeRecords_dateValue_(
        a['Date']
      )
    )
    .map(record =>
      gradeRecords_publicRecord_(
        record
      )
    );
}


/**
 * =====================================================================
 * VALUE CALCULATION
 * =====================================================================
 */
function gradeRecords_calculateValues_(
  categoryKey,
  input
) {

  let rawScore =
    gradeRecords_number_(
      gradeRecords_first_(
        input.rawScore,
        input['Raw Score']
      )
    );

  let maximumScore =
    gradeRecords_number_(
      gradeRecords_first_(
        input.maximumScore,
        input.maxScore,
        input['Maximum Score']
      )
    );

  let percentGrade =
    gradeRecords_number_(
      gradeRecords_first_(
        input.percentGrade,
        input.percent,
        input['Percent Grade']
      )
    );

  let points =
    gradeRecords_number_(
      gradeRecords_first_(
        input.points,
        input['Points']
      )
    );

  if (categoryKey === 'deduction') {

    if (points === null) {
      points = percentGrade;
    }

    if (points === null) {
      points = rawScore;
    }

    if (
      points === null ||
      points === 0
    ) {
      throw new Error(
        'Point Deduction requires a non-zero point value.'
      );
    }

    points =
      points > 0
        ? -points
        : points;

    return {
      rawScore: '',
      maximumScore: '',
      percentGrade: '',
      points: gradeRecords_round2_(points)
    };
  }

  if (
    percentGrade !== null &&
    percentGrade >= 0 &&
    percentGrade <= 1
  ) {
    percentGrade *= 100;
  }

  if (
    percentGrade === null &&
    rawScore !== null &&
    maximumScore !== null
  ) {

    if (maximumScore <= 0) {
      throw new Error(
        'Maximum Score must be greater than zero.'
      );
    }

    percentGrade =
      rawScore /
      maximumScore *
      100;
  }

  if (
    percentGrade === null &&
    categoryKey === 'professional' &&
    rawScore !== null
  ) {

    percentGrade =
      gradeRecords_professionalScoreToPercent_(
        rawScore
      );
  }

  if (percentGrade === null) {
    throw new Error(
      'Enter a Percent Grade or both Raw Score and Maximum Score.'
    );
  }

  if (
    percentGrade < 0 ||
    percentGrade > 100
  ) {
    throw new Error(
      'Percent Grade must be between 0 and 100.'
    );
  }

  return {
    rawScore:
      rawScore === null
        ? ''
        : rawScore,
    maximumScore:
      maximumScore === null
        ? ''
        : maximumScore,
    percentGrade:
      gradeRecords_round2_(
        percentGrade
      ),
    points: ''
  };
}


/**
 * =====================================================================
 * PROFESSIONAL COMPETENCY 0-3 -> PERCENT
 * =====================================================================
 *
 * Reproduces Sheet 2 of the uploaded Professional Competency workbooks.
 * The source table is intentionally nonlinear:
 *
 *   3.00 -> 100%
 *   2.95 -> 95.75%
 *   2.20 -> 77%
 *   2.10 -> 75%
 *   2.00 -> 75%
 *   1.50 -> 65%
 *   1.00 -> 55%
 *   0.50 -> 50%
 *   0.00 -> 45%
 *
 * Scores are rounded to hundredths before conversion, matching the workbook.
 */
function gradeRecords_professionalScoreToPercent_(
  value
) {

  const numeric =
    gradeRecords_number_(value);

  if (
    numeric === null ||
    numeric < 0 ||
    numeric > 3
  ) {
    throw new Error(
      'Professional Competency Raw Score must be between 0 and 3.'
    );
  }

  const score =
    Math.round(
      numeric * 100
    ) / 100;

  let percent;

  if (score >= 2.96) {
    percent =
      (score - 2) * 100;

  } else if (score >= 2.20) {
    percent =
      22 +
      score * 25;

  } else if (score === 2.19) {
    percent = 76.75;

  } else if (score >= 2.11) {
    percent = 76;

  } else if (score >= 1.99) {
    percent = 75;

  } else if (score >= 1.00) {
    percent =
      35 +
      score * 20;

  } else {
    percent =
      45 +
      score * 10;
  }

  return gradeRecords_round2_(
    percent
  );
}


/**
 * =====================================================================
 * SUPERSESSION
 * =====================================================================
 */
function gradeRecords_deactivateMatching_(
  sheet,
  records,
  studentId,
  clinicalTerm,
  categoryKey,
  reason
) {

  const headers =
    gradeRecords_headerMap_(sheet);

  const matches =
    (records || [])
      .filter(record =>
        gradeRecords_isTrue_(
          record['Active']
        ) &&
        gradeRecords_normalize_(
          record['Student ID']
        ) ===
        gradeRecords_normalize_(
          studentId
        ) &&
        gradeRecords_normalize_(
          record['Clinical Term']
        ) ===
        gradeRecords_normalize_(
          clinicalTerm
        ) &&
        gradeRecords_normalizeCategory_(
          record['Grade Category']
        ) ===
        categoryKey
      )
      .sort((a, b) =>
        gradeRecords_dateValue_(
          b['Entered At'] || b['Date']
        ) -
        gradeRecords_dateValue_(
          a['Entered At'] || a['Date']
        )
      );

  const ids = [];

  matches.forEach(record => {

    const rowNumber =
      record.__rowNumber;

    gradeRecords_setCell_(
      sheet,
      rowNumber,
      headers,
      'Active',
      false
    );

    const oldNotes =
      gradeRecords_clean_(
        record['Notes']
      );

    const auditNote =
      '[SUPERSEDED ' +
      gradeRecords_formatDateTime_(
        new Date()
      ) +
      '] ' +
      reason;

    gradeRecords_setCell_(
      sheet,
      rowNumber,
      headers,
      'Notes',
      oldNotes
        ? oldNotes + '\n' + auditNote
        : auditNote
    );

    ids.push(
      gradeRecords_clean_(
        record['Grade Record ID']
      )
    );
  });

  return ids.filter(Boolean);
}


/**
 * =====================================================================
 * STUDENT / TERM / CATEGORY VALIDATION
 * =====================================================================
 */
function gradeRecords_resolveStudent_(
  ss,
  studentId,
  studentName
) {

  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENTS
    );

  const students =
    appSheetToObjects_(sheet);

  const targetId =
    gradeRecords_normalize_(studentId);

  const targetName =
    gradeRecords_normalize_(studentName);

  const matches =
    students.filter(student => {

      if (targetId) {
        return (
          gradeRecords_normalize_(
            student['Student ID']
          ) === targetId
        );
      }

      return (
        targetName &&
        gradeRecords_normalize_(
          student['Student']
        ) === targetName
      );
    });

  if (matches.length === 0) {
    throw new Error(
      'Student was not found in the Students sheet.'
    );
  }

  const active =
    matches.filter(student =>
      !Object.prototype
        .hasOwnProperty
        .call(student, 'Active') ||
      gradeRecords_isTrue_(
        student['Active']
      )
    );

  const selected =
    (active.length ? active : matches)[0];

  return {
    studentId:
      gradeRecords_clean_(
        selected['Student ID']
      ),
    student:
      gradeRecords_clean_(
        selected['Student']
      )
  };
}


function gradeRecords_validateClinicalTerm_(
  value
) {

  const normalized =
    gradeRecords_normalize_(value);

  const term =
    GRADE_RECORDS_CONFIG
      .clinicalTerms
      .find(item =>
        gradeRecords_normalize_(item) ===
        normalized
      );

  if (!term) {
    throw new Error(
      'Clinical Term must be Clinical 1, 2, 3, 4, or 5.'
    );
  }

  return term;
}


function gradeRecords_normalizeCategory_(
  value
) {

  const normalized =
    gradeRecords_normalize_(value);

  if (
    normalized === 'professional' ||
    normalized.includes('professional competency')
  ) {
    return 'professional';
  }

  if (
    normalized === 'clinicaltest' ||
    normalized === 'test' ||
    normalized.includes('clinical test')
  ) {
    return 'clinicalTest';
  }

  if (
    normalized === 'quizzes' ||
    normalized === 'quiz' ||
    normalized.includes('clinical quiz') ||
    normalized.includes('journal')
  ) {
    return 'quizzes';
  }

  if (
    normalized === 'casepresentation' ||
    normalized.includes('case presentation')
  ) {
    return 'casePresentation';
  }

  if (
    normalized === 'deduction' ||
    normalized.includes('point deduction') ||
    normalized.includes('points off')
  ) {
    return 'deduction';
  }

  return '';
}


function gradeRecords_defaultAssignmentName_(
  categoryKey,
  clinicalTerm
) {

  const label =
    GRADE_RECORDS_CONFIG
      .categories[categoryKey]
      .label;

  return (
    clinicalTerm +
    ' ' +
    label
  );
}


/**
 * =====================================================================
 * SHEET SETUP / ACCESS
 * =====================================================================
 */
function gradeRecords_sheetName_() {

  if (
    typeof APP_CONFIG !== 'undefined' &&
    APP_CONFIG.SHEETS &&
    APP_CONFIG.SHEETS.GRADE_RECORDS
  ) {
    return APP_CONFIG.SHEETS.GRADE_RECORDS;
  }

  return GRADE_RECORDS_CONFIG.sheetName;
}


function gradeRecords_getOrCreateSheet_(ss) {

  const sheetName =
    gradeRecords_sheetName_();

  return (
    ss.getSheetByName(sheetName) ||
    ss.insertSheet(sheetName)
  );
}


function gradeRecords_getSheet_(ss) {

  const sheet =
    ss.getSheetByName(
      gradeRecords_sheetName_()
    );

  if (!sheet) {
    throw new Error(
      'Grade Records has not been created. Run setupGradeRecords() first.'
    );
  }

  return sheet;
}


function gradeRecords_ensureHeaders_(sheet) {

  const required =
    GRADE_RECORDS_CONFIG.headers;

  if (sheet.getLastColumn() < 1) {
    sheet
      .getRange(
        1,
        1,
        1,
        required.length
      )
      .setValues([required]);

    return;
  }

  const existing =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(value =>
        gradeRecords_clean_(value)
      );

  const normalized =
    existing.map(value =>
      gradeRecords_normalize_(value)
    );

  const missing =
    required.filter(header =>
      !normalized.includes(
        gradeRecords_normalize_(header)
      )
    );

  if (missing.length > 0) {

    sheet
      .getRange(
        1,
        existing.length + 1,
        1,
        missing.length
      )
      .setValues([missing]);
  }
}


function gradeRecords_applySheetConfiguration_(
  sheet
) {

  const headers =
    gradeRecords_headerMap_(sheet);

  const lastColumn =
    sheet.getLastColumn();

  sheet.setFrozenRows(1);

  sheet
    .getRange(1, 1, 1, lastColumn)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setWrap(true);

  if (!sheet.getFilter()) {
    sheet
      .getRange(
        1,
        1,
        Math.max(sheet.getLastRow(), 2),
        lastColumn
      )
      .createFilter();
  }

  const validationRows =
    Math.max(
      sheet.getMaxRows() - 1,
      1
    );

  gradeRecords_applyListValidation_(
    sheet,
    headers,
    'Clinical Term',
    GRADE_RECORDS_CONFIG.clinicalTerms,
    validationRows
  );

  gradeRecords_applyListValidation_(
    sheet,
    headers,
    'Grade Category',
    Object.keys(
      GRADE_RECORDS_CONFIG.categories
    ).map(key =>
      GRADE_RECORDS_CONFIG
        .categories[key]
        .label
    ),
    validationRows
  );

  gradeRecords_applyListValidation_(
    sheet,
    headers,
    'Evaluation Period',
    [
      'First Half',
      'Second Half',
      'Full Semester',
      'Other'
    ],
    validationRows
  );

  gradeRecords_applyListValidation_(
    sheet,
    headers,
    'Delivery Method',
    [
      'In Person',
      'Paper',
      'Brightspace / LMS',
      'Dashboard Manual Entry',
      'Historical Import',
      'Other'
    ],
    validationRows
  );

  gradeRecords_applyCheckboxValidation_(
    sheet,
    headers,
    'Active',
    validationRows
  );

  gradeRecords_setColumnFormat_(
    sheet,
    headers,
    'Date',
    'm/d/yyyy',
    validationRows
  );

  gradeRecords_setColumnFormat_(
    sheet,
    headers,
    'Entered At',
    'm/d/yyyy h:mm am/pm',
    validationRows
  );

  [
    'Raw Score',
    'Maximum Score',
    'Percent Grade',
    'Points'
  ].forEach(header =>
    gradeRecords_setColumnFormat_(
      sheet,
      headers,
      header,
      '0.00',
      validationRows
    )
  );

  sheet.autoResizeColumns(
    1,
    lastColumn
  );

  const widthOverrides = {
    'Grade Record ID': 165,
    'Student ID': 90,
    'Student': 170,
    'Assignment / Evaluation Name': 230,
    'Instructor': 170,
    'Source ID': 170,
    'Notes': 300,
    'Supersedes Record ID': 170
  };

  Object.keys(widthOverrides)
    .forEach(header => {

      const column =
        headers[header];

      if (column) {
        sheet.setColumnWidth(
          column,
          widthOverrides[header]
        );
      }
    });
}


function gradeRecords_applyListValidation_(
  sheet,
  headers,
  header,
  values,
  rowCount
) {

  const column =
    headers[header];

  if (!column) {
    return;
  }

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireValueInList(
        values,
        true
      )
      .setAllowInvalid(false)
      .build();

  sheet
    .getRange(
      2,
      column,
      rowCount,
      1
    )
    .setDataValidation(rule);
}


function gradeRecords_applyCheckboxValidation_(
  sheet,
  headers,
  header,
  rowCount
) {

  const column =
    headers[header];

  if (!column) {
    return;
  }

  const rule =
    SpreadsheetApp
      .newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false)
      .build();

  sheet
    .getRange(
      2,
      column,
      rowCount,
      1
    )
    .setDataValidation(rule);
}


function gradeRecords_setColumnFormat_(
  sheet,
  headers,
  header,
  format,
  rowCount
) {

  const column =
    headers[header];

  if (!column) {
    return;
  }

  sheet
    .getRange(
      2,
      column,
      rowCount,
      1
    )
    .setNumberFormat(format);
}


/**
 * =====================================================================
 * SHEET RECORD HELPERS
 * =====================================================================
 */
function gradeRecords_getSheetObjects_(sheet) {

  const values =
    sheet
      .getDataRange()
      .getValues();

  if (values.length < 2) {
    return [];
  }

  const headers =
    values[0].map(value =>
      gradeRecords_clean_(value)
    );

  return values
    .slice(1)
    .map((row, index) => {

      const record = {
        __rowNumber: index + 2
      };

      headers.forEach((header, column) => {
        if (header) {
          record[header] = row[column];
        }
      });

      return record;
    })
    .filter(record =>
      gradeRecords_clean_(
        record['Grade Record ID']
      )
    );
}


function gradeRecords_findRowById_(
  sheet,
  gradeRecordId
) {

  const normalized =
    gradeRecords_normalize_(
      gradeRecordId
    );

  const record =
    gradeRecords_getSheetObjects_(sheet)
      .find(item =>
        gradeRecords_normalize_(
          item['Grade Record ID']
        ) ===
        normalized
      );

  if (!record) {
    return null;
  }

  return {
    rowNumber: record.__rowNumber,
    record: record
  };
}


function gradeRecords_headerMap_(sheet) {

  const values =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0];

  const map = {};

  values.forEach((value, index) => {

    const header =
      gradeRecords_clean_(value);

    if (header) {
      map[header] = index + 1;
    }
  });

  return map;
}


function gradeRecords_setCell_(
  sheet,
  rowNumber,
  headers,
  header,
  value
) {

  const column =
    headers[header];

  if (!column) {
    throw new Error(
      'Grade Records is missing required header: ' +
      header
    );
  }

  sheet
    .getRange(
      rowNumber,
      column
    )
    .setValue(value);
}


/**
 * =====================================================================
 * GENERAL HELPERS
 * =====================================================================
 */
function gradeRecords_publicRecord_(record) {

  return {
    gradeRecordId:
      gradeRecords_clean_(
        record['Grade Record ID']
      ),
    studentId:
      gradeRecords_clean_(
        record['Student ID']
      ),
    student:
      gradeRecords_clean_(
        record['Student']
      ),
    clinicalTerm:
      gradeRecords_clean_(
        record['Clinical Term']
      ),
    gradeCategory:
      gradeRecords_clean_(
        record['Grade Category']
      ),
    assignmentName:
      gradeRecords_clean_(
        record['Assignment / Evaluation Name']
      ),
    evaluationPeriod:
      gradeRecords_clean_(
        record['Evaluation Period']
      ),
    date:
      gradeRecords_formatDate_(
        record['Date']
      ),
    instructor:
      gradeRecords_clean_(
        record['Instructor']
      ),
    rawScore:
      gradeRecords_blankNumber_(
        record['Raw Score']
      ),
    maximumScore:
      gradeRecords_blankNumber_(
        record['Maximum Score']
      ),
    percentGrade:
      gradeRecords_blankNumber_(
        record['Percent Grade']
      ),
    points:
      gradeRecords_blankNumber_(
        record['Points']
      ),
    deliveryMethod:
      gradeRecords_clean_(
        record['Delivery Method']
      ),
    source:
      gradeRecords_clean_(
        record['Source']
      ),
    sourceId:
      gradeRecords_clean_(
        record['Source ID']
      ),
    sourceDocument:
      gradeRecords_clean_(
        record['Source Document']
      ),
    sourceFileId:
      gradeRecords_clean_(
        record['Source File ID']
      ),
    enteredBy:
      gradeRecords_clean_(
        record['Entered By']
      ),
    enteredAt:
      gradeRecords_formatDateTime_(
        record['Entered At']
      ),
    notes:
      gradeRecords_clean_(
        record['Notes']
      ),
    active:
      gradeRecords_isTrue_(
        record['Active']
      ),
    supersedesRecordId:
      gradeRecords_clean_(
        record['Supersedes Record ID']
      )
  };
}


function gradeRecords_parseDate_(value) {

  if (!value) {
    return null;
  }

  if (value instanceof Date) {

    if (isNaN(value.getTime())) {
      return null;
    }

    return new Date(
      value.getFullYear(),
      value.getMonth(),
      value.getDate()
    );
  }

  const text =
    String(value).trim();

  const iso =
    text.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})$/
    );

  if (iso) {

    const date =
      new Date(
        Number(iso[1]),
        Number(iso[2]) - 1,
        Number(iso[3])
      );

    return isNaN(date.getTime())
      ? null
      : date;
  }

  const date =
    new Date(text);

  if (isNaN(date.getTime())) {
    return null;
  }

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );
}


function gradeRecords_formatDate_(value) {

  const date =
    gradeRecords_parseDate_(value);

  if (!date) {
    return '';
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


function gradeRecords_formatDateTime_(value) {

  if (!value) {
    return '';
  }

  const date =
    value instanceof Date
      ? value
      : new Date(value);

  if (isNaN(date.getTime())) {
    return '';
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
}


function gradeRecords_dateValue_(value) {

  if (!value) {
    return 0;
  }

  const date =
    value instanceof Date
      ? value
      : new Date(value);

  return isNaN(date.getTime())
    ? 0
    : date.getTime();
}


function gradeRecords_currentUser_() {

  return (
    Session
      .getActiveUser()
      .getEmail() ||
    Session
      .getEffectiveUser()
      .getEmail() ||
    'Unknown User'
  );
}


function gradeRecords_blankNumber_(value) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return null;
  }

  return gradeRecords_number_(value);
}


function gradeRecords_number_(value) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {
    return null;
  }

  if (
    typeof value === 'string'
  ) {

    const text =
      value
        .replace(/%/g, '')
        .trim();

    if (!text) {
      return null;
    }

    value = text;
  }

  const numeric =
    Number(value);

  return isNaN(numeric)
    ? null
    : numeric;
}


function gradeRecords_round2_(value) {

  return Math.round(
    (Number(value) + Number.EPSILON) *
    100
  ) / 100;
}


function gradeRecords_first_() {

  for (
    let index = 0;
    index < arguments.length;
    index++
  ) {

    const value =
      arguments[index];

    if (
      value !== '' &&
      value !== null &&
      value !== undefined
    ) {
      return value;
    }
  }

  return '';
}


function gradeRecords_clean_(value) {

  return String(
    value === null ||
    value === undefined
      ? ''
      : value
  ).trim();
}


function gradeRecords_normalize_(value) {

  return gradeRecords_clean_(value)
    .toLowerCase()
    .replace(/\s+/g, ' ');
}


function gradeRecords_isTrue_(value) {

  if (value === true) {
    return true;
  }

  const normalized =
    gradeRecords_normalize_(value);

  return (
    normalized === 'true' ||
    normalized === 'yes' ||
    normalized === 'y' ||
    normalized === '1' ||
    normalized === 'active'
  );
}


/**
 * =====================================================================
 * SAFE TESTS
 * =====================================================================
 */
function testGradeRecordsProfessionalConversion() {

  const cases = [
    [3.00, 100],
    [2.95, 95.75],
    [2.20, 77],
    [2.19, 76.75],
    [2.18, 76],
    [2.10, 75],
    [2.00, 75],
    [1.99, 75],
    [1.50, 65],
    [1.00, 55],
    [0.50, 50],
    [0.00, 45]
  ];

  cases.forEach(testCase => {

    const actual =
      gradeRecords_professionalScoreToPercent_(
        testCase[0]
      );

    if (actual !== testCase[1]) {
      throw new Error(
        'Conversion failed for ' +
        testCase[0] +
        ': expected ' +
        testCase[1] +
        ', received ' +
        actual
      );
    }
  });

  Logger.log(
    'Professional Competency conversion tests passed.'
  );
}


/**
 * Non-writing regression for Tests A and B. Confirms the same term predicate
 * used by Grade Engine keeps historical Grade Records in Clinical 1 and out
 * of the student's current Clinical 3 summary.
 */
function testGradeRecordsHistoricalTermRegression() {

  const records = [
    {
      'Grade Record ID': 'TEST-A',
      'Student ID': 'DEPT #22',
      'Clinical Term': 'Clinical 1',
      'Grade Category': 'Professional Competency',
      'Percent Grade': 89.38,
      'Active': true
    },
    {
      'Grade Record ID': 'TEST-B',
      'Student ID': 'DEPT #22',
      'Clinical Term': 'Clinical 1',
      'Grade Category': 'Clinical Test',
      'Percent Grade': 89.70,
      'Active': true
    }
  ];

  const forTerm = term =>
    records.filter(record =>
      gradeRecords_normalize_(record['Clinical Term']) ===
      gradeRecords_normalize_(term)
    );

  const clinical1 = forTerm('Clinical 1');
  const clinical3 = forTerm('Clinical 3');

  if (clinical1.length !== 2 || clinical3.length !== 0) {
    throw new Error(
      'Historical Grade Record term regression failed. ' +
      'Clinical 1 count=' + clinical1.length +
      ', Clinical 3 count=' + clinical3.length + '.'
    );
  }

  const clinical1Summary =
    gradeEngine_buildTermSummary_({
      clinicalTerm: 'Clinical 1',
      tracking: [],
      gradeRecords: clinical1
    });

  const clinical3Summary =
    gradeEngine_buildTermSummary_({
      clinicalTerm: 'Clinical 3',
      tracking: [],
      gradeRecords: clinical3
    });

  if (
    !clinical1Summary.categories.professional.available ||
    !clinical1Summary.categories.clinicalTest.available ||
    clinical3Summary.categories.professional.available ||
    clinical3Summary.categories.clinicalTest.available
  ) {
    throw new Error(
      'Grade Engine placed a historical Grade Record in the wrong semester.'
    );
  }

  return {
    success: true,
    clinical1: {
      recordCount: clinical1.length,
      professional: clinical1Summary.categories.professional.percentAverage,
      clinicalTest: clinical1Summary.categories.clinicalTest.percentAverage
    },
    clinical3: {
      recordCount: clinical3.length,
      professionalAvailable: clinical3Summary.categories.professional.available,
      clinicalTestAvailable: clinical3Summary.categories.clinicalTest.available
    }
  };
}


/**
 * Live, writing attachment regression. Run manually only in the deployed
 * Apps Script project. It uses TEST-001 exclusively and refuses to run twice
 * with the same Source ID.
 */
function testNewEntryAttachment_TEST001() {

  const sourceId = 'TEST-001-ATTACHMENT';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const student = gradeRecords_resolveStudent_(ss, 'TEST-001', '');

  const existing =
    gradeRecords_list(student.studentId, '', true)
      .filter(record => record.sourceId === sourceId);

  if (existing.length > 0) {
    throw new Error(
      'The TEST-001 attachment regression already exists. ' +
      'Void or inspect that record instead of creating a duplicate.'
    );
  }

  const trackingSheet =
    appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_TRACKING);

  const trackingBefore =
    appSheetToObjects_(trackingSheet)
      .filter(record =>
        gradeRecords_normalize_(record['Student ID']) ===
        gradeRecords_normalize_(student.studentId)
      ).length;

  const pdfBytes =
    Utilities.newBlob(
      '%PDF-1.4\n% TEST-001 supporting-document regression\n%%EOF',
      'application/pdf',
      'TEST-001.pdf'
    ).getBytes();

  const result =
    dashboardSaveNewEntry({
      studentId: student.studentId,
      clinicalTerm: 'Clinical 1',
      gradeCategory: 'quizzes',
      assignmentName: 'TEST-001 Attachment Regression',
      date: '2026-09-01',
      instructor: 'Automated Regression Test',
      percentGrade: 88,
      source: 'Regression Test',
      sourceId: sourceId,
      notes: 'TEST-001 attachment association regression.',
      attachment: {
        fileName: 'TEST-001.pdf',
        mimeType: 'application/pdf',
        data: Utilities.base64Encode(pdfBytes)
      }
    });

  const matching =
    gradeRecords_list(student.studentId, '', true)
      .filter(record => record.sourceId === sourceId);

  const trackingAfter =
    appSheetToObjects_(trackingSheet)
      .filter(record =>
        gradeRecords_normalize_(record['Student ID']) ===
        gradeRecords_normalize_(student.studentId)
      ).length;

  if (matching.length !== 1) {
    throw new Error(
      'Expected exactly one TEST-001 Grade Record; found ' +
      matching.length + '.'
    );
  }

  const record = matching[0];

  if (
    record.clinicalTerm !== 'Clinical 1' ||
    !record.sourceDocument ||
    !record.sourceFileId ||
    trackingAfter !== trackingBefore
  ) {
    throw new Error(
      'TEST-001 attachment regression failed: ' +
      JSON.stringify({
        clinicalTerm: record.clinicalTerm,
        sourceDocument: record.sourceDocument,
        sourceFileId: record.sourceFileId,
        trackingBefore: trackingBefore,
        trackingAfter: trackingAfter
      })
    );
  }

  const file = DriveApp.getFileById(record.sourceFileId);

  return {
    success: true,
    gradeRecordId: record.gradeRecordId,
    gradeRecordCount: matching.length,
    clinicalTerm: record.clinicalTerm,
    attachment: {
      fileId: file.getId(),
      fileName: file.getName(),
      fileUrl: file.getUrl(),
      folderName: result.attachment.folderName
    },
    clinicalTrackingCountBefore: trackingBefore,
    clinicalTrackingCountAfter: trackingAfter
  };
}


function testGradeRecordsConfiguration() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    gradeRecords_getSheet_(ss);

  const headers =
    gradeRecords_headerMap_(sheet);

  const missing =
    GRADE_RECORDS_CONFIG.headers
      .filter(header =>
        !headers[header]
      );

  if (missing.length > 0) {
    throw new Error(
      'Grade Records is missing: ' +
      missing.join(', ')
    );
  }

  Logger.log(
    'Grade Records configuration test passed.'
  );
}
