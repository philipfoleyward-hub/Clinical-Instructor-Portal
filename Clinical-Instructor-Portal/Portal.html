/**
 * CCMS Clinical Instructor Portal
 * PortalService.gs
 *
 * PURPOSE
 * Thin portal-facing orchestration layer. This file does not own scheduling
 * rules or academic requirement rules. It reuses authoritative CCMS outputs.
 */

const CI_PORTAL_SERVICE = Object.freeze({
  STUDENTS_SHEET: 'Students',
  REQUIREMENT_STATUS_SHEET: 'Student Requirement Status'
});

/**
 * Home read model for a named College CI.
 * Scheduling authority remains SchedulingService.gs.
 */
function portalGetHomeData(evaluatorId, date) {
  const id = String(evaluatorId || '').trim();
  if (!id) throw new Error('Evaluator ID is required.');

  const targetDate = date ? new Date(date) : new Date();
  if (isNaN(targetDate.getTime())) throw new Error('A valid date is required.');

  const evaluator = ciPortalGetEvaluatorById_(id);
  if (!evaluator || !ciPortalIsTrue_(evaluator['Active']) || !ciPortalIsTrue_(evaluator['Portal Access'])) {
    throw new Error('Unable to resolve active portal evaluator.');
  }

  const workspace = scheduling_getCiDayWorkspace(id, targetDate);

  return {
    success: true,
    evaluator: {
      evaluatorId: id,
      evaluatorName: String(evaluator['Evaluator'] || '').trim(),
      evaluatorType: String(evaluator['Evaluator Type'] || '').trim()
    },
    date: targetDate,
    assigned: workspace.assigned === true,
    assignments: workspace.assignments || [],

    // Intentionally empty until their authoritative evaluation/notification
    // paths are inspected and wired. Do not fabricate prototype records.
    inProgressEvaluations: [],
    recentlyCompleted: [],
    notifications: [],
    actionItems: []
  };
}

/**
 * CI-facing competency-opportunity read model.
 *
 * Academic authority: RequirementsEngine output already persisted in
 * Student Requirement Status. This function filters/presents that output;
 * it does not calculate eligibility, prerequisites, or requirement timing.
 */
function portalGetCompetencyOpportunities(studentId) {
  const id = String(studentId || '').trim();
  if (!id) throw new Error('Student ID is required.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const student = portalGetStudentRecord_(ss, id);
  if (!student) throw new Error('Student not found: ' + id);

  const sheet = ss.getSheetByName(CI_PORTAL_SERVICE.REQUIREMENT_STATUS_SHEET);
  if (!sheet) throw new Error('Student Requirement Status sheet not found.');

  const rows = portalSheetToObjects_(sheet);
  const details = rows.filter(row =>
    portalNorm_(row['Student ID']) === portalNorm_(id) &&
    portalNorm_(row['Status Type']) === 'DETAIL' &&
    portalNorm_(row['Requirement Scope']) === 'PROGRAM' &&
    portalNorm_(row['Status']) === 'STILL NEEDED'
  );

  const mapped = details.map(portalRequirementItem_);

  const eligibleNow = mapped.filter(x => portalNorm_(x.timing) === 'ELIGIBLE NOW');
  const prerequisiteReview = mapped.filter(x => portalNorm_(x.timing) === 'PREREQUISITE VERIFICATION NEEDED');
  const future = mapped.filter(x => portalNorm_(x.timing) === 'FUTURE REQUIREMENT');
  const other = mapped.filter(x =>
    ['ELIGIBLE NOW', 'PREREQUISITE VERIFICATION NEEDED', 'FUTURE REQUIREMENT']
      .indexOf(portalNorm_(x.timing)) === -1
  );

  return {
    success: true,
    studentId: id,
    student: String(student['Student'] || student['Student Name'] || '').trim(),
    clinicalTerm: portalFirstNonBlank_(details, 'Clinical Term') || String(student['Clinical Term'] || '').trim(),
    counts: {
      eligibleNow: eligibleNow.length,
      prerequisiteReview: prerequisiteReview.length,
      future: future.length,
      other: other.length
    },
    eligibleNow: eligibleNow,
    prerequisiteReview: prerequisiteReview,
    future: future,
    other: other
  };
}

function portalRequirementItem_(row) {
  return {
    requirementId: String(row['Requirement ID'] || '').trim(),
    name: String(row['Requirement Name'] || '').trim(),
    type: String(row['Requirement Type'] || '').trim(),
    group: String(row['Requirement Group'] || '').trim(),
    timing: String(row['Requirement Timing'] || '').trim(),
    status: String(row['Status'] || '').trim(),
    notes: String(row['Notes'] || '').trim()
  };
}

function portalGetStudentRecord_(ss, studentId) {
  const sheet = ss.getSheetByName(CI_PORTAL_SERVICE.STUDENTS_SHEET);
  if (!sheet) throw new Error('Students sheet not found.');
  const target = portalNorm_(studentId);
  return portalSheetToObjects_(sheet).find(row => portalNorm_(row['Student ID']) === target) || null;
}

function portalSheetToObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return [];
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map(v => String(v || '').trim());
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, i) => { if (header) obj[header] = row[i]; });
    return obj;
  });
}

function portalFirstNonBlank_(rows, header) {
  for (let i = 0; i < rows.length; i++) {
    const value = String(rows[i][header] || '').trim();
    if (value) return value;
  }
  return '';
}

function portalNorm_(value) {
  return String(value == null ? '' : value).trim().toUpperCase();
}
