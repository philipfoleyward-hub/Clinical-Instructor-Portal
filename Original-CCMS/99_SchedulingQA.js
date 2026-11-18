/**
 * =====================================================================
 * CCMS - SCHEDULING FOUNDATION QA - PHASE 1 COMPLETE
 * 99_SchedulingQA.gs
 * =====================================================================
 * READ-ONLY.
 *
 * No production workflow depends on this file.
 *
 * QA coverage:
 * - Scheduling sheet structure
 * - Scheduling relational/data integrity
 * - TEST-001 single-student resolver
 * - Full foundation validation
 * - Real Fall 2026 Class of 2027 cohort resolution
 * - Cohort representation normalization
 * =====================================================================
 */


/**
 * Validate all Scheduling foundation sheet structures.
 */
function testSchedulingSheetStructure() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const expected = {};

  expected[
    APP_CONFIG.SHEETS.CLINICAL_CALENDAR
  ] = [
    'Calendar ID',
    'Clinical Term',
    'Cohort',
    'Date',
    'Day Type',
    'Scheduled Hours',
    'Start Time',
    'End Time',
    'Notes',
    'Active'
  ];

  expected[
    APP_CONFIG.SHEETS.STUDENT_SITE_ASSIGNMENTS
  ] = [
    'Assignment ID',
    'Student ID',
    'Clinical Term',
    'Site ID',
    'Start Date',
    'End Date',
    'Assignment Type',
    'Status',
    'Published At',
    'Published By',
    'Notes'
  ];

  expected[
    APP_CONFIG.SHEETS.STUDENT_ROOM_ROTATIONS
  ] = [
    'Rotation ID',
    'Assignment ID',
    'Student ID',
    'Site ID',
    'Start Date',
    'End Date',
    'Rotation Area',
    'Priority Note',
    'Notes',
    'Status'
  ];

  expected[
    APP_CONFIG.SHEETS.STUDENT_SITE_PREFERENCES
  ] = [
    'Preference ID',
    'Student ID',
    'Clinical Term',
    'Preference 1 Site ID',
    'Preference 2 Site ID',
    'Preference 3 Site ID',
    'Preference 4 Site ID',
    'Conflict or Restriction',
    'Submitted At',
    'Source Response ID'
  ];

  expected[
    APP_CONFIG.SHEETS.CI_AVAILABILITY
  ] = [
    'Availability ID',
    'Evaluator ID',
    'Date',
    'Start Time',
    'End Time',
    'Availability Status',
    'Calendar Event ID',
    'Last Synced'
  ];

  expected[
    APP_CONFIG.SHEETS.CI_ASSIGNMENTS
  ] = [
    'CI Assignment ID',
    'Evaluator ID',
    'Date',
    'Site ID',
    'Start Time',
    'End Time',
    'Assignment Type',
    'Status',
    'Published At',
    'Published By',
    'Notes'
  ];

  expected[
    APP_CONFIG.SHEETS.ATTENDANCE_EVENTS
  ] = [
    'Attendance Event ID',
    'Student ID',
    'Date',
    'Clinical Term',
    'Event Type',
    'Hours',
    'ETO Impact',
    'Make-Up Required',
    'Related Event ID',
    'Source',
    'Approval Status',
    'Approved By',
    'Approved At',
    'Notes',
    'Created At',
    'Created By',
    'Supersedes Event ID',
    'Amendment Reason'
  ];

  expected[
    APP_CONFIG.SHEETS.USERS
  ] = [
    'User ID',
    'Email',
    'Display Name',
    'Role',
    'Evaluator ID',
    'Student ID',
    'Active',
    'Effective From',
    'Effective To',
    'Notes'
  ];

  const results = {};

  Object.keys(expected)
    .forEach(sheetName => {
      const sheet =
        ss.getSheetByName(
          sheetName
        );

      if (!sheet) {
        results[sheetName] = {
          success: false,
          missingSheet: true,
          missingHeaders: []
        };
        return;
      }

      const headers =
        sheet
          .getRange(
            1,
            1,
            1,
            Math.max(
              sheet.getLastColumn(),
              1
            )
          )
          .getValues()[0]
          .map(v =>
            String(
              v || ''
            ).trim()
          );

      const normalized =
        headers.map(
          appNormalize_
        );

      const missing =
        expected[sheetName]
          .filter(h =>
            !normalized.includes(
              appNormalize_(h)
            )
          );

      results[sheetName] = {
        success:
          missing.length === 0,
        missingSheet: false,
        missingHeaders: missing
      };
    });

  const failed =
    Object.keys(results)
      .filter(name =>
        !results[name].success
      );

  const output = {
    success:
      failed.length === 0,
    failedSheets: failed,
    sheets: results
  };

  console.log(
    JSON.stringify(
      output,
      null,
      2
    )
  );

  if (failed.length) {
    throw new Error(
      'Scheduling sheet structure failed for: ' +
      failed.join(', ')
    );
  }

  return output;
}


/**
 * Run production Scheduling integrity validation.
 */
function testSchedulingDataIntegrity() {
  return scheduling_validateSchedulingData();
}


/**
 * Legacy Phase 1 single-student resolver QA.
 *
 * TEST-001 must remain available for this test to resolve successfully.
 */
function testSchedulingTest001() {
  const result =
    scheduling_getStudentScheduleForDate(
      'TEST-001',
      new Date(
        2026,
        8,
        16
      )
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
 * Original Phase 1 foundation QA.
 */
function testSchedulingFoundation() {
  const structure =
    testSchedulingSheetStructure();

  const integrity =
    testSchedulingDataIntegrity();

  return {
    success:
      structure.success &&
      integrity.success,
    structure: structure,
    integrity: integrity
  };
}


/**
 * Real Fall 2026 cohort resolver QA.
 *
 * Tests Class of 2027 on Wednesday September 16, 2026.
 *
 * Expected:
 * - 13 active students
 * - all scheduled
 * - all Clinical 3
 * - all have valid site IDs
 * - all have resolved site names
 * - all have room/rotation assignments
 */
function testSchedulingCohortRealData() {
  const cohort =
    'Class of 2027';

  const date =
    new Date(
      2026,
      8,
      16
    );

  const result =
    scheduling_getCohortScheduleForDate(
      cohort,
      date
    );

  const students =
    result.map(row => ({
      studentId:
        row.studentId,
      student:
        row.student,
      cohort:
        row.cohort,
      date:
        scheduling_formatDate_(
          row.date
        ),
      scheduled:
        row.scheduled,
      status:
        row.status,
      clinicalTerm:
        row.clinicalTerm,
      siteId:
        row.siteId,
      site:
        row.site,
      rotationArea:
        row.rotationArea
    }));

  const problems =
    students.filter(row =>
      row.scheduled !== true ||
      row.status !== 'Scheduled' ||
      appNormalize_(
        row.clinicalTerm
      ) !==
        appNormalize_(
          'Clinical 3'
        ) ||
      !row.siteId ||
      !row.site ||
      !row.rotationArea
    );

  const output = {
    success:
      students.length === 13 &&
      problems.length === 0,
    cohort: cohort,
    date:
      scheduling_formatDate_(
        date
      ),
    expectedStudentCount: 13,
    actualStudentCount:
      students.length,
    scheduledCount:
      students.filter(
        row =>
          row.status ===
          'Scheduled'
      ).length,
    problemCount:
      problems.length,
    problems: problems,
    students: students
  };

  console.log(
    JSON.stringify(
      output,
      null,
      2
    )
  );

  if (
    students.length !== 13
  ) {
    throw new Error(
      'Real cohort QA failed: expected 13 active Class of 2027 students but resolved ' +
      students.length +
      '.'
    );
  }

  if (problems.length) {
    throw new Error(
      'Real cohort QA failed: ' +
      problems.length +
      ' student schedule(s) did not resolve completely. See execution log.'
    );
  }

  return output;
}


/**
 * Verify cohort representation normalization independently.
 *
 * This intentionally tests both the historical compact representation
 * and the current display representation.
 *
 * It does not modify Students or Clinical Calendar.
 */
function testSchedulingCohortNormalization() {
  const date =
    new Date(
      2026,
      8,
      16
    );

  const compact =
    scheduling_getCohortScheduleForDate(
      '2027',
      date
    );

  const display =
    scheduling_getCohortScheduleForDate(
      'Class of 2027',
      date
    );

  const compactIds =
    compact
      .map(row =>
        appNormalize_(
          row.studentId
        )
      )
      .sort();

  const displayIds =
    display
      .map(row =>
        appNormalize_(
          row.studentId
        )
      )
      .sort();

  const same =
    compactIds.length ===
      displayIds.length &&
    compactIds.every(
      (id, index) =>
        id ===
        displayIds[index]
    );

  const output = {
    success:
      same,
    compactRepresentationCount:
      compact.length,
    displayRepresentationCount:
      display.length,
    sameStudentSet:
      same
  };

  console.log(
    JSON.stringify(
      output,
      null,
      2
    )
  );

  if (!same) {
    throw new Error(
      'Cohort normalization QA failed: "2027" and "Class of 2027" did not resolve to the same student set.'
    );
  }

  return output;
}


/**
 * Historical Fall 2026 import baseline.
 *
 * This intentionally preserves the original Phase 1 count expectations
 * as an explicit regression reference. It is NOT the general health gate
 * after Room Rotations become writable.
 */
function testSchedulingFall2026BaselineCounts() {
  const integrity =
    testSchedulingDataIntegrity();

  const expectedCounts = {
    calendarRows: 41,
    siteAssignments: 182,
    roomRotations: 182
  };

  const countCheck =
    integrity.counts.calendarRows ===
      expectedCounts.calendarRows &&
    integrity.counts.siteAssignments ===
      expectedCounts.siteAssignments &&
    integrity.counts.roomRotations ===
      expectedCounts.roomRotations;

  const output = {
    success: countCheck,
    counts: integrity.counts,
    expectedCounts: expectedCounts
  };

  console.log(
    JSON.stringify(
      output,
      null,
      2
    )
  );

  return output;
}


/**
 * Read-only Room Rotations v1 contract QA.
 *
 * Fixture-independent by design. It uses an existing active site assignment
 * from the workbook rather than requiring TEST-001 to have scheduling data.
 *
 * Confirms:
 * - the Status column exists;
 * - scheduling integrity remains clean;
 * - a hypothetical Draft validates against a real parent assignment;
 * - the official resolver remains readable for the same assignment;
 * - mismatched Student ID, mismatched Site ID, and dates outside the parent
 *   assignment are rejected;
 * - no workbook data is written.
 */
function testSchedulingRoomRotationsV1Readiness() {
  const structure =
    testSchedulingSheetStructure();

  const integrity =
    testSchedulingDataIntegrity();

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const assignmentRows =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENT_SITE_ASSIGNMENTS
      )
    )
      .filter(row =>
        String(
          row['Assignment ID'] || ''
        ).trim() &&
        String(
          row['Student ID'] || ''
        ).trim() &&
        String(
          row['Site ID'] || ''
        ).trim() &&
        scheduling_isActiveAssignment_(
          row
        ) &&
        appDateOnly_(
          row['Start Date']
        ) &&
        appDateOnly_(
          row['End Date']
        )
      );

  if (!assignmentRows.length) {
    throw new Error(
      'Room Rotations v1 QA requires at least one valid active Student Site Assignment.'
    );
  }

  const parent =
    assignmentRows[0];

  const validatedDraft =
    scheduling_validateRoomRotationDraft({
      assignmentId:
        parent['Assignment ID'],
      studentId:
        parent['Student ID'],
      siteId:
        parent['Site ID'],
      startDate:
        parent['Start Date'],
      endDate:
        parent['End Date'],
      rotationArea:
        'QA Draft Area'
    });

  const official =
    scheduling_getStudentRoomRotation(
      parent['Student ID'],
      parent['Start Date'],
      parent['Assignment ID']
    );

  const rejectionChecks = {
    studentMismatch: false,
    siteMismatch: false,
    outsideAssignmentDates: false
  };

  try {
    scheduling_validateRoomRotationDraft({
      assignmentId:
        parent['Assignment ID'],
      studentId:
        '__QA_WRONG_STUDENT__',
      siteId:
        parent['Site ID'],
      startDate:
        parent['Start Date'],
      endDate:
        parent['End Date'],
      rotationArea:
        'QA Draft Area'
    });
  } catch (err) {
    rejectionChecks.studentMismatch =
      String(
        err && err.message || err
      ).indexOf(
        'Student ID does not match parent Assignment'
      ) !== -1;
  }

  try {
    scheduling_validateRoomRotationDraft({
      assignmentId:
        parent['Assignment ID'],
      studentId:
        parent['Student ID'],
      siteId:
        '__QA_WRONG_SITE__',
      startDate:
        parent['Start Date'],
      endDate:
        parent['End Date'],
      rotationArea:
        'QA Draft Area'
    });
  } catch (err) {
    rejectionChecks.siteMismatch =
      String(
        err && err.message || err
      ).indexOf(
        'Site ID does not match parent Assignment'
      ) !== -1;
  }

  const outsideDate =
    new Date(
      appDateOnly_(
        parent['End Date']
      ).getTime() +
      24 * 60 * 60 * 1000
    );

  try {
    scheduling_validateRoomRotationDraft({
      assignmentId:
        parent['Assignment ID'],
      studentId:
        parent['Student ID'],
      siteId:
        parent['Site ID'],
      startDate:
        parent['Start Date'],
      endDate:
        outsideDate,
      rotationArea:
        'QA Draft Area'
    });
  } catch (err) {
    rejectionChecks.outsideAssignmentDates =
      String(
        err && err.message || err
      ).indexOf(
        'Room Rotation dates must remain inside parent Assignment'
      ) !== -1;
  }

  const rejectionSuccess =
    rejectionChecks.studentMismatch &&
    rejectionChecks.siteMismatch &&
    rejectionChecks.outsideAssignmentDates;

  const output = {
    success:
      structure.success &&
      integrity.success &&
      integrity.issueCount === 0 &&
      validatedDraft.status ===
        'Draft' &&
      rejectionSuccess,
    structure:
      structure.success,
    integrity: {
      success:
        integrity.success,
      issueCount:
        integrity.issueCount,
      warningCount:
        integrity.warningCount,
      counts:
        integrity.counts
    },
    fixture: {
      source:
        'existing active Student Site Assignment',
      assignmentId:
        parent['Assignment ID'],
      studentId:
        parent['Student ID'],
      siteId:
        parent['Site ID']
    },
    validatedDraft:
      validatedDraft,
    authoritativeRotationId:
      official
        ? official['Rotation ID']
        : null,
    rejectionChecks:
      rejectionChecks,
    writesPerformed: false
  };

  console.log(
    JSON.stringify(
      output,
      null,
      2
    )
  );

  if (!output.success) {
    throw new Error(
      'Scheduling Room Rotations v1 readiness QA failed. Review execution log.'
    );
  }

  return output;
}


/**
 * Complete Scheduling production-health QA.
 *
 * Legacy function name retained for compatibility. Mutable production
 * row counts are reported but no longer treated as a universal pass/fail
 * condition now that Room Rotations support Draft writes.
 */
function testSchedulingPhase1Complete() {
  const structure =
    testSchedulingSheetStructure();

  const integrity =
    testSchedulingDataIntegrity();

  const cohortNormalization =
    testSchedulingCohortNormalization();

  const realCohort =
    testSchedulingCohortRealData();

  const roomRotationsV1 =
    testSchedulingRoomRotationsV1Readiness();

  const historicalBaseline =
    testSchedulingFall2026BaselineCounts();

  const output = {
    success:
      structure.success &&
      integrity.success &&
      integrity.issueCount === 0 &&
      cohortNormalization.success &&
      realCohort.success &&
      roomRotationsV1.success,

    structure:
      structure.success,

    integrity: {
      success:
        integrity.success,
      issueCount:
        integrity.issueCount,
      warningCount:
        integrity.warningCount,
      counts:
        integrity.counts
    },

    historicalBaseline:
      historicalBaseline,

    cohortNormalization:
      cohortNormalization,

    realCohort: {
      success:
        realCohort.success,
      actualStudentCount:
        realCohort.actualStudentCount,
      scheduledCount:
        realCohort.scheduledCount,
      problemCount:
        realCohort.problemCount
    },

    roomRotationsV1:
      roomRotationsV1
  };

  console.log(
    '============================================'
  );
  console.log(
    'CCMS SCHEDULING HEALTH QA'
  );
  console.log(
    '============================================'
  );
  console.log(
    JSON.stringify(
      output,
      null,
      2
    )
  );
  console.log(
    '============================================'
  );

  if (!output.success) {
    throw new Error(
      'Scheduling health QA failed. Review execution log.'
    );
  }

  return output;
}
/**
 * Verify known real CI assignment resolution.
 *
 * EV-004 has a current SITE-004 assignment on August 26, 2026.
 */
function testSchedulingCiAssignmentRealData() {
  const result =
    scheduling_getCiAssignmentsForDate(
      'EV-004',
      new Date(2026, 7, 26)
    );

  const success =
    result.length === 1 &&
    appNormalize_(
      result[0]['Site ID']
    ) === appNormalize_('SITE-004');

  const output = {
    success: success,
    expectedEvaluatorId: 'EV-004',
    expectedSiteId: 'SITE-004',
    actualAssignmentCount: result.length,
    assignments: result
  };

  console.log(
    JSON.stringify(output, null, 2)
  );

  if (!success) {
    throw new Error(
      'CI assignment QA failed for EV-004 on 2026-08-26.'
    );
  }

  return output;
}


/**
 * Verify a CI/date with no assignment returns an intentional empty state.
 */
function testSchedulingCiNoAssignment() {
  const result =
    scheduling_getCiAssignmentsForDate(
      'EV-004',
      new Date(2026, 0, 1)
    );

  const success =
    Array.isArray(result) &&
    result.length === 0;

  if (!success) {
    throw new Error(
      'CI no-assignment QA failed.'
    );
  }

  return {
    success: true,
    assignmentCount: 0
  };
}


/**
 * Verify CI Home scheduling composition.
 */
function testSchedulingCiDayWorkspace() {
  const result =
    scheduling_getCiDayWorkspace(
      'EV-004',
      new Date(2026, 7, 26)
    );

  console.log(
    JSON.stringify(result, null, 2)
  );

  if (
    !result.assigned ||
    !result.assignments.length
  ) {
    throw new Error(
      'CI day workspace QA failed: expected assignment.'
    );
  }

  if (
    appNormalize_(
      result.assignments[0].siteId
    ) !== appNormalize_('SITE-004')
  ) {
    throw new Error(
      'CI day workspace QA failed: expected SITE-004.'
    );
  }

  return {
    success: true,
    evaluatorId: result.evaluatorId,
    assignmentCount:
      result.assignments.length,
    studentCount:
      result.assignments.reduce(
        (total, assignment) =>
          total + assignment.students.length,
        0
      ),
    workspace: result
  };
}


/**
 * Permanent CI scheduling extension regression runner.
 */
function testSchedulingCiReadModel() {
  const knownAssignment =
    testSchedulingCiAssignmentRealData();

  const noAssignment =
    testSchedulingCiNoAssignment();

  const dayWorkspace =
    testSchedulingCiDayWorkspace();

  const output = {
    success:
      knownAssignment.success &&
      noAssignment.success &&
      dayWorkspace.success,

    knownAssignment:
      knownAssignment,

    noAssignment:
      noAssignment,

    dayWorkspace:
      dayWorkspace
  };

  console.log(
    '============================================'
  );
  console.log(
    'CCMS CI SCHEDULING READ MODEL QA'
  );
  console.log(
    '============================================'
  );
  console.log(
    JSON.stringify(output, null, 2)
  );

  if (!output.success) {
    throw new Error(
      'CI Scheduling read-model QA failed.'
    );
  }

  return output;
}
function scheduling_formatTime_(value) {
  if (!value) {
    return '';
  }

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'h:mm a'
    );
  }

  return String(value).trim();
}