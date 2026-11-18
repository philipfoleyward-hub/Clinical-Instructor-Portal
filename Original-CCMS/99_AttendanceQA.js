/**
 * =====================================================================
 * CCMS ATTENDANCE ENTRY QA
 * 99_AttendanceQA.gs
 * =====================================================================
 *
 * Read-only tests are safe.
 *
 * testAttendanceEntryEndToEnd_TEST001() WRITES TWO QA ROWS:
 *   1) an approved QA absence;
 *   2) an approved zero-impact amendment that supersedes the QA absence.
 *
 * The final authoritative TEST-001 balance should therefore return to its
 * pre-test state while preserving both QA rows as an auditable history.
 * =====================================================================
 */

function testAttendanceEntryReadiness() {
  const structure = attendance_validateData();
  const identity = attendance_testCohortIdentity('Class of 2027');

  const result = {
    success: structure.success && identity.success,
    structure: structure,
    cohortIdentity: identity
  };

  console.log(JSON.stringify(result, null, 2));
  if (!result.success) throw new Error('Attendance Entry readiness QA failed.');
  return result;
}


function testAttendanceEntryValidation() {
  const checks = {};

  checks.missingStudent = attendanceQa_expectFailure_(function() {
    attendance_createEvent({
      clinicalTerm: 'Clinical 3',
      date: new Date(),
      eventType: 'ABSENCE',
      hours: 1
    });
  });

  checks.badType = attendanceQa_expectFailure_(function() {
    attendance_createEvent({
      studentId: 'TEST-001',
      clinicalTerm: 'Clinical 3',
      date: new Date(),
      eventType: 'NOT_A_REAL_EVENT',
      hours: 1
    });
  });

  checks.negativeHours = attendanceQa_expectFailure_(function() {
    attendance_createEvent({
      studentId: 'TEST-001',
      clinicalTerm: 'Clinical 3',
      date: new Date(),
      eventType: 'ABSENCE',
      hours: -1
    });
  });

  const success = Object.keys(checks).every(key => checks[key] === true);
  const result = {success: success, checks: checks};
  console.log(JSON.stringify(result, null, 2));
  if (!success) throw new Error('Attendance validation QA failed.');
  return result;
}


/**
 * MUTATING TEST. Run only when intentionally testing the production writer.
 * Leaves two clearly identified QA rows in Attendance Events for provenance.
 */
function testAttendanceEntryEndToEnd_TEST001() {
  const studentId = 'TEST-001';
  const term = 'Clinical 3';

  const rule = attendance_getRule(term, new Date());
  const eventDate = rule.effectiveStartDate || new Date();

  const before = attendance_getStudentTermSummary(studentId, term, eventDate);
  const beforeBalance = Number(before.balances.currentAvailableBalanceHours || 0);
  const beforeMissed = Number(before.attendance.totalMissedHours || 0);

  const created = attendance_createEvent({
    studentId: studentId,
    clinicalTerm: term,
    date: eventDate,
    eventType: 'ABSENCE',
    hours: 1,
    etoImpact: -1,
    makeUpRequired: false,
    approvalStatus: 'APPROVED',
    source: '99_AttendanceQA',
    notes: 'QA ONLY - temporary one-hour absence; superseded in same QA test.',
    reason: 'QA ONLY'
  });

  const afterCreate = attendance_getStudentTermSummary(studentId, term, eventDate);
  const createBalance = Number(afterCreate.balances.currentAvailableBalanceHours || 0);
  const createMissed = Number(afterCreate.attendance.totalMissedHours || 0);

  if (Math.abs(createBalance - (beforeBalance - 1)) > 0.001) {
    throw new Error(
      'QA absence did not deduct exactly 1 ETO hour. Before=' +
      beforeBalance + ', after=' + createBalance
    );
  }
  if (Math.abs(createMissed - (beforeMissed + 1)) > 0.001) {
    throw new Error(
      'QA absence did not add exactly 1 missed hour. Before=' +
      beforeMissed + ', after=' + createMissed
    );
  }

  const amended = attendance_amendEvent(
    created.attendanceEventId,
    {
      eventType: 'ATTENDANCE_ADJUSTMENT',
      hours: 0,
      etoImpact: 0,
      makeUpRequired: false,
      approvalStatus: 'APPROVED',
      source: '99_AttendanceQA',
      notes: 'QA ONLY - superseding zero-impact amendment.',
      reason: 'QA cleanup'
    },
    'QA test completed; supersede temporary absence without deleting history.'
  );

  const afterAmend = attendance_getStudentTermSummary(studentId, term, eventDate);
  const finalBalance = Number(afterAmend.balances.currentAvailableBalanceHours || 0);
  const finalMissed = Number(afterAmend.attendance.totalMissedHours || 0);

  const restored =
    Math.abs(finalBalance - beforeBalance) <= 0.001 &&
    Math.abs(finalMissed - beforeMissed) <= 0.001;

  const result = {
    success: restored,
    studentId: studentId,
    clinicalTerm: term,
    eventDate: eventDate,
    before: {
      balance: beforeBalance,
      missedHours: beforeMissed
    },
    created: {
      attendanceEventId: created.attendanceEventId,
      balance: createBalance,
      missedHours: createMissed
    },
    amended: {
      attendanceEventId: amended.attendanceEventId,
      supersedesEventId: amended.supersedesEventId
    },
    final: {
      balance: finalBalance,
      missedHours: finalMissed
    },
    authoritativeStateRestored: restored,
    historicalQaRowsPreserved: true
  };

  console.log(JSON.stringify(result, null, 2));
  if (!restored) {
    throw new Error('Attendance end-to-end QA did not restore TEST-001 state.');
  }
  return result;
}


function testAttendanceCohortReflectsStudentSummary() {
  const cohort = 'Class of 2027';
  const term = 'Clinical 3';
  const cohortSummary = attendance_getCohortTermSummary(cohort, term);

  const mismatches = [];
  cohortSummary.students.forEach(row => {
    const direct = attendance_getStudentTermSummary(row.studentId, term);
    const a = Number(row.balances.currentAvailableBalanceHours || 0);
    const b = Number(direct.balances.currentAvailableBalanceHours || 0);
    if (Math.abs(a-b) > 0.001) {
      mismatches.push({
        studentId: row.studentId,
        cohortBalance: a,
        directBalance: b
      });
    }
  });

  const result = {
    success: mismatches.length === 0,
    cohort: cohort,
    clinicalTerm: term,
    studentCount: cohortSummary.students.length,
    mismatches: mismatches
  };

  console.log(JSON.stringify(result, null, 2));
  if (!result.success) throw new Error('Cohort/direct attendance mismatch found.');
  return result;
}


function attendanceQa_expectFailure_(fn) {
  try {
    fn();
    return false;
  } catch (error) {
    return true;
  }
}
function testDashboardAttendanceExactCall() {
  var result = dashboardGetCohortAttendanceSummary(
    'Class of 2027',
    'Clinical 3',
    new Date()
  );

  Logger.log(JSON.stringify({
    resultExists: !!result,
    keys: result ? Object.keys(result) : [],
    cohort: result && result.cohort,
    clinicalTerm: result && result.clinicalTerm,
    studentCount: result && result.students
      ? result.students.length
      : 'NO STUDENTS PROPERTY',
    firstStudent: result && result.students && result.students.length
      ? result.students[0]
      : null
  }, null, 2));
}
function testDashboardAttendanceSerialization() {
  var result = dashboardGetCohortAttendanceSummary(
    'Class of 2027',
    'Clinical 3',
    new Date()
  );

  var findings = [];

  function inspect(value, path) {
    if (value instanceof Date) {
      findings.push({
        path: path,
        type: 'Date',
        value: value.toISOString()
      });
      return;
    }

    if (Array.isArray(value)) {
      value.forEach(function(item, index) {
        inspect(item, path + '[' + index + ']');
      });
      return;
    }

    if (value && typeof value === 'object') {
      Object.keys(value).forEach(function(key) {
        inspect(value[key], path + '.' + key);
      });
    }
  }

  inspect(result, 'result');

  Logger.log(JSON.stringify({
    studentCount: result && result.students
      ? result.students.length
      : null,
    dateObjectCount: findings.length,
    dateObjects: findings
  }, null, 2));
}