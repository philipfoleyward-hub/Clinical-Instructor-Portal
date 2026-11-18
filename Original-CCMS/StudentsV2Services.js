/**
 * =====================================================================
 * CCMS STUDENTS V2 PRESENTATION SERVICE
 * StudentsV2Service.gs
 * =====================================================================
 *
 * CANDIDATE / PRESENTATION ADAPTER ONLY
 *
 * PURPOSE
 * -------
 * Supplies the approved Students V2 workspace with presentation-ready data
 * while preserving existing CCMS authorities.
 *
 * AUTHORITIES PRESERVED
 * ---------------------
 * - getStudentDashboard() / Dashboard.gs: central student aggregation
 * - RequirementsEngine: semester/program academic rules
 * - GradeEngine: grade calculations
 * - AttendanceService: attendance/ETO calculations and event authority
 * - SchedulingService: dated site/room-rotation assignment
 * - DashboardDocumentService: student Drive repository/documents
 *
 * THIS FILE DOES NOT
 * ------------------
 * - write academic data
 * - calculate grades
 * - change requirement rules
 * - create a second student record system
 * - infer a CI assignment before CI Coverage is connected
 * - fabricate document links
 */


/**
 * One-call read model for the approved Students V2 workspace.
 *
 * getStudentDashboard() remains the central student aggregation endpoint.
 * This adapter only adds V2-specific presentation reads that are not currently
 * exposed by that DTO.
 */
function dashboardGetStudentV2Workspace(studentId, asOfDate) {
  const id = studentsV2_text_(studentId);

  if (!id) {
    throw new Error('Student ID is required.');
  }

  const date = asOfDate || new Date();

  const dashboard = getStudentDashboard(id);
  const identity = studentsV2_getStudentIdentity_(id);
  const assignment = studentsV2_getTodayAssignment_(id, date);
  const attendance = studentsV2_getAttendance_(id, dashboard, date);
  const programOverall = studentsV2_buildProgramOverall_(dashboard.program);
  const recentActivity = studentsV2_buildRecentActivity_(
    id,
    dashboard,
    attendance
  );

  return {
    success: true,
    studentId: id,
    asOfDate: studentsV2_displayDate_(date),
    identity: identity,
    dashboard: dashboard,
    todayAssignment: assignment,
    attendance: attendance,
    programOverall: programOverall,
    recentActivity: recentActivity,

    // Approved target dependency. Do not infer a person until the
    // date/site CI Coverage calendar is authoritative and connected.
    clinicalInstructor: {
      connected: false,
      status: 'Not Connected',
      name: '',
      note: 'CI assignment will resolve from the dated CI Coverage calendar when that service is connected.'
    }
  };
}


/**
 * Public endpoint requested for Students V2.
 *
 * Returns a presentation timeline only. It does not create a new ledger.
 * Existing authoritative sources remain unchanged.
 */
function dashboardGetStudentRecentActivity(studentId) {
  const id = studentsV2_text_(studentId);

  if (!id) {
    throw new Error('Student ID is required.');
  }

  const dashboard = getStudentDashboard(id);
  const attendance = studentsV2_getAttendance_(
    id,
    dashboard,
    new Date()
  );

  return {
    success: true,
    studentId: id,
    activities: studentsV2_buildRecentActivity_(
      id,
      dashboard,
      attendance
    )
  };
}


/**
 * Read-only dated scheduling adapter.
 */
function dashboardGetStudentTodayAssignment(studentId, asOfDate) {
  const id = studentsV2_text_(studentId);

  if (!id) {
    throw new Error('Student ID is required.');
  }

  return studentsV2_getTodayAssignment_(
    id,
    asOfDate || new Date()
  );
}


/**
 * Student identity fields needed by the V2 presentation but intentionally
 * absent from the legacy Dashboard student DTO.
 *
 * DEPT Number is read independently. Student ID is never parsed to recreate it.
 */
function studentsV2_getStudentIdentity_(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rows = appSheetToObjects_(
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENTS
    )
  ).filter(function(row) {
    return appNormalize_(row['Student ID']) ===
      appNormalize_(studentId);
  });

  if (!rows.length) {
    throw new Error('Student not found: ' + studentId);
  }

  const row =
    typeof dashGetMostAdvancedStudentRow_ === 'function'
      ? dashGetMostAdvancedStudentRow_(rows)
      : rows[0];

  return {
    studentId: studentsV2_text_(row['Student ID']),
    student: studentsV2_text_(row['Student']),
    deptNumber: studentsV2_text_(
      row['DEPT Number'] ||
      row['Dept Number'] ||
      row['Department Number']
    ),
    cohort: studentsV2_text_(row['Cohort']),
    clinicalTerm: studentsV2_text_(row['Clinical Term']),
    active: studentsV2_true_(row['Active']),
    email: studentsV2_text_(row['Email'])
  };
}


function studentsV2_getTodayAssignment_(studentId, date) {
  if (typeof scheduling_getStudentScheduleForDate !== 'function') {
    return {
      available: false,
      scheduled: false,
      status: 'Scheduling Service Unavailable',
      date: studentsV2_displayDate_(date),
      clinicalTerm: '',
      siteId: '',
      site: '',
      rotationArea: '',
      startTime: '',
      endTime: ''
    };
  }

  const result =
    scheduling_getStudentScheduleForDate(
      studentId,
      date
    ) || {};

  const calendar = result.calendar || {};

  return {
    available: true,
    scheduled: Boolean(result.scheduled),
    status: studentsV2_text_(result.status),
    date: studentsV2_displayDate_(result.date || date),
    clinicalTerm: studentsV2_text_(result.clinicalTerm),
    siteId: studentsV2_text_(result.siteId),
    site: studentsV2_text_(result.site),
    rotationArea: studentsV2_text_(result.rotationArea),
    startTime: studentsV2_time_(calendar['Start Time']),
    endTime: studentsV2_time_(calendar['End Time'])
  };
}


function studentsV2_getAttendance_(studentId, dashboard, date) {
  const term =
    dashboard &&
    dashboard.student
      ? studentsV2_text_(dashboard.student.clinicalTerm)
      : '';

  if (!term) {
    return {
      available: false,
      clinicalTerm: '',
      balances: {},
      makeUp: {},
      attendance: {},
      eventCounts: {},
      approvedEvents: [],
      pendingEvents: [],
      rejectedEvents: [],
      warnings: ['Clinical Term is unavailable.']
    };
  }

  try {
    if (typeof attendance_getStudentTermSummary === 'function') {
      return attendance_getStudentTermSummary(
        studentId,
        term,
        date
      );
    }

    if (typeof dashboardGetStudentAttendanceSummary === 'function') {
      return dashboardGetStudentAttendanceSummary(
        studentId,
        term,
        date
      );
    }
  } catch (error) {
    return {
      available: false,
      clinicalTerm: term,
      balances: {},
      makeUp: {},
      attendance: {},
      eventCounts: {},
      approvedEvents: [],
      pendingEvents: [],
      rejectedEvents: [],
      warnings: [error.message]
    };
  }

  return {
    available: false,
    clinicalTerm: term,
    balances: {},
    makeUp: {},
    attendance: {},
    eventCounts: {},
    approvedEvents: [],
    pendingEvents: [],
    rejectedEvents: [],
    warnings: ['AttendanceService is unavailable.']
  };
}


/**
 * Narrow overall Program Progress for the Students V2 summary.
 *
 * IMPORTANT:
 * Fluoroscopy and Head are subgroup summaries within Electives and therefore
 * are NOT added again here. Only the top-level Mandatory + Elective summaries
 * are combined.
 */
function studentsV2_buildProgramOverall_(program) {
  program = program || {};

  const mandatory = program.mandatory || {};
  const elective = program.elective || {};

  const required =
    studentsV2_number_(mandatory.required) +
    studentsV2_number_(elective.required);

  const completed =
    studentsV2_number_(mandatory.completed) +
    studentsV2_number_(elective.completed);

  const remaining =
    Math.max(required - completed, 0);

  return {
    available: required > 0,
    required: required,
    completed: completed,
    remaining: remaining,
    percent: required > 0
      ? Math.max(
          0,
          Math.min(
            100,
            (completed / required) * 100
          )
        )
      : 0,
    note:
      'Overall progress combines the top-level Mandatory and Elective program summaries only; subgroup summaries are not double-counted.'
  };
}


/**
 * Build a newest-first presentation timeline from existing authorities.
 *
 * Sources:
 *   - Dashboard evaluation history (Clinical Tracking)
 *   - Dashboard academic records (Grade Records through Clinical History)
 *   - AttendanceService current events
 *
 * No rows are written and no source-of-truth is replaced.
 */
function studentsV2_buildRecentActivity_(
  studentId,
  dashboard,
  attendance
) {
  const items = [];
  const seen = {};

  const evaluations =
    dashboard && Array.isArray(dashboard.evaluations)
      ? dashboard.evaluations
      : [];

  evaluations.forEach(function(item) {
    const id =
      'evaluation:' +
      studentsV2_text_(item.submissionId);

    if (seen[id]) return;
    seen[id] = true;

    const sourceDocument =
      studentsV2_text_(item.sourceDocument);

    items.push({
      id: id,
      source: 'Clinical Tracking',
      sourceId: studentsV2_text_(item.submissionId),
      kind: studentsV2_activityKind_(
        item.evaluationType
      ),
      title: studentsV2_evaluationTitle_(
        item.evaluationType
      ),
      detail: studentsV2_join_([
        item.examination,
        item.anatomicalArea,
        item.result ||
          item.gradeDisplay
      ], ' • '),
      date: studentsV2_text_(item.date),
      dateValue: studentsV2_number_(item.dateValue),
      clinicalTerm: studentsV2_text_(
        item.attributedClinicalTerm ||
        item.clinicalTerm
      ),
      site: studentsV2_text_(item.clinicalSite),
      status: studentsV2_text_(
        item.coordinatorReviewStatus
      ),
      action: sourceDocument
        ? {
            type: 'url',
            label: 'View Record',
            url: studentsV2_documentUrl_(sourceDocument)
          }
        : {
            type: 'clinicalRecord',
            label: 'View Record',
            url: ''
          }
    });
  });


  const history =
    dashboard && Array.isArray(dashboard.clinicalHistory)
      ? dashboard.clinicalHistory
      : [];

  history.forEach(function(term) {
    (term.academicRecords || []).forEach(function(item) {
      const id =
        'grade:' +
        studentsV2_text_(item.gradeRecordId);

      if (seen[id]) return;
      seen[id] = true;

      items.push({
        id: id,
        source: 'Grade Records',
        sourceId: studentsV2_text_(item.gradeRecordId),
        kind: studentsV2_activityKind_(item.type),
        title: studentsV2_gradeTitle_(item.type),
        detail: studentsV2_join_([
          item.title,
          item.scoreDisplay
        ], ' • '),
        date: studentsV2_text_(item.date),
        dateValue: studentsV2_number_(item.dateValue),
        clinicalTerm: studentsV2_text_(
          item.clinicalTerm ||
          term.clinicalTerm
        ),
        site: studentsV2_text_(item.clinicalSite),
        status: '',
        action: {
          type: 'clinicalRecord',
          label: 'View Grade',
          url: ''
        }
      });
    });
  });


  studentsV2_attendanceEvents_(attendance)
    .forEach(function(item) {
      const eventId =
        studentsV2_text_(
          item.attendanceEventId
        );

      const id =
        'attendance:' +
        eventId;

      if (seen[id]) return;
      seen[id] = true;

      const approval =
        studentsV2_text_(
          item.approvalStatus
        );

      items.push({
        id: id,
        source: 'Attendance Events',
        sourceId: eventId,
        kind: 'attendance',
        title: 'Attendance Event',
        detail: studentsV2_join_([
          studentsV2_pretty_(item.eventType),
          item.hours !== null &&
          item.hours !== undefined &&
          item.hours !== ''
            ? studentsV2_trimNumber_(item.hours) +
              ' hours'
            : '',
          approval &&
          appNormalize_(approval) !== 'approved'
            ? studentsV2_pretty_(approval)
            : ''
        ], ' • '),
        date: studentsV2_displayDate_(
          item.date
        ),
        dateValue: studentsV2_dateValue_(
          item.date
        ),
        clinicalTerm: studentsV2_text_(
          item.clinicalTerm
        ),
        site: '',
        status: approval,
        action: {
          type: 'attendance',
          label: 'View Event',
          url: ''
        }
      });
    });


  return items
    .filter(function(item) {
      return item.dateValue > 0;
    })
    .sort(function(a, b) {
      return b.dateValue - a.dateValue;
    })
    .slice(0, 12);
}


function studentsV2_attendanceEvents_(attendance) {
  attendance = attendance || {};

  return []
    .concat(attendance.approvedEvents || [])
    .concat(attendance.pendingEvents || [])
    .concat(attendance.rejectedEvents || []);
}


function studentsV2_evaluationTitle_(value) {
  const n = appNormalize_(value);

  if (n === 'competency') {
    return 'Competency Evaluation';
  }

  if (n === 'professional competency') {
    return 'Professional Competency';
  }

  if (n === 'iso') {
    return 'ISO Evaluation';
  }

  return studentsV2_pretty_(value) ||
    'Clinical Evaluation';
}


function studentsV2_gradeTitle_(value) {
  const n = appNormalize_(value);

  if (n.indexOf('quiz') !== -1) {
    return 'Clinical Quiz';
  }

  if (n.indexOf('test') !== -1) {
    return 'Clinical Test';
  }

  if (n.indexOf('case') !== -1) {
    return 'Case Presentation';
  }

  if (n.indexOf('professional') !== -1) {
    return 'Professional Competency';
  }

  if (
    n.indexOf('reduction') !== -1 ||
    n.indexOf('deduction') !== -1
  ) {
    return 'Grade Points Reduction';
  }

  return studentsV2_pretty_(value) ||
    'Clinical Grade Record';
}


function studentsV2_activityKind_(value) {
  const n = appNormalize_(value);

  if (n.indexOf('attendance') !== -1) {
    return 'attendance';
  }

  if (
    n.indexOf('quiz') !== -1 ||
    n.indexOf('test') !== -1 ||
    n.indexOf('grade') !== -1
  ) {
    return 'grade';
  }

  if (n.indexOf('professional') !== -1) {
    return 'professional';
  }

  return 'evaluation';
}


function studentsV2_documentUrl_(value) {
  const text = studentsV2_text_(value);

  if (!text) return '';

  if (/^https?:\/\//i.test(text)) {
    return text;
  }

  // A raw Drive file ID is converted to a normal Drive view URL.
  if (/^[A-Za-z0-9_-]{20,}$/.test(text)) {
    return 'https://drive.google.com/open?id=' +
      encodeURIComponent(text);
  }

  // Unknown legacy text is not turned into a fabricated URL.
  return '';
}


function studentsV2_join_(values, separator) {
  return (values || [])
    .map(studentsV2_text_)
    .filter(Boolean)
    .join(separator || ' • ');
}


function studentsV2_pretty_(value) {
  return studentsV2_text_(value)
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\b\w/g, function(ch) {
      return ch.toUpperCase();
    });
}


function studentsV2_text_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}


function studentsV2_number_(value) {
  const number = Number(value);
  return isFinite(number) ? number : 0;
}


function studentsV2_true_(value) {
  if (typeof appIsTrue_ === 'function') {
    return appIsTrue_(value);
  }

  const n = studentsV2_text_(value).toLowerCase();
  return n === 'true' || n === 'yes' || n === '1';
}


function studentsV2_dateValue_(value) {
  if (!value) return 0;

  if (value instanceof Date) {
    return value.getTime();
  }

  const parsed = new Date(value);
  return isNaN(parsed.getTime())
    ? 0
    : parsed.getTime();
}


function studentsV2_displayDate_(value) {
  if (!value) return '';

  if (
    typeof dashDisplayDate_ === 'function'
  ) {
    return dashDisplayDate_(value);
  }

  const date =
    value instanceof Date
      ? value
      : new Date(value);

  if (isNaN(date.getTime())) {
    return studentsV2_text_(value);
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'MMM d, yyyy'
  );
}


function studentsV2_time_(value) {
  if (!value) return '';

  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'h:mm a'
    );
  }

  return studentsV2_text_(value);
}


function studentsV2_trimNumber_(value) {
  const number = Number(value);

  if (!isFinite(number)) {
    return studentsV2_text_(value);
  }

  return String(
    Math.round(number * 100) / 100
  );
}
