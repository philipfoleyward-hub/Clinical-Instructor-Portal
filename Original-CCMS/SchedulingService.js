/**
 * =====================================================================
 * CCMS - SCHEDULING SERVICE - PHASE 1 COMPLETE
 * SchedulingService.gs
 * =====================================================================
 * Scheduling foundation + Room Rotations v1 draft commands.
 *
 * Authorities:
 * Semester Rules -> term windows
 * Clinical Calendar -> scheduled dates
 * Student Site Assignments -> placement history
 * Student Room Rotations -> physical room/area
 * Students -> identity/cohort
 * Clinical Sites -> site identity
 *
 * Writes only Student Room Rotations with Status = Draft through the
 * explicit save command below. It does NOT infer missing assignments,
 * calculate attendance/ETO, assign CIs, publish schedules, or modify
 * academic rules/history.
 *
 * Source data is preserved as entered. Cohort normalization is used only
 * for relational comparison.
 * =====================================================================
 */


/**
 * Return all active Clinical Calendar rows for a Clinical Term.
 */
function scheduling_getClinicalCalendarForTerm(clinicalTerm) {
  if (!clinicalTerm) {
    throw new Error('Clinical Term is required.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rows = appSheetToObjects_(
    appGetSheet_(ss, APP_CONFIG.SHEETS.CLINICAL_CALENDAR)
  );

  return rows
    .filter(row =>
      appNormalize_(row['Clinical Term']) ===
        appNormalize_(clinicalTerm) &&
      scheduling_isActiveRow_(row)
    )
    .sort((a, b) =>
      appDateValue_(a['Date']) -
      appDateValue_(b['Date'])
    );
}


/**
 * Resolve the single active site assignment for a student/date.
 *
 * More than one active assignment is a scheduling conflict and fails
 * loudly rather than guessing.
 */
function scheduling_getStudentSiteAssignment(studentId, date) {
  scheduling_requireStudentIdAndDate_(studentId, date);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rows = appSheetToObjects_(
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENT_SITE_ASSIGNMENTS
    )
  );

  const matches = rows.filter(row =>
    appNormalize_(row['Student ID']) ===
      appNormalize_(studentId) &&
    scheduling_isActiveAssignment_(row) &&
    appDateWithin_(
      date,
      row['Start Date'],
      row['End Date']
    )
  );

  if (matches.length > 1) {
    throw new Error(
      'Scheduling conflict: Student "' +
      studentId +
      '" has more than one active site assignment covering ' +
      scheduling_formatDate_(date) +
      ': ' +
      matches
        .map(row =>
          String(
            row['Assignment ID'] || ''
          ).trim() ||
          '[missing Assignment ID]'
        )
        .join(', ')
    );
  }

  return matches.length
    ? matches[0]
    : null;
}


/**
 * Resolve the single room/area rotation for a student/date.
 *
 * When assignmentId is supplied, the room rotation must belong to that
 * parent site assignment.
 */
function scheduling_getStudentRoomRotation(
  studentId,
  date,
  assignmentId
) {
  scheduling_requireStudentIdAndDate_(
    studentId,
    date
  );

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rows = appSheetToObjects_(
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENT_ROOM_ROTATIONS
    )
  );

  const matches = rows.filter(row => {
    if (
      !scheduling_isAuthoritativeRotation_(row)
    ) {
      return false;
    }

    if (
      appNormalize_(row['Student ID']) !==
      appNormalize_(studentId)
    ) {
      return false;
    }

    if (
      assignmentId &&
      appNormalize_(row['Assignment ID']) !==
      appNormalize_(assignmentId)
    ) {
      return false;
    }

    return appDateWithin_(
      date,
      row['Start Date'],
      row['End Date']
    );
  });

  if (matches.length > 1) {
    throw new Error(
      'Scheduling conflict: Student "' +
      studentId +
      '" has more than one room rotation covering ' +
      scheduling_formatDate_(date) +
      ': ' +
      matches
        .map(row =>
          String(
            row['Rotation ID'] || ''
          ).trim() ||
          '[missing Rotation ID]'
        )
        .join(', ')
    );
  }

  return matches.length
    ? matches[0]
    : null;
}


/**
 * Resolve the expected schedule for one student on one date.
 *
 * Explicit non-error states:
 *   Not Scheduled
 *   Site Not Assigned
 *   Rotation Not Assigned
 *
 * Contradictory relational data fails loudly.
 */
function scheduling_getStudentScheduleForDate(
  studentId,
  date
) {
  scheduling_requireStudentIdAndDate_(
    studentId,
    date
  );

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const targetDate =
    appDateOnly_(date);

  const student =
    scheduling_getStudent_(
      ss,
      studentId
    );

  if (!student) {
    throw new Error(
      'Student not found: ' +
      studentId
    );
  }

  const calendarRows =
    scheduling_getCalendarRowsForDate_(
      ss,
      targetDate,
      student.cohort
    );

  if (calendarRows.length === 0) {
    return {
      studentId: student.studentId,
      student: student.student,
      cohort: student.cohort,
      date: targetDate,
      scheduled: false,
      status: 'Not Scheduled',
      clinicalTerm: '',
      calendar: null,
      assignment: null,
      rotation: null,
      siteId: '',
      site: '',
      rotationArea: ''
    };
  }

  if (calendarRows.length > 1) {
    throw new Error(
      'Clinical Calendar conflict: more than one active row matches ' +
      scheduling_formatDate_(targetDate) +
      ' for cohort "' +
      student.cohort +
      '". Calendar IDs: ' +
      calendarRows
        .map(row =>
          String(
            row['Calendar ID'] || ''
          ).trim() ||
          '[missing Calendar ID]'
        )
        .join(', ')
    );
  }

  const calendar =
    calendarRows[0];

  const calendarTerm =
    String(
      calendar['Clinical Term'] || ''
    ).trim();

  const assignment =
    scheduling_getStudentSiteAssignment(
      student.studentId,
      targetDate
    );

  if (!assignment) {
    return {
      studentId: student.studentId,
      student: student.student,
      cohort: student.cohort,
      date: targetDate,
      scheduled: true,
      status: 'Site Not Assigned',
      clinicalTerm: calendarTerm,
      calendar: calendar,
      assignment: null,
      rotation: null,
      siteId: '',
      site: '',
      rotationArea: ''
    };
  }

  const assignmentTerm =
    String(
      assignment['Clinical Term'] || ''
    ).trim();

  if (
    appNormalize_(assignmentTerm) !==
    appNormalize_(calendarTerm)
  ) {
    throw new Error(
      'Scheduling data integrity conflict for Student "' +
      student.studentId +
      '" on ' +
      scheduling_formatDate_(targetDate) +
      ': Clinical Calendar term "' +
      calendarTerm +
      '" does not match site assignment "' +
      String(
        assignment['Assignment ID'] || ''
      ).trim() +
      '" term "' +
      assignmentTerm +
      '".'
    );
  }

  const siteId =
    String(
      assignment['Site ID'] || ''
    ).trim();

  const site =
    scheduling_getSite_(
      ss,
      siteId
    );

  if (!site) {
    throw new Error(
      'Scheduling data integrity conflict for Student "' +
      student.studentId +
      '" on ' +
      scheduling_formatDate_(targetDate) +
      ': site assignment "' +
      String(
        assignment['Assignment ID'] || ''
      ).trim() +
      '" references unknown Clinical Site ID "' +
      siteId +
      '".'
    );
  }

  const rotation =
    scheduling_getStudentRoomRotation(
      student.studentId,
      targetDate,
      assignment['Assignment ID']
    );

  return {
    studentId: student.studentId,
    student: student.student,
    cohort: student.cohort,
    date: targetDate,
    scheduled: true,
    status:
      rotation
        ? 'Scheduled'
        : 'Rotation Not Assigned',
    clinicalTerm: calendarTerm,
    calendar: calendar,
    assignment: assignment,
    rotation: rotation,
    siteId: siteId,
    site: site.name,
    rotationArea:
      rotation
        ? String(
            rotation['Rotation Area'] || ''
          ).trim()
        : ''
  };
}


/**
 * Resolve all active students in a cohort for a specific date.
 *
 * Cohort representations such as:
 *   2027
 *   "2027"
 *   "Class of 2027"
 *
 * are treated as relationally equivalent without rewriting source data.
 */
function scheduling_getCohortScheduleForDate(
  cohort,
  date
) {
  if (!cohort) {
    throw new Error(
      'Cohort is required.'
    );
  }

  const targetDate =
    appDateOnly_(date);

  if (!targetDate) {
    throw new Error(
      'A valid date is required.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const targetCohort =
    scheduling_normalizeCohort_(
      cohort
    );

  return scheduling_getUniqueActiveStudents_(ss)
    .filter(student =>
      scheduling_normalizeCohort_(
        student.cohort
      ) === targetCohort
    )
    .map(student =>
      scheduling_getStudentScheduleForDate(
        student.studentId,
        targetDate
      )
    )
    .sort((a, b) =>
      a.student.localeCompare(
        b.student
      )
    );
}



/**
 * Return editable Room Rotation workspace rows for a cohort/Clinical Term.
 *
 * Each parent site assignment is returned once with the currently
 * authoritative rotation (legacy blank Status or Published), any Draft
 * rotation for the same assignment, and known historical/current rotation
 * areas for that Site ID.
 *
 * This function is read-only.
 */
function scheduling_getRoomRotationWorkspace(
  cohort,
  clinicalTerm
) {
  if (!cohort) {
    throw new Error('Cohort is required.');
  }

  if (!clinicalTerm) {
    throw new Error('Clinical Term is required.');
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const targetCohort =
    scheduling_normalizeCohort_(
      cohort
    );

  const targetTerm =
    appNormalize_(
      clinicalTerm
    );

  const activeStudentIds =
    new Set(
      scheduling_getUniqueActiveStudents_(ss)
        .filter(student =>
          scheduling_normalizeCohort_(
            student.cohort
          ) === targetCohort
        )
        .map(student =>
          appNormalize_(
            student.studentId
          )
        )
    );

  const assignments =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENT_SITE_ASSIGNMENTS
      )
    )
      .filter(row =>
        activeStudentIds.has(
          appNormalize_(
            row['Student ID']
          )
        ) &&
        appNormalize_(
          row['Clinical Term']
        ) === targetTerm &&
        scheduling_isActiveAssignment_(row)
      );

  const rotations =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENT_ROOM_ROTATIONS
      )
    );

  const studentById = {};

  scheduling_getUniqueActiveStudents_(ss)
    .forEach(student => {
      studentById[
        appNormalize_(student.studentId)
      ] = student;
    });

  return assignments
    .map(assignment => {
      const assignmentId =
        String(
          assignment['Assignment ID'] || ''
        ).trim();

      const assignmentKey =
        appNormalize_(
          assignmentId
        );

      const related =
        rotations.filter(rotation =>
          appNormalize_(
            rotation['Assignment ID']
          ) === assignmentKey
        );

      const authoritative =
        related.filter(
          scheduling_isAuthoritativeRotation_
        );

      if (authoritative.length > 1) {
        throw new Error(
          'Scheduling conflict: Assignment "' +
          assignmentId +
          '" has more than one authoritative room rotation.'
        );
      }

      const drafts =
        related.filter(row =>
          scheduling_rotationStatus_(row) ===
          'draft'
        );

      if (drafts.length > 1) {
        throw new Error(
          'Scheduling conflict: Assignment "' +
          assignmentId +
          '" has more than one Draft room rotation.'
        );
      }

      const student =
        studentById[
          appNormalize_(
            assignment['Student ID']
          )
        ] || null;

      const site =
        scheduling_getSite_(
          ss,
          assignment['Site ID']
        );

      return {
        assignmentId: assignmentId,
        studentId:
          String(
            assignment['Student ID'] || ''
          ).trim(),
        student:
          student
            ? student.student
            : '',
        cohort:
          student
            ? student.cohort
            : '',
        clinicalTerm:
          String(
            assignment['Clinical Term'] || ''
          ).trim(),
        siteId:
          String(
            assignment['Site ID'] || ''
          ).trim(),
        site:
          site
            ? site.name
            : '',
        startDate:
          appDateOnly_(
            assignment['Start Date']
          ),
        endDate:
          appDateOnly_(
            assignment['End Date']
          ),
        assignmentType:
          String(
            assignment['Assignment Type'] || ''
          ).trim(),
        officialRotation:
          authoritative.length
            ? authoritative[0]
            : null,
        draftRotation:
          drafts.length
            ? drafts[0]
            : null,
        rotationAreas:
          scheduling_getKnownRotationAreasForSite_(
            rotations,
            assignment['Site ID']
          )
      };
    })
    .sort((a, b) => {
      const byDate =
        appDateValue_(a.startDate) -
        appDateValue_(b.startDate);

      if (byDate !== 0) {
        return byDate;
      }

      return String(a.student || '')
        .localeCompare(
          String(b.student || '')
        );
    });
}


/**
 * Validate one proposed Room Rotation Draft.
 *
 * This is read-only. The returned normalized record is suitable for
 * scheduling_saveRoomRotationDraft().
 */
function scheduling_validateRoomRotationDraft(input) {
  input = input || {};

  const assignmentId =
    String(
      input.assignmentId || ''
    ).trim();

  const rotationArea =
    String(
      input.rotationArea || ''
    ).trim();

  if (!assignmentId) {
    throw new Error(
      'Assignment ID is required.'
    );
  }

  if (!rotationArea) {
    throw new Error(
      'Rotation Area is required.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const assignments =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENT_SITE_ASSIGNMENTS
      )
    );

  const matches =
    assignments.filter(row =>
      appNormalize_(
        row['Assignment ID']
      ) ===
      appNormalize_(
        assignmentId
      )
    );

  if (matches.length !== 1) {
    throw new Error(
      'Assignment ID "' +
      assignmentId +
      '" must resolve to exactly one Student Site Assignment.'
    );
  }

  const parent =
    matches[0];

  if (
    !scheduling_isActiveAssignment_(
      parent
    )
  ) {
    throw new Error(
      'Assignment "' +
      assignmentId +
      '" is not active.'
    );
  }

  const studentId =
    String(
      parent['Student ID'] || ''
    ).trim();

  const siteId =
    String(
      parent['Site ID'] || ''
    ).trim();

  const startDate =
    appDateOnly_(
      input.startDate ||
      parent['Start Date']
    );

  const endDate =
    appDateOnly_(
      input.endDate ||
      parent['End Date']
    );

  if (!startDate || !endDate) {
    throw new Error(
      'A valid Start Date and End Date are required.'
    );
  }

  if (
    endDate.getTime() <
    startDate.getTime()
  ) {
    throw new Error(
      'End Date cannot precede Start Date.'
    );
  }

  if (
    !appDateWithin_(
      startDate,
      parent['Start Date'],
      parent['End Date']
    ) ||
    !appDateWithin_(
      endDate,
      parent['Start Date'],
      parent['End Date']
    )
  ) {
    throw new Error(
      'Room Rotation dates must remain inside parent Assignment "' +
      assignmentId +
      '".'
    );
  }

  if (
    input.studentId &&
    appNormalize_(input.studentId) !==
    appNormalize_(studentId)
  ) {
    throw new Error(
      'Student ID does not match parent Assignment "' +
      assignmentId +
      '".'
    );
  }

  if (
    input.siteId &&
    appNormalize_(input.siteId) !==
    appNormalize_(siteId)
  ) {
    throw new Error(
      'Site ID does not match parent Assignment "' +
      assignmentId +
      '".'
    );
  }

  return {
    rotationId:
      String(
        input.rotationId || ''
      ).trim(),
    assignmentId: assignmentId,
    studentId: studentId,
    siteId: siteId,
    startDate: startDate,
    endDate: endDate,
    rotationArea: rotationArea,
    priorityNote:
      String(
        input.priorityNote || ''
      ).trim(),
    notes:
      String(
        input.notes || ''
      ).trim(),
    status: 'Draft'
  };
}


/**
 * Create or update a Room Rotation Draft.
 *
 * Drafts never replace the authoritative legacy/Published rotation.
 * If rotationId is omitted and a Draft already exists for the parent
 * Assignment, that Draft is updated rather than creating a duplicate.
 */
function scheduling_saveRoomRotationDraft(input) {
  const record =
    scheduling_validateRoomRotationDraft(
      input
    );

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENT_ROOM_ROTATIONS
    );

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0]
      .map(value =>
        String(
          value || ''
        ).trim()
      );

  const map =
    appHeaderMap_(
      headers
    );

  [
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
  ].forEach(header => {
    if (
      map[
        appNormalize_(header)
      ] === undefined
    ) {
      throw new Error(
        'Student Room Rotations is missing required header "' +
        header +
        '".'
      );
    }
  });

  const rows =
    appSheetToObjects_(sheet);

  const assignmentKey =
    appNormalize_(
      record.assignmentId
    );

  const requestedId =
    appNormalize_(
      record.rotationId
    );

  const draftMatches =
    rows
      .map((row, index) => ({
        row: row,
        sheetRow: index + 2
      }))
      .filter(item =>
        scheduling_rotationStatus_(
          item.row
        ) === 'draft' &&
        appNormalize_(
          item.row['Assignment ID']
        ) === assignmentKey
      );

  if (
    draftMatches.length > 1
  ) {
    throw new Error(
      'Scheduling conflict: Assignment "' +
      record.assignmentId +
      '" already has more than one Draft rotation.'
    );
  }

  let target =
    null;

  if (requestedId) {
    target =
      rows
        .map((row, index) => ({
          row: row,
          sheetRow: index + 2
        }))
        .find(item =>
          appNormalize_(
            item.row['Rotation ID']
          ) === requestedId
        ) ||
      null;

    if (
      target &&
      scheduling_rotationStatus_(
        target.row
      ) !== 'draft'
    ) {
      throw new Error(
        'Only Draft rotations may be updated by scheduling_saveRoomRotationDraft().'
      );
    }
  } else if (
    draftMatches.length === 1
  ) {
    target =
      draftMatches[0];
  }

  const rotationId =
    target
      ? String(
          target.row['Rotation ID'] || ''
        ).trim()
      : (
          record.rotationId ||
          scheduling_newRotationId_()
        );

  const values = {
    'Rotation ID': rotationId,
    'Assignment ID': record.assignmentId,
    'Student ID': record.studentId,
    'Site ID': record.siteId,
    'Start Date': record.startDate,
    'End Date': record.endDate,
    'Rotation Area': record.rotationArea,
    'Priority Note': record.priorityNote,
    'Notes': record.notes,
    'Status': 'Draft'
  };

  const rowValues =
    headers.map(header =>
      Object.prototype.hasOwnProperty.call(
        values,
        header
      )
        ? values[header]
        : (
            target
              ? target.row[header]
              : ''
          )
    );

  let sheetRow;

  if (target) {
    sheetRow =
      target.sheetRow;

    sheet
      .getRange(
        sheetRow,
        1,
        1,
        headers.length
      )
      .setValues([
        rowValues
      ]);
  } else {
    sheet.appendRow(
      rowValues
    );

    sheetRow =
      sheet.getLastRow();
  }

  return {
    success: true,
    operation:
      target
        ? 'updated'
        : 'created',
    sheetRow: sheetRow,
    rotationId: rotationId,
    assignmentId:
      record.assignmentId,
    status: 'Draft'
  };
}


/**
 * Return distinct existing Rotation Area values for one Site ID.
 *
 * This intentionally derives the v1 dropdown from existing source data
 * rather than introducing another configuration sheet.
 */
function scheduling_getKnownRotationAreasForSite(
  siteId
) {
  if (!siteId) {
    return [];
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const rotations =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENT_ROOM_ROTATIONS
      )
    );

  return scheduling_getKnownRotationAreasForSite_(
    rotations,
    siteId
  );
}


/**
 * Validate core Scheduling Phase 1 relational integrity.
 */
function scheduling_validateSchedulingData() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const issues = [];
  const warnings = [];

  scheduling_validateSheetHeaders_(
    ss,
    APP_CONFIG.SHEETS.CLINICAL_CALENDAR,
    [
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
    ]
  );

  scheduling_validateSheetHeaders_(
    ss,
    APP_CONFIG.SHEETS.STUDENT_SITE_ASSIGNMENTS,
    [
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
    ]
  );

  scheduling_validateSheetHeaders_(
    ss,
    APP_CONFIG.SHEETS.STUDENT_ROOM_ROTATIONS,
    [
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
    ]
  );

  const students =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENTS
      )
    );

  const sites =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.CLINICAL_SITES
      )
    );

  const calendar =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.CLINICAL_CALENDAR
      )
    );

  const assignments =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENT_SITE_ASSIGNMENTS
      )
    );

  const rotations =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENT_ROOM_ROTATIONS
      )
    );

  const studentIds =
    new Set(
      students
        .map(row =>
          appNormalize_(
            row['Student ID']
          )
        )
        .filter(Boolean)
    );

  const siteIds =
    new Set(
      sites
        .map(row =>
          appNormalize_(
            row['Clinical Site ID']
          )
        )
        .filter(Boolean)
    );

  const assignmentById = {};

  assignments.forEach(
    (row, index) => {
      const n = index + 2;

      const id =
        String(
          row['Assignment ID'] || ''
        ).trim();

      const studentId =
        String(
          row['Student ID'] || ''
        ).trim();

      const siteId =
        String(
          row['Site ID'] || ''
        ).trim();

      const start =
        appDateOnly_(
          row['Start Date']
        );

      const end =
        appDateOnly_(
          row['End Date']
        );

      if (!id) {
        issues.push(
          'Student Site Assignments row ' +
          n +
          ': missing Assignment ID.'
        );
      } else {
        const key =
          appNormalize_(id);

        if (assignmentById[key]) {
          issues.push(
            'Duplicate Assignment ID "' +
            id +
            '".'
          );
        }

        assignmentById[key] = row;
      }

      if (
        !studentId ||
        !studentIds.has(
          appNormalize_(studentId)
        )
      ) {
        issues.push(
          'Student Site Assignments row ' +
          n +
          ': invalid Student ID "' +
          studentId +
          '".'
        );
      }

      if (
        !siteId ||
        !siteIds.has(
          appNormalize_(siteId)
        )
      ) {
        issues.push(
          'Student Site Assignments row ' +
          n +
          ': invalid Site ID "' +
          siteId +
          '".'
        );
      }

      if (!start || !end) {
        issues.push(
          'Student Site Assignments row ' +
          n +
          ': invalid Start Date or End Date.'
        );
      } else if (
        end.getTime() <
        start.getTime()
      ) {
        issues.push(
          'Student Site Assignments row ' +
          n +
          ': End Date precedes Start Date.'
        );
      }
    }
  );

  for (
    let i = 0;
    i < assignments.length;
    i++
  ) {
    if (
      !scheduling_isActiveAssignment_(
        assignments[i]
      )
    ) {
      continue;
    }

    for (
      let j = i + 1;
      j < assignments.length;
      j++
    ) {
      if (
        !scheduling_isActiveAssignment_(
          assignments[j]
        )
      ) {
        continue;
      }

      if (
        appNormalize_(
          assignments[i]['Student ID']
        ) !==
        appNormalize_(
          assignments[j]['Student ID']
        )
      ) {
        continue;
      }

      if (
        scheduling_rangesOverlap_(
          assignments[i]['Start Date'],
          assignments[i]['End Date'],
          assignments[j]['Start Date'],
          assignments[j]['End Date']
        )
      ) {
        issues.push(
          'Overlapping active site assignments for Student ID "' +
          String(
            assignments[i]['Student ID'] || ''
          ).trim() +
          '": "' +
          String(
            assignments[i]['Assignment ID'] || ''
          ).trim() +
          '" and "' +
          String(
            assignments[j]['Assignment ID'] || ''
          ).trim() +
          '".'
        );
      }
    }
  }

  const rotationById = {};
  const allowedRotationStatuses =
    new Set([
      '',
      'draft',
      'published',
      'superseded'
    ]);

  rotations.forEach(
    (row, index) => {
      const n = index + 2;

      const assignmentId =
        String(
          row['Assignment ID'] || ''
        ).trim();

      const parent =
        assignmentById[
          appNormalize_(assignmentId)
        ];

      const rotationId =
        String(
          row['Rotation ID'] || ''
        ).trim();

      if (!rotationId) {
        issues.push(
          'Student Room Rotations row ' +
          n +
          ': missing Rotation ID.'
        );
      } else {
        const rotationKey =
          appNormalize_(
            rotationId
          );

        if (
          rotationById[
            rotationKey
          ]
        ) {
          issues.push(
            'Duplicate Rotation ID "' +
            rotationId +
            '".'
          );
        }

        rotationById[
          rotationKey
        ] = row;
      }

      const rotationStatus =
        scheduling_rotationStatus_(
          row
        );

      if (
        !allowedRotationStatuses.has(
          rotationStatus
        )
      ) {
        issues.push(
          'Student Room Rotations row ' +
          n +
          ': invalid Status "' +
          String(
            row['Status'] || ''
          ).trim() +
          '".'
        );
      }

      if (!parent) {
        issues.push(
          'Student Room Rotations row ' +
          n +
          ': Assignment ID "' +
          assignmentId +
          '" does not exist.'
        );
        return;
      }

      if (
        appNormalize_(
          row['Student ID']
        ) !==
        appNormalize_(
          parent['Student ID']
        )
      ) {
        issues.push(
          'Student Room Rotations row ' +
          n +
          ': Student ID does not match parent assignment.'
        );
      }

      if (
        appNormalize_(
          row['Site ID']
        ) !==
        appNormalize_(
          parent['Site ID']
        )
      ) {
        issues.push(
          'Student Room Rotations row ' +
          n +
          ': Site ID does not match parent assignment.'
        );
      }

      const start =
        appDateOnly_(
          row['Start Date']
        );

      const end =
        appDateOnly_(
          row['End Date']
        );

      if (!start || !end) {
        issues.push(
          'Student Room Rotations row ' +
          n +
          ': invalid Start Date or End Date.'
        );
        return;
      }

      if (
        end.getTime() <
        start.getTime()
      ) {
        issues.push(
          'Student Room Rotations row ' +
          n +
          ': End Date precedes Start Date.'
        );
        return;
      }

      if (
        !appDateWithin_(
          start,
          parent['Start Date'],
          parent['End Date']
        ) ||
        !appDateWithin_(
          end,
          parent['Start Date'],
          parent['End Date']
        )
      ) {
        issues.push(
          'Student Room Rotations row ' +
          n +
          ': rotation dates extend outside parent assignment.'
        );
      }
    }
  );

  for (
    let i = 0;
    i < rotations.length;
    i++
  ) {
    if (
      !scheduling_isAuthoritativeRotation_(
        rotations[i]
      )
    ) {
      continue;
    }

    for (
      let j = i + 1;
      j < rotations.length;
      j++
    ) {
      if (
        !scheduling_isAuthoritativeRotation_(
          rotations[j]
        )
      ) {
        continue;
      }

      if (
        appNormalize_(
          rotations[i]['Student ID']
        ) !==
        appNormalize_(
          rotations[j]['Student ID']
        )
      ) {
        continue;
      }

      if (
        scheduling_rangesOverlap_(
          rotations[i]['Start Date'],
          rotations[i]['End Date'],
          rotations[j]['Start Date'],
          rotations[j]['End Date']
        )
      ) {
        issues.push(
          'Overlapping authoritative room rotations for Student ID "' +
          String(
            rotations[i]['Student ID'] || ''
          ).trim() +
          '": "' +
          String(
            rotations[i]['Rotation ID'] || ''
          ).trim() +
          '" and "' +
          String(
            rotations[j]['Rotation ID'] || ''
          ).trim() +
          '".'
        );
      }
    }
  }

  calendar.forEach(
    (row, index) => {
      const n = index + 2;

      if (
        !String(
          row['Calendar ID'] || ''
        ).trim()
      ) {
        issues.push(
          'Clinical Calendar row ' +
          n +
          ': missing Calendar ID.'
        );
      }

      const date =
        appDateOnly_(
          row['Date']
        );

      if (!date) {
        issues.push(
          'Clinical Calendar row ' +
          n +
          ': invalid Date.'
        );
        return;
      }

      let term;

      try {
        term =
          clinicalTerm_getForDate(
            date
          );
      } catch (error) {
        issues.push(
          'Clinical Calendar row ' +
          n +
          ': Semester Rules lookup failed: ' +
          error.message
        );
        return;
      }

      if (
        term === 'UNATTRIBUTED'
      ) {
        warnings.push(
          'Clinical Calendar row ' +
          n +
          ': date is outside active Semester Rules windows.'
        );
      } else if (
        appNormalize_(term) !==
        appNormalize_(
          row['Clinical Term']
        )
      ) {
        issues.push(
          'Clinical Calendar row ' +
          n +
          ': stated Clinical Term "' +
          row['Clinical Term'] +
          '" disagrees with Semester Rules attribution "' +
          term +
          '".'
        );
      }
    }
  );

  const result = {
    success:
      issues.length === 0,
    issueCount:
      issues.length,
    warningCount:
      warnings.length,
    issues: issues,
    warnings: warnings,
    counts: {
      calendarRows:
        calendar.length,
      siteAssignments:
        assignments.length,
      roomRotations:
        rotations.length
    }
  };

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
 * Internal: return active Clinical Calendar rows matching date/cohort.
 */
function scheduling_getCalendarRowsForDate_(
  ss,
  date,
  cohort
) {
  const targetCohort =
    scheduling_normalizeCohort_(
      cohort
    );

  return appSheetToObjects_(
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.CLINICAL_CALENDAR
    )
  ).filter(row => {
    if (
      !scheduling_isActiveRow_(row)
    ) {
      return false;
    }

    const rowDate =
      appDateOnly_(
        row['Date']
      );

    if (
      !rowDate ||
      rowDate.getTime() !==
      date.getTime()
    ) {
      return false;
    }

    const rowCohort =
      String(
        row['Cohort'] || ''
      ).trim();

    return (
      !rowCohort ||
      scheduling_normalizeCohort_(
        rowCohort
      ) === targetCohort
    );
  });
}


/**
 * Internal: resolve student identity.
 */
function scheduling_getStudent_(
  ss,
  studentId
) {
  const matches =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.STUDENTS
      )
    ).filter(row =>
      appNormalize_(
        row['Student ID']
      ) ===
      appNormalize_(
        studentId
      )
    );

  if (!matches.length) {
    return null;
  }

  const row =
    matches.find(r =>
      appIsTrue_(r['Active'])
    ) ||
    matches[0];

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
      ).trim()
  };
}


/**
 * Internal: return unique active students by stable Student ID.
 */
function scheduling_getUniqueActiveStudents_(ss) {
  const byId = {};

  appSheetToObjects_(
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.STUDENTS
    )
  )
    .filter(row =>
      appIsTrue_(row['Active'])
    )
    .forEach(row => {
      const id =
        String(
          row['Student ID'] || ''
        ).trim();

      if (!id) {
        return;
      }

      const key =
        appNormalize_(id);

      if (!byId[key]) {
        byId[key] = {
          studentId: id,
          student:
            String(
              row['Student'] || ''
            ).trim(),
          cohort:
            String(
              row['Cohort'] || ''
            ).trim()
        };
      }
    });

  return Object.keys(byId)
    .map(key =>
      byId[key]
    );
}


/**
 * Internal: resolve Clinical Site identity by stable Site ID.
 */
function scheduling_getSite_(
  ss,
  siteId
) {
  const target =
    appNormalize_(
      siteId
    );

  if (!target) {
    return null;
  }

  const row =
    appSheetToObjects_(
      appGetSheet_(
        ss,
        APP_CONFIG.SHEETS.CLINICAL_SITES
      )
    ).find(r =>
      appNormalize_(
        r['Clinical Site ID']
      ) === target
    );

  if (!row) {
    return null;
  }

  return {
    id:
      String(
        row['Clinical Site ID'] || ''
      ).trim(),
    name:
      String(
        row['Clinical Site'] ||
        row['Site'] ||
        row['Clinical Site Name'] ||
        ''
      ).trim(),
    row: row
  };
}


/**
 * Internal: determine whether a Clinical Calendar row is active.
 */
function scheduling_isActiveRow_(row) {
  if (
    !Object.prototype.hasOwnProperty.call(
      row,
      'Active'
    )
  ) {
    return true;
  }

  return appIsTrue_(
    row['Active']
  );
}


/**
 * Internal: determine whether a site assignment participates in
 * active schedule resolution.
 */
function scheduling_isActiveAssignment_(row) {
  const status =
    appNormalize_(
      row['Status']
    );

  return ![
    'cancelled',
    'canceled',
    'superseded',
    'inactive'
  ].includes(status);
}



/**
 * Internal: normalize Room Rotation workflow status.
 *
 * Blank remains authoritative for legacy compatibility.
 */
function scheduling_rotationStatus_(row) {
  return appNormalize_(
    row &&
    Object.prototype.hasOwnProperty.call(
      row,
      'Status'
    )
      ? row['Status']
      : ''
  );
}


/**
 * Internal: legacy blank Status and Published are authoritative.
 * Draft and Superseded do not affect the official schedule resolver.
 */
function scheduling_isAuthoritativeRotation_(row) {
  const status =
    scheduling_rotationStatus_(
      row
    );

  return (
    status === '' ||
    status === 'published'
  );
}


/**
 * Internal: derive distinct Rotation Area options for one Site ID.
 */
function scheduling_getKnownRotationAreasForSite_(
  rotations,
  siteId
) {
  const target =
    appNormalize_(
      siteId
    );

  const seen = {};
  const output = [];

  (rotations || [])
    .forEach(row => {
      if (
        appNormalize_(
          row['Site ID']
        ) !== target
      ) {
        return;
      }

      const area =
        String(
          row['Rotation Area'] || ''
        ).trim();

      const key =
        appNormalize_(
          area
        );

      if (
        !area ||
        !key ||
        seen[key]
      ) {
        return;
      }

      seen[key] = true;
      output.push(area);
    });

  return output.sort((a, b) =>
    a.localeCompare(b)
  );
}


/**
 * Internal: generate a stable draft Rotation ID.
 */
function scheduling_newRotationId_() {
  return (
    'ROT-DRAFT-' +
    Utilities
      .getUuid()
      .replace(/-/g, '')
      .slice(0, 12)
      .toUpperCase()
  );
}


/**
 * Internal: determine whether two inclusive date ranges overlap.
 */
function scheduling_rangesOverlap_(
  startA,
  endA,
  startB,
  endB
) {
  const a1 =
    appDateOnly_(startA);

  const a2 =
    appDateOnly_(endA);

  const b1 =
    appDateOnly_(startB);

  const b2 =
    appDateOnly_(endB);

  if (
    !a1 ||
    !a2 ||
    !b1 ||
    !b2
  ) {
    return false;
  }

  return (
    a1.getTime() <=
      b2.getTime() &&
    b1.getTime() <=
      a2.getTime()
  );
}


/**
 * Internal: shared parameter validation.
 */
function scheduling_requireStudentIdAndDate_(
  studentId,
  date
) {
  if (!studentId) {
    throw new Error(
      'Student ID is required.'
    );
  }

  if (!appDateOnly_(date)) {
    throw new Error(
      'A valid date is required.'
    );
  }
}


/**
 * Internal: validate required sheet headers.
 */
function scheduling_validateSheetHeaders_(
  ss,
  sheetName,
  requiredHeaders
) {
  const sheet =
    appGetSheet_(
      ss,
      sheetName
    );

  if (
    sheet.getLastColumn() < 1
  ) {
    throw new Error(
      'Sheet "' +
      sheetName +
      '" has no headers.'
    );
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0];

  const map =
    appHeaderMap_(
      headers
    );

  const missing =
    requiredHeaders.filter(
      h =>
        map[
          appNormalize_(h)
        ] === undefined
    );

  if (missing.length) {
    throw new Error(
      'Sheet "' +
      sheetName +
      '" is missing required header(s): ' +
      missing.join(', ')
    );
  }

  return true;
}


/**
 * Internal: format date for logs/errors.
 */
function scheduling_formatDate_(date) {
  const value =
    appDateOnly_(date);

  if (!value) {
    return '';
  }

  return Utilities.formatDate(
    value,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


/**
 * Normalize cohort representations for relational comparison.
 *
 * Examples:
 *   2027            -> 2027
 *   "2027"          -> 2027
 *   "Class of 2027" -> 2027
 *
 * Comparison only. Source data is never rewritten.
 */
function scheduling_normalizeCohort_(value) {
  const text =
    String(
      value || ''
    )
      .trim()
      .toLowerCase();

  if (!text) {
    return '';
  }

  const match =
    text.match(
      /\b(20\d{2})\b/
    );

  if (match) {
    return match[1];
  }

  return appNormalize_(
    value
  );
}
/**
 * =====================================================================
 * CI SCHEDULING READ MODEL
 * =====================================================================
 * Read-only extension over the authoritative CI Assignments table.
 *
 * CI Assignments establish evaluator/date/site coverage.
 * Student schedule resolution remains owned by the existing student
 * scheduling functions above.
 *
 * This layer DOES NOT infer CI-to-room relationships.
 * =====================================================================
 */


/**
 * Return all current CI assignments for one evaluator on one date.
 *
 * Multiple assignments on the same date are valid and are preserved.
 */
function scheduling_getCiAssignmentsForDate(evaluatorId, date) {
  if (!evaluatorId) {
    throw new Error('Evaluator ID is required.');
  }

  const targetDate = appDateOnly_(date);

  if (!targetDate) {
    throw new Error('A valid date is required.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rows = appSheetToObjects_(
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.CI_ASSIGNMENTS
    )
  );

  return rows
    .filter(row => {
      if (
        appNormalize_(row['Evaluator ID']) !==
        appNormalize_(evaluatorId)
      ) {
        return false;
      }

      const assignmentDate =
        appDateOnly_(row['Date']);

      if (
        !assignmentDate ||
        assignmentDate.getTime() !==
          targetDate.getTime()
      ) {
        return false;
      }

      return scheduling_isCurrentCiAssignment_(row);
    })
    .sort((a, b) =>
      scheduling_ciTimeValue_(a['Start Time']) -
      scheduling_ciTimeValue_(b['Start Time'])
    );
}


/**
 * Return current CI assignments across an inclusive date range.
 *
 * Intended to support the future "View Full Schedule" CI calendar.
 */
function scheduling_getCiScheduleForRange(
  evaluatorId,
  startDate,
  endDate
) {
  if (!evaluatorId) {
    throw new Error('Evaluator ID is required.');
  }

  const start = appDateOnly_(startDate);
  const end = appDateOnly_(endDate);

  if (!start || !end) {
    throw new Error(
      'Valid start and end dates are required.'
    );
  }

  if (end.getTime() < start.getTime()) {
    throw new Error(
      'End date cannot precede start date.'
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rows = appSheetToObjects_(
    appGetSheet_(
      ss,
      APP_CONFIG.SHEETS.CI_ASSIGNMENTS
    )
  );

  return rows
    .filter(row => {
      if (
        appNormalize_(row['Evaluator ID']) !==
        appNormalize_(evaluatorId)
      ) {
        return false;
      }

      if (!scheduling_isCurrentCiAssignment_(row)) {
        return false;
      }

      const assignmentDate =
        appDateOnly_(row['Date']);

      return (
        assignmentDate &&
        assignmentDate.getTime() >= start.getTime() &&
        assignmentDate.getTime() <= end.getTime()
      );
    })
    .sort((a, b) => {
      const dateDifference =
        appDateOnly_(a['Date']).getTime() -
        appDateOnly_(b['Date']).getTime();

      if (dateDifference !== 0) {
        return dateDifference;
      }

      return (
        scheduling_ciTimeValue_(a['Start Time']) -
        scheduling_ciTimeValue_(b['Start Time'])
      );
    });
}


/**
 * Resolve the CI's operational workspace for one date.
 *
 * Each CI assignment is enriched with:
 * - Clinical Site name
 * - students assigned to that site/date
 * - each student's existing room rotation
 * - each student's Clinical Term
 *
 * IMPORTANT:
 * Student room rotations are displayed as student information only.
 * They do not establish CI-to-room coverage.
 */
function scheduling_getCiDayWorkspace(
  evaluatorId,
  date
) {
  if (!evaluatorId) {
    throw new Error('Evaluator ID is required.');
  }

  const targetDate = appDateOnly_(date);

  if (!targetDate) {
    throw new Error('A valid date is required.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const assignments =
    scheduling_getCiAssignmentsForDate(
      evaluatorId,
      targetDate
    );

  if (!assignments.length) {
    return {
      evaluatorId: evaluatorId,
      date: targetDate,
      assigned: false,
      assignments: []
    };
  }

  const students =
    scheduling_getUniqueActiveStudents_(ss);

  const resolvedAssignments =
    assignments.map(ciAssignment => {
      const siteId = String(
        ciAssignment['Site ID'] || ''
      ).trim();

      const site =
        scheduling_getSite_(ss, siteId);

      if (!site) {
        throw new Error(
          'CI scheduling data integrity conflict: assignment "' +
          String(
            ciAssignment['CI Assignment ID'] || ''
          ).trim() +
          '" references unknown Clinical Site ID "' +
          siteId +
          '".'
        );
      }

      const studentsAtSite =
        students
          .map(student =>
            scheduling_getStudentScheduleForDate(
              student.studentId,
              targetDate
            )
          )
          .filter(studentSchedule =>
            studentSchedule.scheduled === true &&
            appNormalize_(
              studentSchedule.siteId
            ) === appNormalize_(siteId)
          )
          .map(studentSchedule => ({
            studentId:
              studentSchedule.studentId,

            student:
              studentSchedule.student,

            cohort:
              studentSchedule.cohort,

            clinicalTerm:
              studentSchedule.clinicalTerm,

            siteId:
              studentSchedule.siteId,

            site:
              studentSchedule.site,

            rotationArea:
              studentSchedule.rotationArea,

            scheduleStatus:
              studentSchedule.status
          }))
          .sort((a, b) =>
            a.student.localeCompare(b.student)
          );

      return {
        ciAssignmentId: String(
          ciAssignment['CI Assignment ID'] || ''
        ).trim(),

        evaluatorId: String(
          ciAssignment['Evaluator ID'] || ''
        ).trim(),

        date: targetDate,

        siteId: siteId,
        site: site.name,

        startTime:
  scheduling_formatTime_(
    ciAssignment['Start Time']
  ),

endTime:
  scheduling_formatTime_(
    ciAssignment['End Time']
  ),

        assignmentType: String(
          ciAssignment['Assignment Type'] || ''
        ).trim(),

        status: String(
          ciAssignment['Status'] || ''
        ).trim(),

        notes: String(
          ciAssignment['Notes'] || ''
        ).trim(),

        students: studentsAtSite
      };
    });

  return {
    evaluatorId: evaluatorId,
    date: targetDate,
    assigned: true,
    assignments: resolvedAssignments
  };
}


/**
 * CI Assignment rows currently use "Current" for active coverage.
 *
 * Blank status is NOT treated as current because CI Assignments is a
 * published scheduling authority and Home must not infer coverage.
 */
function scheduling_isCurrentCiAssignment_(row) {
  return (
    appNormalize_(row['Status']) ===
    appNormalize_('Current')
  );
}


/**
 * Numeric sorting helper for CI assignment start times.
 * Missing times intentionally sort after timed assignments.
 */
function scheduling_ciTimeValue_(value) {
  if (!value) {
    return Number.MAX_SAFE_INTEGER;
  }

  if (value instanceof Date) {
    return (
      value.getHours() * 60 +
      value.getMinutes()
    );
  }

  const text =
    String(value).trim();

  const match =
    text.match(
      /^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i
    );

  if (!match) {
    return Number.MAX_SAFE_INTEGER;
  }

  let hour =
    Number(match[1]);

  const minute =
    Number(match[2]);

  const meridiem =
    String(match[3] || '').toUpperCase();

  if (meridiem === 'PM' && hour < 12) {
    hour += 12;
  }

  if (meridiem === 'AM' && hour === 12) {
    hour = 0;
  }

  return hour * 60 + minute;
}