/**
 * Dashboard adapter for the existing document generators and DriveService.
 *
 * This file does not create templates, tokens, requirement structures, or a
 * reporting engine. It gives DashboardUI one stable entry point for the
 * already-existing generators and lists files from the established student
 * Drive folder.
 */

function dashboardGenerateStudentDocument(request) {
  request = request || {};

  const studentId = String(request.studentId || '').trim();
  const documentType = appNormalize_(request.documentType || '');
  const clinicalTerm = String(request.clinicalTerm || '').trim();

  if (!studentId) {
    throw new Error('Student ID is required.');
  }

  let result;

  if (documentType === 'semester') {
    return dashboardGenerateSemesterRequirement(request);

  } else if (documentType === 'program') {
    return dashboardGenerateProgramRequirement(request);

  } else if (documentType === 'grades') {
    return dashboardGenerateClinicalGrades(request);

  } else if (documentType === 'progress') {
    throw new Error('Progress Report generation is reserved for a future generator.');

  } else {
    throw new Error('Unsupported document type: ' + (request.documentType || 'blank'));
  }

}


function dashboardGenerateSemesterRequirement(request) {
  request = request || {};
  const studentId = String(request.studentId || '').trim();
  const clinicalTerm = String(request.clinicalTerm || '').trim();

  if (!studentId) throw new Error('Student ID is required.');
  if (!clinicalTerm) throw new Error('Clinical term is required.');

  const result = generateSemesterRequirementSheet(studentId, clinicalTerm);
  return dashboardDocuments_standardizeGenerationResult_(
    result,
    'semester',
    clinicalTerm
  );
}


function dashboardGenerateProgramRequirement(request) {
  request = request || {};
  const studentId = String(request.studentId || '').trim();

  if (!studentId) throw new Error('Student ID is required.');

  const result = generateProgramRequirementSheet(studentId);
  return dashboardDocuments_standardizeGenerationResult_(
    result,
    'program',
    ''
  );
}


function dashboardGenerateClinicalGrades(request) {
  request = request || {};
  const studentId = String(request.studentId || '').trim();
  const requestId = String(request.requestId || 'NO-REQUEST-ID').trim();

  if (!studentId) throw new Error('Student ID is required.');

  console.log('[Grades][' + requestId + '] Dashboard wrapper received request.');
  const started = Date.now();

  const result = generateClinicalGradesSheet(studentId, requestId);

  console.log(
    '[Grades][' + requestId + '] Existing Grades generator returned after ' +
    (Date.now() - started) + ' ms.'
  );

  const response = dashboardDocuments_standardizeGenerationResult_(
    result,
    'grades',
    ''
  );

  console.log('[Grades][' + requestId + '] Plain result returned to dashboard wrapper/client.');
  return response;
}


function dashboardDocuments_standardizeGenerationResult_(result, documentType, clinicalTerm) {
  result = result || {};

  const fileId = String(
    result.documentId ||
    result.spreadsheetId ||
    result.fileId ||
    ''
  ).trim();

  if (!fileId) {
    throw new Error('The generator completed without returning a generated file ID.');
  }

  const fileName = String(result.name || result.fileName || '').trim();
  const viewUrl = String(result.url || result.fileUrl || '').trim();

  return {
    success: true,
    documentType: documentType,
    clinicalTerm: clinicalTerm || '',
    fileId: fileId,
    fileName: fileName || 'Generated document',
    viewUrl: viewUrl,
    pdfUrl: null
  };
}


function dashboardGetStudentDocuments(studentId) {
  const resolved = dashboardDocuments_getStudent_(studentId);

  const studentFolder = dashboardDocuments_findExistingStudentFolder_(resolved);

  const files = [];
  dashboardDocuments_collectFiles_(
    studentFolder,
    '',
    0,
    files,
    150
  );

  files.sort(function(a, b) {
    return b.updatedValue - a.updatedValue;
  });

  return {
    success: true,
    studentId: resolved.studentId,
    folderUrl: studentFolder.getUrl(),
    capabilities: dashboardDocuments_getCapabilities_(),
    files: files.map(function(file) {
      return {
        fileId: file.fileId,
        name: file.name,
        url: file.url,
        mimeType: file.mimeType,
        updated: file.updated,
        location: file.location
      };
    })
  };
}


function dashboardGetDocumentWorkspaceCapabilities() {
  return dashboardDocuments_getCapabilities_();
}


function dashboardDocuments_getCapabilities_() {
  const semesterTerms = [
    'Clinical 1',
    'Clinical 2',
    'Clinical 3',
    'Clinical 4',
    'Clinical 5'
  ].filter(function(term) {
    if (typeof semesterDocument_getTemplateLayout_ !== 'function') {
      return false;
    }

    const layout = semesterDocument_getTemplateLayout_(term);
    return Boolean(layout && layout.templateId);
  });

  const programAvailable =
    typeof generateProgramRequirementSheet === 'function' &&
    typeof PROGRAM_DOCUMENT_TEMPLATE_ID !== 'undefined' &&
    Boolean(PROGRAM_DOCUMENT_TEMPLATE_ID);

  const gradesAvailable =
    typeof generateClinicalGradesSheet === 'function' &&
    typeof CLINICAL_GRADES_TEMPLATE_ID !== 'undefined' &&
    Boolean(CLINICAL_GRADES_TEMPLATE_ID);

  return {
    semester: {
      available: semesterTerms.length > 0,
      terms: semesterTerms
    },
    program: {
      available: programAvailable,
      templateConfigured: programAvailable
    },
    grades: {
      available: gradesAvailable,
      templateConfigured: gradesAvailable
    },
    progressReports: {
      available: false
    }
  };
}


function dashboardDocuments_findExistingStudentFolder_(student) {
  const cohortName = driveService_cohortFolderName_(student.cohort);
  const expectedName = appSafeFileText_(student.student);
  const targetName = appNormalize_(student.student);
  const targetId = appNormalize_(student.studentId);
  const roots = [driveService_getClinicalEvaluationsRoot_()];

  if (typeof competencyIndex_getUniqueFolderByName_ === 'function') {
    try {
      const legacyRoot = competencyIndex_getUniqueFolderByName_('Clinical Evaluations');
      if (legacyRoot && !roots.some(function(root) { return root.getId() === legacyRoot.getId(); })) {
        roots.push(legacyRoot);
      }
    } catch (ignored) {}
  }

  const candidates = [];

  roots.forEach(function(root) {
    const cohortFolders = root.getFoldersByName(cohortName);

    while (cohortFolders.hasNext()) {
      const cohortFolder = cohortFolders.next();
      const exact = cohortFolder.getFoldersByName(expectedName);

      while (exact.hasNext()) {
        candidates.push(exact.next());
      }

      if (candidates.length) {
        continue;
      }

      const folders = cohortFolder.getFolders();
      while (folders.hasNext()) {
        const folder = folders.next();
        const normalized = appNormalize_(folder.getName());

        if (
          normalized === targetName ||
          (targetId && normalized.indexOf(targetId) !== -1) ||
          (targetName && normalized.indexOf(targetName) !== -1)
        ) {
          candidates.push(folder);
        }
      }
    }
  });

  const unique = {};
  candidates.forEach(function(folder) {
    unique[folder.getId()] = folder;
  });
  const matches = Object.keys(unique).map(function(id) { return unique[id]; });

  if (matches.length !== 1) {
    throw new Error(
      matches.length === 0
        ? 'Could not resolve the established Drive folder for ' + student.student + '. No folder was created.'
        : 'More than one established Drive folder could match ' + student.student + '.'
    );
  }

  return matches[0];
}


function dashboardDocuments_getStudent_(studentId) {
  const target = appNormalize_(studentId);

  if (!target) {
    throw new Error('Student ID is required.');
  }

  const rows = appSheetToObjects_(
    appGetSheet_(
      SpreadsheetApp.getActiveSpreadsheet(),
      APP_CONFIG.SHEETS.STUDENTS
    )
  ).filter(function(row) {
    return appNormalize_(row['Student ID']) === target;
  });

  if (!rows.length) {
    throw new Error('Student not found: ' + studentId);
  }

  const row = typeof dashGetMostAdvancedStudentRow_ === 'function'
    ? dashGetMostAdvancedStudentRow_(rows)
    : rows[0];

  const student = String(row['Student'] || '').trim();
  const cohort = String(row['Cohort'] || '').trim();

  if (!student || !cohort) {
    throw new Error('Student name or cohort is missing for: ' + studentId);
  }

  return {
    studentId: String(row['Student ID'] || studentId).trim(),
    student: student,
    cohort: cohort
  };
}


function dashboardDocuments_collectFiles_(folder, path, depth, output, limit) {
  if (output.length >= limit || depth > 2) {
    return;
  }

  const files = folder.getFiles();
  while (files.hasNext() && output.length < limit) {
    const file = files.next();
    const updated = file.getLastUpdated();

    output.push({
      fileId: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      mimeType: file.getMimeType(),
      updated: Utilities.formatDate(
        updated,
        Session.getScriptTimeZone(),
        'MMM d, yyyy h:mm a'
      ),
      updatedValue: updated.getTime(),
      location: path || 'Student Folder'
    });
  }

  if (depth >= 2 || output.length >= limit) {
    return;
  }

  const childFolders = folder.getFolders();
  while (childFolders.hasNext() && output.length < limit) {
    const child = childFolders.next();
    const childPath = path ? path + ' / ' + child.getName() : child.getName();
    dashboardDocuments_collectFiles_(child, childPath, depth + 1, output, limit);
  }
}
