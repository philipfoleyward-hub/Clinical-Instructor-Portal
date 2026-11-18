/**
 * CCMS ORIGINAL PROJECT ONLY — DriveService extension for a companion file.
 * CANDIDATE; install only after reconciling original DriveService HEAD.
 *
 * Never creates a student or cohort folder. Reuses the verified folder of
 * the authorized V5.2 Google Doc AND its PDF, which must be siblings.
 * The official files must have been generated server-side in this transaction.
 */
function driveService_submissionFolder_(options) {
  options = options || {};
  if (!options.officialDocId || !options.officialPdfId) {
    throw new Error('Both official evaluation file IDs are required');
  }
  const docFile = DriveApp.getFileById(options.officialDocId);
  const pdfFile = DriveApp.getFileById(options.officialPdfId);
  const parents = docFile.getParents();
  if (!parents.hasNext()) throw new Error('Official evaluation is not filed in a folder');
  const folder = parents.next();
  if (parents.hasNext()) throw new Error('Official evaluation has ambiguous parents');
  const pdfParents = pdfFile.getParents();
  if (!pdfParents.hasNext() || pdfParents.next().getId() !== folder.getId() || pdfParents.hasNext()) {
    throw new Error('Official Doc and PDF are not in the same evaluation folder');
  }
  if (folder.getName() !== 'Evaluations') {
    throw new Error('Official Doc and PDF are not in an Evaluations folder');
  }
  // Fail closed on the current name-derived original filing path until the
  // stable-ID DriveService folder registry is reconciled and approved.
  // The TEST-001 QA destination must be explicitly allowed by configuration.
  const release = String(PropertiesService.getScriptProperties().getProperty('CI_PORTAL_DOCUMENT_PAIR_MODE') || 'off');
  if (release !== 'test_only') throw new Error('Dual-document output is disabled');
  const qaFolder = String(PropertiesService.getScriptProperties().getProperty('CI_PORTAL_QA_EVALUATIONS_FOLDER_ID') || '');
  if (!qaFolder || folder.getId() !== qaFolder) {
    throw new Error('Dual-document TEST-001 output is not in the approved QA Evaluations folder');
  }
  return folder;
}

function driveService_copySubmissionCompanion_(options) {
  const folder = driveService_submissionFolder_(options);
  const master = DriveApp.getFileById(options.templateId);
  const copy = master.makeCopy(options.fileName, folder);
  return {success: true, fileId: copy.getId(), fileUrl: copy.getUrl(), folderId: folder.getId()};
}
function driveService_saveSubmissionCompanionPdf_(options) {
  const folder = driveService_submissionFolder_(options);
  if (!options.blob) throw new Error('Companion PDF blob is missing');
  const file = folder.createFile(options.blob.copyBlob().setName(options.fileName));
  return {success: true, fileId: file.getId(), fileUrl: file.getUrl(), folderId: folder.getId()};
}