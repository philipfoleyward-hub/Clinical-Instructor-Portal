/**
 * CCMS ORIGINAL PROJECT ONLY — DriveService extension for a companion file.
 *
 * Never creates a student or cohort folder. Files the companion (Student
 * Review Doc/PDF) in the exact folder that already holds the official V5.2
 * Google Doc AND its PDF, which must be siblings. The official files must have
 * been generated server-side in this same submission.
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
