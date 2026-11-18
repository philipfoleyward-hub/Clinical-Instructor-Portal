'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.join(__dirname, '..');
const original = path.join(root, 'Original-CCMS');
const portal = path.join(root, 'Clinical-Instructor-Portal');
const read = file => fs.readFileSync(file, 'utf8');
const gateway = read(path.join(original, 'CIPortalTest001.js'));
const generator = read(path.join(original, '05_EvaluationDocumentGenerator.js'));
const review = read(path.join(original, '05_StudentReviewDocumentService.js'));
const filing = read(path.join(original, '05_DriveServiceCompanion.js'));
const ui = read(path.join(portal, 'CI_Portal_Competency_Approved_Inner.html'));

// Do not replace the newer full gateway with an older test-only candidate.
assert.equal((gateway.match(/function doPost\(/g)||[]).length, 1);
assert.equal((gateway.match(/function ciLink_commitProduction_\(/g)||[]).length, 1);
assert.equal((gateway.match(/function ciLink_commitTest001_\(/g)||[]).length, 1);
const testSection = gateway.split('function ciLink_commitTest001_(data) {')[1].split('/** Production release')[0];
assert(testSection.includes("CI_PORTAL_DOCUMENT_PAIR_MODE"));
assert(testSection.includes("outcome.grade.qaOutput = true"));
assert(testSection.includes("outcome.grade.qaOutputFolderId = qaFolderId"));
assert(testSection.includes('ccmsStudentReview_generate_('));
assert(testSection.indexOf('receipt.officialDocument = doc') < testSection.indexOf('receipt.reviewDocument = reviewDoc'));
assert(testSection.indexOf('receipt.reviewDocument = reviewDoc') < testSection.indexOf('clinicalTracking_addRecord('));
assert(testSection.includes('reviewDocument: reviewDoc'));
assert(testSection.includes('ciLink_duplicate_(id)'));
assert(filing.includes("CI_PORTAL_QA_EVALUATIONS_FOLDER_ID"));
assert(review.includes("'{{SUBMISSION_ID}}': submissionId"));

// Isolated QA folder is an explicit configured ID, never a folder search by name.
const start = generator.indexOf('function evaluationDocument_getQaFolder_(qaFolderId) {');
const stop = generator.indexOf('function evaluationDocument_copyTemplateForQa_(', start);
assert(start > 0 && stop > start);
const guards = generator.slice(start, stop);
const settings = {CI_PORTAL_DOCUMENT_PAIR_MODE: 'test_only', CI_PORTAL_QA_EVALUATIONS_FOLDER_ID: 'approved-id'};
let folderName = 'Evaluations', accesses = 0;
const ctx = {
  PropertiesService: {getScriptProperties: () => ({getProperty: key => settings[key] || ''})},
  DriveApp: {
    getFolderById: id => {accesses++; assert.equal(id, 'approved-id'); return {getName: () => folderName, getId: () => id};},
    getFoldersByName: () => {throw Error('Must not search by folder name when configured')},
    createFolder: () => {throw Error('Must not create a QA folder when configured')}
  }
};
vm.createContext(ctx);
vm.runInContext(guards, ctx);
assert.equal(ctx.evaluationDocument_getQaFolder_('approved-id').getId(), 'approved-id');
assert.equal(accesses, 1);
assert.throws(() => ctx.evaluationDocument_getQaFolder_('wrong-id'), /not authorized/);
settings.CI_PORTAL_DOCUMENT_PAIR_MODE='off';
assert.throws(() => ctx.evaluationDocument_getQaFolder_('approved-id'), /not authorized/);
settings.CI_PORTAL_DOCUMENT_PAIR_MODE='test_only';
folderName='Other';
assert.throws(() => ctx.evaluationDocument_getQaFolder_('approved-id'), /named Evaluations/);

// The existing confirmation stays at the user's point of attention.
for (const action of ['Open evaluation document', 'Open evaluation PDF', 'Open Student Review document', 'Open Student Review PDF']) {
  assert(ui.includes(action), `Missing ${action}`);
}
assert(ui.includes('Student Review receipt was not returned'));
assert(ui.includes('summary.append(missing)'));
assert(ui.includes('id="ci-final-submit" disabled'));
console.log('PASS: dual-doc TEST-001 gateway, explicit folder release, receipt ordering, four-link UI');