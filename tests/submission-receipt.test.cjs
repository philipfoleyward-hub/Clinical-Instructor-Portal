'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '..', 'Clinical-Instructor-Portal', 'CI_Portal_Competency_Approved_Inner.html'), 'utf8');
assert.match(source, /const reviewDoc=msg\.payload\?\.reviewDocument\|\|\{\};/, 'Student Review receipt must come from trusted submit result');
for (const label of ['Open evaluation document', 'Open evaluation PDF', 'Open Student Review document', 'Open Student Review PDF']) {
  assert(source.includes(label), 'Missing receipt action: ' + label);
}
assert(source.includes('(?:docs|drive)'), 'Receipt links must restrict the hostname to approved Google Drive/Docs');
assert.match(source, /Student Review receipt was not returned/, 'Missing companion file must be visible');
assert.match(source, /summary\.append\(missing\)/, 'Companion warning should appear at the submission result');
assert.match(source, /id="ci-final-submit" disabled/, 'Finalize must remain disabled by default');
console.log('PASS: submission receipt displays official and companion links and reports missing companion');
