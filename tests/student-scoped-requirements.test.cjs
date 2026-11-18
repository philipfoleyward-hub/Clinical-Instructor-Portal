'use strict';
const assert = require('node:assert/strict');
const {readFileSync} = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');
const ctx = vm.createContext({});
for (const file of ['Original-CCMS/04_RequirementsEngine.js', 'Clinical-Instructor-Portal/CIPortalEvaluationTransport.js', 'Clinical-Instructor-Portal/PortalDataService.js']) {
  vm.runInContext(readFileSync(path.join(__dirname, '..', file), 'utf8'), ctx);
}
const headers = Array.from(vm.runInContext('REQUIREMENTS_STATUS_HEADERS', ctx));
function row(student, req, done) {
  const v = Array(headers.length).fill('');
  Object.assign(v, {0: student, 2: 'Clinical 3', 3: 'Detail', 4: 'Program', 5: req, 11: done});
  return v;
}
function fixture() {
  const rows = [headers, row('other', 'A', 0), row('TEST-001', 'A', 0),
    row('other', 'B', 1), row('TEST-001', 'retired', 0)];
  const writes = [];
  const sheet = {
    getLastRow: () => rows.length, getMaxRows: () => 1000, getLastColumn: () => headers.length,
    getRange(start, col, count, width) {
      return {
        getValues: () => rows.slice(start-1,start-1+count).map(r => r.slice(col-1,col-1+width)),
        setValues(values) {
          writes.push([start,count]);
          values.forEach((r,i) => { rows[start-1+i] = Array.from(r); });
        }
      };
    }
  };
  return {rows,writes,sheet};
}
test('scoped writes preserve neighboring students and converge on retry', () => {
  const {rows,writes,sheet} = fixture();
  const output = [row('TEST-001','A',1),row('TEST-001','B',0)];
  const other = [rows[1],rows[3]];
  ctx.reqWriteStudentRows_(sheet,'TEST-001',output);
  const after = JSON.stringify(rows);
  ctx.reqWriteStudentRows_(sheet,'TEST-001',output);
  assert.equal(JSON.stringify(rows),after);
  assert.deepEqual([rows[1],rows[3]],other);
  assert.equal(rows[2][11],1);
  assert.equal(rows[4][0],'');
  assert.deepEqual(rows[5],output[1]);
  assert(writes.every(([start,count]) => start!==2 && !(start<=4 && start+count>4)));
});
test('ambiguous output is rejected before the first write', () => {
  const {sheet,writes}=fixture();
  for (const output of [[row('other','A',1)], [row('TEST-001','A',1),row('TEST-001','A',1)]]) {
    assert.throws(()=>ctx.reqWriteStudentRows_(sheet,'TEST-001',output));
  }
  assert.deepEqual(writes,[]);
});
test('exam picker reads current rows for interleaved student IDs', () => {
  const {sheet,rows}=fixture(), book={getSheetByName:()=>sheet};
  const read=()=>Array.from(ctx.ciPortalStudentRequirements_(book,'TEST-001'));
  assert.deepEqual(read().map(r=>r['Requirement ID']),['A','retired']);
  rows[2][11]=1;
  assert.equal(read()[0]['Completed Count'],1);
  assert.equal(ctx.ciPortalStudentRequirements_(book,'absent').length,0);
});
