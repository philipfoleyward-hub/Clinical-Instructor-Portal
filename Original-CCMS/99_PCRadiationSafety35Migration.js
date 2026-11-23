/**
 * ONE-OFF (safe to delete after it has run): adds the missing Clinical 3-5
 * criterion 13 "RADIATION SAFETY CONSCIOUSNESS" (weight 4) to the live
 * Professional Competency Google Form and the approved 3-5 Sheet template.
 * Idempotent: re-running finds the existing question/row and changes nothing.
 * Returns the new Form item ID, which PC_V2_RUBRICS['Clinical 3-5'] must use.
 */
const PC_RS35_TITLE = 'RADIATION SAFETY CONSCIOUSNESS: Assess possibility of pregnancy, use of accurate technique book, uses calipers correctly, effective shielding, and accurate collimation.';

function pcV2_addRadiationSafety35() {
  const form = FormApp.openById(PC_V2_FORM_ID);
  // ---- Form: locate the Clinical 3-5 weighted questions ----
  let current = '', list35 = [];
  form.getItems().forEach(function(item) {
    if (item.getType() === FormApp.ItemType.PAGE_BREAK) {
      const t = pcV2_normalizedText_(item.getTitle());
      current = /Professional Competency.*Clinical Practicum 3.*5/i.test(t) ? '35' : '';
      return;
    }
    if (current === '35' && item.getType() === FormApp.ItemType.LIST &&
        /^Weight:\s*\d+\s*$/i.test(pcV2_text_(item.getHelpText()))) list35.push(item);
  });
  let added = list35.filter(function(x) {return pcV2_normalizedText_(x.getTitle()) === PC_RS35_TITLE;})[0];
  if (!added) {
    if (list35.length !== 26) throw new Error('Expected 26 Clinical 3-5 questions, found ' + list35.length);
    const after = list35[11];  // criterion 12 DISPLAYS PHYSICAL SAFETY CONSCIOUSNESS
    if (!/^DISPLAYS PHYSICAL SAFETY CONSCIOUSNESS/i.test(pcV2_normalizedText_(after.getTitle())))
      throw new Error('Clinical 3-5 criterion 12 is not the expected Physical Safety item; nothing changed.');
    const model = after.asListItem();
    const item = form.addListItem()
      .setTitle(PC_RS35_TITLE)
      .setHelpText('Weight: 4')
      .setChoiceValues(model.getChoices().map(function(c) {return c.getValue();}))
      .setRequired(model.isRequired());
    form.moveItem(item.getIndex(), after.getIndex() + 1);
    added = form.getItemById(item.getId());
  }
  const id = added.getId();

  // ---- Template: insert row 27 below criterion 12 (row 26) ----
  const sheet = SpreadsheetApp.openById(PC_V2_TEMPLATES['Clinical 3-5']).getSheetByName('Sheet1');
  const token = '{{PC_' + id + '_RATING}}';
  if (sheet.getRange(27, 7).getDisplayValue() !== token) {
    if (sheet.getRange(26, 7).getDisplayValue() !== '{{PC_195430175_RATING}}' ||
        sheet.getRange(27, 7).getDisplayValue() !== '{{PC_253887405_RATING}}')
      throw new Error('Template rows 26/27 are not the expected criteria; template unchanged.');
    sheet.insertRowAfter(26);
    sheet.getRange(26, 1, 1, sheet.getMaxColumns()).copyTo(sheet.getRange(27, 1, 1, sheet.getMaxColumns()));
    sheet.setRowHeight(27, sheet.getRowHeight(26));
    sheet.getRange(27, 1).setValue('13.  ' + PC_RS35_TITLE);
    sheet.getRange(27, 7).setValue(token);
    sheet.getRange(27, 8).setValue(4);
    // Renumber the template's printed numbers 13–26 → 14–27 (CCMS also rewrites them per submission).
    for (let r = 28; r <= 43; r++) {
      const cell = sheet.getRange(r, 1), v = String(cell.getValue());
      const m = /^(\d+)\.(\s+[\s\S]*)$/.exec(v);
      if (m && Number(m[1]) >= 13) cell.setValue((Number(m[1]) + 1) + '.' + m[2]);
    }
    SpreadsheetApp.flush();
  }
  // Completeness guards in the totals row must count 27 criteria now.
  ['G43','H43','I43','J43'].forEach(function(a1) {
    const r = sheet.getRange(a1), fx = r.getFormula();
    if (fx && fx.indexOf('<26') >= 0) r.setFormula(fx.split('<26').join('<27'));
  });
  SpreadsheetApp.flush();
  const f = function(a1) {return sheet.getRange(a1).getFormula() || sheet.getRange(a1).getDisplayValue();};
  const out = {formItemId: id, formIndex: added.getIndex(),
    row27: [f('I27'), f('J27')], row26: [f('I26'), f('J26')], G27: f('G27'),
    I43: f('I43'), J43: f('J43'), H43: f('H43'), G43: f('G43'), G45: f('G45'), E43: f('E43'), A50: f('A50'), A53: f('A53'),
    A28: String(sheet.getRange('A28').getValue()).slice(0, 20), A42: String(sheet.getRange('A42').getValue()).slice(0, 20)};
  console.log('PC_RS35 ' + JSON.stringify(out));
  return out;
}
