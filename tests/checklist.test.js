'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const util = require('../js/core/util');

test('parseChecklist() extracts markdown checkboxes and other text', () => {
  const body = 'Intro\n- [ ] First\n- [x] Second\nOutro\n';
  const r = util.parseChecklist(body);

  assert.equal(Array.isArray(r.items), true);
  assert.equal(r.items.length, 2);

  assert.equal(r.items[0].lineIndex, 1);
  assert.equal(r.items[0].checked, false);
  assert.equal(r.items[0].text, 'First');

  assert.equal(r.items[1].lineIndex, 2);
  assert.equal(r.items[1].checked, true);
  assert.equal(r.items[1].text, 'Second');

  assert.equal(r.otherText, 'Intro\nOutro\n');
});

test('toggleChecklistItem() preserves original EOL style', () => {
  const body = 'A\r\n- [ ] One\r\nB\r\n';
  const out = util.toggleChecklistItem(body, 1, true);
  assert.equal(out.indexOf('\r\n') !== -1, true);
  assert.equal(out.split('\r\n')[1], '- [x] One');
});

test('addChecklistItem() appends with blank line when no checklist exists', () => {
  const body = 'Hello';
  const out = util.addChecklistItem(body, 'Item');
  assert.equal(out, 'Hello\n\n- [ ] Item');
});

test('addChecklistItem() inserts after last checklist item when present', () => {
  const body = 'Intro\n- [ ] One\n- [x] Two\nOutro';
  const out = util.addChecklistItem(body, 'Three');
  assert.equal(out, 'Intro\n- [ ] One\n- [x] Two\n- [ ] Three\nOutro');
});

test('mergeNotesAndChecklist() preserves checklist lines and uses original EOL style', () => {
  const original = 'Intro\r\n- [ ] First\r\n- [x] Second\r\nOutro\r\n';
  const notes = 'New intro\nNew outro';
  const out = util.mergeNotesAndChecklist(original, notes);
  assert.equal(out.indexOf('\r\n') !== -1, true);
  assert.equal(out, 'New intro\r\nNew outro\r\n\r\n- [ ] First\r\n- [x] Second');
});

test('mergeNotesAndChecklist() returns checklist-only when notes are empty', () => {
  const original = 'A\n- [ ] One\nB\n';
  const out = util.mergeNotesAndChecklist(original, '   ');
  assert.equal(out, '- [ ] One');
});

test('mergeNotesAndChecklist() returns notes when original has no checklist', () => {
  const original = 'Just notes';
  const out = util.mergeNotesAndChecklist(original, 'New\nText');
  assert.equal(out, 'New\nText');
});
