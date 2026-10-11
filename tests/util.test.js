'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const util = require('../js/core/util');

test('sanitizeId() normalises ids', () => {
  assert.equal(util.sanitizeId('My Theme'), 'my-theme');
  assert.equal(util.sanitizeId('  Hello   World  '), 'hello-world');
  assert.equal(util.sanitizeId('ABC_123'), 'abc123');
  assert.equal(util.sanitizeId('---a---'), 'a');
  assert.equal(util.sanitizeId(''), '');
});

test('isValidHexColor() validates #RRGGBB', () => {
  assert.equal(util.isValidHexColor('#abcdef'), true);
  assert.equal(util.isValidHexColor('#ABCDEF'), true);
  assert.equal(util.isValidHexColor('#12345'), false);
  assert.equal(util.isValidHexColor('123456'), false);
});

test('safeErrorString() is robust', () => {
  assert.equal(util.safeErrorString('boom'), 'boom');
  assert.equal(util.safeErrorString(new Error('nope')), 'nope');
  assert.equal(util.safeErrorString(null), '');
});

test('nowStamp() returns a stable stamp format', () => {
  const s = util.nowStamp();
  assert.match(s, /^\d{8}-\d{6}$/);
});

test('isRealDate() rejects invalid/sentinel dates', () => {
  assert.equal(util.isRealDate(new Date('invalid')), false);
  assert.equal(util.isRealDate(new Date(4501, 0, 1)), false);
  assert.equal(util.isRealDate(new Date()), true);
});

test('isSafeOpenUrl() allows normal links used by the app', () => {
  assert.equal(util.isSafeOpenUrl('onenote:https://contoso.example/notebook'), true);
  assert.equal(util.isSafeOpenUrl('https://contoso.example/path'), true);
});

test('isSafeOpenUrl() blocks unsafe schemes with mixed case and leading/trailing whitespace', () => {
  assert.equal(util.isSafeOpenUrl('  JavaScript:alert(1)  '), false);
  assert.equal(util.isSafeOpenUrl('\tvBsCrIpT:msgbox(1)\n'), false);
  assert.equal(util.isSafeOpenUrl(' data:text/html,hello'), false);
});

test('isSafeOpenUrl() blocks protocol-relative links', () => {
  assert.equal(util.isSafeOpenUrl('//example.com/path'), false);
});

test('parseQuickAddLine() extracts tokens and keeps subject text', () => {
  const p = util.parseQuickAddLine('Prepare Q1 report @tomorrow !high #Work #Finance /doing +private');
  assert.equal(p.subject, 'Prepare Q1 report');
  assert.equal(p.due, 'tomorrow');
  assert.equal(p.priority, 2);
  assert.equal(p.laneId, 'doing');
  assert.deepEqual(p.categories, ['Work', 'Finance']);
  assert.equal(p.sensitivity, 2);
});

test('parseQuickAddLine() supports plus-days and nodue tokens', () => {
  const p1 = util.parseQuickAddLine('Schedule follow-up @+3d');
  assert.equal(p1.subject, 'Schedule follow-up');
  assert.equal(p1.due, 'plus:3');

  const p2 = util.parseQuickAddLine('Tidy backlog @nodue');
  assert.equal(p2.subject, 'Tidy backlog');
  assert.equal(p2.due, 'nodue');

  const p3 = util.parseQuickAddLine('Share update +public');
  assert.equal(p3.subject, 'Share update');
  assert.equal(p3.sensitivity, 0);
});

test('parseQuickAddLine() leaves unknown tokens in subject', () => {
  const p = util.parseQuickAddLine('Investigate @later ???');
  assert.equal(p.subject, 'Investigate @later ???');
  assert.equal(p.due, null);
  assert.equal(p.priority, null);
  assert.equal(p.laneId, '');
  assert.deepEqual(p.categories, []);
  assert.equal(p.sensitivity, null);
});

test('parseQuickAddLine() does not treat numeric hashtags or plain words as tokens', () => {
  const p = util.parseQuickAddLine('Fix #123 public roadmap');
  assert.equal(p.subject, 'Fix #123 public roadmap');
  assert.deepEqual(p.categories, []);
  assert.equal(p.sensitivity, null);
});
