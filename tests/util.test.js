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
