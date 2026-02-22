'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const themeSafety = require('../js/core/theme-safety');

test('isCssLocalOnly() blocks remote loads and IE scriptable CSS', () => {
  assert.equal(themeSafety.isCssLocalOnly('.a{color:#fff}'), true);

  assert.equal(themeSafety.isCssLocalOnly('/* http://example.com */ .a{color:#fff}'), true);
  assert.equal(themeSafety.isCssLocalOnly('url(https://example.com/x.png)'), false);
  assert.equal(themeSafety.isCssLocalOnly('div{background-image:url(//example.com/x.png)}'), false);
  assert.equal(themeSafety.isCssLocalOnly("div{filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(src='//example.com/x.png',sizingMethod='scale');}"), false);
  assert.equal(themeSafety.isCssLocalOnly('@import url("x.css");'), false);
  assert.equal(themeSafety.isCssLocalOnly('div{behavior: url(x.htc);}'), false);
  assert.equal(themeSafety.isCssLocalOnly('div{width: expression(alert(1));}'), false);
  assert.equal(themeSafety.isCssLocalOnly('div{background-image:url(javascript:alert(1))}'), false);
});

test('isSafeLocalCssPath() enforces relative local paths', () => {
  assert.equal(themeSafety.isSafeLocalCssPath('themes/my-theme/theme.css'), true);
  assert.equal(themeSafety.isSafeLocalCssPath('themes/my_theme/theme.css'), true);

  assert.equal(themeSafety.isSafeLocalCssPath('../themes/x.css'), false);
  assert.equal(themeSafety.isSafeLocalCssPath('/themes/x.css'), false);
  assert.equal(themeSafety.isSafeLocalCssPath('./themes/x.css'), false);
  assert.equal(themeSafety.isSafeLocalCssPath('themes\\x\\theme.css'), false);
  assert.equal(themeSafety.isSafeLocalCssPath('C:\\temp\\x.css'), false);
  assert.equal(themeSafety.isSafeLocalCssPath('themes/x/theme.scss'), false);
});
