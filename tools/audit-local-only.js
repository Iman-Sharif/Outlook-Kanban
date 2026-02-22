#!/usr/bin/env node
'use strict';

/*
  Local-only audit

  Goal: prevent accidental introduction of runtime network dependencies.
  This is intentionally conservative for app-owned runtime assets.

  What it checks:
  - HTML: no http/https/protocol-relative resources in src/href/action attributes
  - CSS (app + themes): no @import rules, no remote/protocol-relative url(), no IE scriptable CSS
  - JS (app only): no fetch/XMLHttpRequest/ajax/beacon usage

  Excludes:
  - vendor/ and other third-party files (they may contain URLs in comments)
*/

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');

function readText(relPath) {
  return fs.readFileSync(path.join(ROOT, relPath), 'utf8');
}

function listFilesRecursive(absDir) {
  const out = [];
  const entries = fs.readdirSync(absDir, { withFileTypes: true });
  for (const e of entries) {
    const abs = path.join(absDir, e.name);
    if (e.isDirectory()) {
      out.push(...listFilesRecursive(abs));
    } else {
      out.push(abs);
    }
  }
  return out;
}

function toRel(absPath) {
  return path.relative(ROOT, absPath).replace(/\\/g, '/');
}

function stripCssComments(s) {
  // Remove /* ... */
  return String(s || '').replace(/\/\*[\s\S]*?\*\//g, '');
}

function stripJsComments(s) {
  // Best-effort: remove // and /* */ comments without touching quoted strings.
  // Not a full JS parser; tailored for this ES5 codebase.
  const src = String(s || '');
  let out = '';
  let i = 0;
  let inSingle = false;
  let inDouble = false;
  let inLine = false;
  let inBlock = false;

  while (i < src.length) {
    const ch = src[i];
    const next = (i + 1 < src.length) ? src[i + 1] : '';

    if (inLine) {
      if (ch === '\n') {
        inLine = false;
        out += ch;
      }
      i++;
      continue;
    }

    if (inBlock) {
      if (ch === '*' && next === '/') {
        inBlock = false;
        i += 2;
        continue;
      }
      i++;
      continue;
    }

    if (!inDouble && ch === "'" && !isEscaped(src, i)) {
      inSingle = !inSingle;
      out += ch;
      i++;
      continue;
    }

    if (!inSingle && ch === '"' && !isEscaped(src, i)) {
      inDouble = !inDouble;
      out += ch;
      i++;
      continue;
    }

    if (!inSingle && !inDouble) {
      if (ch === '/' && next === '/') {
        inLine = true;
        i += 2;
        continue;
      }
      if (ch === '/' && next === '*') {
        inBlock = true;
        i += 2;
        continue;
      }
    }

    out += ch;
    i++;
  }

  return out;
}

function isEscaped(src, idx) {
  // Count backslashes immediately preceding idx
  let count = 0;
  for (let i = idx - 1; i >= 0; i--) {
    if (src[i] === '\\') count++;
    else break;
  }
  return (count % 2) === 1;
}

function fail(errors) {
  if (!errors.length) return;
  const msg = ['Local-only audit failed:', ''].concat(errors.map(e => '- ' + e)).join('\n');
  process.stderr.write(msg + '\n');
  process.exit(1);
}

function auditHtml(relPath, errors) {
  const s = readText(relPath);
  const bad = /\b(?:src|href|action)\s*=\s*["']\s*(https?:)?\/\//ig;
  let m;
  while ((m = bad.exec(s)) !== null) {
    errors.push(relPath + ': contains external/protocol-relative ' + m[0].trim());
  }
}

function auditCssFiles(relPaths, errors) {
  for (const relPath of relPaths) {
    const raw = readText(relPath);
    const s = stripCssComments(raw);

    // External loads (conservative)
    if (/\bhttps?:\/\//i.test(s)) {
      errors.push(relPath + ': contains http/https in CSS (disallowed)');
    }

    if (/@import\b/i.test(s)) {
      errors.push(relPath + ': contains @import (disallowed)');
    }
    if (/url\s*\(\s*["']?\s*(https?:)?\/\//i.test(s)) {
      errors.push(relPath + ': contains remote/protocol-relative url(...)');
    }
    // IE filter/legacy patterns that use src='//...'
    if (/\bsrc\s*=\s*["']\s*\/\//i.test(s)) {
      errors.push(relPath + ': contains protocol-relative src= (disallowed)');
    }
    if (/url\s*\(\s*["']?\s*javascript:/i.test(s)) {
      errors.push(relPath + ': contains javascript: url(...)');
    }
    if (/\bsrc\s*=\s*["']\s*javascript:/i.test(s)) {
      errors.push(relPath + ': contains javascript: src= (disallowed)');
    }
    if (/expression\s*\(/i.test(s)) {
      errors.push(relPath + ': contains IE scriptable CSS expression(...)');
    }
    if (/behavior\s*:/i.test(s)) {
      errors.push(relPath + ': contains IE scriptable CSS behavior:');
    }
  }
}

function auditJsFiles(relPaths, errors) {
  for (const relPath of relPaths) {
    const raw = readText(relPath);
    const s = stripJsComments(raw);

    if (/\bfetch\s*\(/.test(s)) {
      errors.push(relPath + ': uses fetch() (disallowed)');
    }
    if (/\bXMLHttpRequest\b/.test(s)) {
      errors.push(relPath + ': references XMLHttpRequest (disallowed)');
    }
    if (/\bnavigator\.sendBeacon\b/.test(s)) {
      errors.push(relPath + ': references navigator.sendBeacon (disallowed)');
    }
    if (/\b\$\.ajax\b|\bjQuery\.ajax\b/.test(s)) {
      errors.push(relPath + ': references $.ajax/jQuery.ajax (disallowed)');
    }

    // ActiveX HTTP clients (IE)
    if (/\bActiveXObject\s*\(\s*["']\s*(msxml2\.(?:server)?xmlhttp(?:\.[0-9.]+)?|microsoft\.xmlhttp|winhttp\.winhttprequest(?:\.[0-9.]+)?)\b/i.test(s)) {
      errors.push(relPath + ': references ActiveX HTTP client (disallowed)');
    }
  }
}

function main() {
  const errors = [];

  // HTML pages shipped with the app
  [
    'kanban.html',
    'upgrade.html',
    'whatsnew.html'
  ].forEach(p => auditHtml(p, errors));

  // CSS: app-owned styles and themes (exclude vendor CSS)
  const cssRoots = [path.join(ROOT, 'css'), path.join(ROOT, 'themes')];
  const cssRel = [];
  for (const absRoot of cssRoots) {
    for (const absFile of listFilesRecursive(absRoot)) {
      if (!/\.css$/i.test(absFile)) continue;
      cssRel.push(toRel(absFile));
    }
  }
  // Only enforce for kfo + themes; bootstrap is vendored (but local). Still safe to include.
  auditCssFiles(cssRel, errors);

  // JS: only app-owned JS (exclude vendor/)
  const jsAbs = listFilesRecursive(path.join(ROOT, 'js'));
  const jsRel = jsAbs
    .filter(p => /\.js$/i.test(p))
    .map(toRel);
  auditJsFiles(jsRel, errors);

  fail(errors);
  process.stdout.write('Local-only audit: OK\n');
}

main();
