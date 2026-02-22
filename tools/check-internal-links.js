#!/usr/bin/env node
'use strict';

/*
  Internal link checker

  Checks Markdown links to local files/directories.
  - Only validates relative links (no http/https/mailto).
  - Ignores pure anchors (#...), and external links.
  - Ignores links inside fenced code blocks.
*/

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');

function listFilesRecursive(absDir) {
  const out = [];
  const entries = fs.readdirSync(absDir, { withFileTypes: true });
  for (const e of entries) {
    const abs = path.join(absDir, e.name);
    if (e.isDirectory()) {
      // skip node_modules if present
      if (e.name === 'node_modules') continue;
      if (e.name === '.git') continue;
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

function stripFencedCodeBlocks(md) {
  const lines = String(md || '').split(/\r?\n/);
  let inFence = false;
  const out = [];
  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.startsWith('```')) {
      inFence = !inFence;
      continue;
    }
    if (!inFence) out.push(line);
  }
  return out.join('\n');
}

function isExternalHref(href) {
  const s = String(href || '').trim();
  return (
    s.startsWith('http://') ||
    s.startsWith('https://') ||
    s.startsWith('mailto:') ||
    s.startsWith('tel:')
  );
}

function normaliseTarget(target) {
  // drop query/fragment for filesystem checks
  const s = String(target || '').trim();
  const hashIdx = s.indexOf('#');
  const qIdx = s.indexOf('?');
  let cut = s;
  if (qIdx !== -1) cut = cut.slice(0, qIdx);
  if (hashIdx !== -1) cut = cut.slice(0, hashIdx);
  return cut;
}

function collectMarkdownLinks(mdText) {
  // Handles inline links/images: [text](target) and ![alt](target)
  // Does not attempt full CommonMark parsing.
  const links = [];
  const re = /!?(\[[^\]]*\])\(([^)]+)\)/g;
  let m;
  while ((m = re.exec(mdText)) !== null) {
    const rawTarget = (m[2] || '').trim();
    // strip surrounding angle brackets: <path>
    const t = rawTarget.startsWith('<') && rawTarget.endsWith('>') ? rawTarget.slice(1, -1).trim() : rawTarget;
    links.push(t);
  }
  return links;
}

function main() {
  const mdFiles = listFilesRecursive(ROOT).filter(p => /\.md$/i.test(p));
  const errors = [];

  for (const absMd of mdFiles) {
    const relMd = toRel(absMd);
    const raw = fs.readFileSync(absMd, 'utf8');
    const md = stripFencedCodeBlocks(raw);
    const links = collectMarkdownLinks(md);

    for (const href of links) {
      const s = String(href || '').trim();
      if (!s) continue;
      if (s.startsWith('#')) continue; // pure anchor
      if (isExternalHref(s)) continue;

      // Ignore other schemes
      if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(s)) continue;

      const target = normaliseTarget(s);
      if (!target) continue;

      const absTarget = path.resolve(path.dirname(absMd), target);
      if (!fs.existsSync(absTarget)) {
        errors.push(relMd + ': broken link -> ' + s);
      }
    }
  }

  if (errors.length) {
    process.stderr.write('Internal link check failed:\n\n' + errors.map(e => '- ' + e).join('\n') + '\n');
    process.exit(1);
  }

  process.stdout.write('Internal link check: OK\n');
}

main();
