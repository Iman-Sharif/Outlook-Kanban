# AGENTS.md

## Project at a glance
- App: **Kanban for Outlook** (local-only Outlook Folder Home Page app)
- Runtime target: **classic Outlook on Windows**
- Maintainer tooling: **Node.js 18+** and shell scripts
- No build step for app runtime assets; release is a staged zip package

## Real repository structure
- `kanban.html` - main app entrypoint loaded by Outlook
- `js/` - app logic
  - `js/app/` UI/controller modules
  - `js/board/` pure board logic (unit tested)
  - `js/core/` shared utilities (unit tested)
  - `js/exchange.js`, `js/outlook/` Outlook integration boundary
  - `js/version.js` app version source
- `css/`, `fonts/`, `vendor/` - local runtime assets (no CDN)
- `themes/` - built-in themes + skeleton
- `docs/` - markdown source + shipped offline HTML docs
- `tests/` - Node test runner tests (`node --test`)
- `tools/` - repo checks (`audit-local-only`, link checker)
- `scripts/package.sh` - release zip packager
- `install.cmd` - Windows install/upgrade/repair/uninstall menu
- `dist/` - generated release artifacts

## Commands (from package.json)
- Install tooling deps:
  - `npm install`
- Run tests:
  - `npm test`
- Run local-only policy check:
  - `npm run check:local-only`
- Run markdown link check:
  - `npm run check:links`
- Run full CI-style check:
  - `npm run check`

## Packaging and release artifacts
- Build release zip on Linux/Fedora/macOS:
  - `bash scripts/package.sh`
- Outputs:
  - `dist/kanban-for-outlook.zip`
  - `dist/kanban-for-outlook.zip.sha256`
- Packaging script stages a clean `kanban-for-outlook/` folder and includes only runtime/offline-doc assets.

## Windows install behavior (for end users)
- `install.cmd` installs to `%USERPROFILE%\kanban-for-outlook`
- Registers Outlook Folder Home Page via HKCU Office Outlook registry keys
- Menu supports Install/Upgrade, Repair, Uninstall, and opening local docs

## Fedora quick start (maintainer workflow)
1. Install prerequisites:
   - `sudo dnf install -y nodejs npm zip`
2. Verify Node version:
   - `node -v` (must be `>=18`)
3. From repo root:
   - `npm install`
   - `npm run check`
4. Build release package:
   - `bash scripts/package.sh`
5. Verify artifacts:
   - `ls -lh dist/kanban-for-outlook.zip dist/kanban-for-outlook.zip.sha256`
   - `sha256sum -c dist/kanban-for-outlook.zip.sha256`

## Release checklist
1. Update release content:
   - `js/version.js`
   - `CHANGELOG.md`
   - `whatsnew.html`
2. Run quality gate:
   - `npm run check`
3. Build package:
   - `bash scripts/package.sh`
4. Validate package contents:
   - `unzip -l dist/kanban-for-outlook.zip`
   - confirm top-level folder is `kanban-for-outlook/`
5. Validate checksum:
   - `sha256sum -c dist/kanban-for-outlook.zip.sha256`
6. Commit release changes and tag:
   - `git tag -a vX.Y.Z -m "vX.Y.Z"`
   - `git push origin main`
   - `git push origin vX.Y.Z`
7. GitHub Release:
   - attach `kanban-for-outlook.zip` and `.sha256`
   - publish notes aligned with `CHANGELOG.md`
