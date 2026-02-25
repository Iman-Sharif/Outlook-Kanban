# Maintainers

This document is for contributors and maintainers.
End users should start with [`START_HERE.html`](../START_HERE.html) and [`docs/SETUP.md`](SETUP.md).

[Docs](README.md) | [Contributing](../CONTRIBUTING.md) | [Changelog](../CHANGELOG.md)

## Project principles

- Local-only by default (no external downloads, update checks, or telemetry)
- Must run inside classic Outlook for Windows as a Folder Home Page (legacy IE engine)
- Prefer small, safe changes; test with real Outlook

## Repository layout (runtime)

```text
.
├─ kanban.html                  # App entrypoint (Folder Home Page)
├─ js/
│  ├─ app/                      # AngularJS controller split into modules
│  ├─ board/                    # Pure board logic (testable)
│  ├─ core/                     # Pure utilities (testable)
│  ├─ exchange.js               # Outlook COM/MAPI bridge
│  ├─ outlook/                  # Outlook adapter boundary
│  └─ version.js                # App version
├─ css/                         # Base UI components + vendored Bootstrap CSS
├─ vendor/                      # Vendored JS libs (local, no CDN)
├─ themes/                      # Built-in themes + skeleton template
└─ install.cmd                  # Local install/upgrade/repair/uninstall menu
```

Key storage model:

- Config/state: stored as Outlook Journal items (`KanbanConfig`, `KanbanState`)
- Lane metadata: stored on tasks via user properties (`KFOLaneId`, `KFOLaneOrder`, `KFOLaneChangedAt`)

## Maintainer checks

Requirements:

- Node.js 18+

Run:

```bash
npm run check
```

This runs:

- local-only runtime audit (prevents accidental network dependencies)
- internal Markdown link check
- Node-only unit tests for pure logic

## Packaging (release zip)

Build the release zip locally:

```bash
bash scripts/package.sh
```

This creates:

- `dist/kanban-for-outlook.zip`
- `dist/kanban-for-outlook.zip.sha256`

## Releasing

Releases are tag-driven.

1) Update version + notes:

- `js/version.js`
- `CHANGELOG.md`
- `whatsnew.html`

2) Run checks:

```bash
npm run check
```

3) Tag and push:

```bash
git tag -a vX.Y.Z -m "vX.Y.Z"
git push origin main
git push origin vX.Y.Z
```

## Signing notes (SmartScreen)

Windows Defender SmartScreen warnings are reputation-based.
The most effective mitigation is shipping a signed Windows installer (`.exe` / `.msi`) alongside the portable zip.

Notes:

- Signing `.cmd` scripts does not reliably remove SmartScreen prompts.
- Timestamping signatures requires contacting a timestamp server during the build/release process (not at runtime).
