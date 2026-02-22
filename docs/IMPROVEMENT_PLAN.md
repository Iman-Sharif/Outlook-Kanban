# Improvement Plan (Next Level)

[README](../README.md) | [Docs index](README.md) | [Architecture](ARCHITECTURE.md) | [Roadmap](../ROADMAP.md)

This document turns the high-level goals in `ROADMAP.md` into an actionable plan.

Assumptions:

- Typical project size: under ~200 tasks
- The app must remain local-only and compatible with classic Outlook for Windows (Folder Home Page / legacy IE engine)

## Objectives

- Reliability: clear errors, strong diagnostics, graceful handling of Outlook policy restrictions
- Maintainability: smaller change surface, clearer boundaries between UI logic and Outlook COM access
- Trust: predictable releases with checksums and strong local-only guarantees
- Daily usability: reduce friction for personal task management inside Outlook

## Phase 1: Reliability + Supportability (v3.1)

Deliverables:

- Expand in-app Diagnostics with:
  - refresh timings (task scan, lane build, filtering)
  - storage health checks (config/state read-write)
  - environment details (Outlook version)
- Replace disruptive `alert()` failures with consistent UI messaging (toast + error details)
- Improve messaging when Folder Home Pages are blocked by policy (with direct links to `docs/TROUBLESHOOTING.md`)

Success criteria:

- Most support requests can be answered with a single Diagnostics copy/paste
- No silent failures when Outlook storage is unavailable

## Phase 2: Maintainability without changing runtime (v3.2)

Deliverables:

- Split the AngularJS controller (previously `js/app.js`) into ES5-compatible modules loaded by script tags
- Introduce an "Outlook adapter" wrapper around `js/exchange.js` so COM exceptions are handled in one place
- Add basic unit tests for pure functions (sorting, filtering, id sanitising, theme safety checks) using Node only

Success criteria:

- Core UI logic can be tested without Outlook
- Outlook COM calls have clear boundaries and consistent error handling

## Phase 3: Daily QoL features (v3.3)

Deliverables:

- Export/import settings (config + UI state) as JSON (local file)
- Optional keyboard shortcuts (opt-in; avoid clashes with Outlook)
- Faster "quick add" flow for creating tasks into a lane

Success criteria:

- Moving to a new machine/profile is straightforward (export/import)

## Phase 4: Local-only guarantees (v3.4)

Deliverables:

- Harden theme safety checks (imported and folder themes)
- Add CI checks that prevent accidental runtime network dependencies

Notes:

- Documentation can contain external links; the local-only guarantee applies to runtime assets (`kanban.html`, `js/`, `css/`, `themes/`).

## Phase 5: Release process + trust signals (v3.5)

Deliverables:

- Tag-driven GitHub Releases that publish:
  - `kanban-for-outlook.zip`
  - `kanban-for-outlook.zip.sha256`
- A maintainer release checklist (`docs/RELEASING.md`)

Success criteria:

- Releases are reproducible and verifiable

## Current status

- Maintainer checks are in place (see `CONTRIBUTING.md`):
  - local-only runtime audit
  - internal Markdown link check
