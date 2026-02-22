# Changelog

[README](README.md) | [Docs index](docs/README.md) | [Roadmap](ROADMAP.md)

All notable changes to this project will be documented in this file.

## 3.5.0

- Add tag-driven GitHub Releases that publish `kanban-for-outlook.zip` + `.sha256`
- Add a maintainer release checklist (`docs/RELEASING.md`)

## 3.4.0

- Harden theme safety checks (block protocol-relative URLs and scriptable URL schemes)
- Validate theme safety at apply-time to defend against edited/corrupted config
- Strengthen local-only audit to catch external URLs in CSS and legacy HTTP clients

## 3.3.0

- Add export/import for settings (config + view state) as local JSON
- Add opt-in keyboard shortcuts (focused on avoiding Outlook conflicts)
- Add lane quick add (create tasks without opening the inspector)

## 3.2.0

- Refactor the AngularJS controller into ES5 modules loaded by script tags
- Add an Outlook adapter wrapper around `js/exchange.js`
- Add Node-only unit tests for pure logic (sorting/filtering/id sanitising/theme safety)

## 3.1.0

- Expand Diagnostics (performance, storage health, environment details)
- Replace disruptive `alert()` errors with consistent toasts + local error details
- Improve messaging when Folder Home Pages are blocked by policy

## 3.0.0

- Rebuilt UI (modern board, settings, setup wizard)
- Local-only hardening (no external links/updates/telemetry in-app)
- Projects as Outlook Tasks folders
- Lanes stored on tasks via `KFO_LaneId` (+ optional ordering via `KFO_LaneOrder`)
- Tools: migrate lanes from Outlook Status, move tasks between projects
- Theme system: built-in light/dark, import CSS, folder themes + skeleton
