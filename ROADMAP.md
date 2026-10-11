# Roadmap

[README](README.md) | [Docs index](docs/README.md) | [Changelog](CHANGELOG.md)

This roadmap expands the product direction for the maintained fork, with a strong focus on:

- UX clarity (especially Settings and first-run experience)
- daily speed (capture, triage, review)
- powerful local-only task management inside classic Outlook
- reliability and maintainability on the legacy host environment

## Product Vision

Make Kanban for Outlook the fastest way to manage personal and team-owned work inside classic Outlook: easy for first-time users, efficient for daily use, and powerful for advanced workflows without sacrificing local-only privacy.

## Constraints and Principles

- Local-only by default (no telemetry, sync service, or remote dependencies)
- Classic Outlook Folder Home Page compatibility remains first-class
- Progressive disclosure: common tasks first, advanced controls on demand
- Safe changes: clear consequences for destructive actions, easy rollback/export
- Performance budget: smooth interaction in large task folders

## Outcomes and Success Metrics

Use these as release gates and quarterly checks:

- Time to first useful board: user can complete setup and move first task in under 2 minutes
- Settings discoverability: user can find and change a setting in under 15 seconds
- Daily throughput: fewer clicks for common operations (capture, move, due-date updates)
- Error recovery: all high-risk actions are reversible or clearly warned before apply
- Stability: no regressions in local-only, host compatibility, and keyboard accessibility

## Delivery Tracks

Roadmap work is delivered across five tracks:

1. Settings and information architecture
2. Board interaction speed and triage
3. Advanced task-management capabilities
4. Performance, reliability, and diagnostics
5. Docs, onboarding, and accessibility

## Phased Delivery Plan

## Phase 1 - Simplify Core UX (Now, 1-2 releases)

Goal: remove confusion in Settings and reduce friction for first-run and common daily actions.

### Deliverables

- Settings redesign with a "Basic" path and an "Advanced" path
  - Basic sections: Appearance, Board, Projects, Lanes, Shortcuts
  - Advanced sections: lane id tools, status sync mapping, theme import details, diagnostics internals
- Settings search upgrade
  - keyword matching plus jump-to result list
  - highlight matching labels within visible section cards
- Sticky Save/Discard bar in Settings
  - visible when `ui.settingsDirty` is true
  - includes "Review changes" summary grouped by section
- Better defaults and safer forms
  - auto-generate lane ids and imported theme ids from names
  - inline validation language rewritten for plain-English clarity
- Better first-run flow
  - setup wizard remains primary path
  - add an obvious "Re-run setup" action in Settings and Help

### User Problems Solved

- "Settings is too dense and hard to scan"
- "I changed something and I am not sure what I changed"
- "Lane ids and advanced options are intimidating"

### Implementation Touchpoints

- `kanban.html` (settings card structure, nav, save actions, setup re-entry)
- `js/app/controller.js` (settings change tracking and per-section diff summary)
- `css/kfo.css` (settings layout hierarchy and sticky action bar)
- `docs/USAGE.md` and `docs/setup.html` (updated settings and onboarding guidance)

### Exit Criteria

- All existing settings remain reachable
- Basic mode hides advanced controls by default
- Save/Discard behavior is deterministic and tested
- `npm run check` passes

## Phase 2 - Accelerate Daily Workflow (Next, 2-3 releases)

Goal: make capturing and triaging tasks dramatically faster for repeat users.

### Deliverables

- Enhanced quick add
  - multi-add input (one task per line)
  - token parsing for quick metadata (examples: due offset, priority, categories, privacy)
- Stronger filter model
  - add lane, priority, has-notes, has-subtasks, and no-due filters
  - saved filter presets for common review contexts
- Review mode
  - one-click focus views: Overdue, Stale, Unassigned lane, No due date
  - targeted bulk actions from each review mode
- Optional temporary sort modes
  - due date, priority, subject, lane age
  - clearly separate from persistent manual order

### User Problems Solved

- "Adding many tasks is slow"
- "I cannot quickly find risky tasks"
- "Board scans are manual and repetitive"

### Implementation Touchpoints

- `kanban.html` (quick add UI, filter controls, review entry points)
- `js/app/controller.js` (token parsing orchestration, review presets, bulk actions)
- `js/board/logic.js` (new filter predicates and temporary sorting behavior)
- `js/core/util.js` (safe token parser utilities)
- `tests/board.test.js` and `tests/util.test.js` (filter/sort/parser coverage)

### Exit Criteria

- Token parsing is deterministic and non-destructive
- Review presets are reversible and do not mutate data unless user confirms actions
- Drag ordering semantics stay intact when sort mode is disabled
- `npm run check` passes

## Phase 3 - Add Power Features (Later, 3-5 releases)

Goal: elevate from a board UI to a high-capability personal task-management tool.

### Deliverables

- Task templates
  - template includes lane, categories, due offset, checklist seed, privacy, priority
  - create task from template in header and lane quick add flow
- Planner views (local)
  - board remains default
  - add optional list/agenda view for due-date planning and batch edits
- Workflow automation (opt-in)
  - simple local rules engine for lane transition behaviors
  - examples: move to Waiting sets due +7d, move to Done marks complete
- Lightweight analytics
  - lane age and cycle-time summaries using local timestamps
  - local export (JSON/CSV) for external analysis

### User Problems Solved

- "I repeat the same task setup over and over"
- "I need planning views beyond lanes"
- "I want insight into flow and bottlenecks"

### Implementation Touchpoints

- `kanban.html` (template controls, view switcher, automation/rules settings)
- `js/app/controller.js` (template CRUD, view-state management, rules execution)
- `js/board/logic.js` (cycle-time computation and summary helpers)
- `js/exchange.js` and `js/outlook/adapter.js` (safe field reads/writes needed for rules)
- `docs/USAGE.md`, `docs/accessibility.html`, `docs/MAINTAINERS.md`
- new tests around templates, rules, analytics invariants

### Exit Criteria

- Templates and automations are explicitly opt-in
- Analytics export excludes sensitive system internals and remains local
- Keyboard flow remains functional in all major interactions
- `npm run check` passes

## Phase 4 - Scale and Reliability (Ongoing)

Goal: keep the experience responsive and stable as data size and feature depth grow.

### Deliverables

- Incremental loading strategy for large folders
- Deferred rendering for lane cards beyond viewport
- Local profiling mode in diagnostics
- Error log quality improvements (faster triage, clearer context)
- Regression suite growth for board logic and settings state transitions

### User Problems Solved

- "Board slows down with many tasks"
- "Hard to debug issues in host-specific environments"

### Implementation Touchpoints

- `js/app/controller.js` (refresh orchestration, diagnostics payload)
- `js/board/logic.js` (data shaping for progressive rendering)
- `tools/` scripts and `tests/` additions
- docs: troubleshooting and known-limit guidance

### Exit Criteria

- Acceptable interaction latency at target folder sizes
- Diagnostics reports actionable root-cause hints
- No local-only policy violations
- `npm run check` passes

## UX Backlog for Settings (Priority Queue)

This queue is intentionally specific because Settings is currently the highest UX leverage area.

1. Basic vs Advanced settings mode toggle
2. Settings result list with section jump links
3. Sticky Save/Discard + change summary
4. Auto-generated lane id and theme id
5. Better language for warnings and reset confirmations
6. One-click "Re-run setup wizard"
7. "Common tasks" shortcuts panel in Settings home

## Feature Backlog for Power Users (Priority Queue)

1. Multi-add quick capture
2. Review mode presets (Overdue/Stale/No due/Unassigned)
3. Template library and "Create from template"
4. Advanced filters (priority, lane, subtasks/notes)
5. Planner view (list/agenda)
6. Opt-in local automation rules
7. Flow analytics and local export

## Risks and Mitigations

- Host constraints (legacy IE engine) may limit modern UI behavior
  - mitigate with conservative DOM patterns and graceful fallback
- Feature growth can increase settings complexity
  - mitigate with progressive disclosure and strong defaults
- Performance regressions from richer filtering and analytics
  - mitigate with profiling mode and incremental processing
- Data integrity risks in automation/rules
  - mitigate with dry-run preview and explicit confirmation for bulk changes

## Validation and Release Gates

Before each release candidate:

- `npm run check` passes (`check:local-only`, link checks, tests)
- Manual keyboard/accessibility pass for core actions
- Compatibility spot-check on supported Outlook/Windows targets
- Updated docs for any changed user workflow
- Changelog entry includes migration notes for behavior changes

## Not in Scope (for this app track)

- Cloud sync or collaboration services
- Telemetry-backed usage analytics
- New Outlook, web Outlook, or macOS runtime support
- Replacing Outlook as source of truth for tasks

## Suggested Execution Order (Immediate)

1. Phase 1 settings simplification foundation
2. Phase 2 capture and review acceleration
3. Phase 3 templates and planner views
4. Phase 4 scale hardening as a continuous stream

This order gives the fastest user-visible value while reducing support burden before adding deeper power features.
