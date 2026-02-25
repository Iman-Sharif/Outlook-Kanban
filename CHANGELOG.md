# Changelog

All notable changes to this project will be documented in this file.

## 1.3.2

- Replace multiple installer scripts with a single `install.cmd` menu (install/upgrade/repair/uninstall)
- Ship offline HTML docs in the release zip (`docs/index.html`, `docs/setup.html`, etc.) and update in-app Help to use them
- Simplify the release zip contents (remove extra `.cmd` files and Markdown docs under `docs/`)
- Extend the local-only audit to include shipped offline docs pages

## 1.3.1

- Fix release zip packaging to include `docs/ACCESSIBILITY.md`
- Fix focus return after long-running tools (Move tasks, Migration, Lane id tool)
- Fix non-drag manual ordering to persist consistently across filtering
- Add `lang` to shipped HTML pages

## 1.3.0

- Add Snooze quick actions (drawer buttons + card Actions presets) computed from today
- Improve keyboard access: focus task cards with Tab and open drawer with Enter/Space
- Improve focus handling for popovers and dialogs (focus trap + return focus)
- Add non-drag alternatives for manual ordering (Actions: move up/down/top/bottom)
- Add accessibility documentation

## 1.2.0

- Add a Select menu plus Shift+click range selection (within a lane)
- Add a responsive Filters menu with an active-count badge
- Show checklist progress and a notes indicator on cards (Comfortable mode)
- Optional: highlight search matches in card titles and notes
- Improve Settings with search, per-section reset, and theme management (apply/rename/remove)
- Make drawer fields clickable (lane, due, categories, priority, privacy)

## 1.1.0

- Edit task notes (Outlook Body) directly in the task drawer, with Undo
- Add a checklist in the drawer using Markdown-style checkboxes (`- [ ]` / `- [x]`)
- Improve task actions with clearer colour cues (green Complete, red Delete)

## 1.0.0

- Initial release of this repository
- Local-only Kanban board for classic Outlook Tasks (Windows desktop)
- Includes: projects (folders), configurable lanes stored on tasks, drag/drop with undo, filters, views, settings export/import, themes, and compact mode
