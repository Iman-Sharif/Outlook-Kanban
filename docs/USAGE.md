# Usage

[README](../README.md) | [Docs index](README.md) | [Setup](SETUP.md) | [Projects](PROJECTS.md) | [Themes](THEMES.md) | [Troubleshooting](TROUBLESHOOTING.md)

## Concepts

- Project: an Outlook Tasks folder
- Lane: stored on a task using an Outlook user property (`KFO_LaneId`)

## Board

- By default, the board uses your Outlook Tasks folder.
- Pick a Project from the header if you want to work within a specific Tasks folder.
- Drag a task card between lanes.
- Double-click a card to open it in Outlook.
- Quick add: click the lightning icon on a lane, type a subject, then press Enter.

> [!NOTE]
> If "Sync Outlook status on move" is enabled and a lane maps to an Outlook Status, dragging a task can update the Outlook task Status.

## Filters

- Search: filters by subject and notes preview
- Category: show tasks in a category (or "No Category")
- Privacy: show All / Private only / Not Private only

> [!TIP]
> Drag-and-drop is disabled while filters are active to avoid saving partial ordering.

## Settings

- Appearance: theme, density, motion, lane width, card fields
- Board: manual ordering, remember filters/project, note preview length, date format, keyboard shortcuts (opt-in)
- Projects: create/link/hide/rename projects
- Lanes: create lanes, reorder them, set colours, optional Outlook Status sync
- Settings transfer: export/import config + view state as JSON (local file)

## Tools

- Migrate lanes: bulk-assign lanes based on existing Outlook Task Status
- Move between projects: move tasks between Outlook folders while keeping lane metadata
- Export / import settings: save and restore config + view state as JSON

Related docs:

- [`docs/PROJECTS.md`](PROJECTS.md)
- [`docs/MIGRATION.md`](MIGRATION.md)
