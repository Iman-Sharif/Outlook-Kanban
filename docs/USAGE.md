# Usage

[Start here](../START_HERE.html) | [Setup](SETUP.md) | [Themes](THEMES.md) | [Docs](README.md)

## Concepts

- Project: an Outlook Tasks folder
- Lane: stored on a task using an Outlook user property (`KFOLaneId`; legacy: `KFO_LaneId`)
- View: a saved combination of mailbox/project + filters (and optional lane layout)

## Board

- By default, the board uses your Outlook Tasks folder.
- Pick a Project from the header if you want to work within a specific Tasks folder.
- Use "New task" in the header to quick add a task in the current folder (default Tasks by default).
- Drag a task card between lanes.
- Click a card to open the task drawer (details, notes, checklist, and quick actions).
- Double-click a card to open it in Outlook.
- Snooze: use the drawer Snooze buttons (or Card Actions snooze presets) to set the due date relative to today.
- Selection: use the checkbox on cards; Shift+click selects a range within the same lane.
- Use the top-bar Select (checkmark) menu for actions like Select all visible / Select all in lane / Invert selection.
- After a lane move, you can Undo from the toast for a short time.
- Quick add: click the lightning icon on a lane, type a subject, then press Enter.
- Header toggles: switch Light/Dark mode and Compact/Comfortable mode.

> [!NOTE]
> If "Sync Outlook status on move" is enabled and a lane maps to an Outlook Status, dragging a task can update the Outlook task Status.

## Filters

- Search: filters by subject and notes preview
- Category: show tasks in a category (or "No Category")
- Due: Any / Overdue / Today / Next 7 days / No due date
- Privacy: show All / Private only / Not Private only
- On narrow panes, the filter dropdowns collapse into a single Filters button.
- Active filters appear as chips under the header (click a chip to clear it).

> [!TIP]
> Drag-and-drop is disabled while filters are active to avoid saving partial ordering.

## Cards (Compact vs Comfortable)

- Comfortable mode can show title, due date, priority, notes preview, and categories (configurable in Settings).
- Compact mode shows only key details (title, due date, priority) to reduce noise; footer action buttons remain.
- Notes preview length is configurable (characters) and only applies in Comfortable mode.
- Optional: enable Settings -> Appearance -> Highlight search to highlight matches in card titles and notes.

> [!TIP]
> If manual ordering is enabled, you can reorder a task without drag-and-drop from Card Actions (Move up/down/top/bottom).

## Notes + checklist

- Notes are the Outlook task Body.
- You can edit notes directly in the task drawer.
- The checklist is stored in the task body as Markdown-style checkboxes (for example `- [ ] Call supplier`). Ticking a checkbox updates the underlying text.

## Views

Use Views to save and quickly re-apply a working context (project + filters).

- Save a view from Settings -> Views.
- Pinned views stay at the top of the picker.

## Settings

- Appearance: theme, density, motion, lane width, card fields, notes preview length (Comfortable mode)
- Board: manual ordering, drag handle only, remember filters/project, date format, categories colour options, keyboard shortcuts (opt-in)
- Projects: create/link/hide/rename projects
- Lanes: create lanes, reorder them, set colours, WIP limits, optional Outlook Status sync
- Lane ids: stored on tasks as `KFOLaneId` (legacy: `KFO_LaneId`) (use Change... if you need to rename an id and migrate tasks)
- Tools: migrate lanes, move tasks between projects, export/import settings, diagnostics

## Tools

- Migrate lanes: bulk-assign lanes based on existing Outlook Task Status
- Move between projects: move tasks between Outlook folders while keeping lane metadata
- Export / import settings: save and restore config + view state as JSON

## Projects (folders)

Projects are Outlook Tasks folders. You can use the board without creating any projects: the default Outlook Tasks folder is always available.

Recommended structure:

- Create a dedicated root folder under Tasks (default: `Kanban Projects`)
- Create one folder per project

Common actions:

- Create a project: Settings -> Projects -> Create
- Link an existing folder: Settings -> Projects -> Link existing (no tasks moved)
- Hide/show projects: removes from the header picker but does not change Outlook
- Move tasks between projects: Settings -> Tools -> Move between projects (keeps lane metadata)

## Migration (from Outlook Status)

If you used a JanBan/Taskboard-style setup that relied on Outlook Status (or multiple folders), you can bulk-assign lanes by mapping Outlook Status values to lane ids.

Recommended flow:

1) Link your existing task folder as a Project.
2) Open Settings -> Tools -> Migrate lanes.
3) Map Outlook Status values to your lanes.
4) Run migration.

Scope options:

- Only tasks without a lane: avoids overwriting tasks that you already organised.
- Treat unknown lanes as unassigned: helps clean up tasks that reference a non-existent lane id.

## Getting help

If something breaks, include:

- Windows version
- Outlook version (classic Outlook)
- whether Folder Home Pages are allowed by your organisation
- diagnostics output from the app (Settings -> Diagnostics)

Related:

- Setup / troubleshooting: [`docs/SETUP.md`](SETUP.md)
- Themes: [`docs/THEMES.md`](THEMES.md)
