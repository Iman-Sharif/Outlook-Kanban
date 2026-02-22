# Projects

[README](../README.md) | [Docs index](README.md) | [Setup](SETUP.md) | [Usage](USAGE.md) | [Migration](MIGRATION.md)

Projects are Outlook Tasks folders.

> [!NOTE]
> Projects are purely a UI grouping. Your tasks remain in Outlook folders; nothing is uploaded anywhere.

## Recommended structure

- Create a dedicated root folder under Tasks (default: `Kanban Projects`)
- Create one folder per project

## Create a project

Settings -> Projects -> Create

This creates a Tasks folder under the configured projects root.

## Link an existing folder

Settings -> Projects -> Link existing

This adds an existing Tasks folder to the project picker (no tasks are moved).

## Hide/show

Hidden projects are removed from the header picker but remain unchanged in Outlook.

## Move tasks between projects

Settings -> Tools -> Move between projects.

The app moves tasks to another Outlook folder and keeps lane metadata (`KFO_LaneId`, `KFO_LaneOrder`).

Related:

- [`docs/USAGE.md`](USAGE.md)
