# Privacy

[README](README.md) | [Docs index](docs/README.md) | [Setup](docs/SETUP.md) | [Security](SECURITY.md)

Kanban for Outlook is designed to run fully locally inside classic Outlook for Windows.

## What data it accesses

- Outlook Tasks from the currently selected project folder
- Outlook Categories (for optional colour display)

## Where it stores data

All data stays in your local Outlook profile:

- Configuration and UI state are stored as Outlook items in your Journal folder (subjects like `KanbanConfig` and `KanbanState`).
- Lane assignment is stored on each task using an Outlook user property (`KFOLaneId`; legacy: `KFO_LaneId`).
- Optional manual ordering is stored on each task using `KFOLaneOrder` (legacy: `KFO_LaneOrder`).

## Network access

- No external downloads
- No update checks
- No telemetry

All JS/CSS dependencies are bundled locally in this repository.

## Themes

Themes are local CSS.
Theme imports are rejected if they contain `http://`, `https://`, `@import`, or IE scriptable CSS (`expression(`, `behavior:`).

## Diagnostics

The in-app Diagnostics view shows what the app stores and uses (config/state/log). Nothing is sent anywhere automatically.
Only share diagnostics if you explicitly choose to copy/paste it.
