# Migration

[README](../README.md) | [Docs index](README.md) | [Usage](USAGE.md) | [Projects](PROJECTS.md) | [Troubleshooting](TROUBLESHOOTING.md)

This fork uses lanes stored on tasks (`KFO_LaneId`).
Older JanBan-based installations may have relied on Outlook Status and/or multiple folders.

> [!IMPORTANT]
> Migration is local-only. It only writes lane metadata onto tasks in the selected project folder.

## Recommended migration flow

1) Link your existing task folder as a Project.
2) Open Settings -> Tools -> Migrate lanes.
3) Map Outlook Status values to your lanes.
4) Run migration.

## Scope options

- Only tasks without a lane: avoids overwriting tasks that you already organised in the new system.
- Treat unknown lanes as unassigned: helps clean up tasks that have a non-existent lane id.

## Notes

- Migration does not move tasks between folders.
- It only writes the lane property locally on each task.
