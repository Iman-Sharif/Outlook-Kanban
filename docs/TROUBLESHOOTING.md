# Troubleshooting

[README](../README.md) | [Docs index](README.md) | [Setup](SETUP.md) | [Usage](USAGE.md) | [SmartScreen](SMARTSCREEN.md)

## "Sorry, this app can only be run as a Home Page"

- Verify the folder Home Page is set to the local `kanban.html`.
- Some environments block Folder Home Pages via policy.

> [!TIP]
> If your org disables Folder Home Pages, the app cannot run. Manual install will not help in that case.

## Blank page / script errors

- Ensure you're using classic Outlook for Windows.
- Some security settings can block ActiveX / scripting in Folder Home Pages.

## Drag-and-drop not working

- Drag-and-drop is disabled while filters are active (search/category/privacy) to avoid saving partial ordering.
- Clear filters and try again.

## Performance

- Large task folders can be slow to load.
- Consider splitting into multiple Projects (folders).
- Set Motion to "Subtle" or "Off".
