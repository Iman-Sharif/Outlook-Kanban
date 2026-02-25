# Accessibility

Kanban for Outlook runs inside classic Outlook's Folder Home Page (legacy IE engine). Some accessibility features depend on what the host exposes, but the app aims to be usable with keyboard and assistive tech.

## Keyboard

- Task cards are focusable: use `Tab` to focus a card, then `Enter`/`Space` to open the task drawer.
- Buttons and inputs throughout the UI are standard focusable controls.
- `Esc` closes open popovers/dialogs (and the task drawer).

Keyboard shortcuts (like `/` for search) are opt-in (Settings -> Board -> Shortcuts).

## Dragging Alternatives

- Moving tasks between lanes: use the Card Actions menu (or bulk actions) to "Move to lane...".
- Reordering within a lane (when manual ordering is enabled): use Card Actions -> Move up/down/top/bottom.

## Focus

- Popovers and dialogs keep focus inside while open.
- When a popover/dialog closes, focus returns to the control that opened it (best-effort).
