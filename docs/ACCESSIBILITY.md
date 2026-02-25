# Accessibility

Kanban for Outlook runs inside classic Outlook's Folder Home Page (legacy IE engine). Some accessibility features depend on what the host exposes, but the app aims to be usable with keyboard and assistive tech.

## Keyboard

- Task cards are focusable: use `Tab` to focus a card.
- Each card is a single tab stop (card buttons are primarily for mouse use).
- `Enter` opens the task drawer.
- `Space` toggles selection for the focused task.
- `↑ ↓ ← →` (or `h j k l`) moves focus between tasks and lanes.
- `a` opens Actions, `c` completes, `o` opens in Outlook, `m` moves to a lane.
- `Delete` deletes the focused task (confirm).
- `n` opens OneNote for the focused task (if available).
- Buttons and inputs throughout the UI are standard focusable controls.
- `Esc` closes open popovers/dialogs (and the task drawer).

Keyboard shortcuts (like `/` for search) are opt-in (Settings -> Board -> Shortcuts).

## Dragging Alternatives

- Moving tasks between lanes: use the Card Actions menu (or bulk actions) to "Move to lane...".
- Reordering within a lane (when manual ordering is enabled): use Card Actions -> Move up/down/top/bottom.

## Focus

- Popovers and dialogs keep focus inside while open.
- When a popover/dialog closes, focus returns to the control that opened it (best-effort).
