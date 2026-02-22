# Theme Skeleton (Copy This Folder)

This folder is a starter template for creating your own KFO theme.

Quick links:

- Theme authoring guide: [`docs/THEME_AUTHORING.md`](../../docs/THEME_AUTHORING.md)
- End-user themes doc: [`docs/THEMES.md`](../../docs/THEMES.md)

## Create a new theme (recommended: Folder theme)

1) Copy `themes/skeleton/` to `themes/<your-id>/`.
2) Edit `themes/<your-id>/theme.css`.
3) Replace the selector prefix everywhere:

```css
.kfo.theme-my-theme
```

with:

```css
.kfo.theme-<your-id>
```

4) In the app: Settings -> Appearance -> Folder theme

- Name: your theme name
- Id: `<your-id>`
- CSS path: `themes/<your-id>/theme.css`

> [!TIP]
> Keep the theme id consistent with the folder name. The preview page auto-detects the id from the folder name.

## Preview without Outlook

Open `themes/<your-id>/preview.html` in a browser.

This page renders representative KFO markup (board, settings cards, modals, toast) so you can iterate on CSS without needing Outlook.

## Local-only expectations

Keep theme assets local.

- Do not use `@import`.
- Do not reference remote URLs (no `http://` / `https://`).

Imported themes are rejected if they contain those patterns, but folder themes should follow the same rule to preserve the "local-only" design.
