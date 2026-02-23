# Themes

[Start here](../START_HERE.html) | [Setup](SETUP.md) | [Usage](USAGE.md) | [Docs](README.md)

Themes are local CSS overrides.

## Built-in themes

- Light mode
- Dark mode

## Theme types

- Built-in: shipped in `themes/kfo-light/` and `themes/kfo-dark/`.
- Folder theme (recommended): a `themes/<id>/theme.css` file you ship locally.
- Imported theme: you import a local `.css` file in Settings; the CSS is stored in your Outlook config.

## Folder themes (recommended for power users)

1) Copy [`themes/skeleton/`](../themes/skeleton/) to `themes/<your-theme>/`
2) Edit `themes/<your-theme>/theme.css`
3) In Settings -> Appearance -> Folder theme:
   - Name: your theme name
   - Id: a simple id (letters/numbers/dash)
   - CSS path: `themes/<your-theme>/theme.css`

Tip: open `themes/<your-theme>/preview.html` to iterate on CSS without Outlook.

### Theme id: how it maps to CSS

The app sets a class on the root element:

```text
<body class="kfo theme-<id> ...">
```

So your theme must scope under:

```css
.kfo.theme-<id> { ... }
.kfo.theme-<id> .kfo-topbar { ... }
```

Use simple ids like `sandstone`, `paper-light`, `midnight-ink`.

## Import a CSS file

Settings -> Appearance -> Import theme.

Security guardrails:

- Theme imports are rejected if they contain `http://`, `https://`, protocol-relative `//`, `@import`, `javascript:`/`vbscript:` URLs, or IE scriptable CSS (`expression(`, `behavior:`)

This keeps the app local-only.

## Authoring tips (classic Outlook / legacy IE)

This app runs inside classic Outlook for Windows using the legacy IE engine.

- Avoid CSS custom properties (variables): IE11 does not support them.
- Prefer simple selectors (deep nesting can be slow).
- Keep contrast high for text, pills, and inputs.
- Test both densities (`density-comfortable`, `density-compact`) and motion modes.

High-signal selectors:

- Top bar: `.kfo-topbar`, `.kfo-input`, `.kfo-select`, `.kfo-btn`, `.kfo-iconBtn`, `.kfo-pill`
- Board: `.kfo-lane`, `.kfo-laneHeader`, `.kfo-laneTitle`
- Cards: `.kfo-task`, `.kfo-taskMeta`, `.kfo-taskBody`, `.kfo-taskFooter`, `.kfo-tag`
- Overlays: `.kfo-modalBackdrop`, `.kfo-modal`, `.kfo-toast`
