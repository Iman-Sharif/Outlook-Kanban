# Themes

[README](../README.md) | [Setup](SETUP.md) | [Usage](USAGE.md) | [Projects](PROJECTS.md) | [Theme Authoring](THEME_AUTHORING.md) | [Troubleshooting](TROUBLESHOOTING.md)

Themes are local CSS.

## Built-in themes

- Professional Light
- Professional Dark

## Folder themes (recommended for power users)

1) Copy [`themes/skeleton/`](../themes/skeleton/) to `themes/<your-theme>/`
2) Edit `themes/<your-theme>/theme.css`
3) In Settings -> Appearance -> Folder theme:
   - Name: your theme name
   - Id: a simple id (letters/numbers/dash)
   - CSS path: `themes/<your-theme>/theme.css`

Tip: open `themes/<your-theme>/preview.html` to iterate on CSS without Outlook.

For creators, see [`docs/THEME_AUTHORING.md`](THEME_AUTHORING.md).

## Import a CSS file

Settings -> Appearance -> Import theme.

Security guardrails:

- Theme imports are rejected if they contain `http://`, `https://`, `@import`, or IE scriptable CSS (`expression(`, `behavior:`)

This keeps the app local-only.
