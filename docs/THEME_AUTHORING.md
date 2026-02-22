# Theme Authoring

[README](../README.md) | [Setup](SETUP.md) | [Usage](USAGE.md) | [Projects](PROJECTS.md) | [Themes](THEMES.md) | [Troubleshooting](TROUBLESHOOTING.md) | [Privacy](../PRIVACY.md) | [Security](../SECURITY.md)

KFO themes are plain CSS files that override the base UI styles in [`css/kfo.css`](../css/kfo.css).

> [!IMPORTANT]
> This app runs inside classic Outlook for Windows using the legacy IE engine. Write CSS that degrades gracefully and keeps text contrast high.

## Theme types

- Built-in: shipped in [`themes/kfo-light/`](../themes/kfo-light/) and [`themes/kfo-dark/`](../themes/kfo-dark/).
- Folder theme (recommended): you create a folder under [`themes/`](../themes/) and point the app at a local `theme.css`.
- Imported theme: you import a local `.css` file in Settings; the CSS is stored in Outlook config.

## Start from the skeleton (recommended)

Use the template:

- [`themes/skeleton/`](../themes/skeleton/)

Steps:

1) Copy `themes/skeleton/` to `themes/<your-id>/`.
2) Edit `themes/<your-id>/theme.css`.
3) Update the selector prefix everywhere:

```css
.kfo.theme-my-theme
```

to:

```css
.kfo.theme-<your-id>
```

4) Open `themes/<your-id>/preview.html` to iterate locally.
5) Add it in-app: Settings -> Appearance -> Folder theme.

## Theme id: how it maps to CSS

The app sets a class on the root element:

```text
<body class="kfo theme-<id> ...">
```

So your theme must scope under:

```css
.kfo.theme-<id> { ... }
.kfo.theme-<id> .kfo-topbar { ... }
```

The app sanitises ids to lowercase `a-z`, `0-9`, and `-`. Use simple ids like `sandstone`, `paper-light`, `midnight-ink`.

## Local-only expectations

> [!NOTE]
> Imported themes are rejected if they contain `http://`, `https://`, `@import`, or IE scriptable CSS (`expression(`, `behavior:`). Folder themes are not automatically scanned, but should follow the same rules.

Practical guidelines:

- Avoid `@import` entirely.
- Avoid remote URLs in `url(...)`.
- If you use `url(...)`, point to a local file you ship with the theme folder.

## What you can style (high-signal selector list)

Core surfaces:

- `.kfo` (root), `.kfo-app`
- `.kfo-topbar`, `.kfo-input`, `.kfo-select`, `.kfo-btn`, `.kfo-iconBtn`, `.kfo-pill`
- `.kfo-board`, `.kfo-lane`, `.kfo-laneHeader`, `.kfo-laneTitle`, `.kfo-laneSub`
- `.kfo-task`, `.kfo-taskHeader`, `.kfo-taskMeta`, `.kfo-taskBody`, `.kfo-taskFooter`
- `.kfo-tag`

Overlays:

- `.kfo-modalBackdrop`, `.kfo-modal`
- `.kfo-toastWrap`, `.kfo-toast`, `.kfo-toast--success`, `.kfo-toast--info`, `.kfo-toast--error`

Settings UI:

- `.kfo-view`, `.kfo-card`, `.kfo-cardTitle`
- `.kfo-rowLabel`, `.kfo-small`, `.kfo-divider`
- `.kfo-listItem`, `.kfo-listItemTitle`, `.kfo-listItemMeta`

You can see representative markup in [`themes/skeleton/preview.html`](../themes/skeleton/preview.html).

## Compatibility + performance tips

- Prefer simple selectors (avoid overly deep nesting).
- Avoid CSS custom properties (variables): IE11 does not support them.
- Use shadows sparingly; large blurred shadows can feel heavy inside Outlook.
- Test both densities: `density-comfortable` and `density-compact`.
- Test all motion modes: `motion-full`, `motion-subtle`, `motion-off`.

## Theme quality checklist

- Scoped under `.kfo.theme-<id>` (does not leak into Outlook UI)
- Readable contrast for text, inputs, pills, and cards
- Modal backdrop + toast are legible and not blinding
- Works on narrow windows (Outlook panes)
- No remote resources; no `@import`
