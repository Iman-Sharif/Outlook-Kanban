# Architecture

[README](../README.md) | [Docs index](README.md) | [Theme Authoring](THEME_AUTHORING.md) | [Security](../SECURITY.md) | [Privacy](../PRIVACY.md)

This project is a static HTML/CSS/JS application that runs inside classic Outlook for Windows as a Folder Home Page.

> [!IMPORTANT]
> The rendering engine is the legacy IE engine. The app integrates with Outlook via COM/MAPI.

## Repository layout

```text
.
├─ kanban.html                 # App entrypoint (Folder Home Page)
├─ js/
│  ├─ app.js                   # AngularJS controller + UI logic
│  ├─ exchange.js              # Outlook COM/MAPI bridge
│  └─ version.js               # App version
├─ css/
│  ├─ kfo.css                  # Base UI components (theme-agnostic)
│  └─ bootstrap.min.css        # Vendor CSS (Glyphicons)
├─ vendor/                     # Vendored JS libs (local, no CDN)
├─ themes/                     # Built-in themes + skeleton template
├─ docs/                       # Documentation
├─ scripts/package.sh          # Release zip builder
└─ install*.cmd / uninstall.cmd# Local install/uninstall helpers
```

## Runtime overview

`kanban.html` loads:

- Vendor libraries from `vendor/` (jQuery, jQuery UI, AngularJS, ui-sortable, Moment, JSON.minify)
- App code from `js/version.js`, `js/exchange.js`, `js/app.js`
- Base styles from `css/kfo.css` + a theme from `themes/.../theme.css`

The AngularJS controller (`taskboardController`) is declared in `js/app.js`.

## Outlook integration

All Outlook access is local and happens through COM:

- `js/exchange.js` detects the Outlook environment (`window.external.OutlookApplication` or ActiveX fallback)
- Tasks and folders are read/written via `outlookNS` (MAPI namespace)

Key concepts:

- Projects = Outlook Task folders
- Lanes are stored per task using Outlook user properties:
  - `KFO_LaneId`
  - `KFO_LaneOrder` (optional manual ordering)

## Local persistence

The app stores configuration/state in the user’s Outlook profile (Journal folder):

- `KanbanConfig` (JSON)
- `KanbanState` (JSON)
- `KanbanErrorLog` (optional; only written if debug logging is enabled)

This keeps the installation portable: replacing files upgrades the UI without migrating a separate database.

## Theme system

Themes are plain CSS overrides.

- Base UI components live in `css/kfo.css`.
- Themes scope under `.kfo.theme-<id>` and override surfaces, borders, and component styling.
- Built-ins live in `themes/kfo-light/` and `themes/kfo-dark/`.
- A creator template lives in `themes/skeleton/`.

See `docs/THEME_AUTHORING.md`.

## Packaging

Releases are shipped as a portable zip.

- Build: `scripts/package.sh`
- CI workflow: `.github/workflows/package.yml`

The zip contains only static assets and scripts (no runtime downloads).
