# Contributing

Thanks for helping improve Kanban for Outlook.

This project runs inside classic Outlook for Windows as a Folder Home Page (legacy IE engine) and integrates with Outlook via COM/MAPI. Changes should be tested with real Outlook.

## Project principles

- Local-only by default (no external downloads, update checks, or telemetry)
- Keep the classic Outlook Folder Home Page experience working
- Prefer small, safe changes with clear rationale
- Documentation is written in UK English

## Development workflow

1) Create a branch.
2) Make the change.
3) Test in Outlook (manual steps).
4) Open a PR with:

- what you changed
- why you changed it
- how you tested it (Outlook version / Windows)

## Maintainer checks (optional)

These checks are for maintainers and contributors. They are not required to run the app.

Requirements:

- Node.js 18+

Run:

```bash
npm run check
```

This currently runs:

- a local-only runtime audit (prevents accidental network dependencies)
- an internal Markdown link check (relative links)

## Maintainer notes

See [`docs/MAINTAINERS.md`](docs/MAINTAINERS.md) for repository layout, packaging, and releasing notes.

## Reporting bugs

If something breaks, include:

- Windows version
- Outlook version (classic Outlook)
- whether Folder Home Pages are allowed by your organisation
- diagnostics output from the app (Settings -> Diagnostics)
