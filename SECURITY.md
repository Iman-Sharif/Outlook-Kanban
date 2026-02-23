# Security Policy

[README](README.md) | [Docs index](docs/README.md) | [Privacy](PRIVACY.md) | [Third-party notices](THIRD_PARTY_NOTICES.md)

## Supported versions

This repository is community-maintained. Security fixes are applied to the latest release.

## Reporting a vulnerability

Please report security issues via GitHub Security Advisories (preferred) or by opening a GitHub issue if the issue is not sensitive.

## Local-only design

This fork aims to minimise risk by:

- Bundling all dependencies locally (no CDN)
- Avoiding network calls (no update checks / telemetry)
- Storing configuration locally in Outlook

## Themes

Themes are local CSS. The app rejects imported theme files that contain:

- `http://` / `https://`
- `@import`
- IE scriptable CSS (`expression(`, `behavior:`)

This helps prevent accidental external loads and reduces risk in the legacy IE rendering engine.
