# Releasing

[README](../README.md) | [Docs index](README.md) | [Changelog](../CHANGELOG.md) | [Signing](SIGNING.md)

This project publishes releases via Git tags.

When you push a tag like `v3.5.0`, GitHub Actions:

- runs maintainer checks (`npm run check`)
- builds `dist/kanban-for-outlook.zip` + `dist/kanban-for-outlook.zip.sha256` using `scripts/package.sh`
- creates a GitHub Release and attaches those files

## Pre-release checklist

1) Update versions and release notes:

- `js/version.js`
- `CHANGELOG.md`
- `whatsnew.html`

2) Run checks locally:

```bash
npm run check
```

3) Build and verify the release zip:

```bash
bash scripts/package.sh
sha256sum -c dist/kanban-for-outlook.zip.sha256
```

4) Commit the changes to `main`.

## Create the release

Create an annotated tag that matches `js/version.js`:

```bash
git tag -a v3.5.0 -m "v3.5.0"
git push origin main
git push origin v3.5.0
```

The release workflow fails if the tag version does not match `js/version.js`.

## Post-release checks

1) Confirm the GitHub Release contains:

- `kanban-for-outlook.zip`
- `kanban-for-outlook.zip.sha256`

2) Verify the checksum on a clean machine:

- Windows (PowerShell):

```powershell
Get-FileHash .\kanban-for-outlook.zip -Algorithm SHA256
```

- Windows (Command Prompt):

```bat
certutil -hashfile kanban-for-outlook.zip SHA256
```

- Linux/macOS:

```bash
sha256sum -c kanban-for-outlook.zip.sha256
```

## Notes

- Do not commit `dist/` artefacts; they are created by CI and attached to releases.
- If you ship a signed installer, follow [`docs/SIGNING.md`](SIGNING.md).
