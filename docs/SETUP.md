# Setup (Local)

Kanban for Outlook is designed to run fully locally as an Outlook Folder Home Page.

[Start here](../START_HERE.html) | [Usage](USAGE.md) | [Themes](THEMES.md) | [Docs](README.md)

## Requirements

- Windows + classic Outlook desktop
- Folder Home Page must be enabled by your policy (some orgs disable it)

> [!NOTE]
> This app is not compatible with the new Outlook or Outlook on the web.

## Install (recommended)

1) Download the release zip:

- Download: [kanban-for-outlook.zip](https://github.com/Iman-Sharif/Outlook-Kanban/releases/latest/download/kanban-for-outlook.zip)
- Checksum (optional): [kanban-for-outlook.zip.sha256](https://github.com/Iman-Sharif/Outlook-Kanban/releases/latest/download/kanban-for-outlook.zip.sha256)

> [!IMPORTANT]
> For installation, do not use GitHub's green `Code` button -> `Download ZIP`, and do not use the Release page's `Source code (zip/tar.gz)` downloads. Those are for developers.

2) Extract the zip (it contains a `kanban-for-outlook` folder).
3) Open that folder and run `install.cmd` (or `install-local.cmd`).
4) Restart Outlook.

The installer copies the app into:

- `%USERPROFILE%\kanban-for-outlook`

and registers `kanban.html` as the Folder Home Page.

If you prefer not to run scripts, use the manual steps below.

## Windows SmartScreen warning

If you see a warning like "Windows protected your PC" when running `install.cmd` / `install-local.cmd`, this is Windows Defender SmartScreen.

### Why it happens

SmartScreen is reputation-based. New scripts and binaries (especially from zips) often show this warning until the download builds trust signals.

### Options

- Continue: click "More info" -> "Run anyway"
- Install manually (no scripts): use the manual install steps below
- Verify integrity: only download from the official GitHub Releases and verify the SHA-256 checksum (if provided)

To verify a checksum on Windows:

- PowerShell: `Get-FileHash .\kanban-for-outlook.zip -Algorithm SHA256`
- Command Prompt: `certutil -hashfile kanban-for-outlook.zip SHA256`

To reduce SmartScreen prompts long-term, releases can be shipped as a signed installer (`.exe` / `.msi`).
Signed installers display a Publisher name and typically trigger fewer warnings once reputation builds.

## Install (manual)

1) Copy the extracted folder to a location you control.
2) In Outlook, right-click the folder you want the board on (typically your top mailbox folder) -> Properties.
3) Home Page tab:
   - Address: select the local `kanban.html`
   - Enable: "Show home page by default for this folder"

> [!TIP]
> If your organisation disables Folder Home Pages via policy, the app cannot run (manual install will not help). See Troubleshooting below.

## Upgrade

Upgrading is a file replacement.

1) Close Outlook.
2) Extract the new release zip to a temporary folder.
3) Run `install.cmd` (safe to re-run), then restart Outlook.

See [`upgrade.html`](../upgrade.html) for the short checklist.

## Uninstall

Run `uninstall.cmd`.

## Notes

- This fork does not use any online setup, update checks, or external assets.
- Configuration is stored locally in Outlook.

## Troubleshooting

### "Sorry, this app can only be run as a Home Page"

- Verify the folder Home Page is set to the local `kanban.html`.
- Some environments block Folder Home Pages via policy.

### Blank page / script errors

- Ensure you're using classic Outlook for Windows.
- Some security settings can block ActiveX / scripting in Folder Home Pages.

### Drag-and-drop not working

- Drag-and-drop is disabled while filters are active (search/category/privacy) to avoid saving partial ordering.
- Clear filters and try again.

### Performance

- Large task folders can be slow to load.
- Consider splitting into multiple Projects (folders).
- Set Motion to "Subtle" or "Off".

## Disclaimer

This project is provided "AS IS" with no warranty. See [`LICENSE`](../LICENSE).
