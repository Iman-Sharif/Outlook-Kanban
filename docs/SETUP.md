# Setup (Local)

Kanban for Outlook is designed to run fully locally as an Outlook Folder Home Page.

[README](../README.md) | [Docs index](README.md) | [Usage](USAGE.md) | [Troubleshooting](TROUBLESHOOTING.md) | [SmartScreen](SMARTSCREEN.md)

## Requirements

- Windows + classic Outlook desktop
- Folder Home Page must be enabled by your policy (some orgs disable it)

> [!NOTE]
> This app is not compatible with the new Outlook or Outlook on the web.

## Install (recommended)

1) Download the release zip:

- Download: [kanban-for-outlook.zip](https://github.com/Iman-Sharif/Kanban-for-Outlook/releases/latest/download/kanban-for-outlook.zip)
- Checksum (optional): [kanban-for-outlook.zip.sha256](https://github.com/Iman-Sharif/Kanban-for-Outlook/releases/latest/download/kanban-for-outlook.zip.sha256)

> [!IMPORTANT]
> For installation, do not use GitHub's green `Code` button -> `Download ZIP`, and do not use the Release page's `Source code (zip/tar.gz)` downloads. Those are for developers.

2) Extract the zip.
3) Run `install.cmd` (or `install-local.cmd`).
4) Restart Outlook.

The installer copies the app into:

- `%USERPROFILE%\kanban-for-outlook`

and registers `kanban.html` as the Folder Home Page.

If you prefer not to run scripts, use the manual steps below.

## Windows SmartScreen warning

If you see a warning like "Windows protected your PC" when running `install.cmd` / `install-local.cmd`, this is Windows Defender SmartScreen.
It is reputation-based and common for new open-source tools.

Options:

- Proceed: click "More info" -> "Run anyway"
- Avoid running scripts: use the manual install steps
- Verify integrity: only download from the official GitHub Releases and verify the SHA-256 checksum (if provided)

See [`docs/SMARTSCREEN.md`](SMARTSCREEN.md) for more detail.

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
> If your organisation disables Folder Home Pages via policy, the app cannot run. See [`docs/TROUBLESHOOTING.md`](TROUBLESHOOTING.md).

## Uninstall

Run `uninstall.cmd`.

## Notes

- This fork does not use any online setup, update checks, or external assets.
- Configuration is stored locally in Outlook.

## Disclaimer

This project is provided "AS IS" with no warranty. See [`DISCLAIMER.md`](../DISCLAIMER.md) and [`LICENSE`](../LICENSE).
