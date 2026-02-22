# Windows SmartScreen

[README](../README.md) | [Docs index](README.md) | [Setup](SETUP.md) | [Signing](SIGNING.md)

Some users will see a warning when running `install.cmd` / `install-local.cmd`:

"Windows protected your PC"

This is Windows Defender SmartScreen.

## Why it happens

SmartScreen is reputation-based. New binaries and scripts (especially from zips) often show this warning until they build trust signals.

## Options for users

1) Continue
   - Click "More info" -> "Run anyway"

2) Install manually (no scripts)
   - See [`docs/SETUP.md`](SETUP.md)

3) Verify
   - Only download from the official GitHub Releases.
   - Verify checksums if the release provides them.

Checksum commands on Windows:

- PowerShell: `Get-FileHash .\kanban-for-outlook.zip -Algorithm SHA256`
- Command Prompt: `certutil -hashfile kanban-for-outlook.zip SHA256`

## Reducing warnings for future releases

The most effective mitigation is to ship a signed Windows installer (`.exe` / `.msi`):

- Code signing certificate (OV or EV)
- Sign the installer during release
- Publisher name will show as the certificate subject

Note: signing a `.cmd` file does not reliably remove SmartScreen prompts.
