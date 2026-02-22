# Signing (Publisher Name)

[README](../README.md) | [Docs index](README.md) | [SmartScreen](SMARTSCREEN.md)

This document is for maintainers.

Windows Defender SmartScreen warnings are reputation-based and common for new open-source downloads.
The most effective mitigation is shipping a signed Windows installer.

## Why signing helps

When an installer (`.exe` / `.msi`) is signed with a code signing certificate:

- Windows can display a Publisher name (the certificate subject)
- SmartScreen prompts are typically reduced once the signed artifact builds reputation

## Certificate types

- OV (Organization Validation): cheaper, still may show SmartScreen until reputation builds
- EV (Extended Validation): more expensive, tends to build SmartScreen trust faster

If you want SmartScreen to show the Publisher as a person, you will need a certificate issued to that identity.

## Recommended release shape

Offer users both:

1) Portable zip (fully local)
2) Signed installer (best UX)

## Signing workflow (high level)

1) Build installer on Windows (Inno Setup / WiX / MSI)
2) Sign with `signtool.exe` using the code signing certificate
3) (Optional but recommended) Timestamp the signature
4) Publish installer + checksums on GitHub Releases

## Notes

- Signing `.cmd` scripts does not reliably remove SmartScreen prompts.
- Timestamping requires contacting a timestamp server during the build/release process (not at runtime).
