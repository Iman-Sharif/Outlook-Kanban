# Roadmap

[README](README.md) | [Docs index](docs/README.md) | [Changelog](CHANGELOG.md)

This file tracks the planned direction for this maintained fork.

## Goals

- Keep the classic Outlook-on-Windows "Folder Home Page" experience working on current Office/Windows releases
- Improve maintainability (structure, documentation, release process)
- Security + privacy first: local-only by default
- Replace or remove dependencies on discontinued/legacy endpoints
- Make support and contribution paths clear (issues, PRs, troubleshooting)

## Milestones

### 0) Project hygiene

- Clear README, credits, and licensing
- Define support channel (GitHub Issues) and contribution flow

### 1) Link + support refresh

- Remove in-app external links and online update mechanics
- Replace legacy support email targets with local diagnostics + GitHub Issues guidance

### 2) Compatibility validation

- Re-test on Windows 10/11 with classic Outlook (2016/2019/2021/Microsoft 365)
- Capture known issues and add reproducible test steps

### 3) Performance + scale

- Improve large-folder performance (optional paging / incremental load)
- Add a lightweight, local-only profiling mode

### 4) Release process

- Move distribution artifacts to GitHub Releases (avoid committing built zip files when possible)
- Add a changelog and versioning policy

### 5) "Next level" track (optional)

- Evaluate a modern Outlook approach (Office Add-in / Graph) as a separate project track
  - Note: this is not a drop-in replacement for the legacy Folder Home Page model
