# Code signing policy

Hypercube Office has applied to the [SignPath Foundation](https://signpath.org) for free code
signing of its Windows installer. Until that is approved, releases are unsigned and Windows
SmartScreen shows "Unknown publisher".

Once approved: free code signing provided by [SignPath.io](https://about.signpath.io),
certificate by [SignPath Foundation](https://signpath.org).

## What gets signed

- Only the Windows x64 installer (`*-Setup-<version>-x64.exe`) and the program files built from
  this repository's source code, by the **Windows installer** GitHub Actions workflow
  ([`.github/workflows/release-windows.yml`](.github/workflows/release-windows.yml)) on GitHub's
  hosted runners. Nothing built on a personal computer is ever submitted for signing.
- Every signed file carries the product name "Hypercube Office" and the release version from
  `apps/shell/package.json`.
- Third-party binaries that ship inside the installer unchanged keep their upstream signature,
  or stay unsigned; they are never signed with this project's certificate.

## Team roles

The project currently has one maintainer, who holds every role:

- Committers and reviewers: [omar-184](https://github.com/omar-184)
- Approvers: [omar-184](https://github.com/omar-184)

Changes from anyone else, including AI-assisted pull requests, are merged only after the
maintainer reviews them. Every release is approved by hand before it is signed. Team members use
multi-factor authentication for GitHub and SignPath.

## Privacy

This program will not transfer any information to other networked systems unless specifically
requested by the user or the person installing or operating it. The only network feature that
runs without a click is the update check, which is off by default; when you turn it on it asks
GitHub Releases whether a newer version exists and sends nothing about you or your documents.
See the [privacy statement](PRIVACY.md).

## Uninstalling

Uninstall from Windows Settings > Apps > Installed apps, or run `Uninstall Hypercube Office.exe`
in the install folder. Silent uninstall: `/S`.

## Reporting a problem

To report a signed file that you think is malicious or was not built from this repository, use
[private vulnerability reporting](SECURITY.md) or contact the SignPath Foundation.
