# Hypercube Office

A light, fully offline office and PDF app for Windows. No AI features, no accounts, no cloud.
Formerly named Hyper-Files (0.1.x); upgrading keeps your settings and save folder.

Hypercube Office opens and edits PDF, Word, Excel and PowerPoint files, and includes a PDF toolbox
(merge, split, compress, convert, protect, sign) plus Acrobat-style PDF editing.

> **Status:** Phase 0. The GenOffice code base is imported with its AI features, sign-in,
> analytics and auto-update removed. See [docs/PLAN.md](docs/PLAN.md).

## Goals

- **Offline only.** The app never connects to the internet on its own: starting it or opening a
  file makes no network request. It goes online only when you ask, for example to load an HTML
  file's web content, insert a picture from a web address, or check for updates (off by default).
  [PRIVACY.md](PRIVACY.md) lists every case.
- **Light.** Runs well on a 4 GB RAM laptop with Windows 10 or 11. Files open in a fast read-only
  view first; the editor loads only when you click Edit.
- **PDF first.** PDF viewing, a PDF24-style toolbox and Acrobat-style editing come before office
  editing.
- **No AI, no accounts, no telemetry.**

## Platforms

Windows 10 and 11 (x64). Other platforms may follow later.

## Built on

Hypercube Office is built from [GenOffice](https://github.com/genspark-ai/genoffice) (Apache-2.0)
with its AI features and sign-in removed; [MuPDF](https://mupdf.com/) (AGPL-3.0) is planned for
the core of the PDF side. Hypercube Office is not affiliated with or endorsed by Genspark, Mainfunc, Inc. or
Artifex Software. See [NOTICE](NOTICE) for attributions.

## Repository layout

```
apps/       Electron apps: shell (home + tabs), docs, sheets, slides, pdf, markdown, html
packages/   Shared engines (docx, pptx, xlsx, pdf2docx, html2docx) and UI/Electron helpers
e2e/        Playwright end-to-end tests for the shell
tools/      Build, lint and asset scripts
docs/       Plan and design notes
```

## Development

Requires Node 22+, npm 10+ and a Rust toolchain (for the sheets engine). See
[CONTRIBUTING.md](CONTRIBUTING.md).

```bash
npm ci
npm run dev        # all editors + shell against Vite dev servers
npm run dist:win   # Windows installer (apps/shell/release/Hypercube-Office-Setup-<version>-x64.exe)
```

## Installing

The **Windows installer** workflow (`.github/workflows/release-windows.yml`) builds the x64
installer from a clean checkout on every version tag, on demand, and on pull requests that touch
packaging, then smoke-tests a silent install and uninstall. Download it from the workflow run's
artifacts. The installer:

- installs per user by default (no admin prompt); "for all users" is offered on the first page;
- shows the AGPL-3.0 license and ships `LICENSE.txt`, `NOTICE.txt` and
  `THIRD-PARTY-NOTICES.txt` in the install folder;
- registers Hypercube Office for PDF, Word, Excel, PowerPoint, CSV, Markdown and HTML files, and its
  last page offers to make Hypercube Office the default app (Windows asks you to confirm in Settings);
- is not code-signed yet, so Windows SmartScreen shows "Unknown publisher" until you choose
  "More info" > "Run anyway".

## Code signing policy

See [CODE_SIGNING_POLICY.md](CODE_SIGNING_POLICY.md): who builds, reviews and approves signed
releases, and what the installer may send over the network. Signing through the SignPath
Foundation has been applied for; releases stay unsigned until it is approved.

## License

Hypercube Office is free software under the [GNU Affero General Public License v3.0](LICENSE).
Third-party components keep their own licenses; see [NOTICE](NOTICE).
