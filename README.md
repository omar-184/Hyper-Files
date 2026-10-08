# Hyper-Files

A light, fully offline office and PDF app for Windows. No AI features, no accounts, no cloud.

Hyper-Files opens and edits PDF, Word, Excel and PowerPoint files, and includes a PDF toolbox
(merge, split, compress, convert, protect, sign) plus Acrobat-style PDF editing.

> **Status:** Phase 0. The GenOffice code base is imported with its AI features, sign-in,
> analytics and auto-update removed. See [docs/PLAN.md](docs/PLAN.md).

## Goals

- **Offline only.** The app never connects to the internet. The single exception is an opt-in
  update check, off by default, that only asks GitHub Releases whether a newer version exists.
- **Light.** Runs well on a 4 GB RAM laptop with Windows 10 or 11. Files open in a fast read-only
  view first; the editor loads only when you click Edit.
- **PDF first.** PDF viewing, a PDF24-style toolbox and Acrobat-style editing come before office
  editing.
- **No AI, no accounts, no telemetry.**

## Platforms

Windows 10 and 11 (x64). Other platforms may follow later.

## Built on

Hyper-Files is built from [GenOffice](https://github.com/genspark-ai/genoffice) (Apache-2.0)
with its AI features and sign-in removed; [MuPDF](https://mupdf.com/) (AGPL-3.0) is planned for
the core of the PDF side. Hyper-Files is not affiliated with or endorsed by Genspark, Mainfunc, Inc. or
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
npm run dist:win   # Windows installer
```

## License

Hyper-Files is free software under the [GNU Affero General Public License v3.0](LICENSE).
Third-party components keep their own licenses; see [NOTICE](NOTICE).
