# Hyper-Files

A light, fully offline office and PDF app for Windows. No AI features, no accounts, no cloud.

Hyper-Files opens and edits PDF, Word, Excel and PowerPoint files, and includes a PDF toolbox
(merge, split, compress, convert, protect, sign) plus Acrobat-style PDF editing.

> **Status:** planning complete, development not started. See [docs/PLAN.md](docs/PLAN.md).

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

Hyper-Files will be built from [GenOffice](https://github.com/genspark-ai/genoffice) (Apache-2.0)
with its AI features and sign-in removed, and with [MuPDF](https://mupdf.com/) (AGPL-3.0) at the
core of the PDF side. Hyper-Files is not affiliated with or endorsed by Genspark, Mainfunc, Inc. or
Artifex Software. See [NOTICE](NOTICE) for attributions.

## Repository layout

```
docs/       Plan and design notes
```

More folders arrive with Phase 0 (importing the GenOffice code base).

## License

Hyper-Files is free software under the [GNU Affero General Public License v3.0](LICENSE).
Third-party components keep their own licenses; see [NOTICE](NOTICE).
