# Hypercube Office build plan

Snapshot of the planning document as of 2026-10-07. The living version is kept in the project's
plan doc; this copy travels with the code so contributors can read it offline.

## Decisions

| Topic                                | Decision                                                        |
| ------------------------------------ | --------------------------------------------------------------- |
| Name                                 | Hypercube Office                                                |
| Platform                             | Windows 10 and 11 (x64)                                         |
| Weakest machine                      | 4 GB RAM                                                        |
| Distribution                         | Personal use and free public release                            |
| Source code                          | Open source, AGPL-3.0                                           |
| First priority                       | PDF tools and PDF editing                                       |
| Language                             | English first; Arabic for everything in a later phase           |
| OCR                                  | Arabic + English, deferred until the light-OCR research is done |
| Old formats (.doc, .xls, .ppt, .odt) | Optional add-on                                                 |
| Default app                          | Installer offers to become default for PDF and Office files     |
| Updates                              | Opt-in update check, off by default                             |
| Network                              | None, apart from the opt-in update check                        |
| AI, accounts, telemetry              | None                                                            |

## Licensing

- **GenOffice** (genspark-ai/genoffice) is Apache-2.0, which may be combined into an AGPL-3.0
  app. We keep its license text, NOTICE and copyright headers, and never use its names or logos.
  The `ee/` directory (GenOffice Enterprise License) is excluded.
- **PDF24** is closed-source freeware; nothing is copied from it. Its feature list is only a
  checklist, rebuilt from open-source engines.
- **MuPDF** and **Ghostscript** are AGPL-3.0, which is why Hypercube Office is AGPL-3.0.

## Engines

| Feature                                                          | Engine                                | License               | Phase |
| ---------------------------------------------------------------- | ------------------------------------- | --------------------- | ----- |
| View, search, print                                              | MuPDF                                 | AGPL-3.0              | 2     |
| Merge, split, reorder, rotate, delete, extract, insert pages     | MuPDF                                 | AGPL-3.0              | 3     |
| Protect, unlock, repair, optimize for web                        | qpdf                                  | Apache-2.0            | 3     |
| Compress (standard)                                              | MuPDF clean + image downsampling      | AGPL-3.0              | 3     |
| Compress (strong), PDF/A, grayscale                              | Ghostscript (converter pack)          | AGPL-3.0              | 3     |
| Images to/from PDF                                               | MuPDF                                 | AGPL-3.0              | 3     |
| Annotate, stamps, forms, drawn signatures                        | MuPDF + GenOffice PDF UI              | AGPL-3.0 / Apache-2.0 | 3     |
| Compare two PDFs                                                 | Own code on MuPDF text extraction     | AGPL-3.0              | 3     |
| Office to PDF                                                    | GenOffice engines; LibreOffice add-on | Apache-2.0 / MPL-2.0  | 3     |
| Edit existing text and images, redaction, certificate signatures | MuPDF (PDFium evaluated in a spike)   | AGPL-3.0 / BSD        | 4     |
| PDF to Word/Excel                                                | LibreOffice add-on                    | MPL-2.0               | 5     |
| OCR (Arabic + English)                                           | To be chosen, see OCR                 | n/a                   | 7     |

## Architecture

- One window with tabs (Home, PDF, Tools, then Docs, Sheets, Slides) instead of GenOffice's six
  separate apps.
- Runtime: Tauri (WebView2) or single-process Electron, decided at the Phase 0 gate by measuring
  installer size, memory and startup on a 4 GB Windows 10 machine.
- Viewer first: files open in a read-only MuPDF view; editors load on demand. Targets on a 4 GB
  laptop: app open under 2 s, a 100-page PDF visible under 1 s.
- Removed from GenOffice: `agent-core`, `ai-provider`, `ai-search`, AI panels, Genspark sign-in,
  telemetry, automatic updates.
- Optional converter pack: LibreOffice + Ghostscript, installed separately.
- Installer offers (never forces) to become the default app; Windows' Default Apps screen confirms.
- Text and layout code is written right-to-left ready, so Arabic can be switched on later.

## Phases

0. **Fork, audit, decide runtime.** Import GenOffice (without `ee/`), rebrand, map every AI,
   sign-in, telemetry and auto-update call, measure Electron vs Tauri, collect a 20-file English
   PDF test set. _Gate:_ runtime chosen; builds and runs on Windows 10 and 11.
1. **Offline and AI-free.** Remove AI, sign-in and telemetry; add the opt-in update check.
   _Check:_ zero network connections with the update check off.
2. **Light PDF viewer and shell.** Tabbed window, MuPDF viewer, recent files, English UI.
   _Check:_ startup and open-time targets met on 4 GB RAM.
3. **PDF toolbox (PDF24 parity),** with batch mode. _Check:_ every tool passes the test set.
4. **Acrobat-style editing (English).** Engine spike, then text/image editing, text boxes,
   cover-and-retype, redaction, form designer, certificate signatures.
5. **Office editing.** Docs, Sheets, Slides back from GenOffice; converter pack.
6. **Arabic everywhere.** RTL interface, Arabic in PDFs and Office files.
7. **OCR (Arabic + English).** After the research below.
8. **Release.** Windows installer, GitHub Releases, offline help, third-party license screen.

## OCR research (deferred)

Compare on 10 Arabic and 10 English scans on a 4 GB laptop:

| Option                                         | Added size (approx.) | Notes                                             |
| ---------------------------------------------- | -------------------- | ------------------------------------------------- |
| Windows built-in OCR (Windows.Media.Ocr)       | ~0 MB                | Arabic needs the Windows language pack; to verify |
| Tesseract "fast" models, Arabic + English only | ~5–10 MB             | Apache-2.0 models                                 |
| RapidOCR / PaddleOCR (ONNX)                    | ~15–30 MB            | Often more accurate on photos; heavier on CPU     |

## Risks

| Risk                                                | Fallback                                                 |
| --------------------------------------------------- | -------------------------------------------------------- |
| Editing existing Arabic PDF text often fails        | Cover-and-retype with proper shaping                     |
| Porting GenOffice's Electron code to Tauri is large | Stay on single-process Electron                          |
| GenOffice is young (0.5.x)                          | Office work comes last; LibreOffice add-on               |
| AI wired deeper than expected                       | Stub AI interfaces so nothing calls out                  |
| Unsigned installer triggers SmartScreen             | Free open-source code signing (e.g. SignPath Foundation) |
| Name not formally trademark-checked                 | Search USPTO, EUIPO, WIPO before public release          |
