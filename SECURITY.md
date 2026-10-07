# Security Policy

## Reporting a Vulnerability

Please report suspected vulnerabilities privately via GitHub's
[private vulnerability reporting](https://github.com/omar-184/Hyper-Files/security/advisories/new)
on this repository. Do not open public issues for security reports. We aim to
acknowledge reports within 72 hours.

## Process Security Posture

All application windows run with the full Electron renderer lockdown:

- `contextIsolation: true`, `nodeIntegration: false`, `sandbox: true` for every
  document window and tab view (docs, sheets, slides, pdf, markdown, html, shell).
- Renderers reach the main process only through typed, validated IPC channels
  (payloads are schema-checked in the main process; sheets uses zod end to end).
- Every `shell.openExternal` call goes through a single shared gate
  (`@genoffice/electron-utils` → `safeExternalUrl`) that parses the URL and
  enforces a protocol allowlist (http/https; pdf link annotations additionally
  allow mailto). `file:`, `javascript:`, and custom schemes are always rejected.
- The app has no accounts, AI services, analytics or auto-update, and makes
  no network requests on its own.

## Threat Model: Rendering Untrusted HTML

Three pipelines render untrusted HTML in a hidden
`BrowserWindow`. Every window is treated as hostile content: full renderer
lockdown (`sandbox: true`, `contextIsolation: true`, `nodeIntegration: false`),
no preload script, no IPC surface. The slides export drives the window
exclusively through `executeJavaScript` and destroys it under a watchdog
timeout; the HTML app's DOCX export and the docs altChunk conversion now do
the same (`Promise.race` against a fixed timeout, then `destroy()`), so a
page whose scripts never yield cannot strand the hidden window. HTML→DOCX
conversion renders the page with scripts enabled and does not sanitize the
markup before conversion — the output is a document, not a sandbox escape;
report it as a vulnerability if you find a way from the converted page into
the main process.

## Out of Scope

- Vulnerabilities that require an already-compromised machine or a modified
  binary. This includes the deliberate environment-variable override points
  for local development (`XLSX_SIDECAR_PATH`): setting them requires control
  of the process environment, which is equivalent to code execution on the
  machine.
