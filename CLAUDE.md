# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

Single-file web form for NGS NT's IGU Daily Test Report (AS 4666:2012). Deployed on GitHub Pages. The entire frontend is `index.html`. The backend is a Google Apps Script web app that lives outside this repo.

- **Live URL:** https://paul-flores-sinche.github.io/ngs-igu-report/
- **GitHub remote:** https://github.com/Paul-Flores-Sinche/ngs-igu-report.git
- **Drive folder for PDFs:** `12YgmAFL5sYvTqwgD3MYsuOlbxU6K_F1B`

## Deploying changes

There is no build step. Edit `index.html`, then push to `main`:

```
git add index.html
git commit -m "..."
git push origin main
```

GitHub Pages serves `main` directly. Changes are live within ~60 seconds.

To preview locally, open `index.html` in a browser (`file://` works for everything except the submit/photo flows which require HTTPS).

## Architecture of index.html

All HTML, CSS, and JavaScript live in this one file (~1400 lines):

| Lines | Content |
|-------|---------|
| 1–191 | `<style>` block — CSS variables, layout, section/item components, print styles |
| 192–941 | HTML body — header, summary card, 10 collapsible sections |
| 942–952 | Submit button, print/reset buttons, status div |
| 954–1387 | `<script>` block — all JS logic |

### Key constants (top of `<script>`)

```js
const SHEET_URL = "https://script.google.com/macros/s/...";  // Apps Script deployment URL
const DRAFT_KEY = 'ngs_igu_draft';                           // localStorage key
```

**When the Apps Script is redeployed, `SHEET_URL` must be updated here.**

### State model

State is never stored in JS variables — it lives entirely in the DOM:
- **Pass/Fail:** `.pf-btn.pass.active` / `.pf-btn.fail.active` / `.pf-btn.na.active`
- **Activity toggles:** `.toggle.yes` / `.toggle.no`

`saveState()` serialises all form inputs + CSS state to `localStorage`. `restoreState()` reverses it on page load.

### Submit flow

`submitReport()` collects data using two DOM-query helpers:

- `getPF(sectionId, itemIndex)` — finds the nth `.pf-row` within `#sectionId` and returns `'Pass'`, `'Fail'`, or `'Not checked'`
- `getAct(sectionId, itemIndex)` — finds the nth `.toggle-group` and returns `'Yes'`, `'No'`, or `'Not completed'`

**Index order is positional in the DOM.** Adding or reordering `.pf-row` or `.toggle-group` elements within a section changes what index corresponds to which item. The `data` object field names (e.g. `s41`, `s92`) must match the Apps Script sheet column mapping.

POSTs JSON with `Content-Type: text/plain` (avoids CORS preflight). Expects back `{status: 'ok', driveUrl: '...'}` or `{status: 'error', message: '...'}`.

Photos (base64 data URLs from `img41`, `img51`, `img61`) are embedded inline in the JSON payload. If no photo was taken, the field is `''`.

### OCR flow (sections 4.1, 5.1, 6.1)

`triggerCamera(id)` → `showPhoto()` → `runOCR()` (Tesseract.js via CDN) → `parseLabel(text, section)` → `fillFromOCR(parsed, section)`

`OCR_FIELDS` maps parsed keys (`manufacturer`, `batch`, `expiry`, etc.) to HTML element IDs. Fields are only filled if currently empty — OCR never overwrites existing data.

`parseLabel()` uses regex heuristics. Section-specific patterns (spacer dimensions, molecular sieve type) are guarded by `if (section === '41')` / `if (section === '51')` blocks.

## Google Apps Script backend

The script is managed separately in the Google Apps Script editor (not in this repo). Key points for working on it:

- Deploy as **Web App → Execute as: Me → Access: Anyone**
- Every change requires a **new deployment** (not editing an existing one); the new URL must be pasted into `SHEET_URL` in `index.html`
- Use the **Executions** panel (left sidebar) in the editor to read `console.log` output and diagnose errors
- The `testDoPost()` function (if present in the script) can be run directly from the editor to test PDF generation without a browser form submission

### Apps Script data contract

The script receives the full `data` object from `submitReport()`. Field names that feed into the Google Sheet columns are:

`date`, `shiftStart`, `shiftFinish`, `totalUnits`, `productionLine`, `completedBy`, `overallResult`, `passCount`, `failCount` — plus section results `s31`…`s103`, spacer bar fields `s41_*`, desiccant fields `s51_*`, sealant fields `s61_*`, and photos `photo41`/`photo51`/`photo61`.
