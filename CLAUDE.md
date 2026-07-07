# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

Single-file web form for NGS NT's IGU Daily Test Report (AS 4666:2012). The entire frontend is `index.html`. The backend (`Code.gs`) is a Google Apps Script web app — it lives in this repo for reference but is deployed separately via the Apps Script editor.

- **Live URL:** https://ngs-daily-report.vercel.app
- **Supervisor Dashboard:** https://ngs-daily-report.vercel.app/dashboard.html
- **Dashboard password:** NGSNT2026sup
- **GitHub remote:** https://github.com/Paul-Flores-Sinche/ngs-igu-report.git
- **Drive folder for saved reports:** `12YgmAFL5sYvTqwgD3MYsuOlbxU6K_F1B`

## Deploying changes

There is no build step. Edit `index.html`, then push to `main`:

```
git add index.html
git commit -m "..."
git push origin main
```

Vercel auto-deploys from `main`. Changes are live within ~60 seconds.

To preview locally, open `index.html` in a browser (`file://` works for everything except the submit/photo flows which require HTTPS).

## Architecture of index.html

All HTML, CSS, and JavaScript live in this one file (~1450 lines):

| Lines | Content |
|-------|---------|
| 1–200 | `<style>` block — CSS variables, layout, section/item components, print styles |
| 200–960 | HTML body — header, summary card, 10 collapsible sections |
| 960–970 | Submit button, print/reset buttons, status div |
| 970–end | `<script>` block — all JS logic |

### Key constants (top of `<script>`)

```js
const SHEET_URL = "https://script.google.com/macros/s/AKfycbx9BAW8_bna3fAE8kgDnaBrH83w38dAgcIaDNR7SKZqk24z0o484drYo6ItcjqiAbZX/exec";  // Apps Script deployment URL
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

## Notes and Signature Fields

96 `id` attributes were added across all textarea and signature inputs so every section can capture a note and a sign-off.

- **Naming convention:** `s{section}_notes` and `s{section}_sig` (e.g. `s31_notes`, `s31_sig`)
- **Special cases:** `s62p_notes`/`s62p_sig` (6.2 Pass), `s62f_notes`/`s62f_sig` (6.2 Fail), `s910_notes`/`s910_sig` (9.10)
- All 48 pairs are collected in `submitReport()` and rendered in `buildHtmlReport()` via `rowNotes()`

## Supervisor Dashboard (dashboard.html)

- Separate file served at `/dashboard.html`
- Password gate: sessionStorage key `ngs_sup_auth`, value `NGSNT2026sup`
- Reads reports via Apps Script `doGet?action=getReports`
- Approve action via Apps Script `doGet?action=approve&rowIndex=N&supervisorName=X`
- Auto-creates `Drive URL`, `Approved By`, `Approved At` columns in the sheet if missing

## Google Apps Script backend (`Code.gs`)

`Code.gs` is the canonical source for the backend. Changes must be pasted into the Apps Script editor and redeployed — the file in this repo is kept in sync manually.

### Deployment rules

- Deploy as **Web App → Execute as: Me → Access: Anyone**
- Every change requires a **new deployment** (not editing an existing one); paste the new URL into `SHEET_URL` in `index.html`
- The **Drive API advanced service** must be enabled: Apps Script editor → Services (+) → Drive API → Add. Required by `generateAndSavePDF()`.
- Use the **Executions** panel to read `console.log` output and diagnose errors

### Key functions

| Function | Purpose |
|----------|---------|
| `doPost(e)` | Web app entry point — parses JSON, calls `logToSheet` then `generateAndSavePDF` |
| `logToSheet(data)` | Appends a row to the "IGU Reports" sheet (creates sheet + headers on first run) |
| `generateAndSavePDF(data)` | Builds HTML via `buildHtmlReport`, uploads to Drive, converts to Google Doc |
| `buildHtmlReport(data, ...)` | Returns a full HTML string for the report document |
| `testDoPost()` | Run directly from the editor to test the full flow without a browser submission |

`CONFIG.FOLDER_ID` and `CONFIG.SHEET_NAME` are the only values that need changing if the Drive folder or sheet name ever changes.

### Apps Script data contract

The script receives the full `data` object from `submitReport()`. Fields consumed by `logToSheet` and `buildHtmlReport`:

**Header fields:** `date`, `applicableCert`, `shiftStart`, `shiftFinish`, `totalUnits`, `productionLine`, `completedBy`, `overallResult`, `passCount`, `failCount`

**Section 1 (sampling):** `s1` (activity), `s1_actual_lot` (worker's actual lot size), `s1_actual_units` (worker's actual units checked)

**Section results:** `s31`–`s35`, `s41`–`s47`, `s51`–`s56`, `s61`–`s69`, `s71`–`s74`, `s81`, `s82`, `s91`–`s910`, `s101`–`s103`

**Material detail fields:** `s41_manufacturer`, `s41_batch`, `s41_product`, `s41_size`, `s41_wall_thickness`, `s41_quantity`, `s41_grade`, `s41_finish`, `s41_date_receipt` — `s51_manufacturer`, `s51_batch`, `s51_product`, `s51_mol_sieve`, `s51_date_receipt`, `s51_expiry` — `s61_manufacturer`, `s61_batch`, `s61_product`, `s61_date_receipt`, `s61_expiry`

**Other:** `desiccantFillTime`, `desiccantAssemTime`, `sealUnitType63`, `photo41`, `photo51`, `photo61`

**Equipment checkboxes:** `eq0`–`eq5` (10.1), `aeq0`–`aeq14` (10.2) — collected as `'Yes'`/`'No'`

**Compliance markings:** `m1`–`m4` (8.1) — collected as `'Yes'`/`'No'`

**Atmospheric readings:** `atm_time_1`–`atm_time_6`, `atm_temp_1`–`atm_temp_6`, `atm_hum_1`–`atm_hum_6`, `atm_pres_1`–`atm_pres_6`

**Notes/signatures:** `s11_notes` through `s103_sig` (48 pairs — see [Notes and Signature Fields](#notes-and-signature-fields))

**Drive URL** is now logged to the sheet via a new column (see [Supervisor Dashboard](#supervisor-dashboard-dashboardhtml))

When adding new fields to `index.html`'s `data` object, also add them to `logToSheet`'s header array and `appendRow` call in `Code.gs`, then redeploy.

## Pending

- Detail tables (4.1, 5.1, 6.1, 2.1) have columns that are too narrow — Google Docs ignores CSS width on HTML blob conversion; needs a Google Docs API approach instead
- Auto-save with localStorage — `DRAFT_KEY` already defined, implementation pending
- SEO work on https://www.ngsnt.com.au/ (separate project)
