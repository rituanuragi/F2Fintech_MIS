# HANDOFF PROMPT — paste everything below into a new Claude chat

I am Shubhi. I am building **F2 Lead Intelligence**, an internal web app for F2 Fintech (loan leads, India).
Please use simple, easy English. Treat the project zip I attach as the source of truth.

## What the app does
Two modes in one React + Express app:
1. **Drive Directory Indexer** — connects a public Google Drive folder, downloads the Excel workbooks (BL, PL, DOCTORS directories, Consolidated Directories, Team X, NEET, LOT, LOGIN, Tata Points, etc.), and indexes every row into a local SQLite (sql.js) database `data/db`. Screens: Dashboard, Directory, Files, Sheets, Global Search, Batch Excel Search, Source Finder, Duplicates, Sync History, Errors, Settings.
2. **Customer Verification Suite** — Customer Lookup (PAN / Aadhaar / phone, field diff, remarks, audit trail), Google Sheets MIS sync, MIS Analyzer, MIS Dashboard (cases by date, status, lender, export), MIS Search Book, and F2 Live Tickets API explorer. Live ticket matching uses phone, lender, name, ticket id and amount.
Every lead gets a **priority** from rules in `server/priority.ts` (editable in the Priority Config modal; separate rules for directory and MIS).

## Stack and layout
React 18 + Vite + Tailwind 4 (`src/`), Express + sql.js + xlsx + exceljs (`server/`), run with `npm install` then `npm run dev` (server on 3000, Vite on 5173 with /api proxy). Env vars in `.env.example` (APPS_SCRIPT_URL, APPS_SCRIPT_TOKEN, DRIVE_ROOT_FOLDER_ID, MIS_SPREADSHEET_ID).
Backend files: db, crawler, indexer, search, priority, normalizer, dateUtils, google, mis, misDashboardRouter, misSearchBook, octSyncService, customerLookupRouter (router name `shubhiRouter`), f2TicketsRouter, healthRouter, and a NEW `server/server.ts`.

## IMPORTANT: what is rebuilt vs original
- `server/server.ts` is NEW. The original `serverApp.ts` was never shared, so these routes were rebuilt from how the frontend uses them: /api/status, dashboard, directory, connect, sync, disconnect, search, export, preview, export-sheet, record/:id, history, errors, priority-rules (+sync/reset), mis-priority-rules, category-types. They have NOT been run. Check them first.
- These screens are only **placeholders** (originals were not shared) and must be rebuilt: BatchExcelSearch, CustomerLookupView, DuplicatesView, SheetsView, SourceFinderView, VerificationAuditModal, AddCategoryTypeModal. DirectoryView, DashboardView, Sidebar, ConnectModal and F2FintechLogo are simple working rewrites, not the originals.
- Nothing has been built or run: no `npm install`, no typecheck, no tests. Do that first and fix errors.

## Fixes already made (verify them, do not redo)
- SheetPreviewModal, GlobalSearchView, HistoryView, FullRowModal: check `res.ok` and show an error instead of crashing / showing a fake green "clean" message.
- MisDashboardView: quick-date pills now come from real `availableDates`; CSV export uses a Blob (a `#` no longer cuts the file) and guards against formula injection (= + - @).
- FullRowModal: uses `match.normPhone / phoneRaw` (the old `match.phone` did not exist), so it no longer matches another customer's ticket.
- FilesView: real status column; "All Directories" count uses the de-duplicated list.
- LeadAuditModal: PAN masked to last 4 characters.
- main.tsx: error boundary around the whole app.
- Global Search (my decision): shows ONLY the location of a match (workbook, sheet, row) for rows in the main directory files. MIS and Google Sheets rows are excluded. Export has only Folder, Workbook, Worksheet, Excel Row, Match Type, Matched Column and drops MIS rows (done in `search.ts exportResults`).
- GET /api/priority-rules no longer changes anything (sync only on POST .../sync).
- Server already returns auto-sync status both at top level and under `autoSync`, so that old issue is not real.

## Known problems still open (from my code review)
1. SECURITY: `proxy-document` and `preview-sheet` in `f2TicketsRouter.ts` fetch any http(s) URL (open proxy / SSRF) and put `filename` straight in a header. Add an allowlist (only our S3 host — ASK ME for the exact host) and clean the filename.
2. `POST /records/:id/check` creates an audit entry every time, and `CustomerLookupView.handleOpenAudit` calls it just to OPEN the trail. Add a read-only route; take the admin name from the logged-in user, not the request body.
3. `search.ts` replaces `sheetName` with a display label (formatTeamXSheetName), so Preview opens the wrong sheet. Keep real `sheetName`, add `sheetLabel`. Same for Duplicates and Batch.
4. Global Search never pages (limit=100, no offset) but the header says "N matches found". Add paging / "load more".
5. MIS Analyzer cards use `filteredResults`, so match rate becomes 100% after clicking "Matched". Use the unfiltered list.
6. MIS Dashboard KPIs read exact keys (statusBreakdown['APPROVED']) but the table uses includes('APPROV'). Use the same rule. Also status/lender dropdowns shrink after filtering, and still default to 2026-10-10 and tab "OCT 26" (hardcoded).
7. Excel upload and Google Sheet use different phone-column detection (browser vs server `detectPhoneColumn`); Batch Excel Search has a third copy. Use the server version everywhere.
8. Batch `splitByMonth` reads `r.date` which the server never sends; `dateHelper.extractMonthYearLabel` guesses year 2026. Return a real date from the server.
9. MIS Search Book polling has a stale closure (effect depends only on selectedWorkbookId). Uploading there silently overwrites MASTER MIS — add a warning/confirm.
10. MisDashboardView calls /mis-dashboard/sync-sheet every 15–60 s per open tab, and each call can trigger a live sync + full reindex. Move to one server-side schedule.
11. PriorityConfigModal runs a sync every time it opens, so deleted rules can come back; sheet-priority cards call movePriority(idx) with the wrong array.
12. PAN/Aadhaar masking is inconsistent (CustomerLookupView masks; FullRowModal shows all raw columns). Hardcoded service-account email, real sample phone numbers, Drive/sheet IDs and logo S3 link in the UI. PDF.js loads worker/cmaps from unpkg. "Ready"/"Synchronized"/"All 11 checks" labels are hardcoded in FilesView(fixed)/DirectoryView/SystemHealthView.
13. Batch connect counts any HTTP 200 as "indexed". Duplicate files are dropped by base name across folders. Any directory not named BL/PL is labelled DOCTORS. GoogleSheetsSyncView says deselected sheets are removed "when the app loads" with no confirmation.
14. CSV/Excel exports in other views should get the same formula-injection guard.

## Open questions for me
- Which website addresses should the document proxy allow (S3 host)?
- Should the MIS Dashboard auto-sync run per open tab, or once on a server schedule? (I lean: server schedule.)
- Do I want to keep the old UI look for the placeholder screens? (I will re-upload the originals if I still have them.)

## What I will attach in the new chat
The project zip `f2-lead-intelligence.zip` (everything above), plus any original files I still have for the placeholder screens, `serverApp.ts`, and `package.json`.

## What I want you to do first
1. Run `npm install`, `npm run typecheck`, fix compile errors.
2. Start the server and test the rebuilt /api routes one by one.
3. Fix open problems 1–3 (security, audit, sheet label), then 4–6.
4. Rebuild the placeholder screens, starting with CustomerLookupView and BatchExcelSearch.
Ask me at most one question at a time. Do not change behaviour that needs a business decision without asking.

## Update v1.3.0
- ONE "Combine Columns" function. The old "Add Category Type" screen is removed. Category types a user added earlier are moved into normal priority rules (once, on server start). New ones are added inside Combine Columns ("Add a new combination": name + file/folder text). Rule field `isCustom` = made by the user; it matches when the file / folder / sheet name CONTAINS the text.
- MIS Google Sheet: all tabs are fetched. If the public export is blocked, all tabs are read through Apps Script (listTabs + readSheet) before an old cache is used. Uploading a multi-tab workbook to MASTER MIS now adds/replaces EVERY uploaded tab (it used to keep only the OCT tab) and keeps existing tabs.
- Batch Excel Search: Pivot Summary table shown in the app; headers (title, 2 column headers, Blanks, Total) can be edited before download; the Excel Pivot sheet uses the same texts. Files: src/components/BatchPivotPanel.tsx, src/utils/pivot.ts, search.ts exportBatchLookup, directoryToolsRouter /batch/export.
- Not run end to end (no npm install / Drive access in the build sandbox). Test: Combine Columns, MASTER MIS Sync Live tab count, Batch pivot + Excel download.

## Update v1.3.1
- Priority + Combine Columns settings are now saved in a fixed folder outside the project (`~/.f2-lead-intelligence`, or env F2_CONFIG_DIR). Re-unzipping the project or starting from another folder no longer loses them. Old ./data/config files are copied once. The folder is printed at server start.
- MIS "Sync Live" no longer says "Synchronized / verified all tabs" when Google could not be reached. It now warns that the saved copy was used and lists its tabs (this was why only OCT 26 showed).


## Update v1.4.0 - Login and roles
- Login page; admin sees all tabs, employees see only the tabs granted to them (the Users & Access screen was removed from the UI). Server: server/auth.ts. First login: admin / admin123, then it forces a new password.
- Admin-only server routes: user management, connect/disconnect, sync, delete-all, priority and MIS sheet writes. Other API routes only need a login (not per-tab).
- Not run in a browser yet (tsc passes; server login flow tested with curl).
