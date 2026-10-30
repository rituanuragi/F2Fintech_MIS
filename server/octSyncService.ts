import * as fs from 'fs';
import * as path from 'path';
import * as XLSXNamespace from 'xlsx';
const XLSX: any = (XLSXNamespace as any).default || XLSXNamespace;
import { getDb, saveDb } from './db.js';
import { indexWorkbookIntoDatabase } from './customerLookupRouter.js';
import { rowHasAnyEntry, cleanWorkbookSheets } from './normalizer.js';
import { invalidateMisSearchBookCache } from './misSearchBook.js';

const F2_ALL_TICKETS_URL = 'https://admin.f2fintech.in/api/v1/get-all-tickets';
const F2_TICKET_DETAIL_BASE_URL = 'https://admin.f2fintech.in/api/v1/get-ticket-with-detail';
const CACHE_DIR = path.resolve(process.cwd(), 'data/cache');
const MASTER_MIS_PATH = path.join(CACHE_DIR, 'MASTER MIS.xlsx');
const GSHEET_MIS_PATH = path.join(CACHE_DIR, 'gsheet_mis_master_mis.xlsx');
const OCT_CSV_PATH = path.join(CACHE_DIR, 'oct26_user_input.csv');
const F2_SESSION_PATH = path.join(CACHE_DIR, 'f2_tickets_session.json');

// Memory cache for ticket details to prevent repeated HTTP calls
const ticketDetailMemoryCache = new Map<number | string, Record<string, any>>();

// State tracker for OCT auto-sync
export const octSyncState = {
  lastSyncAt: null as string | null,
  totalOctRows: 93,
  previousOctRows: 93,
  newRowsAdded: 0,
  updatedRows: 0,
  lastTicketId: null as number | string | null,
  isSyncing: false,
};

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

/**
 * Format any ISO or timestamp date into "D-MMMM-YY" (e.g. "12-October-26")
 */
export function formatToSheetDate(dateVal: any): string {
  if (!dateVal) return '';
  if (typeof dateVal === 'number' && dateVal > 35000 && dateVal < 65000) {
    const days = Math.round(dateVal);
    const utcMs = Date.UTC(1899, 11, 30) + days * 86400000;
    const d = new Date(utcMs);
    const day = d.getUTCDate();
    const month = MONTH_NAMES[d.getUTCMonth()];
    const yr = String(d.getUTCFullYear()).slice(-2);
    return `${day}-${month}-${yr}`;
  }
  const s = String(dateVal).trim();
  if (!s || s === 'null' || s === 'undefined') return '';
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const day = d.getDate();
    const month = MONTH_NAMES[d.getMonth()];
    const yr = String(d.getFullYear()).slice(-2);
    return `${day}-${month}-${yr}`;
  }
  return s;
}

/**
 * Normalize customer name for fuzzy matching
 */
function normalizeName(name: any): string {
  return String(name || '')
    .toUpperCase()
    .replace(/^DR\.?\s+|^MR\.?\s+|^MRS\.?\s+|^MS\.?\s+|^CA\.?\s+/i, '')
    .replace(/[^A-Z0-9]/g, '')
    .trim();
}

/**
 * Fetch detail for an individual ticket with caching
 */
async function fetchTicketDetailWithCache(ticketId: number | string): Promise<Record<string, any> | null> {
  const cleanId = String(ticketId).trim();
  if (ticketDetailMemoryCache.has(cleanId)) {
    return ticketDetailMemoryCache.get(cleanId)!;
  }
  try {
    const res = await fetch(`${F2_TICKET_DETAIL_BASE_URL}/${encodeURIComponent(cleanId)}`, {
      headers: { Accept: 'application/json' },
      signal: AbortSignal.timeout(1500),
    });
    if (!res.ok) return null;
    const json: any = await res.json();
    const data = json?.data;
    if (data && typeof data === 'object') {
      ticketDetailMemoryCache.set(cleanId, data);
      return data;
    }
  } catch (err) {
    // Return null on failure; fallback gracefully to base ticket info
  }
  return null;
}

/**
 * Core function that reconciles the OCT 26 sheet with live OMS tickets,
 * guaranteeing accurate row counts (~104+ rows) and updated case details.
 */
export async function syncOctWithLiveTickets(options: { forceFetchApi?: boolean } = {}): Promise<{
  success: boolean;
  octRowsCount: number;
  previousOctRowsCount: number;
  newRowsAdded: number;
  updatedRowsCount: number;
  totalIndexed: number;
  latestTicketId: number | string | null;
  message: string;
}> {
  if (octSyncState.isSyncing) {
    return {
      success: true,
      octRowsCount: octSyncState.totalOctRows,
      previousOctRowsCount: octSyncState.previousOctRows,
      newRowsAdded: octSyncState.newRowsAdded,
      updatedRowsCount: octSyncState.updatedRows,
      latestTicketId: octSyncState.lastTicketId,
      totalIndexed: 0,
      message: 'Sync is currently in progress.',
    };
  }

  octSyncState.isSyncing = true;
  try {
    if (!fs.existsSync(CACHE_DIR)) {
      fs.mkdirSync(CACHE_DIR, { recursive: true });
    }

    // 1. Fetch live tickets from F2 Fintech OMS
    let liveTickets: Record<string, any>[] = [];
    try {
      const apiRes = await fetch(`${F2_ALL_TICKETS_URL}?companyId=101&page=1&limit=1000`, {
        headers: {
          Accept: 'application/json',
          'Cache-Control': 'no-cache',
        },
        signal: AbortSignal.timeout(10000),
      });
      if (apiRes.ok) {
        const json: any = await apiRes.json();
        const results = json?.data?.results || [];
        if (Array.isArray(results) && results.length > 0) {
          liveTickets = results;
        }
      }
    } catch (apiErr: any) {
      console.warn('[octSyncService] Warning fetching live F2 tickets from API:', apiErr?.message);
    }

    // 2. Combine with f2_tickets_session.json to ensure historical & current tickets are unified
    if (fs.existsSync(F2_SESSION_PATH)) {
      try {
        const sessionRaw = fs.readFileSync(F2_SESSION_PATH, 'utf-8');
        const sessionJson = JSON.parse(sessionRaw);
        const sessionRecords = sessionJson?.records || [];
        if (Array.isArray(sessionRecords)) {
          const map = new Map<string, Record<string, any>>();
          for (const r of sessionRecords) {
            if (r.ticketId) map.set(String(r.ticketId), r);
          }
          // Live API results overwrite cache with freshest statuses
          for (const r of liveTickets) {
            if (r.ticketId) map.set(String(r.ticketId), r);
          }
          liveTickets = Array.from(map.values());
        }
      } catch {}
    }

    // 3. Filter tickets belonging to October 2026 (or existing tickets in October tab to update)
    const octoberTickets = liveTickets.filter((t) => {
      const d = String(t.createdAt || t.created_at || t.applicationDate || '');
      const isOctDate = d.includes('2026-10') || d.includes('10-2026') || /oct.*20?26/i.test(d) || (t.applicationDate && String(t.applicationDate).includes('2026-10'));
      return isOctDate;
    });

    // 4. Open or create MASTER MIS.xlsx
    let workbook: any = null;
    const baseFallbackPath = path.join(CACHE_DIR, 'gsheet_15cVL3BK6yXPTwGGG0s_gp-jW-KxS-70cXn__6C8VCBs.xlsx');
    if (fs.existsSync(MASTER_MIS_PATH)) {
      try {
        workbook = XLSX.readFile(MASTER_MIS_PATH, { cellDates: true, dense: true });
      } catch {}
    }
    if (!workbook && fs.existsSync(GSHEET_MIS_PATH)) {
      try {
        workbook = XLSX.readFile(GSHEET_MIS_PATH, { cellDates: true, dense: true });
      } catch {}
    }
    if (!workbook && fs.existsSync(baseFallbackPath)) {
      try {
        workbook = XLSX.readFile(baseFallbackPath, { cellDates: true, dense: true });
      } catch {}
    }
    if (!workbook) {
      workbook = XLSX.utils.book_new();
    }

    // Find OCT 26 tab
    const octTabName =
      workbook.SheetNames.find((s: string) => s.toUpperCase().includes('OCT 26') || s.toUpperCase() === 'OCT 26') ||
      workbook.SheetNames.find((s: string) => s.toUpperCase().includes('OCT')) ||
      'OCT 26';

    let octWs = workbook.Sheets[octTabName];
    if (!octWs && fs.existsSync(baseFallbackPath)) {
      try {
        const baseWb = XLSX.readFile(baseFallbackPath);
        if (baseWb.Sheets['OCT 26']) {
          octWs = baseWb.Sheets['OCT 26'];
          workbook.Sheets[octTabName] = octWs;
          if (!workbook.SheetNames.includes(octTabName)) {
            workbook.SheetNames.splice(1, 0, octTabName);
          }
        }
      } catch {}
    }
    if (!octWs) {
      octWs = XLSX.utils.aoa_to_sheet([]);
      workbook.Sheets[octTabName] = octWs;
      if (!workbook.SheetNames.includes(octTabName)) {
        workbook.SheetNames.splice(1, 0, octTabName);
      }
    }

    // Parse existing rows into 2D array
    let rawRows = (XLSX.utils.sheet_to_json(octWs, { header: 1, raw: true, defval: '' }) as any[][]) || [];

    // Standard 38 columns used in OCT 26 sheet
    const STANDARD_HEADERS = [
      'TICKET ID',
      'CAM',
      'TVR',
      'LOGIN POC',
      'CREDIT POC',
      'OPS POC',
      'EMP NAME',
      'TL NAME',
      'MANAGER',
      'EMP TYPE',
      'CHANNEL PARTNER MANAGER NAME',
      'CIBIL SCORE',
      'CUSTOMER NAME',
      'CREDIT ELIGIBILTY',
      'BRANCH',
      'CASE CREATED DATE',
      'LAST UPDATED DATE',
      'LAN NO',
      'PHONE NUMBER',
      'LOCATION STATE',
      'CITY',
      'LOAN TYPE',
      'CODE',
      'LENDER',
      'NET LOAN AMOUNT',
      'LOAN AMOUNT',
      'LOAN CATEGORY',
      'STATUS',
      'APPROVAL DATE',
      'DISBURSED DATE',
      'OPS REMARKS',
      'CREDIT REMARKS',
      'CONTRIBUTOR',
      'BANKER NAME',
      'BANKER PHONE NO',
      'BANKER MAIL ID',
      'CASH BACK / SUBVENTION',
      'DEGREE WITH YEAR',
    ];

    let headers: string[] = [];
    if (rawRows.length > 0 && Array.isArray(rawRows[0]) && rawRows[0].length >= 10) {
      headers = rawRows[0].map((h: any, i: number) => String(h || STANDARD_HEADERS[i] || `Col_${i + 1}`).trim());
    } else {
      headers = [...STANDARD_HEADERS];
      rawRows.unshift(headers);
    }

    // STRICT FILTER: Discard any blank rows or template rows that have no real business data!
    // Take ONLY rows with genuine entries (customer name, phone, ticket ID, lender, amount)
    const cleanRows: any[][] = [headers];
    for (let r = 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (row && rowHasAnyEntry(row, headers)) {
        cleanRows.push(row);
      }
    }
    rawRows = cleanRows;

    // Map column indices
    const colIdx = {
      ticketId: headers.findIndex((h) => /^ticket\s*id$/i.test(h)),
      cam: headers.findIndex((h) => /^cam$/i.test(h)),
      tvr: headers.findIndex((h) => /^tvr$/i.test(h)),
      loginPoc: headers.findIndex((h) => /^login\s*poc$/i.test(h)),
      creditPoc: headers.findIndex((h) => /^credit\s*poc$/i.test(h)),
      opsPoc: headers.findIndex((h) => /^ops\s*poc$/i.test(h)),
      empName: headers.findIndex((h) => /^emp\s*name$/i.test(h)),
      tlName: headers.findIndex((h) => /^tl\s*name$/i.test(h)),
      manager: headers.findIndex((h) => /^manager$/i.test(h)),
      empType: headers.findIndex((h) => /^emp\s*type$/i.test(h)),
      cibil: headers.findIndex((h) => /^cibil\s*score$/i.test(h)),
      customerName: headers.findIndex((h) => /^customer\s*name$/i.test(h)),
      creditEligibility: headers.findIndex((h) => /^credit\s*eligibilty$/i.test(h)),
      branch: headers.findIndex((h) => /^branch$/i.test(h)),
      caseCreatedDate: headers.findIndex((h) => /^case\s*created\s*date$/i.test(h)),
      lastUpdatedDate: headers.findIndex((h) => /^last\s*updated\s*date$/i.test(h)),
      lanNo: headers.findIndex((h) => /^lan\s*no$/i.test(h)),
      phone: headers.findIndex((h) => /^phone\s*number$/i.test(h)),
      state: headers.findIndex((h) => /^location\s*state$/i.test(h)),
      city: headers.findIndex((h) => /^city$/i.test(h)),
      loanType: headers.findIndex((h) => /^loan\s*type$/i.test(h)),
      code: headers.findIndex((h) => /^code$/i.test(h)),
      lender: headers.findIndex((h) => /^lender$/i.test(h)),
      netLoanAmount: headers.findIndex((h) => /^net\s*loan\s*amount$/i.test(h)),
      loanAmount: headers.findIndex((h) => /^loan\s*amount$/i.test(h)),
      loanCategory: headers.findIndex((h) => /^loan\s*category$/i.test(h)),
      status: headers.findIndex((h) => /^status$/i.test(h)),
      approvalDate: headers.findIndex((h) => /^approval\s*date$/i.test(h)),
      disbursedDate: headers.findIndex((h) => /^disbursed\s*date$/i.test(h)),
      opsRemarks: headers.findIndex((h) => /^ops\s*remarks$/i.test(h)),
      creditRemarks: headers.findIndex((h) => /^credit\s*remarks$/i.test(h)),
      contributor: headers.findIndex((h) => /^contributor$/i.test(h)),
      bankerName: headers.findIndex((h) => /^banker\s*name$/i.test(h)),
      bankerPhone: headers.findIndex((h) => /^banker\s*phone\s*no$/i.test(h)),
      bankerMail: headers.findIndex((h) => /^banker\s*mail\s*id$/i.test(h)),
    };

    // Ensure fallback indices
    if (colIdx.ticketId === -1) colIdx.ticketId = 0;
    if (colIdx.customerName === -1) colIdx.customerName = 12;
    if (colIdx.phone === -1) colIdx.phone = 18;
    if (colIdx.caseCreatedDate === -1) colIdx.caseCreatedDate = 15;
    if (colIdx.lender === -1) colIdx.lender = 23;
    if (colIdx.loanAmount === -1) colIdx.loanAmount = 25;
    if (colIdx.status === -1) colIdx.status = 27;

    const previousOctRowsCount = Math.max(0, rawRows.length - 1);
    octSyncState.previousOctRows = previousOctRowsCount;

    // Track existing rows by ticketId and normalized customerName
    const existingTicketRowMap = new Map<string, number>();
    const existingNameRowMap = new Map<string, number>();

    for (let r = 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (!row) continue;
      const tid = String(row[colIdx.ticketId] || '').trim();
      if (tid && tid !== 'undefined' && tid !== 'null') {
        existingTicketRowMap.set(tid, r);
      }
      const cname = normalizeName(row[colIdx.customerName]);
      if (cname) {
        existingNameRowMap.set(cname, r);
      }
    }

    let updatedRowsCount = 0;
    const newTicketsToAdd: Record<string, any>[] = [];

    // 5. Match live tickets against sheet rows
    // Also include any live ticket that already exists in the sheet to update its status
    const allTicketsToReconcile = liveTickets.filter((t) => {
      const tid = String(t.ticketId || '').trim();
      return octoberTickets.includes(t) || existingTicketRowMap.has(tid);
    });

    for (const t of allTicketsToReconcile) {
      const tid = String(t.ticketId || '').trim();
      const normCust = normalizeName(t.customerName);

      // Check if ticket already exists by ticketId or by customerName
      let existingRowIdx = existingTicketRowMap.get(tid);
      if (existingRowIdx === undefined && normCust) {
        existingRowIdx = existingNameRowMap.get(normCust);
      }

      if (existingRowIdx !== undefined) {
        // Row exists in sheet: update any missing or updated details!
        const row = rawRows[existingRowIdx];
        let rowChanged = false;

        // If ticket ID was missing in row, fill it
        const currentTid = String(row[colIdx.ticketId] || '').trim();
        if ((!currentTid || currentTid === 'undefined' || currentTid === 'null') && tid) {
          row[colIdx.ticketId] = Number(tid) || tid;
          existingTicketRowMap.set(tid, existingRowIdx);
          rowChanged = true;
        }

        // Check if phone was missing
        const currentPhone = String(row[colIdx.phone] || '').trim();
        const livePhone = t.mobile || t.customerContact || t.phone;
        if (!currentPhone && livePhone) {
          row[colIdx.phone] = String(livePhone).trim();
          rowChanged = true;
        }

        // Update status if live status progressed
        const liveStatus = (t.status || t.ticketStatus || t.loanStatus || '').toUpperCase();
        const currentStatus = String(row[colIdx.status] || '').toUpperCase();
        if (liveStatus && liveStatus !== currentStatus && !['LOGIN', ''].includes(liveStatus)) {
          row[colIdx.status] = liveStatus;
          rowChanged = true;
        }

        // Update lender / provider if blank
        if (!row[colIdx.lender] && (t.provider || t.applicationProvider)) {
          row[colIdx.lender] = t.provider || t.applicationProvider;
          rowChanged = true;
        }

        // Update loan amount if blank
        if (!row[colIdx.loanAmount] && (t.loanAmount || t.applicationAmount)) {
          const amt = t.loanAmount || t.applicationAmount;
          row[colIdx.loanAmount] = amt;
          if (!row[colIdx.netLoanAmount]) row[colIdx.netLoanAmount] = amt;
          rowChanged = true;
        }

        if (rowChanged) {
          updatedRowsCount++;
        }
      } else if (octoberTickets.includes(t)) {
        // Completely new October ticket: queue to fetch details and append!
        newTicketsToAdd.push(t);
      }
    }

    // Sort new tickets in ascending chronological order
    newTicketsToAdd.sort((a, b) => Number(a.ticketId || 0) - Number(b.ticketId || 0));

    // Fetch full details for new tickets in parallel only if contact info is missing
    const ticketsNeedingDetail = newTicketsToAdd.filter((t) => !t.customerContact && !t.mobile);
    const chunkSize = 8;
    for (let i = 0; i < ticketsNeedingDetail.length; i += chunkSize) {
      const chunk = ticketsNeedingDetail.slice(i, i + chunkSize);
      await Promise.all(
        chunk.map(async (t) => {
          if (t.ticketId) {
            const detail = await fetchTicketDetailWithCache(t.ticketId);
            if (detail) {
              Object.assign(t, detail);
            }
          }
        })
      );
    }

    // 6. Append each new ticket as a properly formatted row in OCT 26
    let addedCount = 0;
    for (const t of newTicketsToAdd) {
      const tid = Number(t.ticketId) || String(t.ticketId || '');
      const custName = String(t.customerName || '').toUpperCase();
      const phone = String(t.customerContact || t.mobile || t.phone || '').trim();
      const provider = String(t.provider || t.applicationProvider || t.lender || 'ABFL').trim();
      const amount = t.applicationAmount || t.loanAmount || '';
      const rawDate = t.applicationDate || t.createdAt || t.created_at || '2026-10-12';
      const sheetDate = formatToSheetDate(rawDate);
      const city = String(t.customerLocation || t.city || '').toUpperCase();
      const state = String(t.customerState || t.state || '').toUpperCase();
      const emp = String(t.appliedByName || t.employeeName || '').toUpperCase();
      const mgr = String(t.managerName || '').toUpperCase();
      const loanCat = String(t.loanCategory || 'UNSECURED').toUpperCase();
      const loanType = String(t.loanType || 'OD').toUpperCase();
      const status = String(t.loanStatus || t.status || t.ticketStatus || 'SUBMITTED').toUpperCase();

      const newRow = new Array(headers.length).fill('');
      newRow[colIdx.ticketId] = tid;
      newRow[colIdx.customerName] = custName;
      newRow[colIdx.phone] = phone;
      newRow[colIdx.caseCreatedDate] = sheetDate;
      newRow[colIdx.lastUpdatedDate] = sheetDate;
      newRow[colIdx.state] = state;
      newRow[colIdx.city] = city;
      newRow[colIdx.loanType] = loanType;
      newRow[colIdx.code] = 'F2 FINTECH';
      newRow[colIdx.lender] = provider;
      newRow[colIdx.netLoanAmount] = amount;
      newRow[colIdx.loanAmount] = amount;
      newRow[colIdx.loanCategory] = loanCat;
      newRow[colIdx.status] = status;
      newRow[colIdx.empName] = emp;
      newRow[colIdx.manager] = mgr;
      newRow[colIdx.branch] = city || state || 'DELHI';
      newRow[colIdx.lanNo] = t.applicationId ? `APP${t.applicationId}` : '';
      newRow[colIdx.opsRemarks] = `LOGIN DONE || CASE ID: ${tid}`;
      newRow[colIdx.creditRemarks] = `${custName} case submitted via ${provider}. Loan Amount: ₹${amount}`;

      // Only add if it has genuine business entry
      if (rowHasAnyEntry(newRow, headers)) {
        rawRows.push(newRow);
        addedCount++;
      }
    }

    // FINAL CLEANUP: Ensure 100% of rows written have valid business entries (no blanks!)
    const finalCleanRows = rawRows.filter((r, idx) => idx === 0 || rowHasAnyEntry(r, headers));
    rawRows = finalCleanRows;

    // 7. Write updated 2D array back to worksheet
    const updatedWs = XLSX.utils.aoa_to_sheet(rawRows);
    workbook.Sheets[octTabName] = updatedWs;

    // Clean all worksheets in workbook (stripping blank rows from October 2025 and all tabs)
    cleanWorkbookSheets(workbook);

    const finalBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // Persist to all disk files so all readers / viewers have the freshest file
    const targetDiskFiles = [
      MASTER_MIS_PATH,
      GSHEET_MIS_PATH,
      path.join(CACHE_DIR, 'gsheet_15cVL3BK6yXPTwGGG0s_gp-jW-KxS-70cXn__6C8VCBs.xlsx'),
    ];
    const liveGsheetPath = path.join(CACHE_DIR, 'gsheet_1PDss3DpFJ7OmVdW3bBpAc9r6xlSaCrg6BcYYrc-B744.xlsx');
    if (fs.existsSync(liveGsheetPath)) {
      targetDiskFiles.push(liveGsheetPath);
    }

    for (const fPath of targetDiskFiles) {
      try {
        fs.writeFileSync(fPath, finalBuffer);
      } catch (saveErr) {
        console.warn('[octSyncService] Warning persisting workbook file:', fPath, saveErr);
      }
    }

    // Update CSV export file with ONLY clean rows with entries
    try {
      const csvContent = XLSX.utils.sheet_to_csv(updatedWs);
      fs.writeFileSync(OCT_CSV_PATH, csvContent, 'utf-8');
    } catch {}

    // Invalidate misSearchBook cache so that the UI immediately loads the latest data
    invalidateMisSearchBookCache();

    // 8. Re-index workbook into SQLite database
    const db = await getDb();
    const { totalIndexed, tabDetails } = indexWorkbookIntoDatabase(
      db,
      'mis_master_mis',
      'MASTER MIS',
      finalBuffer
    );

    const octTabDetail = tabDetails.find((t) => t.name.toUpperCase().includes('OCT'));
    const finalOctRows = octTabDetail ? octTabDetail.rows : rawRows.length - 1;

    const nowIso = new Date().toISOString();
    db.run(
      'UPDATE connected_sheets SET rows_count = ?, last_synced_at = ?, status = "synced", sheet_names_json = ? WHERE id = "mis_master_mis";',
      [totalIndexed, nowIso, JSON.stringify(tabDetails)]
    );
    saveDb();

    // 9. Update state tracker
    const maxTid = octoberTickets.reduce((max, t) => {
      const n = Number(t.ticketId);
      return !isNaN(n) && n > max ? n : max;
    }, 0);

    octSyncState.lastSyncAt = nowIso;
    octSyncState.totalOctRows = finalOctRows;
    octSyncState.newRowsAdded = addedCount;
    octSyncState.updatedRows = updatedRowsCount;
    octSyncState.lastTicketId = maxTid || null;

    console.log(
      `[octSyncService] Auto-sync complete: OCT 26 now has ${finalOctRows} rows (+${addedCount} new, ${updatedRowsCount} updated, previous: ${previousOctRowsCount}). Total workbook indexed: ${totalIndexed}.`
    );

    return {
      success: true,
      octRowsCount: finalOctRows,
      previousOctRowsCount,
      newRowsAdded: addedCount,
      updatedRowsCount,
      totalIndexed,
      latestTicketId: maxTid || null,
      message: `OCT 26 synchronized successfully with ${finalOctRows} total rows (+${addedCount} new cases added, ${updatedRowsCount} updated).`,
    };
  } catch (err: any) {
    console.error('[octSyncService] Sync failed with error:', err);
    throw err;
  } finally {
    octSyncState.isSyncing = false;
  }
}
