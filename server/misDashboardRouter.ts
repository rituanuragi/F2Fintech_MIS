import { Router, Request, Response } from 'express';
import * as fs from 'fs';
import * as path from 'path';
import ExcelJS from 'exceljs';
import * as XLSXNamespace from 'xlsx';
const XLSX: any = (XLSXNamespace as any).default || XLSXNamespace;
import { getDb, saveDb } from './db.js';
import { MisService } from './mis.js';
import { indexWorkbookIntoDatabase } from './customerLookupRouter.js';
import { makeRecordKey } from './recordKey.js';

export const misDashboardRouter = Router();

const MONTH_NAMES: Record<string, number> = {
  jan: 1, january: 1,
  feb: 2, february: 2,
  mar: 3, march: 3,
  apr: 4, april: 4,
  may: 5,
  jun: 6, june: 6,
  jul: 7, july: 7,
  aug: 8, august: 8,
  sep: 9, sept: 9, september: 9,
  oct: 10, october: 10,
  nov: 11, november: 11,
  dec: 12, december: 12,
};

const SHORT_MONTH_NAMES = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/**
 * Universal date normalizer that handles any format found in Google Sheets/Excel:
 * - "10-October-26", "10-OCTOBER-2026", "10-October-2026", "10-Oct-26", "7-October-26"
 * - "10/10/2026", "10-10-2026", "10.10.2026"
 * - "2026-10-10", "2026/10/10"
 * - "October 10, 2026", "10 Oct 2026", "10th October 2026"
 * - Excel serial numbers like 46305
 * - ISO timestamps "2026-10-10T00:00:00.000Z"
 * Returns standard "YYYY-MM-DD" string or null.
 */
export function parseDateToIso(val: any): string | null {
  if (val === null || val === undefined) return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    const adj = new Date(val.getTime() + 12 * 3600 * 1000);
    const y = adj.getUTCFullYear();
    const m = String(adj.getUTCMonth() + 1).padStart(2, '0');
    const d = String(adj.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  let s = String(val).trim();
  if (!s || s === '-' || s === '—' || s === 'N/A' || s === 'NA' || s === 'null' || s === 'undefined') return null;

  // Strip trailing time e.g. "10-October-26 14:30:00" or "2026-10-10T14:30:00"
  s = s.replace(/\s+\d{1,2}:\d{2}(?::\d{2})?.*$/, '').replace(/T\d{1,2}:\d{2}(?::\d{2})?.*$/, '').trim();

  // Strip ordinal suffixes like 10th, 1st, 2nd, 3rd (using function replacer to prevent shell expansion issues)
  s = s.replace(/(\d+)(?:st|nd|rd|th)\b/gi, (_match, num) => num);

  // 1. Numeric Excel serial number (approx range 35000 to 65000 = ~1995 to 2078)
  if (/^\d{5}(?:\.\d+)?$/.test(s)) {
    const num = Number(s);
    if (num > 35000 && num < 65000) {
      const days = Math.round(num);
      const utcMs = Date.UTC(1899, 11, 30) + days * 86400000;
      const d = new Date(utcMs);
      const y = d.getUTCFullYear();
      const m = String(d.getUTCMonth() + 1).padStart(2, '0');
      const day = String(d.getUTCDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    }
  }

  // 2. Format like 10-October-26, 10-OCTOBER-2026, 7-October-26, 10-Oct-26, 10/Oct/2026
  const textMonthMatch = s.match(/^(\d{1,2})[\s\-\/\.]([a-zA-Z]+)[\s\-\/\.](\d{2,4})$/);
  if (textMonthMatch) {
    const day = parseInt(textMonthMatch[1], 10);
    const mStr = textMonthMatch[2].toLowerCase();
    let year = parseInt(textMonthMatch[3], 10);
    if (year < 100) year += 2000;
    const mNum = MONTH_NAMES[mStr];
    if (mNum && day >= 1 && day <= 31) {
      return `${year}-${String(mNum).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  // 3. Format like October 10, 2026 or Oct 10 2026
  const textMonthFirstMatch = s.match(/^([a-zA-Z]+)[\s\-\/\.](\d{1,2})[\s\-\/\.](\d{2,4})$/);
  if (textMonthFirstMatch) {
    const mStr = textMonthFirstMatch[1].toLowerCase();
    const day = parseInt(textMonthFirstMatch[2], 10);
    let year = parseInt(textMonthFirstMatch[3], 10);
    if (year < 100) year += 2000;
    const mNum = MONTH_NAMES[mStr];
    if (mNum && day >= 1 && day <= 31) {
      return `${year}-${String(mNum).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  // 4. Format like YYYY-MM-DD or YYYY/MM/DD
  const ymdMatch = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10);
    const month = parseInt(ymdMatch[2], 10);
    const day = parseInt(ymdMatch[3], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  // 5. Format like DD-MM-YYYY or DD/MM/YYYY or DD.MM.YYYY
  const dmyMatch = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (dmyMatch) {
    const p1 = parseInt(dmyMatch[1], 10);
    const p2 = parseInt(dmyMatch[2], 10);
    let year = parseInt(dmyMatch[3], 10);
    if (year < 100) year += 2000;
    let day = p1;
    let month = p2;
    if (p1 > 12 && p2 <= 12) {
      day = p1;
      month = p2;
    } else if (p2 > 12 && p1 <= 12) {
      day = p2;
      month = p1;
    }
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  // 6. Native parse fallback
  const parsed = Date.parse(s);
  if (!isNaN(parsed)) {
    const d = new Date(parsed);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  return null;
}

/**
 * Format ISO string YYYY-MM-DD into human-friendly "DD-MMM-YYYY" (e.g. 10-Oct-2026)
 */
export function formatIsoToDisplay(iso: string): string {
  if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso || '';
  const [y, mStr, dStr] = iso.split('-');
  const m = parseInt(mStr, 10);
  const d = parseInt(dStr, 10);
  const monthName = SHORT_MONTH_NAMES[m] || mStr;
  return `${d}-${monthName}-${y}`;
}

/**
 * Find the primary Date column in a row object or list of column names.
 * Prioritizes: CASE CREATED DATE / CASE CREATE DATE / CREATED DATE / LOGIN DATE
 */
export function pickPrimaryDateColumn(columns: string[]): string {
  if (!columns || columns.length === 0) return 'CASE CREATED DATE';

  // 1. Highest priority: "CASE CREATED DATE" variations
  const caseCreated = columns.find((c) => /case\s*creat(e|ed)?\s*date/i.test(c.trim()));
  if (caseCreated) return caseCreated;

  // 2. "CREATED DATE" / "CREATION DATE"
  const createdDate = columns.find((c) => /creat(e|ed|ion)?\s*date/i.test(c.trim()));
  if (createdDate) return createdDate;

  // 3. "CASE DATE"
  const caseDate = columns.find((c) => /case\s*date/i.test(c.trim()));
  if (caseDate) return caseDate;

  // 4. "LOGIN DATE"
  const loginDate = columns.find((c) => /login\s*date/i.test(c.trim()));
  if (loginDate) return loginDate;

  // 5. Any column with "date"
  const anyDate = columns.find((c) => /date/i.test(c.trim()));
  if (anyDate) return anyDate;

  return columns[0] || 'CASE CREATED DATE';
}

/**
 * Helper to extract and sanitize money numbers
 */
function parseAmount(val: any): number {
  if (val === null || val === undefined) return 0;
  const s = String(val).replace(/[^0-9.]/g, '');
  const num = parseFloat(s);
  return isNaN(num) ? 0 : num;
}

/** Today as yyyy-mm-dd on the server computer (used when the page sends no date). */
function localTodayIso(): string {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// =========================================================================
// 1. GET /api/mis-dashboard/sheets
// List all available MIS Google Sheets (connected in DB & system cache)
// =========================================================================
misDashboardRouter.get('/mis-dashboard/sheets', async (req: Request, res: Response) => {
  try {
    const db = await getDb();
    const sheetsRes = db.exec(`
      SELECT id, name, url, spreadsheet_id, sheet_names_json, rows_count, status, last_synced_at, created_at, priority_rank
      FROM connected_sheets
      ORDER BY COALESCE(priority_rank, 999) ASC, created_at DESC;
    `);

    const connectedSheets = (sheetsRes[0]?.values || []).map((r) => {
      let tabs: { name: string; rows: number; phones?: number }[] = [];
      try {
        const parsed = JSON.parse((r[4] as string) || '[]');
        tabs = Array.isArray(parsed)
          ? parsed.map((t: any) =>
              typeof t === 'string' ? { name: t, rows: 0 } : { name: t.name || 'Tab', rows: t.rows || 0, phones: t.phones || 0 }
            )
          : [];
      } catch {
        tabs = [];
      }

      return {
        id: r[0] as string,
        name: r[1] as string,
        url: r[2] as string,
        spreadsheetId: r[3] as string,
        tabs,
        rowsCount: (r[5] as number) || 0,
        status: (r[6] as string) || 'synced',
        lastSyncedAt: (r[7] as string) || null,
        createdAt: (r[8] as string) || null,
        priorityRank: (r[9] as number) || 1,
      };
    });

    // Also check cached files in data/cache if not already connected
    const cacheDir = path.resolve('./data/cache');
    const additionalFiles = [];
    if (fs.existsSync(cacheDir)) {
      const files = fs.readdirSync(cacheDir);
      for (const f of files) {
        if (!f.endsWith('.xlsx') && !f.endsWith('.xls')) continue;
        const isMaster = /master\s*mis/i.test(f);
        const cleanName = isMaster ? 'MASTER MIS' : f.replace(/\.[^/.]+$/, '');
        const alreadyConnected = connectedSheets.some(
          (cs) => cs.name.toLowerCase() === cleanName.toLowerCase() || cs.url.includes(f)
        );
        if (!alreadyConnected) {
          additionalFiles.push({
            id: `cached_${f}`,
            name: cleanName,
            url: `local://${f}`,
            spreadsheetId: f,
            tabs: [],
            rowsCount: 0,
            status: 'cached',
            lastSyncedAt: null,
            createdAt: null,
            priorityRank: 99,
          });
        }
      }
    }

    res.json({
      success: true,
      sheets: [...connectedSheets, ...additionalFiles],
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Failed to list MIS sheets' });
  }
});

// =========================================================================
// 2. GET /api/mis-dashboard/date-columns
// Detect all date columns in a specific sheet and tab
// =========================================================================
misDashboardRouter.get('/mis-dashboard/date-columns', async (req: Request, res: Response) => {
  const { sheetId, tabName } = req.query;
  try {
    const db = await getDb();
    let sql = `SELECT raw_row_json FROM search_index WHERE source_id = 'mis_sheet'`;
    const params: any[] = [];

    if (sheetId && sheetId !== 'all') {
      sql += ` AND file_id = ?`;
      params.push(sheetId);
    }
    if (tabName && tabName !== 'all') {
      sql += ` AND sheet_name = ?`;
      params.push(tabName);
    }
    sql += ` LIMIT 50;`;

    const resDb = db.exec(sql, params);
    const detectedCols = new Set<string>();

    if (resDb[0]?.values) {
      for (const row of resDb[0].values) {
        try {
          const json = JSON.parse(row[0] as string);
          for (const key of Object.keys(json)) {
            const val = json[key];
            if (/date|dt\b|day\b/i.test(key) || parseDateToIso(val) !== null) {
              detectedCols.add(key.trim());
            }
          }
        } catch {}
      }
    }

    const columnsList = Array.from(detectedCols);
    if (!columnsList.includes('CASE CREATED DATE')) {
      columnsList.unshift('CASE CREATED DATE');
    }
    const defaultColumn = pickPrimaryDateColumn(columnsList);

    res.json({
      success: true,
      columns: columnsList,
      defaultColumn,
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Failed to detect date columns' });
  }
});

// =========================================================================
// 3. GET /api/mis-dashboard/available-dates
// Returns all distinct dates present in that sheet/tab with case counts
// =========================================================================
misDashboardRouter.get('/mis-dashboard/available-dates', async (req: Request, res: Response) => {
  const sheetId = (req.query.sheetId as string) || 'mis_master_mis';
  const tabName = (req.query.tabName as string) || 'all';
  const dateColumn = (req.query.dateColumn as string) || 'CASE CREATED DATE';

  try {
    const db = await getDb();
    let sql = `SELECT sheet_name, raw_row_json FROM search_index WHERE source_id = 'mis_sheet'`;
    const params: any[] = [];

    if (sheetId && sheetId !== 'all') {
      sql += ` AND file_id = ?`;
      params.push(sheetId);
    }
    if (tabName && tabName !== 'all') {
      sql += ` AND sheet_name = ?`;
      params.push(tabName);
    }

    const resDb = db.exec(sql, params);
    const dateCounts: Record<string, { count: number; rawSample: string; tabCount: Record<string, number> }> = {};

    if (resDb[0]?.values) {
      for (const row of resDb[0].values) {
        const sName = row[0] as string;
        try {
          const json = JSON.parse(row[1] as string);
          // Try target dateColumn first, or fallback to any case created date variation
          let rawDate = json[dateColumn];
          if (!rawDate) {
            for (const k of Object.keys(json)) {
              if (k.toLowerCase().trim() === dateColumn.toLowerCase().trim()) {
                rawDate = json[k];
                break;
              }
            }
          }
          if (!rawDate && /case\s*creat/i.test(dateColumn)) {
            for (const k of Object.keys(json)) {
              if (/case\s*creat/i.test(k)) {
                rawDate = json[k];
                break;
              }
            }
          }

          const iso = parseDateToIso(rawDate);
          if (iso) {
            if (!dateCounts[iso]) {
              dateCounts[iso] = { count: 0, rawSample: String(rawDate).trim(), tabCount: {} };
            }
            dateCounts[iso].count += 1;
            dateCounts[iso].tabCount[sName] = (dateCounts[iso].tabCount[sName] || 0) + 1;
          }
        } catch {}
      }
    }

    // Sort descending by date (most recent first)
    const sortedDates = Object.keys(dateCounts)
      .sort((a, b) => b.localeCompare(a))
      .map((iso) => ({
        date: iso,
        displayDate: formatIsoToDisplay(iso),
        rawSample: dateCounts[iso].rawSample,
        count: dateCounts[iso].count,
        tabCount: dateCounts[iso].tabCount,
      }));

    res.json({
      success: true,
      availableDates: sortedDates,
      totalDatesCount: sortedDates.length,
      selectedColumn: dateColumn,
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Failed to fetch available dates' });
  }
});

// =========================================================================
// 4. POST /api/mis-dashboard/query
// Query cases matching the selected date in CASE CREATED DATE column
// =========================================================================
misDashboardRouter.post('/mis-dashboard/query', async (req: Request, res: Response) => {
  const {
    sheetId = 'mis_master_mis',
    tabName = 'all',
    date = localTodayIso(),
    dateColumn = 'CASE CREATED DATE',
    searchQuery = '',
    statusFilter = 'all',
    lenderFilter = 'all',
  } = req.body || {};

  try {
    const db = await getDb();

    // Standardize target date to ISO (YYYY-MM-DD)
    const targetIso = parseDateToIso(date);
    if (!targetIso) {
      return res.status(400).json({
        error: `Invalid date specified: "${date}". Please provide a valid calendar date (e.g. 2026-10-10 or 10-Oct-2026).`,
      });
    }

    let sql = `
      SELECT id, file_name, sheet_name, row_num, customer_name, phone_raw, lender, loan_amount, status_val, raw_row_json, file_id
      FROM search_index
      WHERE source_id = 'mis_sheet'
    `;
    const params: any[] = [];

    if (sheetId && sheetId !== 'all') {
      sql += ` AND file_id = ?`;
      params.push(sheetId);
    }
    if (tabName && tabName !== 'all') {
      sql += ` AND sheet_name = ?`;
      params.push(tabName);
    }

    sql += ` ORDER BY row_num ASC;`;

    const resDb = db.exec(sql, params);
    const matchedEntries: any[] = [];
    const dateHistogram: Record<string, { count: number; rawSample: string }> = {};
    const detectedColsSet = new Set<string>();

    let totalCases = 0;
    const statusBreakdown: Record<string, number> = {};
    const lenderBreakdown: Record<string, number> = {};
    let totalLoanAmount = 0;
    let totalNetLoanAmount = 0;
    const uniquePhonesSet = new Set<string>();
    const uniqueCustomersSet = new Set<string>();

    const searchLower = (searchQuery || '').trim().toLowerCase();
    const filterStatusLower = (statusFilter || 'all').trim().toLowerCase();
    const filterLenderLower = (lenderFilter || 'all').trim().toLowerCase();

    if (resDb[0]?.values) {
      for (const row of resDb[0].values) {
        const id = row[0] as number;
        const fileName = row[1] as string;
        const sName = row[2] as string;
        const rowNum = row[3] as number;
        const custName = (row[4] as string) || '';
        const phone = (row[5] as string) || '';
        const lenderDb = (row[6] as string) || '';
        const amountDb = (row[7] as string) || '';
        const statusDb = (row[8] as string) || '';
        const rawJsonStr = row[9] as string;
        const fileIdDb = (row[10] as string) || '';

        let rawObj: Record<string, any> = {};
        try {
          rawObj = JSON.parse(rawJsonStr);
        } catch {
          rawObj = {};
        }

        // Collect detected date columns
        for (const k of Object.keys(rawObj)) {
          if (/date|dt\b|day\b/i.test(k)) {
            detectedColsSet.add(k.trim());
          }
        }

        // Extract value for target date column
        let matchedColName = dateColumn;
        let cellDateVal = rawObj[dateColumn];

        if (cellDateVal === undefined || cellDateVal === null || String(cellDateVal).trim() === '') {
          // Case-insensitive fallback
          for (const k of Object.keys(rawObj)) {
            if (k.toLowerCase().trim() === dateColumn.toLowerCase().trim()) {
              cellDateVal = rawObj[k];
              matchedColName = k;
              break;
            }
          }
        }

        // If target is case created date, also check variations
        if ((!cellDateVal || String(cellDateVal).trim() === '') && /case\s*creat/i.test(dateColumn)) {
          for (const k of Object.keys(rawObj)) {
            if (/case\s*creat/i.test(k)) {
              cellDateVal = rawObj[k];
              matchedColName = k;
              break;
            }
          }
        }

        const rowIso = parseDateToIso(cellDateVal);
        if (rowIso) {
          if (!dateHistogram[rowIso]) {
            dateHistogram[rowIso] = { count: 0, rawSample: String(cellDateVal).trim() };
          }
          dateHistogram[rowIso].count += 1;
        }

        // Check if row matches target date!
        if (rowIso === targetIso) {
          // Extract comprehensive fields from rawObj
          const ticketId = String(rawObj['TICKET ID'] || rawObj['Ticket Id'] || rawObj['Ticket'] || rawObj['TICKET'] || '').trim();
          // MIS-only: read these straight from the sheet row (exact header, case-insensitive).
          // Blank in the sheet stays blank here - no fallback to database columns or F2 API.
          const sheetVal = (...names: string[]): string => {
            for (const n of names) {
              const want = n.toLowerCase().trim();
              for (const k of Object.keys(rawObj)) {
                if (k.toLowerCase().trim() === want) {
                  const v = rawObj[k];
                  return v === undefined || v === null ? '' : String(v).trim();
                }
              }
            }
            return '';
          };
          const customerName = sheetVal('CUSTOMER NAME', 'Customer Name', 'NAME') || custName;
          const phoneNumber = phone || sheetVal('PHONE NUMBER', 'Phone', 'MOBILE');
          const lender = sheetVal('LENDER', 'Bank', 'BANK NAME');
          const status = sheetVal('STATUS', 'CASE STATUS', 'CURRENT STATUS');
          const loanAmount = sheetVal('LOAN AMOUNT', 'Amount');
          const netLoanAmount = sheetVal('NET LOAN AMOUNT');
          const lastUpdated = String(rawObj['LAST UPDATED DATE'] || rawObj['Last Updated Date'] || '').trim();
          const approvalDate = String(rawObj['APPROVAL DATE'] || rawObj['Approval Date'] || '').trim();
          const disbursedDate = String(rawObj['DISBURSED DATE'] || rawObj['Disbursed Date'] || '').trim();
          const tlName = String(rawObj['TL NAME'] || rawObj['Team Leader'] || rawObj['TL'] || '').trim();
          const manager = String(rawObj['MANAGER'] || rawObj['Manager'] || '').trim();
          const empName = String(rawObj['EMP NAME'] || rawObj['Employee Name'] || rawObj['LOGIN POC'] || '').trim();
          const city = String(rawObj['CITY'] || rawObj['City'] || '').trim();
          const state = String(rawObj['LOCATION STATE'] || rawObj['State'] || '').trim();
          const branch = String(rawObj['BRANCH'] || rawObj['Branch'] || '').trim();
          const cibilScore = String(rawObj['CIBIL SCORE'] || rawObj['CIBIL'] || '').trim();
          const opsRemarks = String(rawObj['OPS REMARKS'] || rawObj['Ops Remarks'] || '').trim();
          const creditRemarks = String(rawObj['CREDIT REMARKS'] || rawObj['Credit Remarks'] || '').trim();
          const bankerName = String(rawObj['BANKER NAME'] || rawObj['Banker Name'] || '').trim();
          const bankerPhone = String(rawObj['BANKER PHONE NO'] || rawObj['Banker Phone'] || '').trim();
          const bankerEmail = String(rawObj['BANKER MAIL ID'] || rawObj['Banker Email'] || '').trim();

          // Apply Status filter if specified
          if (filterStatusLower !== 'all') {
            if (status.toLowerCase() !== filterStatusLower) {
              continue;
            }
          }

          // Apply Lender filter if specified
          if (filterLenderLower !== 'all') {
            if (lender.toLowerCase() !== filterLenderLower) {
              continue;
            }
          }

          // Apply search query filter
          if (searchLower) {
            const rowHaystack = [
              ticketId,
              customerName,
              phoneNumber,
              lender,
              status,
              tlName,
              manager,
              empName,
              city,
              state,
              branch,
              opsRemarks,
              creditRemarks,
              String(cellDateVal),
            ]
              .join(' ')
              .toLowerCase();
            if (!rowHaystack.includes(searchLower)) {
              continue;
            }
          }

          // Track KPIs
          totalCases += 1;
          const statusKey = (status || 'NO STATUS').toUpperCase();
          statusBreakdown[statusKey] = (statusBreakdown[statusKey] || 0) + 1;

          const lenderKey = (lender || 'UNKNOWN').toUpperCase();
          lenderBreakdown[lenderKey] = (lenderBreakdown[lenderKey] || 0) + 1;

          totalLoanAmount += parseAmount(loanAmount);
          totalNetLoanAmount += parseAmount(netLoanAmount);

          if (phoneNumber) uniquePhonesSet.add(phoneNumber.replace(/\D/g, '').slice(-10));
          if (customerName) uniqueCustomersSet.add(customerName.toUpperCase());

          matchedEntries.push({
            id,
            fileId: fileIdDb,
            recordKey: makeRecordKey({ fileId: fileIdDb, sheetName: sName, rowNum }),
            rowNum,
            ticketId,
            customerName,
            phone: phoneNumber,
            lender,
            loanAmount,
            netLoanAmount,
            status,
            caseCreatedDate: String(cellDateVal).trim(),
            lastUpdatedDate: lastUpdated,
            approvalDate,
            disbursedDate,
            tlName,
            manager,
            empName,
            city,
            state,
            branch,
            cibilScore,
            opsRemarks,
            creditRemarks,
            bankerName,
            bankerPhone,
            bankerEmail,
            rawDateValue: String(cellDateVal).trim(),
            normalizedDate: rowIso,
            matchedDateColumn: matchedColName,
            sheetName: sName,
            fileName,
            rawRow: rawObj,
          });
        }
      }
    }

    // Sort available dates descending
    const availableDates = Object.keys(dateHistogram)
      .sort((a, b) => b.localeCompare(a))
      .map((iso) => ({
        date: iso,
        displayDate: formatIsoToDisplay(iso),
        rawSample: dateHistogram[iso].rawSample,
        count: dateHistogram[iso].count,
      }));

    const detectedColsList = Array.from(detectedColsSet);
    if (!detectedColsList.includes('CASE CREATED DATE')) {
      detectedColsList.unshift('CASE CREATED DATE');
    }

    res.json({
      success: true,
      selectedDate: targetIso,
      selectedDisplayDate: formatIsoToDisplay(targetIso),
      selectedDateColumn: dateColumn,
      selectedSheetId: sheetId,
      selectedTab: tabName,
      totalCount: matchedEntries.length,
      entries: matchedEntries,
      availableDates,
      detectedDateColumns: detectedColsList,
      summary: {
        totalCases,
        statusBreakdown,
        lenderBreakdown,
        totalLoanAmount,
        totalNetLoanAmount,
        uniqueCustomers: uniqueCustomersSet.size,
        uniquePhones: uniquePhonesSet.size,
      },
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Failed to query MIS dashboard cases' });
  }
});

// =========================================================================
// 5. POST /api/mis-dashboard/export
// Download formatted Excel workbook of filtered entries on that date
// =========================================================================
misDashboardRouter.post('/mis-dashboard/export', async (req: Request, res: Response) => {
  const {
    sheetId = 'mis_master_mis',
    tabName = 'all',
    date = localTodayIso(),
    dateColumn = 'CASE CREATED DATE',
    searchQuery = '',
    statusFilter = 'all',
    lenderFilter = 'all',
  } = req.body || {};

  try {
    const targetIso = parseDateToIso(date);
    if (!targetIso) {
      return res.status(400).json({ error: 'Invalid date specified for export' });
    }

    // Call query logic internally
    const db = await getDb();
    let sql = `
      SELECT id, file_name, sheet_name, row_num, customer_name, phone_raw, lender, loan_amount, status_val, raw_row_json
      FROM search_index
      WHERE source_id = 'mis_sheet'
    `;
    const params: any[] = [];
    if (sheetId && sheetId !== 'all') {
      sql += ` AND file_id = ?`;
      params.push(sheetId);
    }
    if (tabName && tabName !== 'all') {
      sql += ` AND sheet_name = ?`;
      params.push(tabName);
    }
    sql += ` ORDER BY row_num ASC;`;

    const resDb = db.exec(sql, params);
    const matchedEntries: any[] = [];
    const searchLower = (searchQuery || '').trim().toLowerCase();
    const filterStatusLower = (statusFilter || 'all').trim().toLowerCase();
    const filterLenderLower = (lenderFilter || 'all').trim().toLowerCase();

    if (resDb[0]?.values) {
      for (const row of resDb[0].values) {
        const rawJsonStr = row[9] as string;
        let rawObj: any = {};
        try {
          rawObj = JSON.parse(rawJsonStr);
        } catch {
          rawObj = {};
        }

        let cellDateVal = rawObj[dateColumn];
        if (!cellDateVal) {
          for (const k of Object.keys(rawObj)) {
            if (k.toLowerCase().trim() === dateColumn.toLowerCase().trim()) {
              cellDateVal = rawObj[k];
              break;
            }
          }
        }
        if (!cellDateVal && /case\s*creat/i.test(dateColumn)) {
          for (const k of Object.keys(rawObj)) {
            if (/case\s*creat/i.test(k)) {
              cellDateVal = rawObj[k];
              break;
            }
          }
        }

        const rowIso = parseDateToIso(cellDateVal);
        if (rowIso === targetIso) {
          const status = String(rawObj['STATUS'] || rawObj['Status'] || row[8] || 'LOGIN').trim();
          const lender = String(rawObj['LENDER'] || rawObj['Lender'] || row[6] || '').trim();

          if (filterStatusLower !== 'all' && status.toLowerCase() !== filterStatusLower) continue;
          if (filterLenderLower !== 'all' && lender.toLowerCase() !== filterLenderLower) continue;

          if (searchLower) {
            const haystack = JSON.stringify(rawObj).toLowerCase();
            if (!haystack.includes(searchLower)) continue;
          }

          matchedEntries.push({
            rowNum: row[3],
            ticketId: rawObj['TICKET ID'] || '',
            customerName: rawObj['CUSTOMER NAME'] || row[4] || '',
            phone: rawObj['PHONE NUMBER'] || row[5] || '',
            caseCreatedDate: cellDateVal,
            lastUpdatedDate: rawObj['LAST UPDATED DATE'] || '',
            lender,
            loanAmount: rawObj['LOAN AMOUNT'] || row[7] || '',
            netLoanAmount: rawObj['NET LOAN AMOUNT'] || '',
            status,
            approvalDate: rawObj['APPROVAL DATE'] || '',
            disbursedDate: rawObj['DISBURSED DATE'] || '',
            tlName: rawObj['TL NAME'] || '',
            manager: rawObj['MANAGER'] || '',
            branch: rawObj['BRANCH'] || '',
            city: rawObj['CITY'] || '',
            opsRemarks: rawObj['OPS REMARKS'] || '',
            creditRemarks: rawObj['CREDIT REMARKS'] || '',
            sheetName: row[2],
          });
        }
      }
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'F2 Fintech MIS Dashboard';
    workbook.lastModifiedBy = 'F2 Fintech';
    workbook.created = new Date();

    const displayDate = formatIsoToDisplay(targetIso);
    const sheet = workbook.addWorksheet(`Cases ${displayDate}`.slice(0, 31));

    // Header Title block
    sheet.mergeCells('A1:P1');
    const titleCell = sheet.getCell('A1');
    titleCell.value = `F2 FINTECH MIS DASHBOARD — Cases Created on ${displayDate} (${matchedEntries.length} Total Records)`;
    titleCell.font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E293B' },
    };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.getRow(1).height = 36;

    // Column Headers
    const headers = [
      '#',
      'Ticket ID',
      'Case Created Date',
      'Customer Name',
      'Phone Number',
      'Lender',
      'Status',
      'Loan Amount',
      'Net Loan Amount',
      'TL Name',
      'Manager',
      'Branch',
      'City',
      'Last Updated Date',
      'Disbursed Date',
      'Tab Name',
    ];

    sheet.getRow(2).values = headers;
    sheet.getRow(2).height = 26;
    const headerRow = sheet.getRow(2);
    headerRow.eachCell((cell) => {
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0F766E' }, // Emerald dark
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'medium' },
        right: { style: 'thin' },
      };
    });

    // Populate Data Rows
    matchedEntries.forEach((entry, idx) => {
      const rowIdx = idx + 3;
      const r = sheet.getRow(rowIdx);
      r.values = [
        idx + 1,
        entry.ticketId,
        entry.caseCreatedDate,
        entry.customerName,
        entry.phone,
        entry.lender,
        entry.status,
        entry.loanAmount,
        entry.netLoanAmount,
        entry.tlName,
        entry.manager,
        entry.branch,
        entry.city,
        entry.lastUpdatedDate,
        entry.disbursedDate,
        entry.sheetName,
      ];
      r.height = 22;

      // Alternating row background
      const bgColor = idx % 2 === 0 ? 'FFFFFFFF' : 'FFF8FAFC';
      r.eachCell((cell, colNumber) => {
        cell.font = { name: 'Arial', size: 10 };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: bgColor },
        };
        cell.border = {
          top: { style: 'hair', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'hair', color: { argb: 'FFE2E8F0' } },
          left: { style: 'hair', color: { argb: 'FFE2E8F0' } },
          right: { style: 'hair', color: { argb: 'FFE2E8F0' } },
        };
        // Alignment
        if ([1, 2, 3, 5, 7, 14, 15].includes(colNumber)) {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        } else if ([8, 9].includes(colNumber)) {
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
        }
      });
    });

    // Set Column Widths
    sheet.columns = [
      { width: 6 },
      { width: 14 },
      { width: 20 },
      { width: 28 },
      { width: 16 },
      { width: 16 },
      { width: 16 },
      { width: 18 },
      { width: 18 },
      { width: 20 },
      { width: 20 },
      { width: 18 },
      { width: 18 },
      { width: 20 },
      { width: 20 },
      { width: 16 },
    ];

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="MIS_Cases_${targetIso}.xlsx"`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Export failed' });
  }
});

// =========================================================================
// 6. POST /api/mis-dashboard/sync-sheet
// Trigger immediate live refresh of the Google Sheet with database indexing
// =========================================================================
misDashboardRouter.post('/mis-dashboard/sync-sheet', async (req: Request, res: Response) => {
  const { sheetId = 'mis_master_mis' } = req.body || {};
  try {
    const db = await getDb();
    const sheetRes = db.exec('SELECT id, name, url, spreadsheet_id FROM connected_sheets WHERE id = ? LIMIT 1;', [sheetId]);
    if (!sheetRes.length || !sheetRes[0].values.length) {
      return res.status(404).json({ error: `Sheet with ID "${sheetId}" not found in system` });
    }

    const [id, name, url, spreadsheetId] = sheetRes[0].values[0] as [string, string, string, string];
    const cacheDir = path.resolve('./data/cache');
    let buffer: Buffer | null = null;
    let isLive = false;
    let isPrivateGSheet = false;
    let syncWarning: string | null = null;

    if (url && (url.startsWith('http://') || url.startsWith('https://'))) {
      try {
        const fetchMeta = await MisService.fetchGoogleSheetWithMeta(url, 2);
        buffer = fetchMeta.buffer;
        isLive = fetchMeta.isLive;
        if (!isLive) {
          isPrivateGSheet = true;
          syncWarning =
            'Google Sheet access is Restricted (HTTP 401). To enable automatic 30s live cloud syncing without manual uploads: in Google Sheets, click "Share" → change General Access to "Anyone with the link can view". Alternatively, use "Upload / Update OCT File" to update to 100+ rows directly.';
        }
      } catch (fetchErr: any) {
        syncWarning = fetchErr.message;
      }
    }

    // If Google Sheet was private or fetch threw error, fallback to local cache
    if (!buffer) {
      const candidates = [
        path.join(cacheDir, `gsheet_${spreadsheetId || sheetId}.xlsx`),
        path.join(cacheDir, `gsheet_${sheetId}.xlsx`),
        path.join(cacheDir, 'MASTER MIS.xlsx'),
        path.join(cacheDir, `${sheetId}.xlsx`),
      ];
      for (const p of candidates) {
        if (fs.existsSync(p)) {
          buffer = fs.readFileSync(p);
          break;
        }
      }
    }

    if (!buffer) {
      return res.status(404).json({
        error: `No workbook buffer available for "${name}". Please upload an Excel or CSV file.`,
      });
    }

    // Save active copy to disk cache
    try {
      if (name.toLowerCase().includes('master mis')) {
        fs.writeFileSync(path.join(cacheDir, 'MASTER MIS.xlsx'), buffer);
      }
      fs.writeFileSync(path.join(cacheDir, `gsheet_${sheetId}.xlsx`), buffer);
    } catch {}

    // MIS dashboard uses ONLY the Google Sheet / uploaded MIS data (no F2 API rows are merged in).
    const octResult: any = null;

    // Index workbook completely into SQLite database if not already indexed by octSync
    let totalIndexed = 0;
    let sheetNames: string[] = [];
    let tabDetails: { name: string; rows: number; phones?: number }[] = [];
    let newRowsCount = 0;
    let changedRowsCount = 0;

    if (octResult) {
      totalIndexed = octResult.totalIndexed;
      newRowsCount = octResult.newRowsAdded;
      changedRowsCount = octResult.updatedRowsCount;
      const shRes = db.exec('SELECT sheet_names_json FROM connected_sheets WHERE id = ?;', [sheetId]);
      try {
        tabDetails = JSON.parse((shRes[0]?.values[0]?.[0] as string) || '[]');
      } catch {
        tabDetails = [];
      }
    } else {
      const idxRes = indexWorkbookIntoDatabase(db, sheetId, name, buffer);
      totalIndexed = idxRes.totalIndexed;
      sheetNames = idxRes.sheetNames;
      tabDetails = idxRes.tabDetails;
      newRowsCount = idxRes.newRowsCount;
      changedRowsCount = idxRes.changedRowsCount;

      const now = new Date().toISOString();
      db.run(
        'UPDATE connected_sheets SET rows_count = ?, last_synced_at = ?, status = "synced", sheet_names_json = ? WHERE id = ?;',
        [totalIndexed, now, JSON.stringify(tabDetails), sheetId]
      );
      saveDb();
    }

    // Find OCT tab details if exists
    const octTab = tabDetails.find((t) => t.name.toUpperCase().includes('OCT'));
    const finalOctRows = octResult ? octResult.octRowsCount : (octTab ? octTab.rows : null);

    res.json({
      success: true,
      isLive: true,
      isPrivateGSheet,
      warning: syncWarning,
      totalIndexed,
      newRowsCount,
      changedRowsCount,
      tabDetails,
      octRowsCount: finalOctRows,
      message: octResult
        ? `Live Auto-Sync complete: OCT 26 synchronized with ${finalOctRows} rows (+${octResult.newRowsAdded} new cases added, ${octResult.updatedRowsCount} updated).`
        : `Synchronized ${totalIndexed.toLocaleString()} records (${octTab ? `${octTab.name}: ${octTab.rows} rows` : ''}). ${syncWarning || ''}`,
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Failed to sync MIS sheet' });
  }
});

// =========================================================================
// 7. POST /api/mis-dashboard/upload-oct-file
// Directly upload or update the OCT file (.xlsx or .csv) with all ~100 rows
// =========================================================================
misDashboardRouter.post('/mis-dashboard/upload-oct-file', async (req: Request, res: Response) => {
  const { fileName = 'OCT_26.xlsx', fileBase64, sheetId = 'mis_master_mis', targetTab = 'OCT 26' } = req.body || {};
  if (!fileBase64) {
    return res.status(400).json({ error: 'File content (base64) is required' });
  }

  try {
    const db = await getDb();
    const cleanBase64 = fileBase64.replace(/^data:.*?;base64,/, '');
    const buffer = Buffer.from(cleanBase64, 'base64');
    if (buffer.length < 50) {
      return res.status(400).json({ error: 'Uploaded file is empty or corrupted' });
    }

    const cacheDir = path.resolve('./data/cache');
    if (!fs.existsSync(cacheDir)) fs.mkdirSync(cacheDir, { recursive: true });

    // Parse uploaded workbook
    const uploadedWb = XLSX.read(buffer, { type: 'buffer', cellDates: true, dense: true });
    const uploadedSheetNames = uploadedWb.SheetNames || [];
    if (uploadedSheetNames.length === 0) {
      return res.status(400).json({ error: 'No worksheets found in uploaded file' });
    }

    const masterMisPath = path.join(cacheDir, 'MASTER MIS.xlsx');

    // Merge the upload into the MASTER MIS workbook so no tab is lost:
    //  - one sheet (an exported OCT tab or a CSV): it becomes the OCT tab
    //  - many sheets: EVERY uploaded tab is added or replaced, and tabs that were already
    //    in MASTER MIS (SEPT, AUG, ...) but are not in the upload are kept
    let masterWb: any = null;
    if (fs.existsSync(masterMisPath)) {
      try {
        masterWb = XLSX.readFile(masterMisPath, { cellDates: true });
      } catch {}
    }
    if (!masterWb) {
      masterWb = XLSX.utils.book_new();
    }
    // read again without dense mode so the sheets can be mixed into the master workbook
    const plainUploadedWb = XLSX.read(buffer, { type: 'buffer', cellDates: true });

    if (uploadedSheetNames.length === 1) {
      const octSourceWs = plainUploadedWb.Sheets[uploadedSheetNames[0]];
      const resolvedTabName =
        masterWb.SheetNames.find((s: string) => s.toUpperCase().includes('OCT')) || targetTab || 'OCT 26';
      masterWb.Sheets[resolvedTabName] = octSourceWs;
      if (!masterWb.SheetNames.includes(resolvedTabName)) {
        masterWb.SheetNames.splice(1, 0, resolvedTabName);
      }
    } else {
      for (const tabName of plainUploadedWb.SheetNames) {
        masterWb.Sheets[tabName] = plainUploadedWb.Sheets[tabName];
        if (!masterWb.SheetNames.includes(tabName)) masterWb.SheetNames.push(tabName);
      }
    }
    const finalMasterBuf: Buffer = XLSX.write(masterWb, { type: 'buffer', bookType: 'xlsx' });

    // Persist to disk cache
    fs.writeFileSync(masterMisPath, finalMasterBuf);
    fs.writeFileSync(path.join(cacheDir, `gsheet_${sheetId}.xlsx`), finalMasterBuf);
    if (fileName) {
      try {
        fs.writeFileSync(path.join(cacheDir, fileName), buffer);
      } catch {}
    }

    // Re-index into SQLite
    const { totalIndexed, sheetNames, tabDetails, newRowsCount, changedRowsCount } =
      indexWorkbookIntoDatabase(db, sheetId, 'MASTER MIS', finalMasterBuf);

    const now = new Date().toISOString();
    db.run(
      'UPDATE connected_sheets SET rows_count = ?, last_synced_at = ?, status = "synced", sheet_names_json = ? WHERE id = ?;',
      [totalIndexed, now, JSON.stringify(tabDetails), sheetId]
    );
    saveDb();

    const targetTabDetail = tabDetails.find(
      (t) => t.name.toUpperCase().includes('OCT') || t.name === targetTab
    );

    res.json({
      success: true,
      message: `Successfully updated ${targetTabDetail ? targetTabDetail.name : 'OCT 26'} with ${targetTabDetail ? targetTabDetail.rows : totalIndexed} rows! All customer details, phone numbers, and loan statuses are active and searchable.`,
      rowsCount: targetTabDetail ? targetTabDetail.rows : totalIndexed,
      totalIndexed,
      newRowsCount,
      changedRowsCount,
      tabDetails,
      activeTab: targetTabDetail ? targetTabDetail.name : 'OCT 26',
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Failed to upload and update OCT file' });
  }
});

