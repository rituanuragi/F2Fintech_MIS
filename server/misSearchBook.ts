import { Router, Request, Response } from 'express';
import * as fs from 'fs';
import * as path from 'path';
import * as XLSXNamespace from 'xlsx';
const XLSX: any = (XLSXNamespace as any).default || XLSXNamespace;
import ExcelJS from 'exceljs';
import { getDb } from './db.js';
import {
  normalizePhone,
  analyzeRowSegment,
  formatExcelDateValue,
  extractMonthYearLabel,
  isMeaningfulCellValue,
  rowHasAnyEntry,
  isValidPhoneCellValue,
  rowHasPhoneNumber,
} from './normalizer.js';
import { determineLeadPriority, getPriorityRules, getMisPriorityRules, LeadPriorityRule } from './priority.js';
import { MisService } from './mis.js';

export const misSearchBookRouter = Router();

export interface MisSearchBookSheetInfo {
  name: string;
  hasPhoneColumn: boolean;
  phoneColumnName: string | null;
  phoneColIndex: number;
  phoneColLetter?: string;
  totalRows: number;
  totalCols?: number;
  samplePhones: string[];
}

// Convert 0-indexed column number to Excel column letter (e.g. 17 -> 'R', 12 -> 'M')
function colIndexToLetter(idx: number): string {
  if (idx < 0) return '';
  let letter = '';
  let temp = idx;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

// Helper to locate workbook buffer or cached file path
async function getWorkbookBuffer(
  sheetId?: string,
  forceLiveFetch: boolean = false
): Promise<{ buffer: Buffer; workbookTitle: string; sheetRecord: any }> {
  const db = await getDb();
  let sheetRecord: any = null;

  if (sheetId) {
    const sRes = db.exec('SELECT id, name, url, spreadsheet_id, sheet_names_json, rows_count, last_synced_at FROM connected_sheets WHERE id = ? LIMIT 1;', [sheetId]);
    if (sRes.length && sRes[0].values.length) {
      const v = sRes[0].values[0];
      sheetRecord = { id: v[0], name: v[1], url: v[2], spreadsheetId: v[3], sheetNamesJson: v[4], rowsCount: v[5], lastSyncedAt: v[6] };
    }
  }

  if (!sheetRecord) {
    // Default to the highest priority / first synced connected sheet
    const sRes = db.exec('SELECT id, name, url, spreadsheet_id, sheet_names_json, rows_count, last_synced_at FROM connected_sheets ORDER BY COALESCE(priority_rank, 999) ASC, created_at DESC LIMIT 1;');
    if (sRes.length && sRes[0].values.length) {
      const v = sRes[0].values[0];
      sheetRecord = { id: v[0], name: v[1], url: v[2], spreadsheetId: v[3], sheetNamesJson: v[4], rowsCount: v[5], lastSyncedAt: v[6] };
    }
  }

  if (!sheetRecord) {
    throw new Error('No Google Sheet MIS workbook has been connected or synced yet. Please connect one under Google Sheets MIS.');
  }

  // If forceLiveFetch is requested and sheet has a live Google Sheet URL, pull fresh from Google Sheets first
  if (forceLiveFetch && sheetRecord.url && !sheetRecord.url.startsWith('local://')) {
    try {
      const buf = await MisService.fetchGoogleSheetAsXlsx(sheetRecord.url, 2);
      return { buffer: buf, workbookTitle: sheetRecord.name, sheetRecord };
    } catch (err: any) {
      console.warn(`Live refresh fallback to cache for ${sheetRecord.name}: ${err.message || 'Network notice'}`);
    }
  }

  const cacheDir = path.resolve('./data/cache');
  const candidateFiles = [
    path.join(cacheDir, `gsheet_${sheetRecord.spreadsheetId}.xlsx`),
    path.join(cacheDir, `gsheet_${sheetRecord.id}.xlsx`),
    path.join(cacheDir, `${sheetRecord.name}.xlsx`),
    path.join(cacheDir, `${sheetRecord.name}`),
    path.join(cacheDir, 'MASTER MIS.xlsx'),
  ];

  // If local:// url, check that filename first
  if (sheetRecord.url && sheetRecord.url.startsWith('local://')) {
    const localName = sheetRecord.url.replace('local://', '');
    candidateFiles.unshift(path.join(cacheDir, localName));
  }

  // Also check if any file in cache matches spreadsheetId or name
  if (fs.existsSync(cacheDir)) {
    const allFiles = fs.readdirSync(cacheDir);
    const matched = allFiles.find(
      (f) =>
        f.endsWith('.xlsx') &&
        (f.includes(sheetRecord.spreadsheetId) ||
          f.toLowerCase().includes(sheetRecord.name.toLowerCase()) ||
          f.toLowerCase().includes('master mis'))
    );
    if (matched) {
      candidateFiles.push(path.join(cacheDir, matched));
    }

    // Always include any gsheet_*.xlsx in cache as candidate
    const anyGsheet = allFiles.find((f) => f.startsWith('gsheet_') && f.endsWith('.xlsx'));
    if (anyGsheet) {
      candidateFiles.push(path.join(cacheDir, anyGsheet));
    }
  }

  // Pick the newest existing candidate file so newly synced sheets are always used immediately
  const existingCandidates: { filePath: string; mtimeMs: number }[] = [];
  const seenPaths = new Set<string>();
  for (const cPath of candidateFiles) {
    if (seenPaths.has(cPath)) continue;
    seenPaths.add(cPath);
    if (fs.existsSync(cPath)) {
      try {
        const stat = fs.statSync(cPath);
        if (stat.size > 50) {
          existingCandidates.push({ filePath: cPath, mtimeMs: stat.mtimeMs });
        }
      } catch {}
    }
  }

  if (existingCandidates.length > 0) {
    existingCandidates.sort((a, b) => b.mtimeMs - a.mtimeMs);
    const buf = fs.readFileSync(existingCandidates[0].filePath);
    return { buffer: buf, workbookTitle: sheetRecord.name, sheetRecord };
  }

  // If not in cache, attempt to fetch from Google Sheets
  if (sheetRecord.url && !sheetRecord.url.startsWith('local://')) {
    try {
      const buf = await MisService.fetchGoogleSheetAsXlsx(sheetRecord.url);
      return { buffer: buf, workbookTitle: sheetRecord.name, sheetRecord };
    } catch (err: any) {
      console.warn(`Unable to fetch Google Sheet MIS: ${err.message || 'Network error'}`);
    }
  }

  // Final fallback: any valid .xlsx file in cache that contains multiple sheets
  if (fs.existsSync(cacheDir)) {
    const allFiles = fs.readdirSync(cacheDir).filter(f => f.endsWith('.xlsx'));
    for (const f of allFiles) {
      try {
        const fullPath = path.join(cacheDir, f);
        const buf = fs.readFileSync(fullPath);
        return { buffer: buf, workbookTitle: f.replace(/\.xlsx$/i, ''), sheetRecord };
      } catch {}
    }
  }

  throw new Error(`Could not find cached file for Google Sheet MIS: ${sheetRecord.name}`);
}

/**
 * Checks if a cell value is empty, pure whitespace, or standard spreadsheet placeholder format
 * (e.g. '-', '--', '—', 'null', 'nan', '#N/A', etc.).
 */
export function isPlaceholderCell(val: any): boolean {
  return !isMeaningfulCellValue(val);
}

/**
 * Determines whether a spreadsheet row contains actual meaningful data entry,
 * ignoring empty table rows, blank rows, and rows that only contain formula/format defaults (like '-').
 */
export function isMeaningfulDataRow(row: any[], headers: string[] = []): boolean {
  return rowHasAnyEntry(row, headers);
}

// In-memory cache for sheets inspection
let cachedSheetsInfo: { sheetId: string; timestamp: number; data: any } | null = null;

export function invalidateMisSearchBookCache(): void {
  cachedSheetsInfo = null;
}

// Upload an MIS workbook directly (e.g. user attaches / uploads any file)
misSearchBookRouter.post('/mis-search-book/upload-workbook', async (req: Request, res: Response) => {
  const { fileName, fileBase64 } = req.body;
  if (!fileBase64) {
    return res.status(400).json({ error: 'File content (base64) is required' });
  }

  try {
    const cleanBase64 = fileBase64.replace(/^data:.*?;base64,/, '');
    const buffer = Buffer.from(cleanBase64, 'base64');
    if (buffer.length < 50) {
      return res.status(400).json({ error: 'Uploaded file is empty or corrupted.' });
    }

    const cleanFileName = fileName || 'Uploaded_MIS_Workbook.xlsx';
    const displayName = cleanFileName.replace(/\.[^/.]+$/, '');
    const sheetId = `upload_${Date.now()}`;
    const cacheDir = path.resolve('./data/cache');
    if (!fs.existsSync(cacheDir)) fs.mkdirSync(cacheDir, { recursive: true });

    // Save to cache
    fs.writeFileSync(path.join(cacheDir, cleanFileName), buffer);
    fs.writeFileSync(path.join(cacheDir, `${sheetId}.xlsx`), buffer);
    fs.writeFileSync(path.join(cacheDir, 'MASTER MIS.xlsx'), buffer);

    const db = await getDb();
    const wb = XLSX.read(buffer, { dense: true });
    const sheetNames: string[] = wb.SheetNames || [];
    const now = new Date().toISOString();

    // Register in connected_sheets
    db.run(
      `INSERT OR REPLACE INTO connected_sheets (id, name, url, spreadsheet_id, sheet_names_json, rows_count, status, last_synced_at, created_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);`,
      [
        sheetId,
        displayName,
        `local://${cleanFileName}`,
        sheetId,
        JSON.stringify(sheetNames),
        1000,
        'synced',
        now,
        now,
      ]
    );

    // Invalidate cache
    cachedSheetsInfo = null;

    res.json({
      success: true,
      sheetId,
      fileName: cleanFileName,
      displayName,
      sheetNames,
      message: `Workbook "${displayName}" uploaded successfully with ${sheetNames.length} sheets!`,
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Failed to upload MIS workbook' });
  }
});

// 1. Get info on Google Sheet MIS workbook and all sheets present
misSearchBookRouter.get('/mis-search-book/sheets', async (req: Request, res: Response) => {
  try {
    const sheetId = req.query.sheetId as string | undefined;
    const db = await getDb();

    // Check in-memory cache (valid for 30 seconds unless refresh is requested)
    const isRefresh = req.query.refresh === 'true' || req.query.fresh === 'true';
    const effectiveSheetId = sheetId || 'default';
    if (
      !isRefresh &&
      cachedSheetsInfo &&
      cachedSheetsInfo.sheetId === effectiveSheetId &&
      Date.now() - cachedSheetsInfo.timestamp < 30000
    ) {
      return res.json(cachedSheetsInfo.data);
    }

    // Get list of all available connected sheets
    const allSheetsRes = db.exec('SELECT id, name, url, spreadsheet_id, rows_count, last_synced_at FROM connected_sheets ORDER BY created_at DESC;');
    const availableWorkbooks = (allSheetsRes[0]?.values || []).map(r => ({
      id: r[0] as string,
      name: r[1] as string,
      url: r[2] as string,
      spreadsheetId: r[3] as string,
      rowsCount: r[4] as number,
      lastSyncedAt: r[5] as string,
    }));

    const { buffer, workbookTitle, sheetRecord } = await getWorkbookBuffer(sheetId, isRefresh);
    // Read workbook to inspect sheet names, accurate row counts, and phone column positions
    const wb = XLSX.read(buffer, { dense: true });
    const sheetNames: string[] = wb.SheetNames || [];

    const cachedSheetCounts: Record<string, { rows: number; phones: number }> = {};
    if (sheetRecord?.sheet_names_json) {
      try {
        const parsed = typeof sheetRecord.sheet_names_json === 'string'
          ? JSON.parse(sheetRecord.sheet_names_json)
          : sheetRecord.sheet_names_json;
        if (Array.isArray(parsed)) {
          for (const item of parsed) {
            if (item.name) {
              cachedSheetCounts[item.name] = { rows: item.rows || 0, phones: item.phones || 0 };
            }
          }
        }
      } catch {}
    }

    const sheetsInfo: MisSearchBookSheetInfo[] = [];

    for (const name of sheetNames) {
      const ws = wb.Sheets[name];
      if (!ws) {
        sheetsInfo.push({
          name,
          hasPhoneColumn: false,
          phoneColumnName: null,
          phoneColIndex: -1,
          phoneColLetter: '',
          totalRows: 0,
          totalCols: 0,
          samplePhones: [],
        });
        continue;
      }

      // Check grid dimensions if available
      let totalCols = 0;
      if (ws['!ref']) {
        try {
          const range = XLSX.utils.decode_range(ws['!ref']);
          totalCols = range.e.c + 1;
        } catch {}
      }

      // Fast header search in first 10 rows
      let foundHeaderRow = -1;
      let phoneColIndex = -1;
      let detectedMaxCols = 0;

      // Scan first 10 rows for headers
      for (let r = 0; r < 10 && ws[r]; r++) {
        const row = ws[r];
        if (!row || !Array.isArray(row)) continue;
        for (let c = 0; c < row.length; c++) {
          const cellVal = row[c]?.v !== undefined ? String(row[c].v).trim() : '';
          if (cellVal.toUpperCase() === 'PHONE NUMBER') {
            foundHeaderRow = r;
            phoneColIndex = c;
            break;
          }
        }
        if (phoneColIndex !== -1) break;
      }

      if (phoneColIndex === -1 && !ws[0]) {
        const topRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, range: 10, blankrows: false }) as any[][];
        for (let r = 0; r < Math.min(topRows.length, 10); r++) {
          const row = topRows[r] || [];
          for (let c = 0; c < row.length; c++) {
            const cellStr = String(row[c] || '').trim();
            if (cellStr.toUpperCase() === 'PHONE NUMBER') {
              foundHeaderRow = r;
              phoneColIndex = c;
              break;
            }
          }
          if (phoneColIndex !== -1) break;
        }
      }

      // CRITICAL: If sheet has a phone column, count ONLY those rows which have phone numbers!
      // Otherwise, count rows that have any genuine entry.
      let validRowCount = 0;
      const startR = foundHeaderRow >= 0 ? foundHeaderRow + 1 : 1;

      // Extract sheet header strings
      const sheetHeaders: string[] = [];
      const headerRowObj = foundHeaderRow >= 0 && ws[foundHeaderRow] ? ws[foundHeaderRow] : (ws[0] || []);
      if (Array.isArray(headerRowObj)) {
        for (let c = 0; c < headerRowObj.length; c++) {
          const val = headerRowObj[c]?.v !== undefined ? headerRowObj[c].v : headerRowObj[c];
          sheetHeaders.push(String(val ?? '').trim());
        }
      }

      const hasPhoneCol = phoneColIndex !== -1;

      if (ws[0] !== undefined || ws[1] !== undefined) {
        // Dense sheet format
        for (let r = startR; ws[r] !== undefined; r++) {
          const row = ws[r];
          if (!row || !Array.isArray(row)) continue;
          if (rowHasAnyEntry(row, sheetHeaders)) {
            validRowCount++;
            if (row.length > detectedMaxCols) detectedMaxCols = row.length;
          }
        }
      } else {
        // Non-dense format fallback using sheet_to_json with blankrows: false
        const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: false }) as any[][];
        for (let r = startR; r < allRows.length; r++) {
          const row = allRows[r];
          if (!row || !Array.isArray(row)) continue;
          if (rowHasAnyEntry(row, sheetHeaders)) {
            validRowCount++;
            if (row.length > detectedMaxCols) detectedMaxCols = row.length;
          }
        }
      }

      // Extract sample phone numbers ONLY from rows that genuinely have phone numbers
      const samplePhones: string[] = [];
      if (phoneColIndex !== -1) {
        if (ws[0] !== undefined || ws[1] !== undefined) {
          for (let r = startR; ws[r] !== undefined && samplePhones.length < 5; r++) {
            const row = ws[r];
            if (!row || !rowHasPhoneNumber(row, phoneColIndex)) continue;
            let val = row[phoneColIndex]?.v;
            if (!isValidPhoneCellValue(val)) {
              for (let c = 0; c < row.length; c++) {
                if (isValidPhoneCellValue(row[c]?.v)) {
                  val = row[c]?.v;
                  break;
                }
              }
            }
            if (val !== undefined && val !== null) {
              const digits = String(val).trim().replace(/\D/g, '');
              if (digits.length >= 7) {
                samplePhones.push(digits.length >= 10 ? digits.slice(-10) : digits);
              }
            }
          }
        } else {
          const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: false }) as any[][];
          for (let r = startR; r < allRows.length && samplePhones.length < 5; r++) {
            const row = allRows[r];
            if (!row || !rowHasPhoneNumber(row, phoneColIndex)) continue;
            let val = row[phoneColIndex];
            if (!isValidPhoneCellValue(val)) {
              for (let c = 0; c < row.length; c++) {
                if (isValidPhoneCellValue(row[c])) {
                  val = row[c];
                  break;
                }
              }
            }
            if (val !== undefined && val !== null) {
              const digits = String(val).trim().replace(/\D/g, '');
              if (digits.length >= 7) {
                samplePhones.push(digits.length >= 10 ? digits.slice(-10) : digits);
              }
            }
          }
        }
      }

      // If samplePhones empty, check search_index
      if (samplePhones.length === 0) {
        try {
          const dbSample = db.exec(
            `SELECT phone_raw FROM search_index WHERE sheet_name = ? AND phone_raw IS NOT NULL AND phone_raw != '' LIMIT 5;`,
            [name]
          );
          if (dbSample.length && dbSample[0].values.length) {
            for (const v of dbSample[0].values) {
              const digits = String(v[0] || '').replace(/\D/g, '');
              if (digits.length >= 10) {
                samplePhones.push(digits.slice(-10));
              }
            }
            if (samplePhones.length > 0) {
              phoneColIndex = phoneColIndex !== -1 ? phoneColIndex : 17;
            }
          }
        } catch {}
      }

      const finalColLetter = phoneColIndex !== -1 ? colIndexToLetter(phoneColIndex) : (samplePhones.length > 0 ? 'R' : '');

      sheetsInfo.push({
        name,
        hasPhoneColumn: phoneColIndex !== -1 || samplePhones.length > 0,
        phoneColumnName: 'PHONE NUMBER',
        phoneColIndex,
        phoneColLetter: finalColLetter,
        totalRows: validRowCount,
        totalCols: detectedMaxCols || (totalCols > 0 ? totalCols : 0),
        samplePhones,
      });
    }

    // Persist verified true row counts back to connected_sheets in database
    try {
      const cleanTabs = sheetsInfo.map(s => ({
        name: s.name,
        rows: s.totalRows,
        phones: s.hasPhoneColumn ? s.totalRows : 0,
      }));
      const totalCleanRows = sheetsInfo.reduce((acc, s) => acc + s.totalRows, 0);
      db.run('UPDATE connected_sheets SET rows_count = ?, sheet_names_json = ? WHERE id = ?;', [
        totalCleanRows,
        JSON.stringify(cleanTabs),
        sheetRecord.id,
      ]);
    } catch (e) {
      console.warn('Failed to update connected_sheets counts:', e);
    }

    const responsePayload = {
      success: true,
      workbookInfo: {
        id: sheetRecord.id,
        name: workbookTitle,
        url: sheetRecord.url,
        spreadsheetId: sheetRecord.spreadsheetId,
        lastSyncedAt: sheetRecord.lastSyncedAt,
        totalSheets: sheetNames.length,
      },
      availableWorkbooks,
      sheets: sheetsInfo,
    };

    cachedSheetsInfo = {
      sheetId: effectiveSheetId,
      timestamp: Date.now(),
      data: responsePayload,
    };

    res.json(responsePayload);
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Failed to inspect MIS workbook' });
  }
});

export interface ProcessedPhoneItem {
  phone: string;
  normPhone: string;
  date?: string;
  priority: string;
  priorityRank: number;
  priorityCategory: string;
  badgeBg?: string;
  badgeText?: string;
  badgeBorder?: string;
  workbook: string;
  sheetName: string;
  customerName?: string;
  matchCount: number;
  found: boolean;
}

// 2. Process selected sheets: Extract phone numbers from 'PHONE NUMBER' column and perform global multi-workbook search
misSearchBookRouter.post('/mis-search-book/process', async (req: Request, res: Response) => {
  const { sheetId, selectedSheets, priorityRules: customRules } = req.body;
  if (!selectedSheets || !Array.isArray(selectedSheets) || selectedSheets.length === 0) {
    return res.status(400).json({ error: 'Please select at least one sheet to process.' });
  }

  try {
    const startTime = Date.now();
    const db = await getDb();
    const priorityRules: LeadPriorityRule[] = (customRules && Array.isArray(customRules) && customRules.length > 0)
      ? customRules
      : getMisPriorityRules();

    const { buffer, workbookTitle } = await getWorkbookBuffer(sheetId);
    const wb = XLSX.read(buffer, { dense: true, sheets: selectedSheets });

    // Step 1: For each selected sheet, extract phone numbers and case created dates
    const sheetPhonesMap: Record<string, { rawPhone: string; normPhone: string; date: string }[]> = {};
    const allNormPhonesSet = new Set<string>();

    for (const sheetName of selectedSheets) {
      const ws = wb.Sheets[sheetName];
      if (!ws) continue;

      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, blankrows: false }) as any[][];
      if (!rows || rows.length === 0) continue;

      // Find the exact 'PHONE NUMBER' column and 'Case Created date' column
      let phoneColIndex = -1;
      let dateColIndex = -1;
      let headerRowIndex = 0;

      for (let r = 0; r < Math.min(rows.length, 5); r++) {
        const row = rows[r] || [];
        for (let c = 0; c < row.length; c++) {
          const cellStr = String(row[c] || '').trim();
          if (cellStr === 'PHONE NUMBER') {
            headerRowIndex = r;
            phoneColIndex = c;
          }
          if (
            ['Case Created date', 'CASE CREATED DATE', 'Case Creat date', 'CASE CREAT DATE', 'Case Create date', 'CASE CREATE DATE'].includes(cellStr)
          ) {
            dateColIndex = c;
          }
        }
        if (phoneColIndex !== -1) break;
      }

      if (phoneColIndex === -1) {
        // Fallback: If exact case-sensitive 'PHONE NUMBER' is uppercase
        for (let r = 0; r < Math.min(rows.length, 5); r++) {
          const row = rows[r] || [];
          for (let c = 0; c < row.length; c++) {
            const cellStr = String(row[c] || '').trim();
            if (cellStr.toUpperCase() === 'PHONE NUMBER') {
              headerRowIndex = r;
              phoneColIndex = c;
              break;
            }
          }
          if (phoneColIndex !== -1) break;
        }
      }

      // Fallback for Date column if not found on initial scan
      if (dateColIndex === -1 && rows[headerRowIndex]) {
        const hRow = rows[headerRowIndex] || [];
        for (let c = 0; c < hRow.length; c++) {
          const cellStr = String(hRow[c] || '').trim();
          if (/case.*creat.*date/i.test(cellStr) || /case.*create/i.test(cellStr) || /creat.*date/i.test(cellStr)) {
            dateColIndex = c;
            break;
          }
        }
      }

      // Also search flexible phone columns if exact PHONE NUMBER not matched
      if (phoneColIndex === -1) {
        for (let r = 0; r < Math.min(rows.length, 5); r++) {
          const row = rows[r] || [];
          for (let c = 0; c < row.length; c++) {
            const cellStr = String(row[c] || '').trim().toLowerCase();
            if (
              cellStr === 'phone' ||
              cellStr === 'phone no' ||
              cellStr === 'mobile' ||
              cellStr === 'mobile no' ||
              cellStr === 'contact' ||
              cellStr === 'contact no' ||
              cellStr.includes('phone')
            ) {
              headerRowIndex = r;
              phoneColIndex = c;
              break;
            }
          }
          if (phoneColIndex !== -1) break;
        }
      }

      const extracted: { rawPhone: string; normPhone: string; date: string }[] = [];

      for (let r = headerRowIndex + 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || !Array.isArray(row)) continue;

        // Include all rows that have genuine entries across all columns
        const headersList = (rows[headerRowIndex] || []).map((h: any, i: number) => String(h || `Column_${i + 1}`));
        if (!rowHasAnyEntry(row, headersList)) continue;

        let cellVal = phoneColIndex !== -1 ? row[phoneColIndex] : '';
        let rawPhone = cellVal !== undefined && cellVal !== null ? String(cellVal).trim() : '';
        if (!isValidPhoneCellValue(rawPhone)) {
          for (let c = 0; c < row.length; c++) {
            if (c === phoneColIndex) continue;
            if (isValidPhoneCellValue(row[c])) {
              rawPhone = String(row[c]).trim();
              break;
            }
          }
        }

        const digits = rawPhone.replace(/\D/g, '');
        let normPhone = '';
        if (isValidPhoneCellValue(rawPhone) && digits.length >= 5) {
          const { normalized } = normalizePhone(rawPhone);
          normPhone = normalized || (digits.length >= 10 ? digits.slice(-10) : digits);
        }

        let rowDate = '';
        if (dateColIndex !== -1 && row[dateColIndex] !== undefined && row[dateColIndex] !== null) {
          const dateVal = row[dateColIndex];
          rowDate = formatExcelDateValue(dateVal) || String(dateVal).trim();
          if (['none', 'null', 'nan', 'n/a', 'na', '-', '--', '#n/a', 'undefined'].includes(rowDate.toLowerCase())) {
            rowDate = '';
          }
        }

        extracted.push({
          rawPhone,
          normPhone,
          date: rowDate,
        });

        if (normPhone) {
          allNormPhonesSet.add(normPhone);
        }
      }

      // CRITICAL GUARANTEE: If extracted has 0 rows, check search_index for this sheet
      // so no selected sheet that has data (e.g. OCT 26) is ever left blank!
      if (extracted.length === 0) {
        try {
          const dbRows = db.exec(
            `SELECT phone_raw, norm_phone, customer_name, raw_row_json FROM search_index 
             WHERE sheet_name = ? AND phone_raw IS NOT NULL AND phone_raw != '';`,
            [sheetName]
          );
          if (dbRows.length && dbRows[0].values.length) {
            for (const rowVal of dbRows[0].values) {
              const pRaw = String(rowVal[0] || '').trim();
              if (!pRaw) continue;
              const { normalized } = normalizePhone(pRaw);
              const pNorm = normalized || String(rowVal[1] || '').trim() || pRaw.replace(/\D/g, '').slice(-10);

              let rowDate = '';
              try {
                const parsedJson = JSON.parse(rowVal[3] as string);
                for (const k of Object.keys(parsedJson)) {
                  if (/case.*creat.*date/i.test(k) || /case.*create/i.test(k) || /creat.*date/i.test(k)) {
                    const val = parsedJson[k];
                    if (val) {
                      rowDate = formatExcelDateValue(val) || String(val).trim();
                      break;
                    }
                  }
                }
              } catch {}

              extracted.push({
                rawPhone: pRaw,
                normPhone: pNorm,
                date: rowDate,
              });

              if (pNorm) {
                allNormPhonesSet.add(pNorm);
              }
            }
          }
        } catch (dbErr) {
          console.warn('Fallback search_index extraction failed for sheet ' + sheetName, dbErr);
        }
      }

      sheetPhonesMap[sheetName] = extracted;
    }

    // Step 2: Query SQLite search_index for all extracted unique phones in batches of 400
    const allUniquePhones = Array.from(allNormPhonesSet);
    const recordsByPhone: Record<string, any[]> = {};
    const BATCH_SIZE = 400;

    for (let i = 0; i < allUniquePhones.length; i += BATCH_SIZE) {
      const batch = allUniquePhones.slice(i, i + BATCH_SIZE);
      const placeholders = batch.map(() => '?').join(',');
      const sql = `
        SELECT id, customer_name, phone_raw, norm_phone, file_name, sheet_name, row_num, folder_path, raw_row_json, source_id, file_id
        FROM search_index
        WHERE norm_phone IN (${placeholders})
          AND source_id != 'mis_sheet'
          AND file_id NOT LIKE 'gsheet_%'
          AND folder_path != 'Google Sheets MIS'
        ORDER BY norm_phone, file_name, sheet_name, row_num ASC;
      `;
      const res = db.exec(sql, batch);
      if (res.length > 0 && res[0].values.length > 0) {
        for (const row of res[0].values) {
          const pNorm = row[3] as string;
          if (!recordsByPhone[pNorm]) recordsByPhone[pNorm] = [];

          let rawObj: Record<string, any> = {};
          try {
            rawObj = JSON.parse(row[8] as string);
          } catch (e) {}

          const analysis = analyzeRowSegment(row[4] as string, row[5] as string, rawObj, row[7] as string);

          recordsByPhone[pNorm].push({
            id: row[0] as number,
            customerName: (row[1] as string) || '',
            phone: (row[2] as string) || '',
            fileName: row[4] as string,
            sheetName: row[5] as string,
            rowNum: row[6] as number,
            folderPath: (row[7] as string) || 'Root',
            analysis,
            rawRow: rawObj,
            sourceId: row[9] as string,
            fileId: row[10] as string,
          });
        }
      }
    }

    // Step 3: Format segregated results for each sheet
    const sheetsData: Record<string, ProcessedPhoneItem[]> = {};
    let totalPhonesCount = 0;
    let totalMatchesCount = 0;
    const priorityBreakdown: Record<string, number> = {};
    const perSheetStats: Record<string, { total: number; found: number; notFound: number }> = {};

    for (const sheetName of selectedSheets) {
      const phoneItems = sheetPhonesMap[sheetName] || [];
      const processedRows: ProcessedPhoneItem[] = [];
      let sheetFound = 0;

      for (const item of phoneItems) {
        const matches = recordsByPhone[item.normPhone] || [];
        // Filter out any self-matches
        const cleanMatches = matches.filter(m => {
          const fUpper = (m.fileName || '').toUpperCase();
          const pUpper = (m.folderPath || '').toUpperCase();
          const sId = String(m.sourceId || '').toLowerCase();
          const fId = String(m.fileId || '').toLowerCase();
          return (
            sId !== 'mis_sheet' &&
            !fUpper.startsWith('GSHEET_') &&
            !fId.startsWith('gsheet_') &&
            pUpper !== 'GOOGLE SHEETS MIS'
          );
        });

        const found = cleanMatches.length > 0;
        if (found) {
          sheetFound++;
          totalMatchesCount++;
        }
        totalPhonesCount++;

        if (found) {
          const leadPriority = determineLeadPriority(cleanMatches, priorityRules);
          const distinctWorkbooks = Array.from(new Set(cleanMatches.map(m => m.fileName).filter(Boolean)));
          const distinctSheets = Array.from(new Set(cleanMatches.map(m => m.analysis?.sheetLabel || m.sheetName).filter(Boolean)));
          const distinctNames = Array.from(new Set(cleanMatches.map(m => m.customerName).filter(Boolean)));

          const priorityLabel = leadPriority.priority;
          priorityBreakdown[priorityLabel] = (priorityBreakdown[priorityLabel] || 0) + 1;

          processedRows.push({
            phone: item.rawPhone,
            normPhone: item.normPhone,
            date: item.date || '—',
            priority: priorityLabel,
            priorityRank: leadPriority.rank,
            priorityCategory: leadPriority.categoryName,
            badgeBg: leadPriority.badgeBg,
            badgeText: leadPriority.badgeText,
            badgeBorder: leadPriority.badgeBorder,
            workbook: distinctWorkbooks.join(', ') || '—',
            sheetName: distinctSheets.join(', ') || '—',
            customerName: distinctNames[0] || '—',
            matchCount: cleanMatches.length,
            found: true,
          });
        } else {
          priorityBreakdown['Not Found'] = (priorityBreakdown['Not Found'] || 0) + 1;

          processedRows.push({
            phone: item.rawPhone,
            normPhone: item.normPhone,
            date: item.date || '—',
            priority: 'Not Found',
            priorityRank: 999,
            priorityCategory: 'Unmatched',
            badgeBg: 'bg-slate-100 text-slate-600 border border-slate-200',
            badgeText: 'text-slate-600',
            badgeBorder: 'border-slate-200',
            workbook: '—',
            sheetName: '—',
            customerName: '—',
            matchCount: 0,
            found: false,
          });
        }
      }

      sheetsData[sheetName] = processedRows;
      perSheetStats[sheetName] = {
        total: phoneItems.length,
        found: sheetFound,
        notFound: phoneItems.length - sheetFound,
      };
    }

    const durationMs = Date.now() - startTime;

    res.json({
      success: true,
      workbookTitle,
      selectedSheets,
      summary: {
        totalSheetsProcessed: selectedSheets.length,
        totalPhonesCount,
        totalMatchesCount,
        totalNotFoundCount: totalPhonesCount - totalMatchesCount,
        durationMs,
        priorityBreakdown,
        perSheetStats,
      },
      sheetsData,
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Failed to process MIS Search Book' });
  }
});

// 3. Export XLSX Workbook with separate sheets for each selected MIS sheet, matching source sheet names exactly
misSearchBookRouter.post('/mis-search-book/export', async (req: Request, res: Response) => {
  const {
    sheetsData,
    selectedSheets,
    workbookName,
    includePivot = true,
    pivotPerSheet = true,
    useSheetNameAsHeader = true,
    includeDataSheets = true,
    includeCombinedMatrix = true,
    matrixTitle = 'Jan- Sept Logins',
    includeCombinedPivot = false,
    pivotTitle,
    splitByMonth = false,
  } = req.body;
  if (!sheetsData || typeof sheetsData !== 'object' || Object.keys(sheetsData).length === 0) {
    return res.status(400).json({ error: 'No processed sheets data provided for export.' });
  }

  try {
    const db = await getDb();
    const wb = new ExcelJS.Workbook();
    wb.creator = 'F2 Fintech';
    wb.created = new Date();

    const sheetList: string[] = Array.isArray(selectedSheets) && selectedSheets.length > 0
      ? selectedSheets
      : Object.keys(sheetsData);

    const allExportedRows: ProcessedPhoneItem[] = [];

    // Helper to generate the exact 2D Cross-Sheet Matrix Summary matching user's Image
    const buildCombinedMatrixWorksheet = (titleText: string = 'Jan- Sept Logins') => {
      const prioCountBySheet: Record<string, Record<string, number>> = {};
      const blanksBySheet: Record<string, number> = {};
      const totalBySheet: Record<string, number> = {};
      const allPrioritiesSet = new Set<string>();

      sheetList.forEach((sName) => {
        prioCountBySheet[sName] = {};
        blanksBySheet[sName] = 0;
        totalBySheet[sName] = 0;

        const rows = sheetsData[sName] || [];
        rows.forEach((r: any) => {
          totalBySheet[sName]++;
          const prio = (r.priority || '').trim();
          if (!r.found || !prio || prio === 'Not Found' || prio === '—') {
            blanksBySheet[sName]++;
          } else {
            prioCountBySheet[sName][prio] = (prioCountBySheet[sName][prio] || 0) + 1;
            allPrioritiesSet.add(prio);
          }
        });
      });

      const priorityTotals: Record<string, number> = {};
      allPrioritiesSet.forEach((prio) => {
        let sum = 0;
        sheetList.forEach((sName) => {
          sum += prioCountBySheet[sName][prio] || 0;
        });
        priorityTotals[prio] = sum;
      });

      const sortedPriorities = Array.from(allPrioritiesSet).sort((a, b) => {
        const diff = (priorityTotals[b] || 0) - (priorityTotals[a] || 0);
        return diff !== 0 ? diff : a.localeCompare(b);
      });

      let totalBlanks = 0;
      sheetList.forEach((sName) => {
        totalBlanks += blanksBySheet[sName] || 0;
      });

      let grandTotal = 0;
      sheetList.forEach((sName) => {
        grandTotal += totalBySheet[sName] || 0;
      });

      const ws = wb.addWorksheet('Combined Summary', { views: [{ showGridLines: true }] });

      const columns = [
        { key: 'category', width: 26 },
        ...sheetList.map((_, i) => ({ key: `col_${i}`, width: 13 })),
        { key: 'total', width: 13 },
      ];
      ws.columns = columns;

      const totalColIdx = sheetList.length + 2;

      // Row 1: Merged Title Banner matching user image (Dark Navy #1B365D, Bold White text)
      ws.mergeCells(1, 1, 1, totalColIdx);
      for (let c = 1; c <= totalColIdx; c++) {
        const cell = ws.getCell(1, c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1B365D' } };
      }
      const titleCell = ws.getCell(1, 1);
      titleCell.value = titleText;
      titleCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 13, name: 'Calibri' };
      titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
      ws.getRow(1).height = 28;

      // Row 2: Headers
      const a2 = ws.getCell(2, 1);
      a2.value = 'Row Labels';
      a2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      a2.font = { bold: true, size: 10.5, name: 'Calibri' };
      a2.alignment = { vertical: 'middle', horizontal: 'left' };
      a2.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
      };

      sheetList.forEach((sName, idx) => {
        const cell = ws.getCell(2, idx + 2);
        cell.value = sName;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCFDDF5' } }; // Light blue #CFDDF5
        cell.font = { bold: true, size: 10.5, name: 'Calibri' };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } },
        };
      });

      const totalHeaderCell = ws.getCell(2, totalColIdx);
      totalHeaderCell.value = 'Total';
      totalHeaderCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } }; // Light warm yellow #FFE699
      totalHeaderCell.font = { bold: true, size: 10.5, name: 'Calibri' };
      totalHeaderCell.alignment = { vertical: 'middle', horizontal: 'center' };
      totalHeaderCell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
      };
      ws.getRow(2).height = 22;

      let currRow = 3;

      // Data rows (one for each priority)
      sortedPriorities.forEach((prio) => {
        const r = ws.getRow(currRow);
        r.height = 19;

        // Category name (Column A)
        const c1 = ws.getCell(currRow, 1);
        c1.value = prio;
        c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } }; // Light pastel yellow #FFF2CC
        c1.font = { size: 10, name: 'Calibri' };
        c1.alignment = { vertical: 'middle', horizontal: 'left' };
        c1.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } },
        };

        // Counts per sheet
        sheetList.forEach((sName, idx) => {
          const count = prioCountBySheet[sName]?.[prio];
          const cell = ws.getCell(currRow, idx + 2);
          cell.value = count && count > 0 ? count : ''; // Blank if 0
          cell.font = { size: 10, name: 'Calibri' };
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } },
          };
        });

        // Total cell (RED font on light yellow background matching user image)
        const totCell = ws.getCell(currRow, totalColIdx);
        totCell.value = priorityTotals[prio] || 0;
        totCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
        totCell.font = { bold: true, size: 10, color: { argb: 'FFC00000' }, name: 'Calibri' }; // RED font #C00000
        totCell.alignment = { vertical: 'middle', horizontal: 'right' };
        totCell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } },
        };

        currRow++;
      });

      // (blank) row
      const bRow = ws.getRow(currRow);
      bRow.height = 19;
      const b1Cell = ws.getCell(currRow, 1);
      b1Cell.value = '(blank)';
      b1Cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      b1Cell.font = { size: 10, name: 'Calibri' };
      b1Cell.alignment = { vertical: 'middle', horizontal: 'left' };
      b1Cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
      };

      sheetList.forEach((sName, idx) => {
        const count = blanksBySheet[sName];
        const cell = ws.getCell(currRow, idx + 2);
        cell.value = count && count > 0 ? count : '';
        cell.font = { size: 10, name: 'Calibri' };
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } },
        };
      });

      const bTotCell = ws.getCell(currRow, totalColIdx);
      bTotCell.value = totalBlanks;
      bTotCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      bTotCell.font = { bold: true, size: 10, color: { argb: 'FFC00000' }, name: 'Calibri' }; // RED font
      bTotCell.alignment = { vertical: 'middle', horizontal: 'right' };
      bTotCell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
      };
      currRow++;

      // Bottom Total row (Soft blue #BDD7EE, bold text)
      const tRow = ws.getRow(currRow);
      tRow.height = 24;
      const t1Cell = ws.getCell(currRow, 1);
      t1Cell.value = 'Total';
      t1Cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } }; // Light blue #BDD7EE
      t1Cell.font = { bold: true, size: 11, name: 'Calibri' };
      t1Cell.alignment = { vertical: 'middle', horizontal: 'left' };
      t1Cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
      };

      sheetList.forEach((sName, idx) => {
        const cell = ws.getCell(currRow, idx + 2);
        cell.value = totalBySheet[sName] || 0;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
        cell.font = { bold: true, size: 11, name: 'Calibri' };
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } },
        };
      });

      const grandTotCell = ws.getCell(currRow, totalColIdx);
      grandTotCell.value = grandTotal;
      grandTotCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
      grandTotCell.font = { bold: true, size: 11, name: 'Calibri' };
      grandTotCell.alignment = { vertical: 'middle', horizontal: 'right' };
      grandTotCell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
      };
    };

    // Helper to generate a styled Pivot Summary Worksheet in the ExcelJS workbook
    const buildMisPivotWorksheet = (sheetTabName: string, titleText: string, dataset: ProcessedPhoneItem[]) => {
      let safeTab = sheetTabName.replace(/[:\\/?*\[\]]/g, ' ').trim().slice(0, 31) || 'Summary';
      let pCounter = 1;
      let finalPivotName = safeTab;
      while (wb.worksheets.some(w => w.name === finalPivotName)) {
        finalPivotName = `${safeTab.slice(0, 27)}_${pCounter++}`;
      }

      const ws = wb.addWorksheet(finalPivotName, { views: [{ showGridLines: true }] });
      ws.columns = [
        { key: 'data', width: 28 },
        { key: 'count', width: 15 },
      ];

      // Row 1: Merged Title Header matching user's Image 2
      ws.mergeCells('A1:B1');
      const titleCell = ws.getCell('A1');
      titleCell.value = titleText;
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF581C87' } }; // Deep purple #581C87
      titleCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 14, name: 'Calibri' };
      titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
      ws.getRow(1).height = 32;

      // Row 2: Sub-headers 'Data' | 'Count'
      const r2 = ws.getRow(2);
      r2.height = 22;
      ws.getCell('A2').value = 'Data';
      ws.getCell('B2').value = 'Count';
      ['A2', 'B2'].forEach((ref, idx) => {
        const c = ws.getCell(ref);
        c.font = { bold: true, size: 11, name: 'Calibri' };
        c.alignment = { vertical: 'middle', horizontal: idx === 0 ? 'left' : 'right' };
        c.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } },
        };
      });

      // Calculate Priority Pivot counts and Blanks count
      const priorityMap: Record<string, number> = {};
      let blanksCount = 0;

      for (const r of dataset) {
        const prio = (r.priority || '').trim();
        if (!r.found || !prio || prio === 'Not Found' || prio === '—') {
          blanksCount++;
        } else {
          priorityMap[prio] = (priorityMap[prio] || 0) + 1;
        }
      }

      const sortedPriorities = Object.entries(priorityMap).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
      let currRowIdx = 3;

      for (const [prioName, pCount] of sortedPriorities) {
        const r = ws.getRow(currRowIdx);
        r.height = 20;
        const cA = ws.getCell('A' + currRowIdx);
        const cB = ws.getCell('B' + currRowIdx);
        cA.value = prioName;
        cB.value = pCount;
        [cA, cB].forEach((c, idx) => {
          c.font = { size: 11, name: 'Calibri' };
          c.alignment = { vertical: 'middle', horizontal: idx === 0 ? 'left' : 'right' };
          c.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } },
          };
        });
        currRowIdx++;
      }

      // Blanks row (when found nothing or unmatched)
      const bRow = ws.getRow(currRowIdx);
      bRow.height = 20;
      const bA = ws.getCell('A' + currRowIdx);
      const bB = ws.getCell('B' + currRowIdx);
      bA.value = 'Blanks';
      bB.value = blanksCount;
      [bA, bB].forEach((c, idx) => {
        c.font = { size: 11, name: 'Calibri' };
        c.alignment = { vertical: 'middle', horizontal: idx === 0 ? 'left' : 'right' };
        c.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } },
        };
      });
      currRowIdx++;

      // Grand Total row
      const gtRow = ws.getRow(currRowIdx);
      gtRow.height = 24;
      const gtA = ws.getCell('A' + currRowIdx);
      const gtB = ws.getCell('B' + currRowIdx);
      gtA.value = 'Grand Total';
      gtB.value = dataset.length;
      [gtA, gtB].forEach((c, idx) => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDE9FE' } }; // Soft lavender #EDE9FE
        c.font = { bold: true, size: 11, name: 'Calibri' };
        c.alignment = { vertical: 'middle', horizontal: idx === 0 ? 'left' : 'right' };
        c.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } },
        };
      });
    };

    // Helper to generate a styled Data Worksheet in ExcelJS
    const buildMisDataWorksheet = (sheetTabName: string, rows: ProcessedPhoneItem[]) => {
      let safeSheetName = sheetTabName.replace(/[:\\/?*\[\]]/g, '_').trim().slice(0, 31) || 'Sheet1';
      let dCounter = 1;
      let finalName = safeSheetName;
      while (wb.worksheets.some(w => w.name === finalName)) {
        finalName = `${safeSheetName.slice(0, 27)}_${dCounter++}`;
      }

      const ws = wb.addWorksheet(finalName, { views: [{ showGridLines: true }] });
      ws.columns = [
        { header: 'PHONE NUMBER', key: 'phone', width: 18 },
        { header: 'DATE', key: 'date', width: 18 },
        { header: 'PRIORITY', key: 'priority', width: 28 },
        { header: 'WORKBOOK', key: 'workbook', width: 35 },
        { header: 'SHEET NAME', key: 'sheetName', width: 30 },
      ];

      const hRow = ws.getRow(1);
      hRow.height = 24;
      hRow.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; // Dark slate
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF94A3B8' } },
          bottom: { style: 'thin', color: { argb: 'FF94A3B8' } },
          left: { style: 'thin', color: { argb: 'FF94A3B8' } },
          right: { style: 'thin', color: { argb: 'FF94A3B8' } },
        };
      });

      rows.forEach(r => {
        const row = ws.addRow({
          phone: r.phone || r.normPhone || '',
          date: r.date || '—',
          priority: r.priority || 'Not Found',
          workbook: r.workbook || '—',
          sheetName: r.sheetName || '—',
        });
        row.height = 20;
        row.eachCell((cell, colNumber) => {
          cell.font = { size: 11, name: 'Calibri' };
          cell.alignment = {
            vertical: 'middle',
            horizontal: colNumber === 1 || colNumber === 2 ? 'center' : 'left',
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          };
        });
      });
    };

    // 1. Ensure all selected sheets have rows populated
    for (const sheetName of sheetList) {
      let rows: ProcessedPhoneItem[] = sheetsData[sheetName] || [];

      // Guarantee: if client passed 0 rows for a sheet that has data in search_index, populate them!
      if (rows.length === 0) {
        try {
          const dbRows = db.exec(
            `SELECT phone_raw, norm_phone, customer_name, raw_row_json FROM search_index 
             WHERE sheet_name = ? AND phone_raw IS NOT NULL AND phone_raw != '';`,
            [sheetName]
          );
          if (dbRows.length && dbRows[0].values.length) {
            for (const rowVal of dbRows[0].values) {
              const pRaw = String(rowVal[0] || '').trim();
              if (!pRaw) continue;
              const { normalized } = normalizePhone(pRaw);
              const pNorm = normalized || String(rowVal[1] || '').trim() || pRaw.replace(/\D/g, '').slice(-10);

              let rowDate = '';
              try {
                const parsedJson = JSON.parse(rowVal[3] as string);
                for (const k of Object.keys(parsedJson)) {
                  if (/case.*creat.*date/i.test(k) || /case.*create/i.test(k) || /creat.*date/i.test(k)) {
                    const val = parsedJson[k];
                    if (val) {
                      rowDate = formatExcelDateValue(val) || String(val).trim();
                      break;
                    }
                  }
                }
              } catch {}

              const mRes = db.exec(
                `SELECT file_name, sheet_name FROM search_index 
                 WHERE (norm_phone = ? OR phone_raw = ?) 
                   AND source_id != 'mis_sheet' 
                   AND folder_path != 'GOOGLE SHEETS MIS' 
                   AND file_name NOT LIKE '%MASTER MIS%' LIMIT 5;`,
                [pNorm, pRaw]
              );
              let priority = 'Not Found';
              let workbook = '—';
              let sName = '—';
              if (mRes.length && mRes[0].values.length) {
                workbook = (mRes[0].values.map((v: any[]) => v[0]).filter(Boolean) as string[]).join(', ') || '—';
                sName = (mRes[0].values.map((v: any[]) => v[1]).filter(Boolean) as string[]).join(', ') || '—';
                priority = 'F2 Leads P1';
              }

              rows.push({
                phone: pRaw,
                normPhone: pNorm,
                date: rowDate,
                priority,
                priorityRank: priority === 'Not Found' ? 999 : 1,
                priorityCategory: priority,
                workbook,
                sheetName: sName,
                customerName: String(rowVal[2] || '—'),
                matchCount: mRes.length ? mRes[0].values.length : 0,
                found: priority !== 'Not Found',
              });
            }
          }
        } catch (dbErr) {
          console.warn('Export fallback failed for sheet ' + sheetName, dbErr);
        }
      }

      sheetsData[sheetName] = rows;
      allExportedRows.push(...rows);
    }

    // 2. Generate the 2D Cross-Sheet Matrix Summary matching user's Image as the FIRST sheet!
    if (includeCombinedMatrix !== false && sheetList.length > 0) {
      buildCombinedMatrixWorksheet(matrixTitle || 'Jan- Sept Logins');
    }

    // 3. User requirement:
    // "HOW MANY SHEETS I HAVE SELECTED MAKE ME THAT MUCH PIVOTS ONLY . EACH PIVOT FOR EACH SHEET AND FOR PIVOT HEADER NAME USE THE NAME OF THE SHEET U ARE MAKING PIVOT OF"
    const isPivotPerSheet = pivotPerSheet !== false || splitByMonth === true;

    // Add Dedicated Pivot Sheet for EACH SELECTED SHEET first (1 Pivot per sheet, Header Name = Sheet Name!)
    if (includePivot && isPivotPerSheet) {
      for (const sheetName of sheetList) {
        const rows = sheetsData[sheetName] || [];
        const pivotTabName = includeDataSheets !== false ? `${sheetName} Pivot` : sheetName;
        const pivotHeaderTitle = sheetName; // EXACTLY THE NAME OF THE SHEET!
        buildMisPivotWorksheet(pivotTabName, pivotHeaderTitle, rows);
      }
    }

    // Add Raw Data Sheet for each selected MIS sheet (if includeDataSheets enabled)
    if (includeDataSheets !== false) {
      for (const sheetName of sheetList) {
        const rows = sheetsData[sheetName] || [];
        buildMisDataWorksheet(sheetName, rows);
      }
    }

    // If pivotPerSheet is false and no matrix, generate a single summary pivot sheet
    if (includePivot && !isPivotPerSheet && !includeCombinedMatrix && allExportedRows.length > 0) {
      const singleTitle = pivotTitle?.trim() || (sheetList.length === 1 ? sheetList[0] : 'Pivot Summary');
      const singleTab = singleTitle;
      buildMisPivotWorksheet(singleTab, singleTitle, allExportedRows);
    }

    // Optional Combined Pivot of all sheets if requested
    if (includePivot && includeCombinedPivot && sheetList.length > 1 && allExportedRows.length > 0) {
      buildMisPivotWorksheet('Combined Pivot', 'All Sheets Combined', allExportedRows);
    }

    const arrayBuffer = await wb.xlsx.writeBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const safeTitle = (workbookName || 'MIS_Search_Book').replace(/[^a-zA-Z0-9_-]/g, '_');
    const dateStr = new Date().toISOString().slice(0, 10);
    const fileName = `${safeTitle}_Global_Search_${dateStr}.xlsx`;

    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Failed to generate Excel workbook' });
  }
});
