import * as fs from 'fs';
import * as path from 'path';
import * as XLSXNamespace from 'xlsx';
const XLSX: any = (XLSXNamespace as any).default || XLSXNamespace;
import { getDb } from './db.js';
import {
  normalizePhone,
  cleanAndExtractPhoneNumbers,
  analyzeRowSegment,
  extractSafeCellValue,
  rowHasAnyEntry,
  cleanWorkbookSheets,
  isValidPhoneCellValue,
  rowHasPhoneNumber,
} from './normalizer.js';
import { isGoogleConfigured, listTabs, readSheet } from './google.js';
import { determineLeadPriority, getPriorityRules, evaluateSingleRecordPriority, formatTeamXSheetName } from './priority.js';

export interface MisDetectedSheet {
  sheetName: string;
  totalRows: number;
  detectedPhoneColIndex: number;
  detectedPhoneColName: string;
  headers: string[];
  previewRows: Record<string, any>[];
}

export interface MisEnrichedRow {
  sourceSheet: string;
  sourceRowNum: number;
  inputPhone: string;
  normPhone: string;
  found: boolean;
  priority: string;
  priorityCategory: string;
  priorityRank: number;
  badgeBg?: string;
  badgeText?: string;
  badgeBorder?: string;
  foundWorkbook: string;
  foundSheet: string;
  foundCustomerName: string;
  foundLender: string;
  foundLoanAmount: string;
  foundRowNum: string;
  matchCount: number;
  originalData: Record<string, any>;
}

export class MisService {
  // Extract Google Spreadsheet ID from URL or return raw ID
  static extractSpreadsheetId(urlOrId: string): string | null {
    if (!urlOrId || typeof urlOrId !== 'string') return null;
    const trimmed = urlOrId.trim();

    // Pattern: https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit...
    const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (match) return match[1];

    // Pattern: https://drive.google.com/open?id=1BxiMVs0... or id=1BxiMVs0...
    const idParam = trimmed.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (idParam) return idParam[1];

    // Direct ID alphanumeric 25-60 chars
    if (/^[a-zA-Z0-9_-]{25,60}$/.test(trimmed)) {
      return trimmed;
    }

    return null;
  }

  /**
   * Fallback for private Google Sheets: read EVERY tab through the Apps Script connection
   * (listTabs + readSheet) and build one workbook with all tabs. Returns null if Apps Script
   * is not set up or fails, so the caller can use its old cache fallback.
   */
  static async fetchWorkbookViaAppsScript(spreadsheetId: string): Promise<Buffer | null> {
    if (!isGoogleConfigured()) return null;
    try {
      const tabs = await listTabs(spreadsheetId);
      if (!tabs || tabs.length === 0) return null;
      const wb = XLSX.utils.book_new();
      const used = new Set<string>();
      for (const t of tabs) {
        try {
          const { headers, rows } = await readSheet(spreadsheetId, t.name);
          const aoa = headers.length > 0 ? [headers, ...rows] : rows;
          // Excel tab names: max 31 characters, no  : \\ / ? * [ ]
          let safe = String(t.name || 'Sheet').replace(/[:\\/?*[\]]/g, ' ').trim().slice(0, 31) || 'Sheet';
          let n = 2;
          while (used.has(safe.toLowerCase())) safe = `${safe.slice(0, 28)} ${n++}`;
          used.add(safe.toLowerCase());
          XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), safe);
        } catch (tabErr: any) {
          console.warn(`[mis] Could not read tab "${t.name}" through Apps Script:`, tabErr?.message || tabErr);
        }
      }
      if (wb.SheetNames.length === 0) return null;
      return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
    } catch (err: any) {
      console.warn('[mis] Apps Script all-tabs fetch failed:', err?.message || err);
      return null;
    }
  }

  // Fetch Google Spreadsheet as XLSX binary buffer with metadata indicating live vs cached source
  static async fetchGoogleSheetWithMeta(
    urlOrId: string,
    maxRetries: number = 3
  ): Promise<{ buffer: Buffer; isLive: boolean; source: 'google_live' | 'local_cache' }> {
    const spreadsheetId = this.extractSpreadsheetId(urlOrId);
    if (!spreadsheetId) {
      throw new Error('Invalid Google Sheet URL. Please provide a valid link in the format: https://docs.google.com/spreadsheets/d/<ID>/edit');
    }

    const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
    const cacheDir = path.resolve('./data/cache');
    const cacheFile = path.join(cacheDir, `gsheet_${spreadsheetId}.xlsx`);
    const isMasterMisId =
      spreadsheetId === '1PDss3DpFJ7OmVdW3bBpAc9r6xlSaCrg6BcYYrc-B744' ||
      spreadsheetId === '15cVL3BK6yXPTwGGG0s_gp-jW-KxS-70cXn__6C8VCBs';

    let lastError: any = null;

    // When the public download is blocked, read all tabs through Apps Script before using an old cache
    // (an old cache can hold only one tab, e.g. OCT 26).
    const viaAppsScript = async (): Promise<{ buffer: Buffer; isLive: boolean; source: 'google_live' | 'local_cache' } | null> => {
      const b = await this.fetchWorkbookViaAppsScript(spreadsheetId);
      if (!b) return null;
      try {
        if (!fs.existsSync(cacheDir)) fs.mkdirSync(cacheDir, { recursive: true });
        fs.writeFileSync(cacheFile, b);
        if (isMasterMisId) fs.writeFileSync(path.join(cacheDir, 'MASTER MIS.xlsx'), b);
      } catch {}
      return { buffer: b, isLive: true, source: 'google_live' };
    };

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      if (attempt > 0) {
        // Exponential backoff with jitter: 1s, 2s, 4s + 0-500ms
        const delay = Math.min(1000 * Math.pow(2, attempt - 1) + Math.random() * 500, 10000);
        await new Promise((res) => setTimeout(res, delay));
      }

      try {
        const res = await fetch(exportUrl, {
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/octet-stream',
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0',
          },
          redirect: 'follow',
          signal: AbortSignal.timeout(120000),
        });

        // If redirected to Google login or unauthorized (401/403)
        if (res.url.includes('accounts.google.com') || res.status === 401 || res.status === 403) {
          const viaScript = await viaAppsScript();
          if (viaScript) return viaScript;
          if (fs.existsSync(cacheFile)) {
            return { buffer: fs.readFileSync(cacheFile), isLive: false, source: 'local_cache' };
          }
          const masterMisFallback = path.join(cacheDir, 'MASTER MIS.xlsx');
          if (fs.existsSync(masterMisFallback)) {
            const fallbackBuf = fs.readFileSync(masterMisFallback);
            try { fs.writeFileSync(cacheFile, fallbackBuf); } catch {}
            return { buffer: fallbackBuf, isLive: false, source: 'local_cache' };
          }
          throw new Error(
            'Google Sheet is Private (HTTP 401). To sync automatically, open the sheet > click "Share" > change General Access to "Anyone with the link can view". Alternatively, use "Upload Excel File" to import directly.'
          );
        }

        // Retry on 429 (rate-limit) or 5xx server errors
        if (res.status === 429 || res.status >= 500) {
          const retryAfter = res.headers.get('retry-after');
          if (retryAfter && !isNaN(Number(retryAfter))) {
            await new Promise((r) => setTimeout(r, Math.min(Number(retryAfter) * 1000, 15000)));
          }
          lastError = new Error(`Google Sheets export returned HTTP ${res.status}`);
          continue;
        }

        if (!res.ok) {
          const viaScript = await viaAppsScript();
          if (viaScript) return viaScript;
          if (fs.existsSync(cacheFile)) {
            return { buffer: fs.readFileSync(cacheFile), isLive: false, source: 'local_cache' };
          }
          if (res.status === 404) {
            throw new Error('Google Sheet not found. Please verify the URL.');
          }
          throw new Error(`Google Sheets export returned HTTP ${res.status}. Ensure the sheet is public or upload the file directly.`);
        }

        const contentType = res.headers.get('content-type') || '';
        if (contentType.includes('text/html') && !contentType.includes('spreadsheetml')) {
          const viaScript = await viaAppsScript();
          if (viaScript) return viaScript;
          if (fs.existsSync(cacheFile)) {
            return { buffer: fs.readFileSync(cacheFile), isLive: false, source: 'local_cache' };
          }
          const masterMisFallback = path.join(cacheDir, 'MASTER MIS.xlsx');
          if (fs.existsSync(masterMisFallback)) {
            const fallbackBuf = fs.readFileSync(masterMisFallback);
            try { fs.writeFileSync(cacheFile, fallbackBuf); } catch {}
            return { buffer: fallbackBuf, isLive: false, source: 'local_cache' };
          }
          throw new Error(
            'Google Sheet returned a login page instead of spreadsheet data. Please ensure General Access is set to "Anyone with the link can view", or use "Upload Excel File" directly.'
          );
        }

        const arrayBuf = await res.arrayBuffer();
        const buffer = Buffer.from(arrayBuf);
        if (buffer.length < 50) {
          throw new Error('Downloaded Google Sheet was empty.');
        }

        let finalBuffer = buffer;
        try {
          const downloadedWb = XLSX.read(buffer, { type: 'buffer' });
          const octWs = downloadedWb.Sheets['OCT 26'];
          let octHasData = false;
          if (octWs) {
            const octRawRows = (XLSX.utils.sheet_to_json(octWs, { header: 1, raw: true, defval: '' }) as any[][]) || [];
            const headers = (octRawRows[0] || []).map((h: any) => String(h || ''));
            octHasData = octRawRows.slice(1).some((r: any) => rowHasAnyEntry(r, headers));
          }

          // If downloaded OCT 26 tab has only empty template dashes, merge genuine OCT 26 tab from cached copy
          if (!octHasData) {
            const fallbackPaths = [
              path.join(cacheDir, 'gsheet_15cVL3BK6yXPTwGGG0s_gp-jW-KxS-70cXn__6C8VCBs.xlsx'),
              cacheFile,
              path.join(cacheDir, 'MASTER MIS.xlsx')
            ];
            for (const fbPath of fallbackPaths) {
              if (fs.existsSync(fbPath)) {
                try {
                  const fbWb = XLSX.readFile(fbPath);
                  const fbOct = fbWb.Sheets['OCT 26'];
                  if (fbOct) {
                    const fbRawRows = (XLSX.utils.sheet_to_json(fbOct, { header: 1, raw: true, defval: '' }) as any[][]) || [];
                    const fbHeaders = (fbRawRows[0] || []).map((h: any) => String(h || ''));
                    const hasRealData = fbRawRows.slice(1).some((r: any) => rowHasAnyEntry(r, fbHeaders));
                    if (hasRealData) {
                      // Filter fbOct to contain only clean rows with entries
                      const cleanFbRows = [fbRawRows[0], ...fbRawRows.slice(1).filter((r: any) => rowHasAnyEntry(r, fbHeaders))];
                      downloadedWb.Sheets['OCT 26'] = XLSX.utils.aoa_to_sheet(cleanFbRows);
                      finalBuffer = XLSX.write(downloadedWb, { type: 'buffer', bookType: 'xlsx' });
                      break;
                    }
                  }
                } catch {}
              }
            }
          }
        } catch (mergeErr) {
          console.warn('OCT 26 check/merge error:', mergeErr);
        }

        // Sanitize all sheets in downloaded workbook (strip blank rows from October 2025 and all tabs)
        try {
          const downloadedWb = XLSX.read(finalBuffer, { type: 'buffer' });
          cleanWorkbookSheets(downloadedWb);
          finalBuffer = XLSX.write(downloadedWb, { type: 'buffer', bookType: 'xlsx' });
        } catch (cleanErr) {
          console.warn('Workbook cleaning warning:', cleanErr);
        }

        try {
          if (!fs.existsSync(cacheDir)) fs.mkdirSync(cacheDir, { recursive: true });
          fs.writeFileSync(cacheFile, finalBuffer);
          if (isMasterMisId) {
            fs.writeFileSync(path.join(cacheDir, 'MASTER MIS.xlsx'), finalBuffer);
          }
        } catch {}

        return { buffer: finalBuffer, isLive: true, source: 'google_live' };
      } catch (err: any) {
        lastError = err;
        // If abort/timeout or network error, continue retry loop
        if (attempt < maxRetries) continue;
      }
    }

    const lastTry = await viaAppsScript();
    if (lastTry) return lastTry;
    if (fs.existsSync(cacheFile)) {
      return { buffer: fs.readFileSync(cacheFile), isLive: false, source: 'local_cache' };
    }
    throw new Error(`Failed to connect to Google Sheets after ${maxRetries + 1} attempts: ${lastError?.message || 'Connection timed out'}`);
  }

  // Fetch Google Spreadsheet as XLSX binary buffer with exponential backoff retry and jitter
  static async fetchGoogleSheetAsXlsx(urlOrId: string, maxRetries: number = 3): Promise<Buffer> {
    const res = await this.fetchGoogleSheetWithMeta(urlOrId, maxRetries);
    return res.buffer;
  }

  // Smart case-and-space insensitive phone column detector
  static detectPhoneColumn(headers: string[], sampleRows: any[][] = []): { index: number; name: string } {
    if (!headers || headers.length === 0) {
      return { index: 0, name: 'Column 1' };
    }

    const normalizedHeaders = headers.map(h => String(h || '').trim().toLowerCase().replace(/[\s_\-.]+/g, ''));

    // Step 1: Scan sample rows for columns that actually contain 10-digit Indian mobile numbers
    if (sampleRows && sampleRows.length > 0) {
      const colScores: number[] = new Array(headers.length).fill(0);
      const testRows = sampleRows.slice(0, 30);

      testRows.forEach(row => {
        if (!row) return;
        headers.forEach((_, colIdx) => {
          const cell = row[colIdx];
          const val = typeof cell === 'object' && cell !== null && !(cell instanceof Date) ? (cell.v ?? cell.w ?? '') : cell;
          if (val !== undefined && val !== null) {
            const rawStr = String(val).trim();
            // Handle scientific notation e.g. 9.86837E+09
            let digits = rawStr.replace(/\D/g, '');
            if (/^\d+(?:\.\d+)?[eE]\+?\d+$/.test(rawStr)) {
              const n = Number(rawStr);
              if (!isNaN(n)) digits = Math.floor(n).toString().replace(/\D/g, '');
            }
            // Check 10 digits starting with [5-9], or 11/12 digits with prefix 0 or 91
            if (digits.length === 10 && /^[5-9]\d{9}$/.test(digits)) {
              colScores[colIdx] += 3;
            } else if ((digits.length === 11 || digits.length === 12) && /^[5-9]\d{9}$/.test(digits.slice(-10))) {
              colScores[colIdx] += 2;
            }
          }
        });
      });

      // Bonus points for headers containing phone keywords
      headers.forEach((_, colIdx) => {
        const nh = normalizedHeaders[colIdx];
        if (/phone|mobile|contact|calling|cell|mob|custphone|applicantmobile/i.test(nh)) {
          colScores[colIdx] += 5;
        }
        // Demote non-phone number columns
        if (/serial|srno|sno|account|accno|loanno|appno|token|refno|pincode|zip|amount/i.test(nh)) {
          colScores[colIdx] -= 8;
        }
      });

      let maxCol = -1;
      let maxScore = 0;
      colScores.forEach((score, idx) => {
        if (score > maxScore) {
          maxScore = score;
          maxCol = idx;
        }
      });

      if (maxCol !== -1 && maxScore > 2) {
        return { index: maxCol, name: headers[maxCol] };
      }
    }

    // Step 2: Strict keyword match
    const priorityKeywords = [
      'phonenumber',
      'mobilenumber',
      'contactnumber',
      'callingnumber',
      'cellnumber',
      'number1',
      'phone1',
      'mobile1',
      'mob1',
      'contact1',
      'primaryphone',
      'phone',
      'mobile',
      'contact',
      'phoneno',
      'mobileno',
      'contactno',
      'mobno',
      'cellno',
      'custphone',
      'customerphone',
      'applicantphone',
      'applicantmobile',
      'clientphone',
    ];

    for (const kw of priorityKeywords) {
      const idx = normalizedHeaders.findIndex(nh => nh === kw);
      if (idx !== -1) {
        return { index: idx, name: headers[idx] };
      }
    }

    // Step 3: Partial regex pattern check (skipping serial/account/amount)
    for (let i = 0; i < normalizedHeaders.length; i++) {
      const nh = normalizedHeaders[i];
      if (/serial|srno|sno|account|accno|loanno|appno|token|refno|pincode|zip|amount/i.test(nh)) {
        continue;
      }
      if (/phone|mobile|contact|cell|calling/i.test(nh)) {
        return { index: i, name: headers[i] };
      }
    }

    // Default to first column if nothing found
    return { index: 0, name: headers[0] || 'Column 1' };
  }

  // Detect the true header row index (accounts for title banners, blank rows)
  static detectHeaderRowIndex(rawJson: any[][]): number {
    if (!rawJson || rawJson.length === 0) return 0;

    let bestIndex = 0;
    let bestScore = -1;

    for (let r = 0; r < Math.min(rawJson.length, 10); r++) {
      const row = rawJson[r];
      if (!row || !Array.isArray(row)) continue;

      let validTextCells = 0;
      let keywordHits = 0;

      row.forEach(cell => {
        if (cell === null || cell === undefined) return;
        const s = String(cell).trim();
        if (s.length > 0) {
          validTextCells++;
          if (/name|phone|mob|contact|pan|aadhar|date|address|city|state|status|amount|email/i.test(s)) {
            keywordHits++;
          }
        }
      });

      // Score row: multiple non-empty columns with common header keywords
      const score = validTextCells + (keywordHits * 3);
      if (score > bestScore && validTextCells >= 2) {
        bestScore = score;
        bestIndex = r;
      }
    }

    return bestIndex;
  }

  // Parse workbook buffer and extract sheet metadata & detected phone columns
  static inspectWorkbook(buffer: Buffer): {
    sheets: MisDetectedSheet[];
    workbookName?: string;
  } {
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
    const sheets: MisDetectedSheet[] = [];

    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) continue;

      const rawJson = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' }) as any[][];
      if (!rawJson || rawJson.length === 0) continue;

      const headerRowIndex = this.detectHeaderRowIndex(rawJson);
      const rawHeaders = (rawJson[headerRowIndex] || []).map((h, i) => String(h || `Column_${i + 1}`).trim());
      const dataRows = rawJson.slice(headerRowIndex + 1);

      // Detect phone column
      const detected = this.detectPhoneColumn(rawHeaders, dataRows.slice(0, 30));

      // Count all rows that have genuine business entries across all columns
      const validDataRows = dataRows.filter(row => {
        return rowHasAnyEntry(row, rawHeaders);
      });

      const previewRows = validDataRows.slice(0, 5).map((row, rIdx) => {
        const rowObj: Record<string, any> = { _rowNum: headerRowIndex + 2 + rIdx };
        rawHeaders.forEach((h, cIdx) => {
          rowObj[h] = row[cIdx] !== undefined ? extractSafeCellValue(row[cIdx], h) : '';
        });
        return rowObj;
      });

      sheets.push({
        sheetName,
        totalRows: validDataRows.length,
        detectedPhoneColIndex: detected.index,
        detectedPhoneColName: detected.name,
        headers: rawHeaders,
        previewRows,
      });
    }

    return { sheets };
  }

  // Analyze all extracted MIS rows against the fintech database & priority rules
  static async analyzeRows(
    rowsToAnalyze: {
      sourceSheet: string;
      sourceRowNum: number;
      rawPhone: string;
      originalData: Record<string, any>;
    }[]
  ): Promise<{
    results: MisEnrichedRow[];
    summary: {
      totalRows: number;
      totalSheets: number;
      totalFound: number;
      totalNotFound: number;
      totalShortOrInvalid?: number;
      priorityBreakdown: Record<string, number>;
      durationMs: number;
    };
  }> {
    const startTime = Date.now();
    const db = await getDb();
    const priorityRules = getPriorityRules();

    // Strictly filter out rows where the phone number is missing, empty, blank, or placeholder
    const validRowsToAnalyze = rowsToAnalyze.filter(item => {
      if (!item) return false;
      const raw = String(item.rawPhone || '').trim();
      if (!raw) return false;
      const lower = raw.toLowerCase();
      if (['none', 'null', 'nan', 'n/a', 'na', 'nil', '-', '--', '0', '#n/a', 'undefined', 'not available'].includes(lower)) {
        return false;
      }
      return true;
    });

    // Map each item and perform thorough phone normalization:
    // 1. Strips +91, 91, 091, 0091, 0
    // 2. Strips spaces, dashes, dots, brackets, text labels
    // 3. Handles scientific notation (e.g. 9.87654E+09)
    // 4. Extracts all 10-digit numbers from multi-number cells (e.g. 9985132111 / 9876543210)
    // 5. Excludes entries with length < 10
    const normalizedItems = validRowsToAnalyze.map(item => {
      const raw = String(item.rawPhone || '').trim();
      const phoneInfo = cleanAndExtractPhoneNumbers(raw);
      return {
        ...item,
        phoneInfo,
        normPhone: phoneInfo.primary || '',
        allPhones: phoneInfo.phones,
        isShort: phoneInfo.isShort,
        isInvalid: phoneInfo.isInvalid,
      };
    });

    // Extract all unique candidate phones across all rows (including multi-number cells)
    const uniquePhonesSet = new Set<string>();
    normalizedItems.forEach(item => {
      item.allPhones.forEach(p => uniquePhonesSet.add(p));
    });
    const uniquePhones = Array.from(uniquePhonesSet);
    const recordsByPhone: Record<string, any[]> = {};

    // Query in batches of 400
    const BATCH_SIZE = 400;
    for (let i = 0; i < uniquePhones.length; i += BATCH_SIZE) {
      const batch = uniquePhones.slice(i, i + BATCH_SIZE);
      const placeholders = batch.map(() => '?').join(',');
      const sql = `
        SELECT id, customer_name, phone_raw, norm_phone, file_name, sheet_name, row_num, folder_path, raw_row_json
        FROM search_index
        WHERE norm_phone IN (${placeholders})
          AND source_id != 'mis_sheet'
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
          });
        }
      }
    }

    let totalFound = 0;
    let totalShort = 0;
    const priorityBreakdown: Record<string, number> = {};
    const sheetSet = new Set<string>();

    const results: MisEnrichedRow[] = normalizedItems.map(item => {
      sheetSet.add(item.sourceSheet);

      if (item.isShort) {
        totalShort++;
      }

      // Check if ANY of the extracted valid phone numbers match in the database
      let matched: any[] = [];
      let matchedPhone = item.normPhone;

      for (const p of item.allPhones) {
        if (recordsByPhone[p] && recordsByPhone[p].length > 0) {
          const cleanMatches = recordsByPhone[p].filter(m => {
            const fUpper = (m.fileName || '').toUpperCase();
            const pUpper = (m.folderPath || '').toUpperCase();
            return !fUpper.startsWith('GSHEET_') && pUpper !== 'GOOGLE SHEETS MIS';
          });
          if (cleanMatches.length > 0) {
            matched = cleanMatches;
            matchedPhone = p;
            break;
          }
        }
      }

      const found = matched.length > 0;
      if (found) totalFound++;

      // Determine highest lead priority across all matches using active priority rules
      const leadPriority = determineLeadPriority(matched, priorityRules);

      // Collect matched details
      const distinctWorkbooks = Array.from(new Set(matched.map(m => m.fileName)));
      const distinctSheets = Array.from(
        new Set(
          matched.map(m =>
            formatTeamXSheetName(m.sheetName, m.fileName, m.folderPath, m.rawRow)
          )
        )
      );
      const customerNames = Array.from(new Set(matched.map(m => m.customerName).filter(Boolean)));
      
      // Extract Lender / Loan Amount if present in raw row
      const lenders = Array.from(new Set(matched.map(m => m.rawRow?.LENDER || m.rawRow?.lender || m.rawRow?.Bank || m.rawRow?.['BANK NAME'] || m.rawRow?.['Lender Name'] || '').filter(Boolean)));
      const amounts = Array.from(new Set(matched.map(m => m.rawRow?.ELIGIBILTY || m.rawRow?.Amount || m.rawRow?.['Loan Amt'] || m.rawRow?.['LOAN AMOUNT'] || m.rawRow?.['NET LOAN AMOUNT'] || m.rawRow?.eligible_amount || '').filter(Boolean)));
      const rowNums = matched.map(m => `Row ${m.rowNum}`).slice(0, 5).join(', ');

      // Normalize any date fields inside originalData to prevent 1-day-before shift
      const cleanedOriginalData: Record<string, any> = {};
      if (item.originalData) {
        for (const [k, v] of Object.entries(item.originalData)) {
          cleanedOriginalData[k] = extractSafeCellValue(v, k);
        }
      }

      let priorityLabel = 'Unmatched (Not Found)';
      let priorityText = 'Not Found';
      if (item.isShort) {
        priorityLabel = 'Invalid (< 10 Digits)';
        priorityText = 'Invalid (< 10 Digits)';
      } else if (item.isInvalid) {
        priorityLabel = 'Invalid Phone';
        priorityText = 'Invalid Phone';
      } else if (found) {
        priorityLabel = leadPriority.priority;
        priorityText = leadPriority.priority;
      }

      priorityBreakdown[priorityLabel] = (priorityBreakdown[priorityLabel] || 0) + 1;

      return {
        sourceSheet: item.sourceSheet,
        sourceRowNum: item.sourceRowNum,
        inputPhone: item.rawPhone,
        normPhone: matchedPhone || item.normPhone,
        found,
        priority: priorityText,
        priorityCategory: found ? leadPriority.categoryName : '—',
        priorityRank: found ? leadPriority.rank : 999,
        badgeBg: found ? leadPriority.badgeBg : (item.isShort ? 'bg-amber-50' : 'bg-slate-100'),
        badgeText: found ? leadPriority.badgeText : (item.isShort ? 'text-amber-700' : 'text-slate-500'),
        badgeBorder: found ? leadPriority.badgeBorder : (item.isShort ? 'border-amber-200' : 'border-slate-200'),
        foundWorkbook: found ? distinctWorkbooks.join(', ') : '—',
        foundSheet: found ? distinctSheets.join(', ') : '—',
        foundCustomerName: customerNames[0] || '—',
        foundLender: lenders[0] || '—',
        foundLoanAmount: amounts[0] || '—',
        foundRowNum: found ? (matched.length > 5 ? `${rowNums} (+${matched.length - 5} more)` : rowNums) : '—',
        matchCount: matched.length,
        firstRecordId: found ? matched[0]?.id : undefined,
        firstFileName: found ? matched[0]?.fileName : undefined,
        firstSheetName: found ? matched[0]?.sheetName : undefined,
        originalData: cleanedOriginalData,
      };
    });

    return {
      results,
      summary: {
        totalRows: normalizedItems.length,
        totalSheets: sheetSet.size,
        totalFound,
        totalNotFound: normalizedItems.length - totalFound,
        totalShortOrInvalid: totalShort,
        priorityBreakdown,
        durationMs: Date.now() - startTime,
      },
    };
  }

  // Generate enriched Excel workbook containing all entries with Priority, Found Workbook, and Found Sheet
  static generateEnrichedWorkbook(
    results: MisEnrichedRow[],
    workbookTitle: string = 'MIS_Enriched_Analysis',
    selectedSheets?: string[]
  ): Buffer {
    const workbook = XLSX.utils.book_new();

    // Filter results by selectedSheets if provided
    const filteredResults = selectedSheets && selectedSheets.length > 0
      ? results.filter(r => selectedSheets.includes(r.sourceSheet))
      : results;

    // 1. Build Master Combined Sheet with user's exact requested column order:
    // "phone number, sheet name (of the file in which the phone number is present), priority (from the database we found), sheet name database (from the database we found), workbook name (from the database we found)"
    const masterRows = filteredResults.map(r => {
      const out: Record<string, any> = {
        'Phone Number': r.inputPhone,
        'Sheet Name': r.sourceSheet,
        'Priority': r.priority,
        'Sheet Name Database': r.foundSheet,
        'Workbook Name': r.foundWorkbook,
        'Match Status': r.found ? 'MATCHED' : 'NOT FOUND',
        'Customer Name (Database)': r.foundCustomerName,
        'Lender (Database)': r.foundLender,
        'Loan Amount (Database)': r.foundLoanAmount,
        'Row Number (Database)': r.foundRowNum,
        'Source Row': r.sourceRowNum,
      };

      // Append all original row columns from the uploaded file
      if (r.originalData) {
        Object.entries(r.originalData).forEach(([k, v]) => {
          if (!k.startsWith('_')) {
            out[k] = v;
          }
        });
      }

      return out;
    });

    const masterSheet = XLSX.utils.json_to_sheet(masterRows);
    XLSX.utils.book_append_sheet(workbook, masterSheet, 'Combined Sheet');

    // 2. Matched Sheet (Only Matched in DB)
    const matchedRows = filteredResults.filter(r => r.found).map(r => {
      const out: Record<string, any> = {
        'Phone Number': r.inputPhone,
        'Sheet Name': r.sourceSheet,
        'Priority': r.priority,
        'Workbook Name': r.foundWorkbook,
        'Sheet Name Database': r.foundSheet,
        'Customer Name (Database)': r.foundCustomerName,
        'Lender (Database)': r.foundLender,
        'Loan Amount (Database)': r.foundLoanAmount,
        'Row Number (Database)': r.foundRowNum,
        'Source Row': r.sourceRowNum,
      };
      if (r.originalData) {
        Object.entries(r.originalData).forEach(([k, v]) => {
          if (!k.startsWith('_')) out[k] = v;
        });
      }
      return out;
    });
    if (matchedRows.length > 0) {
      const matchedSheet = XLSX.utils.json_to_sheet(matchedRows);
      XLSX.utils.book_append_sheet(workbook, matchedSheet, 'Matched');
    }

    // 3. Unmatched Sheet (Valid 10-digit leads not found in DB)
    const unmatchedRows = filteredResults.filter(r => !r.found && !r.priority.includes('< 10') && !r.priority.includes('Invalid')).map(r => {
      const out: Record<string, any> = {
        'Phone Number': r.inputPhone,
        'Sheet Name': r.sourceSheet,
        'Status': 'Unmatched (Not Found in Database)',
        'Source Row': r.sourceRowNum,
      };
      if (r.originalData) {
        Object.entries(r.originalData).forEach(([k, v]) => {
          if (!k.startsWith('_')) out[k] = v;
        });
      }
      return out;
    });
    if (unmatchedRows.length > 0) {
      const unmatchedSheet = XLSX.utils.json_to_sheet(unmatchedRows);
      XLSX.utils.book_append_sheet(workbook, unmatchedSheet, 'Unmatched');
    }

    // 4. Ignored Sheet (<10 Digits or Invalid Format)
    const ignoredRows = filteredResults.filter(r => r.priority.includes('< 10') || r.priority.includes('Invalid')).map(r => {
      const out: Record<string, any> = {
        'Phone Number': r.inputPhone,
        'Sheet Name': r.sourceSheet,
        'Reason': r.priority,
        'Source Row': r.sourceRowNum,
      };
      if (r.originalData) {
        Object.entries(r.originalData).forEach(([k, v]) => {
          if (!k.startsWith('_')) out[k] = v;
        });
      }
      return out;
    });
    if (ignoredRows.length > 0) {
      const ignoredSheet = XLSX.utils.json_to_sheet(ignoredRows);
      XLSX.utils.book_append_sheet(workbook, ignoredSheet, 'Ignored');
    }

    // 5. Also include each individual selected sheet
    const sheetsMap: Record<string, MisEnrichedRow[]> = {};
    filteredResults.forEach(r => {
      if (!sheetsMap[r.sourceSheet]) sheetsMap[r.sourceSheet] = [];
      sheetsMap[r.sourceSheet].push(r);
    });

    // For every sheet that was selected, create its individual tab
    Object.entries(sheetsMap).forEach(([sheetName, sheetRows]) => {
      const safeSheetName = sheetName.replace(/[:\\/?*[\]]/g, '_').slice(0, 31);
      const sRows = sheetRows.map(r => {
        const out: Record<string, any> = {
          'Phone Number': r.inputPhone,
          'Priority': r.priority,
          'Workbook Name': r.foundWorkbook,
          'Sheet Name Database': r.foundSheet,
          'Match Status': r.found ? 'MATCHED' : 'NOT FOUND',
          'Customer Name (Database)': r.foundCustomerName,
          'Lender (Database)': r.foundLender,
          'Loan Amount (Database)': r.foundLoanAmount,
        };
        if (r.originalData) {
          Object.entries(r.originalData).forEach(([k, v]) => {
            if (!k.startsWith('_')) {
              out[k] = v;
            }
          });
        }
        return out;
      });

      const ws = XLSX.utils.json_to_sheet(sRows);
      XLSX.utils.book_append_sheet(workbook, ws, safeSheetName);
    });

    return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
  }
}
