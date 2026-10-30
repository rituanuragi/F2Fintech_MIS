/**
 * Google Apps Script Web App Client for F2 Fintech
 * Accesses Google Drive & Google Sheets securely through Apps Script Web App
 * with zero browser exposure and robust server-side retries.
 */

export const DEFAULT_DRIVE_ROOT_FOLDER_ID =
  process.env.DRIVE_ROOT_FOLDER_ID || '1xFVOGEiQNshKBpA8tQoedy9FfPTP4rjn';

export const DEFAULT_MIS_SPREADSHEET_ID =
  process.env.MIS_SPREADSHEET_ID || '15cVL3BK6yXPTwGGG0s_gp-jW-KxS-70cXn__6C8VCBs';

export interface DriveTreeItem {
  id: string;
  name: string;
  mime: string;
  modified?: string;
  size?: number;
  parentId?: string | null;
  path?: string;
}

export interface SheetTabInfo {
  sheetId: number | string;
  name: string;
  rows: number;
  cols: number;
}

export interface SheetDataResult {
  headers: string[];
  rows: any[][];
}

export function isGoogleConfigured(): boolean {
  const url = (process.env.APPS_SCRIPT_URL || '').trim();
  const token = (process.env.APPS_SCRIPT_TOKEN || '').trim();
  return Boolean(url && token);
}

export function getAppsScriptStatus() {
  const url = (process.env.APPS_SCRIPT_URL || '').trim();
  const token = (process.env.APPS_SCRIPT_TOKEN || '').trim();
  return {
    hasUrl: Boolean(url),
    hasToken: Boolean(token),
    driveRootFolderId: DEFAULT_DRIVE_ROOT_FOLDER_ID,
    misSpreadsheetId: DEFAULT_MIS_SPREADSHEET_ID,
  };
}

/**
 * Universal caller to the Google Apps Script Web App.
 * Uses Content-Type: text/plain;charset=utf-8 to eliminate CORS preflight issues.
 * Implements exponential backoff on transient errors (429, 500, 502, 503, 504).
 */
export async function callAppsScript<T = any>(
  action: string,
  payload: Record<string, any> = {}
): Promise<T> {
  const url = (process.env.APPS_SCRIPT_URL || '').trim();
  const token = (process.env.APPS_SCRIPT_TOKEN || '').trim();

  if (!url || !token) {
    throw new Error(
      'Google access not configured: set APPS_SCRIPT_URL and APPS_SCRIPT_TOKEN in environment'
    );
  }

  const requestBody = JSON.stringify({
    token,
    action,
    ...payload,
  });

  let attempt = 0;
  const maxAttempts = 3;
  let lastError: any = null;

  while (attempt < maxAttempts) {
    attempt++;
    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8',
        },
        body: requestBody,
        redirect: 'follow',
      });

      if (!response.ok) {
        if ([429, 500, 502, 503, 504].includes(response.status) && attempt < maxAttempts) {
          const delay = Math.pow(2, attempt) * 500;
          await new Promise((resolve) => setTimeout(resolve, delay));
          continue;
        }
        const text = await response.text();
        throw new Error(`Apps Script responded with HTTP ${response.status}: ${text.slice(0, 200)}`);
      }

      const text = await response.text();
      let parsed: any;
      try {
        parsed = JSON.parse(text);
      } catch (jsonErr) {
        throw new Error(`Apps Script returned invalid JSON response: ${text.slice(0, 200)}`);
      }

      if (parsed && typeof parsed === 'object') {
        if (parsed.ok === false || parsed.error) {
          throw new Error(parsed.error || parsed.message || 'Apps Script returned failure');
        }
      }

      return parsed as T;
    } catch (err: any) {
      lastError = err;
      if (attempt < maxAttempts) {
        const delay = Math.pow(2, attempt) * 500;
        await new Promise((resolve) => setTimeout(resolve, delay));
      }
    }
  }

  throw lastError || new Error(`Apps Script action "${action}" failed after ${maxAttempts} attempts`);
}

/**
 * Check credentials and connectivity. Your Apps Script has no "ping" action,
 * so we call listTabs on the MIS sheet (cheap, and also proves the sheet is readable).
 */
export async function ping(): Promise<{ ok: boolean; time?: string }> {
  await callAppsScript('listTabs', { spreadsheetId: DEFAULT_MIS_SPREADSHEET_ID });
  return { ok: true, time: new Date().toISOString() };
}

/**
 * List the full recursive directory tree under a Google Drive folder.
 */
export async function listTree(folderId?: string): Promise<DriveTreeItem[]> {
  const targetFolderId = folderId || DEFAULT_DRIVE_ROOT_FOLDER_ID;
  const res = await callAppsScript<DriveTreeItem[]>('listTree', {
    folderId: targetFolderId,
  });
  return Array.isArray(res) ? res : [];
}

/**
 * Alias for backward compatibility
 */
export const listFolderTree = listTree;

/**
 * List all worksheet tabs in a Google Spreadsheet.
 */
export async function listTabs(spreadsheetId?: string): Promise<SheetTabInfo[]> {
  const targetId = spreadsheetId || DEFAULT_MIS_SPREADSHEET_ID;
  const res = await callAppsScript<any>('listTabs', {
    spreadsheetId: targetId,
  });

  if (Array.isArray(res)) {
    return res.map((item: any, idx: number) => ({
      sheetId: item.sheetId ?? item.id ?? idx,
      name: item.name ?? item.title ?? `Sheet${idx + 1}`,
      rows: item.rows ?? item.rowCount ?? 0,
      cols: item.cols ?? item.colCount ?? item.columnCount ?? 0,
    }));
  }
  return [];
}

/**
 * Read headers and rows from a specific sheet tab using the Apps Script action "readTab".
 * readTab returns { startRow, totalRows, values } (dates already as yyyy-MM-dd) and reads
 * at most numRows (default 5000) per call, so we page until we have everything (or maxRows).
 */
export async function readSheet(
  spreadsheetId: string,
  tabName: string,
  maxRows?: number
): Promise<SheetDataResult> {
  const targetId = spreadsheetId || DEFAULT_MIS_SPREADSHEET_ID;
  const PAGE = 5000;
  const wanted = maxRows && maxRows > 0 ? maxRows + 1 : Infinity; // +1 for the header row
  const all: any[][] = [];
  let startRow = 1;

  while (all.length < wanted) {
    const numRows = Math.min(PAGE, wanted - all.length);
    const res = await callAppsScript<any>('readTab', {
      spreadsheetId: targetId,
      tab: tabName,
      startRow,
      numRows: Number.isFinite(numRows) ? numRows : PAGE,
    });
    const values: any[][] = Array.isArray(res?.values) ? res.values : [];
    all.push(...values);
    const total = Number(res?.totalRows ?? 0);
    startRow += values.length;
    if (values.length === 0 || startRow > total) break;
  }

  const [headerRow, ...rows] = all;
  const headers: string[] = (headerRow || []).map((c: any) => (c === null || c === undefined ? '' : String(c)));
  return { headers, rows };
}

/**
 * Read sheet values with formatted and raw representations.
 */
export async function readSheetValues(
  spreadsheetId: string,
  tabName: string,
  maxRows?: number
): Promise<{
  rawValues: any[][];
  formattedValues: string[][];
  rowCount: number;
}> {
  const { headers, rows } = await readSheet(spreadsheetId, tabName, maxRows);
  const allRows = headers.length > 0 ? [headers, ...rows] : rows;
  const formattedRows = allRows.map((r) => r.map((c) => (c === null || c === undefined ? '' : String(c))));

  return {
    rawValues: allRows,
    formattedValues: formattedRows,
    rowCount: allRows.length,
  };
}

/**
 * Download a file from Google Drive as a binary Buffer (base64 decoded).
 */
export async function downloadFile(fileId: string): Promise<Buffer> {
  const res = await callAppsScript<any>('downloadFile', { fileId });
  let b64 = '';
  if (typeof res === 'string') {
    b64 = res;
  } else if (res && typeof res === 'object') {
    b64 = res.base64 || res.data || res.content || '';
  }

  if (!b64) {
    throw new Error(`Empty file payload received from Apps Script for file ID: ${fileId}`);
  }

  const cleanB64 = b64.replace(/^data:.*?;base64,/, '');
  return Buffer.from(cleanB64, 'base64');
}

/**
 * Alias for backward compatibility
 */
export const downloadXlsx = downloadFile;
