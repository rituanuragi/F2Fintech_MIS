import { findColumnValue, pathHasSegment } from './columnMatch.js';
import * as XLSXNamespace from 'xlsx';
const XLSX: any = (XLSXNamespace as any).default || XLSXNamespace;

/**
 * Normalization utilities for phone numbers, names, and field recognition
 */

export interface CleanedPhoneResult {
  phones: string[]; // All valid 10-digit mobile numbers extracted
  primary: string | null; // First valid 10-digit mobile number
  isShort: boolean; // True if digits were found but length is smaller than 10 (< 10)
  isInvalid: boolean; // True if cell has no usable digits or dummy data
  original: string;
}

/**
 * Extracts and cleans Indian mobile numbers according to strict fintech matching rules:
 * 1. Handles scientific notation (e.g. 9.87654321E+09) and float formats (9876543210.0)
 * 2. Splits multi-phone cells (separated by /, ,, ;, &, |, \n, or, and)
 * 3. Strips country codes (+91, 91, 091, 0091)
 * 4. Strips leading 0s (09876543210 -> 9876543210)
 * 5. Strips non-digit formatting (spaces, dashes, parentheses, dots, labels like 'Ph:', 'Mob:')
 * 6. Strictly excludes/ignores entries with length < 10 digits (e.g. 12345, 98765)
 * 7. Filters out dummy/repeated numbers (0000000000, 9999999999, 1111111111)
 * 8. Validates 10-digit mobile numbers starting with valid Indian mobile prefixes (5, 6, 7, 8, 9)
 */
export function cleanAndExtractPhoneNumbers(raw: any): CleanedPhoneResult {
  if (raw === null || raw === undefined) {
    return { phones: [], primary: null, isShort: false, isInvalid: true, original: '' };
  }

  let original = String(raw).trim();
  if (!original) {
    return { phones: [], primary: null, isShort: false, isInvalid: true, original: '' };
  }

  const lower = original.toLowerCase();
  if (['none', 'null', 'nan', 'n/a', 'na', 'nil', '-', '--', '0', '#n/a', 'undefined', 'not available'].includes(lower)) {
    return { phones: [], primary: null, isShort: false, isInvalid: true, original };
  }

  // Handle scientific notation directly on raw value or string
  if (typeof raw === 'number') {
    original = Math.floor(raw).toString();
  } else if (/^\d+(?:\.\d+)?[eE]\+?\d+$/.test(original)) {
    const num = Number(original);
    if (!isNaN(num)) {
      original = Math.floor(num).toString();
    }
  } else if (/^\d+\.0+$/.test(original)) {
    original = original.split('.')[0];
  }

  // Split multi-phone cells by common separators (/, ,, ;, &, |, newline, 'or', 'and')
  const chunks = original.split(/[\/,\;\&\|\n\r\t]+|\s+or\s+|\s+and\s+/i);
  const validPhones: string[] = [];
  let foundAnyDigits = false;
  let hasShortLength = false;

  for (const chunk of chunks) {
    let cleanChunk = chunk.trim();
    if (!cleanChunk) continue;

    // Check scientific notation inside chunk
    if (/^\d+(?:\.\d+)?[eE]\+?\d+$/.test(cleanChunk)) {
      const num = Number(cleanChunk);
      if (!isNaN(num)) cleanChunk = Math.floor(num).toString();
    } else if (/^\d+\.0+$/.test(cleanChunk)) {
      cleanChunk = cleanChunk.split('.')[0];
    }

    // Extract only digits
    let digits = cleanChunk.replace(/\D/g, '');
    if (!digits) continue;

    foundAnyDigits = true;

    // RULE: If length is smaller than 10 digits, do NOT consider that entry as valid mobile
    if (digits.length < 10) {
      hasShortLength = true;
      continue;
    }

    // Strip country code prefixes:
    // 0091 + 10 digits (length >= 14)
    if (digits.startsWith('0091') && digits.length >= 14) {
      digits = digits.slice(4);
    }
    // 091 + 10 digits (length >= 13)
    else if (digits.startsWith('091') && digits.length >= 13) {
      digits = digits.slice(3);
    }
    // 91 + 10 digits (length === 12)
    else if (digits.startsWith('91') && digits.length === 12) {
      digits = digits.slice(2);
    }
    // Leading 0 + 10 digits (length === 11)
    else if (digits.startsWith('0') && digits.length === 11) {
      digits = digits.slice(1);
    }

    // If still > 10 digits, take the trailing 10 digits if they form a valid Indian mobile
    if (digits.length > 10) {
      const trailing10 = digits.slice(-10);
      if (/^[5-9]\d{9}$/.test(trailing10)) {
        digits = trailing10;
      }
    }

    // Validate exactly 10 digits starting with valid Indian mobile prefix [5-9]
    if (digits.length === 10 && /^[5-9]\d{9}$/.test(digits)) {
      // Exclude repeated dummy numbers (e.g. 9999999999, 8888888888, 0000000000)
      if (!/^(\d)\1{9}$/.test(digits)) {
        if (!validPhones.includes(digits)) {
          validPhones.push(digits);
        }
      }
    }
  }

  const primary = validPhones[0] || null;
  const isShort = validPhones.length === 0 && foundAnyDigits && hasShortLength;
  const isInvalid = validPhones.length === 0;

  return {
    phones: validPhones,
    primary,
    isShort,
    isInvalid,
    original,
  };
}

export function normalizePhone(raw: any): {
  normalized: string | null;
  original: string;
  allPhones: string[];
  isShort: boolean;
} {
  const result = cleanAndExtractPhoneNumbers(raw);
  return {
    normalized: result.primary,
    original: result.original,
    allPhones: result.phones,
    isShort: result.isShort,
  };
}

export function normalizeName(raw: any): { normalized: string; original: string } {
  if (raw === null || raw === undefined) return { normalized: '', original: '' };
  const original = String(raw).trim();
  if (!original) return { normalized: '', original: '' };

  // Replace multiple whitespaces, tabs, newlines with single space
  const normalized = original
    .toLowerCase()
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return { normalized, original };
}

export function normalizePan(raw: any): string | null {
  if (!raw) return null;
  const str = String(raw).trim().toUpperCase();
  const panMatch = str.match(/[A-Z]{5}[0-9]{4}[A-Z]{1}/);
  return panMatch ? panMatch[0] : (str.length === 10 ? str : null);
}

const FULL_MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

const SHORT_MONTHS: Record<string, number> = {
  jan: 0, january: 0,
  feb: 1, february: 1,
  mar: 2, march: 2,
  apr: 3, april: 3,
  may: 4,
  jun: 5, june: 5,
  jul: 6, july: 6,
  aug: 7, august: 7,
  sep: 8, sept: 8, september: 8,
  oct: 9, october: 9,
  nov: 10, november: 10,
  dec: 11, december: 11,
};

/**
 * Converts an Excel serial date number (e.g. 45759) to exact calendar day without timezone offset.
 */
export function excelSerialToDateString(serial: number): string {
  const days = Math.round(serial);
  // Excel 1900 epoch (accounting for Excel's 1900 leap year bug)
  const utcMs = Date.UTC(1899, 11, 30) + days * 86400000;
  const d = new Date(utcMs);
  const day = d.getUTCDate();
  const month = FULL_MONTHS[d.getUTCMonth()];
  const yearShort = String(d.getUTCFullYear()).slice(-2);
  return `${day}-${month}-${yearShort}`;
}

/**
 * Formats any MIS/Excel date value (Date object, ISO timestamp, Excel serial number, or formatted string)
 * to a consistent human-readable date string without the 1-day-before timezone shift caused by SheetJS.
 */
export function formatExcelDateValue(val: any, isDateColumn: boolean = true): string {
  if (val === null || val === undefined) return '';

  // Handle Date instance directly (SheetJS cellDates: true creates Date slightly before midnight UTC or local midnight)
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    // Add 12 hours so both UTC-midnight-minus-10s (23:59:50Z prev day) and IST-midnight (18:30:00Z prev day)
    // land squarely inside the exact intended calendar date in UTC
    const adjusted = new Date(val.getTime() + 12 * 3600 * 1000);
    const day = adjusted.getUTCDate();
    const month = FULL_MONTHS[adjusted.getUTCMonth()];
    const yearShort = String(adjusted.getUTCFullYear()).slice(-2);
    return `${day}-${month}-${yearShort}`;
  }

  // Handle SheetJS cell object if passed directly
  if (typeof val === 'object') {
    if (val.t === 'd' && val.v instanceof Date) {
      return formatExcelDateValue(val.v, true);
    }
    if (typeof val.w === 'string' && val.w.trim()) {
      return formatExcelDateValue(val.w.trim(), isDateColumn);
    }
    if (val.v !== undefined) {
      return formatExcelDateValue(val.v, isDateColumn);
    }
    return '';
  }

  // Handle numeric Excel serial date (e.g., 44000 to 50000 corresponds to years ~2020-2036)
  if (typeof val === 'number') {
    if (isDateColumn && val > 35000 && val < 65000) {
      return excelSerialToDateString(val);
    }
    return String(val);
  }

  const str = String(val).trim();
  if (!str) return '';

  // Check if string is a pure Excel serial number in a date column
  if (isDateColumn && /^\d{5}(?:\.\d+)?$/.test(str)) {
    const num = Number(str);
    if (num > 35000 && num < 65000) {
      return excelSerialToDateString(num);
    }
  }

  // Handle ISO 8601 strings produced by JSON.stringify(Date) from SheetJS (e.g. "2026-09-02T23:59:50.000Z" or "2025-04-11T18:29:50.000Z")
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z$/i.test(str)) {
    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      const adjusted = new Date(d.getTime() + 12 * 3600 * 1000);
      const day = adjusted.getUTCDate();
      const month = FULL_MONTHS[adjusted.getUTCMonth()];
      const yearShort = String(adjusted.getUTCFullYear()).slice(-2);
      return `${day}-${month}-${yearShort}`;
    }
  }

  // Handle YYYY-MM-DD
  const ymdMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10);
    const mIdx = parseInt(ymdMatch[2], 10) - 1;
    const day = parseInt(ymdMatch[3], 10);
    if (mIdx >= 0 && mIdx < 12 && day >= 1 && day <= 31) {
      return `${day}-${FULL_MONTHS[mIdx]}-${String(year).slice(-2)}`;
    }
  }

  // Handle DD-MMM-YY or DD-Month-YYYY or DD/MMM/YYYY (e.g., "03-Sep-25", "12-Apr-25", "3-September-2026")
  const dmyAlphaMatch = str.match(/^(\d{1,2})[-/\s]+([A-Za-z]+)[-/\s]+(\d{2,4})$/);
  if (dmyAlphaMatch) {
    const day = parseInt(dmyAlphaMatch[1], 10);
    const mKey = dmyAlphaMatch[2].toLowerCase();
    const yearRaw = dmyAlphaMatch[3];
    if (SHORT_MONTHS[mKey] !== undefined && day >= 1 && day <= 31) {
      const month = FULL_MONTHS[SHORT_MONTHS[mKey]];
      const yearShort = yearRaw.length === 4 ? yearRaw.slice(-2) : yearRaw.padStart(2, '0');
      return `${day}-${month}-${yearShort}`;
    }
  }

  // Handle DD/MM/YYYY or DD-MM-YYYY or MM/DD/YY
  const dmyNumMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (dmyNumMatch && isDateColumn) {
    let p1 = parseInt(dmyNumMatch[1], 10);
    let p2 = parseInt(dmyNumMatch[2], 10);
    const yRaw = dmyNumMatch[3];
    const yearShort = yRaw.length === 4 ? yRaw.slice(-2) : yRaw.padStart(2, '0');
    // In Indian MIS sheets, DD/MM/YYYY is standard unless p2 > 12 (which means MM/DD/YY from US locale export)
    let day = p1;
    let mIdx = p2 - 1;
    if (p2 > 12 && p1 <= 12) {
      mIdx = p1 - 1;
      day = p2;
    }
    if (mIdx >= 0 && mIdx < 12 && day >= 1 && day <= 31) {
      return `${day}-${FULL_MONTHS[mIdx]}-${yearShort}`;
    }
  }

  return str;
}

export function isDateColumnHeader(headerName: string = ''): boolean {
  if (!headerName) return false;
  const clean = headerName.trim();
  if (/\b(amount|amt|eligib|eligibility|eligibilty|score|cibil|remark|remarks|status|type|category|name|poc|bank|banker|lender|id|no|number|lan|phone|mobile|pan|aadhar|aadhaar|candidate)\b/i.test(clean)) {
    return false;
  }
  return /\b(date|dob|dt)\b|case\s*creat(e|ed)?\s*date|login\s*(date|time)|created?\s*date|disburs(ed|al)?\s*date|approv(al|ed)?\s*date|sanction\s*date|filing\s*date|application[_\s]*date/i.test(clean);
}

/**
 * Safely extracts cell primitive value from a SheetJS cell object or raw value,
 * preserving exact dates without off-by-one timezone shifts.
 */
export function extractSafeCellValue(cell: any, headerName: string = ''): any {
  if (cell === null || cell === undefined) return '';
  const isDateCol = isDateColumnHeader(headerName);

  if (cell instanceof Date) {
    return formatExcelDateValue(cell, true);
  }

  if (typeof cell === 'object') {
    // If SheetJS parsed this cell as a Date ('d')
    if (cell.t === 'd' || cell.v instanceof Date) {
      return formatExcelDateValue(cell.v, true);
    }
    // If numeric cell in a date column or formatted as a date in cell.w
    if (cell.t === 'n' && typeof cell.v === 'number') {
      if (isDateCol && cell.v > 35000 && cell.v < 65000) {
        return excelSerialToDateString(cell.v);
      }
      if (typeof cell.w === 'string' && /^\d{1,2}[-/][A-Za-z]{3,9}[-/]\d{2,4}$/.test(cell.w.trim())) {
        return formatExcelDateValue(cell.w.trim(), true);
      }
    }
    if (cell.v !== undefined && cell.v !== null) {
      if (typeof cell.v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(cell.v)) {
        return formatExcelDateValue(cell.v, true);
      }
      return cell.v;
    }
    if (cell.w !== undefined && cell.w !== null) {
      return cell.w;
    }
    return '';
  }

  if (typeof cell === 'number' && isDateCol && cell > 35000 && cell < 65000) {
    return excelSerialToDateString(cell);
  }

  if (typeof cell === 'string' && (isDateCol || /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(cell))) {
    return formatExcelDateValue(cell, isDateCol);
  }

  return cell;
}

// Semantic header detection patterns
export const HEADER_PATTERNS = {
  phone: [
    /phone/i,
    /mob/i,
    /mobile/i,
    /contact/i,
    /number\s*[1-9]?/i,
    /calling/i,
    /cell/i,
    /tel/i,
  ],
  name: [
    /customer.*name/i,
    /applicant.*name/i,
    /candidate.*name/i,
    /client.*name/i,
    /first.*name/i,
    /full.*name/i,
    /^name$/i,
    /^name\s/i,
  ],
  pan: [
    /pan/i,
    /pancard/i,
  ],
  city: [
    /city/i,
    /location/i,
    /branch/i,
    /district/i,
  ],
  state: [
    /state/i,
    /region/i,
  ],
  amount: [
    /amount/i,
    /eligib/i,
    /loan.*amt/i,
    /requested/i,
  ],
  lender: [
    /lender/i,
    /bank/i,
    /partner/i,
  ],
  status: [
    /status/i,
    /stage/i,
    /result/i,
  ],
  email: [
    /email/i,
    /mail/i,
  ]
};

export function classifyHeader(headerName: string): string | null {
  const clean = headerName.trim();
  for (const [key, patterns] of Object.entries(HEADER_PATTERNS)) {
    if (patterns.some(p => p.test(clean))) {
      return key;
    }
  }
  return null;
}

export interface RowSegmentAnalysis {
  sheetLabel: string;
  loginVal: string;
  lotVal: string;
  neetVal: string;
  f2Category?: string;
}

/**
 * Intelligent row analyzer:
 * 1. ONLY LOGIN.xlsx -> extracts the month (e.g. "March", "May", "June") -> "LOGIN - May"
 * 2. ONLY LOT.xlsx -> extracts the lot (e.g. "Lot 1", "Lot 2", "Lot 3", "Lot 4") -> "LOT - Lot 1"
 * 3. ONLY NEET.xlsx -> extracts the year (e.g. "2020", "2019", "2021") -> "NEET - 2020"
 * 4. F2 Leads (MAY.xlsx, JAN.xlsx, f2 Leads directory) -> extracts the category (P1, P2, P3, P4, P5) -> e.g. "MAY.xlsx - P2", "JAN.xlsx - P1"
 * 5. All other files (Fresh Data 21k, Fresh Data 30k, ICICI login, directory, tata points, etc.) use their file name
 */
export function analyzeRowSegment(
  fileName: string,
  sheetName: string,
  rawRow?: Record<string, any>,
  folderPath?: string
): RowSegmentAnalysis {
  const row = rawRow || {};
  let loginVal = '';
  let lotVal = '';
  let neetVal = '';
  let f2Category = '';
  let sheetLabel = '';

  const fName = String(fileName || '').trim();
  const fUpper = fName.toUpperCase();
  const baseName = fName.replace(/\.xlsx?$|\.xlsm?$/i, '');
  const folder = String(folderPath || '').trim();
  const folderUpper = folder.toUpperCase();

  // Exclude connected Google Sheets MIS
  if (folderUpper === 'GOOGLE SHEETS MIS' || fUpper.startsWith('GSHEET_')) {
    return { sheetLabel: '', loginVal: '', lotVal: '', neetVal: '', f2Category: '' };
  }

  // Specific file matching to avoid false positives
  const isExactLoginFile = fUpper === 'LOGIN.XLSX' || (fUpper.startsWith('LOGIN.') && !fUpper.includes('ICICI'));
  const isExactLotFile = fUpper === 'LOT.XLSX' || fUpper.startsWith('LOT.');
  const isExactNeetFile = fUpper === 'NEET.XLSX' || fUpper.startsWith('NEET.');

  // F2 Leads files
  const isF2Leads =
    folderUpper.includes('F2 LEADS') ||
    pathHasSegment(folderUpper, 'F2', 'F2 LEADS') ||
    fUpper.includes('F2 LEADS') ||
    fUpper === 'JAN.XLSX' ||
    fUpper === 'MAY.XLSX' ||
    fUpper.startsWith('JAN.') ||
    fUpper.startsWith('MAY.');

  if (isExactLoginFile) {
    // 1. Analyze LOGIN (e.g. "login May", "login March")
    const rawLogin = findColumnValue(row, ['login', 'login month', 'month']);
    if (rawLogin) {
      const cleanMonth = rawLogin.value.replace(/^login[\s_-]*/i, '').trim() || rawLogin.value;
      loginVal = `Login-${cleanMonth}`;
      sheetLabel = `LOGIN - ${cleanMonth}`;
    } else {
      sheetLabel = 'LOGIN';
    }
  } else if (isExactLotFile) {
    // 2. Analyze LOT (e.g. "lot-1", "lot-2", "lot-3", "lot-4")
    const rawLot = findColumnValue(row, ['lot', 'lots', 'lot no', 'lot number']);
    if (rawLot) {
      const lotStr = rawLot.value;
      const numMatch = lotStr.match(/\d+/);
      const lotText = lotStr.replace(/^lot[\s_-]*/i, '').trim() || lotStr;
      lotVal = `Lot-${numMatch ? numMatch[0] : lotText}`;
      sheetLabel = numMatch ? `LOT - Lot ${numMatch[0]}` : `LOT - ${lotStr}`;
    } else {
      sheetLabel = 'LOT';
    }
  } else if (isExactNeetFile) {
    // 3. Analyze NEET (e.g. "neet 2020", "neet 2019", "neet 2021")
    const rawYear = findColumnValue(row, ['year', 'neet year', 'passing year', 'batch']);
    if (rawYear) {
      const cleanYear = rawYear.value.replace(/^neet[\s_-]*/i, '').trim() || rawYear.value;
      neetVal = `NEET-${cleanYear}`;
      sheetLabel = `NEET - ${cleanYear}`;
    } else {
      sheetLabel = 'NEET';
    }
  } else if (isF2Leads) {
    // 4. F2 Leads: Extract category (e.g. P1, P2, P3, P4, P5) from category column or sheetName
    let cat = '';

    // First check rawRow for category column (case insensitive)
    const candidateKeys = [
      'category', 'Category', 'CATEGORY',
      'priority', 'Priority', 'PRIORITY',
      'lead category', 'Lead Category', 'LEAD CATEGORY',
      'lead_category', 'Lead_Category',
      'p', 'P'
    ];
    for (const key of candidateKeys) {
      if (row[key] !== undefined && row[key] !== null) {
        const val = String(row[key]).trim();
        if (val) {
          const pMatch = val.match(/\b(P\s*\d+)\b/i);
          if (pMatch) {
            cat = pMatch[1].replace(/\s+/g, '').toUpperCase();
            break;
          } else if (/^\d+$/.test(val)) {
            cat = `P${val}`;
            break;
          } else if (/^P[A-Z0-9]+$/i.test(val)) {
            cat = val.toUpperCase();
            break;
          }
        }
      }
    }

    // Next check any key in row that contains 'category' or 'priority'
    if (!cat) {
      for (const [colName, val] of Object.entries(row)) {
        if (!val) continue;
        const colLower = colName.toLowerCase().trim();
        if (colLower.includes('category') || colLower === 'priority') {
          const strVal = String(val).trim();
          const pMatch = strVal.match(/\b(P\s*\d+)\b/i);
          if (pMatch) {
            cat = pMatch[1].replace(/\s+/g, '').toUpperCase();
            break;
          } else if (/^\d+$/.test(strVal)) {
            cat = `P${strVal}`;
            break;
          }
        }
      }
    }

    // If still not found, check sheetName (e.g. in MAY.xlsx, sheet names are P1, P2, P3, P4)
    if (!cat && sheetName) {
      const sTrim = sheetName.trim();
      const pMatch = sTrim.match(/\b(P\s*\d+)\b/i);
      if (pMatch) {
        cat = pMatch[1].replace(/\s+/g, '').toUpperCase();
      } else if (/^\d+$/.test(sTrim)) {
        cat = `P${sTrim}`;
      }
    }

    f2Category = cat;
    // Format sheetLabel as requested: "MAY.xlsx - P2", "JAN.xlsx - P1", etc.
    if (cat) {
      sheetLabel = `${fName} - ${cat}`;
    } else {
      sheetLabel = fName;
    }
  } else {
    // 5. All other sheets show their file name
    sheetLabel = fName;
  }

  return { sheetLabel, loginVal, lotVal, neetVal, f2Category };
}

/**
 * Extract Month-Year label (e.g. "Jan 2026", "Sept 2026", "Oct 2026") from various date formats,
 * sheet names, or fallback strings.
 */
export function extractMonthYearLabel(dateVal: any, fallbackStr?: string): string {
  if (!dateVal && !fallbackStr) return 'Other';
  const s = String(dateVal || fallbackStr || '').trim();
  if (!s || s === '—' || s === 'Not Found') return 'Other';

  // 1. Try parsing standard date formats (YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY)
  const isoMatch = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    const year = isoMatch[1];
    const monthNum = parseInt(isoMatch[2], 10);
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];
    if (monthNum >= 1 && monthNum <= 12) {
      return `${months[monthNum - 1]} ${year}`;
    }
  }

  const dmyMatch = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
  if (dmyMatch) {
    let year = dmyMatch[3];
    if (year.length === 2) year = '20' + year;
    const monthNum = parseInt(dmyMatch[2], 10);
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];
    if (monthNum >= 1 && monthNum <= 12) {
      return `${months[monthNum - 1]} ${year}`;
    }
  }

  const monthMap: Record<string, string> = {
    jan: 'Jan', january: 'Jan',
    feb: 'Feb', february: 'Feb',
    mar: 'Mar', march: 'Mar',
    apr: 'Apr', april: 'Apr',
    may: 'May',
    jun: 'Jun', june: 'Jun',
    jul: 'Jul', july: 'Jul',
    aug: 'Aug', august: 'Aug',
    sep: 'Sept', sept: 'Sept', september: 'Sept',
    oct: 'Oct', october: 'Oct',
    nov: 'Nov', november: 'Nov',
    dec: 'Dec', december: 'Dec'
  };

  const monthRegex = /\b(january|jan|february|feb|march|mar|april|apr|may|june|jun|july|jul|august|aug|september|sept|sep|october|oct|november|nov|december|dec)\b/i;
  const mMatch = s.match(monthRegex);
  if (mMatch) {
    const normMonth = monthMap[mMatch[1].toLowerCase()] || 'Jan';
    const yMatch = s.match(/\b(20\d{2}|\d{2})\b/);
    let year = '2026';
    if (yMatch) {
      year = yMatch[1].length === 2 ? '20' + yMatch[1] : yMatch[1];
    }
    return `${normMonth} ${year}`;
  }

  const parsed = Date.parse(s);
  if (!isNaN(parsed)) {
    const d = new Date(parsed);
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];
    return `${months[d.getMonth()]} ${d.getFullYear()}`;
  }

  return 'Other';
}

/**
 * Checks if a cell value contains substantive data rather than being empty,
 * whitespace, or common spreadsheet placeholder/formula default (e.g. '-', '--', 'N/A', '#REF!', etc.).
 */
export function isMeaningfulCellValue(val: any, headerName: string = ''): boolean {
  if (val === null || val === undefined) return false;
  
  let raw = val;
  if (typeof val === 'object') {
    if (val.v !== undefined) raw = val.v;
    else if (val.w !== undefined) raw = val.w;
  }
  if (raw === null || raw === undefined) return false;

  if (typeof raw === 'number') {
    return !isNaN(raw);
  }
  if (typeof raw === 'boolean') return true;
  if (raw instanceof Date) return !isNaN(raw.getTime());

  const str = String(raw).replace(/\u00A0/g, ' ').trim();
  if (str.length === 0) return false;

  const lower = str.toLowerCase();
  if (
    lower === '-' ||
    lower === '--' ||
    lower === '---' ||
    lower === '–' ||
    lower === '—' ||
    lower === '.' ||
    lower === '..' ||
    lower === '...' ||
    lower === 'none' ||
    lower === 'null' ||
    lower === 'nil' ||
    lower === 'nan' ||
    lower === 'n/a' ||
    lower === 'na' ||
    lower === '#n/a' ||
    lower === '#value!' ||
    lower === '#ref!' ||
    lower === '#name?' ||
    lower === '#num!' ||
    lower === '#div/0!' ||
    lower === '#null!' ||
    lower === 'undefined' ||
    lower === 'not available' ||
    lower === 'blank' ||
    lower === 'empty'
  ) {
    return false;
  }

  // Matches strings that only contain dashes, underscores, spaces, dots, or slashes (e.g. " - ", "--", "-/-")
  if (/^[-–—_\s./\\]+$/.test(str)) {
    return false;
  }

  return true;
}

/**
 * Checks whether a single cell contains a valid phone number (excluding blanks, formulas, placeholders, and dummy numbers).
 */
export function isValidPhoneCellValue(val: any): boolean {
  if (val === null || val === undefined) return false;
  const s = String(typeof val === 'object' && val !== null && val.v !== undefined ? val.v : val).trim();
  if (!s) return false;
  const lower = s.toLowerCase();
  if ([
    'none', 'null', 'nan', 'n/a', 'na', 'nil', '-', '--', '—', '0',
    '#n/a', '#value!', '#ref!', '#name?', '#num!', '#div/0!', '#null!',
    'undefined', 'not available', 'blank', 'empty'
  ].includes(lower)) {
    return false;
  }
  const digits = s.replace(/\D/g, '');
  if (digits.length < 7) return false;
  // Discard repeated dummy digits (e.g. 0000000, 1111111, 9999999)
  if (/^(\d)\1{6,}$/.test(digits)) return false;
  return true;
}

/**
 * Checks whether a row contains a valid phone number.
 * If phoneColIndex is provided, checks that column first, then optionally checks other cells if needed.
 */
export function rowHasPhoneNumber(row: any[], phoneColIndex?: number): boolean {
  if (!row || !Array.isArray(row) || row.length === 0) return false;

  // 1. Check primary phone column if index is given and in bounds
  if (phoneColIndex !== undefined && phoneColIndex >= 0 && phoneColIndex < row.length) {
    const cell = row[phoneColIndex];
    if (isValidPhoneCellValue(cell)) {
      return true;
    }
  }

  // 2. Fallback: check other cells in row for a valid phone number
  for (let c = 0; c < row.length; c++) {
    if (c === phoneColIndex) continue;
    const cell = row[c];
    if (isValidPhoneCellValue(cell)) {
      return true;
    }
  }

  return false;
}

/**
 * Strict evaluation of whether an Excel or Google Sheet row contains ANY genuine entry.
 * Explicitly rejects:
 * - Empty rows
 * - Rows where all cells are blanks, whitespace, or dashes/placeholders (e.g. table templates)
 * - Rows that only contain a pre-populated serial number (e.g. 1, 2, 3...) with no other data
 * - Rows that only contain Excel formula artifacts (e.g. CASE AGEING: 46298, AGEING DATE: 3-October-26) with no customer/ticket/phone data
 */
export function rowHasAnyEntry(row: any[], headers: string[] = []): boolean {
  if (!row || !Array.isArray(row) || row.length === 0) return false;

  let hasCustomerName = false;
  let hasPhone = false;
  let hasTicketId = false;
  let hasPanOrAadhaar = false;
  let hasLenderOrAmount = false;
  let otherRealDataCount = 0;
  let meaningfulCells = 0;

  for (let c = 0; c < row.length; c++) {
    const cell = row[c];
    const header = headers[c] || `Column_${c + 1}`;
    if (!isMeaningfulCellValue(cell, header)) continue;

    const val = typeof cell === 'object' && cell !== null && cell.v !== undefined ? cell.v : cell;
    const str = String(val).trim();
    if (!str) continue;

    const hUpper = header.toUpperCase().trim();

    // Ignore pure formula / counter / aging columns
    if (/^(S\.?NO|SR|SERIAL|ROW|INDEX|#)$/i.test(hUpper)) continue;
    if (/\b(AGEING|AGING|CASE\s*AGEING|AGEING\s*DATE|DAYS|DIFF)\b/i.test(hUpper)) continue;

    meaningfulCells++;

    // Check specific semantic columns
    if (/CUSTOMER|APPLICANT|BORROWER|DOCTOR|CLIENT|PARTY|CANDIDATE/i.test(hUpper) && !/POC|BANKER|EMP|TL|MANAGER|LENDER/i.test(hUpper)) {
      if (str.length >= 2 && !/^\d+$/.test(str) && !['none', 'null', 'nan', 'n/a', 'na', '-', '--'].includes(str.toLowerCase())) {
        hasCustomerName = true;
      }
    } else if (/PHONE|MOBILE|CONTACT|CELL/i.test(hUpper) && !/BANKER/i.test(hUpper)) {
      const digits = str.replace(/\D/g, '');
      if (digits.length >= 7) hasPhone = true;
    } else if (/^(TICKET\s*ID|TICKET\s*NO|LEAD\s*ID|APPLICATION\s*ID|CASE\s*ID|LAN\s*NO|APP\s*ID|DEAL\s*ID)$/i.test(hUpper)) {
      if (str.length >= 2 && !['none', 'null', 'nan', 'n/a', 'na', '-', '--'].includes(str.toLowerCase())) {
        hasTicketId = true;
      }
    } else if (/PAN|AADHAAR|AADHAR/i.test(hUpper)) {
      if (str.length >= 5) hasPanOrAadhaar = true;
    } else if (/LENDER|LOAN\s*AMOUNT|NET\s*LOAN\s*AMOUNT|DISBURSED\s*DATE|APPROVAL\s*DATE/i.test(hUpper)) {
      if (str.length >= 2 && !['none', 'null', 'nan', 'n/a', 'na', '-', '--'].includes(str.toLowerCase())) {
        hasLenderOrAmount = true;
      }
    } else {
      // In case headers are unknown (e.g. Column_1, Column_2): inspect cell contents directly
      const cleanDigits = str.replace(/\D/g, '');
      if (cleanDigits.length >= 10 && /^[5-9]\d{9}$/.test(cleanDigits.slice(-10))) {
        hasPhone = true;
      } else if (/^[A-Z]{5}[0-9]{4}[A-Z]$/.test(str.toUpperCase())) {
        hasPanOrAadhaar = true;
      } else if (str.length >= 3 && !/^\d+$/.test(str) && !['none', 'null', 'nan', 'n/a', 'na', '-', '--'].includes(str.toLowerCase())) {
        otherRealDataCount++;
      }
    }
  }

  // A genuine business entry must have at least one substantive identifier
  if (hasCustomerName || hasPhone || hasTicketId || hasPanOrAadhaar) return true;
  if (hasLenderOrAmount && otherRealDataCount >= 2) return true;
  // If headers were generic Column_X but row has multiple real text values
  if (otherRealDataCount >= 2 && meaningfulCells >= 2) return true;

  return false;
}

/**
 * Strips blank rows, empty formula artifacts, and table template rows from all worksheets in an Excel workbook.
 * Ensures the worksheets retain only headers and rows with genuine business entries.
 */
export function cleanWorkbookSheets(workbook: any): {
  totalStripped: number;
  sheetStats: Record<string, { original: number; cleaned: number }>;
} {
  let totalStripped = 0;
  const sheetStats: Record<string, { original: number; cleaned: number }> = {};
  if (!workbook || !workbook.SheetNames || !Array.isArray(workbook.SheetNames)) {
    return { totalStripped: 0, sheetStats };
  }

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const rawRows = (XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }) as any[][]) || [];
    if (rawRows.length <= 1) {
      sheetStats[sheetName] = { original: rawRows.length, cleaned: rawRows.length };
      continue;
    }

    const headers = (rawRows[0] || []).map((h: any, i: number) => String(h || `Column_${i + 1}`).trim());
    const cleanRows: any[][] = [rawRows[0]];

    for (let r = 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (row && rowHasAnyEntry(row, headers)) {
        cleanRows.push(row);
      }
    }

    const stripped = rawRows.length - cleanRows.length;
    totalStripped += stripped;
    sheetStats[sheetName] = { original: rawRows.length, cleaned: cleanRows.length };

    // Update worksheet with clean rows only
    workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(cleanRows);
  }

  return { totalStripped, sheetStats };
}


