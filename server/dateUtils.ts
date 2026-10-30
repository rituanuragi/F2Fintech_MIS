/**
 * Date and currency normalization utilities for F2 Fintech MIS
 */

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

/**
 * Normalizes any date value (Excel serial number, string with month names,
 * DD/MM/YYYY, ISO string, etc.) into ISO format YYYY-MM-DD.
 * Returns null if the value is unparseable or blank.
 */
export function toISODate(value: any): string | null {
  if (value === null || value === undefined) return null;

  // 1. If it's already a JS Date object
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // 2. Excel/Sheets serial number (between 20000 and 80000, roughly 1954 to 2119)
  if (typeof value === 'number' || (!isNaN(Number(value)) && typeof value === 'string' && /^\d+(\.\d+)?$/.test(String(value).trim()))) {
    const num = typeof value === 'number' ? value : parseFloat(String(value).trim());
    if (num >= 20000 && num <= 80000) {
      // Excel epoch starts at 1899-12-30 (accounting for Lotus 1-2-3 1900 leap year bug)
      const wholeDays = Math.floor(num);
      const epoch = new Date(Date.UTC(1899, 11, 30));
      epoch.setUTCDate(epoch.getUTCDate() + wholeDays);
      const y = epoch.getUTCFullYear();
      const m = String(epoch.getUTCMonth() + 1).padStart(2, '0');
      const d = String(epoch.getUTCDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
  }

  let str = String(value).trim();
  if (!str) return null;

  // Quick check for obvious non-dates
  const lower = str.toLowerCase();
  if (['n/a', 'na', 'null', 'none', '-', '--', 'nil', 'undefined', 'not available', '0'].includes(lower)) {
    return null;
  }

  // Collapse whitespace
  str = str.replace(/\s+/g, ' ');

  // Strip trailing time parts (e.g. "14:32:00", "02:32 PM", "2:32pm", "T14:32:00.000Z")
  str = str.replace(/T\d{2}:\d{2}:\d{2}.*$/, '');
  str = str.replace(/\s+\d{1,2}:\d{2}(:\d{2})?(\s*(am|pm))?$/i, '');

  // Strip ordinal suffixes: 1st, 2nd, 3rd, 4th, etc.
  str = str.replace(/\b(\d+)(st|nd|rd|th)\b/gi, '$1');

  // Standardize year helper: 2-digit years 00-69 => 2000-2069, 70-99 => 1970-1999
  const normalizeYear = (yr: number): number => {
    if (yr < 70) return 2000 + yr;
    if (yr < 100) return 1900 + yr;
    return yr;
  };

  const formatResult = (y: number, m: number, d: number): string | null => {
    const ny = normalizeYear(y);
    if (ny < 1900 || ny > 2100) return null;
    if (m < 1 || m > 12) return null;
    const maxDays = new Date(ny, m, 0).getDate();
    if (d < 1 || d > maxDays) return null;
    return `${ny}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  };

  // 3. Match ISO format YYYY-MM-DD or YYYY/MM/DD
  const isoMatch = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (isoMatch) {
    return formatResult(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10), parseInt(isoMatch[3], 10));
  }

  // 4. Match textual month formats:
  // Examples: "2-October-26", "02-Oct-2026", "2 October 2026", "Oct 2, 2026", "October 2 2026", "2-Oct-26"
  // Day first: (\d{1,2})[-/\s]+([a-zA-Z]+)[-/\s,]+(\d{2,4})
  const dayFirstText = str.match(/^(\d{1,2})[-/\s]+([a-zA-Z]+)[-/\s,]+(\d{2,4})$/);
  if (dayFirstText) {
    const d = parseInt(dayFirstText[1], 10);
    const mStr = dayFirstText[2].toLowerCase();
    const m = MONTH_NAMES[mStr];
    const y = parseInt(dayFirstText[3], 10);
    if (m) {
      return formatResult(y, m, d);
    }
  }

  // Month first: ([a-zA-Z]+)[-/\s]+(\d{1,2})[-/\s,]+(\d{2,4})
  const monthFirstText = str.match(/^([a-zA-Z]+)[-/\s]+(\d{1,2})[-/\s,]+(\d{2,4})$/);
  if (monthFirstText) {
    const mStr = monthFirstText[1].toLowerCase();
    const m = MONTH_NAMES[mStr];
    const d = parseInt(monthFirstText[2], 10);
    const y = parseInt(monthFirstText[3], 10);
    if (m) {
      return formatResult(y, m, d);
    }
  }

  // 5. Numeric day-first: DD/MM/YYYY, DD-MM-YYYY, DD.MM.YY
  const numMatch = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (numMatch) {
    let p1 = parseInt(numMatch[1], 10);
    let p2 = parseInt(numMatch[2], 10);
    const y = parseInt(numMatch[3], 10);

    // Default interpretation: Day = p1, Month = p2
    // If day-first is impossible (month > 12) but day <= 12, swap
    if (p2 > 12 && p1 <= 12) {
      const temp = p1;
      p1 = p2;
      p2 = temp;
    }

    return formatResult(y, p2, p1);
  }

  return null;
}

/**
 * Parses financial amount strings (e.g. "1,50,000", "₹ 2.5 L", "3 Lakh", "5 Cr", "500000")
 * into a standard numeric float.
 */
export function parseAmount(val: any): number {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;

  let str = String(val).trim();
  if (!str) return 0;

  // Clean currency symbols and commas
  str = str.replace(/[₹$,]/g, '').trim();

  // Handle Lakh / L
  const lakhMatch = str.match(/^([\d.]+)\s*(lakh|lacs?|lac|l)\b/i);
  if (lakhMatch) {
    const num = parseFloat(lakhMatch[1]);
    return isNaN(num) ? 0 : num * 100000;
  }

  // Handle Crore / Cr
  const crMatch = str.match(/^([\d.]+)\s*(crores?|cr)\b/i);
  if (crMatch) {
    const num = parseFloat(crMatch[1]);
    return isNaN(num) ? 0 : num * 10000000;
  }

  // Handle K / Thousand
  const kMatch = str.match(/^([\d.]+)\s*(k|thousand)\b/i);
  if (kMatch) {
    const num = parseFloat(kMatch[1]);
    return isNaN(num) ? 0 : num * 1000;
  }

  const cleanNum = parseFloat(str.replace(/[^\d.-]/g, ''));
  return isNaN(cleanNum) ? 0 : cleanNum;
}

/**
 * Built-in self-test assertions for toISODate.
 * Returns true if all unit tests pass, throws if any fail.
 */
export function runDateUtilsTests(): { passed: boolean; testCount: number } {
  const tests: [any, string | null][] = [
    [46297, '2026-10-02'], // Excel serial 46297 is 2026-10-02
    ['2-October-26', '2026-10-02'],
    [' 02/10/2026 14:32:00 ', '2026-10-02'],
    ['2nd Oct 2026', '2026-10-02'],
    ['2026-10-02', '2026-10-02'],
    ['', null],
    ['N/A', null],
    ['2.10.26', '2026-10-02'],
    ['10/02/2026', '2026-02-10'], // Day first: 10th Feb 2026
    ['25/10/2026', '2026-10-25'], // Day first: 25th Oct 2026
    ['10/25/2026', '2026-10-25'], // Month > 12 swap: 25th Oct 2026
  ];

  for (const [input, expected] of tests) {
    const res = toISODate(input);
    if (res !== expected) {
      throw new Error(`toISODate failed for input: ${JSON.stringify(input)}. Expected: ${expected}, Got: ${res}`);
    }
  }

  return { passed: true, testCount: tests.length };
}
