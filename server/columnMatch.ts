/**
 * Shared helpers to find a value in a row by column name, even when the header
 * is written differently ("Lot No", "LOT_NO", "lots", "Neet Year", "Login Month").
 * No imports here on purpose, so any server file can use it without cycles.
 */

const EMPTY_VALUES = new Set(['', '-', '—', 'na', 'n/a', 'nan', 'null', 'undefined', 'none']);

export function isEmptyCell(v: any): boolean {
  if (v === undefined || v === null) return true;
  return EMPTY_VALUES.has(String(v).trim().toLowerCase());
}

/** "Lot No" -> "lotno", "LOT_NO" -> "lotno" */
export function squash(s: string): string {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

/** "Lot No" -> ["lot","no"], "lotNo" -> ["lot","no"], "NEET_Year" -> ["neet","year"] */
export function tokens(s: string): string[] {
  return String(s || '')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter(Boolean);
}

/** Does this header match the wanted column name? */
export function headerMatches(header: string, wanted: string): boolean {
  const h = squash(header);
  const w = squash(wanted);
  if (!h || !w) return false;
  if (h === w) return true;
  // whole-word match only (so "lot" matches "Lot No" but NOT "Allotment" or "Pilot")
  const wTokens = tokens(wanted);
  const hTokens = tokens(header);
  if (wTokens.length === 0 || hTokens.length === 0) return false;
  if (wTokens.length === 1) return hTokens.includes(wTokens[0]);
  // multi-word wanted ("passing year"): all its words must be in the header
  return wTokens.every((t) => hTokens.includes(t));
}

/**
 * Find the first non-empty value for any of the wanted column names.
 * Exact names are tried first (in the order given), then flexible matches.
 */
export function findColumnValue(
  row: Record<string, any> | undefined | null,
  wantedNames: string[]
): { column: string; value: string } | null {
  if (!row || typeof row !== 'object') return null;
  const keys = Object.keys(row);
  if (keys.length === 0) return null;

  // 1) exact / squashed-equal match, in the order the user listed the columns
  for (const wanted of wantedNames) {
    const w = squash(wanted);
    if (!w) continue;
    for (const k of keys) {
      if (squash(k) === w && !isEmptyCell(row[k])) {
        return { column: k, value: String(row[k]).trim() };
      }
    }
  }
  // 2) flexible whole-word match
  for (const wanted of wantedNames) {
    for (const k of keys) {
      if (headerMatches(k, wanted) && !isEmptyCell(row[k])) {
        return { column: k, value: String(row[k]).trim() };
      }
    }
  }
  return null;
}

/** Like findColumnValue but returns one value per wanted column (in order), skipping empty ones. */
export function findEachColumnValue(
  row: Record<string, any> | undefined | null,
  wantedNames: string[]
): string[] {
  const out: string[] = [];
  const used = new Set<string>();
  for (const wanted of wantedNames) {
    const hit = findColumnValue(row, [wanted]);
    if (hit && !used.has(hit.column)) {
      used.add(hit.column);
      out.push(hit.value);
    }
  }
  return out;
}

/** True when any path segment equals one of the wanted names (case-insensitive). */
export function pathHasSegment(path: string, ...names: string[]): boolean {
  const segs = String(path || '').toUpperCase().split(/[\\/]+/).map((s) => s.trim()).filter(Boolean);
  const wanted = names.map((n) => n.toUpperCase());
  return segs.some((s) => wanted.includes(s));
}

export function lastSegment(path: string): string {
  const segs = String(path || '').split(/[\\/]+/).map((s) => s.trim()).filter(Boolean);
  return (segs[segs.length - 1] || '').toUpperCase();
}
