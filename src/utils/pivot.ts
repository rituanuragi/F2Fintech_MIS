/**
 * Pivot summary of a batch search: how many numbers fall under each Priority.
 * The server (search.ts exportBatchLookup) uses the same rule, so the table in the app
 * and the Pivot sheet in the downloaded Excel always show the same numbers.
 */
export interface PivotInput {
  found: boolean;
  priority?: string;
}

export interface PivotRow {
  label: string;
  count: number;
}

export interface PivotData {
  rows: PivotRow[];   // one row per Priority, biggest first
  blanks: number;     // not found, or no priority
  total: number;
}

export function buildPivot(results: PivotInput[]): PivotData {
  const map = new Map<string, number>();
  let blanks = 0;
  for (const r of results) {
    const prio = (r.priority || '').trim();
    if (!r.found || !prio || prio === 'Not Found' || prio === '—') {
      blanks++;
    } else {
      map.set(prio, (map.get(prio) || 0) + 1);
    }
  }
  const rows = Array.from(map.entries())
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
  return { rows, blanks, total: results.length };
}

export interface PivotHeaders {
  title: string;
  dataHeader: string;
  countHeader: string;
  blanksLabel: string;
  totalLabel: string;
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];

export function defaultPivotHeaders(now: Date = new Date()): PivotHeaders {
  return {
    title: `${now.getDate()} ${MONTHS[now.getMonth()]} Login`,
    dataHeader: 'Data',
    countHeader: 'Count',
    blanksLabel: 'Blanks',
    totalLabel: 'Grand Total',
  };
}
