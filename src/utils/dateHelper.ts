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
