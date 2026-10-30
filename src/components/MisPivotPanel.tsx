import React, { useEffect, useMemo, useState } from 'react';
import { Table2, Download, Loader2, AlertCircle } from 'lucide-react';
import { buildPivot } from '../utils/pivot';

interface MisPivotPanelProps {
  sheetId: string;
  tabName: string;
  date: string;          // ISO yyyy-mm-dd
  dateColumn: string;
  refreshKey?: string;   // changes when the dashboard data is refreshed
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];

function pivotTitleFor(iso: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!m) return 'Login';
  return `${parseInt(m[3], 10)} ${MONTHS[parseInt(m[2], 10) - 1]} Login`;
}

/** Same number written in different ways = one number (last 10 digits). */
function phoneKey(raw: string): string {
  const d = String(raw || '').replace(/\D/g, '');
  return d.length >= 10 ? d.slice(-10) : '';
}

interface BatchRow {
  inputPhone: string;
  normPhone: string;
  found: boolean;
  priority?: string;
  [k: string]: any;
}

export const MisPivotPanel: React.FC<MisPivotPanelProps> = ({ sheetId, tabName, date, dateColumn, refreshKey }) => {
  const [rows, setRows] = useState<BatchRow[]>([]);
  const [entryCount, setEntryCount] = useState(0);
  const [noPhone, setNoPhone] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [downloading, setDownloading] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      setLoading(true);
      setError(null);
      try {
        // 1. all MIS entries of the selected date (no search / status / lender filter)
        const qRes = await fetch('/api/mis-dashboard/query', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ sheetId, tabName, date, dateColumn, searchQuery: '', statusFilter: 'all', lenderFilter: 'all' }),
        });
        const qData = await qRes.json();
        if (!qRes.ok) throw new Error(qData.error || 'Could not read MIS entries');
        const entries: { phone?: string }[] = qData.entries || [];

        // 2. one entry per unique phone number
        const seen = new Set<string>();
        const phones: string[] = [];
        let missing = 0;
        for (const e of entries) {
          const k = phoneKey(e.phone || '');
          if (!k) { missing++; continue; }
          if (seen.has(k)) continue;
          seen.add(k);
          phones.push(k);
        }

        // 3. look each unique number up in the directory database (priority)
        let results: BatchRow[] = [];
        if (phones.length > 0) {
          const bRes = await fetch('/api/batch', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ phones, scope: 'all' }),
          });
          const bData = await bRes.json();
          if (!bRes.ok) throw new Error(bData.error || 'Priority lookup failed');
          results = bData.results || [];
        }
        if (!cancelled) {
          setRows(results);
          setEntryCount(entries.length);
          setNoPhone(missing);
        }
      } catch (err: any) {
        if (!cancelled) {
          setError(err.message || 'Could not build the pivot');
          setRows([]);
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [sheetId, tabName, date, dateColumn, refreshKey]);

  const pivot = useMemo(() => buildPivot(rows), [rows]);
  const title = pivotTitleFor(date);
  const duplicatesRemoved = Math.max(0, entryCount - noPhone - rows.length);

  const handleDownload = async () => {
    setDownloading(true);
    try {
      const res = await fetch('/api/batch/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ results: rows, format: 'xlsx', includePivot: true, pivotTitle: title, useTitleAsTabName: true }),
      });
      if (!res.ok) throw new Error('Download failed');
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `MIS_Pivot_${date}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (e: any) {
      alert(e.message || 'Download failed');
    } finally {
      setDownloading(false);
    }
  };

  const cell = 'border border-slate-800 px-3 py-2 text-sm';

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
      <div className="px-5 py-3 border-b border-slate-200 bg-slate-50 flex flex-wrap items-center gap-3">
        <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
          <Table2 className="w-4 h-4 text-violet-600" /> Pivot by Priority
        </h3>
        <span className="text-[11px] text-slate-500">
          Each phone number counted once. Priority comes from the directory database; numbers not found are Blanks.
        </span>
        <button
          onClick={handleDownload}
          disabled={downloading || loading || rows.length === 0}
          className="ml-auto inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 disabled:opacity-50"
        >
          {downloading ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Download className="w-3.5 h-3.5" />}
          Download Excel
        </button>
      </div>

      <div className="p-5">
        {loading ? (
          <div className="flex items-center gap-2 text-sm text-slate-500"><Loader2 className="w-4 h-4 animate-spin" /> Building pivot…</div>
        ) : error ? (
          <div className="flex items-center gap-2 text-sm text-rose-700"><AlertCircle className="w-4 h-4" /> {error}</div>
        ) : rows.length === 0 ? (
          <div className="text-sm text-slate-500">No entries with a phone number on this date.</div>
        ) : (
          <>
            <table className="w-full max-w-md border-collapse">
              <thead>
                <tr><th colSpan={2} className="bg-[#5b1a85] text-white text-lg font-bold py-2.5 border border-slate-800">{title}</th></tr>
                <tr>
                  <th className={`${cell} text-left font-bold`}>Data</th>
                  <th className={`${cell} text-right font-bold`}>Count</th>
                </tr>
              </thead>
              <tbody>
                {pivot.rows.map((r) => (
                  <tr key={r.label}>
                    <td className={cell}>{r.label}</td>
                    <td className={`${cell} text-right`}>{r.count}</td>
                  </tr>
                ))}
                <tr>
                  <td className={cell}>Blanks</td>
                  <td className={`${cell} text-right`}>{pivot.blanks}</td>
                </tr>
                <tr className="bg-violet-50 font-bold">
                  <td className={cell}>Grand Total</td>
                  <td className={`${cell} text-right`}>{pivot.total}</td>
                </tr>
              </tbody>
            </table>
            <p className="mt-3 text-[11px] text-slate-500">
              {entryCount} MIS entries on this date → {rows.length} unique numbers
              {duplicatesRemoved > 0 ? ` (${duplicatesRemoved} repeated number${duplicatesRemoved > 1 ? 's' : ''} counted once)` : ''}
              {noPhone > 0 ? `. ${noPhone} entr${noPhone > 1 ? 'ies' : 'y'} without a valid phone number not counted` : ''}.
            </p>
          </>
        )}
      </div>
    </div>
  );
};
