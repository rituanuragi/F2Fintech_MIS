import React, { useEffect, useMemo, useState } from 'react';
import { Layers, Search, Eye, Loader2, CheckCircle2, AlertTriangle, RefreshCw } from 'lucide-react';
import { SheetItem } from '../types';

interface SheetsViewProps {
  initialSearchTerm?: string;
  onPreviewSheet: (fileId: string, sheetName: string) => void;
}

const PAGE = 100;

export const SheetsView: React.FC<SheetsViewProps> = ({ initialSearchTerm = '', onPreviewSheet }) => {
  const [sheets, setSheets] = useState<SheetItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [term, setTerm] = useState(initialSearchTerm);
  const [indexedOnly, setIndexedOnly] = useState<'all' | 'indexed' | 'not-indexed'>('all');
  const [visible, setVisible] = useState(PAGE);

  const load = async () => {
    setLoading(true);
    setError('');
    try {
      const res = await fetch('/api/sheets');
      const data = await res.json().catch(() => null);
      if (!res.ok || !data || !Array.isArray(data.sheets)) {
        throw new Error((data && data.error) || `Could not load sheets (HTTP ${res.status})`);
      }
      setSheets(data.sheets);
    } catch (e: any) {
      setError(e.message || 'Could not load sheets');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);
  useEffect(() => { setTerm(initialSearchTerm); setVisible(PAGE); }, [initialSearchTerm]);

  const filtered = useMemo(() => {
    const t = term.trim().toLowerCase();
    return sheets.filter((s) => {
      if (indexedOnly === 'indexed' && !s.isIndexed) return false;
      if (indexedOnly === 'not-indexed' && s.isIndexed) return false;
      if (!t) return true;
      return (
        s.fileName.toLowerCase().includes(t) ||
        s.sheetName.toLowerCase().includes(t) ||
        s.folderPath.toLowerCase().includes(t) ||
        s.headers.some((h) => String(h).toLowerCase().includes(t))
      );
    });
  }, [sheets, term, indexedOnly]);

  const totalRows = useMemo(() => filtered.reduce((n, s) => n + (s.rowsCount || 0), 0), [filtered]);

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <h2 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
              <Layers className="w-5 h-5 text-blue-600" />
              <span>Worksheets</span>
            </h2>
            <p className="text-xs text-slate-500 mt-1 font-medium">
              Every worksheet found in the connected Drive workbooks. Open one to see its rows.
            </p>
          </div>
          <button
            onClick={load}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold border border-slate-200 cursor-pointer"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin' : ''}`} /> Refresh
          </button>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <div className="relative flex-1 min-w-[240px]">
            <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
            <input
              value={term}
              onChange={(e) => { setTerm(e.target.value); setVisible(PAGE); }}
              placeholder="Search by workbook, sheet, folder or column name"
              className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-500 focus:bg-white focus:ring-1 focus:ring-blue-500 font-medium"
            />
          </div>
          <div className="flex gap-1.5">
            {(['all', 'indexed', 'not-indexed'] as const).map((k) => (
              <button
                key={k}
                onClick={() => { setIndexedOnly(k); setVisible(PAGE); }}
                className={`px-3 py-2 rounded-xl text-xs font-bold cursor-pointer border ${
                  indexedOnly === k ? 'bg-blue-600 text-white border-blue-600' : 'bg-slate-100 text-slate-700 border-slate-200 hover:bg-slate-200'
                }`}
              >
                {k === 'all' ? 'All' : k === 'indexed' ? 'Indexed' : 'Not indexed'}
              </button>
            ))}
          </div>
        </div>

        {!loading && !error && (
          <div className="text-[11px] text-slate-500 font-mono font-semibold">
            {filtered.length.toLocaleString()} sheets · {totalRows.toLocaleString()} rows
          </div>
        )}
      </div>

      {loading && (
        <div className="bg-white border border-slate-200 rounded-2xl p-10 text-center text-slate-500 text-sm">
          <Loader2 className="w-5 h-5 animate-spin mx-auto mb-2 text-blue-600" /> Loading sheets…
        </div>
      )}

      {error && !loading && (
        <div className="bg-rose-50 border border-rose-200 text-rose-700 rounded-2xl p-5 text-sm flex items-center gap-2">
          <AlertTriangle className="w-4 h-4 shrink-0" /> {error}
        </div>
      )}

      {!loading && !error && filtered.length === 0 && (
        <div className="bg-white border border-dashed border-slate-300 rounded-2xl p-10 text-center text-slate-500 text-sm">
          {sheets.length === 0
            ? 'No sheets yet. Connect a Drive folder and run a sync first.'
            : 'No sheets match your search.'}
        </div>
      )}

      {!loading && !error && filtered.length > 0 && (
        <div className="bg-white border border-slate-200/90 rounded-2xl shadow-xs overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs">
              <thead className="bg-slate-50 text-slate-500 uppercase tracking-wider text-[10px] font-bold">
                <tr>
                  <th className="px-4 py-3">Workbook</th>
                  <th className="px-4 py-3">Sheet</th>
                  <th className="px-4 py-3">Folder</th>
                  <th className="px-4 py-3 text-right">Rows</th>
                  <th className="px-4 py-3 text-right">Cols</th>
                  <th className="px-4 py-3">Status</th>
                  <th className="px-4 py-3 text-right">Open</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filtered.slice(0, visible).map((s) => (
                  <tr key={s.id} className="hover:bg-slate-50/70">
                    <td className="px-4 py-2.5 font-semibold text-slate-900 max-w-[260px] truncate" title={s.fileName}>{s.fileName}</td>
                    <td className="px-4 py-2.5 text-slate-700 max-w-[200px] truncate" title={s.sheetName}>{s.sheetName}</td>
                    <td className="px-4 py-2.5 text-slate-500 font-mono max-w-[200px] truncate" title={s.folderPath}>{s.folderPath || '—'}</td>
                    <td className="px-4 py-2.5 text-right font-mono">{(s.rowsCount || 0).toLocaleString()}</td>
                    <td className="px-4 py-2.5 text-right font-mono">{s.colsCount || 0}</td>
                    <td className="px-4 py-2.5">
                      {s.isIndexed ? (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 font-bold">
                          <CheckCircle2 className="w-3 h-3" /> Indexed
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-amber-50 text-amber-700 border border-amber-200 font-bold">
                          {s.status || 'Not indexed'}
                        </span>
                      )}
                    </td>
                    <td className="px-4 py-2.5 text-right">
                      <button
                        onClick={() => onPreviewSheet(s.fileId, s.sheetName)}
                        className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 font-bold cursor-pointer"
                      >
                        <Eye className="w-3.5 h-3.5" /> Preview
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {visible < filtered.length && (
            <div className="p-3 text-center border-t border-slate-100">
              <button
                onClick={() => setVisible((v) => v + PAGE)}
                className="px-4 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold border border-slate-200 cursor-pointer"
              >
                Show more ({(filtered.length - visible).toLocaleString()} left)
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
