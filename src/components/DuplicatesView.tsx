import React, { useCallback, useEffect, useState } from 'react';
import { Copy, Search, Download, Loader2, ChevronDown, ChevronRight, AlertTriangle, Maximize2, UserSearch } from 'lucide-react';
import { DuplicateGroupItem, DuplicateFullStats } from '../types';
import { downloadFromPost } from '../utils/download';

interface DuplicatesViewProps {
  onViewRecord: (id: number) => void;
  onPreviewSheet?: (fileId: string, sheetName: string) => void;
  onLookupCustomer?: (phone: string) => void;
}

type Section = 'cross' | 'within' | 'both';
type CountFilter = 'all' | '2' | '3' | '4' | '5' | 'gt5';

interface DupResponse {
  duplicateGroups: DuplicateGroupItem[];
  totalGroups: number;
  page: number;
  totalPages: number;
  stats: DuplicateFullStats;
}

const SECTIONS: { id: Section; label: string; hint: string }[] = [
  { id: 'cross', label: 'Across files', hint: 'Same phone in more than one workbook or sheet' },
  { id: 'within', label: 'Inside one sheet', hint: 'Same phone repeated in a single sheet' },
  { id: 'both', label: 'Both', hint: 'Every repeated phone' },
];

const COUNTS: { id: CountFilter; label: string; key: 'total' | 'c2' | 'c3' | 'c4' | 'c5' | 'gt5' }[] = [
  { id: 'all', label: 'All', key: 'total' },
  { id: '2', label: '2×', key: 'c2' },
  { id: '3', label: '3×', key: 'c3' },
  { id: '4', label: '4×', key: 'c4' },
  { id: '5', label: '5×', key: 'c5' },
  { id: 'gt5', label: '6+×', key: 'gt5' },
];

export const DuplicatesView: React.FC<DuplicatesViewProps> = ({ onViewRecord, onLookupCustomer }) => {
  const [section, setSection] = useState<Section>('cross');
  const [countFilter, setCountFilter] = useState<CountFilter>('all');
  const [search, setSearch] = useState('');
  const [submitted, setSubmitted] = useState('');
  const [page, setPage] = useState(1);
  const [data, setData] = useState<DupResponse | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [open, setOpen] = useState<Record<string, boolean>>({});
  const [exporting, setExporting] = useState(false);

  const load = useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      const qs = new URLSearchParams({
        section, countFilter, page: String(page), limit: '60', search: submitted,
      });
      const res = await fetch(`/api/duplicates?${qs.toString()}`);
      const json = await res.json().catch(() => null);
      if (!res.ok || !json || !Array.isArray(json.duplicateGroups)) {
        throw new Error((json && json.error) || `Could not load duplicates (HTTP ${res.status})`);
      }
      setData(json);
    } catch (e: any) {
      setData(null);
      setError(e.message || 'Could not load duplicates');
    } finally {
      setLoading(false);
    }
  }, [section, countFilter, page, submitted]);

  useEffect(() => { load(); }, [load]);

  const stats = data?.stats?.[section];

  const handleExport = async () => {
    setExporting(true);
    try {
      await downloadFromPost('/api/duplicates/export', { countFilter, search: submitted }, 'F2_Duplicates.xlsx');
    } catch (e: any) {
      window.alert(e.message || 'Export failed');
    } finally {
      setExporting(false);
    }
  };

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <h2 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
              <Copy className="w-5 h-5 text-blue-600" />
              <span>Duplicate Leads</span>
            </h2>
            <p className="text-xs text-slate-500 mt-1 font-medium">
              Phone numbers that appear more than once in the Drive directory files. MIS sheets are not included.
            </p>
          </div>
          <button
            onClick={handleExport}
            disabled={exporting}
            className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold shadow-xs cursor-pointer disabled:opacity-50"
          >
            {exporting ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Download className="w-3.5 h-3.5" />}
            Export Excel
          </button>
        </div>

        <div className="flex flex-wrap gap-2">
          {SECTIONS.map((s) => (
            <button
              key={s.id}
              title={s.hint}
              onClick={() => { setSection(s.id); setPage(1); }}
              className={`px-4 py-2 rounded-xl text-xs font-bold cursor-pointer border ${
                section === s.id ? 'bg-blue-600 text-white border-blue-600' : 'bg-slate-100 text-slate-700 border-slate-200 hover:bg-slate-200'
              }`}
            >
              {s.label}
              {data?.stats && (
                <span className={`ml-2 font-mono ${section === s.id ? 'text-blue-100' : 'text-slate-400'}`}>
                  {data.stats[s.id].total.toLocaleString()}
                </span>
              )}
            </button>
          ))}
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <form
            onSubmit={(e) => { e.preventDefault(); setPage(1); setSubmitted(search.trim()); }}
            className="relative flex-1 min-w-[240px]"
          >
            <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Filter by phone or name, then press Enter"
              className="w-full pl-10 pr-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-500 focus:bg-white focus:ring-1 focus:ring-blue-500 font-medium"
            />
          </form>
          <div className="flex flex-wrap gap-1.5">
            {COUNTS.map((c) => (
              <button
                key={c.id}
                onClick={() => { setCountFilter(c.id); setPage(1); }}
                className={`px-3 py-2 rounded-xl text-xs font-bold cursor-pointer border ${
                  countFilter === c.id ? 'bg-slate-900 text-white border-slate-900' : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                {c.label}
                {stats && <span className="ml-1.5 font-mono opacity-60">{stats[c.key].toLocaleString()}</span>}
              </button>
            ))}
          </div>
        </div>
      </div>

      {loading && (
        <div className="bg-white border border-slate-200 rounded-2xl p-10 text-center text-slate-500 text-sm">
          <Loader2 className="w-5 h-5 animate-spin mx-auto mb-2 text-blue-600" /> Finding duplicates…
        </div>
      )}

      {error && !loading && (
        <div className="bg-rose-50 border border-rose-200 text-rose-700 rounded-2xl p-5 text-sm flex items-center gap-2">
          <AlertTriangle className="w-4 h-4 shrink-0" /> {error}
        </div>
      )}

      {!loading && !error && data && data.duplicateGroups.length === 0 && (
        <div className="bg-white border border-dashed border-slate-300 rounded-2xl p-10 text-center text-slate-500 text-sm">
          No duplicates for this filter.
        </div>
      )}

      {!loading && !error && data && data.duplicateGroups.length > 0 && (
        <div className="space-y-2.5">
          <div className="text-[11px] text-slate-500 font-mono font-semibold px-1">
            {data.totalGroups.toLocaleString()} groups · page {data.page} of {data.totalPages}
          </div>
          {data.duplicateGroups.map((g) => {
            const isOpen = !!open[g.key];
            const first = g.records[0];
            return (
              <div key={g.key} className="bg-white border border-slate-200/90 rounded-2xl shadow-xs overflow-hidden">
                <button
                  onClick={() => setOpen((o) => ({ ...o, [g.key]: !o[g.key] }))}
                  className="w-full flex items-center gap-3 px-4 py-3 text-left hover:bg-slate-50 cursor-pointer"
                >
                  {isOpen ? <ChevronDown className="w-4 h-4 text-slate-400" /> : <ChevronRight className="w-4 h-4 text-slate-400" />}
                  <span className="font-mono font-bold text-slate-900 text-sm">{g.key}</span>
                  <span className="text-xs text-slate-600 truncate flex-1">{first?.customerName || ''}</span>
                  <span className="px-2 py-0.5 rounded-full bg-rose-50 text-rose-700 border border-rose-200 text-[11px] font-bold">
                    {g.count}× repeated
                  </span>
                  {typeof g.distinctFilesCount === 'number' && (
                    <span className="px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 border border-slate-200 text-[11px] font-bold">
                      {g.distinctFilesCount} file{g.distinctFilesCount === 1 ? '' : 's'}
                    </span>
                  )}
                </button>
                {isOpen && (
                  <div className="border-t border-slate-100 overflow-x-auto">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-50 text-slate-500 uppercase tracking-wider text-[10px] font-bold">
                        <tr>
                          <th className="px-4 py-2">Name</th>
                          <th className="px-4 py-2">Workbook</th>
                          <th className="px-4 py-2">Sheet</th>
                          <th className="px-4 py-2 text-right">Excel row</th>
                          <th className="px-4 py-2 text-right">Actions</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {g.records.map((r) => (
                          <tr key={r.id}>
                            <td className="px-4 py-2 font-semibold text-slate-900">{r.customerName || '—'}</td>
                            <td className="px-4 py-2 text-slate-700">{r.fileName}</td>
                            <td className="px-4 py-2 text-slate-700">{r.sheetName}</td>
                            <td className="px-4 py-2 text-right font-mono">{r.rowNum}</td>
                            <td className="px-4 py-2 text-right whitespace-nowrap">
                              <button
                                onClick={() => onViewRecord(r.id)}
                                className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 font-bold cursor-pointer mr-1.5"
                              >
                                <Maximize2 className="w-3 h-3" /> Row
                              </button>
                              {onLookupCustomer && (
                                <button
                                  onClick={() => onLookupCustomer(r.phone || g.key)}
                                  className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 font-bold cursor-pointer"
                                >
                                  <UserSearch className="w-3 h-3" /> Lookup
                                </button>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            );
          })}

          {data.totalPages > 1 && (
            <div className="flex items-center justify-center gap-2 pt-2">
              <button
                disabled={page <= 1}
                onClick={() => setPage((p) => Math.max(1, p - 1))}
                className="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-xs font-bold border border-slate-200 disabled:opacity-40 cursor-pointer"
              >
                Previous
              </button>
              <span className="text-xs text-slate-500 font-mono">{page} / {data.totalPages}</span>
              <button
                disabled={page >= data.totalPages}
                onClick={() => setPage((p) => p + 1)}
                className="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-xs font-bold border border-slate-200 disabled:opacity-40 cursor-pointer"
              >
                Next
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
