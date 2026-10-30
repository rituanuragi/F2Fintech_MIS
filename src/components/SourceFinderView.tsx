import React, { useState } from 'react';
import { Search, Loader2, Folder, FileSpreadsheet, Eye, Maximize2, AlertCircle } from 'lucide-react';
import { SearchMatchItem } from '../types';

interface Props {
  onViewRecord: (id: number) => void;
  onPreviewSheet: (fileId: string, sheetName: string) => void;
  onLookupCustomer?: (phone: string) => void;
}

const MAX_RESULTS = 20;

/** Find which workbook / sheet / row a phone number or name came from. */
export const SourceFinderView: React.FC<Props> = ({ onViewRecord, onPreviewSheet, onLookupCustomer }) => {
  const [query, setQuery] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [searched, setSearched] = useState(false);
  const [matches, setMatches] = useState<SearchMatchItem[]>([]);
  const [total, setTotal] = useState(0);

  const run = async (e?: React.FormEvent) => {
    e?.preventDefault();
    const q = query.trim();
    if (!q) return;
    setLoading(true);
    setError('');
    setSearched(true);
    try {
      const res = await fetch(`/api/search?type=smart&limit=${MAX_RESULTS}&q=${encodeURIComponent(q)}`);
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data?.error || `Search failed (${res.status})`);
      setMatches(data.matches || []);
      setTotal(data.totalCount ?? (data.matches || []).length);
    } catch (err: any) {
      setMatches([]);
      setTotal(0);
      setError(err?.message || 'Search failed');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-4">
        <div>
          <h2 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
            <Search className="w-5 h-5 text-blue-600" />
            Source Finder
          </h2>
          <p className="text-xs text-slate-500 mt-1 font-medium">
            Type a phone number or name to see which workbook, sheet and row it came from.
          </p>
        </div>
        <form onSubmit={run} className="flex gap-2.5">
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Phone number or name"
            className="flex-1 px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl text-sm text-slate-900 focus:outline-none focus:border-blue-500 focus:bg-white"
          />
          <button
            type="submit"
            disabled={loading || !query.trim()}
            className="px-6 py-3 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white rounded-xl text-sm font-bold flex items-center gap-2 cursor-pointer"
          >
            {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : <Search className="w-4 h-4" />}
            Find source
          </button>
        </form>
      </div>

      {error && (
        <div className="flex items-center gap-2 bg-red-50 border border-red-200 text-red-700 rounded-xl p-4 text-xs font-semibold">
          <AlertCircle className="w-4 h-4" /> {error}
        </div>
      )}

      {searched && !loading && !error && matches.length === 0 && (
        <div className="bg-white border border-slate-200 rounded-2xl p-8 text-center text-sm text-slate-500">
          No source found for “{query.trim()}”.
        </div>
      )}

      {matches.length > 0 && (
        <div className="bg-white border border-slate-200/90 rounded-2xl shadow-xs overflow-hidden">
          <div className="px-5 py-3 border-b border-slate-100 text-xs font-bold text-slate-600">
            {total > matches.length ? `Showing first ${matches.length} of ${total} matches` : `${matches.length} match${matches.length === 1 ? '' : 'es'}`}
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs">
              <thead className="bg-slate-50 text-slate-500 uppercase text-[10px] tracking-wide">
                <tr>
                  <th className="px-4 py-2.5">Folder</th>
                  <th className="px-4 py-2.5">Workbook</th>
                  <th className="px-4 py-2.5">Sheet</th>
                  <th className="px-4 py-2.5">Row</th>
                  <th className="px-4 py-2.5">Matched on</th>
                  <th className="px-4 py-2.5 text-right">Open</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {matches.map((m) => (
                  <tr key={m.id} className="hover:bg-slate-50">
                    <td className="px-4 py-2.5 text-slate-600"><span className="inline-flex items-center gap-1.5"><Folder className="w-3.5 h-3.5 text-slate-400" />{m.folderPath}</span></td>
                    <td className="px-4 py-2.5 font-semibold text-slate-900"><span className="inline-flex items-center gap-1.5"><FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />{m.fileName}</span></td>
                    <td className="px-4 py-2.5 text-slate-700">{m.sheetName}</td>
                    <td className="px-4 py-2.5 font-mono text-slate-700">{m.rowNum}</td>
                    <td className="px-4 py-2.5 text-slate-600">{m.matchedColumn || m.matchType}</td>
                    <td className="px-4 py-2.5">
                      <div className="flex justify-end gap-1.5">
                        <button onClick={() => onPreviewSheet(m.fileId, m.sheetName)} className="p-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 cursor-pointer" title="Preview sheet"><Eye className="w-3.5 h-3.5" /></button>
                        <button onClick={() => onViewRecord(m.id)} className="p-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 cursor-pointer" title="View full row"><Maximize2 className="w-3.5 h-3.5" /></button>
                        {onLookupCustomer && m.normPhone && (
                          <button onClick={() => onLookupCustomer(m.normPhone)} className="px-2 py-1 rounded-lg bg-blue-50 text-blue-700 hover:bg-blue-100 font-bold cursor-pointer">Lookup</button>
                        )}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};
