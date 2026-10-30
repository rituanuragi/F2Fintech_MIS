import React, { useState } from 'react';
import {
  Search,
  PhoneCall,
  User,
  Sparkles,
  Layers,
  FileSpreadsheet,
  Folder,
  Download,
  Eye,
  CheckCircle2,
  Loader2,
  Maximize2,
  Award,
  BookOpen,
} from 'lucide-react';
import { SearchMatchItem } from '../types';
import { BatchExcelSearch } from './BatchExcelSearch';
import { MisSearchBookView } from './MisSearchBookView';

interface GlobalSearchViewProps {
  onViewRecord: (id: number) => void;
  onPreviewSheet: (fileId: string, sheetName: string) => void;
  onLookupCustomer?: (phone: string) => void;
  initialSearchType?: 'phone' | 'name' | 'smart' | 'batch' | 'mis-search-book';
}

export const GlobalSearchView: React.FC<GlobalSearchViewProps> = ({
  onViewRecord,
  onPreviewSheet,
  onLookupCustomer,
  initialSearchType,
}) => {
  const [searchType, setSearchType] = useState<'phone' | 'name' | 'smart' | 'batch' | 'mis-search-book'>(
    initialSearchType || 'phone'
  );
  const [query, setQuery] = useState('');
  const [loading, setLoading] = useState(false);
  const [results, setResults] = useState<{
    matches: SearchMatchItem[];
    totalCount: number;
    searchQuery: string;
    searchType: string;
    durationMs: number;
    sourcesSearched: { filesCount: number; sheetsCount: number; recordsCount: number };
  } | null>(null);
  const [hasSearched, setHasSearched] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  const handleSearch = async (e?: React.FormEvent, customQuery?: string, customType?: 'phone' | 'name' | 'smart') => {
    if (e) e.preventDefault();
    const q = customQuery !== undefined ? customQuery : query;
    const t = customType !== undefined ? customType : (searchType === 'batch' ? 'phone' : searchType);

    if (!q.trim()) return;

    setLoading(true);
    setHasSearched(true);
    try {
      const res = await fetch(`/api/search?q=${encodeURIComponent(q.trim())}&type=${t}&limit=100`);
      const data = await res.json().catch(() => null);
      if (!res.ok || !data || !Array.isArray(data.matches)) {
        throw new Error((data && data.error) || `Search failed (HTTP ${res.status})`);
      }
      setResults(data);
    } catch (err) {
      console.error('Search error:', err);
      setResults(null);
      window.alert(err instanceof Error ? `Search failed: ${err.message}` : 'Search failed.');
    } finally {
      setLoading(false);
    }
  };

  const handleExport = async (format: 'csv' | 'xlsx') => {
    if (!results || results.matches.length === 0) return;
    setIsExporting(true);
    try {
      const res = await fetch('/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ matches: results.matches, format }),
      });
      if (!res.ok) {
        const errBody = await res.json().catch(() => null);
        throw new Error((errBody && errBody.error) || `Export failed (HTTP ${res.status})`);
      }
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `F2_Fintech_Search_${results.searchQuery.replace(/\s+/g, '_')}.${format}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (err) {
      console.error('Export error:', err);
      window.alert(err instanceof Error ? `Export failed: ${err.message}` : 'Export failed.');
    } finally {
      setIsExporting(false);
    }
  };

  const quickSamples = [
    { label: '8096452615 (ICICI)', q: '8096452615', type: 'phone' as const },
    { label: '8657184019 (LOGIN)', q: '8657184019', type: 'phone' as const },
    { label: 'SUMIT GARG (Multi-match)', q: 'SUMIT GARG', type: 'name' as const },
    { label: 'Y REDDY', q: 'Y REDDY', type: 'name' as const },
    { label: 'BAJAJ (Lender Smart)', q: 'BAJAJ', type: 'smart' as const },
    { label: 'Delhi (City Smart)', q: 'Delhi', type: 'smart' as const },
  ];

  return (
    <div className="space-y-6">
      {/* Top Search Controls */}
      <div className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-4">
        <div>
          <h2 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
            <Search className="w-5 h-5 text-blue-600" />
            <span>Global Multi-Workbook Search</span>
          </h2>
          <p className="text-xs text-slate-500 mt-1 font-medium">
            Search instantaneously across all indexed records, all folders, workbooks, and worksheets.
          </p>
        </div>

        {/* Search Mode Tabs */}
        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => setSearchType('phone')}
            className={`inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              searchType === 'phone'
                ? 'bg-blue-600 text-white shadow-xs'
                : 'bg-slate-100 text-slate-700 hover:bg-slate-200 border border-slate-200/80'
            }`}
          >
            <PhoneCall className="w-3.5 h-3.5" />
            <span>Search Phone Number</span>
          </button>

          <button
            type="button"
            onClick={() => setSearchType('name')}
            className={`inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              searchType === 'name'
                ? 'bg-blue-600 text-white shadow-xs'
                : 'bg-slate-100 text-slate-700 hover:bg-slate-200 border border-slate-200/80'
            }`}
          >
            <User className="w-3.5 h-3.5" />
            <span>Search Name</span>
          </button>

          <button
            type="button"
            onClick={() => setSearchType('smart')}
            className={`inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              searchType === 'smart'
                ? 'bg-blue-600 text-white shadow-xs'
                : 'bg-slate-100 text-slate-700 hover:bg-slate-200 border border-slate-200/80'
            }`}
          >
            <Sparkles className="w-3.5 h-3.5 text-amber-300" />
            <span>Smart Search (City, State, PAN, Lender)</span>
          </button>

          <button
            type="button"
            onClick={() => setSearchType('batch')}
            className={`inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              searchType === 'batch'
                ? 'bg-emerald-600 text-white shadow-xs'
                : 'bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200'
            }`}
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
            <span>Upload Excel &amp; Match Sheets</span>
            <span className="px-1.5 py-0.2 rounded bg-emerald-200 text-emerald-800 text-[10px] uppercase font-mono font-bold">
              Batch
            </span>
          </button>

          <button
            type="button"
            onClick={() => setSearchType('mis-search-book')}
            className={`inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              searchType === 'mis-search-book'
                ? 'bg-indigo-600 text-white shadow-xs ring-2 ring-indigo-500/20'
                : 'bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border border-indigo-200'
            }`}
          >
            <BookOpen className="w-3.5 h-3.5 text-indigo-500" />
            <span>MIS Search Book</span>
            <span className="px-1.5 py-0.2 rounded bg-indigo-200 text-indigo-900 text-[10px] uppercase font-mono font-bold">
              MIS Sync
            </span>
          </button>
        </div>

        {/* Regular Search Form (Only shown if searchType !== 'batch' && searchType !== 'mis-search-book') */}
        {searchType !== 'batch' && searchType !== 'mis-search-book' && (
          <>
            <form onSubmit={handleSearch} className="flex gap-2.5">
              <div className="relative flex-1">
                <Search className="w-5 h-5 text-slate-400 absolute left-4 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                  placeholder={
                    searchType === 'phone'
                      ? 'Enter phone number (e.g. 9876543210, +91 98765-43210)...'
                      : searchType === 'name'
                      ? 'Enter customer or applicant name (e.g. Rahul Sharma, Sumit Garg)...'
                      : 'Enter any keyword, PAN, city, loan type, lender or application ID...'
                  }
                  className="w-full pl-12 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-xl text-sm text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-500 focus:bg-white focus:ring-1 focus:ring-blue-500 font-medium transition-all"
                />
              </div>

              <button
                type="submit"
                disabled={loading || !query.trim()}
                className="px-6 py-3 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white rounded-xl text-sm font-bold shadow-sm shadow-blue-500/20 flex items-center gap-2 transition-all cursor-pointer shrink-0"
              >
                {loading ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    <span>Searching...</span>
                  </>
                ) : (
                  <>
                    <Search className="w-4 h-4" />
                    <span>Search</span>
                  </>
                )}
              </button>
            </form>

            {/* Quick sample chips */}
            <div className="flex flex-wrap items-center gap-2 pt-1 text-xs">
              <span className="text-slate-400 font-mono text-[11px] font-semibold">Quick test queries:</span>
              {quickSamples.map((s, idx) => (
                <button
                  key={idx}
                  type="button"
                  onClick={() => {
                    setQuery(s.q);
                    setSearchType(s.type);
                    handleSearch(undefined, s.q, s.type);
                  }}
                  className="px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 hover:text-slate-900 border border-slate-200 text-[11px] font-mono transition-colors cursor-pointer"
                >
                  {s.label}
                </button>
              ))}
            </div>
          </>
        )}
      </div>

      {/* Render MIS Search Book or Batch Excel Search if selected */}
      {searchType === 'mis-search-book' ? (
        <MisSearchBookView
          onViewRecord={onViewRecord}
          onPreviewSheet={onPreviewSheet}
          onLookupCustomer={onLookupCustomer}
        />
      ) : searchType === 'batch' ? (
        <BatchExcelSearch
          onViewRecord={onViewRecord}
          onPreviewSheet={onPreviewSheet}
          onLookupCustomer={onLookupCustomer}
        />
      ) : (
        /* Regular Results Section */
        hasSearched && results && (
          <div className="space-y-4">
            {/* Results Summary Bar */}
            <div className="bg-white border border-slate-200/90 rounded-2xl p-4 flex flex-wrap items-center justify-between gap-3 shadow-xs">
              <div>
                {results.matches.length > 0 ? (
                  <div className="flex items-center gap-2">
                    <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 text-xs font-bold uppercase tracking-wider">
                      <CheckCircle2 className="w-3.5 h-3.5" /> Match Found
                    </span>
                    <span className="text-sm font-bold text-slate-900 font-mono">
                      {results.totalCount.toLocaleString()} {results.totalCount === 1 ? 'match' : 'matches'} found
                    </span>
                    <span className="text-xs text-slate-400 font-mono">
                      ({results.durationMs}ms)
                    </span>
                  </div>
                ) : (
                  <div className="text-sm font-bold text-rose-600 flex items-center gap-2">
                    <span>NO MATCH FOUND</span>
                    <span className="text-xs text-slate-500 font-normal">
                      Search value "{results.searchQuery}" not found in indexed records.
                    </span>
                  </div>
                )}
                <div className="text-[11px] text-slate-500 font-medium mt-1">
                  Shows locations only (workbook, sheet and row) for rows in the main directory files. MIS and Google Sheets rows are not included.
                </div>
                <div className="text-[11px] text-slate-400 font-mono mt-1 font-medium">
                  Searched: {results.sourcesSearched.filesCount} files • {results.sourcesSearched.sheetsCount} worksheets • {results.sourcesSearched.recordsCount.toLocaleString()} indexed records
                </div>
              </div>

              {results.matches.length > 0 && (
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => handleExport('csv')}
                    disabled={isExporting}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold border border-slate-200 transition-colors cursor-pointer disabled:opacity-50"
                  >
                    <Download className="w-3.5 h-3.5 text-slate-500" />
                    <span>Export CSV</span>
                  </button>
                  <button
                    onClick={() => handleExport('xlsx')}
                    disabled={isExporting}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold transition-colors cursor-pointer shadow-xs disabled:opacity-50"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Export XLSX</span>
                  </button>
                </div>
              )}
            </div>

            {/* Results List */}
            {results.matches.length > 0 ? (
              <div className="space-y-3">
                {results.matches.map((match, idx) => (
                  <div
                    key={match.id}
                    className="bg-white border border-slate-200/90 hover:border-blue-400 rounded-2xl p-5 shadow-xs transition-all group"
                  >
                    <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                      {/* Left: Customer Info & Match Details */}
                      <div className="space-y-2 flex-1">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="font-mono text-xs text-slate-400 font-bold">
                            #{idx + 1}
                          </span>
                          <span
                            className={`px-2.5 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider ${
                              match.matchType.includes('Exact')
                                ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                                : 'bg-blue-50 text-blue-700 border border-blue-200'
                            }`}
                          >
                            {match.matchType}
                          </span>

                          {match.priority && match.priority !== '—' && (
                            <span
                              className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-bold font-mono border ${match.priorityBadge?.bg || 'bg-amber-50'} ${match.priorityBadge?.text || 'text-amber-700'} ${match.priorityBadge?.border || 'border-amber-200'}`}
                              title={`Lead Priority: ${match.priority} (Rank #${match.priorityRank || '—'})`}
                            >
                              <Award className="w-3 h-3 text-amber-500" />
                              <span>Priority: {match.priority}</span>
                            </span>
                          )}

                          {match.matchedColumn && (
                            <span className="text-[11px] text-slate-500 font-mono">
                              Matched on <span className="text-slate-800 font-bold">{match.matchedColumn}</span>
                            </span>
                          )}
                        </div>

                        <div className="text-sm font-bold text-slate-900 tracking-tight">
                          Found in {match.fileName}
                          <span className="text-slate-400 font-normal"> → </span>
                          {match.sheetName}
                          <span className="text-slate-400 font-normal"> → </span>
                          Row {match.rowNum}
                        </div>
                      </div>

                      {/* Right: Exact Source Breadcrumbs & Actions */}
                      <div className="flex flex-col sm:flex-row items-start sm:items-center gap-4 lg:border-l lg:border-slate-100 lg:pl-6 shrink-0">
                        {/* Source Breadcrumbs Box */}
                        <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs space-y-1 font-mono text-slate-700 min-w-[240px]">
                          <div className="text-[10px] uppercase tracking-wider text-slate-400 font-bold mb-1 font-sans">
                            Source Hierarchy
                          </div>
                          <div className="flex items-center gap-1.5 text-slate-600 truncate">
                            <Folder className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                            <span className="truncate">{match.folderPath}</span>
                          </div>
                          <div className="flex items-center gap-1.5 text-slate-900 font-bold truncate">
                            <FileSpreadsheet className="w-3.5 h-3.5 text-blue-600 shrink-0" />
                            <span className="truncate">{match.fileName}</span>
                          </div>
                          <div className="flex items-center gap-1.5 text-purple-700 font-bold truncate">
                            <Layers className="w-3.5 h-3.5 text-purple-600 shrink-0" />
                            <span>{match.sheetName}</span>
                            <span className="text-slate-400 font-normal">→</span>
                            <span className="bg-blue-100 text-blue-800 px-1.5 py-0.2 rounded text-[11px] font-bold">
                              Row {match.rowNum}
                            </span>
                          </div>
                        </div>

                        {/* Action Buttons */}
                        <div className="flex flex-col gap-2 w-full sm:w-auto">
                          <button
                            onClick={() => onViewRecord(match.id)}
                            className="inline-flex items-center justify-center gap-1.5 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold shadow-xs transition-all cursor-pointer"
                          >
                            <Maximize2 className="w-3.5 h-3.5" />
                            <span>View Full Row</span>
                          </button>

                          <button
                            onClick={() => onPreviewSheet(match.fileId, match.sheetName)}
                            className="inline-flex items-center justify-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold border border-slate-200 transition-colors cursor-pointer"
                          >
                            <Eye className="w-3.5 h-3.5 text-slate-500" />
                            <span>Preview Sheet</span>
                          </button>

                          {onLookupCustomer && (match.normPhone || match.phoneRaw || match.customerName) && (
                            <button
                              onClick={() => onLookupCustomer(match.normPhone || match.phoneRaw || match.customerName)}
                              className="inline-flex items-center justify-center gap-1.5 px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-xl text-xs font-bold border border-indigo-200 transition-colors cursor-pointer"
                            >
                              <User className="w-3.5 h-3.5 text-indigo-600" />
                              <span>Lookup &amp; Diff</span>
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              <div className="bg-white border border-slate-200/90 rounded-2xl p-12 text-center text-slate-400 space-y-3 shadow-xs">
                <Search className="w-10 h-10 text-slate-300 mx-auto" />
                <h3 className="text-base font-bold text-slate-800">No records found matching "{results.searchQuery}"</h3>
                <p className="text-xs text-slate-500 max-w-md mx-auto">
                  No matching cells were detected across all indexed rows. Try a different spelling, phone number format, or use Smart Search.
                </p>
              </div>
            )}
          </div>
        )
      )}
    </div>
  );
};
