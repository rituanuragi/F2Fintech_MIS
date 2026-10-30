import React, { useState, useEffect } from 'react';
import {
  X,
  FileSpreadsheet,
  Search,
  ChevronLeft,
  ChevronRight,
  Maximize2,
  Table,
  Filter,
  Loader2,
  Hash,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  RotateCcw,
  Check,
  Layers,
  Sparkles,
  Eye,
  SlidersHorizontal,
  Award,
  Download,
  Sliders,
} from 'lucide-react';
import { LeadPriorityRule, SheetPreviewRowItem } from '../types';
import { PriorityConfigModal } from './PriorityConfigModal';

interface SheetPreviewModalProps {
  fileId: string | null;
  sheetName: string | null;
  onClose: () => void;
  onViewRecord?: (id: number) => void;
  onSwitchSheet?: (fileId: string, sheetName: string) => void;
}

type SearchMode = 'all' | 'column' | 'row';

export const SheetPreviewModal: React.FC<SheetPreviewModalProps> = ({
  fileId,
  sheetName,
  onClose,
  onViewRecord,
  onSwitchSheet,
}) => {
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [data, setData] = useState<{
    fileId: string;
    fileName: string;
    sheetName: string;
    headers: string[];
    totalRows: number;
    workbookSheets?: { name: string; rowsCount: number }[];
    rows: SheetPreviewRowItem[];
  } | null>(null);

  // Priority state
  const [isPriorityModalOpen, setIsPriorityModalOpen] = useState(false);
  const [priorityFilter, setPriorityFilter] = useState<string>('all');
  const [isExporting, setIsExporting] = useState(false);
  const [activeRules, setActiveRules] = useState<LeadPriorityRule[]>([]);

  // Load priority rules
  useEffect(() => {
    fetch('/api/priority-rules')
      .then(res => res.json())
      .then(d => {
        if (d.rules) setActiveRules(d.rules);
      })
      .catch(console.error);
  }, []);

  // Pagination
  const [limit, setLimit] = useState(50);
  const [page, setPage] = useState(0);

  // Search modes: 'all' (across columns), 'column' (specific column filter), 'row' (search only the row)
  const [activeMode, setActiveMode] = useState<SearchMode>('all');

  // Search by all columns
  const [generalSearch, setGeneralSearch] = useState('');

  // Search only by Row Number (Requested Feature)
  const [rowSearchInput, setRowSearchInput] = useState('');
  const [activeRowSearch, setActiveRowSearch] = useState<number | null>(null);
  const [isRowRange, setIsRowRange] = useState(false);
  const [rowFromInput, setRowFromInput] = useState('');
  const [rowToInput, setRowToInput] = useState('');
  const [activeRowRange, setActiveRowRange] = useState<{ from: number | null; to: number | null }>({ from: null, to: null });

  // Column-Specific Filter (Requested Feature: Apply and search filter)
  const [selectedColumn, setSelectedColumn] = useState<string>('');
  const [columnOperator, setColumnOperator] = useState<string>('contains');
  const [columnFilterValue, setColumnFilterValue] = useState<string>('');
  const [activeColumnFilter, setActiveColumnFilter] = useState<{
    column: string;
    operator: string;
    value: string;
  } | null>(null);

  // Sorting
  const [sortBy, setSortBy] = useState<string>('row_num');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('asc');

  // Fetch preview data from backend
  const fetchPreview = async (
    targetPage: number,
    opts?: {
      search?: string;
      rowNum?: number | null;
      rowFrom?: number | null;
      rowTo?: number | null;
      colFilter?: { column: string; operator: string; value: string } | null;
      sort?: string;
      order?: 'asc' | 'desc';
    }
  ) => {
    if (!fileId || !sheetName) return;
    setLoading(true);

    try {
      const offset = targetPage * limit;
      const s = opts?.search !== undefined ? opts.search : generalSearch;
      const rNum = opts?.rowNum !== undefined ? opts.rowNum : activeRowSearch;
      const rFrom = opts?.rowFrom !== undefined ? opts.rowFrom : activeRowRange.from;
      const rTo = opts?.rowTo !== undefined ? opts.rowTo : activeRowRange.to;
      const colF = opts?.colFilter !== undefined ? opts.colFilter : activeColumnFilter;
      const sBy = opts?.sort !== undefined ? opts.sort : sortBy;
      const sOrder = opts?.order !== undefined ? opts.order : sortOrder;

      const params = new URLSearchParams({
        fileId,
        sheetName,
        limit: String(limit),
        offset: String(offset),
      });

      if (s.trim()) params.append('search', s.trim());
      if (rNum !== null && !isNaN(rNum)) params.append('rowNumber', String(rNum));
      if (rFrom !== null && !isNaN(rFrom)) params.append('rowFrom', String(rFrom));
      if (rTo !== null && !isNaN(rTo)) params.append('rowTo', String(rTo));
      if (colF && colF.column && colF.value.trim()) {
        params.append('columnName', colF.column);
        params.append('columnOperator', colF.operator);
        params.append('columnValue', colF.value.trim());
      }
      if (sBy) params.append('sortBy', sBy);
      if (sOrder) params.append('sortOrder', sOrder);

      setLoadError(null);
      const res = await fetch(`/api/preview?${params.toString()}`);
      const resJson = await res.json().catch(() => null);
      if (!res.ok || !resJson || !Array.isArray(resJson.rows) || typeof resJson.totalRows !== 'number') {
        throw new Error((resJson && resJson.error) || `Preview failed (HTTP ${res.status})`);
      }
      setData(resJson);

      // Auto-select first column if none selected yet
      if (resJson.headers && resJson.headers.length > 0 && !selectedColumn) {
        setSelectedColumn(resJson.headers[0]);
      }
    } catch (err) {
      console.error('Error fetching sheet preview:', err);
      setData(null);
      setLoadError(err instanceof Error ? err.message : 'Could not load this sheet.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (fileId && sheetName) {
      setPage(0);
      fetchPreview(0);
    }
  }, [fileId, sheetName, limit]);

  // Handle General Search
  const handleGeneralSearchSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setPage(0);
    setActiveRowSearch(null);
    setActiveRowRange({ from: null, to: null });
    setActiveColumnFilter(null);
    fetchPreview(0, {
      search: generalSearch,
      rowNum: null,
      rowFrom: null,
      rowTo: null,
      colFilter: null,
    });
  };

  // Handle Search Only by Row Number (Requested Feature)
  const handleRowSearchSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setPage(0);
    setGeneralSearch('');
    setActiveColumnFilter(null);

    if (isRowRange) {
      const from = rowFromInput.trim() ? parseInt(rowFromInput.trim(), 10) : null;
      const to = rowToInput.trim() ? parseInt(rowToInput.trim(), 10) : null;
      setActiveRowRange({ from, to });
      setActiveRowSearch(null);
      fetchPreview(0, {
        search: '',
        rowNum: null,
        rowFrom: from,
        rowTo: to,
        colFilter: null,
      });
    } else {
      const rowNum = rowSearchInput.trim() ? parseInt(rowSearchInput.trim(), 10) : null;
      setActiveRowSearch(rowNum);
      setActiveRowRange({ from: null, to: null });
      fetchPreview(0, {
        search: '',
        rowNum,
        rowFrom: null,
        rowTo: null,
        colFilter: null,
      });
    }
  };

  // Handle Apply Column Filter (Requested Feature)
  const handleApplyColumnFilter = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedColumn || !columnFilterValue.trim()) return;

    setPage(0);
    setActiveRowSearch(null);
    setActiveRowRange({ from: null, to: null });

    const newFilter = {
      column: selectedColumn,
      operator: columnOperator,
      value: columnFilterValue.trim(),
    };
    setActiveColumnFilter(newFilter);

    fetchPreview(0, {
      rowNum: null,
      rowFrom: null,
      rowTo: null,
      colFilter: newFilter,
    });
  };

  // Clear all filters and resets view
  const clearAllFilters = () => {
    setGeneralSearch('');
    setRowSearchInput('');
    setRowFromInput('');
    setRowToInput('');
    setActiveRowSearch(null);
    setActiveRowRange({ from: null, to: null });
    setColumnFilterValue('');
    setActiveColumnFilter(null);
    setPage(0);
    fetchPreview(0, {
      search: '',
      rowNum: null,
      rowFrom: null,
      rowTo: null,
      colFilter: null,
    });
  };

  const hasAnyFilterActive =
    !!generalSearch.trim() ||
    activeRowSearch !== null ||
    activeRowRange.from !== null ||
    activeRowRange.to !== null ||
    activeColumnFilter !== null ||
    priorityFilter !== 'all';

  // Export worksheet with Priority column included
  const handleExportSheet = async (format: 'xlsx' | 'csv') => {
    if (!fileId || !sheetName) return;
    setIsExporting(true);
    try {
      const res = await fetch('/api/export-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileId, sheetName, format }),
      });
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${sheetName}_with_Priority.${format}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (err) {
      console.error('Export sheet error:', err);
    } finally {
      setIsExporting(false);
    }
  };

  if (!fileId || !sheetName) return null;

  const totalPages = data ? Math.ceil(data.totalRows / limit) : 0;

  // Filter rows by priority if chosen
  const displayedRows = (data?.rows || []).filter(row => {
    if (priorityFilter === 'all') return true;
    return (
      row.priorityCategory === priorityFilter ||
      row.priority === priorityFilter ||
      row.priority?.toLowerCase().includes(priorityFilter.toLowerCase())
    );
  });

  // Sheet columns excluding Priority if already present to avoid duplication
  const regularHeaders = (data?.headers || []).filter(
    h => h.trim().toLowerCase() !== 'priority'
  );

  return (
    <div className="fixed inset-0 z-50 bg-black/85 backdrop-blur-sm flex items-center justify-center p-2 sm:p-5">
      <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-7xl h-[94vh] flex flex-col shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        
        {/* Top Header */}
        <div className="p-4 sm:p-5 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 bg-slate-950/90">
          <div>
            <div className="flex items-center gap-2 text-xs font-mono text-purple-400 font-bold uppercase tracking-wider mb-1">
              <Table className="w-4 h-4" />
              <span>Spreadsheet Grid Explorer</span>
            </div>
            <h2 className="text-base sm:text-lg font-bold text-white flex items-center gap-2 truncate">
              <FileSpreadsheet className="w-5 h-5 text-blue-400 shrink-0" />
              <span className="truncate max-w-xs">{data?.fileName || fileId}</span>
              <span className="text-slate-500 font-normal">→</span>
              <span className="text-emerald-400 font-mono">{sheetName}</span>
            </h2>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            {/* Priority Rules Modal Button */}
            <button
              type="button"
              onClick={() => setIsPriorityModalOpen(true)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-amber-500/10 hover:bg-amber-500/20 text-amber-300 border border-amber-500/30 text-xs font-bold transition-all shadow-sm cursor-pointer"
              title="Change priority rank of all files in directory"
            >
              <Sliders className="w-3.5 h-3.5 text-amber-400" />
              <span>Lead Priority Rules</span>
              <span className="hidden md:inline text-[10px] px-1.5 py-0.5 rounded bg-amber-500/20 text-amber-200 font-mono">
                #1 F2 &gt; #2 NEET...
              </span>
            </button>

            {/* Export Sheet with Priority */}
            <div className="flex items-center gap-1 bg-slate-900 border border-slate-800 p-0.5 rounded-lg">
              <button
                type="button"
                onClick={() => handleExportSheet('xlsx')}
                disabled={isExporting}
                className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 text-white text-xs font-semibold shadow-sm transition-colors cursor-pointer"
                title="Download worksheet with Priority column included"
              >
                <Download className="w-3 h-3" />
                <span>Export XLSX</span>
              </button>

              <button
                type="button"
                onClick={() => handleExportSheet('csv')}
                disabled={isExporting}
                className="inline-flex items-center gap-1 px-2 py-1 rounded hover:bg-slate-800 text-slate-300 disabled:opacity-50 text-xs font-mono transition-colors cursor-pointer"
                title="Download CSV with Priority column included"
              >
                <span>CSV</span>
              </button>
            </div>

            <span className="text-xs font-mono text-slate-300 bg-slate-800/90 px-3 py-1.5 rounded-lg border border-slate-700">
              {data ? data.totalRows.toLocaleString() : 0} rows {hasAnyFilterActive ? 'matched' : 'total'}
            </span>

            <button
              onClick={onClose}
              className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"
              title="Close Preview"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Workbook Sheet Tabs Bar (if multiple sheets exist in this workbook) */}
        {data?.workbookSheets && data.workbookSheets.length > 1 && (
          <div className="px-5 py-2 border-b border-slate-800 bg-slate-950/70 flex items-center gap-2 overflow-x-auto">
            <span className="text-[10px] font-mono uppercase tracking-wider text-slate-500 font-bold shrink-0">
              Workbook Sheets ({data.workbookSheets.length}):
            </span>
            {data.workbookSheets.map((ws) => {
              const isCurrent = ws.name === (data.sheetName || sheetName);
              return (
                <button
                  key={ws.name}
                  type="button"
                  onClick={() => {
                    if (!isCurrent && onSwitchSheet) {
                      onSwitchSheet(data.fileId || fileId, ws.name);
                    }
                  }}
                  className={`px-2.5 py-1 rounded-lg text-xs font-mono flex items-center gap-1.5 shrink-0 transition-all cursor-pointer ${
                    isCurrent
                      ? 'bg-emerald-600 text-white font-bold shadow-xs'
                      : 'bg-slate-900 hover:bg-slate-800 text-slate-300 border border-slate-800'
                  }`}
                >
                  <Layers className="w-3 h-3" />
                  <span>{ws.name}</span>
                  <span className={`text-[10px] ${isCurrent ? 'text-emerald-100' : 'text-slate-500'}`}>
                    ({ws.rowsCount.toLocaleString()})
                  </span>
                </button>
              );
            })}
          </div>
        )}

        {/* Mode Selector & Toolbar */}
        <div className="px-5 py-3 border-b border-slate-800 bg-slate-900/90 space-y-3">
          {/* Search Mode Toggles & Priority Filter */}
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2">
              <div className="flex items-center gap-1.5 bg-slate-950 border border-slate-800 p-1 rounded-xl">
                <button
                  type="button"
                  onClick={() => setActiveMode('all')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all cursor-pointer ${
                    activeMode === 'all'
                      ? 'bg-emerald-600 text-white shadow-sm'
                      : 'text-slate-400 hover:text-slate-200'
                  }`}
                >
                  <Search className="w-3.5 h-3.5" />
                  <span>Search All Columns</span>
                </button>

                <button
                  type="button"
                  onClick={() => setActiveMode('column')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all cursor-pointer ${
                    activeMode === 'column'
                      ? 'bg-emerald-600 text-white shadow-sm'
                      : 'text-slate-400 hover:text-slate-200'
                  }`}
                >
                  <SlidersHorizontal className="w-3.5 h-3.5" />
                  <span>Apply Column Filter</span>
                  {activeColumnFilter && (
                    <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
                  )}
                </button>

                <button
                  type="button"
                  onClick={() => setActiveMode('row')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all cursor-pointer ${
                    activeMode === 'row'
                      ? 'bg-gradient-to-r from-purple-600 to-indigo-600 text-white shadow-sm ring-1 ring-purple-400'
                      : 'text-slate-400 hover:text-purple-300'
                  }`}
                >
                  <Hash className="w-3.5 h-3.5 text-purple-400" />
                  <span>Search Only Row #</span>
                  {(activeRowSearch !== null || activeRowRange.from !== null) && (
                    <span className="w-2 h-2 rounded-full bg-purple-400 animate-pulse"></span>
                  )}
                </button>
              </div>

              {/* Priority Filter Dropdown */}
              <div className="flex items-center gap-1.5 bg-slate-950 border border-slate-800 px-3 py-1.5 rounded-xl">
                <Award className="w-3.5 h-3.5 text-amber-400" />
                <span className="text-xs text-amber-400 font-mono font-bold">Priority:</span>
                <select
                  value={priorityFilter}
                  onChange={(e) => setPriorityFilter(e.target.value)}
                  className="bg-transparent text-xs text-slate-200 font-mono focus:outline-none cursor-pointer"
                >
                  <option value="all" className="bg-slate-900 text-slate-200">
                    All Leads ({data?.rows.length || 0})
                  </option>
                  {activeRules.map(r => (
                    <option key={r.id} value={r.name} className="bg-slate-900 text-slate-200">
                      #{r.rank} {r.name}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Pagination Size Selector */}
            <div className="flex items-center gap-2 text-xs font-mono text-slate-400">
              <span>Show per page:</span>
              <select
                value={limit}
                onChange={(e) => {
                  setLimit(Number(e.target.value));
                  setPage(0);
                }}
                className="bg-slate-950 border border-slate-800 rounded px-2 py-1 text-slate-200 text-xs focus:outline-none"
              >
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
                <option value={250}>250</option>
                <option value={500}>500</option>
              </select>
            </div>
          </div>

          {/* MODE 1: Search Across All Columns */}
          {activeMode === 'all' && (
            <form onSubmit={handleGeneralSearchSubmit} className="flex flex-wrap items-center gap-2">
              <div className="relative flex-1 max-w-lg">
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={generalSearch}
                  onChange={(e) => setGeneralSearch(e.target.value)}
                  placeholder="Search this sheet across all columns (e.g. name, state, phone, PAN)..."
                  className="w-full pl-8 pr-3 py-2 bg-slate-950 border border-slate-700/80 rounded-xl text-xs text-slate-100 placeholder:text-slate-500 focus:outline-none focus:border-emerald-500 font-sans"
                />
              </div>

              <button
                type="submit"
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-xl shadow-sm transition-all cursor-pointer flex items-center gap-1.5"
              >
                <Search className="w-3.5 h-3.5" />
                <span>Search</span>
              </button>

              {generalSearch && (
                <button
                  type="button"
                  onClick={() => {
                    setGeneralSearch('');
                    fetchPreview(0, { search: '' });
                  }}
                  className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-xl border border-slate-700 transition-colors cursor-pointer"
                >
                  Clear
                </button>
              )}
            </form>
          )}

          {/* MODE 2: Apply and Search Filter by Specific Column (Requested Feature) */}
          {activeMode === 'column' && (
            <form onSubmit={handleApplyColumnFilter} className="flex flex-wrap items-center gap-2">
              {/* Column dropdown */}
              <div className="flex items-center gap-1.5 bg-slate-950 border border-slate-700/80 rounded-xl px-2.5 py-1">
                <span className="text-[11px] text-slate-500 font-mono">Column:</span>
                <select
                  value={selectedColumn}
                  onChange={(e) => setSelectedColumn(e.target.value)}
                  className="bg-transparent text-slate-200 text-xs font-semibold focus:outline-none py-1 cursor-pointer"
                >
                  {data?.headers.map((h, idx) => (
                    <option key={idx} value={h} className="bg-slate-900 text-white">
                      {h}
                    </option>
                  ))}
                </select>
              </div>

              {/* Operator dropdown */}
              <div className="flex items-center gap-1.5 bg-slate-950 border border-slate-700/80 rounded-xl px-2.5 py-1">
                <span className="text-[11px] text-slate-500 font-mono">Condition:</span>
                <select
                  value={columnOperator}
                  onChange={(e) => setColumnOperator(e.target.value)}
                  className="bg-transparent text-slate-200 text-xs font-semibold focus:outline-none py-1 cursor-pointer"
                >
                  <option value="contains" className="bg-slate-900 text-white">Contains</option>
                  <option value="equals" className="bg-slate-900 text-white">Equals (Exact)</option>
                  <option value="startsWith" className="bg-slate-900 text-white">Starts With</option>
                  <option value="endsWith" className="bg-slate-900 text-white">Ends With</option>
                  <option value="notEquals" className="bg-slate-900 text-white">Does Not Equal</option>
                </select>
              </div>

              {/* Value Input */}
              <div className="relative flex-1 max-w-xs">
                <input
                  type="text"
                  value={columnFilterValue}
                  onChange={(e) => setColumnFilterValue(e.target.value)}
                  placeholder={`Value for ${selectedColumn || 'column'}...`}
                  className="w-full px-3 py-2 bg-slate-950 border border-slate-700/80 rounded-xl text-xs text-slate-100 placeholder:text-slate-500 focus:outline-none focus:border-emerald-500 font-sans"
                />
              </div>

              {/* Apply Filter Button */}
              <button
                type="submit"
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-xl shadow-md shadow-emerald-900/30 transition-all cursor-pointer flex items-center gap-1.5"
              >
                <Filter className="w-3.5 h-3.5" />
                <span>Apply Filter</span>
              </button>

              {activeColumnFilter && (
                <button
                  type="button"
                  onClick={() => {
                    setActiveColumnFilter(null);
                    setColumnFilterValue('');
                    fetchPreview(0, { colFilter: null });
                  }}
                  className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-xl border border-slate-700 transition-colors cursor-pointer"
                >
                  Reset Filter
                </button>
              )}
            </form>
          )}

          {/* MODE 3: Search Only the Row (Requested Feature) */}
          {activeMode === 'row' && (
            <form onSubmit={handleRowSearchSubmit} className="flex flex-wrap items-center gap-3">
              {!isRowRange ? (
                // Exact Row Search
                <div className="flex items-center gap-2">
                  <div className="relative flex items-center">
                    <span className="absolute left-3 text-purple-400 font-mono text-xs font-bold"># Row</span>
                    <input
                      type="number"
                      min={1}
                      value={rowSearchInput}
                      onChange={(e) => setRowSearchInput(e.target.value)}
                      placeholder="e.g. 2, 57046, 305377..."
                      className="pl-16 pr-3 py-2 bg-slate-950 border border-purple-500/50 rounded-xl text-xs text-white placeholder:text-slate-500 focus:outline-none focus:border-purple-400 font-mono w-56 shadow-inner"
                    />
                  </div>

                  <button
                    type="submit"
                    className="px-4 py-2 bg-purple-600 hover:bg-purple-500 text-white text-xs font-bold rounded-xl shadow-md shadow-purple-900/30 transition-all cursor-pointer flex items-center gap-1.5"
                  >
                    <Search className="w-3.5 h-3.5" />
                    <span>Search Exact Row</span>
                  </button>
                </div>
              ) : (
                // Row Range Search
                <div className="flex items-center gap-2">
                  <div className="flex items-center gap-1.5 bg-slate-950 border border-purple-500/40 rounded-xl px-3 py-1">
                    <span className="text-[11px] text-purple-400 font-mono font-bold">From Row:</span>
                    <input
                      type="number"
                      min={1}
                      value={rowFromInput}
                      onChange={(e) => setRowFromInput(e.target.value)}
                      placeholder="1"
                      className="w-20 bg-transparent text-white text-xs font-mono focus:outline-none py-1"
                    />
                  </div>

                  <div className="flex items-center gap-1.5 bg-slate-950 border border-purple-500/40 rounded-xl px-3 py-1">
                    <span className="text-[11px] text-purple-400 font-mono font-bold">To Row:</span>
                    <input
                      type="number"
                      min={1}
                      value={rowToInput}
                      onChange={(e) => setRowToInput(e.target.value)}
                      placeholder="100"
                      className="w-20 bg-transparent text-white text-xs font-mono focus:outline-none py-1"
                    />
                  </div>

                  <button
                    type="submit"
                    className="px-4 py-2 bg-purple-600 hover:bg-purple-500 text-white text-xs font-bold rounded-xl shadow-md shadow-purple-900/30 transition-all cursor-pointer flex items-center gap-1.5"
                  >
                    <Search className="w-3.5 h-3.5" />
                    <span>Filter Row Range</span>
                  </button>
                </div>
              )}

              {/* Range Toggle */}
              <button
                type="button"
                onClick={() => setIsRowRange(!isRowRange)}
                className="text-xs text-purple-300 hover:text-purple-200 underline font-mono cursor-pointer"
              >
                {isRowRange ? 'Switch to Exact Row #' : 'Switch to Row Range (From/To)'}
              </button>

              {(activeRowSearch !== null || activeRowRange.from !== null) && (
                <button
                  type="button"
                  onClick={() => {
                    setActiveRowSearch(null);
                    setActiveRowRange({ from: null, to: null });
                    setRowSearchInput('');
                    setRowFromInput('');
                    setRowToInput('');
                    fetchPreview(0, { rowNum: null, rowFrom: null, rowTo: null });
                  }}
                  className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs rounded-xl border border-slate-700 transition-colors cursor-pointer"
                >
                  Clear Row Search
                </button>
              )}
            </form>
          )}

          {/* Active Filter Badges Bar */}
          {hasAnyFilterActive && (
            <div className="flex flex-wrap items-center gap-2 pt-1">
              <span className="text-[11px] font-mono text-slate-500 uppercase font-semibold">
                Active Filters:
              </span>

              {activeRowSearch !== null && (
                <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-mono font-bold bg-purple-500/20 text-purple-300 border border-purple-500/40">
                  <Hash className="w-3 h-3" />
                  Exact Row #{activeRowSearch}
                  <button
                    onClick={() => {
                      setActiveRowSearch(null);
                      setRowSearchInput('');
                      fetchPreview(0, { rowNum: null });
                    }}
                    className="hover:text-white ml-1 cursor-pointer"
                  >
                    ×
                  </button>
                </span>
              )}

              {(activeRowRange.from !== null || activeRowRange.to !== null) && (
                <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-mono font-bold bg-purple-500/20 text-purple-300 border border-purple-500/40">
                  <Hash className="w-3 h-3" />
                  Rows {activeRowRange.from || 1} to {activeRowRange.to || '∞'}
                  <button
                    onClick={() => {
                      setActiveRowRange({ from: null, to: null });
                      setRowFromInput('');
                      setRowToInput('');
                      fetchPreview(0, { rowFrom: null, rowTo: null });
                    }}
                    className="hover:text-white ml-1 cursor-pointer"
                  >
                    ×
                  </button>
                </span>
              )}

              {activeColumnFilter && (
                <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-mono font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/40">
                  <Filter className="w-3 h-3" />
                  {activeColumnFilter.column} {activeColumnFilter.operator} "{activeColumnFilter.value}"
                  <button
                    onClick={() => {
                      setActiveColumnFilter(null);
                      setColumnFilterValue('');
                      fetchPreview(0, { colFilter: null });
                    }}
                    className="hover:text-white ml-1 cursor-pointer"
                  >
                    ×
                  </button>
                </span>
              )}

              {generalSearch && (
                <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-mono font-bold bg-blue-500/20 text-blue-300 border border-blue-500/40">
                  <Search className="w-3 h-3" />
                  Text: "{generalSearch}"
                  <button
                    onClick={() => {
                      setGeneralSearch('');
                      fetchPreview(0, { search: '' });
                    }}
                    className="hover:text-white ml-1 cursor-pointer"
                  >
                    ×
                  </button>
                </span>
              )}

              <button
                onClick={clearAllFilters}
                className="text-[11px] text-slate-400 hover:text-rose-400 underline font-mono ml-auto cursor-pointer flex items-center gap-1"
              >
                <RotateCcw className="w-3 h-3" />
                <span>Reset All</span>
              </button>
            </div>
          )}
        </div>

        {/* Spreadsheet Data Grid */}
        <div className="flex-1 overflow-auto bg-slate-950/40 relative">
          {loading && (
            <div className="absolute inset-0 flex items-center justify-center bg-slate-950/60 backdrop-blur-xs z-30">
              <div className="flex items-center gap-2.5 px-5 py-2.5 rounded-xl bg-slate-900 border border-slate-700 text-xs font-mono text-emerald-400 shadow-xl">
                <Loader2 className="w-4 h-4 animate-spin text-emerald-400" />
                <span>Searching sheet records...</span>
              </div>
            </div>
          )}

          {data && data.headers ? (
            <table className="w-full text-left text-xs border-collapse min-w-max">
              <thead className="bg-slate-900 text-slate-400 sticky top-0 z-10 font-mono text-[11px] shadow-sm">
                <tr className="border-b border-slate-800">
                  {/* # Row Header */}
                  <th
                    onClick={() => {
                      const nextOrder = sortBy === 'row_num' && sortOrder === 'asc' ? 'desc' : 'asc';
                      setSortBy('row_num');
                      setSortOrder(nextOrder);
                      fetchPreview(page, { sort: 'row_num', order: nextOrder });
                    }}
                    className="py-2.5 px-3 bg-slate-900 border-r border-slate-800 text-center w-20 font-bold sticky left-0 z-20 shadow-md cursor-pointer hover:bg-slate-800 transition-colors select-none"
                    title="Click to sort by Row Number"
                  >
                    <div className="flex items-center justify-center gap-1 text-slate-300">
                      <span># Row</span>
                      {sortBy === 'row_num' && (
                        sortOrder === 'asc' ? <ArrowUp className="w-3 h-3 text-purple-400" /> : <ArrowDown className="w-3 h-3 text-purple-400" />
                      )}
                    </div>
                  </th>

                  {/* Priority Column Header */}
                  <th
                    className="py-2.5 px-3 bg-slate-900 border-r border-slate-800 text-amber-300 font-bold tracking-wide whitespace-nowrap select-none sticky left-20 z-20 shadow-md"
                    title="Priority of Lead across all files in directory"
                  >
                    <div className="flex items-center gap-1.5">
                      <Award className="w-3.5 h-3.5 text-amber-400" />
                      <span>Priority</span>
                    </div>
                  </th>

                  {/* All Sheet Column Headers */}
                  {regularHeaders.map((col, idx) => {
                    const isFilteredCol = activeColumnFilter?.column === col;
                    return (
                      <th
                        key={idx}
                        className={`py-2.5 px-3 border-r border-slate-800 text-slate-300 font-semibold tracking-wide whitespace-nowrap bg-slate-900 transition-colors select-none ${
                          isFilteredCol ? 'bg-emerald-950/40 text-emerald-300' : ''
                        }`}
                      >
                        <div className="flex items-center justify-between gap-2">
                          <span className="truncate">{col}</span>
                          <button
                            type="button"
                            onClick={() => {
                              setSelectedColumn(col);
                              setActiveMode('column');
                            }}
                            className="p-1 rounded hover:bg-slate-800 text-slate-500 hover:text-emerald-400 transition-colors cursor-pointer"
                            title={`Filter specifically by ${col}`}
                          >
                            <Filter className="w-3 h-3" />
                          </button>
                        </div>
                      </th>
                    );
                  })}

                  <th className="py-2.5 px-3 text-right bg-slate-900 font-bold text-slate-400 w-24">
                    Actions
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60 font-mono text-slate-200">
                {displayedRows.length === 0 ? (
                  <tr>
                    <td colSpan={regularHeaders.length + 3} className="py-16 text-center text-slate-500">
                      <div className="space-y-2">
                        <Search className="w-8 h-8 text-slate-600 mx-auto" />
                        <div className="text-sm font-semibold text-slate-300">No rows matched your search filter</div>
                        <p className="text-xs text-slate-500">
                          Try clearing filters or changing the row search parameters.
                        </p>
                        {hasAnyFilterActive && (
                          <button
                            onClick={clearAllFilters}
                            className="px-3 py-1 bg-slate-800 hover:bg-slate-700 text-emerald-400 rounded-lg text-xs font-mono inline-flex items-center gap-1 mt-2 cursor-pointer"
                          >
                            <RotateCcw className="w-3 h-3" /> Reset Filters
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                ) : (
                  displayedRows.map((row) => {
                    const isExactRowMatch = activeRowSearch === row.rowNum;
                    return (
                      <tr
                        key={row.rowNum}
                        className={`hover:bg-slate-800/50 transition-colors group ${
                          isExactRowMatch ? 'bg-purple-950/30 border-y-2 border-purple-500/60' : ''
                        }`}
                      >
                        {/* Excel Row Number Cell */}
                        <td
                          onClick={() => onViewRecord && onViewRecord(row.id)}
                          className={`py-2 px-3 text-center font-bold border-r border-slate-800 sticky left-0 z-10 select-none cursor-pointer transition-colors ${
                            isExactRowMatch
                              ? 'bg-purple-900/60 text-purple-200 ring-1 ring-purple-400'
                              : 'bg-slate-950 text-slate-400 hover:text-emerald-400 hover:bg-slate-900'
                          }`}
                          title={`Excel Row #${row.rowNum} — Click to view full record`}
                        >
                          <div className="flex items-center justify-center gap-1">
                            <span>{row.rowNum}</span>
                            {isExactRowMatch && (
                              <span className="w-1.5 h-1.5 rounded-full bg-purple-400 animate-ping"></span>
                            )}
                          </div>
                        </td>

                        {/* Priority Column Cell */}
                        <td className="py-2 px-3 text-left border-r border-slate-800 whitespace-nowrap bg-slate-950/80 sticky left-20 z-10">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span
                              className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full font-mono text-[11px] font-bold border ${row.priorityBadge?.bg || 'bg-amber-500/15'} ${row.priorityBadge?.text || 'text-amber-300'} ${row.priorityBadge?.border || 'border-amber-500/30'}`}
                              title={`Priority: ${row.priority || '—'} (Rank #${row.priorityRank || '—'})`}
                            >
                              <Award className="w-3 h-3 text-amber-400 shrink-0" />
                              <span>{row.priority || '—'}</span>
                            </span>
                            {row.occurrencesCount > 1 && (
                              <span
                                className="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-mono bg-blue-500/15 text-blue-300 border border-blue-500/30 cursor-help"
                                title={`Lead found in ${row.occurrencesCount} files across directory: ${(row.matchingFiles || []).join(', ')}`}
                              >
                                {row.occurrencesCount} files
                              </span>
                            )}
                          </div>
                        </td>

                        {/* All Sheet Cells */}
                        {regularHeaders.map((col, cIdx) => {
                          const cellVal = row.values[col];
                          const strVal = cellVal !== null && cellVal !== undefined ? String(cellVal) : '';

                          // Check if cell matches search or filter
                          let isCellMatch = false;
                          if (generalSearch.trim() && strVal.toLowerCase().includes(generalSearch.toLowerCase())) {
                            isCellMatch = true;
                          }
                          if (activeColumnFilter && activeColumnFilter.column === col) {
                            const q = activeColumnFilter.value.toLowerCase();
                            const valLower = strVal.toLowerCase();
                            if (activeColumnFilter.operator === 'equals' && valLower === q) isCellMatch = true;
                            else if (activeColumnFilter.operator === 'startsWith' && valLower.startsWith(q)) isCellMatch = true;
                            else if (activeColumnFilter.operator === 'endsWith' && valLower.endsWith(q)) isCellMatch = true;
                            else if (activeColumnFilter.operator === 'contains' && valLower.includes(q)) isCellMatch = true;
                          }

                          return (
                            <td
                              key={cIdx}
                              className={`py-2 px-3 border-r border-slate-800/60 whitespace-nowrap max-w-md truncate font-sans text-xs select-all ${
                                isCellMatch
                                  ? 'bg-emerald-500/20 text-emerald-200 font-semibold'
                                  : 'text-slate-200'
                              }`}
                            >
                              {strVal === '' ? (
                                <span className="text-slate-700 italic font-mono text-[10px]">&mdash;</span>
                              ) : (
                                strVal
                              )}
                            </td>
                          );
                        })}

                        {/* Actions column */}
                        <td className="py-2 px-3 text-right">
                          <button
                            onClick={() => onViewRecord && onViewRecord(row.id)}
                            className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-slate-800 hover:bg-emerald-600 text-slate-300 hover:text-white text-[11px] font-mono transition-colors cursor-pointer"
                            title="Inspect full row with all original columns"
                          >
                            <Eye className="w-3 h-3" />
                            <span>View</span>
                          </button>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          ) : null}
        </div>

        {/* Footer: Pagination & Navigation */}
        <div className="p-3.5 border-t border-slate-800 bg-slate-950 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-400 font-mono">
          <div className="flex items-center gap-3">
            <span>
              {loadError && <span className="text-rose-400">Error: {loadError}</span>}
              Showing {data && data.rows.length > 0 ? page * limit + 1 : 0} – {Math.min((page + 1) * limit, data?.totalRows || 0)} of {(data?.totalRows ?? 0).toLocaleString()} rows
            </span>
            {hasAnyFilterActive && (
              <span className="text-emerald-400 font-semibold">
                (Filter Active)
              </span>
            )}
          </div>

          {/* Pagination Buttons */}
          <div className="flex items-center gap-2">
            <button
              disabled={page === 0 || loading}
              onClick={() => {
                const newPage = page - 1;
                setPage(newPage);
                fetchPreview(newPage);
              }}
              className="px-2.5 py-1 rounded bg-slate-800 hover:bg-slate-700 disabled:opacity-30 text-slate-300 transition-colors cursor-pointer flex items-center gap-1"
            >
              <ChevronLeft className="w-4 h-4" />
              <span>Prev</span>
            </button>

            <span className="px-2 text-slate-300 font-bold">
              Page {page + 1} of {Math.max(1, totalPages)}
            </span>

            <button
              disabled={page >= totalPages - 1 || loading}
              onClick={() => {
                const newPage = page + 1;
                setPage(newPage);
                fetchPreview(newPage);
              }}
              className="px-2.5 py-1 rounded bg-slate-800 hover:bg-slate-700 disabled:opacity-30 text-slate-300 transition-colors cursor-pointer flex items-center gap-1"
            >
              <span>Next</span>
              <ChevronRight className="w-4 h-4" />
            </button>

            <button
              onClick={onClose}
              className="ml-3 px-4 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 transition-colors cursor-pointer font-sans text-xs font-semibold"
            >
              Close
            </button>
          </div>
        </div>
      </div>

      {/* Lead Priority Rules Modal */}
      <PriorityConfigModal
        isOpen={isPriorityModalOpen}
        onClose={() => setIsPriorityModalOpen(false)}
        onRulesUpdated={(updatedRules) => {
          setActiveRules(updatedRules);
          fetchPreview(page);
        }}
      />
    </div>
  );
};
