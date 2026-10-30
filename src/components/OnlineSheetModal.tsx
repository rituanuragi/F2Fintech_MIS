import React, { useState, useEffect, useMemo } from 'react';
import {
  X,
  FileSpreadsheet,
  Download,
  ExternalLink,
  Search,
  ChevronLeft,
  ChevronRight,
  Maximize2,
  Minimize2,
  Copy,
  Check,
  RotateCcw,
  Layers,
  ArrowUpDown,
  Filter,
  Loader2,
  Table,
} from 'lucide-react';

export interface OnlineSheetModalProps {
  isOpen: boolean;
  onClose: () => void;
  documentUrl: string | null;
  fileName?: string;
  ticketId?: string | number | null;
  activityAuthor?: string | null;
  comment?: string | null;
}

interface ParsedSheetData {
  headers: string[];
  rows: any[][];
  totalRows: number;
  totalCols: number;
}

interface ParsedWorkbook {
  success: boolean;
  fileName: string;
  sheetNames: string[];
  sheets: Record<string, ParsedSheetData>;
  error?: string;
}

export const OnlineSheetModal: React.FC<OnlineSheetModalProps> = ({
  isOpen,
  onClose,
  documentUrl,
  fileName,
  ticketId,
  activityAuthor,
  comment,
}) => {
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const [workbook, setWorkbook] = useState<ParsedWorkbook | null>(null);
  const [activeSheetName, setActiveSheetName] = useState<string>('');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [page, setPage] = useState<number>(0);
  const [rowsPerPage, setRowsPerPage] = useState<number>(50);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);
  const [copiedCellVal, setCopiedCellVal] = useState<string | null>(null);
  const [sortCol, setSortCol] = useState<number | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');

  // Close on Escape key
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && isOpen) {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

  // Fetch parsed sheet when modal opens or documentUrl changes
  useEffect(() => {
    if (!isOpen || !documentUrl) {
      setWorkbook(null);
      setError(null);
      return;
    }

    let isMounted = true;
    setLoading(true);
    setError(null);
    setSearchQuery('');
    setPage(0);
    setSortCol(null);

    const fetchUrl = `/api/f2-tickets/preview-sheet?url=${encodeURIComponent(
      documentUrl
    )}&filename=${encodeURIComponent(fileName || '')}`;

    fetch(fetchUrl)
      .then(async (res) => {
        if (!res.ok) {
          const errData = await res.json().catch(() => ({}));
          throw new Error(errData.error || `Server returned ${res.status}`);
        }
        return res.json();
      })
      .then((data: ParsedWorkbook) => {
        if (!isMounted) return;
        if (data.success && data.sheetNames && data.sheetNames.length > 0) {
          setWorkbook(data);
          setActiveSheetName(data.sheetNames[0]);
        } else {
          setError(data.error || 'Failed to parse spreadsheet content.');
        }
      })
      .catch((err) => {
        if (!isMounted) return;
        setError(err.message || 'Unable to load and parse sheet online.');
      })
      .finally(() => {
        if (isMounted) setLoading(false);
      });

    return () => {
      isMounted = false;
    };
  }, [isOpen, documentUrl, fileName]);

  // Current sheet data
  const currentSheet: ParsedSheetData | null = useMemo(() => {
    if (!workbook || !activeSheetName || !workbook.sheets[activeSheetName]) return null;
    return workbook.sheets[activeSheetName];
  }, [workbook, activeSheetName]);

  // Filtered & sorted rows
  const filteredRows = useMemo(() => {
    if (!currentSheet) return [];
    let rows = currentSheet.rows;

    // Search filter across all columns
    const q = searchQuery.trim().toLowerCase();
    if (q) {
      rows = rows.filter((r) =>
        r.some((cell) => cell !== null && cell !== undefined && String(cell).toLowerCase().includes(q))
      );
    }

    // Column sorting
    if (sortCol !== null && sortCol >= 0) {
      rows = [...rows].sort((a, b) => {
        const valA = a[sortCol] ?? '';
        const valB = b[sortCol] ?? '';
        const numA = Number(valA);
        const numB = Number(valB);

        let cmp = 0;
        if (!isNaN(numA) && !isNaN(numB) && String(valA).trim() !== '' && String(valB).trim() !== '') {
          cmp = numA - numB;
        } else {
          cmp = String(valA).localeCompare(String(valB), undefined, { numeric: true });
        }
        return sortDirection === 'asc' ? cmp : -cmp;
      });
    }

    return rows;
  }, [currentSheet, searchQuery, sortCol, sortDirection]);

  // Paginated rows
  const totalFiltered = filteredRows.length;
  const totalPages = Math.ceil(totalFiltered / rowsPerPage) || 1;
  const paginatedRows = useMemo(() => {
    const start = page * rowsPerPage;
    return filteredRows.slice(start, start + rowsPerPage);
  }, [filteredRows, page, rowsPerPage]);

  const handleCopyCell = (text: string) => {
    if (!text) return;
    navigator.clipboard.writeText(text);
    setCopiedCellVal(text);
    setTimeout(() => setCopiedCellVal(null), 1800);
  };

  const handleSort = (colIdx: number) => {
    if (sortCol === colIdx) {
      if (sortDirection === 'asc') {
        setSortDirection('desc');
      } else {
        setSortCol(null);
        setSortDirection('asc');
      }
    } else {
      setSortCol(colIdx);
      setSortDirection('asc');
    }
  };

  const resolvedFileName =
    fileName || workbook?.fileName || (documentUrl ? documentUrl.split('/').pop()?.split('?')[0] : 'Spreadsheet.xlsx');

  const downloadUrl = documentUrl
    ? `/api/f2-tickets/proxy-document?url=${encodeURIComponent(
        documentUrl
      )}&download=1&filename=${encodeURIComponent(resolvedFileName || 'sheet.xlsx')}`
    : null;

  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-xs flex items-center justify-center p-2 sm:p-4 animate-in fade-in duration-150"
      onClick={onClose}
    >
      <div
        className={`bg-white rounded-2xl sm:rounded-3xl shadow-2xl border border-slate-200 flex flex-col overflow-hidden transition-all ${
          isFullscreen
            ? 'w-full h-full rounded-none'
            : 'w-full max-w-[96vw] xl:max-w-7xl h-[92vh] max-h-[95vh]'
        }`}
        onClick={(e) => e.stopPropagation()}
      >
        {/* ==================================================================== */}
        {/* TOP BAR: File Info & Controls */}
        {/* ==================================================================== */}
        <div className="px-5 py-3.5 bg-slate-900 text-white flex flex-wrap items-center justify-between gap-3 border-b border-slate-800 shrink-0">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-10 h-10 rounded-xl bg-emerald-500/20 text-emerald-400 border border-emerald-400/30 flex items-center justify-center shrink-0">
              <FileSpreadsheet className="w-5 h-5" />
            </div>

            <div className="min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="px-2 py-0.5 rounded text-[10px] font-extrabold uppercase bg-emerald-500/20 text-emerald-300 border border-emerald-400/30">
                  Online Sheet Preview
                </span>
                {ticketId && (
                  <span className="px-2 py-0.5 rounded text-[10px] font-mono font-bold bg-blue-500/20 text-blue-300 border border-blue-400/30">
                    Ticket #{ticketId}
                  </span>
                )}
                {activityAuthor && (
                  <span className="text-[11px] text-slate-300 font-medium truncate">
                    Attached by <strong className="text-white">{activityAuthor}</strong>
                  </span>
                )}
              </div>
              <h3
                className="text-sm sm:text-base font-extrabold text-white truncate max-w-lg mt-0.5"
                title={resolvedFileName || ''}
              >
                {resolvedFileName}
              </h3>
            </div>
          </div>

          {/* Action buttons: Download, Open URL, Fullscreen, Close */}
          <div className="flex items-center gap-2 shrink-0">
            {downloadUrl && (
              <a
                href={downloadUrl}
                download={resolvedFileName || 'sheet.xlsx'}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs transition shadow-xs cursor-pointer"
                title="Download spreadsheet to device"
              >
                <Download className="w-3.5 h-3.5" />
                <span className="hidden md:inline">Download</span>
              </a>
            )}

            {documentUrl && (
              <a
                href={documentUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-white/10 hover:bg-white/20 text-slate-200 font-bold text-xs transition"
                title="Open raw file URL in new tab"
              >
                <ExternalLink className="w-3.5 h-3.5" />
                <span className="hidden lg:inline">Raw Link</span>
              </a>
            )}

            <button
              type="button"
              onClick={() => setIsFullscreen(!isFullscreen)}
              className="p-2 rounded-xl bg-white/10 hover:bg-white/20 text-slate-300 hover:text-white transition cursor-pointer"
              title={isFullscreen ? 'Exit Fullscreen' : 'Fullscreen'}
            >
              {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
            </button>

            <button
              type="button"
              onClick={onClose}
              className="p-2 rounded-xl bg-white/10 hover:bg-rose-500/30 hover:text-rose-200 text-slate-300 transition cursor-pointer"
              title="Close (Esc)"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Remark Comment Context Banner (if provided) */}
        {comment && (
          <div className="px-5 py-2 bg-amber-50 border-b border-amber-200/70 text-xs text-amber-900 flex items-center justify-between gap-3 shrink-0">
            <div className="flex items-center gap-2 truncate">
              <span className="font-extrabold uppercase text-[10px] px-1.5 py-0.5 rounded bg-amber-200/60 text-amber-900">
                Remark Note
              </span>
              <span className="truncate italic">"{comment}"</span>
            </div>
            <span className="text-[10px] text-amber-700 shrink-0 font-medium">
              Sheet preview rendered securely in-app without downloading
            </span>
          </div>
        )}

        {/* ==================================================================== */}
        {/* TABS & TOOLBAR */}
        {/* ==================================================================== */}
        <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
          {/* Sheet Selector Tabs */}
          <div className="flex items-center gap-1.5 overflow-x-auto max-w-full pb-1 sm:pb-0">
            <span className="text-xs font-bold text-slate-500 flex items-center gap-1 mr-1 shrink-0">
              <Layers className="w-3.5 h-3.5 text-slate-400" />
              <span>Sheets:</span>
            </span>

            {workbook?.sheetNames.map((sName) => {
              const sData = workbook.sheets[sName];
              const isActive = sName === activeSheetName;
              return (
                <button
                  key={sName}
                  type="button"
                  onClick={() => {
                    setActiveSheetName(sName);
                    setPage(0);
                    setSortCol(null);
                  }}
                  className={`px-3 py-1.5 rounded-xl text-xs font-bold transition flex items-center gap-2 shrink-0 cursor-pointer ${
                    isActive
                      ? 'bg-emerald-600 text-white shadow-xs'
                      : 'bg-white text-slate-700 hover:bg-slate-200/80 border border-slate-200'
                  }`}
                >
                  <FileSpreadsheet className={`w-3.5 h-3.5 ${isActive ? 'text-white' : 'text-emerald-600'}`} />
                  <span>{sName}</span>
                  {sData && (
                    <span
                      className={`text-[10px] px-1.5 py-0.2 rounded-full font-mono ${
                        isActive ? 'bg-emerald-700 text-emerald-100' : 'bg-slate-100 text-slate-500'
                      }`}
                    >
                      {sData.totalRows.toLocaleString()} rows
                    </span>
                  )}
                </button>
              );
            })}
          </div>

          {/* Quick Search inside Sheet & Rows Per Page */}
          <div className="flex items-center gap-3 w-full sm:w-auto justify-between sm:justify-end">
            <div className="relative flex-1 sm:w-64">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => {
                  setSearchQuery(e.target.value);
                  setPage(0);
                }}
                placeholder="Search cells in this sheet..."
                className="w-full pl-8 pr-8 py-1.5 rounded-xl border border-slate-200 text-xs bg-white focus:outline-hidden focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 font-medium"
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => {
                    setSearchQuery('');
                    setPage(0);
                  }}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              )}
            </div>

            <div className="flex items-center gap-1.5 text-xs text-slate-500 shrink-0">
              <span className="hidden md:inline">Rows:</span>
              <select
                value={rowsPerPage}
                onChange={(e) => {
                  setRowsPerPage(Number(e.target.value));
                  setPage(0);
                }}
                className="px-2 py-1.5 rounded-xl border border-slate-200 bg-white text-xs font-bold text-slate-700 focus:outline-hidden cursor-pointer"
              >
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
                <option value={200}>200</option>
              </select>
            </div>
          </div>
        </div>

        {/* ==================================================================== */}
        {/* MAIN BODY: Grid / Loader / Error */}
        {/* ==================================================================== */}
        <div className="flex-1 overflow-auto bg-slate-100/70 relative">
          {loading ? (
            <div className="h-full flex flex-col items-center justify-center p-8 space-y-3">
              <div className="w-12 h-12 rounded-2xl bg-emerald-100 text-emerald-600 flex items-center justify-center animate-pulse">
                <FileSpreadsheet className="w-6 h-6" />
              </div>
              <div className="flex items-center gap-2 text-slate-700 font-bold text-sm">
                <Loader2 className="w-4 h-4 animate-spin text-emerald-600" />
                <span>Reading and parsing spreadsheet online...</span>
              </div>
              <p className="text-xs text-slate-400 max-w-sm text-center">
                Fetching workbook data directly to display all columns, sheets, and rows in-browser.
              </p>
            </div>
          ) : error ? (
            <div className="h-full flex flex-col items-center justify-center p-8 space-y-4 max-w-md mx-auto text-center">
              <div className="w-12 h-12 rounded-2xl bg-rose-100 text-rose-600 flex items-center justify-center">
                <X className="w-6 h-6" />
              </div>
              <div>
                <h4 className="font-extrabold text-slate-800 text-base">Unable to Preview Sheet</h4>
                <p className="text-xs text-slate-500 mt-1">{error}</p>
              </div>
              {downloadUrl && (
                <div className="flex items-center gap-2 pt-2">
                  <a
                    href={downloadUrl}
                    download={resolvedFileName || 'sheet.xlsx'}
                    className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs"
                  >
                    <Download className="w-4 h-4" />
                    <span>Download File Directly</span>
                  </a>
                </div>
              )}
            </div>
          ) : !currentSheet || currentSheet.headers.length === 0 ? (
            <div className="h-full flex flex-col items-center justify-center p-8 text-center text-slate-400 space-y-2">
              <Table className="w-10 h-10 mx-auto text-slate-300" />
              <p className="font-bold text-sm text-slate-700">This sheet appears to be empty.</p>
              <p className="text-xs text-slate-400">No rows or header columns discovered in {activeSheetName}.</p>
            </div>
          ) : (
            <div className="min-w-full inline-block align-middle pb-16">
              <table className="w-full border-collapse text-left text-xs bg-white select-text">
                {/* Table Header */}
                <thead className="sticky top-0 z-20 bg-slate-900 text-white shadow-xs">
                  <tr>
                    {/* Row Number Column */}
                    <th className="sticky left-0 z-30 bg-slate-950 px-3 py-2.5 text-[10px] font-mono font-extrabold text-slate-400 border-b border-r border-slate-800 text-center w-12 shrink-0 select-none">
                      #
                    </th>
                    {currentSheet.headers.map((header, hIdx) => {
                      const isSorted = sortCol === hIdx;
                      return (
                        <th
                          key={hIdx}
                          onClick={() => handleSort(hIdx)}
                          className="px-3.5 py-2.5 text-xs font-bold text-slate-100 border-b border-r border-slate-800 hover:bg-slate-800/80 transition cursor-pointer select-none whitespace-nowrap min-w-[120px]"
                          title={`Sort by column ${header}`}
                        >
                          <div className="flex items-center justify-between gap-2">
                            <span className="truncate">{header}</span>
                            <div className="text-slate-400 shrink-0">
                              {isSorted ? (
                                <span className="text-emerald-400 font-mono text-[10px]">
                                  {sortDirection === 'asc' ? '▲' : '▼'}
                                </span>
                              ) : (
                                <ArrowUpDown className="w-3 h-3 opacity-40 hover:opacity-100" />
                              )}
                            </div>
                          </div>
                        </th>
                      );
                    })}
                  </tr>
                </thead>

                {/* Table Body */}
                <tbody className="divide-y divide-slate-100">
                  {paginatedRows.length === 0 ? (
                    <tr>
                      <td
                        colSpan={currentSheet.headers.length + 1}
                        className="py-12 text-center text-slate-400 text-xs"
                      >
                        No rows matching filter "{searchQuery}".
                      </td>
                    </tr>
                  ) : (
                    paginatedRows.map((row, rIdx) => {
                      const rowNum = page * rowsPerPage + rIdx + 1;
                      const isEven = rIdx % 2 === 0;

                      return (
                        <tr
                          key={rIdx}
                          className={`${
                            isEven ? 'bg-white' : 'bg-slate-50/70'
                          } hover:bg-emerald-50/50 transition-colors group`}
                        >
                          {/* Row Number */}
                          <td className="sticky left-0 z-10 bg-slate-100/90 group-hover:bg-emerald-100/80 px-2 py-2 text-center font-mono text-[10px] text-slate-500 font-bold border-r border-slate-200 select-none">
                            {rowNum}
                          </td>

                          {/* Row Cells */}
                          {currentSheet.headers.map((_, cIdx) => {
                            const rawCell = row[cIdx];
                            const cellStr =
                              rawCell !== null && rawCell !== undefined ? String(rawCell).trim() : '';

                            const isNumeric =
                              cellStr !== '' &&
                              !isNaN(Number(cellStr.replace(/,/g, ''))) &&
                              cellStr.length < 15;

                            return (
                              <td
                                key={cIdx}
                                onClick={() => handleCopyCell(cellStr)}
                                className={`px-3.5 py-2 border-r border-slate-100 text-slate-800 break-words max-w-xs transition-colors cursor-pointer hover:bg-emerald-100/40 relative ${
                                  isNumeric ? 'font-mono text-slate-900 text-right' : 'font-sans'
                                }`}
                                title="Click to copy cell value"
                              >
                                {cellStr || <span className="text-slate-300 italic">—</span>}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          )}

          {/* Quick Copied Tooltip Toast */}
          {copiedCellVal !== null && (
            <div className="fixed bottom-16 right-8 z-50 bg-slate-900 text-white text-xs px-3.5 py-2 rounded-xl shadow-xl flex items-center gap-2 border border-slate-700 animate-in fade-in duration-100">
              <Check className="w-3.5 h-3.5 text-emerald-400" />
              <span>Copied cell value to clipboard</span>
            </div>
          )}
        </div>

        {/* ==================================================================== */}
        {/* FOOTER: Pagination & Summary */}
        {/* ==================================================================== */}
        <div className="px-5 py-3 bg-white border-t border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-3 text-xs text-slate-600">
            {currentSheet && (
              <span>
                Showing{' '}
                <strong className="text-slate-900">
                  {totalFiltered === 0 ? 0 : page * rowsPerPage + 1}
                </strong>{' '}
                to{' '}
                <strong className="text-slate-900">
                  {Math.min((page + 1) * rowsPerPage, totalFiltered)}
                </strong>{' '}
                of <strong className="text-slate-900">{totalFiltered.toLocaleString()}</strong> rows{' '}
                {searchQuery && `(filtered from ${currentSheet.totalRows.toLocaleString()})`}
                {' • '}
                <strong className="text-slate-900">{currentSheet.headers.length}</strong> columns
              </span>
            )}
            <span className="text-[11px] text-slate-400 hidden sm:inline">
              Tip: Click any cell to copy value
            </span>
          </div>

          {/* Pagination Controls */}
          <div className="flex items-center gap-2">
            <button
              type="button"
              disabled={page <= 0}
              onClick={() => setPage((p) => Math.max(0, p - 1))}
              className="inline-flex items-center gap-1 px-3 py-1.5 rounded-xl border border-slate-200 text-xs font-bold text-slate-700 hover:bg-slate-100 disabled:opacity-30 disabled:pointer-events-none transition cursor-pointer"
            >
              <ChevronLeft className="w-4 h-4" />
              <span>Prev</span>
            </button>

            <span className="text-xs font-mono font-bold text-slate-700 px-2">
              Page {page + 1} of {totalPages}
            </span>

            <button
              type="button"
              disabled={page >= totalPages - 1}
              onClick={() => setPage((p) => Math.min(totalPages - 1, p + 1))}
              className="inline-flex items-center gap-1 px-3 py-1.5 rounded-xl border border-slate-200 text-xs font-bold text-slate-700 hover:bg-slate-100 disabled:opacity-30 disabled:pointer-events-none transition cursor-pointer"
            >
              <span>Next</span>
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
