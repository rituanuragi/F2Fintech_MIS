import React, { useState, useEffect, useMemo } from 'react';
import {
  X,
  FileSpreadsheet,
  Download,
  ExternalLink,
  Search,
  ChevronLeft,
  ChevronRight,
  ZoomIn,
  ZoomOut,
  RotateCcw,
  Loader2,
  AlertCircle,
  Table,
  Layers,
} from 'lucide-react';
import * as XLSXNamespace from 'xlsx';
const XLSX: any = (XLSXNamespace as any).default || XLSXNamespace;

export interface OnlineSheetViewerModalProps {
  url: string;
  fileName: string;
  title?: string;
  onClose: () => void;
}

interface SheetData {
  headers: string[];
  rows: (string | number | boolean | null)[][];
  totalRows: number;
  totalCols: number;
}

interface ApiResponse {
  success: boolean;
  fileName: string;
  sheetNames: string[];
  sheets: Record<string, SheetData>;
  error?: string;
}

export const OnlineSheetViewerModal: React.FC<OnlineSheetViewerModalProps> = ({
  url,
  fileName,
  title,
  onClose,
}) => {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState<string>('');
  const [sheets, setSheets] = useState<Record<string, SheetData>>({});

  // Search & Pagination & Zoom State
  const [searchQuery, setSearchQuery] = useState('');
  const [pageSize, setPageSize] = useState<number | 'all'>(50);
  const [currentPage, setCurrentPage] = useState(0);
  const [zoomLevel, setZoomLevel] = useState<number>(100);
  const [syncVersionTrigger, setSyncVersionTrigger] = useState<number>(0);
  const [lastLiveUpdateText, setLastLiveUpdateText] = useState<string | null>(null);
  const prevSyncVerRef = React.useRef<number | null>(null);

  // Poll /api/mis/auto-sync-status while modal is open so new MIS Sheet entries appear automatically
  useEffect(() => {
    const checkSync = async () => {
      try {
        const res = await fetch('/api/mis/auto-sync-status');
        const data = await res.json();
        if (data?.success && data.autoSync) {
          const ver = data.autoSync.syncVersion || 0;
          if (prevSyncVerRef.current !== null && ver > prevSyncVerRef.current) {
            prevSyncVerRef.current = ver;
            setSyncVersionTrigger(ver);
            setLastLiveUpdateText(`Auto-updated at ${new Date().toLocaleTimeString()}`);
          } else if (prevSyncVerRef.current === null) {
            prevSyncVerRef.current = ver;
          }
        }
      } catch {
        // ignore
      }
    };
    checkSync();
    const interval = setInterval(checkSync, 10000);
    return () => clearInterval(interval);
  }, []);

  // Close on Escape key
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [onClose]);

  // Load and parse the sheet online
  useEffect(() => {
    let isCancelled = false;

    const loadSheet = async (isBackgroundRefresh: boolean) => {
      if (!isBackgroundRefresh) {
        setLoading(true);
      }
      setError(null);

      // Strategy 1: Try backend preview-sheet endpoint
      try {
        const previewApiUrl = `/api/f2-tickets/preview-sheet?url=${encodeURIComponent(
          url
        )}&filename=${encodeURIComponent(fileName)}&_t=${Date.now()}`;

        const res = await fetch(previewApiUrl);
        const data: ApiResponse = await res.json();

        if (!isCancelled && data.success && data.sheetNames && data.sheetNames.length > 0) {
          setSheetNames(data.sheetNames);
          setActiveSheet((prev) => (prev && data.sheetNames.includes(prev) ? prev : data.sheetNames[0]));
          setSheets(data.sheets || {});
          setLoading(false);
          return;
        }

        if (data.error && !isCancelled) {
          throw new Error(data.error);
        }
      } catch (backendErr: any) {
        console.warn('Backend sheet preview returned error, trying client-side proxy parse:', backendErr);
      }

      // Strategy 2: Fallback to downloading via proxy and parsing in browser with XLSX
      try {
        const proxyUrl = `/api/f2-tickets/proxy-document?url=${encodeURIComponent(url)}`;
        const proxyRes = await fetch(proxyUrl);
        if (!proxyRes.ok) {
          throw new Error(`Failed to fetch document through proxy (HTTP ${proxyRes.status})`);
        }

        const arrayBuffer = await proxyRes.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
        const sNames = workbook.SheetNames || [];

        if (sNames.length === 0) {
          throw new Error('Workbook contains no readable sheets.');
        }

        const sheetsRecord: Record<string, SheetData> = {};

        for (const sName of sNames) {
          const ws = workbook.Sheets[sName];
          if (!ws) continue;
          const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });
          if (!rawRows || rawRows.length === 0) {
            sheetsRecord[sName] = { headers: [], rows: [], totalRows: 0, totalCols: 0 };
            continue;
          }

          let headerIdx = 0;
          for (let i = 0; i < Math.min(rawRows.length, 5); i++) {
            if (Array.isArray(rawRows[i]) && rawRows[i].some((c: any) => c !== null && c !== undefined && String(c).trim() !== '')) {
              headerIdx = i;
              break;
            }
          }

          const hRow = rawRows[headerIdx] || [];
          const dRows = rawRows.slice(headerIdx + 1);

          let maxC = hRow.length;
          for (const r of dRows) {
            if (Array.isArray(r) && r.length > maxC) maxC = r.length;
          }
          maxC = Math.max(maxC, 1);

          const hdrs = Array.from({ length: maxC }, (_, i) => {
            const v = hRow[i];
            if (v !== null && v !== undefined && String(v).trim() !== '') {
              return String(v).trim();
            }
            let letter = '';
            let n = i;
            while (n >= 0) {
              letter = String.fromCharCode((n % 26) + 65) + letter;
              n = Math.floor(n / 26) - 1;
            }
            return `Col ${letter}`;
          });

          const normRows = dRows.map((r: any[]) => {
            const rowArr = Array.isArray(r) ? r : [];
            const padded: any[] = [];
            for (let c = 0; c < maxC; c++) {
              const cell = rowArr[c];
              if (cell === null || cell === undefined) {
                padded.push('');
              } else if (cell instanceof Date) {
                padded.push(cell.toLocaleDateString('en-GB'));
              } else {
                padded.push(String(cell));
              }
            }
            return padded;
          });

          sheetsRecord[sName] = {
            headers: hdrs,
            rows: normRows,
            totalRows: normRows.length,
            totalCols: hdrs.length,
          };
        }

        if (!isCancelled) {
          setSheetNames(sNames);
          setActiveSheet((prev) => (prev && sNames.includes(prev) ? prev : sNames[0]));
          setSheets(sheetsRecord);
          setLoading(false);
        }
      } catch (err: any) {
        if (!isCancelled) {
          setError(err.message || 'Could not parse spreadsheet data.');
          setLoading(false);
        }
      }
    };

    loadSheet(syncVersionTrigger > 0);

    return () => {
      isCancelled = true;
    };
  }, [url, fileName, syncVersionTrigger]);

  const currentSheetData: SheetData | undefined = sheets[activeSheet];

  // Filter rows based on search query
  const filteredRows = useMemo(() => {
    if (!currentSheetData) return [];
    if (!searchQuery.trim()) return currentSheetData.rows;

    const q = searchQuery.toLowerCase().trim();
    return currentSheetData.rows.filter((row) =>
      row.some((cell) => cell !== null && cell !== undefined && String(cell).toLowerCase().includes(q))
    );
  }, [currentSheetData, searchQuery]);

  // Paginated rows
  const paginatedRows = useMemo(() => {
    if (pageSize === 'all') return filteredRows;
    const start = currentPage * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, currentPage, pageSize]);

  const totalPages = pageSize === 'all' ? 1 : Math.ceil(filteredRows.length / pageSize);

  const downloadUrl = `/api/f2-tickets/proxy-document?url=${encodeURIComponent(
    url
  )}&download=1&filename=${encodeURIComponent(fileName || 'spreadsheet.xlsx')}`;

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-xs flex items-center justify-center p-2 sm:p-4"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-2xl sm:rounded-3xl w-full max-w-[96vw] h-[92vh] flex flex-col overflow-hidden shadow-2xl border border-slate-200 animate-in fade-in zoom-in-95 duration-150"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Top Header Bar */}
        <div className="px-5 py-3.5 bg-gradient-to-r from-emerald-950 via-slate-900 to-slate-950 text-white flex flex-wrap items-center justify-between gap-3 shrink-0 border-b border-emerald-800/40">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-10 h-10 rounded-xl bg-emerald-600/30 border border-emerald-500/40 text-emerald-400 flex items-center justify-center shrink-0 shadow-xs">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div className="min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="px-2 py-0.5 rounded-md bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 font-extrabold text-[10px] uppercase tracking-wider">
                  Online Sheet Viewer
                </span>
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-teal-500/20 text-teal-200 border border-teal-500/30 font-bold text-[10px]">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                  Live Auto-Sync
                </span>
                {lastLiveUpdateText && (
                  <span className="px-2 py-0.5 rounded-md bg-indigo-500/30 text-indigo-200 border border-indigo-400/30 font-bold text-[10px]">
                    {lastLiveUpdateText}
                  </span>
                )}
                <h3 className="font-extrabold text-sm sm:text-base text-white truncate max-w-sm sm:max-w-md" title={fileName}>
                  {fileName}
                </h3>
              </div>
              <div className="flex items-center gap-2 text-xs text-slate-300 font-mono mt-0.5">
                {title && <span className="text-emerald-400 font-bold">{title}</span>}
                {currentSheetData && (
                  <span>
                    • {currentSheetData.totalRows.toLocaleString()} rows • {currentSheetData.totalCols} columns
                  </span>
                )}
                {sheetNames.length > 1 && (
                  <span className="text-slate-400">
                    • {sheetNames.length} sheets in workbook
                  </span>
                )}
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <a
              href={downloadUrl}
              download={fileName || 'spreadsheet.xlsx'}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs transition cursor-pointer shadow-xs"
              title="Download original Excel file"
            >
              <Download className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Download Excel</span>
            </a>

            <a
              href={url}
              target="_blank"
              rel="noopener noreferrer"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 font-bold text-xs transition cursor-pointer"
              title="Open document source URL"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Source Link</span>
            </a>

            <button
              type="button"
              onClick={onClose}
              className="p-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white transition cursor-pointer"
              title="Close Viewer (Esc)"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Sheet Tabs Bar (if multiple sheets exist) */}
        {sheetNames.length > 1 && (
          <div className="px-5 py-2 bg-slate-100/90 border-b border-slate-200 flex items-center gap-1.5 overflow-x-auto shrink-0 scrollbar-thin">
            <div className="flex items-center gap-1.5 text-xs text-slate-500 font-bold mr-1 shrink-0">
              <Layers className="w-3.5 h-3.5 text-emerald-600" />
              <span>Sheets:</span>
            </div>
            {sheetNames.map((sName) => {
              const isCurrent = sName === activeSheet;
              const count = sheets[sName]?.totalRows ?? 0;
              return (
                <button
                  key={sName}
                  type="button"
                  onClick={() => {
                    setActiveSheet(sName);
                    setCurrentPage(0);
                  }}
                  className={`px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer whitespace-nowrap flex items-center gap-1.5 ${
                    isCurrent
                      ? 'bg-emerald-600 text-white shadow-2xs font-extrabold'
                      : 'bg-white hover:bg-slate-200 text-slate-700 border border-slate-200'
                  }`}
                >
                  <FileSpreadsheet className={`w-3 h-3 ${isCurrent ? 'text-emerald-100' : 'text-slate-400'}`} />
                  <span>{sName}</span>
                  <span className={`text-[10px] px-1.5 py-0.2 rounded-md ${isCurrent ? 'bg-emerald-700/80 text-emerald-100' : 'bg-slate-100 text-slate-500'}`}>
                    {count}
                  </span>
                </button>
              );
            })}
          </div>
        )}

        {/* Toolbar Bar: Search, Pagination & Zoom Controls */}
        <div className="px-5 py-2.5 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-2.5 flex-1 min-w-[240px] max-w-md">
            <div className="relative flex-1">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => {
                  setSearchQuery(e.target.value);
                  setCurrentPage(0);
                }}
                placeholder="Search across all cells in this sheet..."
                className="w-full pl-8 pr-7 py-1.5 rounded-xl border border-slate-300 bg-white text-xs text-slate-800 placeholder-slate-400 focus:outline-hidden focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-600 font-medium"
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700 p-0.5 cursor-pointer"
                >
                  <X className="w-3 h-3" />
                </button>
              )}
            </div>
            {searchQuery && (
              <span className="text-[11px] font-bold text-slate-600 shrink-0">
                {filteredRows.length} matching row{filteredRows.length === 1 ? '' : 's'}
              </span>
            )}
          </div>

          <div className="flex items-center gap-3 flex-wrap">
            {/* Rows per page selector */}
            <div className="flex items-center gap-1.5 text-xs text-slate-600">
              <span className="text-slate-400 text-[11px] font-bold">Rows:</span>
              {([25, 50, 100, 'all'] as const).map((sz) => (
                <button
                  key={String(sz)}
                  type="button"
                  onClick={() => {
                    setPageSize(sz);
                    setCurrentPage(0);
                  }}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold cursor-pointer transition ${
                    pageSize === sz
                      ? 'bg-emerald-600 text-white'
                      : 'bg-white text-slate-700 border border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  {sz === 'all' ? 'All' : sz}
                </button>
              ))}
            </div>

            {/* Pagination Controls (if not 'all') */}
            {pageSize !== 'all' && totalPages > 1 && (
              <div className="flex items-center gap-1 text-xs">
                <button
                  type="button"
                  disabled={currentPage === 0}
                  onClick={() => setCurrentPage((p) => Math.max(0, p - 1))}
                  className="p-1 rounded-lg bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 disabled:opacity-40 cursor-pointer"
                  title="Previous page"
                >
                  <ChevronLeft className="w-3.5 h-3.5" />
                </button>
                <span className="font-mono font-bold text-slate-700 text-xs px-1">
                  Page {currentPage + 1} of {totalPages}
                </span>
                <button
                  type="button"
                  disabled={currentPage >= totalPages - 1}
                  onClick={() => setCurrentPage((p) => Math.min(totalPages - 1, p + 1))}
                  className="p-1 rounded-lg bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 disabled:opacity-40 cursor-pointer"
                  title="Next page"
                >
                  <ChevronRight className="w-3.5 h-3.5" />
                </button>
              </div>
            )}

            {/* Zoom Controls */}
            <div className="flex items-center gap-1 border-l border-slate-200 pl-2">
              <button
                type="button"
                onClick={() => setZoomLevel((z) => Math.max(80, z - 10))}
                className="p-1 rounded-lg bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 cursor-pointer"
                title="Zoom Out"
              >
                <ZoomOut className="w-3.5 h-3.5" />
              </button>
              <span className="font-mono text-[11px] font-bold text-slate-600 w-10 text-center">
                {zoomLevel}%
              </span>
              <button
                type="button"
                onClick={() => setZoomLevel((z) => Math.min(140, z + 10))}
                className="p-1 rounded-lg bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 cursor-pointer"
                title="Zoom In"
              >
                <ZoomIn className="w-3.5 h-3.5" />
              </button>
              {zoomLevel !== 100 && (
                <button
                  type="button"
                  onClick={() => setZoomLevel(100)}
                  className="p-1 rounded-lg bg-slate-100 text-slate-500 hover:text-slate-800 cursor-pointer text-[10px]"
                  title="Reset Zoom"
                >
                  <RotateCcw className="w-3 h-3" />
                </button>
              )}
            </div>
          </div>
        </div>

        {/* Spreadsheet Content Area */}
        <div className="flex-1 overflow-auto bg-slate-100/60 relative">
          {loading ? (
            <div className="absolute inset-0 flex flex-col items-center justify-center space-y-3 bg-white/90">
              <Loader2 className="w-8 h-8 text-emerald-600 animate-spin" />
              <div className="text-center">
                <p className="font-extrabold text-sm text-slate-800">
                  Loading Excel Spreadsheet...
                </p>
                <p className="text-xs text-slate-500 mt-1 max-w-sm">
                  Fetching and parsing online sheet data for <span className="font-mono font-bold text-slate-700">{fileName}</span>
                </p>
              </div>
            </div>
          ) : error ? (
            <div className="p-10 max-w-md mx-auto my-auto text-center space-y-4 bg-white rounded-2xl border border-rose-200 shadow-sm mt-12">
              <AlertCircle className="w-10 h-10 text-rose-500 mx-auto" />
              <div>
                <h4 className="font-extrabold text-slate-800 text-base">
                  Could Not Preview Sheet Online
                </h4>
                <p className="text-xs text-slate-500 mt-1">{error}</p>
              </div>
              <div className="flex flex-col gap-2 pt-2">
                <a
                  href={downloadUrl}
                  download={fileName || 'spreadsheet.xlsx'}
                  className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shadow-xs"
                >
                  <Download className="w-4 h-4" />
                  <span>Download Original Excel File Instead</span>
                </a>
                <a
                  href={url}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="w-full inline-flex items-center justify-center gap-2 px-4 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs"
                >
                  <ExternalLink className="w-4 h-4" />
                  <span>Open Direct Link</span>
                </a>
              </div>
            </div>
          ) : !currentSheetData || currentSheetData.headers.length === 0 ? (
            <div className="p-12 text-center text-slate-400 space-y-2 mt-8">
              <FileSpreadsheet className="w-10 h-10 mx-auto text-slate-300" />
              <p className="font-bold text-sm text-slate-700">Empty Sheet</p>
              <p className="text-xs text-slate-400">
                The sheet "{activeSheet}" contains no rows or columns.
              </p>
            </div>
          ) : (
            <div
              className="inline-block min-w-full align-middle"
              style={{ transform: `scale(${zoomLevel / 100})`, transformOrigin: 'top left' }}
            >
              <table className="min-w-full border-collapse border border-slate-200 bg-white text-xs select-text">
                {/* Table Header */}
                <thead className="sticky top-0 z-10 bg-slate-900 text-white shadow-xs">
                  <tr>
                    <th className="sticky left-0 z-20 w-12 min-w-[48px] px-2 py-2.5 bg-slate-950 border border-slate-800 text-center font-mono font-bold text-[10px] text-slate-400 select-none">
                      #
                    </th>
                    {currentSheetData.headers.map((header, colIdx) => (
                      <th
                        key={colIdx}
                        className="px-3.5 py-2.5 bg-slate-900 border border-slate-800 text-left font-extrabold text-xs text-slate-100 whitespace-nowrap min-w-[120px] max-w-[320px] truncate tracking-tight shadow-xs"
                        title={header}
                      >
                        <div className="flex items-center justify-between gap-1.5">
                          <span className="truncate">{header}</span>
                          <span className="text-[9px] font-mono text-slate-400 font-bold uppercase shrink-0">
                            {(() => {
                              let letter = '';
                              let n = colIdx;
                              while (n >= 0) {
                                letter = String.fromCharCode((n % 26) + 65) + letter;
                                n = Math.floor(n / 26) - 1;
                              }
                              return letter;
                            })()}
                          </span>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>

                {/* Table Body */}
                <tbody className="divide-y divide-slate-200 font-sans">
                  {paginatedRows.length === 0 ? (
                    <tr>
                      <td
                        colSpan={currentSheetData.headers.length + 1}
                        className="p-8 text-center text-slate-400 text-xs italic"
                      >
                        No matching rows found in this sheet for "{searchQuery}".
                      </td>
                    </tr>
                  ) : (
                    paginatedRows.map((row, rowIdx) => {
                      const actualRowNum =
                        (pageSize === 'all' ? 0 : currentPage * (pageSize as number)) + rowIdx + 1;
                      const isEven = rowIdx % 2 === 0;

                      return (
                        <tr
                          key={rowIdx}
                          className={`hover:bg-blue-50/70 transition-colors ${
                            isEven ? 'bg-white' : 'bg-slate-50/70'
                          }`}
                        >
                          <td className="sticky left-0 z-10 px-2 py-2 bg-slate-100 border-r border-slate-200 text-center font-mono text-[10px] font-bold text-slate-500 select-none">
                            {actualRowNum}
                          </td>
                          {currentSheetData.headers.map((_, colIdx) => {
                            const rawCell = row[colIdx];
                            const cellStr = rawCell !== null && rawCell !== undefined ? String(rawCell) : '';
                            const isSearchMatch =
                              searchQuery.trim() !== '' &&
                              cellStr.toLowerCase().includes(searchQuery.toLowerCase().trim());

                            return (
                              <td
                                key={colIdx}
                                className={`px-3 py-2 border-r border-b border-slate-200/80 text-slate-800 whitespace-nowrap max-w-[360px] truncate ${
                                  isSearchMatch ? 'bg-amber-100/90 font-bold text-amber-950' : ''
                                }`}
                                title={cellStr}
                              >
                                {cellStr || <span className="text-slate-300 select-none">—</span>}
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
        </div>

        {/* Bottom Status Bar */}
        <div className="px-5 py-2.5 bg-slate-50 border-t border-slate-200 text-xs text-slate-500 flex flex-wrap items-center justify-between gap-3 shrink-0">
          <div className="flex items-center gap-2">
            <span className="font-bold text-slate-700">Sheet:</span>
            <span className="px-2 py-0.5 rounded bg-white border border-slate-200 font-mono font-bold text-slate-800 text-[11px]">
              {activeSheet || 'Sheet1'}
            </span>
            {currentSheetData && (
              <span className="text-slate-500">
                • Showing {paginatedRows.length} of {filteredRows.length} rows (Total: {currentSheetData.totalRows.toLocaleString()} rows)
              </span>
            )}
          </div>

          <div className="flex items-center gap-3">
            <span className="text-slate-400 text-[11px]">
              Tip: Use the search bar to filter rows in real-time
            </span>
            <button
              type="button"
              onClick={onClose}
              className="px-3 py-1 rounded-lg bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold text-xs cursor-pointer transition"
            >
              Close
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
