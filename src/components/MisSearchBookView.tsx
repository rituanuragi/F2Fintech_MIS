import React, { useState, useEffect, useRef } from 'react';
import {
  BookOpen,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  Loader2,
  Download,
  Filter,
  Search,
  CheckSquare,
  Square,
  RefreshCw,
  Layers,
  ChevronRight,
  ExternalLink,
  Copy,
  Check,
  Sparkles,
  ArrowRight,
  Sliders,
  Eye,
  FileCheck,
  Table,
  Upload,
  Edit2,
} from 'lucide-react';
import { MisSearchBookSheetInfo, MisSearchBookRow, MisSearchBookSummary, LeadPriorityRule } from '../types';
import { PriorityConfigModal } from './PriorityConfigModal';
import { extractMonthYearLabel } from '../utils/dateHelper';

interface MisSearchBookViewProps {
  onViewRecord?: (id: number) => void;
  onPreviewSheet?: (fileId: string, sheetName: string) => void;
  onLookupCustomer?: (phone: string) => void;
}

export const MisSearchBookView: React.FC<MisSearchBookViewProps> = ({
  onViewRecord,
  onPreviewSheet,
  onLookupCustomer,
}) => {
  // Workbook & sheet selection states
  const [loadingWorkbook, setLoadingWorkbook] = useState(true);
  const [workbookInfo, setWorkbookInfo] = useState<{
    id: string;
    name: string;
    url: string;
    spreadsheetId: string;
    lastSyncedAt: string;
    totalSheets: number;
  } | null>(null);
  const [availableWorkbooks, setAvailableWorkbooks] = useState<
    Array<{ id: string; name: string; url: string; rowsCount: number; lastSyncedAt: string }>
  >([]);
  const [selectedWorkbookId, setSelectedWorkbookId] = useState<string>('');
  const [sheets, setSheets] = useState<MisSearchBookSheetInfo[]>([]);
  const [selectedSheetNames, setSelectedSheetNames] = useState<Set<string>>(new Set());
  const [sheetSearch, setSheetSearch] = useState('');
  const [filterWithPhoneOnly, setFilterWithPhoneOnly] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Processing & results states
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingStage, setProcessingStage] = useState<string>('');
  const [resultsSummary, setResultsSummary] = useState<MisSearchBookSummary | null>(null);
  const [sheetsData, setSheetsData] = useState<Record<string, MisSearchBookRow[]>>({});
  const [activeSheetTab, setActiveSheetTab] = useState<string>('');
  const [tableSearch, setTableSearch] = useState('');
  const [priorityFilter, setPriorityFilter] = useState<string>('all');
  const [statusFilter, setStatusFilter] = useState<'all' | 'found' | 'not_found'>('all');
  const [isExporting, setIsExporting] = useState(false);
  const [copiedPhone, setCopiedPhone] = useState<string | null>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 50;

  // Separate MIS Priority configuration states
  const [isPriorityModalOpen, setIsPriorityModalOpen] = useState(false);
  const [misPriorityRules, setMisPriorityRules] = useState<LeadPriorityRule[]>([]);
  const [isSyncingPriority, setIsSyncingPriority] = useState(false);
  const [syncPriorityMessage, setSyncPriorityMessage] = useState<string | null>(null);

  const handleSyncMisPriority = async () => {
    setIsSyncingPriority(true);
    setSyncPriorityMessage(null);
    try {
      const res = await fetch('/api/mis-priority-rules/sync', { method: 'POST' });
      const data = await res.json();
      if (data.rules) {
        setMisPriorityRules(data.rules);
      }
      setSyncPriorityMessage(
        data.addedCount > 0
          ? `+${data.addedCount} new added at end`
          : 'Priority Synced'
      );
      setTimeout(() => setSyncPriorityMessage(null), 3500);
    } catch (err) {
      console.error('Failed to sync MIS priority:', err);
    } finally {
      setIsSyncingPriority(false);
    }
  };

  // Pivot Summary & Combined Matrix states matching user's Image
  const [includeCombinedMatrix, setIncludeCombinedMatrix] = useState(true); // 2D Cross-sheet matrix matching user image
  const [matrixTitle, setMatrixTitle] = useState('Jan- Sept Logins');
  const [isEditingMatrixTitle, setIsEditingMatrixTitle] = useState(false);
  const [activeViewMode, setActiveViewMode] = useState<'matrix' | 'pivots'>('matrix'); // Default to matrix view!
  const [includePivotSheet, setIncludePivotSheet] = useState(true);
  const [pivotPerSheet, setPivotPerSheet] = useState(true); // Exactly 1 pivot for each selected sheet!
  const [useSheetNameAsHeader, setUseSheetNameAsHeader] = useState(true); // Pivot header = sheet name!
  const [includeDataSheets, setIncludeDataSheets] = useState(true);
  const [includeCombinedPivot, setIncludeCombinedPivot] = useState(false);
  const [activePivotSheetTab, setActivePivotSheetTab] = useState<string>('');

  // Real-time Auto-Sync Tracking
  const [autoSyncInfo, setAutoSyncInfo] = useState<{
    enabled: boolean;
    intervalSeconds: number;
    lastCheckedAt: string | null;
    lastNewRowsCount: number;
    lastChangedRowsCount: number;
    syncVersion: number;
  }>({
    enabled: true,
    intervalSeconds: 30,
    lastCheckedAt: null,
    lastNewRowsCount: 0,
    lastChangedRowsCount: 0,
    syncVersion: 1,
  });
  const lastSyncVersionRef = useRef<number>(1);
  const hasResultsRef = useRef<boolean>(false);
  useEffect(() => {
    hasResultsRef.current = !!resultsSummary;
  }, [resultsSummary]);

  // Fetch sheets for the MIS workbook
  const fetchWorkbookSheets = async (targetId?: string, forceRefresh: boolean = false, silent: boolean = false) => {
    if (!silent) {
      setLoadingWorkbook(true);
      setErrorMessage(null);
    }
    try {
      const params = new URLSearchParams();
      if (targetId) params.set('sheetId', targetId);
      if (forceRefresh) params.set('refresh', 'true');
      const url = `/api/mis-search-book/sheets${params.toString() ? `?${params.toString()}` : ''}`;
      const res = await fetch(url);
      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Failed to inspect Google Sheet MIS workbook');
      }

      setWorkbookInfo(data.workbookInfo);
      setAvailableWorkbooks(data.availableWorkbooks || []);
      if (!selectedWorkbookId && data.workbookInfo?.id) {
        setSelectedWorkbookId(data.workbookInfo.id);
      }
      setSheets(data.sheets || []);

      if (!silent) {
        // Default selection: select all sheets that have the PHONE NUMBER column
        const phoneSheets = (data.sheets || [])
          .filter((s: MisSearchBookSheetInfo) => s.hasPhoneColumn)
          .map((s: MisSearchBookSheetInfo) => s.name);

        if (phoneSheets.length > 0) {
          setSelectedSheetNames(new Set(phoneSheets));
        } else {
          setSelectedSheetNames(new Set((data.sheets || []).map((s: MisSearchBookSheetInfo) => s.name)));
        }
      }
    } catch (err: any) {
      if (!silent) {
        setErrorMessage(err.message || 'Error loading Google Sheet MIS');
      }
    } finally {
      if (!silent) {
        setLoadingWorkbook(false);
      }
    }
  };

  useEffect(() => {
    fetchWorkbookSheets();
    // Load separate MIS priority rules
    fetch('/api/mis-priority-rules')
      .then(res => res.json())
      .then(data => {
        if (data.rules) setMisPriorityRules(data.rules);
      })
      .catch(console.error);

    // Poll auto-sync status every 10s so when a new entry is added in the MIS Sheet, sheets & results auto-update!
    const poll = setInterval(async () => {
      try {
        const res = await fetch('/api/mis/auto-sync-status');
        if (!res.ok) return;
        const d = await res.json();
        if (d && d.success) {
          setAutoSyncInfo({
            enabled: d.enabled ?? true,
            intervalSeconds: d.intervalSeconds ?? 30,
            lastCheckedAt: d.lastCheckedAt ?? null,
            lastNewRowsCount: d.lastNewRowsCount ?? 0,
            lastChangedRowsCount: d.lastChangedRowsCount ?? 0,
            syncVersion: d.syncVersion ?? 1,
          });
          if (d.syncVersion && d.syncVersion > lastSyncVersionRef.current) {
            lastSyncVersionRef.current = d.syncVersion;
            fetchWorkbookSheets(selectedWorkbookId || undefined, false, true);
            if (hasResultsRef.current) {
              handleProceed(true);
            }
          }
        }
      } catch {}
    }, 10000);
    return () => clearInterval(poll);
  }, [selectedWorkbookId]);

  const handleWorkbookChange = (newSheetId: string) => {
    setSelectedWorkbookId(newSheetId);
    setResultsSummary(null);
    setSheetsData({});
    fetchWorkbookSheets(newSheetId);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    setErrorMessage(null);
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const base64 = (event.target?.result as string)?.split(',')[1];
        if (!base64) throw new Error('Failed to read file data');

        const res = await fetch('/api/mis-search-book/upload-workbook', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fileName: file.name,
            fileBase64: base64,
          }),
        });

        const data = await res.json();
        if (!res.ok || !data.success) {
          throw new Error(data.error || 'Failed to upload workbook');
        }

        setSelectedWorkbookId(data.sheetId);
        setResultsSummary(null);
        setSheetsData({});
        await fetchWorkbookSheets(data.sheetId);
      } catch (err: any) {
        setErrorMessage(err.message || 'Error processing uploaded file');
      } finally {
        setIsUploading(false);
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };
    reader.readAsDataURL(file);
  };

  // Toggle sheet selection
  const handleToggleSheet = (name: string) => {
    setSelectedSheetNames((prev) => {
      const next = new Set(prev);
      if (next.has(name)) {
        next.delete(name);
      } else {
        next.add(name);
      }
      return next;
    });
  };

  // Select all sheets that have entries
  const handleSelectAll = () => {
    setSelectedSheetNames(new Set(sheets.filter((s) => s.totalRows > 0).map((s) => s.name)));
  };

  // Deselect all sheets
  const handleDeselectAll = () => {
    setSelectedSheetNames(new Set());
  };

  // Select only sheets with PHONE NUMBER column and valid entries
  const handleSelectOnlyPhone = () => {
    const valid = sheets.filter((s) => s.hasPhoneColumn && s.totalRows > 0).map((s) => s.name);
    setSelectedSheetNames(new Set(valid));
  };

  // Copy phone number
  const handleCopy = (phone: string) => {
    navigator.clipboard.writeText(phone);
    setCopiedPhone(phone);
    setTimeout(() => setCopiedPhone(null), 1500);
  };

  // Calculate Pivot Table Summary data for the active sheet tab or all combined
  const pivotSummaryData = React.useMemo(() => {
    if (!sheetsData || Object.keys(sheetsData).length === 0) return null;
    const currentSheetName = activePivotSheetTab || activeSheetTab || Object.keys(sheetsData)[0];
    const isAllCombined = currentSheetName === 'ALL_COMBINED';

    const sheetsToAggregate = isAllCombined
      ? Object.keys(sheetsData)
      : (currentSheetName && sheetsData[currentSheetName] ? [currentSheetName] : Object.keys(sheetsData));

    const priorityMap: Record<string, number> = {};
    let blanksCount = 0;
    let totalCount = 0;

    sheetsToAggregate.forEach((sName) => {
      const rows = sheetsData[sName] || [];
      rows.forEach((r) => {
        totalCount++;
        const prio = (r.priority || '').trim();
        if (!r.found || !prio || prio === 'Not Found' || prio === '—') {
          blanksCount++;
        } else {
          priorityMap[prio] = (priorityMap[prio] || 0) + 1;
        }
      });
    });

    const items = Object.entries(priorityMap)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));

    // The header banner name is EXACTLY the sheet name!
    const bannerTitle = isAllCombined ? 'All Sheets Combined' : currentSheetName;

    return {
      sheetName: currentSheetName,
      bannerTitle,
      isAllCombined,
      items,
      blanksCount,
      grandTotal: totalCount,
    };
  }, [sheetsData, activePivotSheetTab, activeSheetTab]);

  // Calculate 2D Cross-Sheet Matrix Summary matching user's Image
  const matrixData = React.useMemo(() => {
    if (!sheetsData || Object.keys(sheetsData).length === 0) return null;
    const sheetList = Array.from(selectedSheetNames).filter((s) => sheetsData[s]);
    if (sheetList.length === 0) return null;

    const prioCountBySheet: Record<string, Record<string, number>> = {};
    const blanksBySheet: Record<string, number> = {};
    const totalBySheet: Record<string, number> = {};
    const allPrioritiesSet = new Set<string>();

    sheetList.forEach((sName) => {
      prioCountBySheet[sName] = {};
      blanksBySheet[sName] = 0;
      totalBySheet[sName] = 0;

      const rows = sheetsData[sName] || [];
      rows.forEach((r) => {
        totalBySheet[sName]++;
        const prio = (r.priority || '').trim();
        if (!r.found || !prio || prio === 'Not Found' || prio === '—') {
          blanksBySheet[sName]++;
        } else {
          prioCountBySheet[sName][prio] = (prioCountBySheet[sName][prio] || 0) + 1;
          allPrioritiesSet.add(prio);
        }
      });
    });

    const priorityTotals: Record<string, number> = {};
    allPrioritiesSet.forEach((prio) => {
      let sum = 0;
      sheetList.forEach((sName) => {
        sum += prioCountBySheet[sName][prio] || 0;
      });
      priorityTotals[prio] = sum;
    });

    const sortedPriorities = Array.from(allPrioritiesSet).sort((a, b) => {
      const diff = (priorityTotals[b] || 0) - (priorityTotals[a] || 0);
      return diff !== 0 ? diff : a.localeCompare(b);
    });

    let totalBlanks = 0;
    sheetList.forEach((sName) => {
      totalBlanks += blanksBySheet[sName] || 0;
    });

    let grandTotal = 0;
    sheetList.forEach((sName) => {
      grandTotal += totalBySheet[sName] || 0;
    });

    return {
      sheetList,
      prioCountBySheet,
      blanksBySheet,
      totalBySheet,
      sortedPriorities,
      priorityTotals,
      totalBlanks,
      grandTotal,
    };
  }, [sheetsData, selectedSheetNames]);

  // Execute Search Book generation
  const handleProceed = async (silentBackground: boolean = false) => {
    const selectedList = Array.from(selectedSheetNames);
    if (selectedList.length === 0) {
      if (!silentBackground) {
        setErrorMessage('Please select at least one sheet to proceed.');
      }
      return;
    }

    if (!silentBackground) {
      setIsProcessing(true);
      setErrorMessage(null);
      setProcessingStage('Reading selected sheets & extracting PHONE NUMBER column...');
    }

    try {
      const res = await fetch('/api/mis-search-book/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sheetId: selectedWorkbookId,
          selectedSheets: selectedList,
          priorityRules: misPriorityRules.length > 0 ? misPriorityRules : undefined,
        }),
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Failed to process selected sheets');
      }

      setResultsSummary(data.summary);
      setSheetsData(data.sheetsData || {});

      // Auto-set matrix title based on selected sheets if possible
      if (selectedList.length > 1) {
        const first = selectedList[0].split(' ')[0];
        const last = selectedList[selectedList.length - 1].split(' ')[0];
        if (first && last && first !== last) {
          setMatrixTitle(`${first}- ${last} Logins`);
        }
      }

      // Set active tab to first selected sheet with records
      const firstSheetWithData = selectedList.find(
        (name) => (data.sheetsData?.[name]?.length || 0) > 0
      );
      const chosenFirst = firstSheetWithData || selectedList[0] || '';
      setActiveSheetTab(chosenFirst);
      setActivePivotSheetTab(chosenFirst);
      setCurrentPage(1);
    } catch (err: any) {
      console.error('MIS Search Book process error:', err);
      setErrorMessage(err.message || 'Error executing search across sheets');
    } finally {
      setIsProcessing(false);
      setProcessingStage('');
    }
  };

  // Download the generated multi-sheet Excel file (.xlsx)
  const handleDownloadXlsx = async () => {
    if (!sheetsData || Object.keys(sheetsData).length === 0) return;
    setIsExporting(true);
    try {
      const res = await fetch('/api/mis-search-book/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sheetsData,
          selectedSheets: Array.from(selectedSheetNames),
          workbookName: workbookInfo?.name || 'MIS_Search_Book',
          includeCombinedMatrix, // 2D matrix matching image
          matrixTitle: matrixTitle.trim() || 'Jan- Sept Logins',
          includePivot: includePivotSheet,
          pivotPerSheet, // Exactly 1 pivot for each selected sheet
          useSheetNameAsHeader, // Header in each pivot = sheet name
          includeDataSheets,
          includeCombinedPivot,
        }),
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || 'Export failed');
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const dateStr = new Date().toISOString().slice(0, 10);
      const safeTitle = (workbookInfo?.name || 'MIS_Search_Book').replace(/[^a-zA-Z0-9_-]/g, '_');
      a.download = `${safeTitle}_Global_Search_${dateStr}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      console.error('Download error:', err);
      setErrorMessage(err.message || 'Failed to download Excel file');
    } finally {
      setIsExporting(false);
    }
  };

  // Filter sheets in selection grid: Only show worksheets that have entries (totalRows > 0)
  const filteredSheets = sheets.filter((s) => {
    if (s.totalRows === 0) {
      return false;
    }
    if (sheetSearch && !s.name.toLowerCase().includes(sheetSearch.toLowerCase())) {
      return false;
    }
    if (filterWithPhoneOnly && !s.hasPhoneColumn) {
      return false;
    }
    return true;
  });

  // Filter rows in active sheet table
  const activeRows = (activeSheetTab && sheetsData[activeSheetTab]) || [];
  const filteredActiveRows = activeRows.filter((row) => {
    if (statusFilter === 'found' && !row.found) return false;
    if (statusFilter === 'not_found' && row.found) return false;

    if (priorityFilter !== 'all') {
      if (priorityFilter === 'Not Found' && row.found) return false;
      if (priorityFilter !== 'Not Found' && !row.priority.toLowerCase().includes(priorityFilter.toLowerCase())) {
        return false;
      }
    }

    if (tableSearch) {
      const q = tableSearch.toLowerCase();
      const matchPhone = row.phone.toLowerCase().includes(q) || row.normPhone.includes(q);
      const matchDate = (row.date || '').toLowerCase().includes(q);
      const matchWb = row.workbook.toLowerCase().includes(q);
      const matchSheet = row.sheetName.toLowerCase().includes(q);
      const matchPrio = row.priority.toLowerCase().includes(q);
      const matchCust = (row.customerName || '').toLowerCase().includes(q);
      return matchPhone || matchDate || matchWb || matchSheet || matchPrio || matchCust;
    }
    return true;
  });

  const totalPages = Math.max(1, Math.ceil(filteredActiveRows.length / pageSize));
  const paginatedRows = filteredActiveRows.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  return (
    <div className="space-y-6">
      {/* View Header Banner */}
      <div className="bg-gradient-to-r from-blue-700 via-indigo-700 to-slate-900 rounded-2xl p-6 text-white shadow-md relative overflow-hidden">
        <div className="relative z-10 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 mb-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-mono font-bold bg-blue-500/30 text-blue-200 border border-blue-400/30 uppercase tracking-wider">
                MIS Sync • Global Multi-Workbook Search
              </span>
              <span className="px-2 py-0.5 rounded-full text-[10px] font-mono font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                Case-Sensitive PHONE NUMBER
              </span>
            </div>
            <h1 className="text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <BookOpen className="w-6 h-6 text-blue-300" />
              <span>MIS Search Book</span>
            </h1>
            <p className="text-xs text-blue-100/90 mt-1 max-w-3xl leading-relaxed">
              Scan all sheets from your synced Google Sheet MIS workbook, automatically fetch customer contact numbers from the exact{' '}
              <strong className="text-white underline decoration-amber-400 font-mono">PHONE NUMBER</strong> column, run a high-speed cross-match against all Drive workbooks, assign lead priorities, and generate a segregated multi-sheet Excel file.
            </p>
          </div>

          <div className="flex items-center gap-2 self-start md:self-auto shrink-0 flex-wrap">
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              className="hidden"
            />
            <button
              onClick={() => fileInputRef.current?.click()}
              disabled={isUploading || isProcessing}
              className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold flex items-center gap-2 transition-all cursor-pointer shadow-sm disabled:opacity-50"
            >
              {isUploading ? (
                <>
                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                  <span>Uploading File...</span>
                </>
              ) : (
                <>
                  <Upload className="w-3.5 h-3.5" />
                  <span>Upload MIS File (.xlsx)</span>
                </>
              )}
            </button>

            <button
              type="button"
              onClick={() => setIsPriorityModalOpen(true)}
              className="px-4 py-2.5 bg-purple-600 hover:bg-purple-500 text-white rounded-xl text-xs font-bold flex items-center gap-2 transition-all cursor-pointer shadow-sm"
              title="Select and configure priority hierarchy specifically for MIS Search Book"
            >
              <Sliders className="w-3.5 h-3.5" />
              <span>MIS Priority Rules</span>
              <span className="text-[10px] px-1.5 py-0.2 rounded bg-black/20 font-mono">
                {misPriorityRules.length > 0 ? `${misPriorityRules.length}` : 'Separate'}
              </span>
            </button>

            <button
              onClick={() => fetchWorkbookSheets(selectedWorkbookId, true)}
              disabled={loadingWorkbook || isProcessing || isUploading}
              className="px-4 py-2.5 bg-white/10 hover:bg-white/20 border border-white/20 rounded-xl text-xs font-bold text-white flex items-center gap-2 transition-all cursor-pointer disabled:opacity-50"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${loadingWorkbook ? 'animate-spin' : ''}`} />
              <span>Refresh Sheets</span>
            </button>
          </div>
        </div>
      </div>

      {/* Error alert if any */}
      {errorMessage && (
        <div className="bg-rose-50 border border-rose-200 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-rose-800 text-xs">
          <div className="flex items-start gap-2.5">
            <AlertCircle className="w-4 h-4 shrink-0 text-rose-600 mt-0.5" />
            <div>
              <span className="font-bold">Error:</span> {errorMessage}
              <div className="text-[11px] text-rose-600 mt-0.5">
                You can upload your MIS Excel file directly using the button below or on the right.
              </div>
            </div>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-1.5 bg-rose-600 hover:bg-rose-700 text-white rounded-lg font-bold flex items-center gap-1.5 transition-colors cursor-pointer"
            >
              <Upload className="w-3.5 h-3.5" />
              <span>Upload MIS File (.xlsx)</span>
            </button>
            <button
              onClick={() => setErrorMessage(null)}
              className="px-3 py-1.5 bg-white border border-rose-200 hover:bg-rose-100 text-rose-700 rounded-lg font-bold transition-colors cursor-pointer"
            >
              Dismiss
            </button>
          </div>
        </div>
      )}

      {/* SECTION 1: Synced Google Sheet MIS Workbook Details & Sheet Selector */}
      <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-xs space-y-6">
        {/* Workbook Info Top Bar */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-slate-100">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-10 h-10 rounded-xl bg-blue-50 border border-blue-200 flex items-center justify-center text-blue-600 shrink-0">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div className="min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-xs font-bold uppercase tracking-wider text-slate-400">
                  Active Synced MIS Workbook:
                </span>
                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 text-[10px] font-bold">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span> Live Auto-Sync Active (Every {autoSyncInfo.intervalSeconds}s)
                </span>
                {autoSyncInfo.lastCheckedAt && (
                  <span className="text-[10px] font-mono text-slate-500">
                    • Auto-checked {new Date(autoSyncInfo.lastCheckedAt).toLocaleTimeString()}
                  </span>
                )}
              </div>
              <h3 className="text-base font-extrabold text-slate-900 truncate">
                {workbookInfo?.name || 'Loading Google Sheet MIS...'}
              </h3>
              <p className="text-[11px] text-emerald-700 font-medium mt-0.5">
                New entries added in the MIS Sheet automatically sync &amp; update without manually clicking Sync.
              </p>
            </div>
          </div>

          {/* Switch workbook dropdown if multiple connected */}
          {availableWorkbooks.length > 1 && (
            <div className="flex items-center gap-2 shrink-0">
              <span className="text-xs font-medium text-slate-500">Switch Workbook:</span>
              <select
                value={selectedWorkbookId}
                onChange={(e) => handleWorkbookChange(e.target.value)}
                className="px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:outline-none focus:ring-1 focus:ring-blue-500"
              >
                {availableWorkbooks.map((wb) => (
                  <option key={wb.id} value={wb.id}>
                    {wb.name} ({wb.rowsCount?.toLocaleString() || 0} rows)
                  </option>
                ))}
              </select>
            </div>
          )}
        </div>

        {/* Sheet Selection Controls Bar */}
        <div className="space-y-3">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h4 className="text-sm font-black text-slate-900 flex items-center gap-2">
                <Layers className="w-4 h-4 text-blue-600" />
                <span>Select Worksheets to Include</span>
                <span className="px-2 py-0.5 rounded-full bg-blue-50 text-blue-700 border border-blue-200 text-[11px] font-mono font-bold">
                  {selectedSheetNames.size} of {sheets.length} selected
                </span>
              </h4>
              <p className="text-xs text-slate-500 mt-0.5">
                Check or uncheck individual sheets. For example, select <strong className="text-slate-700 font-mono">JAN 2026</strong> to scan its customer contacts and create a segregated <strong className="text-slate-700 font-mono">JAN 2026</strong> tab in the generated workbook.
              </p>
            </div>

            {/* Quick Bulk Selection Buttons */}
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={handleSelectAll}
                className="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold transition-colors flex items-center gap-1.5 cursor-pointer"
              >
                <CheckSquare className="w-3.5 h-3.5 text-blue-600" />
                <span>Select All</span>
              </button>

              <button
                type="button"
                onClick={handleDeselectAll}
                className="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold transition-colors flex items-center gap-1.5 cursor-pointer"
              >
                <Square className="w-3.5 h-3.5 text-slate-400" />
                <span>Deselect All</span>
              </button>

              <button
                type="button"
                onClick={handleSelectOnlyPhone}
                className="px-2.5 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 text-xs font-semibold transition-colors flex items-center gap-1.5 cursor-pointer"
                title="Select all sheets where column 'PHONE NUMBER' exists"
              >
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                <span>Only with PHONE NUMBER</span>
              </button>
            </div>
          </div>

          {/* Search filter for sheet names */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pt-1">
            <div className="relative flex-1 max-w-sm">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={sheetSearch}
                onChange={(e) => setSheetSearch(e.target.value)}
                placeholder="Filter sheets (e.g. JAN, 2026, OCT)..."
                className="w-full pl-9 pr-3 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-800 placeholder:text-slate-400 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </div>

            <label className="flex items-center gap-2 text-xs text-slate-600 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={filterWithPhoneOnly}
                onChange={(e) => setFilterWithPhoneOnly(e.target.checked)}
                className="rounded border-slate-300 text-blue-600 focus:ring-blue-500"
              />
              <span>Hide sheets without PHONE NUMBER column</span>
            </label>
          </div>
        </div>

        {/* Sheet Grid Card View */}
        {loadingWorkbook ? (
          <div className="p-8 text-center bg-slate-50 rounded-xl border border-slate-200 space-y-2">
            <Loader2 className="w-6 h-6 animate-spin text-blue-600 mx-auto" />
            <p className="text-xs font-medium text-slate-600">Inspecting sheets and detecting PHONE NUMBER columns...</p>
          </div>
        ) : filteredSheets.length === 0 ? (
          <div className="p-6 text-center bg-slate-50 rounded-xl border border-dashed border-slate-300 text-xs text-slate-500">
            No worksheets match your filter. Try adjusting search query.
          </div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-2.5 max-h-80 overflow-y-auto p-1 border border-slate-200/80 rounded-xl bg-slate-50/50">
            {filteredSheets.map((sheet) => {
              const isSelected = selectedSheetNames.has(sheet.name);
              return (
                <div
                  key={sheet.name}
                  onClick={() => handleToggleSheet(sheet.name)}
                  className={`p-3 rounded-xl border transition-all cursor-pointer flex flex-col justify-between select-none ${
                    isSelected
                      ? 'bg-blue-50/80 border-blue-300 shadow-xs ring-1 ring-blue-400/40'
                      : 'bg-white border-slate-200/80 hover:border-slate-300 hover:bg-slate-50/80 opacity-80'
                  }`}
                >
                  <div className="flex items-start justify-between gap-2">
                    <div className="flex items-center gap-2 min-w-0">
                      <div
                        className={`w-4 h-4 rounded flex items-center justify-center shrink-0 ${
                          isSelected ? 'bg-blue-600 text-white' : 'border border-slate-300 bg-white'
                        }`}
                      >
                        {isSelected && <Check className="w-3 h-3 stroke-[3]" />}
                      </div>
                      <span className="text-xs font-extrabold text-slate-900 truncate font-mono">
                        {sheet.name}
                      </span>
                    </div>

                    {sheet.hasPhoneColumn ? (
                      <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-200 shrink-0">
                        PHONE OK
                      </span>
                    ) : (
                      <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-500 border border-slate-200 shrink-0">
                        No Phone Col
                      </span>
                    )}
                  </div>

                  <div className="mt-2.5 flex items-center justify-between text-[11px] font-mono">
                    <span className="text-slate-700 font-bold">
                      {sheet.totalRows.toLocaleString()} rows
                      {sheet.totalCols && sheet.totalCols > 0 ? (
                        <span className="text-slate-400 font-normal text-[10px] ml-1">
                          ({sheet.totalCols} cols)
                        </span>
                      ) : null}
                    </span>
                    {sheet.hasPhoneColumn && (
                      <span
                        className="text-emerald-700 font-bold bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-200 text-[10px]"
                        title={`PHONE NUMBER column is located in Column ${sheet.phoneColLetter || ''} (Column #${sheet.phoneColIndex + 1})`}
                      >
                        Phone: Col {sheet.phoneColLetter ? `${sheet.phoneColLetter} ` : ''}(#{sheet.phoneColIndex + 1})
                      </span>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* Proceed Action Button & Priority Configuration */}
        <div className="pt-2 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div className="text-xs text-slate-500">
            Ready to extract phone numbers from{' '}
            <strong className="text-slate-900 font-bold">{selectedSheetNames.size} selected sheets</strong> and cross-match
            against the entire lead registry.
          </div>

          <div className="flex flex-wrap items-center gap-3 shrink-0">
            <button
              type="button"
              onClick={handleSyncMisPriority}
              disabled={isSyncingPriority}
              className="px-3.5 py-3 bg-white hover:bg-slate-50 text-slate-700 border border-slate-200 rounded-xl text-xs font-bold flex items-center justify-center gap-2 transition-all cursor-pointer shadow-xs disabled:opacity-50"
              title="Refresh and sync all directories & sheets into priority list (new directories automatically placed at the last)"
            >
              <RefreshCw className={`w-3.5 h-3.5 text-slate-600 ${isSyncingPriority ? 'animate-spin text-purple-600' : ''}`} />
              <span>{isSyncingPriority ? 'Syncing...' : syncPriorityMessage || 'Sync Priority'}</span>
            </button>

            <button
              type="button"
              onClick={() => setIsPriorityModalOpen(true)}
              className="px-4 py-3 bg-purple-50 hover:bg-purple-100 text-purple-700 border border-purple-200 rounded-xl text-xs font-bold flex items-center justify-center gap-2 transition-all cursor-pointer shadow-xs"
              title="Configure priority rules separately for this MIS Search Book"
            >
              <Sliders className="w-4 h-4 text-purple-600" />
              <span>MIS Priority Rules</span>
              <span className="text-[10px] px-1.5 py-0.5 rounded bg-purple-200 text-purple-800 font-mono font-bold">
                {misPriorityRules.length > 0 ? `${misPriorityRules.length} Rules` : 'Dedicated'}
              </span>
            </button>

            <button
              type="button"
              onClick={() => handleProceed(false)}
              disabled={isProcessing || selectedSheetNames.size === 0}
              className="px-6 py-3 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white rounded-xl text-sm font-bold shadow-md shadow-blue-500/25 flex items-center justify-center gap-2.5 transition-all cursor-pointer"
            >
              {isProcessing ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin" />
                  <span>Processing Global Multi-Workbook Search...</span>
                </>
              ) : (
                <>
                  <Sparkles className="w-4 h-4 text-amber-300" />
                  <span>Proceed &amp; Generate Global Search Sheet</span>
                  <ArrowRight className="w-4 h-4" />
                </>
              )}
            </button>
          </div>
        </div>
      </div>

      {/* Processing indicator card */}
      {isProcessing && (
        <div className="bg-white border border-blue-200 rounded-2xl p-8 text-center shadow-xs space-y-3">
          <Loader2 className="w-8 h-8 animate-spin text-blue-600 mx-auto" />
          <h3 className="text-base font-black text-slate-900">Scanning Customer Phone Numbers Across Sheets</h3>
          <p className="text-xs text-slate-600 max-w-md mx-auto">{processingStage || 'Processing...'}</p>
          <div className="text-[11px] font-mono text-slate-400">
            Cross-referencing against: Fresh Data 21k, Fresh Data 30k, ICICI Bank, LOGIN, NEET, LOT, Tata Points, Consolidated Directory...
          </div>
        </div>
      )}

      {/* SECTION 2: Results, KPIs & Download */}
      {resultsSummary && (
        <div className="space-y-6">
          {/* Summary KPIs & Download Bar */}
          <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-xs space-y-6">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-slate-100">
              <div>
                <div className="flex items-center gap-2">
                  <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 text-xs font-bold uppercase tracking-wider">
                    <CheckCircle2 className="w-3.5 h-3.5" /> Generation Complete
                  </span>
                  <span className="text-xs font-mono text-slate-400">
                    ({resultsSummary.durationMs}ms)
                  </span>
                </div>
                <h2 className="text-xl font-black text-slate-900 mt-1">
                  Global Search Sheets Successfully Created
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  The generated workbook contains segregated sheets exactly matching your MIS sheet names ({resultsSummary.totalSheetsProcessed} sheets processed).
                </p>
              </div>

              {/* Big Prominent Download Button & 1 Pivot Per Selected Sheet Options */}
              <div className="flex flex-col lg:flex-row lg:items-center gap-3">
                <div className="flex flex-wrap items-center gap-3 bg-slate-50 border border-slate-200 px-3.5 py-2 rounded-xl text-xs shadow-xs">
                  {/* Combined Summary Sheet (2D Cross-sheet Matrix matching user image) */}
                  <label
                    className={`flex items-center gap-2 cursor-pointer select-none px-2 py-0.5 rounded transition-all ${
                      includeCombinedMatrix
                        ? 'bg-blue-100 text-blue-950 border border-blue-300 font-bold'
                        : 'text-slate-700 hover:bg-slate-100 font-semibold'
                    }`}
                    title="Include 2D Cross-sheet summary matrix matching user reference image"
                  >
                    <input
                      type="checkbox"
                      checked={includeCombinedMatrix}
                      onChange={(e) => setIncludeCombinedMatrix(e.target.checked)}
                      className="rounded border-blue-300 text-blue-600 focus:ring-blue-500 w-4 h-4 cursor-pointer"
                    />
                    <span>Combined Sheet (Matrix)</span>
                  </label>

                  {/* Include Pivot Sheets */}
                  <label className="flex items-center gap-2 pl-3 border-l border-slate-200 cursor-pointer select-none text-slate-800 font-bold">
                    <input
                      type="checkbox"
                      checked={includePivotSheet}
                      onChange={(e) => setIncludePivotSheet(e.target.checked)}
                      className="rounded border-slate-300 text-purple-600 focus:ring-purple-500 w-4 h-4 cursor-pointer"
                    />
                    <span>Sheet Pivots</span>
                  </label>

                  {includePivotSheet && (
                    <>
                      {/* 1. Exactly 1 Pivot per Selected Sheet */}
                      <label
                        className={`flex items-center gap-1.5 pl-3 border-l border-slate-200 cursor-pointer select-none px-2 py-0.5 rounded transition-all ${
                          pivotPerSheet
                            ? 'bg-purple-100 text-purple-950 border border-purple-300 font-bold'
                            : 'text-purple-700 hover:bg-purple-50 font-semibold'
                        }`}
                        title="Generate 1 dedicated pivot sheet for each selected sheet"
                      >
                        <input
                          type="checkbox"
                          checked={pivotPerSheet}
                          onChange={(e) => setPivotPerSheet(e.target.checked)}
                          className="rounded border-purple-300 text-purple-600 focus:ring-purple-500 w-3.5 h-3.5 cursor-pointer"
                        />
                        <span className="text-[11px]">1 Pivot per Sheet</span>
                        <span className="text-[10px] bg-purple-200 text-purple-900 font-mono px-1.5 py-0.2 rounded font-bold">
                          {selectedSheetNames.size} Pivots
                        </span>
                      </label>

                      {/* 2. Use Sheet Name as Pivot Header */}
                      <label
                        className="flex items-center gap-1.5 pl-3 border-l border-slate-200 cursor-pointer select-none text-slate-700 hover:text-slate-900 font-medium"
                        title="Use the exact sheet name as the pivot header banner (e.g. 'JAN 2026')"
                      >
                        <input
                          type="checkbox"
                          checked={useSheetNameAsHeader}
                          onChange={(e) => setUseSheetNameAsHeader(e.target.checked)}
                          className="rounded border-slate-300 text-purple-600 focus:ring-purple-500 w-3.5 h-3.5 cursor-pointer"
                        />
                        <span className="text-[11px] font-semibold">Header = Sheet Name</span>
                      </label>
                    </>
                  )}

                  {/* 3. Include Data Sheets */}
                  <label
                    className="flex items-center gap-1.5 pl-3 border-l border-slate-200 cursor-pointer select-none text-slate-600 hover:text-slate-900 font-medium"
                    title="Include the raw data sheets alongside each pivot"
                  >
                    <input
                      type="checkbox"
                      checked={includeDataSheets}
                      onChange={(e) => setIncludeDataSheets(e.target.checked)}
                      className="rounded border-slate-300 text-blue-600 focus:ring-blue-500 w-3.5 h-3.5 cursor-pointer"
                    />
                    <span className="text-[11px]">Data Sheets</span>
                  </label>
                </div>

                <button
                  type="button"
                  onClick={handleDownloadXlsx}
                  disabled={isExporting}
                  className="px-6 py-3.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white rounded-xl text-sm font-black shadow-lg shadow-emerald-600/25 flex items-center justify-center gap-2.5 transition-all cursor-pointer shrink-0"
                >
                  {isExporting ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>Preparing Workbook...</span>
                    </>
                  ) : (
                    <>
                      <Download className="w-4 h-4 stroke-[2.5]" />
                      <span>Download MIS Search Book (.xlsx)</span>
                    </>
                  )}
                </button>
              </div>
            </div>

            {/* KPI Cards Grid */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
              <div className="bg-slate-50 rounded-xl p-4 border border-slate-200">
                <span className="text-slate-500 text-[11px] font-bold uppercase tracking-wider block font-sans">
                  Sheets Processed
                </span>
                <span className="text-2xl font-black text-slate-900 mt-1 block font-mono">
                  {resultsSummary.totalSheetsProcessed}
                </span>
                <span className="text-[11px] text-slate-400 mt-0.5 block font-mono">
                  Segregated as tabs
                </span>
              </div>

              <div className="bg-blue-50/70 rounded-xl p-4 border border-blue-200">
                <span className="text-blue-700 text-[11px] font-bold uppercase tracking-wider block font-sans">
                  Total Phones Scanned
                </span>
                <span className="text-2xl font-black text-blue-900 mt-1 block font-mono">
                  {resultsSummary.totalPhonesCount.toLocaleString()}
                </span>
                <span className="text-[11px] text-blue-600 mt-0.5 block font-mono">
                  from PHONE NUMBER column
                </span>
              </div>

              <div className="bg-emerald-50/70 rounded-xl p-4 border border-emerald-200">
                <span className="text-emerald-700 text-[11px] font-bold uppercase tracking-wider block font-sans">
                  Matched in Drive
                </span>
                <span className="text-2xl font-black text-emerald-900 mt-1 block font-mono">
                  {resultsSummary.totalMatchesCount.toLocaleString()}
                </span>
                <span className="text-[11px] text-emerald-700 font-bold mt-0.5 block font-mono">
                  {resultsSummary.totalPhonesCount > 0
                    ? `${Math.round((resultsSummary.totalMatchesCount / resultsSummary.totalPhonesCount) * 100)}% match rate`
                    : '0%'}
                </span>
              </div>

              <div className="bg-amber-50/70 rounded-xl p-4 border border-amber-200">
                <span className="text-amber-800 text-[11px] font-bold uppercase tracking-wider block font-sans">
                  Unmatched / Not Found
                </span>
                <span className="text-2xl font-black text-amber-900 mt-1 block font-mono">
                  {resultsSummary.totalNotFoundCount.toLocaleString()}
                </span>
                <span className="text-[11px] text-amber-700 mt-0.5 block font-mono">
                  Priority: Not Found
                </span>
              </div>
            </div>

            {/* Lead Priority Breakdown Pills */}
            <div className="pt-1">
              <span className="text-xs font-bold text-slate-500 uppercase tracking-wider block mb-2 font-mono">
                Global Priority Breakdown across selected sheets:
              </span>
              <div className="flex flex-wrap items-center gap-2">
                {Object.entries(resultsSummary.priorityBreakdown).map(([label, count]) => {
                  const isNotFound = label === 'Not Found';
                  return (
                    <span
                      key={label}
                      className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-lg text-xs font-bold font-mono border ${
                        isNotFound
                          ? 'bg-slate-100 text-slate-600 border-slate-200'
                          : 'bg-blue-50 text-blue-800 border-blue-200'
                      }`}
                    >
                      <span>{label}:</span>
                      <span className="font-extrabold">{count.toLocaleString()}</span>
                    </span>
                  );
                })}
              </div>
            </div>
          </div>

          {/* View Mode Toggle Bar: Combined Sheet Matrix vs Individual Sheet Pivots */}
          <div className="flex flex-wrap items-center justify-between gap-3 bg-white p-3.5 rounded-2xl border border-slate-200 shadow-xs">
            <div className="flex items-center gap-1.5 p-1 bg-slate-100 rounded-xl">
              <button
                type="button"
                onClick={() => setActiveViewMode('matrix')}
                className={`px-4 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-2 cursor-pointer ${
                  activeViewMode === 'matrix'
                    ? 'bg-blue-600 text-white shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Table className="w-4 h-4" />
                <span>Combined Sheet (Summary Matrix)</span>
                <span className={`text-[10px] px-2 py-0.5 rounded-full font-mono font-bold ${
                  activeViewMode === 'matrix' ? 'bg-white/20 text-white' : 'bg-slate-200 text-slate-700'
                }`}>
                  Image Format
                </span>
              </button>

              <button
                type="button"
                onClick={() => setActiveViewMode('pivots')}
                className={`px-4 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-2 cursor-pointer ${
                  activeViewMode === 'pivots'
                    ? 'bg-purple-900 text-white shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <FileSpreadsheet className="w-4 h-4" />
                <span>Individual Sheet Pivots</span>
                <span className={`text-[10px] px-2 py-0.5 rounded-full font-mono font-bold ${
                  activeViewMode === 'pivots' ? 'bg-white/20 text-white' : 'bg-slate-200 text-slate-700'
                }`}>
                  {selectedSheetNames.size} Pivots
                </span>
              </button>
            </div>

            <div className="flex items-center gap-3 text-xs text-slate-600 font-mono">
              <span className="flex items-center gap-1 text-emerald-700 font-bold">
                <CheckCircle2 className="w-3.5 h-3.5" /> Both formats included in Excel export
              </span>
            </div>
          </div>

          {/* VIEW 1: Combined Sheet (2D Cross-Sheet Summary Matrix) matching user's reference image */}
          {activeViewMode === 'matrix' && matrixData && (
            <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-xs space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-blue-50 border border-blue-200 flex items-center justify-center text-blue-700">
                    <Table className="w-4 h-4" />
                  </div>
                  <div>
                    <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                      Combined Sheet Preview
                      <span className="text-[10px] px-2 py-0.5 rounded-full bg-blue-100 text-blue-800 font-mono font-bold">
                        {matrixData.sheetList.length} Sheets Consolidated
                      </span>
                    </h3>
                    <p className="text-xs text-slate-500">
                      Consolidated 2D cross-tabulation matrix matching your reference format.
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-2 text-xs">
                  <span className="text-slate-500 font-medium">Header Title:</span>
                  {isEditingMatrixTitle ? (
                    <div className="flex items-center gap-1.5">
                      <input
                        type="text"
                        value={matrixTitle}
                        onChange={(e) => setMatrixTitle(e.target.value)}
                        className="px-2.5 py-1 text-xs border border-blue-400 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 font-bold text-slate-800"
                        autoFocus
                      />
                      <button
                        type="button"
                        onClick={() => setIsEditingMatrixTitle(false)}
                        className="px-2.5 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold transition-colors cursor-pointer"
                      >
                        Done
                      </button>
                    </div>
                  ) : (
                    <button
                      type="button"
                      onClick={() => setIsEditingMatrixTitle(true)}
                      className="flex items-center gap-1.5 px-3 py-1 bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold rounded-lg border border-slate-300 transition-colors cursor-pointer"
                      title="Click to customize title banner"
                    >
                      <span>"{matrixTitle}"</span>
                      <Edit2 className="w-3 h-3 text-slate-500" />
                    </button>
                  )}
                </div>
              </div>

              {/* The Authentic 2D Matrix Table matching the user's reference image */}
              <div className="overflow-x-auto rounded-xl border border-slate-300 shadow-xs max-h-[620px] overflow-y-auto">
                <table className="w-full text-xs border-collapse">
                  {/* Row 1: Merged Title Banner in Dark Navy #1B365D */}
                  <thead>
                    <tr>
                      <th
                        colSpan={matrixData.sheetList.length + 2}
                        className="bg-[#1B365D] text-white font-bold text-center py-2.5 px-4 text-sm tracking-wide border border-black/40"
                      >
                        {matrixTitle || 'Jan- Sept Logins'}
                      </th>
                    </tr>
                    {/* Row 2: Headers */}
                    <tr>
                      <th className="bg-[#FFF2CC] border border-slate-400/80 px-4 py-2 text-left font-bold text-slate-900 min-w-[200px]">
                        Row Labels
                      </th>
                      {matrixData.sheetList.map((sName) => (
                        <th
                          key={sName}
                          className="bg-[#CFDDF5] border border-slate-400/80 px-3 py-2 text-center font-bold text-slate-900 min-w-[90px] whitespace-nowrap"
                        >
                          {sName}
                        </th>
                      ))}
                      <th className="bg-[#FFE699] border border-slate-400/80 px-3 py-2 text-center font-black text-slate-950 min-w-[90px]">
                        Total
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {/* Priority Rows */}
                    {matrixData.sortedPriorities.map((prio) => (
                      <tr key={prio} className="hover:bg-amber-50/40 transition-colors">
                        <td className="bg-[#FFF2CC] border border-slate-400/60 px-4 py-1.5 font-medium text-slate-900 text-left">
                          {prio}
                        </td>
                        {matrixData.sheetList.map((sName) => {
                          const count = matrixData.prioCountBySheet[sName]?.[prio];
                          return (
                            <td
                              key={sName}
                              className="bg-white border border-slate-400/60 px-3 py-1.5 text-right font-mono text-slate-800"
                            >
                              {count && count > 0 ? count.toLocaleString() : ''}
                            </td>
                          );
                        })}
                        <td className="bg-[#FFF2CC] border border-slate-400/60 px-3 py-1.5 text-right font-bold font-mono text-[#C00000]">
                          {(matrixData.priorityTotals[prio] || 0).toLocaleString()}
                        </td>
                      </tr>
                    ))}

                    {/* (blank) row */}
                    <tr className="hover:bg-amber-50/40 transition-colors">
                      <td className="bg-[#FFF2CC] border border-slate-400/60 px-4 py-1.5 italic text-slate-700 text-left">
                        (blank)
                      </td>
                      {matrixData.sheetList.map((sName) => {
                        const count = matrixData.blanksBySheet[sName];
                        return (
                          <td
                            key={sName}
                            className="bg-white border border-slate-400/60 px-3 py-1.5 text-right font-mono text-slate-600"
                          >
                            {count && count > 0 ? count.toLocaleString() : ''}
                          </td>
                        );
                      })}
                      <td className="bg-[#FFF2CC] border border-slate-400/60 px-3 py-1.5 text-right font-bold font-mono text-[#C00000]">
                        {matrixData.totalBlanks.toLocaleString()}
                      </td>
                    </tr>

                    {/* Bottom Total Row in soft blue #BDD7EE */}
                    <tr className="font-bold border-t-2 border-slate-400">
                      <td className="bg-[#BDD7EE] border border-slate-400/80 px-4 py-2 text-slate-950 font-black text-left">
                        Total
                      </td>
                      {matrixData.sheetList.map((sName) => (
                        <td
                          key={sName}
                          className="bg-[#BDD7EE] border border-slate-400/80 px-3 py-2 text-right font-mono font-bold text-slate-900"
                        >
                          {(matrixData.totalBySheet[sName] || 0).toLocaleString()}
                        </td>
                      ))}
                      <td className="bg-[#BDD7EE] border border-slate-400/80 px-3 py-2 text-right font-mono font-black text-slate-950">
                        {matrixData.grandTotal.toLocaleString()}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* VIEW 2: Dedicated Pivot Table Summary Cards (1 Pivot per selected sheet) */}
          {activeViewMode === 'pivots' && pivotSummaryData && (
            <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-xs space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-purple-50 border border-purple-200 flex items-center justify-center text-purple-600">
                    <FileSpreadsheet className="w-4 h-4" />
                  </div>
                  <div>
                    <h3 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                      Pivot Table Summary Preview
                      <span className="text-[10px] px-2 py-0.5 rounded-full bg-purple-100 text-purple-800 font-mono font-bold">
                        {selectedSheetNames.size} Pivots to be Generated
                      </span>
                    </h3>
                    <p className="text-xs text-slate-500">
                      1 dedicated Pivot Sheet for each selected sheet. Header name matches the sheet name exactly.
                    </p>
                  </div>
                </div>

                {/* Tab Name indicator badge */}
                <div className="text-[11px] font-mono px-3 py-1 bg-purple-50 text-purple-800 border border-purple-200 rounded-lg">
                  Excel Tab:{' '}
                  <strong className="text-purple-950 font-extrabold">
                    {pivotSummaryData.isAllCombined
                      ? '"Combined Pivot"'
                      : `"${pivotSummaryData.sheetName} Pivot"`}
                  </strong>
                  {' • '}Header:{' '}
                  <strong className="text-purple-950 font-extrabold">
                    "{pivotSummaryData.bannerTitle}"
                  </strong>
                </div>
              </div>

              {/* Clean Sheet Tabs to switch between selected sheet pivots (No noisy date parsing) */}
              <div className="flex flex-wrap items-center gap-2 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
                <span className="text-xs font-bold text-slate-500 uppercase tracking-wider font-mono mr-1">
                  Preview Pivot For:
                </span>
                {Object.keys(sheetsData).map((sName) => {
                  const isCurrent = (activePivotSheetTab || activeSheetTab) === sName;
                  const rowCount = sheetsData[sName]?.length || 0;
                  return (
                    <button
                      key={sName}
                      type="button"
                      onClick={() => setActivePivotSheetTab(sName)}
                      className={`px-3 py-1.5 rounded-xl text-xs font-mono font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                        isCurrent
                          ? 'bg-purple-900 text-white shadow-xs ring-2 ring-purple-600'
                          : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-200'
                      }`}
                      title={`Preview pivot for ${sName}`}
                    >
                      <span>{sName}</span>
                      <span
                        className={`text-[10px] px-1.5 py-0.2 rounded font-bold ${
                          isCurrent ? 'bg-purple-800 text-purple-200' : 'bg-slate-100 text-slate-600'
                        }`}
                      >
                        {rowCount} rows
                      </span>
                    </button>
                  );
                })}
                {Object.keys(sheetsData).length > 1 && (
                  <button
                    type="button"
                    onClick={() => setActivePivotSheetTab('ALL_COMBINED')}
                    className={`px-3 py-1.5 rounded-xl text-xs font-mono font-bold transition-all cursor-pointer ${
                      activePivotSheetTab === 'ALL_COMBINED'
                        ? 'bg-purple-900 text-white shadow-xs ring-2 ring-purple-600'
                        : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-200'
                    }`}
                  >
                    All Sheets Combined ({resultsSummary.totalPhonesCount} rows)
                  </button>
                )}
              </div>

              {/* Authentic Excel-styled Pivot Table Card */}
              <div className="flex flex-wrap items-start gap-6">
                <div className="w-full max-w-sm rounded-xl overflow-hidden border-2 border-purple-900 shadow-md bg-white text-slate-900">
                  {/* Purple Header Banner matching Image 2 - Header Name IS The Sheet Name */}
                  <div className="bg-[#581C87] text-white text-center py-2.5 px-4 font-bold text-sm tracking-wide flex items-center justify-center gap-2">
                    <span>{pivotSummaryData.bannerTitle}</span>
                  </div>

                  {/* Data & Count Columns */}
                  <div className="grid grid-cols-2 bg-slate-100 border-b border-slate-300 text-xs font-bold text-slate-800 px-4 py-2 uppercase tracking-wider">
                    <div>Data</div>
                    <div className="text-right">Count</div>
                  </div>

                  {/* Priority Rows */}
                  <div className="divide-y divide-slate-200 text-xs font-medium">
                    {pivotSummaryData.items.map((item) => (
                      <div
                        key={item.name}
                        onClick={() => setPriorityFilter(priorityFilter === item.name ? 'all' : item.name)}
                        className="grid grid-cols-2 px-4 py-2 hover:bg-purple-50 transition-colors cursor-pointer group"
                        title={`Click to filter table by ${item.name}`}
                      >
                        <div className="truncate text-slate-800 font-semibold group-hover:text-purple-700">
                          {item.name}
                        </div>
                        <div className="text-right font-mono text-slate-700 font-bold">
                          {item.count.toLocaleString()}
                        </div>
                      </div>
                    ))}

                    {/* Blanks row (when found nothing) */}
                    <div
                      onClick={() => setStatusFilter(statusFilter === 'not_found' ? 'all' : 'not_found')}
                      className="grid grid-cols-2 px-4 py-2 hover:bg-slate-100 transition-colors cursor-pointer group bg-slate-50/60"
                      title="Click to filter by Not Found (Blanks)"
                    >
                      <div className="truncate text-slate-600 italic group-hover:text-slate-900">
                        Blanks
                      </div>
                      <div className="text-right font-mono text-slate-600 font-bold">
                        {pivotSummaryData.blanksCount.toLocaleString()}
                      </div>
                    </div>

                    {/* Grand Total Row */}
                    <div className="grid grid-cols-2 px-4 py-2.5 bg-[#EDE9FE] font-bold text-purple-950 text-xs border-t-2 border-purple-300">
                      <div>Grand Total</div>
                      <div className="text-right font-mono font-black text-purple-900">
                        {pivotSummaryData.grandTotal.toLocaleString()}
                      </div>
                    </div>
                  </div>
                </div>

                {/* Info side card with breakdown of pivots that will be generated */}
                <div className="flex-1 min-w-[280px] space-y-3 pt-1">
                  <div className="p-4 rounded-xl bg-purple-50 border border-purple-200 text-xs text-purple-900 space-y-3">
                    <div className="font-bold flex items-center gap-1.5 text-purple-800">
                      <Sparkles className="w-4 h-4 text-purple-600" />
                      <span>{selectedSheetNames.size} Dedicated Pivot Sheets in Downloaded Excel</span>
                    </div>
                    <p className="text-[12px] text-purple-950/80 leading-relaxed">
                      You selected <strong>{selectedSheetNames.size} sheets</strong>. The export will generate{' '}
                      <strong>{selectedSheetNames.size} separate Pivot sheets</strong> (1 Pivot for each sheet). Each pivot’s header banner is automatically named after its sheet:
                    </p>

                    <div className="bg-white/90 rounded-lg p-2.5 border border-purple-200 space-y-1.5 max-h-48 overflow-y-auto font-mono text-[11px]">
                      {Array.from(selectedSheetNames).map((name, idx) => (
                        <div key={name} className="flex items-center justify-between text-slate-700">
                          <span className="font-bold text-purple-950">
                            #{idx + 1}. Tab: "{name} Pivot"
                          </span>
                          <span className="text-purple-700 font-semibold">
                            Header: "{name}"
                          </span>
                        </div>
                      ))}
                    </div>

                    <div className="pt-1 flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => setIsPriorityModalOpen(true)}
                        className="px-3 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-xs font-bold inline-flex items-center gap-1.5 transition-colors cursor-pointer"
                      >
                        <Sliders className="w-3.5 h-3.5" />
                        <span>Customize MIS Priority Rules</span>
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* SECTION 3: Interactive On-Screen Tabbed Sheet Inspector */}
          <div className="bg-white border border-slate-200 rounded-2xl shadow-xs overflow-hidden">
            {/* Top Sheet Tabs Nav */}
            <div className="bg-slate-50/80 px-4 pt-4 border-b border-slate-200 flex items-center justify-between gap-3 overflow-x-auto">
              <div className="flex items-center gap-2 overflow-x-auto pb-3">
                <span className="text-xs font-bold text-slate-400 uppercase tracking-wider font-mono mr-1 shrink-0">
                  Sheets:
                </span>
                {Object.keys(sheetsData).map((sheetName) => {
                  const isActive = activeSheetTab === sheetName;
                  const rowCount = sheetsData[sheetName]?.length || 0;
                  const foundCount = resultsSummary.perSheetStats?.[sheetName]?.found || 0;
                  return (
                    <button
                      key={sheetName}
                      type="button"
                      onClick={() => {
                        setActiveSheetTab(sheetName);
                        setCurrentPage(1);
                      }}
                      className={`px-3.5 py-2 rounded-xl text-xs font-bold font-mono transition-all cursor-pointer shrink-0 flex items-center gap-2 ${
                        isActive
                          ? 'bg-blue-600 text-white shadow-sm'
                          : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-200'
                      }`}
                    >
                      <Table className={`w-3.5 h-3.5 ${isActive ? 'text-white' : 'text-slate-400'}`} />
                      <span>{sheetName}</span>
                      <span
                        className={`px-1.5 py-0.2 rounded-full text-[10px] font-mono ${
                          isActive ? 'bg-white/20 text-white' : 'bg-slate-100 text-slate-600'
                        }`}
                      >
                        {rowCount}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Inside Sheet Filters Bar */}
            <div className="p-4 border-b border-slate-100 bg-white flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="flex items-center gap-2 min-w-0">
                <span className="text-xs font-bold text-slate-900 font-mono">
                  Sheet: <span className="text-blue-600">{activeSheetTab}</span>
                </span>
                <span className="text-xs text-slate-400 font-mono">
                  ({filteredActiveRows.length} of {activeRows.length} rows)
                </span>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                {/* Search within sheet */}
                <div className="relative">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={tableSearch}
                    onChange={(e) => {
                      setTableSearch(e.target.value);
                      setCurrentPage(1);
                    }}
                    placeholder="Search phone, priority, workbook..."
                    className="pl-8 pr-3 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-800 placeholder:text-slate-400 focus:outline-none focus:ring-1 focus:ring-blue-500 w-56"
                  />
                </div>

                {/* Status filter: All / Matched / Not Found */}
                <select
                  value={statusFilter}
                  onChange={(e) => {
                    setStatusFilter(e.target.value as any);
                    setCurrentPage(1);
                  }}
                  className="px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:outline-none focus:ring-1 focus:ring-blue-500"
                >
                  <option value="all">All Records</option>
                  <option value="found">Matched Only</option>
                  <option value="not_found">Not Found Only</option>
                </select>
              </div>
            </div>

            {/* Segregated Table */}
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead>
                  <tr className="bg-slate-50/90 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider text-[11px] font-mono">
                    <th className="py-3 px-4 w-12 text-center">#</th>
                    <th className="py-3 px-4">PHONE NUMBER</th>
                    <th className="py-3 px-4">DATE</th>
                    <th className="py-3 px-4">PRIORITY</th>
                    <th className="py-3 px-4">WORKBOOK</th>
                    <th className="py-3 px-4">SHEET NAME</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 font-sans">
                  {paginatedRows.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="py-12 text-center text-slate-400 text-xs font-medium">
                        No records match the current filter in sheet "{activeSheetTab}".
                      </td>
                    </tr>
                  ) : (
                    paginatedRows.map((row, idx) => {
                      const absoluteIdx = (currentPage - 1) * pageSize + idx + 1;
                      return (
                        <tr key={idx} className="hover:bg-slate-50/70 transition-colors">
                          <td className="py-3 px-4 text-center font-mono text-[11px] text-slate-400">
                            {absoluteIdx}
                          </td>

                          {/* 1. PHONE NUMBER */}
                          <td className="py-3 px-4 font-mono font-bold text-slate-900">
                            <div className="flex items-center gap-2">
                              <span>{row.phone}</span>
                              <button
                                type="button"
                                onClick={() => handleCopy(row.phone)}
                                className="text-slate-400 hover:text-blue-600 transition-colors cursor-pointer"
                                title="Copy phone number"
                              >
                                {copiedPhone === row.phone ? (
                                  <Check className="w-3.5 h-3.5 text-emerald-600" />
                                ) : (
                                  <Copy className="w-3.5 h-3.5" />
                                )}
                              </button>
                            </div>
                          </td>

                          {/* 2. DATE (Case Created date) */}
                          <td className="py-3 px-4 font-mono text-slate-700 whitespace-nowrap">
                            {row.date && row.date !== '—' ? (
                              <span className="px-2 py-0.5 rounded bg-slate-100 text-slate-800 text-[11px] font-semibold border border-slate-200">
                                {row.date}
                              </span>
                            ) : (
                              <span className="text-slate-400 font-mono text-[11px]">—</span>
                            )}
                          </td>

                          {/* 3. PRIORITY */}
                          <td className="py-3 px-4">
                            {row.found ? (
                              <span
                                className={`inline-flex items-center px-2.5 py-1 rounded-md text-[11px] font-mono font-extrabold border ${
                                  row.badgeBg || 'bg-blue-50 text-blue-700 border-blue-200'
                                }`}
                              >
                                {row.priority}
                              </span>
                            ) : (
                              <span className="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-mono font-bold bg-slate-100 text-slate-500 border border-slate-200">
                                Not Found
                              </span>
                            )}
                          </td>

                          {/* 4. WORKBOOK */}
                          <td className="py-3 px-4 font-medium text-slate-800">
                            {row.found ? (
                              <div className="flex items-center gap-1.5 flex-wrap">
                                {row.workbook.split(', ').map((wbName, wIdx) => (
                                  <span
                                    key={wIdx}
                                    className="px-2 py-0.5 rounded bg-slate-100 text-slate-800 font-mono text-[11px] border border-slate-200 font-semibold"
                                  >
                                    {wbName}
                                  </span>
                                ))}
                              </div>
                            ) : (
                              <span className="text-slate-400 font-mono">—</span>
                            )}
                          </td>

                          {/* 4. SHEET NAME */}
                          <td className="py-3 px-4 font-medium text-slate-700">
                            {row.found ? (
                              <div className="flex items-center gap-1.5 flex-wrap">
                                {row.sheetName.split(', ').map((sName, sIdx) => (
                                  <span
                                    key={sIdx}
                                    className="px-2 py-0.5 rounded bg-indigo-50 text-indigo-700 font-mono text-[11px] border border-indigo-200"
                                  >
                                    {sName}
                                  </span>
                                ))}
                              </div>
                            ) : (
                              <span className="text-slate-400 font-mono">—</span>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls */}
            {totalPages > 1 && (
              <div className="p-4 bg-slate-50/70 border-t border-slate-100 flex items-center justify-between text-xs text-slate-600 font-mono">
                <div>
                  Showing {Math.min(filteredActiveRows.length, (currentPage - 1) * pageSize + 1)} to{' '}
                  {Math.min(filteredActiveRows.length, currentPage * pageSize)} of {filteredActiveRows.length} rows
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                    disabled={currentPage <= 1}
                    className="px-3 py-1 bg-white border border-slate-200 rounded-lg font-bold disabled:opacity-40 hover:bg-slate-100 cursor-pointer"
                  >
                    Previous
                  </button>
                  <span>
                    Page {currentPage} of {totalPages}
                  </span>
                  <button
                    type="button"
                    onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                    disabled={currentPage >= totalPages}
                    className="px-3 py-1 bg-white border border-slate-200 rounded-lg font-bold disabled:opacity-40 hover:bg-slate-100 cursor-pointer"
                  >
                    Next
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Dedicated Priority Rules Configuration Modal for MIS Search Book */}
      <PriorityConfigModal
        isOpen={isPriorityModalOpen}
        onClose={() => setIsPriorityModalOpen(false)}
        onRulesUpdated={(newRules) => {
          setMisPriorityRules(newRules);
          if (resultsSummary && Object.keys(sheetsData).length > 0) {
            handleProceed();
          }
        }}
        scope="mis"
        title="MIS Search Book Dedicated Priority Rules"
        description="Select and configure priority hierarchy specifically for this MIS Search Book. When a lead matches multiple files, the top-ranked category in this list is chosen."
      />
    </div>
  );
};
