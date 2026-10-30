import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  CalendarDays,
  FileSpreadsheet,
  Layers,
  Search,
  Filter,
  Download,
  RefreshCw,
  ExternalLink,
  Copy,
  Check,
  ChevronLeft,
  ChevronRight,
  TrendingUp,
  Building2,
  DollarSign,
  UserCheck,
  CheckCircle2,
  Clock,
  AlertCircle,
  Eye,
  SlidersHorizontal,
  ChevronDown,
  ChevronUp,
  Sparkles,
  PhoneCall,
  Calendar,
  UploadCloud,
  ShieldAlert,
  Info,
  Zap,
} from 'lucide-react';
import {
  MisDashboardSheet,
  MisDashboardEntry,
  MisDashboardDateSummary,
  MisDashboardSummary,
} from '../types';
import { MisPivotPanel } from './MisPivotPanel';

interface MisDashboardViewProps {
  onViewRecord?: (id: number | string) => void;
  onLookupCustomer?: (phone: string) => void;
  onNavigateToTicket?: (ticketId: string | number) => void;
}

// Format numbers into Indian currency style (e.g. ₹ 45,00,000 or ₹ 45.0 L)
function formatIndianCurrency(amount: number): string {
  if (!amount || isNaN(amount)) return '₹ 0';
  if (amount >= 10000000) {
    return `₹ ${(amount / 10000000).toFixed(2)} Cr`;
  }
  if (amount >= 100000) {
    return `₹ ${(amount / 100000).toFixed(2)} L`;
  }
  return `₹ ${amount.toLocaleString('en-IN')}`;
}

// Today's date on THIS computer as yyyy-mm-dd (not UTC, so it is right after midnight in India)
function todayIso(): string {
  const d = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${d.getFullYear()}-${mm}-${dd}`;
}

// Find the worksheet tab of the current month, e.g. "OCT 26", "OCT 2026", "October 2026"
function pickCurrentMonthTab(tabs: { name: string }[] | undefined): string | null {
  if (!tabs || tabs.length === 0) return null;
  const now = new Date();
  const mon = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'][now.getMonth()];
  const yy = String(now.getFullYear()).slice(-2);
  const yyyy = String(now.getFullYear());
  const hits = tabs.filter((t) => t.name.trim().toUpperCase().startsWith(mon));
  const exact = hits.find((t) => {
    const n = t.name.toUpperCase();
    return new RegExp(`\\b(${yyyy}|${yy})\\b`).test(n);
  });
  return (exact || null)?.name ?? null;
}

export const MisDashboardView: React.FC<MisDashboardViewProps> = ({
  onViewRecord,
  onLookupCustomer,
  onNavigateToTicket,
}) => {
  // 1. Sheet selection state
  const [sheets, setSheets] = useState<MisDashboardSheet[]>([]);
  const [selectedSheetId, setSelectedSheetId] = useState<string>('mis_master_mis');
  const [selectedTab, setSelectedTab] = useState<string>('OCT 26'); // Default to current month tab
  const [isLoadingSheets, setIsLoadingSheets] = useState(false);

  // 2. Date column selection
  const [dateColumns, setDateColumns] = useState<string[]>(['CASE CREATED DATE']);
  const [selectedDateColumn, setSelectedDateColumn] = useState<string>('CASE CREATED DATE');

  // 3. Date Selection (defaults to 10 Oct 2026 as per user requirement)
  const [selectedDate, setSelectedDate] = useState<string>(todayIso());
  const [availableDates, setAvailableDates] = useState<MisDashboardDateSummary[]>([]);
  const [isLoadingDates, setIsLoadingDates] = useState(false);

  // 4. Query results & KPIs
  const [entries, setEntries] = useState<MisDashboardEntry[]>([]);
  const [summary, setSummary] = useState<MisDashboardSummary | null>(null);
  const [isLoadingEntries, setIsLoadingEntries] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // 5. Filters & Search
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [lenderFilter, setLenderFilter] = useState('all');
  const [sortBy, setSortBy] = useState<'row' | 'name' | 'amount' | 'status'>('row');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('asc');

  // 6. UI Interactions
  const [expandedRowId, setExpandedRowId] = useState<number | null>(null);
  const [copiedPhone, setCopiedPhone] = useState<string | null>(null);
  const [isExporting, setIsExporting] = useState(false);
  const [isSyncingSheet, setIsSyncingSheet] = useState(false);
  const [syncMessage, setSyncMessage] = useState<string | null>(null);
  const [showAllDatesDrawer, setShowAllDatesDrawer] = useState(false);
  const [showPivot, setShowPivot] = useState(false);

  // 7. Auto-Update state
  const [isAutoUpdate, setIsAutoUpdate] = useState<boolean>(true);
  const [autoUpdateInterval, setAutoUpdateInterval] = useState<number>(30); // in seconds
  const [countdown, setCountdown] = useState<number>(30);
  const [lastUpdatedTime, setLastUpdatedTime] = useState<string>(
    new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
  );
  const [isSilentUpdating, setIsSilentUpdating] = useState<boolean>(false);

  // Upload and Private Sheet Warning states
  const [isUploadingOct, setIsUploadingOct] = useState(false);
  const [privateSheetWarning, setPrivateSheetWarning] = useState<string | null>(null);
  const [uploadSuccessMessage, setUploadSuccessMessage] = useState<string | null>(null);
  const octFileInputRef = useRef<HTMLInputElement | null>(null);

  // Load available sheets on mount
  useEffect(() => {
    fetchSheets();
  }, []);

  const fetchSheets = async (isBackground = false) => {
    if (!isBackground) setIsLoadingSheets(true);
    try {
      const res = await fetch('/api/mis-dashboard/sheets');
      const data = await res.json();
      if (data.sheets && data.sheets.length > 0) {
        setSheets(data.sheets);
        // If sheet is not yet selected or we want to update tabs
        if (!selectedSheetId) {
          const master = data.sheets.find(
            (s: MisDashboardSheet) => s.id === 'mis_master_mis' || /master\s*mis/i.test(s.name)
          );
          const initialId = master ? master.id : data.sheets[0].id;
          setSelectedSheetId(initialId);

          // Open the tab of the current month; if there is none, the first tab
          const targetSheet = master || data.sheets[0];
          const currentTab = pickCurrentMonthTab(targetSheet.tabs);
          if (currentTab) {
            setSelectedTab(currentTab);
          } else if (targetSheet.tabs && targetSheet.tabs.length > 0) {
            setSelectedTab(targetSheet.tabs[0].name);
          } else {
            setSelectedTab('all');
          }
        }
      }
    } catch (err: any) {
      if (!isBackground) {
        console.error('Failed to load MIS sheets:', err);
        setErrorMessage('Could not load MIS Google Sheets. Please check network connection.');
      }
    } finally {
      if (!isBackground) setIsLoadingSheets(false);
    }
  };

  // Fetch detected date columns when sheet or tab changes
  useEffect(() => {
    if (!selectedSheetId) return;
    const fetchDateColumns = async () => {
      try {
        const res = await fetch(`/api/mis-dashboard/date-columns?sheetId=${encodeURIComponent(selectedSheetId)}&tabName=${encodeURIComponent(selectedTab)}`);
        const data = await res.json();
        if (data.columns && data.columns.length > 0) {
          setDateColumns(data.columns);
          if (data.defaultColumn && !selectedDateColumn) {
            setSelectedDateColumn(data.defaultColumn);
          }
        }
      } catch (err) {
        console.warn('Failed to fetch date columns:', err);
      }
    };
    fetchDateColumns();
  }, [selectedSheetId, selectedTab]);

  // Fetch available dates for the selected sheet, tab & date column
  useEffect(() => {
    if (!selectedSheetId) return;
    const fetchAvailableDates = async () => {
      setIsLoadingDates(true);
      try {
        const res = await fetch(
          `/api/mis-dashboard/available-dates?sheetId=${encodeURIComponent(selectedSheetId)}&tabName=${encodeURIComponent(selectedTab)}&dateColumn=${encodeURIComponent(selectedDateColumn)}`
        );
        const data = await res.json();
        if (data.availableDates) {
          setAvailableDates(data.availableDates);
        }
      } catch (err) {
        console.warn('Failed to fetch available dates:', err);
      } finally {
        setIsLoadingDates(false);
      }
    };
    fetchAvailableDates();
  }, [selectedSheetId, selectedTab, selectedDateColumn]);

  // Query entries whenever selectedDate, sheet, tab, or dateColumn changes
  useEffect(() => {
    if (!selectedDate || !selectedSheetId) return;
    executeDateQuery();
  }, [selectedDate, selectedSheetId, selectedTab, selectedDateColumn]);

  const executeDateQuery = async (isBackground = false) => {
    if (isBackground) {
      setIsSilentUpdating(true);
    } else {
      setIsLoadingEntries(true);
    }
    setErrorMessage(null);
    try {
      const res = await fetch('/api/mis-dashboard/query', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sheetId: selectedSheetId,
          tabName: selectedTab,
          date: selectedDate,
          dateColumn: selectedDateColumn,
          searchQuery,
          statusFilter,
          lenderFilter,
        }),
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to query MIS dashboard entries');
      }

      setEntries(data.entries || []);
      setSummary(data.summary || null);
      if (data.availableDates && data.availableDates.length > 0) {
        setAvailableDates(data.availableDates);
      }
      setLastUpdatedTime(
        new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
      );
    } catch (err: any) {
      if (!isBackground) {
        console.error('Error fetching date cases:', err);
        setErrorMessage(err.message || 'Failed to retrieve cases for the selected date');
        setEntries([]);
        setSummary(null);
      }
    } finally {
      if (isBackground) {
        setIsSilentUpdating(false);
      } else {
        setIsLoadingEntries(false);
      }
    }
  };

  // Auto-Update interval countdown and background trigger with live sheet polling
  useEffect(() => {
    if (!isAutoUpdate) return;

    const timer = setInterval(() => {
      setCountdown((prev) => {
        if (prev <= 1) {
          // Trigger background auto-sync and query
          (async () => {
            setIsSilentUpdating(true);
            try {
              const syncRes = await fetch('/api/mis-dashboard/sync-sheet', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ sheetId: selectedSheetId }),
              });
              const syncData = await syncRes.json();
              if (syncData.isPrivateGSheet) {
                setPrivateSheetWarning(syncData.warning);
              } else if (syncData.isLive) {
                setPrivateSheetWarning(null);
              }
              await fetchSheets(true);
              await executeDateQuery(true);
            } catch {
              await executeDateQuery(true);
            } finally {
              setIsSilentUpdating(false);
            }
          })();
          return autoUpdateInterval;
        }
        return prev - 1;
      });
    }, 1000);

    return () => clearInterval(timer);
  }, [
    isAutoUpdate,
    autoUpdateInterval,
    selectedDate,
    selectedSheetId,
    selectedTab,
    selectedDateColumn,
    searchQuery,
    statusFilter,
    lenderFilter,
  ]);

  // Reset countdown when user changes interval or changes key filters
  useEffect(() => {
    setCountdown(autoUpdateInterval);
  }, [autoUpdateInterval, selectedDate, selectedSheetId, selectedTab, selectedDateColumn]);

  // Sync / Refresh Google Sheet
  const handleSyncSheet = async () => {
    setIsSyncingSheet(true);
    setSyncMessage('Syncing MIS Sheet & recalculating row count...');
    try {
      const res = await fetch('/api/mis-dashboard/sync-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheetId: selectedSheetId }),
      });
      const data = await res.json();
      if (data.isPrivateGSheet) {
        setPrivateSheetWarning(data.warning);
      } else {
        setPrivateSheetWarning(null);
      }
      setSyncMessage(data.message || 'Sheet refreshed successfully!');
      setTimeout(() => setSyncMessage(null), 5000);
      await fetchSheets();
      await executeDateQuery();
    } catch (err: any) {
      setSyncMessage(`Sync failed: ${err.message || 'Network error'}`);
      setTimeout(() => setSyncMessage(null), 5000);
    } finally {
      setIsSyncingSheet(false);
    }
  };

  // Directly upload or update the OCT file (.xlsx or .csv)
  const handleUploadOctFile = async (file: File) => {
    if (!file) return;
    setIsUploadingOct(true);
    setErrorMessage(null);
    setUploadSuccessMessage(null);

    try {
      const base64 = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as string);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });

      const res = await fetch('/api/mis-dashboard/upload-oct-file', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: file.name,
          fileBase64: base64,
          sheetId: selectedSheetId || 'mis_master_mis',
          targetTab: selectedTab.toUpperCase().includes('OCT') ? selectedTab : 'OCT 26',
        }),
      });

      const data = await res.json();
      if (!res.ok) {
        throw new Error(data.error || 'Failed to upload and update OCT file');
      }

      setUploadSuccessMessage(data.message);
      setPrivateSheetWarning(null);
      setTimeout(() => setUploadSuccessMessage(null), 8000);

      await fetchSheets();
      if (data.activeTab) {
        setSelectedTab(data.activeTab);
      }
      await executeDateQuery();
    } catch (err: any) {
      setErrorMessage(`Upload failed: ${err.message}`);
    } finally {
      setIsUploadingOct(false);
      if (octFileInputRef.current) octFileInputRef.current.value = '';
    }
  };

  // Export to Excel file
  const handleExportExcel = async () => {
    setIsExporting(true);
    try {
      const res = await fetch('/api/mis-dashboard/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sheetId: selectedSheetId,
          tabName: selectedTab,
          date: selectedDate,
          dateColumn: selectedDateColumn,
          searchQuery,
          statusFilter,
          lenderFilter,
        }),
      });

      if (!res.ok) {
        throw new Error('Export request failed');
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `MIS_Cases_${selectedDate}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      console.error('Export failed:', err);
      alert(`Export failed: ${err.message || 'Unknown error'}`);
    } finally {
      setIsExporting(false);
    }
  };

  // Export client-side CSV
  const handleExportCsv = () => {
    if (entries.length === 0) return;
    const headers = [
      'Ticket ID',
      'Case Created Date',
      'Customer Name',
      'Phone Number',
      'Lender',
      'Status',
      'Loan Amount',
      'Net Loan Amount',
      'TL Name',
      'Manager',
      'City',
      'Branch',
      'Last Updated Date',
      'Disbursed Date',
    ];
    // Safe CSV cell: always quoted, quotes doubled, and cells that start with
    // = + - @ get a leading apostrophe so Excel cannot run them as formulas.
    const csvCell = (v: any) => {
      let t = String(v ?? '');
      if (/^[=+\-@\t\r]/.test(t)) t = `'${t}`;
      return `"${t.replace(/"/g, '""')}"`;
    };
    const rows = filteredAndSortedEntries.map((e) =>
      [
        e.ticketId,
        e.caseCreatedDate,
        e.customerName,
        e.phone,
        e.lender,
        e.status,
        e.loanAmount,
        e.netLoanAmount,
        e.tlName,
        e.manager,
        e.city,
        e.branch,
        e.lastUpdatedDate,
        e.disbursedDate,
      ].map(csvCell)
    );

    // Blob download (a data: URI cuts the file at the first "#", e.g. in "#N/A")
    const csvText = '\uFEFF' + [headers.map(csvCell).join(','), ...rows.map((r) => r.join(','))].join('\r\n');
    const blobUrl = URL.createObjectURL(new Blob([csvText], { type: 'text/csv;charset=utf-8;' }));
    const link = document.createElement('a');
    link.setAttribute('href', blobUrl);
    link.setAttribute('download', `MIS_Cases_${selectedDate}.csv`);
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 1000);
  };

  // Copy phone helper
  const handleCopyPhone = (phone: string, e: React.MouseEvent) => {
    e.stopPropagation();
    navigator.clipboard.writeText(phone);
    setCopiedPhone(phone);
    setTimeout(() => setCopiedPhone(null), 2000);
  };

  // Copy all phones for the date
  const handleCopyAllPhones = () => {
    const phones = filteredAndSortedEntries
      .map((e) => e.phone)
      .filter(Boolean)
      .filter((v, i, a) => a.indexOf(v) === i);
    if (phones.length === 0) return;
    navigator.clipboard.writeText(phones.join('\n'));
    alert(`Copied ${phones.length} unique phone number(s) to clipboard!`);
  };

  // Navigate date by 1 day
  const handleDateOffset = (deltaDays: number) => {
    try {
      const [y, m, d] = selectedDate.split('-').map(Number);
      const curr = new Date(Date.UTC(y, m - 1, d));
      curr.setUTCDate(curr.getUTCDate() + deltaDays);
      const newIso = curr.toISOString().split('T')[0];
      setSelectedDate(newIso);
    } catch {
      // Fallback
    }
  };

  // Format the selected date for human display
  const humanReadableDate = useMemo(() => {
    try {
      const [y, m, d] = selectedDate.split('-').map(Number);
      const dateObj = new Date(Date.UTC(y, m - 1, d));
      return dateObj.toLocaleDateString('en-GB', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric',
      });
    } catch {
      return selectedDate;
    }
  }, [selectedDate]);

  // Distinct lenders and statuses in current entries for filter dropdowns
  const distinctLenders = useMemo(() => {
    const set = new Set<string>();
    entries.forEach((e) => {
      if (e.lender) set.add(e.lender.toUpperCase());
    });
    return Array.from(set).sort();
  }, [entries]);

  const distinctStatuses = useMemo(() => {
    const set = new Set<string>();
    entries.forEach((e) => {
      if (e.status) set.add(e.status.toUpperCase());
    });
    return Array.from(set).sort();
  }, [entries]);

  // Client-side filtering & sorting for instant feel
  const filteredAndSortedEntries = useMemo(() => {
    let result = entries;

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      result = result.filter((e) => {
        return (
          e.customerName.toLowerCase().includes(q) ||
          e.phone.toLowerCase().includes(q) ||
          e.ticketId.toLowerCase().includes(q) ||
          e.lender.toLowerCase().includes(q) ||
          e.status.toLowerCase().includes(q) ||
          e.tlName.toLowerCase().includes(q) ||
          e.city.toLowerCase().includes(q) ||
          e.caseCreatedDate.toLowerCase().includes(q)
        );
      });
    }

    if (statusFilter !== 'all') {
      result = result.filter((e) => e.status.toUpperCase() === statusFilter.toUpperCase());
    }

    if (lenderFilter !== 'all') {
      result = result.filter((e) => e.lender.toUpperCase() === lenderFilter.toUpperCase());
    }

    // Sort
    result = [...result].sort((a, b) => {
      let comp = 0;
      if (sortBy === 'row') comp = a.rowNum - b.rowNum;
      else if (sortBy === 'name') comp = a.customerName.localeCompare(b.customerName);
      else if (sortBy === 'status') comp = a.status.localeCompare(b.status);
      else if (sortBy === 'amount') {
        const numA = parseFloat(a.loanAmount.replace(/[^0-9.]/g, '')) || 0;
        const numB = parseFloat(b.loanAmount.replace(/[^0-9.]/g, '')) || 0;
        comp = numA - numB;
      }
      return sortOrder === 'asc' ? comp : -comp;
    });

    return result;
  }, [entries, searchQuery, statusFilter, lenderFilter, sortBy, sortOrder]);

  const currentSheet = useMemo(() => {
    return sheets.find((s) => s.id === selectedSheetId) || sheets[0] || null;
  }, [sheets, selectedSheetId]);

  return (
    <div className="flex flex-col h-full bg-slate-50 overflow-y-auto">
      {/* Top Header & Context Banner */}
      <div className="bg-white border-b border-slate-200 px-6 py-4 shadow-xs sticky top-0 z-20">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2">
              <div className="p-2 bg-emerald-50 text-emerald-700 rounded-lg border border-emerald-200 shadow-xs">
                <CalendarDays className="w-5 h-5" />
              </div>
              <h1 className="text-xl font-bold text-slate-900 tracking-tight">
                MIS Dashboard — Daily Case Tracker
              </h1>
              <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-300">
                Live Google Sheets
              </span>
            </div>
            <p className="text-xs text-slate-500 mt-1">
              Select any connected MIS Google Sheet, pick a calendar date, and instantly inspect all case entries matching the Case Created Date.
            </p>
          </div>

          {/* Auto-Update & Quick sync Controls */}
          <div className="flex flex-wrap items-center gap-2">
            {syncMessage && (
              <span className="text-xs font-medium text-emerald-700 bg-emerald-50 border border-emerald-200 px-2.5 py-1 rounded-md animate-fade-in">
                {syncMessage}
              </span>
            )}

            {/* Live Auto-Update Pill */}
            <div className="inline-flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-md px-2 py-1 shadow-2xs">
              <button
                type="button"
                onClick={() => setIsAutoUpdate(!isAutoUpdate)}
                className={`inline-flex items-center gap-1.5 text-xs font-semibold px-2 py-0.5 rounded transition-all cursor-pointer ${
                  isAutoUpdate
                    ? 'bg-emerald-100 text-emerald-800 hover:bg-emerald-200'
                    : 'bg-slate-200 text-slate-600 hover:bg-slate-300'
                }`}
                title={isAutoUpdate ? 'Click to pause auto-update' : 'Click to resume auto-update'}
              >
                <span
                  className={`w-2 h-2 rounded-full ${
                    isAutoUpdate
                      ? isSilentUpdating
                        ? 'bg-emerald-600 animate-ping'
                        : 'bg-emerald-500 animate-pulse'
                      : 'bg-slate-400'
                  }`}
                />
                <span>{isAutoUpdate ? 'Auto-Update ON' : 'Auto-Update Paused'}</span>
              </button>
              {isAutoUpdate && (
                <span className="text-[11px] text-slate-500 font-medium whitespace-nowrap">
                  {isSilentUpdating ? 'Refreshing...' : `in ${countdown}s`}
                </span>
              )}
              <select
                value={autoUpdateInterval}
                onChange={(e) => {
                  const val = Number(e.target.value);
                  setAutoUpdateInterval(val);
                  setCountdown(val);
                }}
                disabled={!isAutoUpdate}
                className="text-[11px] bg-white border border-slate-300 rounded px-1.5 py-0.5 text-slate-700 cursor-pointer focus:outline-hidden"
                title="Select Auto-Update interval"
              >
                <option value={15}>15s</option>
                <option value={30}>30s</option>
                <option value={60}>60s</option>
                <option value={120}>2m</option>
              </select>
            </div>

            {/* Manual Immediate Refresh */}
            <button
              onClick={() => executeDateQuery(false)}
              disabled={isLoadingEntries}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors shadow-2xs disabled:opacity-50 cursor-pointer"
              title={`Last updated at ${lastUpdatedTime}. Click to refresh immediately.`}
            >
              <RefreshCw className={`w-3.5 h-3.5 text-slate-500 ${isLoadingEntries ? 'animate-spin' : ''}`} />
              <span className="hidden sm:inline">Refresh</span>
            </button>

            {/* Sync Live Sheet */}
            <button
              onClick={handleSyncSheet}
              disabled={isSyncingSheet}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors shadow-2xs disabled:opacity-50 cursor-pointer"
              title="Refresh / Resync Google Sheet"
            >
              <RefreshCw className={`w-3.5 h-3.5 text-slate-500 ${isSyncingSheet ? 'animate-spin' : ''}`} />
              <span>{isSyncingSheet ? 'Syncing...' : 'Sync Live Sheet'}</span>
            </button>

            {/* Upload / Update OCT File */}
            <input
              ref={octFileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) handleUploadOctFile(file);
              }}
            />
            <button
              type="button"
              onClick={() => octFileInputRef.current?.click()}
              disabled={isUploadingOct}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 transition-colors shadow-2xs disabled:opacity-50 cursor-pointer"
              title="Upload or update the OCT file (.xlsx or .csv) with latest ~100 rows"
            >
              <UploadCloud className={`w-3.5 h-3.5 ${isUploadingOct ? 'animate-bounce' : ''}`} />
              <span>{isUploadingOct ? 'Updating Rows...' : 'Upload / Update OCT File'}</span>
            </button>

            {/* Export to Excel */}
            <button
              onClick={handleExportExcel}
              disabled={isExporting || entries.length === 0}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-white bg-emerald-600 rounded-md hover:bg-emerald-700 transition-colors shadow-xs disabled:opacity-50 cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>{isExporting ? 'Exporting...' : 'Export to Excel'}</span>
            </button>
          </div>
        </div>

        {/* Banners for Private Google Sheet Notice or Upload Success */}
        {uploadSuccessMessage && (
          <div className="mt-3 p-3 bg-emerald-50 border border-emerald-200 rounded-xl text-emerald-900 text-xs flex items-center justify-between gap-2 shadow-2xs animate-fade-in">
            <div className="flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
              <span className="font-semibold">{uploadSuccessMessage}</span>
            </div>
            <button
              onClick={() => setUploadSuccessMessage(null)}
              className="text-emerald-500 hover:text-emerald-800 text-xs cursor-pointer font-bold px-1"
            >
              ✕
            </button>
          </div>
        )}

        {privateSheetWarning && (
          <div className="mt-3 p-3.5 bg-amber-50/90 border border-amber-200 rounded-xl text-amber-900 text-xs flex items-start justify-between gap-3 shadow-2xs animate-fade-in">
            <div className="flex items-start gap-2.5">
              <ShieldAlert className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
              <div>
                <p className="font-bold text-amber-950">
                  Live Cloud Auto-Sync Notice: Google Sheet Access Restricted (HTTP 401)
                </p>
                <p className="text-[11px] text-amber-800 mt-0.5 leading-relaxed">
                  Google returned HTTP 401 Unauthorized for the connected sheet link. To allow the server to automatically download the latest ~100 rows on every 30s auto-sync cycle:
                  in Google Sheets, click <strong>Share</strong> &rarr; change General Access to <strong>&ldquo;Anyone with the link can view&rdquo;</strong>.
                </p>
                <div className="mt-2 flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() => octFileInputRef.current?.click()}
                    className="px-2.5 py-1 bg-amber-600 hover:bg-amber-700 text-white font-bold rounded-lg text-[11px] transition cursor-pointer shadow-2xs flex items-center gap-1"
                  >
                    <UploadCloud className="w-3 h-3" />
                    <span>Upload Latest OCT File (100+ Rows)</span>
                  </button>
                  {currentSheet?.url && !currentSheet.url.startsWith('local://') && (
                    <a
                      href={currentSheet.url}
                      target="_blank"
                      rel="noreferrer"
                      className="px-2.5 py-1 bg-white hover:bg-amber-100 text-amber-800 border border-amber-300 font-medium rounded-lg text-[11px] transition inline-flex items-center gap-1"
                    >
                      <ExternalLink className="w-3 h-3" />
                      <span>Open Sheet Share Settings</span>
                    </a>
                  )}
                </div>
              </div>
            </div>
            <button
              onClick={() => setPrivateSheetWarning(null)}
              className="text-amber-500 hover:text-amber-800 text-xs font-bold cursor-pointer px-1"
              title="Dismiss notice"
            >
              ✕
            </button>
          </div>
        )}

        {/* Control Bar: Sheet Selector, Tab Selector, Date Column, and Interactive Calendar */}
        <div className="mt-4 pt-4 border-t border-slate-100 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-12 gap-3 items-end">
          {/* 1. Google Sheet Selector (Cols 4) */}
          <div className="lg:col-span-4">
            <label className="block text-xs font-semibold text-slate-700 mb-1 flex items-center justify-between">
              <span className="flex items-center gap-1">
                <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
                Select MIS Google Sheet
              </span>
              {currentSheet && (
                <span className="text-[11px] font-normal text-slate-400">
                  {currentSheet.rowsCount?.toLocaleString()} total rows
                </span>
              )}
            </label>
            <div className="relative">
              <select
                value={selectedSheetId}
                onChange={(e) => {
                  setSelectedSheetId(e.target.value);
                  const chosen = sheets.find((s) => s.id === e.target.value);
                  if (chosen && chosen.tabs && chosen.tabs.length > 0) {
                    const oct = chosen.tabs.find((t) => t.name.toUpperCase().includes('OCT'));
                    setSelectedTab(oct ? oct.name : chosen.tabs[0].name);
                  } else {
                    setSelectedTab('all');
                  }
                }}
                disabled={isLoadingSheets}
                className="w-full pl-3 pr-8 py-2 bg-white border border-slate-300 rounded-lg text-xs font-medium text-slate-800 shadow-2xs focus:outline-hidden focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all appearance-none cursor-pointer"
              >
                {sheets.map((s) => (
                  <option key={s.id} value={s.id}>
                    {s.name} {s.rowsCount ? `(${s.rowsCount.toLocaleString()} rows)` : ''}
                  </option>
                ))}
              </select>
              <ChevronDown className="w-4 h-4 text-slate-400 absolute right-2.5 top-2.5 pointer-events-none" />
            </div>
          </div>

          {/* 2. Worksheet / Tab Selector (Cols 3) */}
          <div className="lg:col-span-3">
            <label className="block text-xs font-semibold text-slate-700 mb-1 flex items-center justify-between">
              <span className="flex items-center gap-1">
                <Layers className="w-3.5 h-3.5 text-indigo-600" />
                Worksheet Tab
              </span>
              <span className="text-[11px] font-normal text-slate-400">
                {currentSheet?.tabs?.length || 0} tabs
              </span>
            </label>
            <div className="relative">
              <select
                value={selectedTab}
                onChange={(e) => setSelectedTab(e.target.value)}
                className="w-full pl-3 pr-8 py-2 bg-white border border-slate-300 rounded-lg text-xs font-medium text-slate-800 shadow-2xs focus:outline-hidden focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all appearance-none cursor-pointer"
              >
                <option value="all">★ All Worksheet Tabs (Whole Sheet)</option>
                {currentSheet?.tabs?.map((t) => (
                  <option key={t.name} value={t.name}>
                    {t.name} {t.rows ? `(${t.rows} rows)` : ''}
                  </option>
                ))}
              </select>
              <ChevronDown className="w-4 h-4 text-slate-400 absolute right-2.5 top-2.5 pointer-events-none" />
            </div>
          </div>

          {/* 3. Date Column Filter (Cols 2) */}
          <div className="lg:col-span-2">
            <label className="block text-xs font-semibold text-slate-700 mb-1 flex items-center justify-between">
              <span className="flex items-center gap-1">
                <Filter className="w-3.5 h-3.5 text-amber-600" />
                Date Column
              </span>
              <span className="text-[10px] bg-amber-50 text-amber-700 px-1 py-0.2 rounded font-mono">
                Auto
              </span>
            </label>
            <div className="relative">
              <select
                value={selectedDateColumn}
                onChange={(e) => setSelectedDateColumn(e.target.value)}
                className="w-full pl-2.5 pr-7 py-2 bg-white border border-slate-300 rounded-lg text-xs font-medium text-slate-800 shadow-2xs focus:outline-hidden focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all appearance-none cursor-pointer truncate"
              >
                {dateColumns.map((col) => (
                  <option key={col} value={col}>
                    {col}
                  </option>
                ))}
              </select>
              <ChevronDown className="w-4 h-4 text-slate-400 absolute right-2 top-2.5 pointer-events-none" />
            </div>
          </div>

          {/* 4. Interactive Date Calendar Picker (Cols 3) */}
          <div className="lg:col-span-3">
            <label className="block text-xs font-semibold text-slate-700 mb-1 flex items-center justify-between">
              <span className="flex items-center gap-1 text-emerald-800 font-bold">
                <Calendar className="w-3.5 h-3.5 text-emerald-600" />
                Select Calendar Date
              </span>
              <span className="text-[11px] font-semibold text-emerald-600 bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-200">
                {entries.length} cases found
              </span>
            </label>
            <div className="flex items-center gap-1">
              <button
                type="button"
                onClick={() => handleDateOffset(-1)}
                className="p-2 bg-white border border-slate-300 rounded-lg hover:bg-slate-100 text-slate-600 transition-colors shadow-2xs"
                title="Previous Day"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => {
                  if (e.target.value) setSelectedDate(e.target.value);
                }}
                className="w-full px-2.5 py-2 bg-emerald-50/50 border border-emerald-300 rounded-lg text-xs font-bold text-emerald-950 shadow-2xs focus:outline-hidden focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all cursor-pointer"
              />
              <button
                type="button"
                onClick={() => handleDateOffset(1)}
                className="p-2 bg-white border border-slate-300 rounded-lg hover:bg-slate-100 text-slate-600 transition-colors shadow-2xs"
                title="Next Day"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>

        {/* Date Presets and Quick-Select Pills */}
        <div className="mt-3 pt-2.5 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
          <div className="flex flex-wrap items-center gap-1.5 text-xs">
            <span className="text-slate-400 font-medium text-[11px] flex items-center gap-1 mr-1">
              <Clock className="w-3 h-3 text-slate-400" /> Quick Dates:
            </span>

            {/* Quick dates come from the real data (newest first), not hardcoded numbers */}
            {[...availableDates]
              .sort((x, y) => (x.date < y.date ? 1 : -1))
              .slice(0, 5)
              .map((d) => (
                <button
                  key={d.date}
                  onClick={() => setSelectedDate(d.date)}
                  className={`px-2.5 py-1 rounded-md text-xs font-semibold transition-all border ${
                    selectedDate === d.date
                      ? 'bg-emerald-600 text-white border-emerald-700 shadow-xs'
                      : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-50 hover:border-slate-300'
                  }`}
                >
                  {d.displayDate}
                  <span className={`ml-1.5 px-1 py-0.2 rounded text-[10px] ${
                    selectedDate === d.date ? 'bg-emerald-800 text-emerald-100' : 'bg-slate-100 text-slate-600'
                  }`}>
                    {d.count} cases
                  </span>
                </button>
              ))}
            {availableDates.length === 0 && (
              <span className="text-[11px] text-slate-400">No dates loaded yet</span>
            )}

            {availableDates.length > 5 && (
              <button
                onClick={() => setShowAllDatesDrawer(!showAllDatesDrawer)}
                className="px-2 py-1 rounded-md text-xs font-medium text-indigo-700 bg-indigo-50 border border-indigo-200 hover:bg-indigo-100 transition-colors flex items-center gap-1"
              >
                <span>More Dates ({availableDates.length})</span>
                {showAllDatesDrawer ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />}
              </button>
            )}
          </div>

          {/* Universal date format normalization indicator */}
          <div className="flex items-center gap-1 text-[11px] text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200 font-medium">
            <CheckCircle2 className="w-3 h-3 text-emerald-600" />
            <span>Format Flexible: Matches 10-October-26, 10-OCTOBER-2026, 10/10/2026 & serials</span>
          </div>
        </div>

        {/* Collapsible drawer showing ALL available dates with case counts */}
        {showAllDatesDrawer && (
          <div className="mt-3 p-3 bg-slate-50 rounded-lg border border-slate-200 max-h-40 overflow-y-auto">
            <div className="text-[11px] font-semibold text-slate-600 mb-2 flex items-center justify-between">
              <span>All Active Dates with Case Entries in {selectedTab === 'all' ? 'All Tabs' : selectedTab}:</span>
              <span className="text-slate-400 font-normal">{availableDates.length} distinct dates</span>
            </div>
            <div className="flex flex-wrap gap-1.5">
              {availableDates.map((d) => (
                <button
                  key={d.date}
                  onClick={() => {
                    setSelectedDate(d.date);
                    setShowAllDatesDrawer(false);
                  }}
                  className={`px-2 py-0.5 rounded text-[11px] font-medium transition-colors border ${
                    selectedDate === d.date
                      ? 'bg-emerald-600 text-white border-emerald-700'
                      : 'bg-white text-slate-700 border-slate-200 hover:border-emerald-300 hover:bg-emerald-50'
                  }`}
                >
                  {d.displayDate} ({d.count})
                </button>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Main Content Area */}
      <div className="p-6 space-y-5 flex-1">
        {/* Date Title Banner */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
          <div>
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold uppercase tracking-wider text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                Selected Date Entry
              </span>
              <span className="text-sm font-semibold text-slate-500">
                Column: <code className="text-slate-800 font-bold bg-slate-100 px-1 py-0.5 rounded">{selectedDateColumn}</code>
              </span>
            </div>
            <h2 className="text-xl font-bold text-slate-900 mt-1">
              {humanReadableDate}
            </h2>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={() => setShowPivot((v) => !v)}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border rounded-lg transition-colors shadow-2xs ${showPivot ? 'text-white bg-violet-600 border-violet-600' : 'text-slate-700 bg-white border-slate-300 hover:bg-slate-50'}`}
              title="Pivot of this date's entries by Priority (each number counted once)"
            >
              <Layers className="w-3.5 h-3.5" />
              <span>{showPivot ? 'Hide Pivot' : 'Pivot by Priority'}</span>
            </button>
            <button
              onClick={handleCopyAllPhones}
              disabled={filteredAndSortedEntries.length === 0}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors shadow-2xs disabled:opacity-50"
              title="Copy all customer phone numbers from this date"
            >
              <Copy className="w-3.5 h-3.5 text-slate-500" />
              <span>Copy All Phones</span>
            </button>
            <button
              onClick={handleExportCsv}
              disabled={filteredAndSortedEntries.length === 0}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors shadow-2xs disabled:opacity-50"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" />
              <span>Export CSV</span>
            </button>
          </div>
        </div>

        {showPivot && selectedDate && selectedSheetId && (
          <MisPivotPanel
            sheetId={selectedSheetId}
            tabName={selectedTab}
            date={selectedDate}
            dateColumn={selectedDateColumn}
          />
        )}

        {/* KPI / Metric Summary Cards for the Selected Date */}
        <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-3">
          {/* 1. Total Cases */}
          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs relative overflow-hidden">
            <div className="flex items-center justify-between">
              <span className="text-xs font-medium text-slate-500">Cases Created</span>
              <div className="p-1.5 bg-emerald-50 text-emerald-600 rounded-lg">
                <CalendarDays className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-1">
              <span className="text-2xl font-extrabold text-slate-900">
                {summary?.totalCases || entries.length}
              </span>
              <span className="text-xs text-slate-400 font-medium">cases</span>
            </div>
            <div className="mt-1 text-[11px] text-emerald-700 font-medium">
              on {selectedDate}
            </div>
          </div>

          {/* 2. Total Loan Amount */}
          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs relative overflow-hidden">
            <div className="flex items-center justify-between">
              <span className="text-xs font-medium text-slate-500">Total Loan Amount</span>
              <div className="p-1.5 bg-blue-50 text-blue-600 rounded-lg">
                <DollarSign className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-1">
              <span className="text-xl font-bold text-slate-900">
                {summary ? formatIndianCurrency(summary.totalLoanAmount) : '₹ 0'}
              </span>
            </div>
            <div className="mt-1 text-[11px] text-slate-500">
              Net: {summary ? formatIndianCurrency(summary.totalNetLoanAmount) : '₹ 0'}
            </div>
          </div>

          {/* 3. Approved Cases */}
          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs relative overflow-hidden">
            <div className="flex items-center justify-between">
              <span className="text-xs font-medium text-slate-500">Approved</span>
              <div className="p-1.5 bg-emerald-50 text-emerald-600 rounded-lg">
                <CheckCircle2 className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-1">
              <span className="text-2xl font-extrabold text-emerald-700">
                {summary?.statusBreakdown?.['APPROVED'] || 0}
              </span>
              <span className="text-xs text-slate-400">cases</span>
            </div>
            <div className="mt-1 text-[11px] text-slate-400">
              Disbursed: {summary?.statusBreakdown?.['DISBURSED'] || 0}
            </div>
          </div>

          {/* 4. Login / In Progress */}
          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs relative overflow-hidden">
            <div className="flex items-center justify-between">
              <span className="text-xs font-medium text-slate-500">Login / Pipeline</span>
              <div className="p-1.5 bg-amber-50 text-amber-600 rounded-lg">
                <Clock className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-1">
              <span className="text-2xl font-extrabold text-amber-700">
                {summary?.statusBreakdown?.['LOGIN'] || 0}
              </span>
              <span className="text-xs text-slate-400">cases</span>
            </div>
            <div className="mt-1 text-[11px] text-slate-400">
              Rejected: {summary?.statusBreakdown?.['REJECTED'] || 0}
            </div>
          </div>

          {/* 5. Lenders Involved */}
          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs relative overflow-hidden">
            <div className="flex items-center justify-between">
              <span className="text-xs font-medium text-slate-500">Active Lenders</span>
              <div className="p-1.5 bg-indigo-50 text-indigo-600 rounded-lg">
                <Building2 className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-1">
              <span className="text-2xl font-extrabold text-slate-900">
                {summary ? Object.keys(summary.lenderBreakdown).length : 0}
              </span>
              <span className="text-xs text-slate-400">banks</span>
            </div>
            <div className="mt-1 text-[11px] text-slate-500 truncate" title={distinctLenders.join(', ')}>
              {distinctLenders.slice(0, 2).join(', ')}
            </div>
          </div>

          {/* 6. Unique Customers */}
          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs relative overflow-hidden">
            <div className="flex items-center justify-between">
              <span className="text-xs font-medium text-slate-500">Unique Applicants</span>
              <div className="p-1.5 bg-purple-50 text-purple-600 rounded-lg">
                <UserCheck className="w-4 h-4" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-1">
              <span className="text-2xl font-extrabold text-slate-900">
                {summary?.uniqueCustomers || entries.length}
              </span>
              <span className="text-xs text-slate-400">leads</span>
            </div>
            <div className="mt-1 text-[11px] text-slate-400">
              {summary?.uniquePhones || 0} verified phones
            </div>
          </div>
        </div>

        {/* Search, Filter & Controls Strip */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs space-y-3">
          <div className="flex flex-col md:flex-row items-center justify-between gap-3">
            {/* Search Input */}
            <div className="relative w-full md:w-80">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5 pointer-events-none" />
              <input
                type="text"
                placeholder="Search by customer, phone, ticket, lender, TL..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-9 pr-3 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded-lg text-slate-800 placeholder-slate-400 focus:outline-hidden focus:ring-2 focus:ring-emerald-500 focus:bg-white transition-all"
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2.5 top-2 text-xs text-slate-400 hover:text-slate-600"
                >
                  ✕
                </button>
              )}
            </div>

            {/* Filter controls */}
            <div className="flex flex-wrap items-center gap-2 w-full md:w-auto">
              {/* Status Filter */}
              <div className="flex items-center gap-1.5 text-xs text-slate-600">
                <span className="font-medium text-slate-500 text-[11px]">Status:</span>
                <select
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                  className="px-2 py-1 bg-slate-50 border border-slate-300 rounded-md text-xs font-medium text-slate-700 cursor-pointer focus:outline-hidden focus:ring-1 focus:ring-emerald-500"
                >
                  <option value="all">All Statuses ({entries.length})</option>
                  {distinctStatuses.map((st) => (
                    <option key={st} value={st}>
                      {st} ({entries.filter((e) => e.status.toUpperCase() === st).length})
                    </option>
                  ))}
                </select>
              </div>

              {/* Lender Filter */}
              <div className="flex items-center gap-1.5 text-xs text-slate-600">
                <span className="font-medium text-slate-500 text-[11px]">Lender:</span>
                <select
                  value={lenderFilter}
                  onChange={(e) => setLenderFilter(e.target.value)}
                  className="px-2 py-1 bg-slate-50 border border-slate-300 rounded-md text-xs font-medium text-slate-700 cursor-pointer focus:outline-hidden focus:ring-1 focus:ring-emerald-500"
                >
                  <option value="all">All Lenders ({entries.length})</option>
                  {distinctLenders.map((l) => (
                    <option key={l} value={l}>
                      {l} ({entries.filter((e) => e.lender.toUpperCase() === l).length})
                    </option>
                  ))}
                </select>
              </div>

              {/* Sort By */}
              <div className="flex items-center gap-1.5 text-xs text-slate-600">
                <span className="font-medium text-slate-500 text-[11px]">Sort:</span>
                <select
                  value={sortBy}
                  onChange={(e) => setSortBy(e.target.value as any)}
                  className="px-2 py-1 bg-slate-50 border border-slate-300 rounded-md text-xs font-medium text-slate-700 cursor-pointer focus:outline-hidden focus:ring-1 focus:ring-emerald-500"
                >
                  <option value="row">Original Row #</option>
                  <option value="name">Customer Name</option>
                  <option value="amount">Loan Amount</option>
                  <option value="status">Status</option>
                </select>
                <button
                  onClick={() => setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc')}
                  className="p-1 border border-slate-300 rounded hover:bg-slate-100 text-slate-600"
                  title="Toggle Asc/Desc"
                >
                  {sortOrder === 'asc' ? '↑' : '↓'}
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Error message display if any */}
        {errorMessage && (
          <div className="p-4 bg-rose-50 border border-rose-200 rounded-xl text-rose-800 text-xs flex items-center gap-2">
            <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
            <span>{errorMessage}</span>
          </div>
        )}

        {/* Cases Data Table */}
        <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
          <div className="px-5 py-3 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                Cases Created on {humanReadableDate}
              </span>
              <span className="px-2 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">
                {filteredAndSortedEntries.length} Records
              </span>
            </div>
            <span className="text-[11px] text-slate-500">
              Tab: <strong className="text-slate-700">{selectedTab === 'all' ? 'All Tabs' : selectedTab}</strong> | Workbook:{' '}
              <strong className="text-slate-700">{currentSheet?.name || 'MIS'}</strong>
            </span>
          </div>

          {isLoadingEntries ? (
            <div className="p-12 text-center">
              <RefreshCw className="w-8 h-8 text-emerald-600 animate-spin mx-auto mb-2" />
              <p className="text-xs font-semibold text-slate-700">Loading cases for {selectedDate}...</p>
              <p className="text-[11px] text-slate-400 mt-1">Scanning {selectedDateColumn} across sheet records</p>
            </div>
          ) : filteredAndSortedEntries.length === 0 ? (
            <div className="p-12 text-center space-y-3">
              <div className="w-12 h-12 bg-slate-100 rounded-full flex items-center justify-center mx-auto text-slate-400">
                <CalendarDays className="w-6 h-6" />
              </div>
              <h3 className="text-sm font-bold text-slate-800">
                No cases found created on {humanReadableDate}
              </h3>
              <p className="text-xs text-slate-500 max-w-md mx-auto">
                No rows in <strong>{selectedTab === 'all' ? 'All Tabs' : selectedTab}</strong> match date{' '}
                <code className="text-emerald-700 font-bold bg-emerald-50 px-1 py-0.5 rounded">{selectedDate}</code> in column{' '}
                <code className="bg-slate-100 px-1 py-0.5 rounded">{selectedDateColumn}</code>.
              </p>

              {availableDates.length > 0 && (
                <div className="pt-2">
                  <p className="text-xs font-semibold text-slate-600 mb-2">
                    Try selecting one of the closest dates that have active cases:
                  </p>
                  <div className="flex flex-wrap justify-center gap-1.5">
                    {availableDates.slice(0, 6).map((d) => (
                      <button
                        key={d.date}
                        onClick={() => setSelectedDate(d.date)}
                        className="px-2.5 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 rounded-md text-xs font-medium transition-colors"
                      >
                        {d.displayDate} ({d.count} cases)
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-100/70 border-b border-slate-200 text-slate-600 font-semibold uppercase text-[10px] tracking-wider">
                    <th className="py-2.5 px-3 text-center w-12">#</th>
                    <th className="py-2.5 px-3">Ticket ID</th>
                    <th className="py-2.5 px-3">Case Created Date</th>
                    <th className="py-2.5 px-3">Customer Name</th>
                    <th className="py-2.5 px-3">Phone</th>
                    <th className="py-2.5 px-3">Lender</th>
                    <th className="py-2.5 px-3 text-right">Loan Amount</th>
                    <th className="py-2.5 px-3 text-center">Status</th>
                    <th className="py-2.5 px-3">TL / Manager</th>
                    <th className="py-2.5 px-3">City / Branch</th>
                    <th className="py-2.5 px-3">Last Updated</th>
                    <th className="py-2.5 px-3 text-center w-28">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredAndSortedEntries.map((entry, idx) => {
                    const isExpanded = expandedRowId === entry.id;
                    const statusUpper = (entry.status || '').toUpperCase();

                    let statusBadgeClass = 'bg-slate-100 text-slate-700 border-slate-200';
                    if (statusUpper.includes('APPROV')) {
                      statusBadgeClass = 'bg-emerald-100 text-emerald-800 border-emerald-300';
                    } else if (statusUpper.includes('DISBURS')) {
                      statusBadgeClass = 'bg-blue-100 text-blue-800 border-blue-300';
                    } else if (statusUpper.includes('LOGIN') || statusUpper.includes('PROGRESS')) {
                      statusBadgeClass = 'bg-amber-100 text-amber-800 border-amber-300';
                    } else if (statusUpper.includes('REJECT') || statusUpper.includes('DECLINE')) {
                      statusBadgeClass = 'bg-rose-100 text-rose-800 border-rose-300';
                    } else if (statusUpper.includes('HOLD')) {
                      statusBadgeClass = 'bg-purple-100 text-purple-800 border-purple-300';
                    }

                    return (
                      <React.Fragment key={`${entry.id}-${idx}`}>
                        <tr
                          onClick={() => setExpandedRowId(isExpanded ? null : entry.id)}
                          className={`hover:bg-slate-50/80 transition-colors cursor-pointer ${
                            idx % 2 === 1 ? 'bg-slate-50/40' : 'bg-white'
                          } ${isExpanded ? 'bg-emerald-50/30' : ''}`}
                        >
                          {/* Row # */}
                          <td className="py-2.5 px-3 text-center text-slate-400 font-mono text-[11px]">
                            {idx + 1}
                          </td>

                          {/* Ticket ID */}
                          <td className="py-2.5 px-3">
                            {entry.ticketId ? (
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  if (onNavigateToTicket) onNavigateToTicket(entry.ticketId);
                                }}
                                className="font-mono font-bold text-slate-900 hover:text-indigo-600 transition-colors flex items-center gap-1 group"
                              >
                                <span>{entry.ticketId}</span>
                                <ExternalLink className="w-2.5 h-2.5 opacity-0 group-hover:opacity-100 text-indigo-500" />
                              </button>
                            ) : (
                              <span className="text-slate-300">—</span>
                            )}
                          </td>

                          {/* Case Created Date (Highlighted) */}
                          <td className="py-2.5 px-3">
                            <span
                              className="inline-flex items-center gap-1 px-2 py-0.5 rounded font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 text-[11px]"
                              title={`Normalized ISO: ${entry.normalizedDate} | Column: ${entry.matchedDateColumn}`}
                            >
                              <Calendar className="w-3 h-3 text-emerald-600" />
                              {entry.caseCreatedDate || entry.rawDateValue}
                            </span>
                          </td>

                          {/* Customer Name */}
                          <td className="py-2.5 px-3 font-semibold text-slate-900">
                            <div className="flex items-center gap-1.5">
                              <span className="truncate max-w-[160px]">{entry.customerName || '—'}</span>
                            </div>
                          </td>

                          {/* Phone Number */}
                          <td className="py-2.5 px-3">
                            {entry.phone ? (
                              <div className="flex items-center gap-1">
                                <span className="font-mono font-medium text-slate-800">
                                  {entry.phone}
                                </span>
                                <button
                                  type="button"
                                  onClick={(e) => handleCopyPhone(entry.phone, e)}
                                  className="p-1 text-slate-400 hover:text-slate-700 transition-colors rounded hover:bg-slate-200"
                                  title="Copy phone"
                                >
                                  {copiedPhone === entry.phone ? (
                                    <Check className="w-3 h-3 text-emerald-600" />
                                  ) : (
                                    <Copy className="w-3 h-3" />
                                  )}
                                </button>
                              </div>
                            ) : (
                              <span className="text-slate-300">—</span>
                            )}
                          </td>

                          {/* Lender */}
                          <td className="py-2.5 px-3">
                            {entry.lender ? (
                              <span className="px-2 py-0.5 rounded text-[11px] font-semibold bg-slate-100 text-slate-700 border border-slate-200">
                                {entry.lender}
                              </span>
                            ) : (
                              <span className="text-slate-300">—</span>
                            )}
                          </td>

                          {/* Loan Amount */}
                          <td className="py-2.5 px-3 text-right font-mono font-semibold text-slate-900">
                            {entry.loanAmount ? `₹ ${entry.loanAmount}` : <span className="text-slate-300 font-normal">—</span>}
                          </td>

                          {/* Status */}
                          <td className="py-2.5 px-3 text-center">
                            <span className={`inline-block px-2 py-0.5 rounded text-[10px] font-bold border ${statusBadgeClass}`}>
                              {entry.status || 'LOGIN'}
                            </span>
                          </td>

                          {/* TL / Manager */}
                          <td className="py-2.5 px-3 text-slate-600 truncate max-w-[140px]" title={`TL: ${entry.tlName} | Mgr: ${entry.manager}`}>
                            {entry.tlName || entry.manager ? (
                              <span>{entry.tlName || entry.manager}</span>
                            ) : (
                              <span className="text-slate-300">—</span>
                            )}
                          </td>

                          {/* City / Branch */}
                          <td className="py-2.5 px-3 text-slate-600 truncate max-w-[120px]">
                            {entry.city || entry.branch ? (
                              <span>{entry.city || entry.branch}</span>
                            ) : (
                              <span className="text-slate-300">—</span>
                            )}
                          </td>

                          {/* Last Updated */}
                          <td className="py-2.5 px-3 text-slate-500 font-mono text-[11px]">
                            {entry.lastUpdatedDate || <span className="text-slate-300">—</span>}
                          </td>

                          {/* Actions */}
                          <td className="py-2.5 px-3 text-center" onClick={(e) => e.stopPropagation()}>
                            <div className="flex items-center justify-center gap-1">
                              {onViewRecord && (
                                <button
                                  type="button"
                                  onClick={() => onViewRecord(entry.recordKey || entry.id)}
                                  className="p-1.5 text-slate-500 hover:text-emerald-700 hover:bg-emerald-50 rounded transition-colors"
                                  title="View Full Row Details"
                                >
                                  <Eye className="w-3.5 h-3.5" />
                                </button>
                              )}
                              {onLookupCustomer && entry.phone && (
                                <button
                                  type="button"
                                  onClick={() => onLookupCustomer(entry.phone)}
                                  className="p-1.5 text-slate-500 hover:text-indigo-700 hover:bg-indigo-50 rounded transition-colors"
                                  title="Lookup in Customer Operations"
                                >
                                  <UserCheck className="w-3.5 h-3.5" />
                                </button>
                              )}
                              <button
                                type="button"
                                onClick={() => setExpandedRowId(isExpanded ? null : entry.id)}
                                className="p-1.5 text-slate-400 hover:text-slate-700 rounded transition-colors"
                                title="Expand Details"
                              >
                                {isExpanded ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                              </button>
                            </div>
                          </td>
                        </tr>

                        {/* Expanded details row */}
                        {isExpanded && (
                          <tr className="bg-emerald-50/20 border-b border-emerald-100">
                            <td colSpan={12} className="py-3 px-6">
                              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                                {/* Remarks section */}
                                <div className="space-y-1.5 md:col-span-2">
                                  {entry.opsRemarks && (
                                    <div className="bg-white p-2.5 rounded-lg border border-slate-200">
                                      <span className="font-bold text-slate-700 block text-[11px] mb-0.5">
                                        Operations Remarks (Ops):
                                      </span>
                                      <p className="text-slate-600 text-xs leading-relaxed font-sans">
                                        {entry.opsRemarks}
                                      </p>
                                    </div>
                                  )}

                                  {entry.creditRemarks && (
                                    <div className="bg-white p-2.5 rounded-lg border border-slate-200">
                                      <span className="font-bold text-slate-700 block text-[11px] mb-0.5">
                                        Credit Remarks:
                                      </span>
                                      <p className="text-slate-600 text-xs leading-relaxed font-sans">
                                        {entry.creditRemarks}
                                      </p>
                                    </div>
                                  )}
                                </div>

                                {/* Additional Metadata column */}
                                <div className="bg-white p-3 rounded-lg border border-slate-200 space-y-1.5 text-[11px]">
                                  <div className="font-bold text-slate-800 border-b border-slate-100 pb-1 mb-1">
                                    Case Details & POC
                                  </div>
                                  {entry.netLoanAmount && (
                                    <div className="flex justify-between">
                                      <span className="text-slate-400">Net Loan Amount:</span>
                                      <span className="font-bold text-slate-800">₹ {entry.netLoanAmount}</span>
                                    </div>
                                  )}
                                  {entry.approvalDate && (
                                    <div className="flex justify-between">
                                      <span className="text-slate-400">Approval Date:</span>
                                      <span className="font-medium text-slate-700">{entry.approvalDate}</span>
                                    </div>
                                  )}
                                  {entry.disbursedDate && (
                                    <div className="flex justify-between">
                                      <span className="text-slate-400">Disbursed Date:</span>
                                      <span className="font-medium text-slate-700">{entry.disbursedDate}</span>
                                    </div>
                                  )}
                                  {entry.empName && (
                                    <div className="flex justify-between">
                                      <span className="text-slate-400">Emp / Login POC:</span>
                                      <span className="font-medium text-slate-700">{entry.empName}</span>
                                    </div>
                                  )}
                                  {entry.bankerName && (
                                    <div className="flex justify-between">
                                      <span className="text-slate-400">Banker:</span>
                                      <span className="font-medium text-slate-700">{entry.bankerName}</span>
                                    </div>
                                  )}
                                  {entry.bankerPhone && (
                                    <div className="flex justify-between">
                                      <span className="text-slate-400">Banker Phone:</span>
                                      <span className="font-mono text-slate-700">{entry.bankerPhone}</span>
                                    </div>
                                  )}
                                  <div className="pt-2 flex gap-2">
                                    {onViewRecord && (
                                      <button
                                        type="button"
                                        onClick={() => onViewRecord(entry.recordKey || entry.id)}
                                        className="w-full text-center py-1 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded font-medium text-[11px] transition-colors"
                                      >
                                        Open Full Row
                                      </button>
                                    )}
                                  </div>
                                </div>
                              </div>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          {/* Table Footer */}
          <div className="px-5 py-3 border-t border-slate-200 bg-slate-50 flex flex-col sm:flex-row items-center justify-between text-xs text-slate-500 gap-2">
            <div>
              Showing {filteredAndSortedEntries.length} of {entries.length} cases created on{' '}
              <strong className="text-slate-800">{humanReadableDate}</strong>
            </div>
            <div className="flex items-center gap-3">
              <span>
                Sheet: <strong>{currentSheet?.name}</strong> ({selectedTab})
              </span>
              <button
                onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })}
                className="text-emerald-700 font-semibold hover:underline"
              >
                Back to Top ↑
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
