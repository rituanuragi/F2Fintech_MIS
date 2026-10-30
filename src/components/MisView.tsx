import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  TableProperties,
  Link2,
  Upload,
  FileSpreadsheet,
  Download,
  Search,
  CheckCircle2,
  AlertCircle,
  Loader2,
  Sparkles,
  Sliders,
  RefreshCw,
  Layers,
  ChevronDown,
  ChevronRight,
  Copy,
  Check,
  CheckSquare,
  Square,
  Filter,
  ExternalLink,
  ShieldCheck,
  Eye,
  FileText,
  Trash2,
  PlusCircle,
  MinusCircle,
  X,
  FolderTree,
  Plus,
} from 'lucide-react';
import { MisDetectedSheet, MisEnrichedRow, MisSummary, LeadPriorityRule } from '../types';
import { PriorityConfigModal } from './PriorityConfigModal';

interface MisViewProps {
  onViewRecord?: (id: number) => void;
  onPreviewSheet?: (fileId: string, sheetName: string) => void;
  onLookupCustomer?: (phone: string) => void;
}

export const MisView: React.FC<MisViewProps> = ({
  onViewRecord,
  onPreviewSheet,
  onLookupCustomer,
}) => {
  // Input mode: 'google_sheet' or 'excel_file'
  const [inputMode, setInputMode] = useState<'google_sheet' | 'excel_file'>('google_sheet');

  // Google Sheet state
  const [googleSheetUrl, setGoogleSheetUrl] = useState(
    'https://docs.google.com/spreadsheets/d/15cVL3BK6yXPTwGGG0s_gp-jW-KxS-70cXn__6C8VCBs/edit?gid=614906163#gid=614906163'
  );
  const [autoSyncVersion, setAutoSyncVersion] = useState<number>(0);
  const [lastAutoUpdateBanner, setLastAutoUpdateBanner] = useState<string | null>(null);
  const prevSyncVerRef = useRef<number | null>(null);

  // Excel File state
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [excelFileName, setExcelFileName] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Parsed sheets & column mappings
  const [detectedSheets, setDetectedSheets] = useState<MisDetectedSheet[]>([]);
  const [sheetPhoneCols, setSheetPhoneCols] = useState<Record<string, string>>({});
  const [selectedSheets, setSelectedSheets] = useState<Set<string>>(new Set());
  const [sheetSearchQuery, setSheetSearchQuery] = useState('');
  const [workbookBuffer, setWorkbookBuffer] = useState<ArrayBuffer | null>(null);

  // Processing & Loading states
  const [isLoading, setIsLoading] = useState(false);
  const [loadingMessage, setLoadingMessage] = useState('');
  const [isExporting, setIsExporting] = useState(false);
  const [exportSuccess, setExportSuccess] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // Analysis Results
  const [results, setResults] = useState<MisEnrichedRow[]>([]);
  const [summary, setSummary] = useState<MisSummary | null>(null);
  const [analysisSourceTitle, setAnalysisSourceTitle] = useState('');

  // Filters & Pagination
  const [filterStatus, setFilterStatus] = useState<'all' | 'matched' | 'unmatched' | 'invalid'>('all');
  const [filterSheet, setFilterSheet] = useState<string>('all');
  const [filterPriority, setFilterPriority] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);
  const [copiedPhone, setCopiedPhone] = useState<string | null>(null);
  const [expandedRowIndex, setExpandedRowIndex] = useState<number | null>(null);

  // Priority modal
  const [isPriorityModalOpen, setIsPriorityModalOpen] = useState(false);
  const [priorityRules, setPriorityRules] = useState<LeadPriorityRule[]>([]);
  const [isSyncingPriority, setIsSyncingPriority] = useState(false);
  const [syncPriorityMessage, setSyncPriorityMessage] = useState<string | null>(null);

  const handleSyncGlobalPriority = async () => {
    setIsSyncingPriority(true);
    setSyncPriorityMessage(null);
    try {
      const res = await fetch('/api/priority-rules/sync', { method: 'POST' });
      const data = await res.json();
      if (data.rules) setPriorityRules(data.rules);
      setSyncPriorityMessage(
        data.addedCount > 0
          ? `+${data.addedCount} new added at end`
          : 'Priority Synced'
      );
      setTimeout(() => setSyncPriorityMessage(null), 3500);
    } catch (err) {
      console.error('Failed to sync priority rules:', err);
    } finally {
      setIsSyncingPriority(false);
    }
  };

  // Open the Priority screen directly on "Combine Columns"
  const [priorityOpenCombine, setPriorityOpenCombine] = useState(false);

  // Fetch active priority rules
  useEffect(() => {
    fetch('/api/priority-rules')
      .then(res => res.json())
      .then(data => {
        if (data.rules) setPriorityRules(data.rules);
      })
      .catch(console.error);
  }, []);

  // Poll background MIS Auto-Sync engine so any new entry added in Google Sheet auto-updates active analysis
  useEffect(() => {
    const checkAutoSync = async () => {
      try {
        const res = await fetch('/api/mis/auto-sync-status');
        const data = await res.json();
        if (data?.success && data.autoSync) {
          const ver = data.autoSync.syncVersion || 0;
          setAutoSyncVersion(ver);
          if (prevSyncVerRef.current !== null && ver > prevSyncVerRef.current) {
            prevSyncVerRef.current = ver;
            const diff = data.autoSync.lastNewRowsCount || 0;
            setLastAutoUpdateBanner(
              diff > 0
                ? `Auto-Synced +${diff} new row(s) from MIS Sheet at ${new Date().toLocaleTimeString()}`
                : `Auto-Synced latest MIS Sheet changes at ${new Date().toLocaleTimeString()}`
            );
            if (inputMode === 'google_sheet' && googleSheetUrl.trim() && results.length > 0 && !isLoading) {
              handleProcessGoogleSheet(selectedSheets.size > 0 ? Array.from(selectedSheets) : undefined);
            }
          } else if (prevSyncVerRef.current === null) {
            prevSyncVerRef.current = ver;
          }
        }
      } catch {
        // ignore transient network errors
      }
    };
    checkAutoSync();
    const timer = setInterval(checkAutoSync, 10000);
    return () => clearInterval(timer);
  }, [inputMode, googleSheetUrl, results.length, isLoading, selectedSheets]);

  // Helper to copy phone number
  const handleCopy = (phone: string) => {
    navigator.clipboard.writeText(phone);
    setCopiedPhone(phone);
    setTimeout(() => setCopiedPhone(null), 1500);
  };

  // Format cell values from SheetJS without 1-day-before timezone shift on dates
  const formatClientCellValue = (val: any): any => {
    if (val === null || val === undefined) return '';
    if (val instanceof Date) {
      if (isNaN(val.getTime())) return '';
      const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      const adjusted = new Date(val.getTime() + 12 * 3600 * 1000);
      const day = adjusted.getUTCDate();
      const month = months[adjusted.getUTCMonth()];
      const yearShort = String(adjusted.getUTCFullYear()).slice(-2);
      return `${day}-${month}-${yearShort}`;
    }
    return val;
  };

  // Toggle single sheet selection (Add / Remove)
  const handleToggleSheet = (sheetName: string) => {
    setSelectedSheets(prev => {
      const next = new Set(prev);
      if (next.has(sheetName)) {
        next.delete(sheetName);
      } else {
        next.add(sheetName);
      }
      return next;
    });
  };

  // Select all sheets
  const handleSelectAllSheets = () => {
    setSelectedSheets(new Set(detectedSheets.map(s => s.sheetName)));
  };

  // Deselect all sheets
  const handleDeselectAllSheets = () => {
    setSelectedSheets(new Set());
  };

  // Reset all state for a new analysis
  const handleReset = () => {
    setDetectedSheets([]);
    setSheetPhoneCols({});
    setSelectedSheets(new Set());
    setWorkbookBuffer(null);
    setResults([]);
    setSummary(null);
    setErrorMessage(null);
    setExcelFile(null);
    setExcelFileName('');
    setGoogleSheetUrl('');
    setSheetSearchQuery('');
    setCurrentPage(1);
  };

  // 1. Process Google Sheet: Fetch, Inspect sheets, and Run Analysis
  const handleProcessGoogleSheet = async (sheetsToInclude?: string[]) => {
    if (!googleSheetUrl.trim()) {
      setErrorMessage('Please enter a valid Google Sheet URL.');
      return;
    }

    setIsLoading(true);
    setErrorMessage(null);
    setLoadingMessage('Fetching Google Sheet across all sheets...');

    try {
      const res = await fetch('/api/mis/process-google-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: googleSheetUrl.trim(),
          customPhoneColumns: sheetPhoneCols,
          selectedSheets: sheetsToInclude || (selectedSheets.size > 0 ? Array.from(selectedSheets) : undefined),
        }),
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Failed to process Google Sheet');
      }

      setDetectedSheets(data.sheets || []);
      const initCols: Record<string, string> = { ...sheetPhoneCols };
      const allSheetNames: string[] = [];

      (data.sheets || []).forEach((s: MisDetectedSheet) => {
        allSheetNames.push(s.sheetName);
        if (!initCols[s.sheetName]) {
          initCols[s.sheetName] = s.detectedPhoneColName;
        }
      });
      setSheetPhoneCols(initCols);

      if (selectedSheets.size === 0) {
        setSelectedSheets(new Set(allSheetNames));
      }

      setResults(data.results || []);
      setSummary(data.summary || null);
      setAnalysisSourceTitle('Google Sheet Analysis');
      setCurrentPage(1);
    } catch (err: any) {
      setErrorMessage(err.message || 'Error communicating with server');
    } finally {
      setIsLoading(false);
      setLoadingMessage('');
    }
  };

  // 2. Process Uploaded Excel File: Parse all sheets client-side, detect phone column, and analyze
  const handleExcelUpload = async (file: File) => {
    setExcelFile(file);
    setExcelFileName(file.name);
    setIsLoading(true);
    setErrorMessage(null);
    setResults([]);
    setSummary(null);
    setLoadingMessage(`Reading ${file.name} across all sheets...`);

    try {
      const arrayBuf = await file.arrayBuffer();
      setWorkbookBuffer(arrayBuf);

      const workbook = XLSX.read(arrayBuf, { type: 'array', cellDates: true });
      const detected: MisDetectedSheet[] = [];
      const initCols: Record<string, string> = {};
      const allSheetNames: string[] = [];

      const rowsToAnalyze: {
        sourceSheet: string;
        sourceRowNum: number;
        rawPhone: string;
        originalData: Record<string, any>;
      }[] = [];

      for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) continue;

        const rawJson = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' }) as any[][];
        if (!rawJson || rawJson.length === 0) continue;

        // Find header row
        let headerRowIndex = 0;
        for (let r = 0; r < Math.min(rawJson.length, 10); r++) {
          const row = rawJson[r];
          if (row && row.some(cell => String(cell || '').trim().length > 0)) {
            headerRowIndex = r;
            break;
          }
        }

        const headers = (rawJson[headerRowIndex] || []).map((h, i) => String(h || `Column_${i + 1}`).trim());
        const dataRows = rawJson.slice(headerRowIndex + 1).filter(r => r && r.some(cell => String(cell || '').trim().length > 0));

        // Auto-detect phone column (case and space insensitive)
        let phoneColIdx = 0;
        const normalizedHeaders = headers.map(h => h.toLowerCase().replace(/[\s_\-.]+/g, ''));
        const priorityKeywords = ['phonenumber', 'mobilenumber', 'contactnumber', 'phone', 'mobile', 'contact', 'phoneno', 'mobno', 'phone1', 'number1'];

        for (const kw of priorityKeywords) {
          const idx = normalizedHeaders.findIndex(nh => nh === kw);
          if (idx !== -1) {
            phoneColIdx = idx;
            break;
          }
        }

        if (phoneColIdx === 0) {
          const patternIdx = normalizedHeaders.findIndex(nh => /phone|mobile|contact/i.test(nh));
          if (patternIdx !== -1) phoneColIdx = patternIdx;
        }

        const phoneColName = headers[phoneColIdx] || 'Column 1';
        initCols[sheetName] = phoneColName;
        allSheetNames.push(sheetName);

        // Filter out completely empty rows and rows where phone number is blank/missing
        const validDataRows = dataRows.filter(row => {
          if (!row || !row.some(cell => String(cell || '').trim().length > 0)) return false;
          const pVal = row[phoneColIdx] !== undefined && row[phoneColIdx] !== null ? String(row[phoneColIdx]).trim() : '';
          if (!pVal) return false;
          const pLower = pVal.toLowerCase();
          if (['none', 'null', 'nan', 'n/a', 'na', 'nil', '-', '--', '0', '#n/a', 'undefined', 'not available'].includes(pLower)) {
            return false;
          }
          return true;
        });

        detected.push({
          sheetName,
          totalRows: validDataRows.length,
          detectedPhoneColIndex: phoneColIdx,
          detectedPhoneColName: phoneColName,
          headers,
          previewRows: validDataRows.slice(0, 3).map((r, idx) => {
            const obj: Record<string, any> = { _rowNum: headerRowIndex + 2 + idx };
            headers.forEach((h, c) => {
              obj[h] = r[c] !== undefined ? formatClientCellValue(r[c]) : '';
            });
            return obj;
          }),
        });

        // Collect rows for this sheet (only valid rows with actual phone numbers)
        validDataRows.forEach((row, rIdx) => {
          const rawPhone = String(row[phoneColIdx] || '').trim();
          const rowObj: Record<string, any> = {};
          headers.forEach((h, cIdx) => {
            rowObj[h] = row[cIdx] !== undefined ? formatClientCellValue(row[cIdx]) : '';
          });

          rowsToAnalyze.push({
            sourceSheet: sheetName,
            sourceRowNum: headerRowIndex + 2 + rIdx,
            rawPhone,
            originalData: rowObj,
          });
        });
      }

      setDetectedSheets(detected);
      setSheetPhoneCols(initCols);
      setSelectedSheets(new Set(allSheetNames));

      if (rowsToAnalyze.length === 0) {
        throw new Error('No valid data rows found in the uploaded workbook.');
      }

      setLoadingMessage(`Analyzing ${rowsToAnalyze.length.toLocaleString()} rows against fintech database & priorities...`);

      // Send to server for database matching and priority ranking
      const res = await fetch('/api/mis/analyze-rows', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows: rowsToAnalyze }),
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Failed to analyze rows');
      }

      setResults(data.results || []);
      setSummary(data.summary || null);
      setAnalysisSourceTitle(file.name);
      setCurrentPage(1);
    } catch (err: any) {
      setErrorMessage(err.message || 'Error processing Excel file');
    } finally {
      setIsLoading(false);
      setLoadingMessage('');
    }
  };

  // 3. Re-run analysis if user changed sheet selections or phone column mappings
  const handleApplySheetSelectionAndColumns = async () => {
    if (!workbookBuffer && !googleSheetUrl) return;

    if (selectedSheets.size === 0) {
      setErrorMessage('Please select at least one sheet to include in the analysis.');
      return;
    }

    if (inputMode === 'google_sheet') {
      handleProcessGoogleSheet(Array.from(selectedSheets));
      return;
    }

    if (!workbookBuffer) return;

    setIsLoading(true);
    setErrorMessage(null);
    setLoadingMessage(`Re-analyzing ${selectedSheets.size} selected sheet(s)...`);

    try {
      const workbook = XLSX.read(workbookBuffer, { type: 'array', cellDates: true });
      const rowsToAnalyze: {
        sourceSheet: string;
        sourceRowNum: number;
        rawPhone: string;
        originalData: Record<string, any>;
      }[] = [];

      for (const sheetInfo of detectedSheets) {
        const sheetName = sheetInfo.sheetName;
        // Skip sheet if not selected
        if (!selectedSheets.has(sheetName)) continue;

        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) continue;

        const rawJson = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' }) as any[][];
        if (!rawJson || rawJson.length === 0) continue;

        let headerRowIndex = 0;
        for (let r = 0; r < Math.min(rawJson.length, 10); r++) {
          const row = rawJson[r];
          if (row && row.some(cell => String(cell || '').trim().length > 0)) {
            headerRowIndex = r;
            break;
          }
        }

        const headers = (rawJson[headerRowIndex] || []).map((h, i) => String(h || `Column_${i + 1}`).trim());
        const selectedCol = sheetPhoneCols[sheetName] || sheetInfo.detectedPhoneColName;
        let colIdx = headers.findIndex(h => h.trim().toLowerCase() === selectedCol.trim().toLowerCase());
        if (colIdx === -1) colIdx = 0;

        const dataRows = rawJson.slice(headerRowIndex + 1);
        dataRows.forEach((row, rIdx) => {
          // 1. Skip completely empty rows
          if (!row || !row.some(cell => String(cell || '').trim().length > 0)) return;

          // 2. Skip rows where phone number is blank/missing/placeholder
          const rawPhone = row[colIdx] !== undefined && row[colIdx] !== null ? String(row[colIdx]).trim() : '';
          if (!rawPhone) return;
          const lowerPhone = rawPhone.toLowerCase();
          if (['none', 'null', 'nan', 'n/a', 'na', 'nil', '-', '--', '0', '#n/a', 'undefined', 'not available'].includes(lowerPhone)) {
            return;
          }

          const rowObj: Record<string, any> = {};
          headers.forEach((h, cIdx) => {
            rowObj[h] = row[cIdx] !== undefined ? formatClientCellValue(row[cIdx]) : '';
          });

          rowsToAnalyze.push({
            sourceSheet: sheetName,
            sourceRowNum: headerRowIndex + 2 + rIdx,
            rawPhone,
            originalData: rowObj,
          });
        });
      }

      if (rowsToAnalyze.length === 0) {
        throw new Error('The selected sheet(s) have no data rows.');
      }

      const res = await fetch('/api/mis/analyze-rows', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows: rowsToAnalyze }),
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Failed to re-analyze');
      }

      setResults(data.results || []);
      setSummary(data.summary || null);
      setCurrentPage(1);
    } catch (err: any) {
      setErrorMessage(err.message || 'Failed to re-analyze');
    } finally {
      setIsLoading(false);
      setLoadingMessage('');
    }
  };

  // 4. Export Enriched Excel File (.xlsx)
  // Contains:
  // - Sheet 1: "Combined Sheet" with:
  //   Phone Number, Sheet Name (of file in which present), Priority (from db), Sheet Name Database, Workbook Name
  // - Following Sheets: Each of the individual sheets that were selected by the user!
  const handleExportExcel = async () => {
    if (results.length === 0) return;
    setIsExporting(true);
    setExportError(null);
    setExportSuccess(false);

    try {
      const selectedList = selectedSheets.size > 0 
        ? Array.from(selectedSheets) 
        : Array.from(new Set(results.map(r => r.sourceSheet)));

      const filtered = results.filter(r => selectedList.includes(r.sourceSheet));

      if (filtered.length === 0) {
        throw new Error('No records found in the selected sheet(s) to export. Please select at least one sheet.');
      }

      // 1. Build workbook client-side with SheetJS
      const workbook = XLSX.utils.book_new();

      // Master Combined Sheet
      const masterRows = filtered.map(r => {
        const out: Record<string, any> = {
          'Phone Number': r.inputPhone,
          'Sheet Name': r.sourceSheet,
          'Priority': r.priority,
          'Sheet Name Database': r.foundSheet,
          'Workbook Name': r.foundWorkbook,
          'Match Status': r.found ? 'MATCHED' : 'NOT FOUND',
          'Customer Name (Database)': r.foundCustomerName,
          'Lender (Database)': r.foundLender,
          'Loan Amount (Database)': r.foundLoanAmount,
          'Row Number (Database)': r.foundRowNum,
          'Source Row': r.sourceRowNum,
        };

        if (r.originalData) {
          Object.entries(r.originalData).forEach(([k, v]) => {
            if (!k.startsWith('_')) {
              out[k] = v;
            }
          });
        }

        return out;
      });

      const masterSheet = XLSX.utils.json_to_sheet(masterRows);
      XLSX.utils.book_append_sheet(workbook, masterSheet, 'Combined Sheet');

      // 2. Matched Sheet (Only Matched in DB)
      const matchedRows = filtered.filter(r => r.found).map(r => {
        const out: Record<string, any> = {
          'Phone Number': r.inputPhone,
          'Sheet Name': r.sourceSheet,
          'Priority': r.priority,
          'Workbook Name': r.foundWorkbook,
          'Sheet Name Database': r.foundSheet,
          'Customer Name (Database)': r.foundCustomerName,
          'Lender (Database)': r.foundLender,
          'Loan Amount (Database)': r.foundLoanAmount,
          'Row Number (Database)': r.foundRowNum,
          'Source Row': r.sourceRowNum,
        };
        if (r.originalData) {
          Object.entries(r.originalData).forEach(([k, v]) => {
            if (!k.startsWith('_')) out[k] = v;
          });
        }
        return out;
      });
      if (matchedRows.length > 0) {
        const matchedSheet = XLSX.utils.json_to_sheet(matchedRows);
        XLSX.utils.book_append_sheet(workbook, matchedSheet, 'Matched');
      }

      // 3. Unmatched Sheet (Valid 10-digit leads not found in DB)
      const unmatchedRows = filtered.filter(r => !r.found && !r.priority.includes('< 10') && !r.priority.includes('Invalid')).map(r => {
        const out: Record<string, any> = {
          'Phone Number': r.inputPhone,
          'Sheet Name': r.sourceSheet,
          'Status': 'Unmatched (Not Found in Database)',
          'Source Row': r.sourceRowNum,
        };
        if (r.originalData) {
          Object.entries(r.originalData).forEach(([k, v]) => {
            if (!k.startsWith('_')) out[k] = v;
          });
        }
        return out;
      });
      if (unmatchedRows.length > 0) {
        const unmatchedSheet = XLSX.utils.json_to_sheet(unmatchedRows);
        XLSX.utils.book_append_sheet(workbook, unmatchedSheet, 'Unmatched');
      }

      // 4. Ignored Sheet (<10 Digits or Invalid Format)
      const ignoredRows = filtered.filter(r => r.priority.includes('< 10') || r.priority.includes('Invalid')).map(r => {
        const out: Record<string, any> = {
          'Phone Number': r.inputPhone,
          'Sheet Name': r.sourceSheet,
          'Reason': r.priority,
          'Source Row': r.sourceRowNum,
        };
        if (r.originalData) {
          Object.entries(r.originalData).forEach(([k, v]) => {
            if (!k.startsWith('_')) out[k] = v;
          });
        }
        return out;
      });
      if (ignoredRows.length > 0) {
        const ignoredSheet = XLSX.utils.json_to_sheet(ignoredRows);
        XLSX.utils.book_append_sheet(workbook, ignoredSheet, 'Ignored');
      }

      // 5. Individual Selected Sheets
      const sheetsMap: Record<string, MisEnrichedRow[]> = {};
      filtered.forEach(r => {
        if (!sheetsMap[r.sourceSheet]) sheetsMap[r.sourceSheet] = [];
        sheetsMap[r.sourceSheet].push(r);
      });

      Object.entries(sheetsMap).forEach(([sheetName, sheetRows]) => {
        const safeSheetName = sheetName.replace(/[:\\/?*[\]]/g, '_').slice(0, 31);
        const sRows = sheetRows.map(r => {
          const out: Record<string, any> = {
            'Phone Number': r.inputPhone,
            'Priority': r.priority,
            'Workbook Name': r.foundWorkbook,
            'Sheet Name Database': r.foundSheet,
            'Match Status': r.found ? 'MATCHED' : 'NOT FOUND',
            'Customer Name (Database)': r.foundCustomerName,
            'Lender (Database)': r.foundLender,
            'Loan Amount (Database)': r.foundLoanAmount,
          };
          if (r.originalData) {
            Object.entries(r.originalData).forEach(([k, v]) => {
              if (!k.startsWith('_')) {
                out[k] = v;
              }
            });
          }
          return out;
        });

        const ws = XLSX.utils.json_to_sheet(sRows);
        XLSX.utils.book_append_sheet(workbook, ws, safeSheetName);
      });

      const dateStr = new Date().toISOString().slice(0, 10);
      const safeTitle = (analysisSourceTitle || 'MIS_Enriched').replace(/[^a-zA-Z0-9_-]/g, '_');
      const filename = `${safeTitle}_${dateStr}.xlsx`;

      // Trigger download using SheetJS writeFile + Blob fallback
      try {
        XLSX.writeFile(workbook, filename);
      } catch (writeErr) {
        // Blob URL fallback
        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
          document.body.removeChild(a);
          window.URL.revokeObjectURL(url);
        }, 1500);
      }

      setExportSuccess(true);
      setTimeout(() => setExportSuccess(false), 4000);
    } catch (err: any) {
      console.error('Export error:', err);
      setExportError(err.message || 'Failed to download Excel file');
    } finally {
      setIsExporting(false);
    }
  };

  // Filter results: only include selected sheets + filter status + filter priority + search query
  const filteredResults = results.filter(row => {
    // Only show rows from selected sheets
    if (selectedSheets.size > 0 && !selectedSheets.has(row.sourceSheet)) {
      return false;
    }

    // Status filter
    if (filterStatus === 'matched' && !row.found) return false;
    if (filterStatus === 'unmatched' && (row.found || row.priority.includes('< 10') || row.priority.includes('Invalid'))) return false;
    if (filterStatus === 'invalid' && !row.priority.includes('< 10') && !row.priority.includes('Invalid')) return false;

    // Sheet filter
    if (filterSheet !== 'all' && row.sourceSheet !== filterSheet) return false;

    // Priority filter
    if (filterPriority !== 'all' && row.priority !== filterPriority) return false;

    // Search query
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      const matchPhone = row.inputPhone.toLowerCase().includes(q) || row.normPhone.includes(q);
      const matchName = (row.foundCustomerName || '').toLowerCase().includes(q);
      const matchWb = (row.foundWorkbook || '').toLowerCase().includes(q);
      const matchSheet = (row.foundSheet || '').toLowerCase().includes(q);
      const matchPriority = (row.priority || '').toLowerCase().includes(q);
      const matchSrcSheet = (row.sourceSheet || '').toLowerCase().includes(q);
      return matchPhone || matchName || matchWb || matchSheet || matchPriority || matchSrcSheet;
    }

    return true;
  });

  const totalPages = Math.ceil(filteredResults.length / pageSize) || 1;
  const paginatedRows = filteredResults.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  // Available unique priorities and source sheets from the currently selected sheets
  const activeResults = results.filter(r => selectedSheets.size === 0 || selectedSheets.has(r.sourceSheet));
  const uniquePriorities = Array.from(new Set(activeResults.map(r => r.priority))).filter(Boolean);
  const uniqueSourceSheets = Array.from(new Set(activeResults.map(r => r.sourceSheet))).filter(Boolean);

  // Visible sheets in the mapping list (supports filtering by sheetSearchQuery)
  const visibleDetectedSheets = detectedSheets.filter(s =>
    s.sheetName.toLowerCase().includes(sheetSearchQuery.toLowerCase())
  );

  const totalSelectedRows = detectedSheets
    .filter(s => selectedSheets.has(s.sheetName))
    .reduce((sum, s) => sum + s.totalRows, 0);

  return (
    <div className="flex-1 bg-slate-50 flex flex-col min-h-0 overflow-y-auto">
      {/* Top Banner / Header */}
      <div className="bg-white border-b border-slate-200 px-6 py-5">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-700 flex items-center justify-center text-white shadow-sm">
                <TableProperties className="w-5 h-5" />
              </div>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <h1 className="text-xl font-bold text-slate-900 tracking-tight">MIS Analyzer</h1>
                  <span className="px-2 py-0.5 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200">
                    Lead Enrichment & Priority
                  </span>
                  <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-teal-50 text-teal-800 border border-teal-200">
                    <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                    Auto-Sync Active (Zero Manual Sync Needed)
                  </span>
                  {lastAutoUpdateBanner && (
                    <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-indigo-50 text-indigo-700 border border-indigo-200">
                      {lastAutoUpdateBanner}
                    </span>
                  )}
                </div>
                <p className="text-xs text-slate-500 mt-0.5">
                  Analyze Google Sheets or Excel workbooks. Select/remove sheets, auto-detect Phone number columns, match against your database, and download the Combined Sheet with Priority, Workbook & Sheet names.
                </p>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() => { setPriorityOpenCombine(true); setIsPriorityModalOpen(true); }}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold text-slate-700 bg-white border border-slate-200 hover:bg-slate-50 hover:border-slate-300 shadow-2xs transition-all cursor-pointer"
              title="Combine Columns: choose which columns are added to the Priority column"
            >
              <FolderTree className="w-3.5 h-3.5 text-violet-600" />
              <span>Combine Columns</span>
            </button>
            <button
              type="button"
              onClick={handleSyncGlobalPriority}
              disabled={isSyncingPriority}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-medium text-slate-700 bg-white border border-slate-200 hover:bg-slate-50 hover:border-slate-300 shadow-2xs transition-all cursor-pointer disabled:opacity-50"
              title="Refresh and sync all directories & files into the priority list (new directories automatically added at the end)"
            >
              <RefreshCw className={`w-3.5 h-3.5 text-slate-500 ${isSyncingPriority ? 'animate-spin text-emerald-600' : ''}`} />
              <span>{isSyncingPriority ? 'Syncing...' : syncPriorityMessage || 'Sync Priority'}</span>
            </button>
            <button
              type="button"
              onClick={() => { setPriorityOpenCombine(false); setIsPriorityModalOpen(true); }}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-medium text-slate-700 bg-white border border-slate-200 hover:bg-slate-50 hover:border-slate-300 shadow-2xs transition-all cursor-pointer"
            >
              <Sliders className="w-3.5 h-3.5 text-slate-500" />
              <span>Priority Rules</span>
            </button>
            {results.length > 0 && (
              <button
                type="button"
                onClick={handleReset}
                className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 transition-all cursor-pointer"
              >
                <RefreshCw className="w-3.5 h-3.5" />
                <span>New MIS Analysis</span>
              </button>
            )}
          </div>
        </div>
      </div>

      <div className="p-6 space-y-6 max-w-7xl mx-auto w-full">
        {/* Step 1: Input Choice (Google Sheet or Excel File) */}
        {results.length === 0 && (
          <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
            <div className="border-b border-slate-200 px-6 py-4 bg-slate-50/70 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="w-6 h-6 rounded-full bg-emerald-600 text-white flex items-center justify-center text-xs font-bold">1</span>
                <h2 className="text-sm font-semibold text-slate-800">Select Input Source</h2>
              </div>
              <div className="flex items-center p-1 bg-slate-200/70 rounded-lg text-xs font-medium">
                <button
                  type="button"
                  onClick={() => setInputMode('google_sheet')}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md transition-all ${
                    inputMode === 'google_sheet'
                      ? 'bg-white text-emerald-700 shadow-2xs font-semibold'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <Link2 className="w-3.5 h-3.5" />
                  <span>Google Sheet Link</span>
                </button>
                <button
                  type="button"
                  onClick={() => setInputMode('excel_file')}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md transition-all ${
                    inputMode === 'excel_file'
                      ? 'bg-white text-emerald-700 shadow-2xs font-semibold'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  <span>Upload Excel File</span>
                </button>
              </div>
            </div>

            <div className="p-6">
              {inputMode === 'google_sheet' ? (
                /* Google Sheet Link Input */
                <div className="space-y-4">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 uppercase tracking-wider mb-1.5">
                      Google Spreadsheet URL
                    </label>
                    <div className="flex items-center gap-3">
                      <div className="relative flex-1">
                        <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400">
                          <Link2 className="w-4 h-4" />
                        </div>
                        <input
                          type="url"
                          value={googleSheetUrl}
                          onChange={e => setGoogleSheetUrl(e.target.value)}
                          placeholder="https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit"
                          className="w-full pl-10 pr-4 py-2.5 bg-slate-50/50 border border-slate-300 rounded-lg text-sm text-slate-800 placeholder-slate-400 focus:outline-hidden focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all font-mono text-xs"
                          onKeyDown={e => {
                            if (e.key === 'Enter') handleProcessGoogleSheet();
                          }}
                        />
                      </div>
                      <button
                        type="button"
                        onClick={() => handleProcessGoogleSheet()}
                        disabled={isLoading || !googleSheetUrl.trim()}
                        className="inline-flex items-center gap-2 px-5 py-2.5 rounded-lg text-sm font-semibold text-white bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 disabled:cursor-not-allowed shadow-sm transition-all shrink-0 cursor-pointer"
                      >
                        {isLoading ? (
                          <>
                            <Loader2 className="w-4 h-4 animate-spin" />
                            <span>Processing...</span>
                          </>
                        ) : (
                          <>
                            <Sparkles className="w-4 h-4" />
                            <span>Analyze All Sheets</span>
                          </>
                        )}
                      </button>
                    </div>
                  </div>

                  <div className="p-3.5 bg-emerald-50/60 border border-emerald-100 rounded-lg flex items-start gap-2.5 text-xs text-emerald-800">
                    <ShieldCheck className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                    <div>
                      <span className="font-semibold">Sharing Requirement:</span> Make sure your Google Sheet is shared with{' '}
                      <span className="underline font-medium">"Anyone with the link can view"</span> so the server can securely access all sheets. Or download as Excel and upload on the right tab!
                    </div>
                  </div>
                </div>
              ) : (
                /* Excel File Upload Drag & Drop */
                <div>
                  <div
                    onClick={() => fileInputRef.current?.click()}
                    className="border-2 border-dashed border-slate-300 hover:border-emerald-500 bg-slate-50/60 hover:bg-emerald-50/20 rounded-xl p-8 text-center cursor-pointer transition-all group"
                  >
                    <input
                      ref={fileInputRef}
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      onChange={e => {
                        const f = e.target.files?.[0];
                        if (f) handleExcelUpload(f);
                      }}
                      className="hidden"
                    />
                    <div className="w-12 h-12 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center mx-auto mb-3 group-hover:scale-105 transition-transform">
                      <Upload className="w-6 h-6" />
                    </div>
                    <div className="text-sm font-semibold text-slate-800">
                      {excelFileName ? excelFileName : 'Click to select or drag and drop your Excel workbook'}
                    </div>
                    <p className="text-xs text-slate-500 mt-1">
                      Supports multi-sheet workbooks in .XLSX, .XLS, or .CSV format
                    </p>
                  </div>
                </div>
              )}

              {/* Error Alert */}
              {errorMessage && (
                <div className="mt-4 p-3.5 bg-rose-50 border border-rose-200 rounded-lg flex items-start gap-2.5 text-xs text-rose-800">
                  <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                  <div>
                    <span className="font-semibold">Error:</span> {errorMessage}
                  </div>
                </div>
              )}

              {/* Loading Banner */}
              {isLoading && (
                <div className="mt-4 p-4 bg-emerald-50 border border-emerald-200 rounded-lg flex items-center gap-3 text-xs text-emerald-800 animate-pulse">
                  <Loader2 className="w-5 h-5 text-emerald-600 animate-spin shrink-0" />
                  <div>
                    <div className="font-semibold">{loadingMessage || 'Analyzing data...'}</div>
                    <div className="text-emerald-700/80 text-[11px] mt-0.5">
                      Extracting sheets, detecting Phone number columns, querying SQLite index, and mapping priority rules.
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* Detected Sheets & Column Review + ADD / REMOVE SHEETS CONTROLS */}
        {detectedSheets.length > 0 && (
          <div className="bg-white rounded-xl border border-slate-200 shadow-2xs p-5">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-4 border-b border-slate-100 pb-3">
              <div>
                <div className="flex items-center gap-2">
                  <Layers className="w-4 h-4 text-emerald-600" />
                  <h3 className="text-xs font-bold text-slate-900 uppercase tracking-wider">
                    Sheets & Phone Column Mappings ({detectedSheets.length} Sheets Detected)
                  </h3>
                  <span className="px-2 py-0.5 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200">
                    {selectedSheets.size} Selected ({totalSelectedRows.toLocaleString()} rows)
                  </span>
                </div>
                <p className="text-xs text-slate-500 mt-1">
                  Check or uncheck sheets to add/remove them from analysis and final download. The column named "Phone number" is automatically mapped.
                </p>
              </div>

              {/* Quick Actions: Select All, Deselect All, Sheet Search */}
              <div className="flex flex-wrap items-center gap-2">
                <div className="relative">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={sheetSearchQuery}
                    onChange={e => setSheetSearchQuery(e.target.value)}
                    placeholder="Filter sheets (e.g. EMP, SEPT)..."
                    className="w-48 pl-8 pr-3 py-1 bg-slate-50 border border-slate-200 rounded-lg text-xs text-slate-700 placeholder-slate-400 focus:bg-white focus:outline-hidden focus:ring-1 focus:ring-emerald-500"
                  />
                  {sheetSearchQuery && (
                    <button
                      onClick={() => setSheetSearchQuery('')}
                      className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                    >
                      <X className="w-3 h-3" />
                    </button>
                  )}
                </div>

                <button
                  type="button"
                  onClick={handleSelectAllSheets}
                  className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 transition-colors cursor-pointer"
                >
                  <CheckSquare className="w-3.5 h-3.5 text-emerald-600" />
                  <span>Select All</span>
                </button>

                <button
                  type="button"
                  onClick={handleDeselectAllSheets}
                  className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 transition-colors cursor-pointer"
                >
                  <Square className="w-3.5 h-3.5 text-slate-400" />
                  <span>Deselect All</span>
                </button>
              </div>
            </div>

            {/* Sheets Grid */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 max-h-[420px] overflow-y-auto pr-1">
              {visibleDetectedSheets.map(s => {
                const isSelected = selectedSheets.has(s.sheetName);

                return (
                  <div
                    key={s.sheetName}
                    className={`p-3.5 rounded-xl border transition-all ${
                      isSelected
                        ? 'bg-white border-emerald-300 shadow-2xs ring-1 ring-emerald-50'
                        : 'bg-slate-50/70 border-slate-200 opacity-60'
                    }`}
                  >
                    <div className="flex items-start justify-between gap-2 mb-2">
                      <label className="flex items-center gap-2 cursor-pointer select-none">
                        <input
                          type="checkbox"
                          checked={isSelected}
                          onChange={() => handleToggleSheet(s.sheetName)}
                          className="w-4 h-4 rounded text-emerald-600 focus:ring-emerald-500 border-slate-300 cursor-pointer"
                        />
                        <span className={`text-xs font-bold truncate max-w-[170px] ${isSelected ? 'text-slate-900' : 'text-slate-400 line-through'}`} title={s.sheetName}>
                          {s.sheetName}
                        </span>
                      </label>

                      <div className="flex items-center gap-1.5 shrink-0">
                        <span className="text-[10px] font-mono text-slate-500 bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
                          {s.totalRows.toLocaleString()} rows
                        </span>
                        <button
                          type="button"
                          onClick={() => handleToggleSheet(s.sheetName)}
                          className={`text-[10px] font-semibold px-2 py-0.5 rounded transition-colors cursor-pointer ${
                            isSelected
                              ? 'text-rose-600 hover:bg-rose-50'
                              : 'text-emerald-700 hover:bg-emerald-50'
                          }`}
                          title={isSelected ? 'Remove from final export' : 'Add to final export'}
                        >
                          {isSelected ? 'Remove' : '+ Add'}
                        </button>
                      </div>
                    </div>

                    <div className="flex items-center gap-1.5 text-xs pt-1.5 border-t border-slate-100">
                      <span className="text-slate-500 text-[11px] shrink-0">Phone Col:</span>
                      <select
                        disabled={!isSelected}
                        value={sheetPhoneCols[s.sheetName] || s.detectedPhoneColName}
                        onChange={e => {
                          setSheetPhoneCols(prev => ({
                            ...prev,
                            [s.sheetName]: e.target.value,
                          }));
                        }}
                        className={`flex-1 bg-white border border-slate-300 rounded px-2 py-1 text-xs text-slate-800 font-medium focus:ring-1 focus:ring-emerald-500 ${
                          !isSelected ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : ''
                        }`}
                      >
                        {s.headers.map(h => (
                          <option key={h} value={h}>
                            {h}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Action Bar for applying updated sheet selections and phone columns */}
            <div className="mt-4 pt-3 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-3">
              <div className="text-xs text-slate-600">
                <span className="font-semibold text-slate-900">{selectedSheets.size}</span> sheets included in final download & matching.
                {detectedSheets.length - selectedSheets.size > 0 && (
                  <span className="text-slate-400 ml-1.5">
                    ({detectedSheets.length - selectedSheets.size} sheets excluded)
                  </span>
                )}
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleApplySheetSelectionAndColumns}
                  disabled={isLoading || selectedSheets.size === 0}
                  className="inline-flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 transition-all cursor-pointer shadow-sm"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isLoading ? 'animate-spin' : ''}`} />
                  <span>Apply & Re-analyze Selected Sheets ({selectedSheets.size})</span>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Results Section */}
        {results.length > 0 && summary && (
          <div className="space-y-6">
            {/* KPI Summary Cards */}
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                <div className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">Total Scanned</div>
                <div className="text-xl font-bold text-slate-900 mt-1">{filteredResults.length.toLocaleString()}</div>
                <div className="text-[11px] text-slate-500 mt-0.5">{selectedSheets.size} Sheet(s)</div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-emerald-200 shadow-2xs bg-emerald-50/30">
                <div className="text-[11px] font-semibold text-emerald-700 uppercase tracking-wider">Matched in DB</div>
                <div className="text-xl font-bold text-emerald-700 mt-1">
                  {filteredResults.filter(r => r.found).length.toLocaleString()}
                </div>
                <div className="text-[11px] text-emerald-600 font-medium mt-0.5">
                  {filteredResults.length > 0
                    ? `${((filteredResults.filter(r => r.found).length / filteredResults.length) * 100).toFixed(1)}% match rate`
                    : '0%'}
                </div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                <div className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">Unmatched (Valid)</div>
                <div className="text-xl font-bold text-slate-600 mt-1">
                  {filteredResults.filter(r => !r.found && !r.priority.includes('< 10') && !r.priority.includes('Invalid')).length.toLocaleString()}
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5">Valid 10-digit leads</div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-amber-200 shadow-2xs bg-amber-50/20">
                <div className="text-[11px] font-semibold text-amber-800 uppercase tracking-wider">Ignored (&lt;10 Digits)</div>
                <div className="text-xl font-bold text-amber-700 mt-1">
                  {filteredResults.filter(r => r.priority.includes('< 10') || r.priority.includes('Invalid')).length.toLocaleString()}
                </div>
                <div className="text-[11px] text-amber-600/80 mt-0.5">Excluded from DB check</div>
              </div>

              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                <div className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">Duration</div>
                <div className="text-xl font-bold text-slate-900 mt-1">{(summary.durationMs / 1000).toFixed(2)}s</div>
                <div className="text-[11px] text-slate-400 mt-0.5">Normalized search</div>
              </div>

              {/* Priority Highlights */}
              <div className="bg-white p-4 rounded-xl border border-emerald-200 shadow-2xs bg-emerald-50/20 col-span-2 sm:col-span-1 lg:col-span-1">
                <div className="text-[11px] font-semibold text-slate-700 uppercase tracking-wider mb-1 flex items-center justify-between">
                  <span>Top Priorities</span>
                </div>
                <div className="flex flex-wrap gap-1 max-h-12 overflow-y-auto">
                  {Object.entries(summary.priorityBreakdown)
                    .filter(([p]) => !p.includes('Unmatched') && !p.includes('Not Found') && !p.includes('Invalid'))
                    .slice(0, 3)
                    .map(([priority, count]) => (
                      <span
                        key={priority}
                        className="inline-flex items-center gap-1 text-[10px] font-medium px-1.5 py-0.5 rounded-full bg-slate-100 text-slate-700 border border-slate-200"
                      >
                        <span className="truncate max-w-[80px]">{priority}:</span>
                        <span className="font-bold text-emerald-700">{count}</span>
                      </span>
                    ))}
                </div>
              </div>
            </div>

            {/* Main Interactive Table Card */}
            <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
              {/* Table Toolbar */}
              <div className="p-4 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-slate-50/60">
                <div className="flex flex-wrap items-center gap-2">
                  {/* Status Filter Tabs */}
                  <div className="inline-flex p-1 bg-slate-200/80 rounded-lg text-xs font-medium">
                    <button
                      onClick={() => {
                        setFilterStatus('all');
                        setCurrentPage(1);
                      }}
                      className={`px-3 py-1 rounded-md transition-all ${
                        filterStatus === 'all' ? 'bg-white text-slate-900 shadow-2xs font-semibold' : 'text-slate-600'
                      }`}
                    >
                      All ({results.filter(r => selectedSheets.size === 0 || selectedSheets.has(r.sourceSheet)).length.toLocaleString()})
                    </button>
                    <button
                      onClick={() => {
                        setFilterStatus('matched');
                        setCurrentPage(1);
                      }}
                      className={`px-3 py-1 rounded-md transition-all ${
                        filterStatus === 'matched' ? 'bg-white text-emerald-700 shadow-2xs font-semibold' : 'text-slate-600'
                      }`}
                    >
                      Matched ({results.filter(r => (selectedSheets.size === 0 || selectedSheets.has(r.sourceSheet)) && r.found).length.toLocaleString()})
                    </button>
                    <button
                      onClick={() => {
                        setFilterStatus('unmatched');
                        setCurrentPage(1);
                      }}
                      className={`px-3 py-1 rounded-md transition-all ${
                        filterStatus === 'unmatched' ? 'bg-white text-slate-800 shadow-2xs font-semibold' : 'text-slate-600'
                      }`}
                    >
                      Unmatched ({results.filter(r => (selectedSheets.size === 0 || selectedSheets.has(r.sourceSheet)) && !r.found && !r.priority.includes('< 10') && !r.priority.includes('Invalid')).length.toLocaleString()})
                    </button>
                    <button
                      onClick={() => {
                        setFilterStatus('invalid');
                        setCurrentPage(1);
                      }}
                      className={`px-3 py-1 rounded-md transition-all ${
                        filterStatus === 'invalid' ? 'bg-white text-amber-800 shadow-2xs font-semibold' : 'text-slate-600'
                      }`}
                    >
                      &lt;10 Digits ({results.filter(r => (selectedSheets.size === 0 || selectedSheets.has(r.sourceSheet)) && (r.priority.includes('< 10') || r.priority.includes('Invalid'))).length.toLocaleString()})
                    </button>
                  </div>

                  {/* Filter by Source Sheet */}
                  {uniqueSourceSheets.length > 1 && (
                    <select
                      value={filterSheet}
                      onChange={e => {
                        setFilterSheet(e.target.value);
                        setCurrentPage(1);
                      }}
                      className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-700 font-medium focus:ring-1 focus:ring-emerald-500 shadow-2xs"
                    >
                      <option value="all">All Selected Sheets ({uniqueSourceSheets.length})</option>
                      {uniqueSourceSheets.map(s => (
                        <option key={s} value={s}>
                          Sheet: {s}
                        </option>
                      ))}
                    </select>
                  )}

                  {/* Filter by Priority */}
                  {uniquePriorities.length > 0 && (
                    <select
                      value={filterPriority}
                      onChange={e => {
                        setFilterPriority(e.target.value);
                        setCurrentPage(1);
                      }}
                      className="bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-700 font-medium focus:ring-1 focus:ring-emerald-500 shadow-2xs"
                    >
                      <option value="all">All Priorities</option>
                      {uniquePriorities.map(p => (
                        <option key={p} value={p}>
                          {p}
                        </option>
                      ))}
                    </select>
                  )}
                </div>

                {/* Right side: Search Box & Download Button */}
                <div className="flex flex-wrap items-center gap-3">
                  <div className="relative">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={e => {
                        setSearchQuery(e.target.value);
                        setCurrentPage(1);
                      }}
                      placeholder="Search phone, name, workbook..."
                      className="w-48 sm:w-56 pl-9 pr-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs text-slate-800 placeholder-slate-400 focus:outline-hidden focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 shadow-2xs"
                    />
                  </div>

                  {/* Export Button: Creates Combined Sheet + individual selected sheets */}
                  <button
                    type="button"
                    onClick={handleExportExcel}
                    disabled={isExporting}
                    className={`inline-flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-bold text-white shadow-sm transition-all shrink-0 cursor-pointer ${
                      exportSuccess
                        ? 'bg-emerald-700 hover:bg-emerald-800 ring-2 ring-emerald-400'
                        : 'bg-emerald-600 hover:bg-emerald-700 active:scale-95'
                    } disabled:opacity-50 disabled:cursor-not-allowed`}
                    title={`Export ${selectedSheets.size} selected sheet(s) and Combined Master Sheet`}
                  >
                    {isExporting ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin text-white" />
                        <span>Generating Excel...</span>
                      </>
                    ) : exportSuccess ? (
                      <>
                        <CheckCircle2 className="w-4 h-4 text-emerald-200" />
                        <span>Downloaded Excel!</span>
                      </>
                    ) : (
                      <>
                        <Download className="w-4 h-4" />
                        <span>Download Final Excel ({selectedSheets.size} Sheets)</span>
                      </>
                    )}
                  </button>
                </div>
              </div>

              {/* Export error banner */}
              {exportError && (
                <div className="px-4 py-2.5 bg-rose-50 border-b border-rose-200 flex items-center justify-between text-xs text-rose-800">
                  <div className="flex items-center gap-2">
                    <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
                    <span>{exportError}</span>
                  </div>
                  <button
                    onClick={() => setExportError(null)}
                    className="text-rose-500 hover:text-rose-700"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              )}

              {/* Table Body */}
              <div className="overflow-x-auto min-h-[350px]">
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-100/70 text-slate-700 font-semibold uppercase tracking-wider text-[10px]">
                      <th className="py-2.5 px-3 w-10 text-center">#</th>
                      <th className="py-2.5 px-3">Phone Number</th>
                      <th className="py-2.5 px-3">Sheet Name</th>
                      <th className="py-2.5 px-3">Priority</th>
                      <th className="py-2.5 px-3">Sheet Name Database</th>
                      <th className="py-2.5 px-3">Workbook Name</th>
                      <th className="py-2.5 px-3">Customer Name (Database)</th>
                      <th className="py-2.5 px-3">Lender / Amount</th>
                      <th className="py-2.5 px-3 text-center">Status</th>
                      <th className="py-2.5 px-3 text-center">Details</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-slate-700">
                    {paginatedRows.length === 0 ? (
                      <tr>
                        <td colSpan={10} className="py-12 text-center text-slate-400">
                          No matching records found for the selected filter or sheets.
                        </td>
                      </tr>
                    ) : (
                      paginatedRows.map((row, idx) => {
                        const globalIndex = (currentPage - 1) * pageSize + idx + 1;
                        const isExpanded = expandedRowIndex === globalIndex;

                        return (
                          <React.Fragment key={`${row.sourceSheet}_${row.sourceRowNum}_${idx}`}>
                            <tr className={`hover:bg-slate-50/80 transition-colors ${row.found ? '' : 'bg-slate-50/30'}`}>
                              <td className="py-2.5 px-3 text-center text-slate-400 font-mono text-[11px]">
                                {globalIndex}
                              </td>

                              <td className="py-2.5 px-3 font-mono font-medium text-slate-900">
                                <div className="flex items-center gap-1.5">
                                  <span>{row.inputPhone || '—'}</span>
                                  {row.inputPhone && (
                                    <button
                                      onClick={() => handleCopy(row.inputPhone)}
                                      className="text-slate-400 hover:text-slate-600 transition-colors cursor-pointer"
                                      title="Copy phone"
                                    >
                                      {copiedPhone === row.inputPhone ? (
                                        <Check className="w-3 h-3 text-emerald-600" />
                                      ) : (
                                        <Copy className="w-3 h-3" />
                                      )}
                                    </button>
                                  )}
                                </div>
                              </td>

                              <td className="py-2.5 px-3">
                                <span className="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-mono bg-blue-50 text-blue-700 border border-blue-200">
                                  {row.sourceSheet}
                                </span>
                              </td>

                              <td className="py-2.5 px-3">
                                {row.found ? (
                                  <span
                                    className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-semibold ${
                                      row.badgeBg || 'bg-emerald-50'
                                    } ${row.badgeText || 'text-emerald-700'} border ${
                                      row.badgeBorder || 'border-emerald-200'
                                    }`}
                                  >
                                    <Sparkles className="w-2.5 h-2.5" />
                                    <span>{row.priority}</span>
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-medium bg-slate-100 text-slate-500 border border-slate-200">
                                    Not Found
                                  </span>
                                )}
                              </td>

                              <td className="py-2.5 px-3">
                                {row.found ? (
                                  <button
                                    type="button"
                                    onClick={() =>
                                      onPreviewSheet?.(
                                        row.firstFileName || row.foundWorkbook.split(',')[0].trim(),
                                        row.firstSheetName || row.foundSheet.split(',')[0].trim()
                                      )
                                    }
                                    className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium bg-slate-100 hover:bg-emerald-50 text-slate-800 hover:text-emerald-700 border border-slate-200 hover:border-emerald-300 transition-colors cursor-pointer"
                                    title="Click to preview worksheet"
                                  >
                                    <Eye className="w-2.5 h-2.5 text-slate-400" />
                                    <span>{row.foundSheet}</span>
                                  </button>
                                ) : (
                                  <span className="text-slate-400">—</span>
                                )}
                              </td>

                              <td className="py-2.5 px-3 font-medium text-slate-900">
                                {row.found ? (
                                  <div className="flex items-center gap-1.5">
                                    <FileSpreadsheet className="w-3.5 h-3.5 text-blue-600 shrink-0" />
                                    <span className="font-semibold text-slate-800">{row.foundWorkbook}</span>
                                  </div>
                                ) : (
                                  <span className="text-slate-400">—</span>
                                )}
                              </td>

                              <td className="py-2.5 px-3 font-medium text-slate-800">
                                {row.foundCustomerName !== '—' ? (
                                  <span>{row.foundCustomerName}</span>
                                ) : (
                                  <span className="text-slate-400">—</span>
                                )}
                              </td>

                              <td className="py-2.5 px-3 text-slate-600">
                                {row.foundLender !== '—' || row.foundLoanAmount !== '—' ? (
                                  <div className="truncate max-w-[150px]">
                                    {row.foundLender !== '—' && <span className="font-medium">{row.foundLender} </span>}
                                    {row.foundLoanAmount !== '—' && (
                                      <span className="text-emerald-700 font-semibold">({row.foundLoanAmount})</span>
                                    )}
                                  </div>
                                ) : (
                                  <span className="text-slate-400">—</span>
                                )}
                              </td>

                              <td className="py-2.5 px-3 text-center">
                                {row.found ? (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200">
                                    <CheckCircle2 className="w-2.5 h-2.5 text-emerald-600" />
                                    <span>MATCHED</span>
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full border border-slate-200">
                                    <span>NOT FOUND</span>
                                  </span>
                                )}
                              </td>

                              <td className="py-2.5 px-3 text-center">
                                <div className="flex items-center justify-center gap-1">
                                  {row.found && row.firstRecordId && (
                                    <button
                                      type="button"
                                      onClick={() => onViewRecord?.(row.firstRecordId!)}
                                      className="px-2 py-1 rounded bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 text-[10px] font-bold transition-colors cursor-pointer"
                                      title="Inspect matched database record"
                                    >
                                      Row
                                    </button>
                                  )}
                                  <button
                                    onClick={() => setExpandedRowIndex(isExpanded ? null : globalIndex)}
                                    className="text-slate-400 hover:text-slate-600 p-1 rounded hover:bg-slate-100 transition-colors cursor-pointer"
                                    title="View full original row"
                                  >
                                    {isExpanded ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />}
                                  </button>
                                </div>
                              </td>
                            </tr>

                            {/* Expanded Row Drawer showing original sheet data */}
                            {isExpanded && (
                              <tr className="bg-slate-50/90 border-b border-slate-200">
                                <td colSpan={10} className="p-4">
                                  <div className="bg-white p-3.5 rounded-lg border border-slate-200 shadow-2xs">
                                    <div className="flex flex-wrap items-center justify-between gap-2 text-xs font-semibold text-slate-800 mb-2 border-b border-slate-100 pb-2">
                                      <div className="flex items-center gap-2">
                                        <FileText className="w-3.5 h-3.5 text-emerald-600" />
                                        <span>
                                          Original Row Data from Sheet: <span className="font-mono text-emerald-700">{row.sourceSheet}</span> (Row {row.sourceRowNum})
                                        </span>
                                      </div>
                                      <div className="flex flex-wrap items-center gap-2">
                                        {row.foundRowNum !== '—' && (
                                          <span className="text-[11px] font-normal text-slate-500">
                                            Database Match: {row.foundRowNum}
                                          </span>
                                        )}
                                        {row.found && row.firstRecordId && (
                                          <button
                                            type="button"
                                            onClick={() => onViewRecord?.(row.firstRecordId!)}
                                            className="px-2.5 py-1 rounded-md bg-blue-600 hover:bg-blue-700 text-white text-[11px] font-bold transition-colors cursor-pointer"
                                          >
                                            Inspect Matched Database Row
                                          </button>
                                        )}
                                        {row.found && (row.firstFileName || row.foundWorkbook !== '—') && (
                                          <button
                                            type="button"
                                            onClick={() =>
                                              onPreviewSheet?.(
                                                row.firstFileName || row.foundWorkbook.split(',')[0].trim(),
                                                row.firstSheetName || row.foundSheet.split(',')[0].trim()
                                              )
                                            }
                                            className="px-2.5 py-1 rounded-md bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-[11px] font-semibold transition-colors cursor-pointer"
                                          >
                                            Preview Database Sheet
                                          </button>
                                        )}
                                        {onLookupCustomer && row.inputPhone && (
                                          <button
                                            type="button"
                                            onClick={() => onLookupCustomer(row.inputPhone)}
                                            className="px-2.5 py-1 rounded-md bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 text-[11px] font-bold transition-colors cursor-pointer"
                                          >
                                            Customer Lookup &amp; Diff
                                          </button>
                                        )}
                                      </div>
                                    </div>

                                    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-2.5 text-xs">
                                      {row.originalData &&
                                        Object.entries(row.originalData)
                                          .filter(([k]) => !k.startsWith('_'))
                                          .map(([key, value]) => (
                                            <div key={key} className="p-2 bg-slate-50 rounded border border-slate-100">
                                              <div className="text-[10px] text-slate-400 uppercase font-medium truncate" title={key}>
                                                {key}
                                              </div>
                                              <div className="text-slate-800 font-medium truncate mt-0.5" title={String(value)}>
                                                {String(value || '—')}
                                              </div>
                                            </div>
                                          ))}
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>

              {/* Pagination Footer */}
              <div className="p-4 border-t border-slate-200 bg-slate-50/70 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-600">
                <div>
                  Showing{' '}
                  <span className="font-semibold text-slate-800">
                    {filteredResults.length === 0 ? 0 : (currentPage - 1) * pageSize + 1}
                  </span>{' '}
                  to{' '}
                  <span className="font-semibold text-slate-800">
                    {Math.min(currentPage * pageSize, filteredResults.length)}
                  </span>{' '}
                  of <span className="font-semibold text-slate-800">{filteredResults.length.toLocaleString()}</span> entries
                </div>

                <div className="flex items-center gap-2">
                  <span className="text-slate-500">Rows per page:</span>
                  <select
                    value={pageSize}
                    onChange={e => {
                      setPageSize(Number(e.target.value));
                      setCurrentPage(1);
                    }}
                    className="bg-white border border-slate-300 rounded px-2 py-1 text-xs"
                  >
                    <option value={25}>25</option>
                    <option value={50}>50</option>
                    <option value={100}>100</option>
                    <option value={250}>250</option>
                  </select>

                  <div className="flex items-center gap-1 ml-2">
                    <button
                      onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                      disabled={currentPage <= 1}
                      className="px-2.5 py-1 rounded bg-white border border-slate-300 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-50 transition-colors shadow-2xs cursor-pointer"
                    >
                      Prev
                    </button>
                    <span className="px-2 py-1 text-slate-700 font-medium">
                      {currentPage} / {totalPages}
                    </span>
                    <button
                      onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                      disabled={currentPage >= totalPages}
                      className="px-2.5 py-1 rounded bg-white border border-slate-300 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-50 transition-colors shadow-2xs cursor-pointer"
                    >
                      Next
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Priority Configuration Modal */}
      {isPriorityModalOpen && (
        <PriorityConfigModal
          isOpen={isPriorityModalOpen}
          openCombineOnOpen={priorityOpenCombine}
          onClose={() => {
            setIsPriorityModalOpen(false);
            // Refresh priority rules
            fetch('/api/priority-rules')
              .then(res => res.json())
              .then(data => {
                if (data.rules) setPriorityRules(data.rules);
              })
              .catch(console.error);
          }}
          onRulesUpdated={updated => {
            setPriorityRules(updated);
          }}
        />
      )}

    </div>
  );
};
