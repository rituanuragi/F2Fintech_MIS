import React, { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload, Loader2, AlertTriangle, Download, ChevronDown, ChevronRight, FileSpreadsheet, Ticket,
  FolderTree, Layers, Search, Maximize2, UserSearch, X,
} from 'lucide-react';
import { apiFetch } from '../utils/api';

/* ------------------------------------------------------------------ types (match server/batchCustomerLookup.ts) */
interface MisEntry {
  id: number; fileId: string; fileName: string; sheetName: string; rowNum: number; sheetRank: number;
  customerName: string; lender: string; loanAmount: string; status: string; ticketId: string;
  loginDate: string; approvalDate: string; disbursedDate: string; teamLeader: string; manager: string;
  employee: string; city: string; state: string; cibil: string; pan: string; opsRemarks: string; creditRemarks: string;
}
interface DirEntry {
  id: number; folderPath: string; fileName: string; sheetName: string; rowNum: number; customerName: string;
  priority: string; rank: number; badge: { bg?: string; text?: string; border?: string };
}
interface LiveTicket {
  ticketId: string | number; formattedTicketId?: string; customerName?: string; customerContact?: string;
  applicationProvider?: string; standardizedProvider?: string; ticketStatus?: string; loanStatus?: string;
  applicationAmount?: string; applicationDate?: string; matchScore?: number; matchReasons?: string[];
  matchedParameters?: { phoneMatched: boolean; lenderMatched: boolean; nameMatched: boolean; ticketIdMatched: boolean; amountMatched: boolean };
}
interface BatchResult {
  inputPhone: string; normPhone: string; inputCount: number; sources: string[];
  status: 'in_mis' | 'not_in_mis' | 'invalid';
  mis: null | {
    entryCount: number; workbookCount: number; tabCount: number;
    books: { fileName: string; sheetName: string; entries: number; rank: number }[];
    statuses: Record<string, number>; lenders: Record<string, number>;
    primary: MisEntry; entries: MisEntry[]; entriesHidden: number;
  };
  directory: null | { entryCount: number; fileCount: number; entries: DirEntry[]; entriesHidden: number };
  api: null | { matched: boolean; ticket: LiveTicket | null };
}
interface BatchSummary {
  rowsReceived: number; uniqueNumbers: number; duplicatesInInput: number; invalid: number; inMis: number; notInMis: number;
  totalMisEntries: number; inMisAndDirectory: number; inMisAndApi: number; inAllThree: number; inMisOnly: number; multipleMisEntries: number;
}
interface ParsedSheet {
  name: string; headers: string[]; rows: any[][]; firstDataExcelRow: number; phoneIndex: number; include: boolean;
}

interface Props {
  onOpenRowModal?: (id: number) => void;
  onNavigateToTicket?: (ticketId: string | number) => void;
  onLookupSingle?: (phone: string) => void;
}

/* ------------------------------------------------------------------ helpers */
type Tier = 'all3' | 'mis_dir' | 'mis_api' | 'mis_only' | 'not_in_mis' | 'invalid';
const tierOf = (r: BatchResult): Tier => {
  if (r.status === 'invalid') return 'invalid';
  if (r.status === 'not_in_mis') return 'not_in_mis';
  const dir = (r.directory?.entryCount || 0) > 0;
  const api = !!r.api?.matched;
  if (dir && api) return 'all3';
  if (dir) return 'mis_dir';
  if (api) return 'mis_api';
  return 'mis_only';
};

const TIERS: { id: Tier; label: string; bar: string; chip: string }[] = [
  { id: 'all3', label: 'MIS + Directory + API', bar: 'bg-emerald-500', chip: 'bg-emerald-50 text-emerald-700 border-emerald-200' },
  { id: 'mis_dir', label: 'MIS + Directory', bar: 'bg-blue-500', chip: 'bg-blue-50 text-blue-700 border-blue-200' },
  { id: 'mis_api', label: 'MIS + API', bar: 'bg-indigo-500', chip: 'bg-indigo-50 text-indigo-700 border-indigo-200' },
  { id: 'mis_only', label: 'MIS only', bar: 'bg-slate-400', chip: 'bg-slate-100 text-slate-700 border-slate-200' },
  { id: 'not_in_mis', label: 'Not in MIS', bar: 'bg-rose-500', chip: 'bg-rose-50 text-rose-700 border-rose-200' },
  { id: 'invalid', label: 'Invalid number', bar: 'bg-amber-400', chip: 'bg-amber-50 text-amber-700 border-amber-200' },
];

type Filter = 'all' | 'in_mis' | 'multi' | Tier;

const show = (v: any) => {
  const s = v === null || v === undefined ? '' : String(v).trim();
  return s || '—';
};
const cellText = (v: any): string => {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v).trim();
};
// Stop Excel from running a cell as a formula
const safe = (v: any) => {
  const s = v === null || v === undefined ? '' : String(v);
  return /^[=+\-@\t\r]/.test(s) ? `'${s}` : s;
};
const statusTone = (s?: string) => {
  const t = (s || '').toUpperCase();
  if (/DISBURS/.test(t)) return 'bg-emerald-50 text-emerald-700 border-emerald-200';
  if (/APPROV/.test(t)) return 'bg-blue-50 text-blue-700 border-blue-200';
  if (/REJECT|DECLIN|CANCEL/.test(t)) return 'bg-rose-50 text-rose-700 border-rose-200';
  return 'bg-slate-100 text-slate-600 border-slate-200';
};
const topEntry = (m: Record<string, number>) => Object.entries(m).sort((a, b) => b[1] - a[1])[0]?.[0] || '';

const PAGE = 50;

/* ------------------------------------------------------------------ component */
export const BatchCustomerLookup: React.FC<Props> = ({ onOpenRowModal, onNavigateToTicket, onLookupSingle }) => {
  const fileInput = useRef<HTMLInputElement>(null);
  const [pasted, setPasted] = useState('');
  const [fileName, setFileName] = useState('');
  const [sheets, setSheets] = useState<ParsedSheet[]>([]);
  const [reading, setReading] = useState(false);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const [results, setResults] = useState<BatchResult[]>([]);
  const [summary, setSummary] = useState<BatchSummary | null>(null);
  const [durationMs, setDurationMs] = useState(0);

  const [filter, setFilter] = useState<Filter>('all');
  const [text, setText] = useState('');
  const [lenderFilter, setLenderFilter] = useState('all');
  const [statusFilter, setStatusFilter] = useState('all');
  const [page, setPage] = useState(1);
  const [open, setOpen] = useState<Set<string>>(new Set());

  /* ---------- reading an Excel / CSV file ---------- */
  const handleFile = async (file: File) => {
    setReading(true);
    setError(null);
    setFileName(file.name);
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
      const parsed: ParsedSheet[] = [];
      for (const name of wb.SheetNames) {
        const ws = wb.Sheets[name];
        if (!ws) continue;
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: true }) as any[][];
        if (raw.length === 0) continue;
        // the server decides where the header row and the phone column are (one rule for every screen)
        const det = await apiFetch<{ headerRowIndex: number; headers: string[]; index: number }>('/api/phone-column/detect', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ rows: raw.slice(0, 40) }),
        });
        const dataRows = raw.slice(det.headerRowIndex + 1);
        if (dataRows.every((r) => !r || r.every((c) => cellText(c) === ''))) continue;
        parsed.push({
          name, headers: det.headers, rows: dataRows, firstDataExcelRow: det.headerRowIndex + 2,
          phoneIndex: det.index, include: true,
        });
      }
      if (parsed.length === 0) throw new Error('No sheet with data rows was found in this file.');
      setSheets(parsed);
    } catch (e: any) {
      setSheets([]);
      setFileName('');
      setError(e.message || 'Could not read this file.');
    } finally {
      setReading(false);
    }
  };

  const updateSheet = (name: string, patch: Partial<ParsedSheet>) =>
    setSheets((prev) => prev.map((s) => (s.name === name ? { ...s, ...patch } : s)));

  const pastedNumbers = useMemo(() => pasted.split(/[\s,;|]+/).map((t) => t.trim()).filter(Boolean), [pasted]);
  const fileNumbers = useMemo(() => {
    const out: { rawPhone: string; sourceSheet: string; sourceRow: number }[] = [];
    for (const s of sheets) {
      if (!s.include) continue;
      s.rows.forEach((row, i) => {
        const v = cellText(row?.[s.phoneIndex]);
        if (v) out.push({ rawPhone: v, sourceSheet: s.name, sourceRow: s.firstDataExcelRow + i });
      });
    }
    return out;
  }, [sheets]);
  const totalInput = pastedNumbers.length + fileNumbers.length;

  /* ---------- run ---------- */
  const run = async () => {
    setRunning(true);
    setError(null);
    setResults([]);
    setSummary(null);
    setOpen(new Set());
    setPage(1);
    setFilter('all');
    try {
      const phones = [
        ...pastedNumbers.map((p) => ({ rawPhone: p, sourceSheet: 'Pasted', sourceRow: undefined as number | undefined })),
        ...fileNumbers,
      ];
      const out = await apiFetch<{ results: BatchResult[]; summary: BatchSummary; durationMs: number }>('/api/records/batch-lookup', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ phones }),
        timeoutMs: 300000,
      });
      setResults(out.results);
      setSummary(out.summary);
      setDurationMs(out.durationMs);
    } catch (e: any) {
      setError(e.message || 'Batch lookup failed.');
    } finally {
      setRunning(false);
    }
  };

  /* ---------- derived numbers ---------- */
  const tierCounts = useMemo(() => {
    const c: Record<Tier, number> = { all3: 0, mis_dir: 0, mis_api: 0, mis_only: 0, not_in_mis: 0, invalid: 0 };
    results.forEach((r) => (c[tierOf(r)] += 1));
    return c;
  }, [results]);

  const lenderOptions = useMemo(
    () => Array.from(new Set(results.filter((r) => r.mis).map((r) => (r.mis!.primary.lender || '').toUpperCase()).filter(Boolean))).sort(),
    [results]
  );
  const statusOptions = useMemo(
    () => Array.from(new Set(results.filter((r) => r.mis).map((r) => (r.mis!.primary.status || '').toUpperCase()).filter(Boolean))).sort(),
    [results]
  );

  const visible = useMemo(() => {
    const q = text.trim().toLowerCase();
    return results.filter((r) => {
      if (filter === 'in_mis' && r.status !== 'in_mis') return false;
      if (filter === 'multi' && !((r.mis?.entryCount || 0) > 1)) return false;
      if (['all3', 'mis_dir', 'mis_api', 'mis_only', 'not_in_mis', 'invalid'].includes(filter) && tierOf(r) !== filter) return false;
      if (lenderFilter !== 'all' && (r.mis?.primary.lender || '').toUpperCase() !== lenderFilter) return false;
      if (statusFilter !== 'all' && (r.mis?.primary.status || '').toUpperCase() !== statusFilter) return false;
      if (q) {
        const hay = [r.inputPhone, r.normPhone, r.mis?.primary.customerName, r.mis?.primary.ticketId].join(' ').toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [results, filter, text, lenderFilter, statusFilter]);

  const pageCount = Math.max(1, Math.ceil(visible.length / PAGE));
  const pageRows = visible.slice((page - 1) * PAGE, page * PAGE);
  const toggle = (k: string) =>
    setOpen((prev) => {
      const n = new Set(prev);
      n.has(k) ? n.delete(k) : n.add(k);
      return n;
    });

  /* ---------- export (4 sheets, formula-safe) ---------- */
  const exportExcel = () => {
    const list = visible;
    const numbers = list.map((r) => {
      const p = r.mis?.primary;
      const t = r.api?.ticket;
      return {
        Phone: safe(r.normPhone || r.inputPhone),
        Result: r.status === 'in_mis' ? 'Found in MIS' : r.status === 'not_in_mis' ? 'Not in MIS' : 'Invalid number',
        Coverage: TIERS.find((x) => x.id === tierOf(r))?.label || '',
        'Times in your list': r.inputCount,
        'Where in your list': safe(r.sources.join(', ')),
        'MIS entries': r.mis?.entryCount ?? 0,
        'MIS workbooks': r.mis?.workbookCount ?? 0,
        'MIS tabs': r.mis?.tabCount ?? 0,
        'Primary workbook': safe(p?.fileName),
        'Primary tab': safe(p?.sheetName),
        'Primary Excel row': p?.rowNum ?? '',
        Name: safe(p?.customerName),
        Lender: safe(p?.lender),
        'Loan amount': safe(p?.loanAmount),
        'MIS status': safe(p?.status),
        'Login date': safe(p?.loginDate),
        'MIS ticket ID': safe(p?.ticketId),
        'Directory entries': r.directory?.entryCount ?? 0,
        'Directory files': r.directory?.fileCount ?? 0,
        'Top directory priority': safe(r.directory?.entries[0]?.priority),
        'Live ticket found': r.api?.matched ? 'Yes' : r.status === 'in_mis' ? 'No' : '',
        'Live ticket ID': safe(t?.formattedTicketId || t?.ticketId),
        'Live ticket status': safe(t?.ticketStatus || t?.loanStatus),
        'Live provider': safe(t?.applicationProvider || t?.standardizedProvider),
      };
    });
    const misRows = list.flatMap((r) =>
      (r.mis?.entries || []).map((e) => ({
        Phone: safe(r.normPhone), Workbook: safe(e.fileName), Tab: safe(e.sheetName), 'Excel row': e.rowNum,
        Name: safe(e.customerName), Lender: safe(e.lender), 'Loan amount': safe(e.loanAmount), Status: safe(e.status),
        'Ticket ID': safe(e.ticketId), 'Login date': safe(e.loginDate), 'Approval date': safe(e.approvalDate),
        'Disbursed date': safe(e.disbursedDate), 'Team leader': safe(e.teamLeader), Manager: safe(e.manager),
        Employee: safe(e.employee), City: safe(e.city), State: safe(e.state), CIBIL: safe(e.cibil), PAN: safe(e.pan),
        'Ops remarks': safe(e.opsRemarks), 'Credit remarks': safe(e.creditRemarks),
      }))
    );
    const dirRows = list.flatMap((r) =>
      (r.directory?.entries || []).map((d) => ({
        Phone: safe(r.normPhone), Folder: safe(d.folderPath), Workbook: safe(d.fileName), Sheet: safe(d.sheetName),
        'Excel row': d.rowNum, Name: safe(d.customerName), Priority: safe(d.priority),
      }))
    );
    const apiRows = list
      .filter((r) => r.api?.ticket)
      .map((r) => {
        const t = r.api!.ticket!;
        return {
          Phone: safe(r.normPhone), 'Ticket ID': safe(t.formattedTicketId || t.ticketId), Name: safe(t.customerName),
          Provider: safe(t.applicationProvider || t.standardizedProvider), Status: safe(t.ticketStatus || t.loanStatus),
          Amount: safe(t.applicationAmount), 'Match score': t.matchScore ?? '', 'Why it matched': safe((t.matchReasons || []).join('; ')),
        };
      });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(numbers), 'Numbers');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(misRows), 'MIS entries');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dirRows), 'Directory');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(apiRows), 'Live tickets');
    XLSX.writeFile(wb, `F2_Batch_Customer_Lookup_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const clearAll = () => {
    setPasted('');
    setSheets([]);
    setFileName('');
    setResults([]);
    setSummary(null);
    setError(null);
  };

  /* ================================================================== render */
  return (
    <div className="space-y-6">
      {/* ---------- input ---------- */}
      <div className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <h3 className="text-base font-black text-slate-900 tracking-tight">Batch lookup</h3>
            <p className="text-xs text-slate-500 mt-1 font-medium max-w-3xl">
              Paste many numbers or upload an Excel file. Every number is checked in the MIS first. Numbers that are not in the MIS stop there.
              Numbers that are in the MIS are then checked in the main directory files and in the live tickets.
            </p>
          </div>
          {(totalInput > 0 || results.length > 0) && (
            <button onClick={clearAll} className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 text-xs font-bold text-slate-600 hover:bg-slate-50 cursor-pointer">
              <X className="w-3.5 h-3.5" /> Clear
            </button>
          )}
        </div>

        <div className="grid lg:grid-cols-2 gap-4">
          <div>
            <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wide">Paste numbers</label>
            <textarea
              value={pasted}
              onChange={(e) => setPasted(e.target.value)}
              rows={6}
              placeholder={'8830636983\n9876543210, 9123456780'}
              className="mt-1.5 w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl text-sm font-mono text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-500 focus:bg-white focus:ring-1 focus:ring-blue-500"
            />
            <div className="text-[11px] text-slate-500 mt-1 font-medium">{pastedNumbers.length.toLocaleString()} numbers pasted. Separate with new lines, spaces or commas.</div>
          </div>

          <div>
            <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wide">Or upload a file</label>
            <input
              ref={fileInput}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) handleFile(f);
                e.target.value = '';
              }}
            />
            <button
              onClick={() => fileInput.current?.click()}
              disabled={reading || running}
              className="mt-1.5 w-full h-[148px] border-2 border-dashed border-slate-300 hover:border-blue-400 hover:bg-blue-50/40 rounded-xl flex flex-col items-center justify-center gap-2 text-slate-600 disabled:opacity-60 cursor-pointer"
            >
              {reading ? <Loader2 className="w-6 h-6 animate-spin" /> : <Upload className="w-6 h-6 text-blue-600" />}
              <span className="text-sm font-bold">{reading ? 'Reading file…' : fileName ? 'Choose a different file' : 'Choose an Excel or CSV file'}</span>
              {fileName && !reading && (
                <span className="text-[11px] text-slate-500 inline-flex items-center gap-1 font-medium">
                  <FileSpreadsheet className="w-3.5 h-3.5" /> {fileName} · {fileNumbers.length.toLocaleString()} numbers
                </span>
              )}
            </button>
          </div>
        </div>

        {sheets.length > 0 && (
          <div className="overflow-x-auto border border-slate-200 rounded-xl">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 text-slate-500 text-left text-[11px] uppercase tracking-wide">
                <tr>
                  <th className="px-3 py-2 font-bold">Use</th>
                  <th className="px-3 py-2 font-bold">Sheet</th>
                  <th className="px-3 py-2 font-bold text-right">Rows</th>
                  <th className="px-3 py-2 font-bold">Phone column</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {sheets.map((s) => (
                  <tr key={s.name}>
                    <td className="px-3 py-2"><input type="checkbox" checked={s.include} onChange={(e) => updateSheet(s.name, { include: e.target.checked })} /></td>
                    <td className="px-3 py-2 font-bold text-slate-800">{s.name}</td>
                    <td className="px-3 py-2 text-right tabular-nums">{s.rows.length.toLocaleString()}</td>
                    <td className="px-3 py-2">
                      <select value={s.phoneIndex} onChange={(e) => updateSheet(s.name, { phoneIndex: Number(e.target.value) })} className="px-2 py-1.5 rounded-lg border border-slate-300 bg-white text-xs">
                        {s.headers.map((h, i) => (
                          <option key={i} value={i}>{h || `Column ${i + 1}`}</option>
                        ))}
                      </select>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        <div className="flex flex-wrap items-center gap-3">
          <button
            onClick={run}
            disabled={running || totalInput === 0}
            className="px-6 py-3 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white rounded-xl text-sm font-bold shadow-sm flex items-center gap-2 cursor-pointer"
          >
            {running ? <Loader2 className="w-4 h-4 animate-spin" /> : <Search className="w-4 h-4" />}
            {running ? 'Checking…' : `Check ${totalInput.toLocaleString()} numbers`}
          </button>
          <span className="text-[11px] text-slate-500 font-medium">Up to 5,000 numbers per batch.</span>
        </div>
      </div>

      {error && (
        <div className="bg-rose-50 border border-rose-200 text-rose-700 rounded-2xl p-4 text-sm flex items-start gap-2">
          <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5" /> <span>{error}</span>
        </div>
      )}

      {/* ---------- results ---------- */}
      {summary && (
        <div className="space-y-5">
          {/* KPI strip */}
          <div className="grid grid-cols-2 md:grid-cols-4 xl:grid-cols-6 gap-3">
            {[
              ['Numbers checked', summary.uniqueNumbers, 'text-slate-900'],
              ['Found in MIS', summary.inMis, 'text-emerald-700'],
              ['Not in MIS', summary.notInMis, 'text-rose-700'],
              ['MIS entries in total', summary.totalMisEntries, 'text-blue-700'],
              ['More than one MIS entry', summary.multipleMisEntries, 'text-amber-700'],
              ['Also in directory', summary.inMisAndDirectory, 'text-blue-700'],
              ['Also in live tickets', summary.inMisAndApi, 'text-indigo-700'],
              ['In all three', summary.inAllThree, 'text-emerald-700'],
              ['MIS only', summary.inMisOnly, 'text-slate-700'],
              ['Invalid numbers', summary.invalid, 'text-amber-700'],
              ['Repeated in your list', summary.duplicatesInInput, 'text-slate-700'],
              ['Time taken', `${(durationMs / 1000).toFixed(1)}s`, 'text-slate-700'],
            ].map(([label, value, tone]) => (
              <div key={label as string} className="bg-white border border-slate-200/90 rounded-xl px-4 py-3">
                <div className="text-[11px] font-bold text-slate-500">{label}</div>
                <div className={`text-xl font-black tabular-nums ${tone}`}>{typeof value === 'number' ? value.toLocaleString() : value}</div>
              </div>
            ))}
          </div>

          {/* coverage bar */}
          <div className="bg-white border border-slate-200/90 rounded-2xl p-5 space-y-3">
            <div className="text-xs font-black text-slate-800">Where each number was found</div>
            <div className="flex h-3 rounded-full overflow-hidden bg-slate-100">
              {TIERS.map((t) => {
                const n = tierCounts[t.id];
                if (!n) return null;
                return <div key={t.id} className={t.bar} style={{ width: `${(n / results.length) * 100}%` }} title={`${t.label}: ${n}`} />;
              })}
            </div>
            <div className="flex flex-wrap gap-2">
              {TIERS.map((t) => (
                <button
                  key={t.id}
                  onClick={() => { setFilter(filter === t.id ? 'all' : t.id); setPage(1); }}
                  className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-full border text-xs font-bold cursor-pointer ${t.chip} ${filter === t.id ? 'ring-2 ring-offset-1 ring-blue-500' : ''}`}
                >
                  <span className={`w-2 h-2 rounded-full ${t.bar}`} />
                  {t.label}
                  <span className="tabular-nums">{tierCounts[t.id].toLocaleString()}</span>
                </button>
              ))}
            </div>
          </div>

          {/* filters */}
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex rounded-xl border border-slate-200 overflow-hidden bg-white">
              {([['all', 'All'], ['in_mis', 'In MIS'], ['multi', 'Several MIS entries'], ['not_in_mis', 'Not in MIS']] as [Filter, string][]).map(([id, label]) => (
                <button
                  key={id}
                  onClick={() => { setFilter(id); setPage(1); }}
                  className={`px-3.5 py-2 text-xs font-bold cursor-pointer ${filter === id ? 'bg-blue-600 text-white' : 'text-slate-700 hover:bg-slate-50'}`}
                >
                  {label}
                </button>
              ))}
            </div>
            <div className="relative flex-1 min-w-[200px]">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                value={text}
                onChange={(e) => { setText(e.target.value); setPage(1); }}
                placeholder="Filter by number, name or ticket ID"
                className="w-full pl-9 pr-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:border-blue-500"
              />
            </div>
            <select value={lenderFilter} onChange={(e) => { setLenderFilter(e.target.value); setPage(1); }} className="px-3 py-2 rounded-xl border border-slate-200 bg-white text-xs font-bold">
              <option value="all">All lenders</option>
              {lenderOptions.map((l) => <option key={l} value={l}>{l}</option>)}
            </select>
            <select value={statusFilter} onChange={(e) => { setStatusFilter(e.target.value); setPage(1); }} className="px-3 py-2 rounded-xl border border-slate-200 bg-white text-xs font-bold">
              <option value="all">All MIS statuses</option>
              {statusOptions.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
            <button onClick={exportExcel} className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold cursor-pointer">
              <Download className="w-4 h-4" /> Export {visible.length.toLocaleString()} to Excel
            </button>
          </div>

          {/* table */}
          <div className="bg-white border border-slate-200/90 rounded-2xl overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead className="bg-slate-50 text-slate-500 text-left text-[11px] uppercase tracking-wide">
                  <tr>
                    <th className="px-3 py-2.5 w-8" />
                    <th className="px-3 py-2.5 font-bold">Number</th>
                    <th className="px-3 py-2.5 font-bold">Name</th>
                    <th className="px-3 py-2.5 font-bold">Found in</th>
                    <th className="px-3 py-2.5 font-bold">Primary MIS entry</th>
                    <th className="px-3 py-2.5 font-bold">Lender</th>
                    <th className="px-3 py-2.5 font-bold">MIS status</th>
                    <th className="px-3 py-2.5 font-bold text-right">Loan amount</th>
                  </tr>
                </thead>
                <tbody>
                  {visible.length === 0 && (
                    <tr><td colSpan={8} className="px-4 py-10 text-center text-slate-500 text-sm">No number matches these filters.</td></tr>
                  )}
                  {pageRows.map((r, i) => {
                    const key = `${r.normPhone || r.inputPhone}-${i}-${page}`;
                    const isOpen = open.has(key);
                    const p = r.mis?.primary;
                    const tier = tierOf(r);
                    return (
                      <React.Fragment key={key}>
                        <tr
                          onClick={() => r.status === 'in_mis' && toggle(key)}
                          className={`border-t border-slate-100 ${r.status === 'in_mis' ? 'cursor-pointer hover:bg-slate-50' : ''} ${r.status !== 'in_mis' ? 'bg-slate-50/50' : ''}`}
                        >
                          <td className="px-3 py-2.5 text-slate-400">
                            {r.status === 'in_mis' ? (isOpen ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />) : null}
                          </td>
                          <td className="px-3 py-2.5 font-mono font-bold text-slate-900">
                            {r.normPhone || r.inputPhone}
                            {r.inputCount > 1 && <span className="ml-2 text-[10px] font-bold text-amber-700 bg-amber-50 border border-amber-200 rounded-full px-1.5 py-0.5">×{r.inputCount} in list</span>}
                          </td>
                          <td className="px-3 py-2.5 text-slate-800 font-semibold">{show(p?.customerName)}</td>
                          <td className="px-3 py-2.5">
                            {r.status === 'in_mis' ? (
                              <div className="flex flex-wrap gap-1.5">
                                <span className="px-2 py-0.5 rounded-full border text-[11px] font-bold bg-emerald-50 text-emerald-700 border-emerald-200">
                                  MIS ×{r.mis!.entryCount}
                                </span>
                                {(r.directory?.entryCount || 0) > 0 && (
                                  <span className="px-2 py-0.5 rounded-full border text-[11px] font-bold bg-blue-50 text-blue-700 border-blue-200">
                                    Directory ×{r.directory!.entryCount}
                                  </span>
                                )}
                                {r.api?.matched && (
                                  <span className="px-2 py-0.5 rounded-full border text-[11px] font-bold bg-indigo-50 text-indigo-700 border-indigo-200">
                                    Live ticket
                                  </span>
                                )}
                              </div>
                            ) : (
                              <span className={`px-2 py-0.5 rounded-full border text-[11px] font-bold ${TIERS.find((t) => t.id === tier)!.chip}`}>
                                {r.status === 'invalid' ? 'Invalid number' : 'Not found in MIS'}
                              </span>
                            )}
                          </td>
                          <td className="px-3 py-2.5 text-xs text-slate-600">{p ? `${p.fileName} · ${p.sheetName} · row ${p.rowNum}` : '—'}</td>
                          <td className="px-3 py-2.5 text-slate-700 font-semibold">{show(p?.lender)}</td>
                          <td className="px-3 py-2.5">
                            {p?.status ? <span className={`px-2 py-0.5 rounded-full border text-[11px] font-bold ${statusTone(p.status)}`}>{p.status}</span> : '—'}
                          </td>
                          <td className="px-3 py-2.5 text-right tabular-nums text-slate-800">{show(p?.loanAmount)}</td>
                        </tr>

                        {isOpen && r.mis && (
                          <tr className="bg-slate-50/70">
                            <td colSpan={8} className="px-4 py-4">
                              <div className="grid xl:grid-cols-3 gap-4">
                                {/* MIS */}
                                <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-3">
                                  <div className="flex items-center gap-2 text-xs font-black text-slate-800"><Layers className="w-4 h-4 text-emerald-600" /> MIS · {r.mis.entryCount} {r.mis.entryCount === 1 ? 'entry' : 'entries'}</div>
                                  <div className="text-[11px] text-slate-500 font-medium">{r.mis.workbookCount} workbook{r.mis.workbookCount === 1 ? '' : 's'} · {r.mis.tabCount} tab{r.mis.tabCount === 1 ? '' : 's'}</div>
                                  <div className="space-y-1">
                                    {r.mis.books.map((b, bi) => (
                                      <div key={bi} className="flex items-center justify-between text-xs bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1.5">
                                        <span className="text-slate-700 font-semibold">{b.fileName} · {b.sheetName}</span>
                                        <span className="font-black tabular-nums text-slate-900">×{b.entries}</span>
                                      </div>
                                    ))}
                                  </div>
                                  <div className="flex flex-wrap gap-1.5">
                                    {Object.entries(r.mis.statuses).map(([s, n]) => (
                                      <span key={s} className={`px-2 py-0.5 rounded-full border text-[11px] font-bold ${statusTone(s)}`}>{s} ×{n}</span>
                                    ))}
                                    {Object.entries(r.mis.lenders).map(([l, n]) => (
                                      <span key={l} className="px-2 py-0.5 rounded-full border text-[11px] font-bold bg-slate-100 text-slate-700 border-slate-200">{l} ×{n}</span>
                                    ))}
                                  </div>
                                  <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
                                    {r.mis.entries.map((e, ei) => (
                                      <div key={e.id} className="border border-slate-200 rounded-lg p-2.5 text-xs space-y-1">
                                        <div className="flex items-center justify-between gap-2">
                                          <span className="font-bold text-slate-800">{ei === 0 ? 'Primary · ' : ''}{e.fileName} · {e.sheetName} · row {e.rowNum}</span>
                                          {onOpenRowModal && (
                                            <button onClick={(ev) => { ev.stopPropagation(); onOpenRowModal(e.id); }} className="inline-flex items-center gap-1 text-blue-600 font-bold cursor-pointer shrink-0">
                                              <Maximize2 className="w-3 h-3" /> Full row
                                            </button>
                                          )}
                                        </div>
                                        <div className="grid grid-cols-2 gap-x-3 gap-y-0.5 text-slate-600">
                                          <span>Lender: <b className="text-slate-800">{show(e.lender)}</b></span>
                                          <span>Status: <b className="text-slate-800">{show(e.status)}</b></span>
                                          <span>Amount: <b className="text-slate-800">{show(e.loanAmount)}</b></span>
                                          <span>Ticket: <b className="text-slate-800">{show(e.ticketId)}</b></span>
                                          <span>Login: <b className="text-slate-800">{show(e.loginDate)}</b></span>
                                          <span>Approval: <b className="text-slate-800">{show(e.approvalDate)}</b></span>
                                          <span>Disbursed: <b className="text-slate-800">{show(e.disbursedDate)}</b></span>
                                          <span>TL: <b className="text-slate-800">{show(e.teamLeader)}</b></span>
                                          <span>Manager: <b className="text-slate-800">{show(e.manager)}</b></span>
                                          <span>City: <b className="text-slate-800">{show(e.city)}</b></span>
                                          <span>CIBIL: <b className="text-slate-800">{show(e.cibil)}</b></span>
                                          <span>PAN: <b className="text-slate-800">{show(e.pan)}</b></span>
                                        </div>
                                        {(e.opsRemarks || e.creditRemarks) && (
                                          <div className="text-slate-500">{[e.opsRemarks && `Ops: ${e.opsRemarks}`, e.creditRemarks && `Credit: ${e.creditRemarks}`].filter(Boolean).join(' · ')}</div>
                                        )}
                                      </div>
                                    ))}
                                    {r.mis.entriesHidden > 0 && <div className="text-[11px] text-slate-500 font-medium">+ {r.mis.entriesHidden} more MIS entries not shown</div>}
                                  </div>
                                </div>

                                {/* Directory */}
                                <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-3">
                                  <div className="flex items-center gap-2 text-xs font-black text-slate-800"><FolderTree className="w-4 h-4 text-blue-600" /> Main directory · {r.directory?.entryCount || 0} {(r.directory?.entryCount || 0) === 1 ? 'entry' : 'entries'}</div>
                                  {(r.directory?.entryCount || 0) === 0 ? (
                                    <div className="text-xs text-slate-500 font-medium">This number is not in any main directory file.</div>
                                  ) : (
                                    <>
                                      <div className="text-[11px] text-slate-500 font-medium">{r.directory!.fileCount} file{r.directory!.fileCount === 1 ? '' : 's'}</div>
                                      <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
                                        {r.directory!.entries.map((d) => (
                                          <div key={d.id} className="border border-slate-200 rounded-lg p-2.5 text-xs space-y-1">
                                            <div className="flex items-center justify-between gap-2">
                                              <span className="font-bold text-slate-800">{d.fileName}</span>
                                              {d.priority && (
                                                <span
                                                  className="px-2 py-0.5 rounded-full border text-[11px] font-bold"
                                                  style={d.badge?.bg ? { background: d.badge.bg, color: d.badge.text, borderColor: d.badge.border } : undefined}
                                                >
                                                  {d.priority}
                                                </span>
                                              )}
                                            </div>
                                            <div className="text-slate-600">{d.folderPath || '—'} · {d.sheetName} · row {d.rowNum}</div>
                                            {onOpenRowModal && (
                                              <button onClick={(ev) => { ev.stopPropagation(); onOpenRowModal(d.id); }} className="inline-flex items-center gap-1 text-blue-600 font-bold cursor-pointer">
                                                <Maximize2 className="w-3 h-3" /> Full row
                                              </button>
                                            )}
                                          </div>
                                        ))}
                                        {r.directory!.entriesHidden > 0 && <div className="text-[11px] text-slate-500 font-medium">+ {r.directory!.entriesHidden} more directory entries not shown</div>}
                                      </div>
                                    </>
                                  )}
                                </div>

                                {/* Live ticket */}
                                <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-3">
                                  <div className="flex items-center gap-2 text-xs font-black text-slate-800"><Ticket className="w-4 h-4 text-indigo-600" /> Live ticket API</div>
                                  {!r.api?.ticket ? (
                                    <div className="text-xs text-slate-500 font-medium">No live ticket has this phone number.</div>
                                  ) : (
                                    <div className="space-y-2 text-xs">
                                      <div className="flex items-center justify-between gap-2">
                                        <span className="font-black text-slate-900">{r.api.ticket.formattedTicketId || r.api.ticket.ticketId}</span>
                                        <span className="px-2 py-0.5 rounded-full border text-[11px] font-bold bg-indigo-50 text-indigo-700 border-indigo-200">Score {r.api.ticket.matchScore}</span>
                                      </div>
                                      <div className="grid grid-cols-2 gap-x-3 gap-y-0.5 text-slate-600">
                                        <span>Name: <b className="text-slate-800">{show(r.api.ticket.customerName)}</b></span>
                                        <span>Status: <b className="text-slate-800">{show(r.api.ticket.ticketStatus || r.api.ticket.loanStatus)}</b></span>
                                        <span>Provider: <b className="text-slate-800">{show(r.api.ticket.applicationProvider || r.api.ticket.standardizedProvider)}</b></span>
                                        <span>Amount: <b className="text-slate-800">{show(r.api.ticket.applicationAmount)}</b></span>
                                      </div>
                                      {r.api.ticket.matchedParameters && (
                                        <div className="flex flex-wrap gap-1.5">
                                          {([
                                            ['Phone', r.api.ticket.matchedParameters.phoneMatched],
                                            ['Lender', r.api.ticket.matchedParameters.lenderMatched],
                                            ['Name', r.api.ticket.matchedParameters.nameMatched],
                                            ['Ticket ID', r.api.ticket.matchedParameters.ticketIdMatched],
                                            ['Amount', r.api.ticket.matchedParameters.amountMatched],
                                          ] as [string, boolean][]).map(([label, ok]) => (
                                            <span key={label} className={`px-2 py-0.5 rounded-full border text-[11px] font-bold ${ok ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-white text-slate-400 border-slate-200'}`}>
                                              {label} {ok ? '✓' : '✗'}
                                            </span>
                                          ))}
                                        </div>
                                      )}
                                      {onNavigateToTicket && (
                                        <button onClick={(ev) => { ev.stopPropagation(); onNavigateToTicket(r.api!.ticket!.ticketId); }} className="inline-flex items-center gap-1 text-blue-600 font-bold cursor-pointer">
                                          Open ticket
                                        </button>
                                      )}
                                    </div>
                                  )}
                                  {onLookupSingle && (
                                    <button onClick={(ev) => { ev.stopPropagation(); onLookupSingle(r.normPhone); }} className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 text-xs font-bold text-slate-700 hover:bg-slate-50 cursor-pointer">
                                      <UserSearch className="w-3.5 h-3.5" /> Open in single lookup
                                    </button>
                                  )}
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
            {pageCount > 1 && (
              <div className="flex items-center justify-between px-4 py-3 border-t border-slate-100 text-xs text-slate-600 font-medium">
                <span>Page {page} of {pageCount} · {visible.length.toLocaleString()} numbers</span>
                <div className="flex gap-2">
                  <button onClick={() => setPage((p) => Math.max(1, p - 1))} disabled={page <= 1} className="px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 disabled:opacity-40 font-bold cursor-pointer">Previous</button>
                  <button onClick={() => setPage((p) => Math.min(pageCount, p + 1))} disabled={page >= pageCount} className="px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 disabled:opacity-40 font-bold cursor-pointer">Next</button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
