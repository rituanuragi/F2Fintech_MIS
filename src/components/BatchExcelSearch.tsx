import React, { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload, ListChecks, Loader2, Download, CheckCircle2, XCircle, ChevronDown, ChevronRight,
  Maximize2, UserSearch, AlertTriangle, FileSpreadsheet,
} from 'lucide-react';
import { downloadFromPost } from '../utils/download';
import { BatchPivotPanel } from './BatchPivotPanel';
import { defaultPivotHeaders, PivotHeaders } from '../utils/pivot';

interface BatchExcelSearchProps {
  onViewRecord: (id: number) => void;
  onPreviewSheet?: (fileId: string, sheetName: string) => void;
  onLookupCustomer?: (phone: string) => void;
}

interface BatchRecord {
  id: number;
  customerName: string;
  phone: string;
  fileName: string;
  sheetName: string;
  sheetLabel: string;
  login: string;
  lot: string;
  neet: string;
  rowNum: number;
  folderPath: string;
}

interface BatchResult {
  inputPhone: string;
  normPhone: string;
  found: boolean;
  priority: string;
  priorityCategory: string;
  priorityRank: number;
  badgeText: string;
  sheetsName: string;
  login: string;
  lot: string;
  neet: string;
  workbookNames: string;
  customerName: string;
  rowNumbers: string;
  matchCount: number;
  records: BatchRecord[];
}

interface BatchResponse {
  results: BatchResult[];
  summary: { totalUploaded: number; totalFound: number; totalNotFound: number; durationMs: number; scope: string };
}

const PHONE_HEADER = /phone|mobile|contact|number|mob\b|calling/i;
const tenDigits = (v: any) => String(v ?? '').replace(/\D/g, '').slice(-10).length === 10;

/** Pick the column that holds phone numbers: by header name first, then by how many cells look like a phone. */
function detectPhoneColumn(headers: string[], rows: Record<string, any>[]): string | null {
  const sample = rows.slice(0, 300);
  const byHeader = headers.filter((h) => PHONE_HEADER.test(h));
  const score = (h: string) => sample.filter((r) => tenDigits(r[h])).length;
  const pool = byHeader.length ? byHeader : headers;
  let best: string | null = null;
  let bestScore = 0;
  for (const h of pool) {
    const s = score(h);
    if (s > bestScore) { best = h; bestScore = s; }
  }
  return best;
}

export const BatchExcelSearch: React.FC<BatchExcelSearchProps> = ({ onViewRecord, onLookupCustomer }) => {
  const fileRef = useRef<HTMLInputElement>(null);
  const [fileName, setFileName] = useState('');
  const [headers, setHeaders] = useState<string[]>([]);
  const [rows, setRows] = useState<Record<string, any>[]>([]);
  const [phoneCol, setPhoneCol] = useState<string>('');
  const [pasted, setPasted] = useState('');
  const [scope, setScope] = useState<'neet_login_lot' | 'all'>('neet_login_lot');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [data, setData] = useState<BatchResponse | null>(null);
  const [view, setView] = useState<'all' | 'found' | 'missing'>('all');
  const [open, setOpen] = useState<Record<string, boolean>>({});
  const [exporting, setExporting] = useState<'' | 'xlsx' | 'csv'>('');
  // Pivot table: header texts the user can change before downloading
  const [pivotHeaders, setPivotHeaders] = useState<PivotHeaders>(() => defaultPivotHeaders());
  const [includePivot, setIncludePivot] = useState(true);
  const [useTitleAsTabName, setUseTitleAsTabName] = useState(false);

  const readFile = async (file: File) => {
    setError('');
    setData(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: false });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<Record<string, any>>(sheet, { defval: '', raw: false });
      if (json.length === 0) throw new Error('The first sheet of this file is empty.');
      const hs = Object.keys(json[0]);
      setFileName(file.name);
      setHeaders(hs);
      setRows(json);
      setPhoneCol(detectPhoneColumn(hs, json) || '');
      setPasted('');
    } catch (e: any) {
      setError(e.message || 'Could not read this file.');
    }
  };

  const phones = useMemo(() => {
    if (pasted.trim()) {
      return pasted.split(/[\s,;]+/).map((s) => s.trim()).filter(Boolean).map((p) => ({ rawPhone: p }));
    }
    if (!phoneCol) return [];
    return rows
      .map((r) => ({ rawPhone: String(r[phoneCol] ?? '').trim(), originalData: r }))
      .filter((p) => p.rawPhone);
  }, [pasted, phoneCol, rows]);

  const run = async () => {
    if (phones.length === 0) return;
    setLoading(true);
    setError('');
    setData(null);
    try {
      const res = await fetch('/api/batch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ phones, scope }),
      });
      const json = await res.json().catch(() => null);
      if (!res.ok || !json || !Array.isArray(json.results)) {
        throw new Error((json && json.error) || `Batch search failed (HTTP ${res.status})`);
      }
      setData(json);
      setView('all');
      setOpen({});
    } catch (e: any) {
      setError(e.message || 'Batch search failed');
    } finally {
      setLoading(false);
    }
  };

  const shown = useMemo(() => {
    if (!data) return [];
    return data.results.filter((r) => (view === 'all' ? true : view === 'found' ? r.found : !r.found));
  }, [data, view]);

  const exportAs = async (format: 'xlsx' | 'csv') => {
    if (!data) return;
    setExporting(format);
    try {
      // The server only needs the summary columns, so the row details are left out of the upload.
      const slim = data.results.map(({ records, ...rest }) => rest);
      await downloadFromPost(
        '/api/batch/export',
        {
          results: slim,
          format,
          includePivot,
          pivotTitle: pivotHeaders.title,
          dataHeader: pivotHeaders.dataHeader,
          countHeader: pivotHeaders.countHeader,
          blanksLabel: pivotHeaders.blanksLabel,
          totalLabel: pivotHeaders.totalLabel,
          useTitleAsTabName,
        },
        `F2_Batch_Lookup.${format}`
      );
    } catch (e: any) {
      window.alert(e.message || 'Export failed');
    } finally {
      setExporting('');
    }
  };

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-5">
        <div>
          <h2 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
            <ListChecks className="w-5 h-5 text-blue-600" />
            <span>Batch Excel Search</span>
          </h2>
          <p className="text-xs text-slate-500 mt-1 font-medium">
            Upload an Excel or CSV with phone numbers, or paste them. Each number is looked up in the Drive directory files.
          </p>
        </div>

        <div className="grid md:grid-cols-2 gap-4">
          <div
            onClick={() => fileRef.current?.click()}
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => {
              e.preventDefault();
              const f = e.dataTransfer.files?.[0];
              if (f) readFile(f);
            }}
            className="border-2 border-dashed border-slate-300 hover:border-blue-400 rounded-2xl p-6 text-center cursor-pointer bg-slate-50 hover:bg-blue-50/40 transition-colors"
          >
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) readFile(f);
                e.target.value = '';
              }}
            />
            {fileName ? (
              <>
                <FileSpreadsheet className="w-7 h-7 mx-auto text-emerald-600" />
                <div className="text-sm font-bold text-slate-900 mt-2 break-all">{fileName}</div>
                <div className="text-[11px] text-slate-500 font-mono mt-0.5">{rows.length.toLocaleString()} rows · click to replace</div>
              </>
            ) : (
              <>
                <Upload className="w-7 h-7 mx-auto text-slate-400" />
                <div className="text-sm font-bold text-slate-700 mt-2">Drop an Excel or CSV here</div>
                <div className="text-[11px] text-slate-400 mt-0.5">or click to choose a file (.xlsx, .xls, .csv)</div>
              </>
            )}
          </div>

          <div>
            <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Or paste numbers</label>
            <textarea
              value={pasted}
              onChange={(e) => setPasted(e.target.value)}
              rows={5}
              placeholder={'9876543210\n9123456780'}
              className="mt-1.5 w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm font-mono text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-500 focus:bg-white focus:ring-1 focus:ring-blue-500"
            />
          </div>
        </div>

        {rows.length > 0 && !pasted.trim() && (
          <div className="flex flex-wrap items-center gap-3 text-xs">
            <span className="font-bold text-slate-600">Phone column</span>
            <select
              value={phoneCol}
              onChange={(e) => setPhoneCol(e.target.value)}
              className="px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:border-blue-500"
            >
              <option value="">Choose a column…</option>
              {headers.map((h) => <option key={h} value={h}>{h}</option>)}
            </select>
            {!phoneCol && (
              <span className="inline-flex items-center gap-1 text-amber-700">
                <AlertTriangle className="w-3.5 h-3.5" /> No phone column found. Please pick one.
              </span>
            )}
          </div>
        )}

        <div className="flex flex-wrap items-center gap-3">
          <div className="flex gap-1.5">
            {([
              { id: 'neet_login_lot', label: 'NEET / Login / Lot' },
              { id: 'all', label: 'All directory files' },
            ] as const).map((s) => (
              <button
                key={s.id}
                onClick={() => setScope(s.id)}
                className={`px-3 py-2 rounded-xl text-xs font-bold cursor-pointer border ${
                  scope === s.id ? 'bg-slate-900 text-white border-slate-900' : 'bg-white text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                {s.label}
              </button>
            ))}
          </div>
          <button
            onClick={run}
            disabled={loading || phones.length === 0}
            className="ml-auto px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white rounded-xl text-sm font-bold shadow-sm flex items-center gap-2 cursor-pointer"
          >
            {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : <ListChecks className="w-4 h-4" />}
            Search {phones.length > 0 ? `${phones.length.toLocaleString()} numbers` : ''}
          </button>
        </div>

        {error && (
          <div className="bg-rose-50 border border-rose-200 text-rose-700 rounded-xl p-3 text-xs flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 shrink-0" /> {error}
          </div>
        )}
      </div>

      {data && (
        <div className="space-y-4">
          <div className="bg-white border border-slate-200/90 rounded-2xl p-4 shadow-xs flex flex-wrap items-center gap-4">
            <div className="flex gap-1.5">
              {([
                { id: 'all', label: `All ${data.summary.totalUploaded.toLocaleString()}` },
                { id: 'found', label: `Found ${data.summary.totalFound.toLocaleString()}` },
                { id: 'missing', label: `Not found ${data.summary.totalNotFound.toLocaleString()}` },
              ] as const).map((t) => (
                <button
                  key={t.id}
                  onClick={() => setView(t.id)}
                  className={`px-3.5 py-2 rounded-xl text-xs font-bold cursor-pointer border ${
                    view === t.id ? 'bg-blue-600 text-white border-blue-600' : 'bg-slate-100 text-slate-700 border-slate-200 hover:bg-slate-200'
                  }`}
                >
                  {t.label}
                </button>
              ))}
            </div>
            <span className="text-[11px] text-slate-400 font-mono">{(data.summary.durationMs / 1000).toFixed(1)}s</span>
            <div className="ml-auto flex gap-2">
              <button
                onClick={() => exportAs('csv')}
                disabled={!!exporting}
                className="inline-flex items-center gap-1.5 px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold border border-slate-200 cursor-pointer disabled:opacity-50"
              >
                {exporting === 'csv' ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Download className="w-3.5 h-3.5" />} CSV
              </button>
              <button
                onClick={() => exportAs('xlsx')}
                disabled={!!exporting}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold cursor-pointer disabled:opacity-50"
              >
                {exporting === 'xlsx' ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Download className="w-3.5 h-3.5" />} Excel
              </button>
            </div>
          </div>

          <BatchPivotPanel
            results={data.results}
            headers={pivotHeaders}
            onHeadersChange={setPivotHeaders}
            includePivot={includePivot}
            onIncludePivotChange={setIncludePivot}
            useTitleAsTabName={useTitleAsTabName}
            onUseTitleAsTabNameChange={setUseTitleAsTabName}
          />

          <div className="bg-white border border-slate-200/90 rounded-2xl shadow-xs overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-slate-500 uppercase tracking-wider text-[10px] font-bold">
                  <tr>
                    <th className="px-3 py-3 w-8"></th>
                    <th className="px-3 py-3">Phone</th>
                    <th className="px-3 py-3">Status</th>
                    <th className="px-3 py-3">Priority</th>
                    <th className="px-3 py-3">Sheet</th>
                    <th className="px-3 py-3">Login</th>
                    <th className="px-3 py-3">Lot</th>
                    <th className="px-3 py-3">NEET</th>
                    <th className="px-3 py-3">Name</th>
                    <th className="px-3 py-3 text-right">Matches</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {shown.slice(0, 1000).map((r, i) => {
                    const key = `${r.normPhone}-${i}`;
                    const isOpen = !!open[key];
                    return (
                      <React.Fragment key={key}>
                        <tr
                          className={`${r.found ? 'cursor-pointer hover:bg-slate-50' : ''}`}
                          onClick={() => r.found && setOpen((o) => ({ ...o, [key]: !o[key] }))}
                        >
                          <td className="px-3 py-2.5 text-slate-400">
                            {r.found ? (isOpen ? <ChevronDown className="w-4 h-4" /> : <ChevronRight className="w-4 h-4" />) : null}
                          </td>
                          <td className="px-3 py-2.5 font-mono font-bold text-slate-900">{r.inputPhone}</td>
                          <td className="px-3 py-2.5">
                            {r.found ? (
                              <span className="inline-flex items-center gap-1 text-emerald-700 font-bold"><CheckCircle2 className="w-3.5 h-3.5" /> Found</span>
                            ) : (
                              <span className="inline-flex items-center gap-1 text-rose-600 font-bold"><XCircle className="w-3.5 h-3.5" /> Not found</span>
                            )}
                          </td>
                          <td className="px-3 py-2.5">
                            {r.found ? (
                              <span className={`inline-block px-2 py-0.5 rounded-md bg-slate-900 font-bold ${r.badgeText || 'text-white'}`}>{r.priority}</span>
                            ) : '—'}
                          </td>
                          <td className="px-3 py-2.5 text-slate-700 max-w-[220px] truncate" title={r.sheetsName}>{r.sheetsName}</td>
                          <td className="px-3 py-2.5 text-slate-700">{r.login}</td>
                          <td className="px-3 py-2.5 text-slate-700">{r.lot}</td>
                          <td className="px-3 py-2.5 text-slate-700">{r.neet}</td>
                          <td className="px-3 py-2.5 text-slate-700 max-w-[160px] truncate">{r.customerName}</td>
                          <td className="px-3 py-2.5 text-right font-mono">{r.matchCount}</td>
                        </tr>
                        {isOpen && (
                          <tr>
                            <td colSpan={10} className="bg-slate-50/70 px-6 py-3">
                              <table className="w-full text-[11px]">
                                <thead className="text-slate-400 uppercase tracking-wider text-[10px] font-bold">
                                  <tr>
                                    <th className="text-left py-1">Workbook</th>
                                    <th className="text-left py-1">Sheet</th>
                                    <th className="text-right py-1">Excel row</th>
                                    <th className="text-right py-1">Actions</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-200/70">
                                  {r.records.map((rec) => (
                                    <tr key={rec.id}>
                                      <td className="py-1.5 text-slate-800 font-semibold">{rec.fileName}</td>
                                      <td className="py-1.5 text-slate-600">{rec.sheetLabel || rec.sheetName}</td>
                                      <td className="py-1.5 text-right font-mono">{rec.rowNum}</td>
                                      <td className="py-1.5 text-right whitespace-nowrap">
                                        <button
                                          onClick={(e) => { e.stopPropagation(); onViewRecord(rec.id); }}
                                          className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 font-bold cursor-pointer mr-1.5"
                                        >
                                          <Maximize2 className="w-3 h-3" /> Row
                                        </button>
                                        {onLookupCustomer && (
                                          <button
                                            onClick={(e) => { e.stopPropagation(); onLookupCustomer(r.normPhone || r.inputPhone); }}
                                            className="inline-flex items-center gap-1 px-2 py-1 rounded-lg bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 font-bold cursor-pointer"
                                          >
                                            <UserSearch className="w-3 h-3" /> Lookup
                                          </button>
                                        )}
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {shown.length > 1000 && (
              <div className="p-3 text-center text-[11px] text-slate-500 border-t border-slate-100">
                Showing the first 1,000 of {shown.length.toLocaleString()} rows. The Excel and CSV exports include all of them.
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
