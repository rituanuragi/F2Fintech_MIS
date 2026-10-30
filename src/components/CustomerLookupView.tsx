import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  UserSearch, Search, Loader2, AlertTriangle, History, MessageSquare, CheckCheck, Ticket, FileSpreadsheet,
  Maximize2, GitCompare, Building2, Phone, CreditCard, Calendar, ArrowRight,
} from 'lucide-react';
import { AuthUser, CustomerLookupRecord, CustomerSearchResult } from '../types';
import { LeadAuditModal } from './LeadAuditModal';
import { RemarksModal } from './RemarksModal';
import { BatchCustomerLookup } from './BatchCustomerLookup';

interface CustomerLookupViewProps {
  initialQuery?: string;
  currentUser?: AuthUser | null;
  onOpenRowModal?: (id: number) => void;
  onNavigateToMIS?: () => void;
  onNavigateToTicket?: (ticketId: string | number) => void;
  onPreviewSheet?: (fileId: string, sheetName: string) => void;
}

type Field = 'all' | 'phone' | 'pan' | 'aadhar';

const FIELDS: { id: Field; label: string; placeholder: string }[] = [
  { id: 'all', label: 'Any', placeholder: 'Phone, PAN, name…' },
  { id: 'phone', label: 'Phone', placeholder: '10-digit mobile number' },
  { id: 'pan', label: 'PAN', placeholder: 'ABCDE1234F' },
  { id: 'aadhar', label: 'Aadhaar', placeholder: '12-digit Aadhaar number' },
];

// Show only the last 4 characters of an ID number.
const mask = (v?: string) => {
  const s = String(v || '').trim();
  if (!s) return '—';
  return s.length <= 4 ? s : `${'X'.repeat(s.length - 4)}${s.slice(-4)}`;
};

const show = (v: any) => {
  const s = v === null || v === undefined ? '' : String(v).trim();
  return s || '—';
};

// Fields compared side by side when a customer has more than one record.
const DIFF_FIELDS: { label: string; get: (r: CustomerLookupRecord) => string }[] = [
  { label: 'Name', get: (r) => r.customerName || '' },
  { label: 'Phone', get: (r) => r.phoneNumber || '' },
  { label: 'PAN', get: (r) => mask(r.panNumber) },
  { label: 'Aadhaar', get: (r) => mask(r.aadharNumber) },
  { label: 'Lender', get: (r) => r.lenderName || '' },
  { label: 'Loan amount', get: (r) => String(r.loanAmount ?? '') },
  { label: 'Status', get: (r) => r.status || '' },
  { label: 'City', get: (r) => r.city || '' },
  { label: 'Login date', get: (r) => r.loginDate || '' },
  { label: 'Approval date', get: (r) => r.approvalDate || '' },
  { label: 'Disbursed date', get: (r) => r.disbursedDate || '' },
];

const statusTone = (s?: string) => {
  const t = (s || '').toUpperCase();
  if (/DISBURS/.test(t)) return 'bg-emerald-50 text-emerald-700 border-emerald-200';
  if (/APPROV/.test(t)) return 'bg-blue-50 text-blue-700 border-blue-200';
  if (/REJECT|DECLIN|CANCEL/.test(t)) return 'bg-rose-50 text-rose-700 border-rose-200';
  return 'bg-slate-100 text-slate-600 border-slate-200';
};

export const CustomerLookupView: React.FC<CustomerLookupViewProps> = ({
  initialQuery = '',
  currentUser,
  onOpenRowModal,
  onNavigateToMIS,
  onNavigateToTicket,
}) => {
  const [mode, setMode] = useState<'single' | 'batch'>('single');
  const [field, setField] = useState<Field>('all');
  const [query, setQuery] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [result, setResult] = useState<CustomerSearchResult | null>(null);
  const [adminName, setAdminName] = useState('Admin');
  const [auditRecord, setAuditRecord] = useState<CustomerLookupRecord | null>(null);
  const [remarksRecord, setRemarksRecord] = useState<CustomerLookupRecord | null>(null);
  const [marking, setMarking] = useState<number | null>(null);
  const lastRun = useRef('');

  const runSearch = useCallback(async (q: string, f: Field) => {
    const text = q.trim();
    if (!text) return;
    lastRun.current = `${f}|${text}`;
    setLoading(true);
    setError('');
    try {
      const res = await fetch(`/api/records/search?field=${f}&q=${encodeURIComponent(text)}`);
      const json = await res.json().catch(() => null);
      if (!res.ok || !json || !Array.isArray(json.records)) {
        throw new Error((json && json.error) || `Search failed (HTTP ${res.status})`);
      }
      setResult(json);
    } catch (e: any) {
      setResult(null);
      setError(e.message || 'Search failed');
    } finally {
      setLoading(false);
    }
  }, []);

  // A phone sent from another screen runs straight away.
  useEffect(() => {
    const q = (initialQuery || '').trim();
    if (!q) return;
    const f: Field = /^\d{10,}$/.test(q.replace(/\D/g, '')) && /^[\d\s+()-]+$/.test(q) ? 'phone' : 'all';
    setQuery(q);
    setField(f);
    if (lastRun.current !== `${f}|${q}`) runSearch(q, f);
  }, [initialQuery, runSearch]);

  const replaceRecord = (id: number, patch: Partial<CustomerLookupRecord>) => {
    setResult((prev) =>
      prev ? { ...prev, records: prev.records.map((r) => (r.id === id ? { ...r, ...patch } : r)) } : prev
    );
    setAuditRecord((r) => (r && r.id === id ? { ...r, ...patch } : r));
    setRemarksRecord((r) => (r && r.id === id ? { ...r, ...patch } : r));
  };

  // Opening the trail only reads it. It never adds an entry.
  const openAudit = async (rec: CustomerLookupRecord) => {
    setAuditRecord(rec);
    try {
      const res = await fetch(`/api/records/${rec.id}/checks`);
      if (!res.ok) return;
      const d = await res.json();
      replaceRecord(rec.id, { checkHistory: d.checkHistory || [], checkCount: d.checkCount || 0 });
    } catch { /* keep what we already have */ }
  };

  // Marking a lead as checked is a deliberate action and adds one entry.
  const markChecked = async (rec: CustomerLookupRecord) => {
    setMarking(rec.id);
    try {
      const res = await fetch(`/api/records/${rec.id}/check`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ adminName }),
      });
      const d = await res.json().catch(() => null);
      if (!res.ok || !d) throw new Error((d && d.error) || `HTTP ${res.status}`);
      replaceRecord(rec.id, { checkHistory: d.checkHistory, checkCount: d.checkCount });
    } catch (e: any) {
      window.alert(`Could not record the check: ${e.message}`);
    } finally {
      setMarking(null);
    }
  };

  const addRemark = async (recordId: number, text: string) => {
    const res = await fetch(`/api/records/${recordId}/remarks`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text, author: currentUser?.username || 'admin', authorName: adminName }),
    });
    const d = await res.json().catch(() => null);
    if (!res.ok || !d?.remark) throw new Error((d && d.error) || `HTTP ${res.status}`);
    const rec = result?.records.find((r) => r.id === recordId);
    replaceRecord(recordId, { remarks: [d.remark, ...(rec?.remarks || [])] });
  };

  const deleteRemark = async (recordId: number, remarkId: string) => {
    const res = await fetch(`/api/records/${recordId}/remarks/${remarkId}`, { method: 'DELETE' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const rec = result?.records.find((r) => r.id === recordId);
    replaceRecord(recordId, { remarks: (rec?.remarks || []).filter((r) => r.id !== remarkId) });
  };

  const records = result?.records || [];
  const diffRows = useMemo(() => {
    if (records.length < 2) return [];
    return DIFF_FIELDS.map((f) => {
      const values = records.map((r) => f.get(r));
      const distinct = new Set(values.filter((v) => v && v !== '—'));
      return { label: f.label, values, differs: distinct.size > 1 };
    });
  }, [records]);

  const placeholder = FIELDS.find((f) => f.id === field)?.placeholder;

  return (
    <div className="space-y-6">
      <div className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-4">
        <div>
          <h2 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
            <UserSearch className="w-5 h-5 text-blue-600" />
            <span>Customer Lookup</span>
          </h2>
          <p className="text-xs text-slate-500 mt-1 font-medium">
            Find one customer across MIS sheets, Drive directory files and live tickets. PAN and Aadhaar show only their last 4 characters.
          </p>
        </div>

        <div className="inline-flex rounded-xl border border-slate-200 overflow-hidden bg-slate-50">
          {([['single', 'Single lookup'], ['batch', 'Batch lookup']] as const).map(([id, label]) => (
            <button
              key={id}
              onClick={() => setMode(id)}
              className={`px-4 py-2 text-xs font-bold cursor-pointer ${mode === id ? 'bg-slate-900 text-white' : 'text-slate-700 hover:bg-slate-100'}`}
            >
              {label}
            </button>
          ))}
        </div>

        {mode === 'single' && (<>
        <div className="flex flex-wrap gap-2">
          {FIELDS.map((f) => (
            <button
              key={f.id}
              onClick={() => setField(f.id)}
              className={`px-4 py-2 rounded-xl text-xs font-bold cursor-pointer border ${
                field === f.id ? 'bg-blue-600 text-white border-blue-600' : 'bg-slate-100 text-slate-700 border-slate-200 hover:bg-slate-200'
              }`}
            >
              {f.label}
            </button>
          ))}
        </div>

        <form onSubmit={(e) => { e.preventDefault(); runSearch(query, field); }} className="flex gap-2.5">
          <div className="relative flex-1">
            <Search className="w-5 h-5 text-slate-400 absolute left-4 top-1/2 -translate-y-1/2" />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder={placeholder}
              className="w-full pl-12 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-xl text-sm text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-500 focus:bg-white focus:ring-1 focus:ring-blue-500 font-medium"
            />
          </div>
          <button
            type="submit"
            disabled={loading || !query.trim()}
            className="px-6 py-3 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white rounded-xl text-sm font-bold shadow-sm flex items-center gap-2 cursor-pointer shrink-0"
          >
            {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : <Search className="w-4 h-4" />} Look up
          </button>
        </form>
        </>)}
      </div>

      {mode === 'batch' && (
        <BatchCustomerLookup
          onOpenRowModal={onOpenRowModal}
          onNavigateToTicket={onNavigateToTicket}
          onLookupSingle={(phone) => {
            setMode('single');
            setField('phone');
            setQuery(phone);
            runSearch(phone, 'phone');
          }}
        />
      )}

      {mode === 'single' && (<>

      {error && (
        <div className="bg-rose-50 border border-rose-200 text-rose-700 rounded-2xl p-4 text-sm flex items-center gap-2">
          <AlertTriangle className="w-4 h-4 shrink-0" /> {error}
        </div>
      )}

      {result && records.length === 0 && !loading && (
        <div className="bg-white border border-dashed border-slate-300 rounded-2xl p-10 text-center text-slate-500 text-sm space-y-2">
          <div>
            {result.notInMis ? (
              <>
                <strong className="text-slate-700">Not found in MIS</strong> for “{result.searchQuery}”.
              </>
            ) : (
              <>No customer found for “{result.searchQuery}”.</>
            )}
          </div>
          {onNavigateToMIS && (
            <button onClick={onNavigateToMIS} className="inline-flex items-center gap-1 text-blue-600 font-bold text-xs cursor-pointer">
              Check that the MIS sheet is connected <ArrowRight className="w-3 h-3" />
            </button>
          )}
        </div>
      )}

      {records.length > 0 && (
        <div className="space-y-4">
          <div className="text-[11px] text-slate-500 font-mono font-semibold px-1">
            {records.length} record{records.length === 1 ? '' : 's'}
            {result?.variantDetails?.distinctPanCount && result.variantDetails.distinctPanCount > 1
              ? ` · ${result.variantDetails.distinctPanCount} different PANs`
              : ''}
          </div>

          {result?.variantDetails?.distinctPanCount && result.variantDetails.distinctPanCount > 1 && (
            <div className="bg-amber-50 border border-amber-200 text-amber-800 rounded-xl p-3 text-xs font-semibold flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 shrink-0" />
              These records carry more than one PAN. Check that they belong to the same person.
            </div>
          )}

          {records.map((r) => {
            const live = r.liveTicketMatch;
            return (
              <div key={r.id} className="bg-white border border-slate-200/90 rounded-2xl shadow-xs overflow-hidden">
                <div className="px-5 py-4 flex flex-wrap items-start gap-3 border-b border-slate-100">
                  <div className="flex-1 min-w-[220px]">
                    <div className="flex flex-wrap items-center gap-2">
                      <h3 className="text-base font-black text-slate-900">{r.customerName || r.doctorName || 'Applicant'}</h3>
                      <span className={`px-2 py-0.5 rounded-full border text-[11px] font-bold ${statusTone(r.status)}`}>{show(r.status)}</span>
                      {r.isHighestPriority && (
                        <span className="px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200 text-[11px] font-bold">
                          Top priority
                        </span>
                      )}
                      {r.isDriveRecord && (
                        <span className="px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 border border-slate-200 text-[11px] font-bold">
                          Drive file
                        </span>
                      )}
                    </div>
                    <div className="text-[11px] text-slate-500 font-mono mt-1 flex items-center gap-1.5 flex-wrap">
                      <FileSpreadsheet className="w-3 h-3" />
                      {show(r.fileName)} · {show(r.sheetName)}{r.rowNum ? ` · row ${r.rowNum}` : ''}
                    </div>
                    {r.chosenPriorityReason && <div className="text-[11px] text-slate-400 mt-1">{r.chosenPriorityReason}</div>}
                  </div>

                  <div className="flex flex-wrap gap-2">
                    <button
                      onClick={() => markChecked(r)}
                      disabled={marking === r.id}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold cursor-pointer disabled:opacity-50"
                    >
                      {marking === r.id ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <CheckCheck className="w-3.5 h-3.5" />}
                      Mark checked
                    </button>
                    <button
                      onClick={() => openAudit(r)}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold border border-slate-200 cursor-pointer"
                    >
                      <History className="w-3.5 h-3.5" /> Audit {typeof r.checkCount === 'number' ? `(${r.checkCount})` : ''}
                    </button>
                    <button
                      onClick={() => setRemarksRecord(r)}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-semibold border border-slate-200 cursor-pointer"
                    >
                      <MessageSquare className="w-3.5 h-3.5" /> Remarks ({r.remarks?.length || 0})
                    </button>
                    {onOpenRowModal && (
                      <button
                        onClick={() => onOpenRowModal(r.id)}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 rounded-xl text-xs font-bold border border-blue-200 cursor-pointer"
                      >
                        <Maximize2 className="w-3.5 h-3.5" /> Full row
                      </button>
                    )}
                  </div>
                </div>

                <div className="p-5 grid sm:grid-cols-2 lg:grid-cols-4 gap-x-6 gap-y-3 text-xs">
                  {[
                    { icon: Phone, label: 'Phone', value: r.phoneNumber },
                    { icon: CreditCard, label: 'PAN', value: mask(r.panNumber) },
                    { icon: CreditCard, label: 'Aadhaar', value: mask(r.aadharNumber) },
                    { icon: Building2, label: 'Lender', value: r.lenderName },
                    { icon: Building2, label: 'Loan amount', value: r.loanAmount },
                    { icon: Building2, label: 'City / State', value: [r.city, r.state].filter(Boolean).join(', ') },
                    { icon: Building2, label: 'Manager', value: r.manager },
                    { icon: Building2, label: 'Team leader', value: r.teamLeader },
                    { icon: Calendar, label: 'Login date', value: r.loginDate },
                    { icon: Calendar, label: 'Approval date', value: r.approvalDate },
                    { icon: Calendar, label: 'Disbursed date', value: r.disbursedDate },
                    { icon: Building2, label: 'CIBIL', value: r.cibilScore },
                  ].map((f) => (
                    <div key={f.label}>
                      <div className="text-[10px] uppercase tracking-wider font-bold text-slate-400 flex items-center gap-1">
                        <f.icon className="w-3 h-3" /> {f.label}
                      </div>
                      <div className="font-semibold text-slate-900 mt-0.5 break-words">{show(f.value)}</div>
                    </div>
                  ))}
                </div>

                {(r.opsRemarks || r.creditRemarks) && (
                  <div className="px-5 pb-4 grid sm:grid-cols-2 gap-3 text-xs">
                    {r.opsRemarks && (
                      <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
                        <div className="text-[10px] uppercase tracking-wider font-bold text-slate-400 mb-1">Ops remarks</div>
                        {r.opsRemarks}
                      </div>
                    )}
                    {r.creditRemarks && (
                      <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
                        <div className="text-[10px] uppercase tracking-wider font-bold text-slate-400 mb-1">Credit remarks</div>
                        {r.creditRemarks}
                      </div>
                    )}
                  </div>
                )}

                {live && live.matched && (
                  <div className="mx-5 mb-4 bg-indigo-50/60 border border-indigo-200 rounded-xl p-3.5 text-xs">
                    <div className="flex flex-wrap items-center gap-2">
                      <Ticket className="w-4 h-4 text-indigo-600" />
                      <span className="font-black text-indigo-900">Live ticket {live.formattedTicketId}</span>
                      <span className="px-2 py-0.5 rounded-full bg-white border border-indigo-200 text-indigo-700 font-bold">
                        {live.matchScore}% match
                      </span>
                      {live.ticketStatus && <span className="text-indigo-700 font-semibold">{live.ticketStatus}</span>}
                      {onNavigateToTicket && (
                        <button
                          onClick={() => onNavigateToTicket(live.ticketId)}
                          className="ml-auto inline-flex items-center gap-1 text-indigo-700 font-bold cursor-pointer"
                        >
                          Open ticket <ArrowRight className="w-3 h-3" />
                        </button>
                      )}
                    </div>
                    <div className="mt-2 flex flex-wrap gap-1.5">
                      {([
                        ['Phone', live.matchedParameters?.phoneMatched],
                        ['Lender', live.matchedParameters?.lenderMatched],
                        ['Name', live.matchedParameters?.nameMatched],
                        ['Ticket ID', live.matchedParameters?.ticketIdMatched],
                        ['Amount', live.matchedParameters?.amountMatched],
                      ] as [string, boolean | undefined][]).map(([l, ok]) => (
                        <span
                          key={l}
                          className={`px-2 py-0.5 rounded-md border text-[11px] font-bold ${
                            ok ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-white text-slate-400 border-slate-200'
                          }`}
                        >
                          {ok ? '✓' : '–'} {l}
                        </span>
                      ))}
                    </div>
                    <div className="mt-2 text-indigo-900/80">
                      {[live.customerName, live.standardizedProvider || live.applicationProvider, live.applicationAmount]
                        .filter(Boolean).join(' · ')}
                    </div>
                  </div>
                )}

                {r.driveOccurrences && r.driveOccurrences.length > 0 && (
                  <div className="mx-5 mb-4 text-xs">
                    <div className="text-[10px] uppercase tracking-wider font-bold text-slate-400 mb-1.5">
                      Also in the Drive directory ({r.driveOccurrences.length})
                    </div>
                    <div className="flex flex-wrap gap-1.5">
                      {r.driveOccurrences.slice(0, 12).map((d) => (
                        <button
                          key={d.id}
                          onClick={() => onOpenRowModal?.(d.id)}
                          className="px-2 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 font-semibold cursor-pointer"
                        >
                          {d.fileName} · {d.sheetName}
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                {r.otherMisMatches && r.otherMisMatches.length > 0 && (
                  <div className="mx-5 mb-4 text-xs">
                    <div className="text-[10px] uppercase tracking-wider font-bold text-slate-400 mb-1.5">
                      Other MIS entries for this phone ({r.otherMisMatches.length})
                    </div>
                    <div className="space-y-1">
                      {r.otherMisMatches.slice(0, 8).map((m) => (
                        <div key={m.id} className="flex flex-wrap gap-x-3 text-slate-600">
                          <span className="font-semibold text-slate-800">{m.fileName} · {m.sheetName}</span>
                          {m.rowNum ? <span className="font-mono">row {m.rowNum}</span> : null}
                          {m.lenderName && <span>{m.lenderName}</span>}
                          {m.status && <span>{m.status}</span>}
                          {m.loginDate && <span>{m.loginDate}</span>}
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            );
          })}

          {diffRows.length > 0 && (
            <div className="bg-white border border-slate-200/90 rounded-2xl shadow-xs overflow-hidden">
              <div className="px-5 py-3 border-b border-slate-100 flex items-center gap-2">
                <GitCompare className="w-4 h-4 text-blue-600" />
                <h3 className="text-sm font-black text-slate-900">Compare records</h3>
                <span className="text-[11px] text-slate-400">Rows that differ are highlighted</span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-50 text-slate-500 uppercase tracking-wider text-[10px] font-bold">
                    <tr>
                      <th className="px-4 py-2.5">Field</th>
                      {records.map((r, i) => (
                        <th key={r.id} className="px-4 py-2.5 normal-case">
                          #{i + 1} · {r.sheetName || r.fileName}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {diffRows.map((d) => (
                      <tr key={d.label} className={d.differs ? 'bg-amber-50/60' : ''}>
                        <td className="px-4 py-2 font-bold text-slate-700">{d.label}</td>
                        {d.values.map((v, i) => (
                          <td key={i} className={`px-4 py-2 ${d.differs ? 'font-bold text-amber-900' : 'text-slate-700'}`}>{show(v)}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}

      </>)}

      <LeadAuditModal record={auditRecord} isOpen={auditRecord !== null} onClose={() => setAuditRecord(null)} />
      <RemarksModal
        record={remarksRecord}
        isOpen={remarksRecord !== null}
        onClose={() => setRemarksRecord(null)}
        onAddRemark={addRemark}
        onDeleteRemark={deleteRemark}
        currentUser={{ name: adminName, role: currentUser?.role || 'admin' }}
      />
    </div>
  );
};
