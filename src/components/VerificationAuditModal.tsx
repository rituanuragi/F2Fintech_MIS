import React from 'react';
import { X, ShieldCheck, ShieldAlert, ShieldX, CheckCircle2, XCircle } from 'lucide-react';

export interface VerificationAuditReport {
  sourceUrl: string;
  endpointIdentifier: string;
  syncTimestamp: string;
  sourceReportedTotal: number | null;
  totalBatchesProcessed: number;
  totalRawRecordsFetched: number;
  totalUniqueRecordsFetched: number;
  duplicatesDetected: number;
  recordsInsertedOrUpdated: number;
  failedRequestsAndRetries: {
    failedRequests: number;
    retriesCount: number;
    errors: string[];
  };
  paginationCompleted: boolean;
  recordCountChecksPassed: boolean;
  status: 'SUCCESS' | 'PARTIAL' | 'FAILED';
  verificationDetails: string;
  unresolvedDiscrepancies: string[];
}

interface Props {
  isOpen: boolean;
  onClose: () => void;
  report: VerificationAuditReport | null;
  title?: string;
}

const fmt = (n: number | null | undefined) => (typeof n === 'number' ? n.toLocaleString() : '—');

const STATUS_STYLE = {
  SUCCESS: { chip: 'bg-emerald-50 text-emerald-700 border-emerald-200', icon: ShieldCheck, label: 'Verified' },
  PARTIAL: { chip: 'bg-amber-50 text-amber-700 border-amber-200', icon: ShieldAlert, label: 'Partly verified' },
  FAILED: { chip: 'bg-rose-50 text-rose-700 border-rose-200', icon: ShieldX, label: 'Failed' },
} as const;

export const VerificationAuditModal: React.FC<Props> = ({ isOpen, onClose, report, title }) => {
  if (!isOpen || !report) return null;

  const st = STATUS_STYLE[report.status] || STATUS_STYLE.PARTIAL;
  const StatusIcon = st.icon;
  const failed = report.failedRequestsAndRetries || { failedRequests: 0, retriesCount: 0, errors: [] };

  // Each line is one check, shown as pass or fail.
  const checks: { label: string; ok: boolean; detail: string }[] = [
    {
      label: 'Record count matches the source',
      ok: report.recordCountChecksPassed,
      detail: `Source reported ${fmt(report.sourceReportedTotal)} · stored ${fmt(report.recordsInsertedOrUpdated)}`,
    },
    {
      label: 'All tabs and pages were read',
      ok: report.paginationCompleted,
      detail: `${fmt(report.totalBatchesProcessed)} tabs processed`,
    },
    {
      label: 'No failed requests',
      ok: failed.failedRequests === 0,
      detail: `${fmt(failed.failedRequests)} failed · ${fmt(failed.retriesCount)} retries`,
    },
    {
      label: 'No unresolved differences',
      ok: (report.unresolvedDiscrepancies || []).length === 0,
      detail: `${(report.unresolvedDiscrepancies || []).length} open`,
    },
  ];

  const numbers: { label: string; value: string }[] = [
    { label: 'Source total', value: fmt(report.sourceReportedTotal) },
    { label: 'Raw records fetched', value: fmt(report.totalRawRecordsFetched) },
    { label: 'Unique records', value: fmt(report.totalUniqueRecordsFetched) },
    { label: 'Duplicates found', value: fmt(report.duplicatesDetected) },
    { label: 'Stored or updated', value: fmt(report.recordsInsertedOrUpdated) },
    { label: 'Tabs processed', value: fmt(report.totalBatchesProcessed) },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs" onClick={onClose}>
      <div
        className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-3xl overflow-hidden flex flex-col max-h-[90vh]"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between gap-3">
          <div className="flex items-center gap-2.5 min-w-0">
            <div className="p-2 rounded-xl bg-indigo-500/20 text-indigo-300 border border-indigo-400/30">
              <ShieldCheck className="w-5 h-5" />
            </div>
            <div className="min-w-0">
              <h3 className="text-sm font-black tracking-tight truncate">{title || 'Data Verification Audit'}</h3>
              <p className="text-[11px] text-slate-400 font-mono truncate">{report.endpointIdentifier}</p>
            </div>
          </div>
          <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-white/10 cursor-pointer" aria-label="Close">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="p-6 space-y-5 overflow-y-auto">
          <div className="flex flex-wrap items-center gap-3">
            <span className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-full border text-xs font-bold ${st.chip}`}>
              <StatusIcon className="w-3.5 h-3.5" /> {st.label}
            </span>
            <span className="text-[11px] text-slate-500 font-mono">
              Synced {report.syncTimestamp ? new Date(report.syncTimestamp).toLocaleString() : '—'}
            </span>
          </div>

          {report.verificationDetails && (
            <p className="text-sm text-slate-700 bg-slate-50 border border-slate-200 rounded-xl p-3.5">{report.verificationDetails}</p>
          )}

          <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
            {numbers.map((n) => (
              <div key={n.label} className="border border-slate-200 rounded-xl p-3">
                <div className="text-[10px] uppercase tracking-wider font-bold text-slate-400">{n.label}</div>
                <div className="text-lg font-black text-slate-900 font-mono mt-0.5">{n.value}</div>
              </div>
            ))}
          </div>

          <div className="border border-slate-200 rounded-xl divide-y divide-slate-100">
            {checks.map((c) => (
              <div key={c.label} className="flex items-center gap-3 px-4 py-2.5">
                {c.ok ? <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" /> : <XCircle className="w-4 h-4 text-rose-600 shrink-0" />}
                <span className="text-xs font-bold text-slate-900 flex-1">{c.label}</span>
                <span className="text-[11px] text-slate-500 font-mono">{c.detail}</span>
              </div>
            ))}
          </div>

          {failed.errors.length > 0 && (
            <div>
              <div className="text-[10px] uppercase tracking-wider font-bold text-rose-600 mb-1.5">Request errors</div>
              <ul className="space-y-1 text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-xl p-3">
                {failed.errors.map((e, i) => <li key={i} className="font-mono break-words">{e}</li>)}
              </ul>
            </div>
          )}

          {(report.unresolvedDiscrepancies || []).length > 0 && (
            <div>
              <div className="text-[10px] uppercase tracking-wider font-bold text-amber-600 mb-1.5">Unresolved differences</div>
              <ul className="space-y-1 text-xs text-amber-800 bg-amber-50 border border-amber-200 rounded-xl p-3">
                {report.unresolvedDiscrepancies.map((d, i) => <li key={i}>{d}</li>)}
              </ul>
            </div>
          )}

          {report.sourceUrl && (
            <div className="text-[11px] text-slate-400 font-mono break-all">Source: {report.sourceUrl}</div>
          )}
        </div>
      </div>
    </div>
  );
};
