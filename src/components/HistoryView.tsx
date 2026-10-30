import React, { useState, useEffect } from 'react';
import {
  History,
  AlertTriangle,
  CheckCircle2,
  Clock,
  FileSpreadsheet,
  Layers,
  Database,
  RefreshCw,
} from 'lucide-react';
import { SyncRunItem, SyncErrorItem } from '../types';

interface HistoryViewProps {
  initialSection?: 'runs' | 'errors';
  onTriggerSync?: () => void;
}

export const HistoryView: React.FC<HistoryViewProps> = ({
  initialSection = 'runs',
  onTriggerSync,
}) => {
  const [runs, setRuns] = useState<SyncRunItem[]>([]);
  const [errors, setErrors] = useState<SyncErrorItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const loadHistory = async () => {
    setLoading(true);
    setLoadError(null);
    try {
      const [rRes, eRes] = await Promise.all([
        fetch('/api/history'),
        fetch('/api/errors'),
      ]);
      const rData = await rRes.json().catch(() => null);
      const eData = await eRes.json().catch(() => null);
      if (!rRes.ok || !rData) {
        throw new Error((rData && rData.error) || `Could not load sync history (HTTP ${rRes.status})`);
      }
      if (!eRes.ok || !eData) {
        throw new Error((eData && eData.error) || `Could not load error log (HTTP ${eRes.status})`);
      }
      setRuns(Array.isArray(rData.runs) ? rData.runs : []);
      setErrors(Array.isArray(eData.errors) ? eData.errors : []);
    } catch (err) {
      console.error(err);
      setRuns([]);
      setErrors([]);
      setLoadError(err instanceof Error ? err.message : 'Could not load the audit logs.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadHistory();
  }, []);

  const formatDate = (isoStr?: string | null) => {
    if (!isoStr) return 'In progress...';
    return new Date(isoStr).toLocaleString('en-IN', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: true,
    });
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200/90 pb-4">
        <div>
          <h2 className="text-xl font-bold text-slate-900 tracking-tight flex items-center gap-2">
            <History className="w-5 h-5 text-blue-600" />
            Audit &amp; Sync History
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            Complete operational audit log of all Google Drive synchronizations, indexed rows, and warning diagnostics.
          </p>
        </div>

        <div className="flex items-center gap-2">
          {onTriggerSync && (
            <button
              onClick={onTriggerSync}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg shadow-2xs transition-colors cursor-pointer"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Run Incremental Sync</span>
            </button>
          )}
          <button
            onClick={loadHistory}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg border border-slate-200 transition-colors cursor-pointer"
          >
            <RefreshCw className="w-3.5 h-3.5 text-slate-500" />
            <span>Refresh Audit Logs</span>
          </button>
        </div>
      </div>

      {loadError && (
        <div className="p-3.5 rounded-xl bg-rose-50 border border-rose-200 text-xs text-rose-800 font-semibold">
          Could not load the logs: {loadError}. The tables below may be empty because of this error.
        </div>
      )}

      {/* Sync Runs Table */}
      <div className={`bg-white border rounded-2xl overflow-hidden shadow-xs ${initialSection === 'runs' ? 'border-blue-300 ring-1 ring-blue-100' : 'border-slate-200/90'}`}>
        <div className="p-4 border-b border-slate-100 bg-slate-50/70 font-semibold text-xs text-slate-700 uppercase tracking-wider font-mono">
          Recent Sync Operations
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-50 text-slate-500 uppercase font-mono text-[10px] tracking-wider border-b border-slate-100">
              <tr>
                <th className="py-3 px-4">Status</th>
                <th className="py-3 px-4">Started At</th>
                <th className="py-3 px-4">Completed At</th>
                <th className="py-3 px-4 text-right">Duration</th>
                <th className="py-3 px-4 text-right">Files</th>
                <th className="py-3 px-4 text-right">Sheets</th>
                <th className="py-3 px-4 text-right">Rows Ingested</th>
                <th className="py-3 px-4 text-right">Errors</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 font-mono text-slate-700">
              {loading ? (
                <tr>
                  <td colSpan={8} className="py-8 text-center text-slate-400">
                    Loading sync history...
                  </td>
                </tr>
              ) : runs.length === 0 ? (
                <tr>
                  <td colSpan={8} className="py-8 text-center text-slate-400">
                    {loadError ? 'Sync history could not be loaded.' : 'No sync operations recorded yet.'}
                  </td>
                </tr>
              ) : (
                runs.map((run) => (
                  <tr key={run.id} className="hover:bg-slate-50 transition-colors">
                    <td className="py-3 px-4">
                      {run.status === 'completed' ? (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">
                          <CheckCircle2 className="w-3 h-3" /> Completed
                        </span>
                      ) : run.status === 'running' ? (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-200 animate-pulse">
                          Running
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-50 text-rose-700 border border-rose-200">
                          <AlertTriangle className="w-3 h-3" /> Failed
                        </span>
                      )}
                    </td>
                    <td className="py-3 px-4 text-slate-600">{formatDate(run.startedAt)}</td>
                    <td className="py-3 px-4 text-slate-500">{formatDate(run.completedAt)}</td>
                    <td className="py-3 px-4 text-right text-slate-600">
                      {run.durationMs ? `${(run.durationMs / 1000).toFixed(1)}s` : '—'}
                    </td>
                    <td className="py-3 px-4 text-right font-bold text-slate-800">
                      {run.filesProcessed} / {run.filesDiscovered}
                    </td>
                    <td className="py-3 px-4 text-right text-slate-800 font-bold">
                      {run.sheetsProcessed}
                    </td>
                    <td className="py-3 px-4 text-right font-bold text-blue-700">
                      {run.rowsProcessed.toLocaleString()}
                    </td>
                    <td className="py-3 px-4 text-right">
                      {run.errorsCount > 0 ? (
                        <span className="text-rose-600 font-bold">{run.errorsCount}</span>
                      ) : (
                        <span className="text-slate-400">0</span>
                      )}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Sync Errors & Diagnostics Table */}
      <div className="bg-white border border-slate-200/90 rounded-2xl overflow-hidden shadow-xs">
        <div className="p-4 border-b border-slate-100 bg-rose-50/50 font-semibold text-xs text-rose-700 uppercase tracking-wider font-mono flex items-center gap-2">
          <AlertTriangle className="w-4 h-4 text-rose-600" />
          <span>Error Log &amp; Workbook Diagnostics ({errors.length})</span>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-50 text-slate-500 uppercase font-mono text-[10px] tracking-wider border-b border-slate-100">
              <tr>
                <th className="py-3 px-4">Workbook Name</th>
                <th className="py-3 px-4">Sheet</th>
                <th className="py-3 px-4">Diagnostic Details</th>
                <th className="py-3 px-4 text-right">Timestamp</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 font-mono text-slate-700">
              {errors.length === 0 ? (
                <tr>
                  {loadError ? (
                    <td colSpan={4} className="py-8 text-center text-rose-600 font-sans">
                      Error log could not be loaded, so sync status is unknown.
                    </td>
                  ) : (
                    <td colSpan={4} className="py-8 text-center text-emerald-600 font-sans">
                      ✓ Clean sync status — No corrupted or unreadable files encountered.
                    </td>
                  )}
                </tr>
              ) : (
                errors.map((err) => (
                  <tr key={err.id} className="hover:bg-slate-50 transition-colors">
                    <td className="py-3 px-4 font-bold text-slate-800">
                      {err.fileName || 'Unknown File'}
                    </td>
                    <td className="py-3 px-4 text-blue-700">
                      {err.sheetName || 'Whole File'}
                    </td>
                    <td className="py-3 px-4 text-rose-700 font-sans text-xs">
                      {err.errorMessage}
                    </td>
                    <td className="py-3 px-4 text-right text-slate-400 text-[11px]">
                      {formatDate(err.timestamp)}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
