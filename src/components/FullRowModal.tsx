import React, { useState, useEffect } from 'react';
import {
  X,
  Copy,
  Check,
  Folder,
  FileSpreadsheet,
  Layers,
  Search,
  ExternalLink,
  Table,
  Ticket,
  ArrowRight,
} from 'lucide-react';
import { SearchMatchItem, LiveTicketMatchInfo } from '../types';

interface FullRowModalProps {
  recordId: number | string | null;
  onClose: () => void;
  onPreviewSheet?: (fileId: string, sheetName: string) => void;
  onLookupCustomer?: (phone: string) => void;
  onNavigateToTicket?: (ticketId: string | number) => void;
}

export const FullRowModal: React.FC<FullRowModalProps> = ({
  recordId,
  onClose,
  onPreviewSheet,
  onLookupCustomer,
  onNavigateToTicket,
}) => {
  const [loading, setLoading] = useState(false);
  const [data, setData] = useState<{
    match: SearchMatchItem | null;
    originalRow: Record<string, any>;
    sourceBreadcrumbs: {
      sourceName: string;
      folderPath: string;
      fileName: string;
      sheetName: string;
      rowNum: number;
    } | null;
  } | null>(null);
  const [liveTicketMatch, setLiveTicketMatch] = useState<LiveTicketMatchInfo | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [copied, setCopied] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);

  useEffect(() => {
    if (!recordId) return;
    setLiveTicketMatch(null);
    setLoadError(null);
    const fetchRow = async () => {
      setLoading(true);
      try {
        const res = await fetch(`/api/record/${encodeURIComponent(String(recordId))}`);
        const result = await res.json().catch(() => null);
        if (!res.ok || !result || !result.originalRow) {
          throw new Error((result && result.error) || `Could not load record (HTTP ${res.status})`);
        }
        setData(result);

        // Use the phone the search already matched. Only fall back to a column
        // whose NAME looks like a phone column (never guess from any 10-digit value).
        const phoneFromColumn = result.originalRow
          ? Object.entries(result.originalRow).find(
              ([k, v]) =>
                /phone|mobile|contact|cell/i.test(k) &&
                /^\d{10}$/.test(String(v ?? '').replace(/\D/g, '').slice(-10))
            )?.[1]
          : undefined;
        const phone =
          result?.match?.normPhone ||
          result?.match?.phoneRaw ||
          phoneFromColumn;
        const lender =
          result?.match?.lender ||
          (result?.originalRow &&
            Object.entries(result.originalRow).find(([k]) =>
              /lender|bank|provider/i.test(k)
            )?.[1]);
        const name =
          result?.match?.customerName ||
          (result?.originalRow &&
            Object.entries(result.originalRow).find(([k]) =>
              /name|customer/i.test(k)
            )?.[1]);
        const ticketId =
          (result?.originalRow &&
            Object.entries(result.originalRow).find(([k]) =>
              /^(ticket\s*id|ticket\s*no|ticket|lead\s*id|app\s*id)$/i.test(k)
            )?.[1]);

        if (phone || lender || name || ticketId) {
          fetch(
            `/api/f2-tickets/check-match?phone=${encodeURIComponent(
              String(phone || '')
            )}&lender=${encodeURIComponent(String(lender || ''))}&name=${encodeURIComponent(
              String(name || '')
            )}&ticketId=${encodeURIComponent(String(ticketId || ''))}`
          )
            .then((r) => r.json())
            .then((m) => {
              if (m?.matched && m?.match) {
                setLiveTicketMatch(m.match);
              }
            })
            .catch(() => {});
        }
      } catch (err) {
        console.error(err);
        setData(null);
        setLoadError(err instanceof Error ? err.message : 'Could not load this record.');
      } finally {
        setLoading(false);
      }
    };
    fetchRow();
  }, [recordId]);

  if (!recordId) return null;

  const entries = data?.originalRow ? Object.entries(data.originalRow) : [];
  const filteredEntries = entries.filter(
    ([key, val]) =>
      key.toLowerCase().includes(searchTerm.toLowerCase()) ||
      String(val).toLowerCase().includes(searchTerm.toLowerCase())
  );

  const copyRowJson = () => {
    if (!data) return;
    navigator.clipboard.writeText(JSON.stringify(data.originalRow, null, 2));
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-4xl max-h-[90vh] flex flex-col shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        {/* Header */}
        <div className="p-5 border-b border-slate-800 flex items-start justify-between gap-4 bg-slate-950/60">
          <div>
            <div className="flex items-center gap-2 text-xs font-mono text-emerald-400 font-bold uppercase tracking-wider mb-1">
              <Table className="w-4 h-4" />
              <span>Complete Original Record</span>
            </div>
            <h2 className="text-lg font-bold text-white tracking-tight">
              Excel Row {data?.sourceBreadcrumbs?.rowNum} — All Columns &amp; Values
            </h2>

            {data?.sourceBreadcrumbs && (
              <div className="flex flex-wrap items-center gap-2 text-xs font-mono text-slate-400 mt-2">
                <span className="flex items-center gap-1 text-slate-300">
                  <Folder className="w-3.5 h-3.5 text-emerald-400" />
                  {data.sourceBreadcrumbs.folderPath}
                </span>
                <span>/</span>
                <span className="flex items-center gap-1 text-slate-200 font-semibold">
                  <FileSpreadsheet className="w-3.5 h-3.5 text-blue-400" />
                  {data.sourceBreadcrumbs.fileName}
                </span>
                <span>/</span>
                <span className="flex items-center gap-1 text-purple-300">
                  <Layers className="w-3.5 h-3.5 text-purple-400" />
                  {data.sourceBreadcrumbs.sheetName}
                </span>
              </div>
            )}
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={copyRowJson}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-medium border border-slate-700 transition-colors cursor-pointer"
            >
              {copied ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
              <span>{copied ? 'Copied JSON' : 'Copy Row'}</span>
            </button>
            <button
              onClick={onClose}
              className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Clean Live Ticket Button: Only shown if data is present in live API */}
        {liveTicketMatch && onNavigateToTicket && (
          <div className="mx-5 mt-3 flex items-center justify-end">
            <button
              type="button"
              onClick={() => {
                onClose();
                onNavigateToTicket(liveTicketMatch.ticketId);
              }}
              className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs transition cursor-pointer shadow-xs"
              title={`Open Live Ticket #${liveTicketMatch.ticketId} in Live Tickets API`}
            >
              <Ticket className="w-3.5 h-3.5 text-amber-300" />
              <span>Live Ticket #{liveTicketMatch.ticketId}</span>
              <ArrowRight className="w-3.5 h-3.5 text-blue-200" />
            </button>
          </div>
        )}

        {/* Filter bar */}
        <div className="px-5 py-3 border-b border-slate-800/80 bg-slate-900/60 flex items-center justify-between gap-4">
          <div className="relative flex-1 max-w-sm">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="Filter columns or values..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-8 pr-3 py-1.5 bg-slate-950 border border-slate-800 rounded-lg text-xs text-slate-200 placeholder:text-slate-500 focus:outline-none focus:border-emerald-500"
            />
          </div>

          <span className="text-xs text-slate-400 font-mono">
            {filteredEntries.length} of {entries.length} columns shown
          </span>
        </div>

        {/* Column table */}
        <div className="flex-1 overflow-y-auto p-5">
          {loading ? (
            <div className="text-center py-12 text-slate-500 text-xs">
              Loading original record data...
            </div>
          ) : loadError ? (
            <div className="text-center py-12 text-rose-400 text-xs font-semibold">
              Could not load this record: {loadError}
            </div>
          ) : (
            <div className="border border-slate-800 rounded-xl overflow-hidden shadow-inner">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-800/80 text-slate-400 uppercase font-mono text-[10px] tracking-wider border-b border-slate-800">
                  <tr>
                    <th className="py-2.5 px-4 w-1/3">Original Column Name</th>
                    <th className="py-2.5 px-4">Original Cell Value</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60 font-mono">
                  {filteredEntries.map(([col, val], idx) => (
                    <tr key={idx} className="hover:bg-slate-800/30 transition-colors">
                      <td className="py-2.5 px-4 text-slate-400 font-semibold bg-slate-950/40 select-all">
                        {col}
                      </td>
                      <td className="py-2.5 px-4 text-slate-100 select-all font-sans break-words">
                        {val === null || val === undefined || val === '' ? (
                          <span className="text-slate-600 italic font-mono text-[11px]">&lt;blank cell&gt;</span>
                        ) : (
                          String(val)
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="p-4 border-t border-slate-800 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-400 bg-slate-950/40 font-mono">
          <span>Excel Row #{data?.sourceBreadcrumbs?.rowNum}</span>
          <div className="flex flex-wrap items-center gap-2">
            {onPreviewSheet && data?.sourceBreadcrumbs && (
              <button
                onClick={() => {
                  const targetFile = data.match?.fileId || data.sourceBreadcrumbs!.fileName;
                  const targetSheet = data.sourceBreadcrumbs!.sheetName;
                  onClose();
                  onPreviewSheet(targetFile, targetSheet);
                }}
                className="px-3.5 py-1.5 rounded-lg bg-emerald-600/20 hover:bg-emerald-600 text-emerald-300 hover:text-white border border-emerald-500/30 font-semibold transition-colors cursor-pointer"
              >
                Preview Worksheet
              </button>
            )}
            {onLookupCustomer && (data?.match?.normPhone || data?.match?.phoneRaw || data?.match?.customerName) && (
              <button
                onClick={() => {
                  const q = data.match?.normPhone || data.match?.phoneRaw || data.match?.customerName || '';
                  onClose();
                  onLookupCustomer(q);
                }}
                className="px-3.5 py-1.5 rounded-lg bg-indigo-600/20 hover:bg-indigo-600 text-indigo-300 hover:text-white border border-indigo-500/30 font-semibold transition-colors cursor-pointer"
              >
                Customer Lookup &amp; Diff
              </button>
            )}
            <button
              onClick={onClose}
              className="px-4 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 transition-colors cursor-pointer"
            >
              Close
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
