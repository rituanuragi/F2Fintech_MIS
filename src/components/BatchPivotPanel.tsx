import React, { useMemo } from 'react';
import { Table2, RotateCcw, Pencil } from 'lucide-react';
import { buildPivot, defaultPivotHeaders, PivotHeaders, PivotInput } from '../utils/pivot';

interface BatchPivotPanelProps {
  results: PivotInput[];
  headers: PivotHeaders;
  onHeadersChange: (next: PivotHeaders) => void;
  includePivot: boolean;
  onIncludePivotChange: (v: boolean) => void;
  useTitleAsTabName: boolean;
  onUseTitleAsTabNameChange: (v: boolean) => void;
}

const FIELDS: { key: keyof PivotHeaders; label: string }[] = [
  { key: 'title', label: 'Pivot title (top banner)' },
  { key: 'dataHeader', label: 'First column header' },
  { key: 'countHeader', label: 'Second column header' },
  { key: 'blanksLabel', label: 'Blanks row name' },
  { key: 'totalLabel', label: 'Total row name' },
];

export const BatchPivotPanel: React.FC<BatchPivotPanelProps> = ({
  results, headers, onHeadersChange, includePivot, onIncludePivotChange, useTitleAsTabName, onUseTitleAsTabNameChange,
}) => {
  const pivot = useMemo(() => buildPivot(results), [results]);
  // an empty box falls back to the default text, same as the Excel file does
  const defaults = defaultPivotHeaders();
  const show = (k: keyof PivotHeaders) => (headers[k] || '').trim() || defaults[k];

  return (
    <div className="bg-white border border-slate-200/90 rounded-2xl shadow-xs overflow-hidden">
      <div className="px-5 py-3.5 border-b border-slate-100 flex flex-wrap items-center gap-3">
        <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
          <Table2 className="w-4 h-4 text-violet-600" />
          Pivot Summary
        </h3>
        <span className="text-[11px] text-slate-500">Count of numbers by Priority. This is the same table that goes in the Excel file.</span>
        <button
          type="button"
          onClick={() => onHeadersChange(defaultPivotHeaders())}
          className="ml-auto inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-[11px] font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 cursor-pointer"
        >
          <RotateCcw className="w-3 h-3" /> Reset headers
        </button>
      </div>

      <div className="grid lg:grid-cols-[minmax(0,1fr)_320px] gap-5 p-5">
        <div className="space-y-3">
          <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
            <Pencil className="w-3 h-3" /> Edit headers (before you download)
          </div>
          <div className="grid sm:grid-cols-2 gap-3">
            {FIELDS.map((f) => (
              <label key={f.key} className={`block text-[11px] font-semibold text-slate-600 ${f.key === 'title' ? 'sm:col-span-2' : ''}`}>
                {f.label}
                <input
                  value={headers[f.key]}
                  onChange={(e) => onHeadersChange({ ...headers, [f.key]: e.target.value })}
                  placeholder={defaults[f.key]}
                  maxLength={120}
                  className="mt-1 w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-xs text-slate-900 focus:outline-none focus:border-violet-500 focus:bg-white"
                />
              </label>
            ))}
          </div>
          <div className="flex flex-col gap-1.5 pt-1">
            <label className="inline-flex items-center gap-2 text-xs text-slate-700 font-medium cursor-pointer">
              <input type="checkbox" checked={includePivot} onChange={(e) => onIncludePivotChange(e.target.checked)} />
              Add the Pivot sheet to the Excel download
            </label>
            <label className={`inline-flex items-center gap-2 text-xs font-medium cursor-pointer ${includePivot ? 'text-slate-700' : 'text-slate-400'}`}>
              <input
                type="checkbox"
                checked={useTitleAsTabName}
                disabled={!includePivot}
                onChange={(e) => onUseTitleAsTabNameChange(e.target.checked)}
              />
              Use the pivot title as the sheet tab name
            </label>
          </div>
        </div>

        <div className="overflow-hidden rounded-xl border border-slate-300 self-start">
          <table className="w-full text-xs">
            <thead>
              <tr>
                <th colSpan={2} className="px-3 py-2.5 text-center text-sm font-bold text-white bg-purple-900">{show('title')}</th>
              </tr>
              <tr className="bg-white border-b border-slate-300">
                <th className="px-3 py-2 text-left font-bold text-slate-900">{show('dataHeader')}</th>
                <th className="px-3 py-2 text-right font-bold text-slate-900">{show('countHeader')}</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {pivot.rows.map((r) => (
                <tr key={r.label}>
                  <td className="px-3 py-2 text-slate-800">{r.label}</td>
                  <td className="px-3 py-2 text-right font-mono text-slate-800">{r.count.toLocaleString()}</td>
                </tr>
              ))}
              <tr>
                <td className="px-3 py-2 text-slate-800">{show('blanksLabel')}</td>
                <td className="px-3 py-2 text-right font-mono text-slate-800">{pivot.blanks.toLocaleString()}</td>
              </tr>
              <tr className="bg-violet-100 font-bold">
                <td className="px-3 py-2 text-slate-900">{show('totalLabel')}</td>
                <td className="px-3 py-2 text-right font-mono text-slate-900">{pivot.total.toLocaleString()}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
