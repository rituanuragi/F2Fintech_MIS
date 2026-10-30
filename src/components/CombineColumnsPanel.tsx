import React, { useState } from 'react';
import { X, Save, Link2, Plus, Check, Trash2 } from 'lucide-react';
import { LeadPriorityRule } from '../types';

// ---------------------------------------------------------------------------
// Combine Columns panel: choose which columns are added to the Priority label
// ---------------------------------------------------------------------------
const SEPARATORS: { value: string; label: string }[] = [
  { value: '-', label: 'Dash  ( - )' },
  { value: ' ', label: 'Space' },
  { value: '_', label: 'Underscore  ( _ )' },
  { value: '|', label: 'Bar  ( | )' },
  { value: '/', label: 'Slash  ( / )' },
];

const AUTO_COLUMN: Record<string, string> = { neet: 'Year', lot: 'Lot', login: 'Login' };

const autoPrefix = (rule: LeadPriorityRule): string => {
  if (rule.categoryKey === 'neet') return 'NEET';
  if (rule.categoryKey === 'lot') return 'Lot';
  if (rule.categoryKey === 'login') return 'Login';
  return (rule.name || '').replace(/\.(xlsx?|xlsm?|csv)$/i, '').trim();
};

const previewLabel = (rule: LeadPriorityRule): string => {
  const sep = rule.combineSeparator ?? '-';
  const prefix = (rule.combineLabel || '').trim() || autoPrefix(rule);
  const cols = (rule.combineColumns || []).filter(Boolean);
  const parts = cols.length ? cols.map((c) => `{${c}}`) : AUTO_COLUMN[rule.categoryKey] ? [`{${AUTO_COLUMN[rule.categoryKey]}}`] : [];
  return [prefix, ...parts].join(sep);
};

interface CombinePanelProps {
  rules: LeadPriorityRule[];
  columnsByRule: Record<string, string[]>;
  allColumns: string[];
  loading: boolean;
  saving: boolean;
  onChange: (ruleId: string, patch: Partial<LeadPriorityRule>) => void;
  onClose: () => void;
  onSave: () => void;
  onAddRule: (name: string, pattern: string) => void;
  onDeleteRule: (ruleId: string) => void;
}

export const CombineColumnsPanel: React.FC<CombinePanelProps> = ({
  rules, columnsByRule, allColumns, loading, saving, onChange, onClose, onSave, onAddRule, onDeleteRule,
}) => {
  const [newName, setNewName] = useState('');
  const [newPattern, setNewPattern] = useState('');
  const submitNew = () => {
    if (!newName.trim()) return;
    onAddRule(newName.trim(), (newPattern.trim() || newName.trim()));
    setNewName('');
    setNewPattern('');
  };
  const [filter, setFilter] = useState('');
  const [extra, setExtra] = useState<Record<string, string>>({});

  const visible = rules.filter((r) => !filter.trim() || r.name.toLowerCase().includes(filter.trim().toLowerCase()));

  const toggleColumn = (rule: LeadPriorityRule, col: string) => {
    const cur = rule.combineColumns || [];
    const next = cur.some((c) => c.toLowerCase() === col.toLowerCase())
      ? cur.filter((c) => c.toLowerCase() !== col.toLowerCase())
      : [...cur, col];
    onChange(rule.id, { combineColumns: next });
  };

  const addExtra = (rule: LeadPriorityRule) => {
    const v = (extra[rule.id] || '').trim();
    if (!v) return;
    const cur = rule.combineColumns || [];
    if (!cur.some((c) => c.toLowerCase() === v.toLowerCase())) {
      onChange(rule.id, { combineColumns: [...cur, v] });
    }
    setExtra((e) => ({ ...e, [rule.id]: '' }));
  };

  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-slate-950/85 backdrop-blur-sm">
      <div className="bg-slate-900 border border-slate-700 rounded-2xl w-full max-w-3xl max-h-[90vh] flex flex-col shadow-2xl shadow-black/60 overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-sky-500/10 border border-sky-500/30 text-sky-400 flex items-center justify-center">
              <Link2 className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-bold text-white">Combine Columns in Priority</h3>
              <p className="text-xs text-slate-400">
                Pick the columns whose values are added to the Priority column in Global Search and Batch Search.
              </p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800 cursor-pointer">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="p-6 overflow-y-auto space-y-3 flex-1">
          <div className="p-3 rounded-xl bg-slate-800/40 border border-slate-700/60 text-xs text-slate-300 space-y-1">
            <p>
              Example: <span className="font-mono text-purple-300">NEET</span> + column <span className="font-mono text-purple-300">Year</span> shows{' '}
              <span className="font-mono text-emerald-300">NEET-2023</span>. <span className="font-mono text-amber-300">LOT</span> + <span className="font-mono text-amber-300">lots</span> shows{' '}
              <span className="font-mono text-emerald-300">Lot-1</span>. <span className="font-mono text-orange-300">LOGIN</span> + <span className="font-mono text-orange-300">Month</span> shows{' '}
              <span className="font-mono text-emerald-300">Login-May</span>.
            </p>
            <p className="text-slate-400">
              NEET, LOT and LOGIN already read their own year / lot / month column. Choose columns only when you want something different or more. Run your search again to see the new labels.
            </p>
          </div>

          <input
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
            placeholder="Find a file or directory rule..."
            className="w-full px-3 py-2 rounded-lg bg-slate-800 border border-slate-700 text-xs text-slate-200 placeholder-slate-500 outline-none focus:border-sky-500"
          />


          <div className="p-3.5 rounded-xl bg-slate-800/40 border border-dashed border-slate-600 space-y-2">
            <div className="text-[11px] font-bold text-slate-200">Add a new combination</div>
            <p className="text-[10px] text-slate-400">
              For a file or folder that is not in the list below (e.g. <span className="font-mono text-slate-300">DENTAL</span>). Its name only has to contain the text you type. Then pick its columns in the new card.
            </p>
            <div className="grid grid-cols-2 gap-2">
              <input
                value={newName}
                onChange={(e) => setNewName(e.target.value)}
                placeholder="Name, e.g. DENTAL"
                className="px-2.5 py-1.5 rounded-lg bg-slate-900 border border-slate-700 text-xs text-slate-200 placeholder-slate-600 outline-none focus:border-sky-500"
              />
              <input
                value={newPattern}
                onChange={(e) => setNewPattern(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); submitNew(); } }}
                placeholder="File / folder contains (optional)"
                className="px-2.5 py-1.5 rounded-lg bg-slate-900 border border-slate-700 text-xs text-slate-200 placeholder-slate-600 outline-none focus:border-sky-500"
              />
            </div>
            <button
              type="button"
              onClick={submitNew}
              disabled={!newName.trim()}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[11px] font-bold bg-sky-600 hover:bg-sky-500 text-white cursor-pointer disabled:opacity-40"
            >
              <Plus className="w-3.5 h-3.5" /> Add combination
            </button>
          </div>

          {loading && <div className="py-6 text-center text-xs text-slate-500">Reading column names from your files...</div>}

          {visible.map((rule) => {
            const available = (columnsByRule[rule.id] && columnsByRule[rule.id].length ? columnsByRule[rule.id] : allColumns).slice(0, 60);
            const selected = rule.combineColumns || [];
            const hasCombo = selected.length > 0;
            return (
              <div key={rule.id} className={`p-3.5 rounded-xl border space-y-3 ${hasCombo ? 'bg-sky-950/20 border-sky-500/40' : 'bg-slate-800/50 border-slate-700/60'}`}>
                <div className="flex items-center justify-between gap-3 flex-wrap">
                  <div className="flex items-center gap-2 min-w-0">
                    <span className="font-bold text-slate-100 text-xs truncate">{rule.name}</span>
                    <span className="text-[10px] font-mono text-slate-400">Rank #{rule.rank}</span>
                    {rule.isCustom && (
                      <button
                        type="button"
                        onClick={() => onDeleteRule(rule.id)}
                        className="p-1 rounded text-rose-300 hover:text-rose-200 hover:bg-rose-500/10 cursor-pointer"
                        title="Remove this combination"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-[10px] text-slate-400">Shows as</span>
                    <span className="px-2 py-0.5 rounded-md bg-slate-950 border border-emerald-500/30 font-mono text-[11px] text-emerald-300">{previewLabel(rule)}</span>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <label className="text-[10px] text-slate-400 space-y-1 block">
                    <span>Text before the values</span>
                    <input
                      value={rule.combineLabel ?? ''}
                      onChange={(e) => onChange(rule.id, { combineLabel: e.target.value })}
                      placeholder={autoPrefix(rule)}
                      className="w-full px-2.5 py-1.5 rounded-lg bg-slate-900 border border-slate-700 text-xs text-slate-200 outline-none focus:border-sky-500"
                    />
                  </label>
                  <label className="text-[10px] text-slate-400 space-y-1 block">
                    <span>Join with</span>
                    <select
                      value={rule.combineSeparator ?? '-'}
                      onChange={(e) => onChange(rule.id, { combineSeparator: e.target.value })}
                      className="w-full px-2.5 py-1.5 rounded-lg bg-slate-900 border border-slate-700 text-xs text-slate-200 outline-none focus:border-sky-500"
                    >
                      {SEPARATORS.map((s) => (
                        <option key={s.value} value={s.value}>{s.label}</option>
                      ))}
                    </select>
                  </label>
                </div>

                <div className="space-y-1.5">
                  <div className="text-[10px] text-slate-400 flex items-center justify-between">
                    <span>Columns to add (click to select, order = order of selection)</span>
                    {hasCombo && (
                      <button
                        type="button"
                        onClick={() => onChange(rule.id, { combineColumns: [] })}
                        className="text-rose-300 hover:text-rose-200 cursor-pointer"
                      >
                        Clear
                      </button>
                    )}
                  </div>
                  <div className="flex flex-wrap gap-1.5">
                    {selected.map((c) => (
                      <button
                        key={`sel-${c}`}
                        type="button"
                        onClick={() => toggleColumn(rule, c)}
                        className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-mono bg-sky-500/25 border border-sky-400/60 text-sky-100 cursor-pointer"
                      >
                        <Check className="w-3 h-3" />
                        {c}
                      </button>
                    ))}
                    {available
                      .filter((c) => !selected.some((s) => s.toLowerCase() === c.toLowerCase()))
                      .map((c) => (
                        <button
                          key={`av-${c}`}
                          type="button"
                          onClick={() => toggleColumn(rule, c)}
                          className="px-2 py-0.5 rounded-full text-[11px] font-mono bg-slate-900 border border-slate-700 text-slate-300 hover:border-sky-500 hover:text-sky-200 cursor-pointer"
                        >
                          {c}
                        </button>
                      ))}
                    {available.length === 0 && selected.length === 0 && (
                      <span className="text-[11px] text-slate-500">No column names found yet. Sync your files, or type a column name below.</span>
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    <input
                      value={extra[rule.id] || ''}
                      onChange={(e) => setExtra((x) => ({ ...x, [rule.id]: e.target.value }))}
                      onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); addExtra(rule); } }}
                      placeholder="Type another column name and press Enter"
                      className="flex-1 px-2.5 py-1.5 rounded-lg bg-slate-900 border border-slate-700 text-xs text-slate-200 placeholder-slate-600 outline-none focus:border-sky-500"
                    />
                    <button
                      type="button"
                      onClick={() => addExtra(rule)}
                      className="p-1.5 rounded-lg bg-slate-900 border border-slate-700 text-slate-300 hover:text-white hover:border-sky-500 cursor-pointer"
                      title="Add column"
                    >
                      <Plus className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        <div className="px-6 py-4 border-t border-slate-800 bg-slate-900/80 flex items-center justify-end gap-2.5">
          <button onClick={onClose} className="px-4 py-2 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 cursor-pointer">
            Back
          </button>
          <button
            onClick={onSave}
            disabled={saving}
            className="inline-flex items-center gap-2 px-5 py-2 rounded-xl text-xs font-bold bg-sky-600 hover:bg-sky-500 text-white shadow-lg shadow-sky-900/40 cursor-pointer disabled:opacity-50"
          >
            <Save className="w-4 h-4" />
            <span>{saving ? 'Saving...' : 'Save combinations'}</span>
          </button>
        </div>
      </div>
    </div>
  );
};

