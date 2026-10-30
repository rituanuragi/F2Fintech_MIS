import React, { useState, useEffect } from 'react';
import {
  X,
  ArrowUp,
  ArrowDown,
  ChevronsUp,
  ChevronsDown,
  RotateCcw,
  Save,
  Sliders,
  CheckCircle2,
  FileSpreadsheet,
  Layers,
  Sparkles,
  Info,
  Calendar,
  AlertCircle,
  RefreshCw,
  FolderTree,
  Folder,
  Trash2,
  Link2,
  Plus,
  Check,
} from 'lucide-react';
import { LeadPriorityRule } from '../types';
import { CombineColumnsPanel } from './CombineColumnsPanel';

const CORE_SYSTEM_RULE_IDS = new Set([
  'team_x',
  'file_jan',
  'file_may',
  'f2_leads',
  'neet',
  'fresh_data_21k',
  'fresh_data_30k',
  'tata_points',
  'login',
  'icic_login',
  'lot',
  'directory',
  'file_bl',
  'dir_pl',
  'dir_doctors',
  'mis_dir_pl',
  'mis_dir_doctors',
  'mis_dir_bl',
]);

const isExcludedPriorityItem = (nameOrId: string): boolean => {
  if (!nameOrId) return true;
  const s = String(nameOrId).toLowerCase().trim();
  return (
    s === 'sample_neet' ||
    s === 'file_doctor_list' ||
    s === 'doctor_list' ||
    s.includes('master mis') ||
    s.startsWith('gsheet_') ||
    s.includes('google sheets mis') ||
    s.includes('f2_tickets') ||
    s.includes('session.json') ||
    s.includes('cache')
  );
};

const isValidRule = (r: LeadPriorityRule): boolean => {
  if (!r) return false;
  if (isExcludedPriorityItem(r.id) || isExcludedPriorityItem(r.name) || isExcludedPriorityItem(r.categoryKey)) return false;
  if (r.description && isExcludedPriorityItem(r.description)) return false;
  return true;
};


interface PriorityConfigModalProps {
  isOpen: boolean;
  onClose: () => void;
  onRulesUpdated?: (rules: LeadPriorityRule[]) => void;
  scope?: 'global' | 'mis';
  title?: string;
  description?: string;
  /** Open straight on the Combine Columns screen (used by the MIS screen). */
  openCombineOnOpen?: boolean;
}

export const PriorityConfigModal: React.FC<PriorityConfigModalProps> = ({
  isOpen,
  onClose,
  onRulesUpdated,
  scope = 'global',
  title,
  description,
  openCombineOnOpen = false,
}) => {
  const [rules, setRules] = useState<LeadPriorityRule[]>([]);
  const [defaultRules, setDefaultRules] = useState<LeadPriorityRule[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);
  const [syncFeedback, setSyncFeedback] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [neetYearOrder, setNeetYearOrder] = useState<'desc' | 'asc'>('desc');
  const [lastSyncedTime, setLastSyncedTime] = useState<string | null>(null);
  const [showCombine, setShowCombine] = useState(false);
  const [columnsByRule, setColumnsByRule] = useState<Record<string, string[]>>({});
  const [allColumns, setAllColumns] = useState<string[]>([]);
  const [columnsLoading, setColumnsLoading] = useState(false);

  const endpoint = scope === 'mis' ? '/api/mis-priority-rules' : '/api/priority-rules';
  const syncEndpoint = scope === 'mis' ? '/api/mis-priority-rules/sync' : '/api/priority-rules/sync';
  const resetEndpoint = scope === 'mis' ? '/api/mis-priority-rules/reset' : '/api/priority-rules/reset';

  // Load rules and automatically trigger sync on open so newly added directories appear at the last
  useEffect(() => {
    if (isOpen) {
      handleSyncDirectories(true);
      if (openCombineOnOpen) openCombine();
    }
    if (!isOpen) setShowCombine(false);
  }, [isOpen, scope]);

  const fetchRules = async () => {
    setIsLoading(true);
    try {
      const res = await fetch(endpoint);
      const data = await res.json();
      if (data.rules) {
        const validRules = (data.rules as LeadPriorityRule[]).filter(isValidRule);
        setRules(validRules);
        const neetRule = validRules.find((r: LeadPriorityRule) => r.categoryKey === 'neet');
        if (neetRule) {
          setNeetYearOrder(neetRule.neetYearOrder || 'desc');
        }
      }
      if (data.defaultRules) {
        setDefaultRules(data.defaultRules);
      }
      if (data.lastSyncedAt) {
        setLastSyncedTime(new Date(data.lastSyncedAt).toLocaleTimeString());
      }
    } catch (err) {
      console.error('Failed to load priority rules:', err);
    } finally {
      setIsLoading(false);
    }
  };

  // Sync / Refresh with all directories and files:
  // If a new directory is added, it is automatically appended at the LAST of the priority list!
  const handleSyncDirectories = async (silent: boolean = false) => {
    if (!silent) setIsSyncing(true);
    else setIsLoading(true);
    setErrorMessage(null);

    try {
      const res = await fetch(syncEndpoint, { method: 'POST' });
      const data = await res.json();
      if (data.rules) {
        const validRules = (data.rules as LeadPriorityRule[]).filter(isValidRule);
        setRules(validRules);
        const neetRule = validRules.find((r: LeadPriorityRule) => r.categoryKey === 'neet');
        if (neetRule) {
          setNeetYearOrder(neetRule.neetYearOrder || 'desc');
        }
        if (onRulesUpdated) {
          onRulesUpdated(validRules);
        }
      }

      setLastSyncedTime(new Date().toLocaleTimeString());

      if (!silent) {
        if (data.addedCount && data.addedCount > 0) {
          setSyncFeedback(`Synced successfully! ${data.addedCount} new directory/source(s) were automatically added at the end of the priority list.`);
        } else {
          setSyncFeedback('All main directory files are up-to-date in the priority hierarchy.');
        }
        setTimeout(() => setSyncFeedback(null), 4500);
      }
    } catch (err: any) {
      console.error('Failed to sync priority with directories:', err);
      if (!silent) {
        setErrorMessage(`Sync failed: ${err.message}`);
        setTimeout(() => setErrorMessage(null), 4000);
      }
      // Fallback to fetch regular rules
      await fetchRules();
    } finally {
      setIsSyncing(false);
      setIsLoading(false);
    }
  };

  const openCombine = async () => {
    setShowCombine(true);
    setColumnsLoading(true);
    try {
      const res = await fetch('/api/priority-rules/columns');
      if (res.ok) {
        const data = await res.json();
        setColumnsByRule(data.columns || {});
        setAllColumns(data.all || []);
      }
    } catch (err) {
      console.warn('Could not load column names:', err);
    } finally {
      setColumnsLoading(false);
    }
  };

  const updateRuleFields = (ruleId: string, patch: Partial<LeadPriorityRule>) => {
    setRules((prev) => prev.map((r) => (r.id === ruleId ? { ...r, ...patch } : r)));
  };

  const handleAddCombination = (name: string, pattern: string) => {
    const id = `custom_${Date.now()}`;
    const newRule: LeadPriorityRule = {
      id,
      rank: rules.length + 1,
      name,
      categoryKey: id,
      description: `Custom combination for "${pattern}"`,
      filePatterns: [pattern],
      folderPatterns: [pattern],
      neetYearOrder: 'desc',
      color: '#64748b',
      badgeBg: 'bg-indigo-500/15',
      badgeText: 'text-indigo-400',
      badgeBorder: 'border-indigo-500/30',
      isCustom: true,
      combineLabel: name,
      combineColumns: [],
      combineSeparator: '-',
    };
    setRules((prev) => [...prev, newRule]);
  };

  const handleDeleteCombination = (ruleId: string) => {
    setRules((prev) => prev.filter((r) => r.id !== ruleId).map((r, i) => ({ ...r, rank: i + 1 })));
  };

  const handleSaveCombinations = async () => {
    await handleSave();
    setShowCombine(false);
  };

  const handleDeleteRule = async (ruleId: string) => {
    const updated = rules.filter(r => r.id !== ruleId).map((r, i) => ({ ...r, rank: i + 1 }));
    setRules(updated);
    try {
      await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rules: updated }),
      });
      if (onRulesUpdated) {
        onRulesUpdated(updated);
      }
      setSyncFeedback('Rule removed from priority list.');
      setTimeout(() => setSyncFeedback(null), 3000);
    } catch (err: any) {
      setErrorMessage(`Failed to remove rule: ${err.message}`);
      setTimeout(() => setErrorMessage(null), 4000);
    }
  };

  const moveUp = (index: number) => {
    if (index === 0) return;
    const newRules = [...rules];
    const temp = newRules[index - 1];
    newRules[index - 1] = newRules[index];
    newRules[index] = temp;
    // re-index ranks
    const reindexed = newRules.map((r, i) => ({ ...r, rank: i + 1 }));
    setRules(reindexed);
  };

  const moveDown = (index: number) => {
    if (index === rules.length - 1) return;
    const newRules = [...rules];
    const temp = newRules[index + 1];
    newRules[index + 1] = newRules[index];
    newRules[index] = temp;
    // re-index ranks
    const reindexed = newRules.map((r, i) => ({ ...r, rank: i + 1 }));
    setRules(reindexed);
  };

  const moveToTop = (index: number) => {
    if (index === 0) return;
    const newRules = [...rules];
    const [item] = newRules.splice(index, 1);
    newRules.unshift(item);
    setRules(newRules.map((r, i) => ({ ...r, rank: i + 1 })));
  };

  const moveToBottom = (index: number) => {
    if (index === rules.length - 1) return;
    const newRules = [...rules];
    const [item] = newRules.splice(index, 1);
    newRules.push(item);
    setRules(newRules.map((r, i) => ({ ...r, rank: i + 1 })));
  };

  const handleNeetOrderChange = (order: 'desc' | 'asc') => {
    setNeetYearOrder(order);
    setRules(prev =>
      prev.map(r => {
        if (r.categoryKey === 'neet') {
          return {
            ...r,
            neetYearOrder: order,
            description:
              order === 'asc'
                ? 'NEET.xlsx (ordered by year 15, 19, 20, 22, 23... - Ascending)'
                : 'NEET.xlsx (ordered by year 23, 22, 20, 19, 15... - Recent years first)',
          };
        }
        return r;
      })
    );
  };

  const handleSave = async () => {
    setIsSaving(true);
    try {
      const res = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rules }),
      });
      const data = await res.json();
      if (data.rules) {
        setRules(data.rules);
        if (onRulesUpdated) {
          onRulesUpdated(data.rules);
        }
        setSaveSuccess(true);
        setTimeout(() => setSaveSuccess(false), 2500);
      }
    } catch (err: any) {
      setErrorMessage(`Failed to save rules: ${err.message}`);
      setTimeout(() => setErrorMessage(null), 4000);
    } finally {
      setIsSaving(false);
    }
  };

  const handleReset = async () => {
    setIsLoading(true);
    setErrorMessage(null);
    try {
      const res = await fetch(resetEndpoint, { method: 'POST' });
      const data = await res.json();
      if (data.rules) {
        setRules(data.rules);
        setNeetYearOrder('desc');
        if (onRulesUpdated) {
          onRulesUpdated(data.rules);
        }
        setSaveSuccess(true);
        setTimeout(() => setSaveSuccess(false), 2500);
      }
    } catch (err: any) {
      setErrorMessage(`Failed to reset: ${err.message}`);
      setTimeout(() => setErrorMessage(null), 4000);
    } finally {
      setIsLoading(false);
    }
  };

  if (!isOpen) return null;

  const headerTitle = title || (scope === 'mis' ? 'MIS Search Book Priority Configuration' : 'Directory Lead Priority Configuration');
  const badgeLabel = scope === 'mis' ? 'MIS Dedicated' : 'Global Default';
  const headerDesc = description || (scope === 'mis'
    ? 'Configure the priority order specifically for this MIS Search Book. When a lead matches multiple files, the top-ranked category is chosen.'
    : 'Change the priority rank of all files in the directory. Lower rank number = Higher Priority lead.');

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-3xl max-h-[90vh] flex flex-col shadow-2xl shadow-black/60 overflow-hidden">
        {/* Header */}
        <div className="px-6 py-4 border-b border-slate-800 flex items-center justify-between bg-slate-900/60">
          <div className="flex items-center gap-3">
            <div className={`w-10 h-10 rounded-xl ${scope === 'mis' ? 'bg-purple-500/10 border-purple-500/20 text-purple-400' : 'bg-emerald-500/10 border-emerald-500/20 text-emerald-400'} border flex items-center justify-center`}>
              <Sliders className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base font-bold text-white flex items-center gap-2">
                {headerTitle}
                <span className={`text-[11px] px-2 py-0.5 rounded-full ${scope === 'mis' ? 'bg-purple-500/20 text-purple-300' : 'bg-emerald-500/20 text-emerald-300'} font-mono`}>
                  {badgeLabel}
                </span>
              </h2>
              <p className="text-xs text-slate-400">
                {headerDesc}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => handleSyncDirectories(false)}
              disabled={isSyncing || isLoading}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-blue-600/20 hover:bg-blue-600/30 text-blue-300 border border-blue-500/40 hover:border-blue-500/60 transition-all cursor-pointer shadow-xs disabled:opacity-50"
              title="Refresh and sync all directories. Newly added directories are automatically placed at the last."
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin text-blue-400' : 'text-blue-300'}`} />
              <span>{isSyncing ? 'Syncing...' : 'Sync Directories'}</span>
            </button>
            <button
              type="button"
              onClick={openCombine}
              disabled={isLoading}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold bg-sky-600/20 hover:bg-sky-600/30 text-sky-300 border border-sky-500/40 hover:border-sky-500/60 transition-all cursor-pointer shadow-xs disabled:opacity-50"
              title="Choose columns (like Year, Lot, Month) to add to the Priority column"
            >
              <Link2 className="w-3.5 h-3.5" />
              <span>Combine Columns</span>
            </button>
            <button
              onClick={onClose}
              className="p-2 text-slate-400 hover:text-white rounded-lg hover:bg-slate-800 transition-colors cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Content */}
        <div className="p-6 overflow-y-auto space-y-4 flex-1">
          {/* Explanatory banner */}
          <div className="p-3.5 rounded-xl bg-slate-800/40 border border-slate-700/60 flex items-start gap-3 text-xs text-slate-300">
            <Info className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
            <div className="space-y-1">
              <p>
                <strong>Multi-Match Resolution Rule:</strong> If a phone number is present in multiple files (e.g., in <span className="text-amber-400 font-mono">Lot</span>, <span className="text-orange-400 font-mono">Login</span>, and <span className="text-cyan-400 font-mono">Fresh Data 21k</span>), the system assigns the highest priority category: <span className="text-cyan-400 font-bold font-mono">Fresh Data 21k</span> in the Priority column.
              </p>
              <p className="text-slate-400 text-[11px]">
                Use the <strong>Up ⬆️</strong> and <strong>Down ⬇️</strong> controls to adjust precedence.
              </p>
            </div>
          </div>

          {/* Directory Refresh & Auto-Sync Explanatory Card */}
          <div className="p-3.5 rounded-xl bg-blue-950/20 border border-blue-800/40 flex items-center justify-between gap-3 text-xs text-blue-200">
            <div className="flex items-start gap-2.5">
              <FolderTree className="w-4 h-4 text-blue-400 shrink-0 mt-0.5" />
              <div>
                <span className="font-bold text-blue-100 flex items-center gap-2">
                  Auto-Sync Directories &amp; Sources
                  {lastSyncedTime && (
                    <span className="text-[10px] text-blue-400/80 font-mono font-normal">
                      • Last synced: {lastSyncedTime}
                    </span>
                  )}
                </span>
                <p className="text-[11px] text-blue-300/80 mt-0.5">
                  Whenever you add a new directory or start the app, it automatically syncs and appends the new directory at the <strong>very last</strong> of the priority hierarchy.
                </p>
              </div>
            </div>
            <button
              type="button"
              onClick={() => handleSyncDirectories(false)}
              disabled={isSyncing || isLoading}
              className="px-2.5 py-1 rounded-lg bg-blue-600/30 hover:bg-blue-600/50 text-blue-200 border border-blue-500/40 text-[11px] font-bold shrink-0 transition-colors cursor-pointer flex items-center gap-1.5"
            >
              <RefreshCw className={`w-3 h-3 ${isSyncing ? 'animate-spin' : ''}`} />
              <span>Refresh Now</span>
            </button>
          </div>

          {syncFeedback && (
            <div className="p-3 rounded-xl bg-emerald-500/20 border border-emerald-500/40 text-emerald-200 text-xs flex items-center gap-2 animate-in fade-in">
              <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
              <span>{syncFeedback}</span>
            </div>
          )}

          {errorMessage && (
            <div className="p-3 rounded-xl bg-rose-500/20 border border-rose-500/40 text-rose-200 text-xs flex items-center gap-2">
              <AlertCircle className="w-4 h-4 text-rose-400 shrink-0" />
              <span>{errorMessage}</span>
            </div>
          )}

          {/* NEET Year Sorting Options */}
          <div className="p-3.5 rounded-xl bg-purple-950/20 border border-purple-800/40 flex items-center justify-between">
            <div className="flex items-center gap-2.5">
              <Calendar className="w-4 h-4 text-purple-400" />
              <div>
                <span className="text-xs font-bold text-purple-200">NEET Year Tie-Breaker Ordering:</span>
                <p className="text-[11px] text-purple-300/70">
                  Controls precedence when comparing different NEET batch years (20, 22, 23, etc.)
                </p>
              </div>
            </div>
            <div className="flex items-center bg-slate-900 border border-purple-700/40 rounded-lg p-0.5 text-xs">
              <button
                type="button"
                onClick={() => handleNeetOrderChange('desc')}
                className={`px-3 py-1 rounded-md text-[11px] font-medium transition-colors cursor-pointer ${
                  neetYearOrder === 'desc'
                    ? 'bg-purple-600 text-white font-bold'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                Recent Years First (23, 22, 20...)
              </button>
              <button
                type="button"
                onClick={() => handleNeetOrderChange('asc')}
                className={`px-3 py-1 rounded-md text-[11px] font-medium transition-colors cursor-pointer ${
                  neetYearOrder === 'asc'
                    ? 'bg-purple-600 text-white font-bold'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                Ascending (20, 22, 23...)
              </button>
            </div>
          </div>

          {/* Rules List */}
          <div className="space-y-2">
            <div className="flex items-center justify-between text-xs text-slate-400 px-2 font-mono">
              <span>LEAD CATEGORY &amp; FILE MATCH</span>
              <span>PRIORITY CONTROLS</span>
            </div>

            {isLoading ? (
              <div className="py-12 text-center text-slate-500 text-xs">
                Loading priority hierarchy...
              </div>
            ) : (
              rules.map((rule, idx) => (
                <div
                  key={rule.id}
                  className={`flex items-center justify-between p-3 rounded-xl border transition-all ${
                    idx === 0
                      ? 'bg-emerald-950/20 border-emerald-500/40 shadow-sm'
                      : 'bg-slate-800/50 border-slate-700/60 hover:border-slate-600'
                  }`}
                >
                  <div className="flex items-center gap-3.5">
                    {/* Rank Badge */}
                    <div
                      className={`w-7 h-7 rounded-lg flex items-center justify-center font-mono font-bold text-xs shrink-0 ${
                        idx === 0
                          ? 'bg-emerald-500 text-slate-950 shadow-md shadow-emerald-500/20'
                          : idx === 1
                          ? 'bg-purple-500 text-white'
                          : idx === 2
                          ? 'bg-cyan-500 text-slate-950'
                          : 'bg-slate-700 text-slate-200'
                      }`}
                    >
                      #{rule.rank}
                    </div>

                    <div>
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="font-bold text-slate-100 text-xs">{rule.name}</span>
                        <span
                          className={`text-[10px] px-2 py-0.5 rounded-full font-mono font-semibold border ${rule.badgeBg} ${rule.badgeText} ${rule.badgeBorder}`}
                        >
                          Rank #{rule.rank}
                        </span>

                        {/* Directory / Sheet / Workfile Type Tag */}
                        {(rule.id.startsWith('dir_') || rule.id.startsWith('mis_dir_') || rule.description?.startsWith('Directory:')) ? (
                          <span className="text-[10px] px-2 py-0.5 rounded-full font-medium bg-blue-500/20 text-blue-300 border border-blue-500/40 flex items-center gap-1">
                            <Folder className="w-2.5 h-2.5" />
                            Directory
                          </span>
                        ) : rule.id.startsWith('mis_sheet_') ? (
                          <span className="text-[10px] px-2 py-0.5 rounded-full font-medium bg-purple-500/20 text-purple-300 border border-purple-500/40 flex items-center gap-1">
                            <FileSpreadsheet className="w-2.5 h-2.5" />
                            MIS Sheet
                          </span>
                        ) : (
                          <span className="text-[10px] px-2 py-0.5 rounded-full font-medium bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 flex items-center gap-1">
                            <FileSpreadsheet className="w-2.5 h-2.5" />
                            Workfile
                          </span>
                        )}

                        {rule.categoryKey === 'neet' && (
                          <span className="text-[10px] px-2 py-0.5 rounded-full font-mono font-bold bg-purple-500/20 text-purple-300 border border-purple-500/40">
                            {neetYearOrder === 'asc' ? 'Ascending (20, 22, 23...)' : 'Recent First (23, 22, 20...)'}
                          </span>
                        )}
                        {(rule.categoryKey === 'f2_leads' || rule.categoryKey === 'file_jan' || rule.categoryKey === 'file_may' || rule.id === 'file_jan' || rule.id === 'file_may') && (
                          <span className="text-[10px] px-2 py-0.5 rounded-full font-mono font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/40">
                            P1 &gt; P2 &gt; P3 &gt; P4 &gt; P5
                          </span>
                        )}
                        {idx === 0 && (
                          <span className="text-[10px] font-bold text-emerald-400 bg-emerald-500/10 px-1.5 py-0.5 rounded border border-emerald-500/20">
                            Highest Priority
                          </span>
                        )}
                      </div>
                      <p className="text-[11px] text-slate-400 mt-0.5 flex items-center gap-2 flex-wrap">
                        {rule.categoryKey === 'neet' ? (
                          <>
                            <span className="text-purple-200 font-medium">
                              {neetYearOrder === 'asc'
                                ? 'NEET.xlsx (ordered by year 15, 19, 20, 22, 23... - Older batches first)'
                                : 'NEET.xlsx (ordered by year 23, 22, 20, 19, 15... - Recent batches first)'}
                            </span>
                            <span className="text-[10px] font-mono text-purple-300 bg-purple-950/80 px-2 py-0.5 rounded border border-purple-700/60 font-semibold">
                              {neetYearOrder === 'asc' ? 'Tie-Breaker: 20 < 22 < 23 (Older batches win)' : 'Tie-Breaker: 23 > 22 > 20 (Recent batches win)'}
                            </span>
                          </>
                        ) : (rule.categoryKey === 'f2_leads' || rule.categoryKey === 'file_jan' || rule.categoryKey === 'file_may' || rule.id === 'file_jan' || rule.id === 'file_may') ? (
                          <>
                            <span className="text-emerald-300 font-medium">
                              {rule.description || `${rule.name} workfile in DOCTORS / f2 Leads`}
                            </span>
                            <span className="text-[10px] font-mono text-emerald-300 bg-emerald-950/80 px-2 py-0.5 rounded border border-emerald-700/60 font-semibold">
                              Sub-order: P1 (1.01) &gt; P2 (1.02) &gt; P3 (1.03) &gt; P4 (1.04) &gt; P5 (1.05)
                            </span>
                          </>
                        ) : (
                          <span>{rule.description}</span>
                        )}
                      </p>
                    </div>
                  </div>

                  {/* Move Controls */}
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => moveToTop(idx)}
                      disabled={idx === 0}
                      title="Move to Top (#1 Highest Priority)"
                      className="p-1.5 rounded-lg bg-slate-900 border border-slate-700 text-slate-300 hover:text-emerald-400 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed transition-colors cursor-pointer"
                    >
                      <ChevronsUp className="w-4 h-4" />
                    </button>

                    <button
                      type="button"
                      onClick={() => moveUp(idx)}
                      disabled={idx === 0}
                      title="Increase priority rank"
                      className="p-1.5 rounded-lg bg-slate-900 border border-slate-700 text-slate-300 hover:text-white hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed transition-colors cursor-pointer"
                    >
                      <ArrowUp className="w-4 h-4" />
                    </button>

                    <button
                      type="button"
                      onClick={() => moveDown(idx)}
                      disabled={idx === rules.length - 1}
                      title="Decrease priority rank"
                      className="p-1.5 rounded-lg bg-slate-900 border border-slate-700 text-slate-300 hover:text-white hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed transition-colors cursor-pointer"
                    >
                      <ArrowDown className="w-4 h-4" />
                    </button>

                    <button
                      type="button"
                      onClick={() => moveToBottom(idx)}
                      disabled={idx === rules.length - 1}
                      title="Move to Bottom (Lowest Priority)"
                      className="p-1.5 rounded-lg bg-slate-900 border border-slate-700 text-slate-300 hover:text-rose-400 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed transition-colors cursor-pointer"
                    >
                      <ChevronsDown className="w-4 h-4" />
                    </button>

                    {!CORE_SYSTEM_RULE_IDS.has(rule.id) && (
                      <button
                        type="button"
                        onClick={() => handleDeleteRule(rule.id)}
                        title="Delete rule from priority list"
                        className="p-1.5 rounded-lg bg-rose-500/10 border border-rose-500/30 text-rose-400 hover:bg-rose-500/25 hover:text-rose-300 transition-colors cursor-pointer ml-1"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    )}
                  </div>
                </div>
              ))
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t border-slate-800 bg-slate-900/80 flex items-center justify-between">
          <button
            type="button"
            onClick={handleReset}
            disabled={isLoading || isSaving}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-slate-400 hover:text-white hover:bg-slate-800 border border-transparent hover:border-slate-700 transition-colors cursor-pointer"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Reset to Standard Default</span>
          </button>

          <div className="flex items-center gap-2.5">
            {saveSuccess && (
              <span className="inline-flex items-center gap-1 text-xs text-emerald-400 font-medium animate-in fade-in">
                <CheckCircle2 className="w-4 h-4" />
                <span>Saved &amp; Active</span>
              </span>
            )}

            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 rounded-xl text-xs font-medium text-slate-300 hover:bg-slate-800 transition-colors cursor-pointer"
            >
              Close
            </button>

            <button
              type="button"
              onClick={handleSave}
              disabled={isSaving || isLoading}
              className="inline-flex items-center gap-2 px-5 py-2 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-lg shadow-emerald-900/40 transition-all cursor-pointer disabled:opacity-50"
            >
              <Save className="w-4 h-4" />
              <span>{isSaving ? 'Saving...' : 'Apply & Save Priority'}</span>
            </button>
          </div>
        </div>
      </div>

      {showCombine && (
        <CombineColumnsPanel
          rules={rules}
          columnsByRule={columnsByRule}
          allColumns={allColumns}
          loading={columnsLoading}
          saving={isSaving}
          onChange={updateRuleFields}
          onClose={() => setShowCombine(false)}
          onSave={handleSaveCombinations}
          onAddRule={handleAddCombination}
          onDeleteRule={handleDeleteCombination}
        />
      )}
    </div>
  );
};
