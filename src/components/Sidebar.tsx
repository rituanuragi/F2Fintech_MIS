import React from 'react';
import {
  LayoutDashboard, FolderTree, FileSpreadsheet, Layers, Search, ListChecks, Compass, Copy,
  History, AlertTriangle, Settings, UserSearch, Table2, FileText, BarChart3, BookOpen, Ticket, Sliders,
} from 'lucide-react';

export type TabType =
  | 'dashboard' | 'directory' | 'files' | 'sheets' | 'search' | 'batch-search' | 'source-finder'
  | 'duplicates' | 'history' | 'errors' | 'settings' | 'customer-lookup' | 'google-sheets-sync'
  | 'mis' | 'mis-dashboard' | 'mis-search-book' | 'f2-tickets';

interface Props {
  currentTab: TabType;
  onSelectTab: (tab: TabType) => void;
  onOpenPriorityConfig?: () => void;
  stats?: { totalRows: number; filesCount: number; sheetsCount: number; duplicateGroups: number; errorsCount: number } | null;
}

const GROUPS: { title: string; items: { id: TabType; label: string; icon: any }[] }[] = [
  { title: 'Drive Directory', items: [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'directory', label: 'Directory', icon: FolderTree },
    { id: 'files', label: 'Files', icon: FileSpreadsheet },
    { id: 'sheets', label: 'Sheets', icon: Layers },
    { id: 'search', label: 'Global Search', icon: Search },
    { id: 'batch-search', label: 'Batch Excel Search', icon: ListChecks },
    { id: 'source-finder', label: 'Source Finder', icon: Compass },
    { id: 'duplicates', label: 'Duplicates', icon: Copy },
    { id: 'history', label: 'Sync History', icon: History },
    { id: 'errors', label: 'Errors', icon: AlertTriangle },
    { id: 'settings', label: 'Settings', icon: Settings },
  ] },
  { title: 'Customer & MIS', items: [
    { id: 'customer-lookup', label: 'Customer Lookup', icon: UserSearch },
    { id: 'google-sheets-sync', label: 'Google Sheets MIS', icon: Table2 },
    { id: 'mis', label: 'MIS Analyzer', icon: FileText },
    { id: 'mis-dashboard', label: 'MIS Dashboard', icon: BarChart3 },
    { id: 'mis-search-book', label: 'MIS Search Book', icon: BookOpen },
    { id: 'f2-tickets', label: 'F2 Live Tickets', icon: Ticket },
  ] },
];

export const TAB_LABELS: Record<string, string> = {
  ...Object.fromEntries(GROUPS.flatMap((g) => g.items.map((i) => [i.id, i.label]))),
  'priority-config': 'Lead Priorities',
};

export const Sidebar: React.FC<Props> = ({ currentTab, onSelectTab, onOpenPriorityConfig, stats }) => {
  const groups = GROUPS;
  return (
  <aside className="w-60 shrink-0 bg-slate-900 text-slate-200 border-r border-slate-800 overflow-y-auto hidden md:block">
    <div className="p-4 space-y-5">
      {groups.map((g) => (
        <div key={g.title}>
          <div className="text-[10px] uppercase tracking-widest text-slate-500 font-bold mb-2">{g.title}</div>
          <div className="space-y-0.5">
            {g.items.map(({ id, label, icon: Icon }) => (
              <button
                key={id}
                onClick={() => onSelectTab(id)}
                className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-xs font-semibold text-left transition-colors cursor-pointer ${
                  currentTab === id ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 text-slate-300'
                }`}
              >
                <Icon className="w-4 h-4 shrink-0" />
                <span className="truncate">{label}</span>
              </button>
            ))}
          </div>
        </div>
      ))}
      {onOpenPriorityConfig && (
        <button onClick={onOpenPriorityConfig} className="w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-xs font-semibold bg-slate-800 hover:bg-slate-700 cursor-pointer">
          <Sliders className="w-4 h-4 text-amber-400" /> Lead Priorities
        </button>
      )}
      {stats && (
        <div className="text-[11px] text-slate-400 font-mono space-y-0.5 pt-2 border-t border-slate-800">
          <div>{stats.filesCount} files · {stats.sheetsCount} sheets</div>
          <div>{stats.totalRows.toLocaleString()} rows</div>
        </div>
      )}
    </div>
  </aside>
  );
};
