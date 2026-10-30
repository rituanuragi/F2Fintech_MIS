import React from 'react';
import { DashboardStats, ActiveSource } from '../types';

interface Props {
  stats: DashboardStats | null;
  activeSource: ActiveSource | null;
  onNavigate: (tab: any) => void;
  onOpenConnect: () => void;
  onOpenPriorityConfig?: () => void;
}

export const DashboardView: React.FC<Props> = ({ stats, activeSource, onNavigate, onOpenConnect }) => {
  if (!activeSource || !stats) {
    return (
      <div className="bg-white border border-slate-200 rounded-2xl p-10 text-center space-y-3 text-slate-700">
        <h2 className="text-lg font-bold">No Drive directory connected yet</h2>
        <button onClick={onOpenConnect} className="px-4 py-2 rounded-xl bg-blue-600 text-white text-xs font-bold cursor-pointer">Connect Drive folder</button>
      </div>
    );
  }
  const cards: [string, number, string][] = [
    ['Files', stats.filesCount, 'files'], ['Sheets', stats.sheetsCount, 'sheets'],
    ['Indexed rows', stats.indexedRows, 'search'], ['Unique phones', stats.uniquePhones, 'search'],
    ['Duplicate groups', stats.duplicateGroups, 'duplicates'], ['Folders', stats.foldersCount, 'directory'],
  ];
  return (
    <div className="space-y-6 text-slate-800">
      <h2 className="text-xl font-bold">{activeSource.name}</h2>
      <div className="grid grid-cols-2 lg:grid-cols-3 gap-4">
        {cards.map(([label, value, tab]) => (
          <button key={label} onClick={() => onNavigate(tab)} className="bg-white border border-slate-200 rounded-2xl p-5 text-left hover:border-blue-400 cursor-pointer">
            <div className="text-[11px] uppercase font-bold text-slate-400">{label}</div>
            <div className="text-2xl font-black font-mono">{(value || 0).toLocaleString()}</div>
          </button>
        ))}
      </div>
    </div>
  );
};
