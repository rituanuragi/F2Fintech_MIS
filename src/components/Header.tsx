import React from 'react';
import {
  RefreshCw,
  FolderInput,
  Loader2,
  Sliders,
  ExternalLink,
  Search,
  FolderTree,
  Sparkles,
  PanelLeftClose,
  PanelLeftOpen,
} from 'lucide-react';
import { ActiveSource, SyncProgress, AppMode } from '../types';
import { F2FintechLogo } from './F2FintechLogo';

interface HeaderProps {
  appMode: AppMode;
  onChangeAppMode: (mode: AppMode) => void;
  activeSource: ActiveSource | null;
  isSyncing: boolean;
  progress: SyncProgress;
  onSync: (full: boolean) => void;
  onChangeDirectory: () => void;
  onDisconnect: () => void;
  onOpenPriorityConfig?: () => void;
  sidebarOpen: boolean;
  onToggleSidebar: () => void;
}

const LinkOrLabel: React.FC<any> = ({ isLink, href, children, ...rest }) =>
  isLink
    ? <a href={href} target="_blank" rel="noopener noreferrer" {...rest}>{children}</a>
    : <div {...rest}>{children}</div>;

export const Header: React.FC<HeaderProps> = ({
  appMode,
  onChangeAppMode,
  activeSource,
  isSyncing,
  progress,
  onSync,
  onChangeDirectory,
  onDisconnect,
  onOpenPriorityConfig,
  sidebarOpen,
  onToggleSidebar,
}) => {
  return (
    <header className="bg-white border-b border-slate-200/90 text-slate-800 px-4 sm:px-6 py-2.5 flex flex-wrap items-center justify-between gap-3 sticky top-0 z-30 shadow-xs">
      {/* Left: Active Directory / Context Info */}
      <div className="flex items-center gap-3">
        <button
          type="button"
          onClick={onToggleSidebar}
          className="hidden md:inline-flex p-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 cursor-pointer transition-colors"
          title={sidebarOpen ? 'Hide tabs (full screen)' : 'Show tabs'}
          aria-label={sidebarOpen ? 'Hide tabs' : 'Show tabs'}
        >
          {sidebarOpen ? <PanelLeftClose className="w-4 h-4" /> : <PanelLeftOpen className="w-4 h-4" />}
        </button>
        <F2FintechLogo variant="icon" size="sm" className="hidden sm:inline-flex" />
        {appMode === 'directory' ? (
          activeSource ? (
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold text-slate-500 uppercase tracking-wider hidden lg:inline">
                Drive:
              </span>
              <LinkOrLabel
                isLink={true}
                href={activeSource.rootUrl}
                className="px-3 py-1.5 rounded-xl bg-blue-50/80 hover:bg-blue-100 text-blue-700 font-semibold text-xs flex items-center gap-2 border border-blue-200/80 transition-colors shadow-2xs"
                title={`Google Drive Root ID: ${activeSource.rootFolderId}`}
              >
                <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                <span className="truncate max-w-[140px] sm:max-w-xs">{activeSource.name}</span>
                {<ExternalLink className="w-3 h-3 text-blue-500 shrink-0" />}
              </LinkOrLabel>
            </div>
          ) : (
            <div className="flex items-center gap-2 text-xs font-semibold text-slate-400">
              <span className="w-2 h-2 rounded-full bg-amber-400"></span>
              <span>No Drive connected</span>
            </div>
          )
        ) : (
          <div className="flex items-center gap-2">
            <span className="px-2.5 py-1 rounded-xl bg-slate-900 text-white font-black text-xs tracking-tight flex items-center gap-1.5 shadow-2xs">
              <Sparkles className="w-3.5 h-3.5 text-amber-400" />
              <span>Customer Verification Suite</span>
            </span>
            <span className="text-xs text-slate-500 hidden sm:inline font-medium">
              Multi-ID Lookup, Field Diff & Google Sheets MIS
            </span>
          </div>
        )}
      </div>

      {/* Center: App Switcher Navigation Bar (Independent Switch) */}
      <div className="flex items-center bg-slate-100 p-1 rounded-xl border border-slate-200/80 shadow-2xs">
        <button
          type="button"
          onClick={() => onChangeAppMode('directory')}
          className={`flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-xs font-bold tracking-wide transition-all cursor-pointer ${
            appMode === 'directory'
              ? 'bg-white text-blue-700 shadow-xs ring-1 ring-slate-200'
              : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
          }`}
          title="Google Drive recursive directory crawler, workbook indexer & batch search"
        >
          <FolderTree className="w-3.5 h-3.5" />
          <span>Drive Directory Indexer</span>
        </button>

        <button
          type="button"
          onClick={() => onChangeAppMode('customer-lookup')}
          className={`flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-xs font-bold tracking-wide transition-all cursor-pointer ${
            appMode === 'customer-lookup'
              ? 'bg-white text-indigo-700 shadow-xs ring-1 ring-slate-200'
              : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
          }`}
          title="Customer PAN/Aadhaar/Phone search, base record diff comparison & Google Sheets MIS"
        >
          <Search className="w-3.5 h-3.5 text-indigo-600" />
          <span>Customer Lookup & MIS</span>
        </button>

        <button
          type="button"
          onClick={() => onChangeAppMode('f2-tickets')}
          className={`flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-xs font-bold tracking-wide transition-all cursor-pointer ${
            appMode === 'f2-tickets'
              ? 'bg-white text-emerald-700 shadow-xs ring-1 ring-slate-200'
              : 'text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
          }`}
          title="Live F2 Fintech All Tickets API Explorer (In-Memory, Zero Limit)"
        >
          <Sparkles className="w-3.5 h-3.5 text-emerald-600" />
          <span>F2 Live Tickets API</span>
        </button>
      </div>

      {/* Right: Actions & Sync Controls */}
      <div className="flex items-center gap-2">
        {isSyncing && (
          <div className="flex items-center gap-2 bg-blue-50/90 border border-blue-200 rounded-xl px-2.5 py-1 text-xs shadow-2xs">
            <Loader2 className="w-3.5 h-3.5 animate-spin text-blue-600 shrink-0" />
            <span className="font-semibold text-blue-700 truncate max-w-[160px] sm:max-w-[220px]">
              {progress.message || 'Syncing...'}
            </span>
          </div>
        )}

        {<button
          type="button"
          onClick={onOpenPriorityConfig}
          className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-xl bg-slate-800 hover:bg-slate-900 text-white transition-all cursor-pointer shadow-xs"
          title="Configure Lead Priority order"
        >
          <Sliders className="w-3.5 h-3.5 text-amber-400" />
          <span className="hidden md:inline">Priorities</span>
        </button>}

        {appMode === 'directory' ? (
          activeSource ? (
            <button
              onClick={() => onSync(false)}
              disabled={isSyncing}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-xl bg-blue-600 hover:bg-blue-700 text-white shadow-sm shadow-blue-500/20 disabled:opacity-50 transition-all cursor-pointer"
              title="Detect new or modified files"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin' : ''}`} />
              <span className="hidden sm:inline">Sync</span>
            </button>
          ) : (
            <button
              onClick={onChangeDirectory}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-bold rounded-xl bg-blue-600 hover:bg-blue-700 text-white shadow-md shadow-blue-500/20 transition-colors cursor-pointer"
            >
              <FolderInput className="w-3.5 h-3.5" />
              <span>Connect Drive</span>
            </button>
          )
        ) : null}
      </div>
    </header>
  );
};
