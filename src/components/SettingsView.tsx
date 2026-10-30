import React from 'react';
import {
  Settings,
  FolderInput,
  FolderSync,
  LogOut,
  ExternalLink,
  CheckCircle2,
  HardDrive,
  Sparkles,
  ShieldCheck,
  Download,
  Terminal,
  Laptop,
} from 'lucide-react';
import { ActiveSource } from '../types';
import { F2FintechLogo } from './F2FintechLogo';

interface SettingsViewProps {
  activeSource: ActiveSource | null;
  onOpenConnect: () => void;
  onSync: (full: boolean) => void;
  onDisconnect: () => void;
  onOpenPriorityConfig?: () => void;
}

export const SettingsView: React.FC<SettingsViewProps> = ({
  activeSource,
  onOpenConnect,
  onSync,
  onDisconnect,
  onOpenPriorityConfig,
}) => {
  return (
    <div className="space-y-6 max-w-4xl">
      {/* Header */}
      <div className="border-b border-slate-200/90 pb-4">
        <h2 className="text-xl font-bold text-slate-900 tracking-tight flex items-center gap-2">
          <Settings className="w-5 h-5 text-blue-600" />
          Directory Configuration &amp; Sources
        </h2>
        <p className="text-xs text-slate-500 mt-1">
          Manage connected Google Drive root repositories, active search index isolation, and organization branding.
        </p>
      </div>

      {/* Brand & Logo Status Card */}
      <div className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-4">
        <div className="flex items-center justify-between border-b border-slate-100 pb-4">
          <div className="flex items-center gap-3">
            <F2FintechLogo variant="icon" size="md" />
            <div>
              <span className="text-[10px] font-mono text-blue-600 uppercase tracking-widest block font-bold">
                ORGANIZATION BRANDING
              </span>
              <h3 className="text-base font-bold text-slate-900">
                F2 Fintech Official Logo &amp; Identity
              </h3>
            </div>
          </div>
          <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-blue-50 text-blue-700 border border-blue-200 text-xs font-semibold">
            <ShieldCheck className="w-3.5 h-3.5 text-blue-600" /> Embedded Automatically
          </span>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200/80 flex flex-col items-center justify-center text-center">
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-3">
              Vertical Lockup (Sidebar &amp; Mobile)
            </span>
            <F2FintechLogo variant="full" size="lg" subtitleText="MIS &amp; Lead Search" />
          </div>

          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200/80 flex flex-col items-center justify-center text-center">
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-3">
              Horizontal Lockup (Headers &amp; Modals)
            </span>
            <div className="py-4">
              <F2FintechLogo variant="horizontal" size="md" subtitleText="Data &amp; Lead Intelligence" />
            </div>
          </div>
        </div>

        <div className="p-3.5 rounded-xl bg-emerald-50 border border-emerald-200 text-xs text-emerald-800 flex items-start gap-2.5">
          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
          <div>
            <strong className="font-semibold block text-emerald-900">No Google Drive Upload Required:</strong>
            The official logo you provided has been integrated directly into the web application's vector rendering engine. It renders crystal-clear at every screen resolution and DPI without needing image hosting in Google Drive.
          </div>
        </div>
      </div>

      {/* Active Source Card */}
      <div className="bg-white border border-slate-200/90 rounded-2xl p-6 shadow-xs space-y-5">
        <div className="flex items-center justify-between border-b border-slate-100 pb-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-blue-50 border border-blue-200 flex items-center justify-center text-blue-600 font-bold">
              <HardDrive className="w-5 h-5" />
            </div>
            <div>
              <span className="text-[10px] font-mono text-blue-600 uppercase tracking-widest block font-bold">
                ACTIVE REPOSITORY
              </span>
              <h3 className="text-base font-bold text-slate-900">
                {activeSource ? activeSource.name : 'No Active Source Connected'}
              </h3>
            </div>
          </div>

          {activeSource && (
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 text-xs font-semibold">
              <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500" /> Active &amp; Online
            </span>
          )}
        </div>

        {activeSource && (
          <div className="space-y-3 font-mono text-xs">
            <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-1">
              <span className="text-slate-500 text-[11px] block uppercase font-bold">Google Drive Root Folder URL</span>
              <a
                href={activeSource.rootUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="text-blue-600 hover:text-blue-700 hover:underline flex items-center gap-1.5 break-all text-xs font-medium"
              >
                <span>{activeSource.rootUrl}</span>
                <ExternalLink className="w-3.5 h-3.5 shrink-0" />
              </a>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                <span className="text-slate-500 text-[11px] block uppercase font-bold">Folder ID</span>
                <span className="text-slate-800 text-xs">{activeSource.rootFolderId}</span>
              </div>
              <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                <span className="text-slate-500 text-[11px] block uppercase font-bold">Internal Source ID</span>
                <span className="text-slate-800 text-xs">{activeSource.id}</span>
              </div>
            </div>
          </div>
        )}

        {/* Action buttons */}
        <div className="pt-2 flex flex-wrap items-center gap-3">
          <button
            onClick={onOpenConnect}
            className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold shadow-md shadow-blue-500/20 transition-all cursor-pointer"
          >
            <FolderInput className="w-4 h-4" />
            <span>Connect Another / Change Directory</span>
          </button>

          <button
            onClick={() => onSync(true)}
            className="inline-flex items-center gap-2 px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 rounded-xl text-xs font-semibold transition-all cursor-pointer"
          >
            <FolderSync className="w-4 h-4 text-slate-500" />
            <span>Rebuild Complete Index (Full Resync)</span>
          </button>

          {onOpenPriorityConfig && (
            <button
              onClick={onOpenPriorityConfig}
              className="inline-flex items-center gap-2 px-4 py-2 bg-slate-800 hover:bg-slate-900 text-white rounded-xl text-xs font-bold transition-all cursor-pointer shadow-2xs"
            >
              <Sparkles className="w-4 h-4 text-amber-400" />
              <span>Lead Priority Rules (#1–#9)</span>
            </button>
          )}

          <button
            onClick={onDisconnect}
            className="inline-flex items-center gap-2 px-4 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-xl text-xs font-semibold transition-all cursor-pointer ml-auto"
          >
            <LogOut className="w-4 h-4" />
            <span>Disconnect Source</span>
          </button>
        </div>
      </div>
    </div>
  );
};
