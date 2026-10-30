import React, { useState, useEffect } from 'react';
import { Header } from './components/Header';
import { Sidebar, TabType } from './components/Sidebar';
import { DashboardView } from './components/DashboardView';
import { DirectoryView } from './components/DirectoryView';
import { FilesView } from './components/FilesView';
import { SheetsView } from './components/SheetsView';
import { GlobalSearchView } from './components/GlobalSearchView';
import { BatchExcelSearch } from './components/BatchExcelSearch';
import { MisView } from './components/MisView';
import { DuplicatesView } from './components/DuplicatesView';
import { SourceFinderView } from './components/SourceFinderView';
import { HistoryView } from './components/HistoryView';
import { SettingsView } from './components/SettingsView';
import { FullRowModal } from './components/FullRowModal';
import { SheetPreviewModal } from './components/SheetPreviewModal';
import { ConnectModal } from './components/ConnectModal';
import { PriorityConfigModal } from './components/PriorityConfigModal';
import { CustomerLookupView } from './components/CustomerLookupView';
import { GoogleSheetsSyncView } from './components/GoogleSheetsSyncView';
import { F2TicketsApiView } from './components/F2TicketsApiView';
import { MisSearchBookView } from './components/MisSearchBookView';
import { MisDashboardView } from './components/MisDashboardView';
import {
  ActiveSource,
  SyncProgress,
  DashboardStats,
  DiscoveredFolderItem,
  DiscoveredFileItem,
  AppMode,
} from './types';

export default function App() {
  const [appMode, setAppMode] = useState<AppMode>('directory');
  const [currentTab, setCurrentTab] = useState<TabType>('dashboard');
  const [activeSource, setActiveSource] = useState<ActiveSource | null>(null);
  const [isSyncing, setIsSyncing] = useState(false);
  const [progress, setProgress] = useState<SyncProgress>({
    stage: 'completed',
    message: 'System ready',
  });
  const [stats, setStats] = useState<DashboardStats | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [versionMismatch, setVersionMismatch] = useState(false);
  const [dataVersion, setDataVersion] = useState(0);
  const wasSyncingRef = React.useRef(false);
  const hasAutoSyncedRef = React.useRef(false);
  const lastSyncErrorRef = React.useRef<string | null>(null);

  const [folders, setFolders] = useState<DiscoveredFolderItem[]>([]);
  const [files, setFiles] = useState<DiscoveredFileItem[]>([]);

  // Modals & Cross-View Navigation state
  const [activeRecordId, setActiveRecordId] = useState<number | string | null>(null);
  const [previewTarget, setPreviewTarget] = useState<{ fileId: string; sheetName: string } | null>(null);
  const [isConnectOpen, setIsConnectOpen] = useState(false);
  const [isPriorityModalOpen, setIsPriorityModalOpen] = useState(false);
  const [initialCustomerQuery, setInitialCustomerQuery] = useState<string>('');
  const [initialSheetsFilter, setInitialSheetsFilter] = useState<string>('');
  const [initialTicketId, setInitialTicketId] = useState<string | null>(null);
  // Sidebar show/hide (full-screen mode). Remembered between visits.
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(() => {
    try { return localStorage.getItem('f2_sidebar_open') !== '0'; } catch { return true; }
  });
  const toggleSidebar = () => setSidebarOpen((v) => {
    const next = !v;
    try { localStorage.setItem('f2_sidebar_open', next ? '1' : '0'); } catch {}
    return next;
  });

  const handleNavigateToTicket = (ticketId: string | number) => {
    setInitialTicketId(String(ticketId));
    setAppMode('f2-tickets');
    setCurrentTab('f2-tickets');
  };

  const handleBackToCustomerLookup = () => {
    setAppMode('customer-lookup');
    setCurrentTab('customer-lookup');
  };

  // Safe JSON fetcher that avoids SyntaxError on HTML responses during server restarts
  const safeFetch = async <T,>(url: string): Promise<T | null> => {
    try {
      const res = await fetch(url);
      if (!res.ok) return null;
      const contentType = res.headers.get('content-type');
      if (!contentType || !contentType.includes('application/json')) return null;
      return (await res.json()) as T;
    } catch {
      return null;
    }
  };

  const prevMisSyncVersionRef = React.useRef<number | null>(null);

  // Poll status & sync state
  const fetchStatus = async () => {
    const data = await safeFetch<{
      isConnected: boolean;
      activeSource: ActiveSource | null;
      isSyncing: boolean;
      progress: SyncProgress;
      misAutoSync?: {
        syncVersion?: number;
      };
    }>('/api/status');
    if (data) {
      setVersionMismatch((data as any).version !== __APP_VERSION__);
      if (data.activeSource) {
        setActiveSource(data.activeSource);
      }
      setIsSyncing(data.isSyncing || false);
      if (data.progress) {
        if (data.progress.stage === 'error') {
          // Show the real reason once instead of hiding it behind "System ready".
          if (lastSyncErrorRef.current !== data.progress.message) {
            lastSyncErrorRef.current = data.progress.message;
            setNotice(`Drive sync failed: ${data.progress.message}`);
          }
          setProgress({ stage: 'completed', message: 'System ready' });
        } else {
          setProgress(data.progress);
        }
      }
      const syncVer = data.misAutoSync?.syncVersion;
      if (typeof syncVer === 'number') {
        if (prevMisSyncVersionRef.current !== null && syncVer > prevMisSyncVersionRef.current) {
          prevMisSyncVersionRef.current = syncVer;
          fetchStats();
          fetchDirectory();
        } else if (prevMisSyncVersionRef.current === null) {
          prevMisSyncVersionRef.current = syncVer;
        }
      }
    }
  };

  // Fetch Dashboard Stats
  const fetchStats = async () => {
    const data = await safeFetch<{
      hasConnectedSource: boolean;
      activeSource?: ActiveSource;
      stats: DashboardStats | null;
    }>('/api/dashboard');
    if (data) {
      if (data.hasConnectedSource && data.stats) {
        setStats({ ...data.stats, errorsCount: 0 });
        if (data.activeSource) {
          setActiveSource(data.activeSource);
        }
      }
    }
  };

  // Fetch Directory & Files
  const fetchDirectory = async () => {
    const data = await safeFetch<{
      folders: DiscoveredFolderItem[];
      files: DiscoveredFileItem[];
    }>('/api/directory');
    if (data) {
      setFolders(data.folders || []);
      setFiles(data.files || []);
    }
  };

  const handleSync = async (full: boolean) => {
    try {
      setIsSyncing(true);
      setProgress({
        stage: 'discovering',
        message: full ? 'Starting full resync...' : 'Syncing Drive directory & MIS sheets...',
      });
      const res = await fetch('/api/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fullResync: full }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        setIsSyncing(false);
        setNotice((err && err.error) || `Sync could not start (HTTP ${res.status})`);
        return;
      }
      await fetchStatus();
    } catch (err) {
      console.error('Sync trigger failed:', err);
    }
  };

  // Automatically refresh and sync both priority sections on app start
  useEffect(() => {
    if (!hasAutoSyncedRef.current) {
      hasAutoSyncedRef.current = true;
      // Refresh priority rules and check permanent sheets without heavy re-indexing
      fetch('/api/priority-rules/sync', { method: 'POST' }).catch((err) => {
        console.warn('Auto-sync priority rules failed on app startup:', err);
      });
      fetch('/api/mis-priority-rules/sync', { method: 'POST' }).catch((err) => {
        console.warn('Auto-sync MIS priority rules failed on app startup:', err);
      });
    }
  }, []);

  useEffect(() => {
    if (wasSyncingRef.current && !isSyncing) {
      fetchStats();
      fetchDirectory();
      if (['dashboard','directory','files','sheets','duplicates','history','errors'].includes(currentTab)) {
        setDataVersion((v) => v + 1); // reload the open screen with fresh data
      }
    }
    wasSyncingRef.current = isSyncing;
  }, [isSyncing, currentTab]);

  // Light background refresh so numbers stay current without clicking anything
  useEffect(() => {
    const t = setInterval(() => { fetchStats(); }, 15000);
    return () => clearInterval(t);
  }, []);

  useEffect(() => {
    fetchStatus();
    fetchStats();
    fetchDirectory();

    const interval = setInterval(() => {
      fetchStatus();
      if (isSyncing) {
        fetchStats();
        fetchDirectory();
      }
    }, 2000);

    return () => clearInterval(interval);
  }, [isSyncing]);

  const handleConnect = async (rootUrl: string, name?: string) => {
    const res = await fetch('/api/connect', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rootUrl, name }),
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || 'Connection failed');
    }
    setIsSyncing(true);
    await fetchStatus();
    await fetchStats();
    await fetchDirectory();
  };

  const handleDisconnect = async () => {
    try {
      await fetch('/api/disconnect', { method: 'POST' });
      setActiveSource(null);
      setStats(null);
      setFolders([]);
      setFiles([]);
    } catch (err) {
      console.error('Disconnect failed:', err);
    }
  };

  const handleSelectTab = (tab: TabType) => {
    setCurrentTab(tab);
    if (tab === 'customer-lookup' || tab === 'google-sheets-sync') {
      setAppMode('customer-lookup');
    } else if (tab === 'f2-tickets') {
      setAppMode('f2-tickets');
    } else {
      setAppMode('directory');
    }
  };

  const handleChangeAppMode = (mode: AppMode) => {
    setAppMode(mode);
    if (mode === 'customer-lookup') {
      setCurrentTab('customer-lookup');
    } else if (mode === 'f2-tickets') {
      setCurrentTab('f2-tickets');
    } else {
      if (['customer-lookup', 'google-sheets-sync', 'f2-tickets'].includes(currentTab)) {
        setCurrentTab('dashboard');
      }
    }
  };

  const handleLookupCustomer = (phoneOrQuery: string) => {
    setInitialCustomerQuery(phoneOrQuery);
    setAppMode('customer-lookup');
    setCurrentTab('customer-lookup');
  };

  const handleManageSheets = (workbookName: string) => {
    setInitialSheetsFilter(workbookName);
    setAppMode('directory');
    setCurrentTab('sheets');
  };

  return (
    <div className="min-h-screen bg-[#f3f6fc] text-slate-800 flex flex-col font-sans">
      {/* Top Header with App Switcher */}
      <Header
        appMode={appMode}
        onChangeAppMode={handleChangeAppMode}
        activeSource={activeSource}
        isSyncing={isSyncing}
        progress={progress}
        onSync={handleSync}
        onChangeDirectory={() => setIsConnectOpen(true)}
        onDisconnect={handleDisconnect}
        onOpenPriorityConfig={() => setIsPriorityModalOpen(true)}
        sidebarOpen={sidebarOpen}
        onToggleSidebar={toggleSidebar}
      />

      <div className="flex flex-1">
        {/* Navigation Sidebar */}
        {sidebarOpen && <Sidebar
          currentTab={currentTab}
          onSelectTab={(tab) => {
            if (tab !== 'sheets') setInitialSheetsFilter('');
            handleSelectTab(tab);
          }}
          onOpenPriorityConfig={() => setIsPriorityModalOpen(true)}
          stats={
            stats
              ? {
                  totalRows: stats.totalRows,
                  filesCount: stats.filesCount,
                  sheetsCount: stats.sheetsCount,
                  duplicateGroups: stats.duplicateGroups,
                  errorsCount: stats.errorsCount,
                }
              : null
          }
        />}

        {/* Main Content Area */}
        <main className={`flex-1 p-6 lg:p-8 mx-auto w-full overflow-y-auto min-w-0 ${sidebarOpen ? 'max-w-7xl' : 'max-w-none'}`}>
          {versionMismatch && (
            <div className="mb-4 rounded-xl border border-red-300 bg-red-50 p-4 text-sm text-red-800">
              <b>An old server is still running.</b> This page is new, but the server answering it is an older copy, so you would see old numbers.
              Close every terminal window, then run <code className="font-mono">npm run dev</code> again (it now closes old servers by itself).
            </div>
          )}
          {notice && (
            <div className="mb-4 flex items-start justify-between gap-3 rounded-xl border border-amber-300 bg-amber-50 p-3 text-xs text-amber-900">
              <span>{notice} See the <b>Errors</b> tab for details. The app will retry automatically.</span>
              <button onClick={() => setNotice(null)} className="font-bold cursor-pointer">Dismiss</button>
            </div>
          )}

          {/* Shubhi Project Integrated Features */}
          <div className={currentTab === 'f2-tickets' ? 'block' : 'hidden'}>
            <F2TicketsApiView
              initialTicketId={initialTicketId}
              onBackToLookup={handleBackToCustomerLookup}
              onClearInitialTicketId={() => setInitialTicketId(null)}
            />
          </div>

          <div className={currentTab === 'customer-lookup' ? 'block' : 'hidden'}>
            <CustomerLookupView
              initialQuery={initialCustomerQuery}
              onOpenRowModal={(id) => setActiveRecordId(id)}
              onNavigateToMIS={() => handleSelectTab('google-sheets-sync')}
              onNavigateToTicket={handleNavigateToTicket}
              onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
            />
          </div>

          {currentTab === 'google-sheets-sync' && (
            <GoogleSheetsSyncView
              onNavigateToLookup={() => handleSelectTab('customer-lookup')}
            />
          )}

          <div key={dataVersion}>
          {/* Directory & Deep Indexer Features */}
          {currentTab === 'dashboard' && (
            <DashboardView
              stats={stats}
              activeSource={activeSource}
              onNavigate={handleSelectTab}
              onOpenConnect={() => setIsConnectOpen(true)}
              onOpenPriorityConfig={() => setIsPriorityModalOpen(true)}
            />
          )}

          {currentTab === 'directory' && (
            <DirectoryView
              folders={folders}
              files={files}
              onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
              onManageSheets={handleManageSheets}
            />
          )}

          {currentTab === 'files' && (
            <FilesView
              files={files}
              onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
              onManageSheets={handleManageSheets}
            />
          )}

          {currentTab === 'sheets' && (
            <SheetsView
              initialSearchTerm={initialSheetsFilter}
              onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
            />
          )}

          {currentTab === 'mis-dashboard' && (
            <MisDashboardView
              onViewRecord={(id) => setActiveRecordId(id)}
              onLookupCustomer={handleLookupCustomer}
              onNavigateToTicket={handleNavigateToTicket}
            />
          )}

          {currentTab === 'search' && (
            <GlobalSearchView
              onViewRecord={(id) => setActiveRecordId(id)}
              onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
              onLookupCustomer={handleLookupCustomer}
            />
          )}

          {currentTab === 'mis-search-book' && (
            <MisSearchBookView
              onViewRecord={(id) => setActiveRecordId(id)}
              onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
              onLookupCustomer={handleLookupCustomer}
            />
          )}

          {currentTab === 'batch-search' && (
            <BatchExcelSearch
              onViewRecord={(id) => setActiveRecordId(id)}
              onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
              onLookupCustomer={handleLookupCustomer}
            />
          )}

          {currentTab === 'mis' && (
            <MisView
              onViewRecord={(id) => setActiveRecordId(id)}
              onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
              onLookupCustomer={handleLookupCustomer}
            />
          )}

          {currentTab === 'source-finder' && (
            <SourceFinderView
              onViewRecord={(id) => setActiveRecordId(id)}
              onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
              onLookupCustomer={handleLookupCustomer}
            />
          )}

          {currentTab === 'duplicates' && (
            <DuplicatesView
              onViewRecord={(id) => setActiveRecordId(id)}
              onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
              onLookupCustomer={handleLookupCustomer}
            />
          )}

          {currentTab === 'history' && (
            <HistoryView
              initialSection="runs"
              onTriggerSync={() => handleSync(false)}
            />
          )}

          {currentTab === 'errors' && (
            <HistoryView
              initialSection="errors"
              onTriggerSync={() => handleSync(false)}
            />
          )}

          {currentTab === 'settings' && (
            <SettingsView
              activeSource={activeSource}
              onOpenConnect={() => setIsConnectOpen(true)}
              onSync={handleSync}
              onDisconnect={handleDisconnect}
              onOpenPriorityConfig={() => setIsPriorityModalOpen(true)}
            />
          )}
          </div>
        </main>
      </div>

      {/* Full Record Modal */}
      {activeRecordId !== null && (
        <FullRowModal
          recordId={activeRecordId}
          onClose={() => setActiveRecordId(null)}
          onPreviewSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
          onLookupCustomer={handleLookupCustomer}
          onNavigateToTicket={handleNavigateToTicket}
        />
      )}

      {/* Sheet Grid Preview Modal */}
      {previewTarget !== null && (
        <SheetPreviewModal
          fileId={previewTarget.fileId}
          sheetName={previewTarget.sheetName}
          onClose={() => setPreviewTarget(null)}
          onViewRecord={(id) => setActiveRecordId(id)}
          onSwitchSheet={(fileId, sheetName) => setPreviewTarget({ fileId, sheetName })}
        />
      )}

      {/* Connect / Change Directory Modal */}
      <ConnectModal
        isOpen={isConnectOpen}
        onClose={() => setIsConnectOpen(false)}
        onConnect={handleConnect}
        currentRootUrl={activeSource?.rootUrl}
      />

      {/* Global Priority Configuration Modal */}
      <PriorityConfigModal
        isOpen={isPriorityModalOpen}
        onClose={() => setIsPriorityModalOpen(false)}
      />
    </div>
  );
}
