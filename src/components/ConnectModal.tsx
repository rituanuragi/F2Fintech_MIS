import React, { useState } from 'react';
import { X, FolderInput, Loader2 } from 'lucide-react';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  onConnect: (rootUrl: string, name?: string) => Promise<void>;
  currentRootUrl?: string;
}

export const ConnectModal: React.FC<Props> = ({ isOpen, onClose, onConnect, currentRootUrl }) => {
  const [url, setUrl] = useState(currentRootUrl || '');
  const [name, setName] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  if (!isOpen) return null;

  const submit = async () => {
    setBusy(true); setError(null);
    try { await onConnect(url.trim(), name.trim() || undefined); onClose(); }
    catch (e: any) { setError(e?.message || 'Connection failed'); }
    finally { setBusy(false); }
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl w-full max-w-lg p-6 space-y-4 text-slate-800">
        <div className="flex items-center justify-between">
          <h3 className="font-bold flex items-center gap-2"><FolderInput className="w-4 h-4 text-blue-600" /> Connect a Google Drive folder</h3>
          <button onClick={onClose} className="p-1 rounded hover:bg-slate-100 cursor-pointer"><X className="w-4 h-4" /></button>
        </div>
        <p className="text-xs text-slate-500">Paste the folder link. It must be shared as "Anyone with the link can view".</p>
        <input value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://drive.google.com/drive/folders/..."
          className="w-full border border-slate-300 rounded-xl px-3 py-2 text-xs" />
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Name (optional)"
          className="w-full border border-slate-300 rounded-xl px-3 py-2 text-xs" />
        {error && <div className="text-xs text-rose-600 font-semibold">{error}</div>}
        <div className="flex justify-end gap-2">
          <button onClick={onClose} className="px-4 py-2 text-xs rounded-xl bg-slate-100 cursor-pointer">Cancel</button>
          <button onClick={submit} disabled={!url.trim() || busy}
            className="px-4 py-2 text-xs font-bold rounded-xl bg-blue-600 text-white disabled:opacity-50 inline-flex items-center gap-2 cursor-pointer">
            {busy && <Loader2 className="w-3.5 h-3.5 animate-spin" />} Connect &amp; index
          </button>
        </div>
      </div>
    </div>
  );
};
