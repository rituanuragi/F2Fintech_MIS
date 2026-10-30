import React from 'react';
import { Folder } from 'lucide-react';
import { DiscoveredFolderItem, DiscoveredFileItem } from '../types';

interface Props {
  folders: DiscoveredFolderItem[];
  files: DiscoveredFileItem[];
  onPreviewSheet: (fileId: string, sheetName: string) => void;
  onManageSheets?: (fileName: string) => void;
}

/** Simple folder list with file counts. Replace with the old tree view when available. */
export const DirectoryView: React.FC<Props> = ({ folders, files }) => {
  const counts = new Map<string, number>();
  files.forEach((f) => counts.set(f.folderPath, (counts.get(f.folderPath) || 0) + 1));
  const paths = Array.from(new Set([...folders.map((f) => f.path), ...counts.keys()])).sort();
  return (
    <div className="bg-white border border-slate-200 rounded-2xl p-5 text-slate-800">
      <h2 className="text-lg font-bold mb-3">Directory ({paths.length} folders)</h2>
      <ul className="divide-y divide-slate-100 text-xs font-mono">
        {paths.map((p) => (
          <li key={p} className="py-2 flex items-center gap-2">
            <Folder className="w-4 h-4 text-emerald-600" /> <span className="flex-1">{p}</span>
            <span className="text-slate-400">{counts.get(p) || 0} files</span>
          </li>
        ))}
      </ul>
    </div>
  );
};
