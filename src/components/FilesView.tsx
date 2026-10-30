import React, { useState } from 'react';
import {
  FileSpreadsheet,
  Search,
  ExternalLink,
  Eye,
  CheckCircle2,
  AlertTriangle,
  Layers,
  Building2,
  Stethoscope,
  CreditCard,
} from 'lucide-react';
import { DiscoveredFileItem } from '../types';

interface FilesViewProps {
  files: DiscoveredFileItem[];
  onPreviewSheet: (fileId: string, sheetName: string) => void;
  onManageSheets?: (fileName: string) => void;
}

export const FilesView: React.FC<FilesViewProps> = ({ files, onPreviewSheet, onManageSheets }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<'ALL' | 'BL' | 'DOCTORS' | 'PL'>('ALL');

  const formatBytes = (bytes: number) => {
    if (!bytes) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${parseFloat((bytes / Math.pow(k, i)).toFixed(2))} ${sizes[i]}`;
  };

  const getFileCategory = (f: DiscoveredFileItem): 'BL' | 'DOCTORS' | 'PL' => {
    const fPath = (f.folderPath || '').toLowerCase();
    const fName = (f.name || '').toLowerCase();
    if (fPath === 'bl' || fName === 'bl' || fName === 'bl.xlsx') return 'BL';
    if (fPath === 'pl' || fName.includes('icici')) return 'PL';
    return 'DOCTORS';
  };

  const cleanBaseName = (n: string) =>
    (n || '').replace(/\.(xlsx|xls|csv)$/i, '').replace(/^[a-zA-Z0-9_-]{25,60}_/, '').trim().toLowerCase();

  const deduplicatedFiles = React.useMemo(() => {
    const seen = new Set<string>();
    const result: DiscoveredFileItem[] = [];
    const sorted = [...files].sort((a, b) => {
      const aIsXlsx = a.name.toLowerCase().endsWith('.xlsx') ? 1 : 0;
      const bIsXlsx = b.name.toLowerCase().endsWith('.xlsx') ? 1 : 0;
      return aIsXlsx - bIsXlsx;
    });

    for (const f of sorted) {
      const key = cleanBaseName(f.name);
      if (seen.has(key)) continue;
      seen.add(key);
      result.push({
        ...f,
        name: f.name.replace(/\.(xlsx|xls|csv)$/i, ''),
      });
    }
    return result;
  }, [files]);

  const blCount = deduplicatedFiles.filter(f => getFileCategory(f) === 'BL').length;
  const doctorsCount = deduplicatedFiles.filter(f => getFileCategory(f) === 'DOCTORS').length;
  const plCount = deduplicatedFiles.filter(f => getFileCategory(f) === 'PL').length;

  const filteredFiles = deduplicatedFiles.filter(f => {
    const matchesSearch =
      f.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      f.folderPath.toLowerCase().includes(searchTerm.toLowerCase());
    if (!matchesSearch) return false;

    if (selectedCategory === 'ALL') return true;
    return getFileCategory(f) === selectedCategory;
  });

  return (
    <div className="space-y-6">
      {/* Header & Search */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200 pb-4">
        <div>
          <h2 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
            <FileSpreadsheet className="w-5 h-5 text-blue-600" />
            <span>Files &amp; Excel Workbooks</span>
          </h2>
          <p className="text-xs text-slate-500 mt-1 font-medium">
            Complete inventory of all discovered workbooks across <strong>BL</strong>, <strong>DOCTORS</strong>, and <strong>PL</strong> directories.
          </p>
        </div>

        <div className="flex items-center gap-2">
          <div className="relative w-full sm:w-64">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="Search files..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-9 pr-3 py-2 bg-white border border-slate-300 rounded-xl text-xs text-slate-800 placeholder:text-slate-400 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-colors shadow-2xs"
            />
          </div>
        </div>
      </div>

      {/* Directory Filter Pills */}
      <div className="flex items-center gap-2 overflow-x-auto pb-1">
        <button
          onClick={() => setSelectedCategory('ALL')}
          className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
            selectedCategory === 'ALL'
              ? 'bg-slate-900 text-white shadow-xs'
              : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-50'
          }`}
        >
          <span>All Directories</span>
          <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-white/20">{deduplicatedFiles.length}</span>
        </button>

        <button
          onClick={() => setSelectedCategory('BL')}
          className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
            selectedCategory === 'BL'
              ? 'bg-emerald-600 text-white shadow-xs'
              : 'bg-white border border-emerald-200 text-emerald-700 hover:bg-emerald-50'
          }`}
        >
          <Building2 className="w-3.5 h-3.5" />
          <span>BL Directory</span>
          <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-emerald-100 text-emerald-800 font-mono">
            {blCount}
          </span>
        </button>

        <button
          onClick={() => setSelectedCategory('DOCTORS')}
          className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
            selectedCategory === 'DOCTORS'
              ? 'bg-blue-600 text-white shadow-xs'
              : 'bg-white border border-blue-200 text-blue-700 hover:bg-blue-50'
          }`}
        >
          <Stethoscope className="w-3.5 h-3.5" />
          <span>DOCTORS Directory</span>
          <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-blue-100 text-blue-800 font-mono">
            {doctorsCount}
          </span>
        </button>

        <button
          onClick={() => setSelectedCategory('PL')}
          className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
            selectedCategory === 'PL'
              ? 'bg-purple-600 text-white shadow-xs'
              : 'bg-white border border-purple-200 text-purple-700 hover:bg-purple-50'
          }`}
        >
          <CreditCard className="w-3.5 h-3.5" />
          <span>PL Directory</span>
          <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-purple-100 text-purple-800 font-mono">
            {plCount}
          </span>
        </button>
      </div>

      {/* Files Table */}
      <div className="bg-white border border-slate-200/90 rounded-2xl overflow-hidden shadow-xs">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-50 text-slate-600 uppercase font-sans text-[11px] font-bold tracking-wider border-b border-slate-200">
              <tr>
                <th className="py-3 px-4">Directory &amp; Path</th>
                <th className="py-3 px-4">Workbook Name</th>
                <th className="py-3 px-4">Sheets</th>
                <th className="py-3 px-4 text-right">Total Rows</th>
                <th className="py-3 px-4 text-right">Size</th>
                <th className="py-3 px-4 text-center">Status</th>
                <th className="py-3 px-4 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {filteredFiles.length === 0 ? (
                <tr>
                  <td colSpan={7} className="text-center py-10 text-slate-400 font-medium">
                    No files found matching current filter
                  </td>
                </tr>
              ) : (
                filteredFiles.map((file) => {
                  const cat = getFileCategory(file);
                  let catBadge = 'bg-slate-100 text-slate-700 border-slate-200';
                  if (cat === 'BL') catBadge = 'bg-emerald-50 text-emerald-700 border-emerald-200';
                  if (cat === 'DOCTORS') catBadge = 'bg-blue-50 text-blue-700 border-blue-200';
                  if (cat === 'PL') catBadge = 'bg-purple-50 text-purple-700 border-purple-200';

                  return (
                    <tr key={file.id} className="hover:bg-blue-50/40 transition-colors group">
                      <td className="py-3.5 px-4 font-mono text-[11px]">
                        <span className={`px-2 py-0.5 rounded-md border font-semibold ${catBadge}`}>
                          {file.folderPath}
                        </span>
                      </td>
                      <td className="py-3.5 px-4 font-bold text-slate-900 flex items-center gap-2">
                        <FileSpreadsheet className="w-4 h-4 text-blue-600 shrink-0" />
                        <span className="truncate max-w-xs">{file.name}</span>
                      </td>
                      <td className="py-3.5 px-4 font-mono">
                        {onManageSheets ? (
                          <button
                            type="button"
                            onClick={() => onManageSheets(file.name)}
                            className="inline-flex items-center gap-1 text-slate-700 hover:text-purple-700 bg-slate-100 hover:bg-purple-50 px-2 py-0.5 rounded-md text-[11px] font-semibold border border-slate-200 hover:border-purple-200 transition-colors cursor-pointer"
                            title="Filter worksheets for this workbook in Sheets Manager"
                          >
                            <Layers className="w-3 h-3 text-purple-600" />
                            <span>{file.sheets?.length || file.sheetsCount || 1} sheet(s)</span>
                          </button>
                        ) : (
                          <span className="inline-flex items-center gap-1 text-slate-700 bg-slate-100 px-2 py-0.5 rounded-md text-[11px] font-semibold border border-slate-200">
                            <Layers className="w-3 h-3 text-purple-600" />
                            {file.sheets?.length || file.sheetsCount || 1} sheet(s)
                          </span>
                        )}
                      </td>
                      <td className="py-3.5 px-4 text-right font-mono font-bold text-emerald-600">
                        {file.totalRows.toLocaleString()}
                      </td>
                      <td className="py-3.5 px-4 text-right font-mono text-slate-500">
                        {formatBytes(file.size)}
                      </td>
                      <td className="py-3.5 px-4 text-center">
                        {(() => {
                          const st = (file.status || '').toLowerCase();
                          if (!st || ['indexed', 'ready', 'completed', 'synced', 'success', 'active'].includes(st)) {
                            return (
                              <span className="inline-flex items-center gap-1 text-emerald-600 font-semibold text-[11px]">
                                <CheckCircle2 className="w-3.5 h-3.5" /> Ready
                              </span>
                            );
                          }
                          const bad = ['error', 'failed', 'corrupt', 'unreadable'].some((w) => st.includes(w));
                          return (
                            <span
                              className={`inline-flex items-center gap-1 font-semibold text-[11px] ${bad ? 'text-rose-600' : 'text-amber-600'}`}
                              title={file.status}
                            >
                              <AlertTriangle className="w-3.5 h-3.5" /> {file.status}
                            </span>
                          );
                        })()}
                      </td>
                      <td className="py-3.5 px-4 text-right">
                        <div className="flex items-center justify-end gap-1.5">
                          <button
                            onClick={() => onPreviewSheet(file.id, file.sheets?.[0]?.name || 'Sheet1')}
                            className="p-1.5 rounded-lg bg-slate-100 hover:bg-emerald-50 text-slate-600 hover:text-emerald-600 transition-colors cursor-pointer"
                            title="Preview first sheet"
                          >
                            <Eye className="w-4 h-4" />
                          </button>
                          <a
                            href={`https://drive.google.com/file/d/${file.id}/view`}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="p-1.5 rounded-lg bg-slate-100 hover:bg-blue-50 text-slate-600 hover:text-blue-600 transition-colors"
                            title="Open in Google Drive"
                          >
                            <ExternalLink className="w-4 h-4" />
                          </a>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
