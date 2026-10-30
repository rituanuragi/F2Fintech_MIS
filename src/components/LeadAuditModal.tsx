import React from 'react';
import { X, History, UserCheck, Shield, Clock, FileSpreadsheet, Building, Calendar } from 'lucide-react';
import { CustomerLookupRecord, LeadCheckItem } from '../types';

interface LeadAuditModalProps {
  record: CustomerLookupRecord | null;
  isOpen: boolean;
  onClose: () => void;
}

// Show only the last 4 characters of a PAN (e.g. XXXXXX234F)
const maskPan = (pan: string) => {
  const p = String(pan || '').trim();
  return p.length <= 4 ? p : `${'X'.repeat(p.length - 4)}${p.slice(-4)}`;
};

export const LeadAuditModal: React.FC<LeadAuditModalProps> = ({ record, isOpen, onClose }) => {
  if (!isOpen || !record) return null;

  const history: LeadCheckItem[] = record.checkHistory || [];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh]">
        {/* Modal Header */}
        <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-2 rounded-xl bg-indigo-500/20 text-indigo-400 border border-indigo-400/30">
              <History className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-bold">Audit Trail & Google Sheet Remarks</h3>
              <p className="text-xs text-slate-300">
                Verification history & MIS remarks for <strong className="text-white">{record.customerName || record.doctorName || 'Customer'}</strong>
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-white/10 transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Lead Context Summary */}
        <div className="px-6 py-3 bg-slate-50 border-b border-slate-200/80 flex flex-wrap items-center justify-between gap-2 text-xs">
          <div className="flex items-center gap-3 flex-wrap">
            <div>
              <span className="text-slate-400 font-semibold uppercase text-[10px] block">PHONE</span>
              <span className="font-mono font-bold text-slate-900">{record.phoneNumber}</span>
            </div>
            {record.fileName && (
              <div>
                <span className="text-slate-400 font-semibold uppercase text-[10px] block">SOURCE FILE</span>
                <span className="font-mono font-bold text-slate-900">{record.fileName}</span>
              </div>
            )}
            {(record.sheetName || record.rawSheetName) && (
              <div>
                <span className="text-slate-400 font-semibold uppercase text-[10px] block">SHEET NAME</span>
                <span className="font-mono font-bold text-emerald-700">{record.sheetName || record.rawSheetName}</span>
              </div>
            )}
            {record.panNumber && record.panNumber !== '—' && (
              <div>
                <span className="text-slate-400 font-semibold uppercase text-[10px] block">PAN</span>
                <span className="font-mono font-bold text-slate-900">{maskPan(record.panNumber)}</span>
              </div>
            )}
            {record.lenderName && (
              <div>
                <span className="text-slate-400 font-semibold uppercase text-[10px] block">LENDER</span>
                <span className="font-bold text-slate-900">{record.lenderName}</span>
              </div>
            )}
            {record.loginDate && (
              <div>
                <span className="text-slate-400 font-semibold uppercase text-[10px] block">LOGIN DATE</span>
                <span className="font-mono font-bold text-slate-900">{record.loginDate}</span>
              </div>
            )}
            {record.approvalDate && (
              <div>
                <span className="text-slate-400 font-semibold uppercase text-[10px] block">APPROVED DATE</span>
                <span className="font-mono font-bold text-emerald-700">{record.approvalDate}</span>
              </div>
            )}
          </div>

          <div className="flex items-center gap-2">
            <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-indigo-50 text-indigo-700 border border-indigo-200">
              {record.checkCount || history.length || 1} Total Checks
            </span>
          </div>
        </div>

        {/* Modal Scrollable Body */}
        <div className="p-6 overflow-y-auto flex-1 space-y-5">
          {/* 1. SEPARATE OPS REMARKS SECTION */}
          <div className="space-y-1.5">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <div className="flex items-center gap-2">
                <span className="px-2 py-0.5 rounded-md text-[10px] font-extrabold uppercase tracking-wider bg-amber-100 text-amber-900 border border-amber-300">
                  OPS REMARKS
                </span>
                <span className="text-xs font-bold text-slate-700">Operations & Disbursal Trail</span>
              </div>
              <span className="text-[10px] font-mono font-semibold text-slate-500">
                File: {record.fileName || 'MIS'}
                {record.sheetName || record.rawSheetName ? ` • Sheet: ${record.sheetName || record.rawSheetName}` : ''}
              </span>
            </div>
            <div className="p-4 rounded-xl bg-amber-50/60 border border-amber-200/90 shadow-2xs">
              <p className="text-xs text-slate-800 leading-relaxed font-medium whitespace-pre-wrap">
                {record.opsRemarks || 'None recorded in MIS Google Sheet.'}
              </p>
            </div>
          </div>

          {/* 2. SEPARATE CREDIT REMARKS SECTION */}
          <div className="space-y-1.5">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <div className="flex items-center gap-2">
                <span className="px-2 py-0.5 rounded-md text-[10px] font-extrabold uppercase tracking-wider bg-purple-100 text-purple-900 border border-purple-300">
                  CREDIT REMARKS
                </span>
                <span className="text-xs font-bold text-slate-700">Underwriting & Eligibility Assessment</span>
              </div>
              <span className="text-[10px] font-mono font-semibold text-slate-500">
                File: {record.fileName || 'MIS'}
                {record.sheetName || record.rawSheetName ? ` • Sheet: ${record.sheetName || record.rawSheetName}` : ''}
              </span>
            </div>
            <div className="p-4 rounded-xl bg-purple-50/60 border border-purple-200/90 shadow-2xs">
              <p className="text-xs text-slate-800 leading-relaxed font-medium whitespace-pre-wrap">
                {record.creditRemarks || 'None recorded in MIS Google Sheet.'}
              </p>
            </div>
          </div>

          {/* 3. VERIFICATION AUDIT TRAIL TIMELINE */}
          <div className="space-y-2 pt-2 border-t border-slate-100">
            <div className="flex items-center justify-between">
              <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1.5">
                <Shield className="w-3.5 h-3.5 text-indigo-600" />
                <span>Verification Audit History</span>
              </h4>
              <span className="text-[10px] text-slate-400">Automated Audit Log</span>
            </div>

            {history.length === 0 ? (
              <div className="text-center py-6 text-slate-500 bg-slate-50 rounded-xl border border-slate-200/70">
                <UserCheck className="w-8 h-8 mx-auto text-slate-300 mb-1" />
                <p className="font-semibold text-xs text-slate-700">First Verification Recorded</p>
                <p className="text-[11px] text-slate-400 mt-0.5">This record was just logged as verified.</p>
              </div>
            ) : (
              <div className="relative pl-6 space-y-3 before:absolute before:left-2 before:top-2 before:bottom-2 before:w-0.5 before:bg-slate-200">
                {history.map((item, idx) => {
                  const date = new Date(item.timestamp);
                  const formattedDate = !isNaN(date.getTime())
                    ? date.toLocaleString('en-IN', {
                        dateStyle: 'medium',
                        timeStyle: 'medium',
                      })
                    : item.timestamp;

                  return (
                    <div key={item.id || idx} className="relative flex items-start gap-3 group">
                      <span className="absolute -left-6 top-1.5 w-3.5 h-3.5 rounded-full border-2 border-white bg-indigo-600 ring-4 ring-indigo-50 shrink-0" />
                      <div className="flex-1 bg-white border border-slate-200/90 rounded-xl p-3 shadow-2xs hover:border-indigo-300 transition-colors">
                        <div className="flex items-center justify-between text-xs mb-1">
                          <span className="font-bold text-slate-800 flex items-center gap-1.5">
                            <Shield className="w-3.5 h-3.5 text-indigo-600" />
                            {item.adminName || 'Admin'}
                          </span>
                          <span className="text-[10px] font-mono text-slate-400">
                            #{history.length - idx}
                          </span>
                        </div>
                        <div className="flex items-center gap-1.5 text-[11px] text-slate-500 font-mono">
                          <Clock className="w-3 h-3 text-slate-400" />
                          <span>{formattedDate}</span>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs transition-colors cursor-pointer"
          >
            Close Audit Trail
          </button>
        </div>
      </div>
    </div>
  );
};
