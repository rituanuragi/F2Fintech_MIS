import React, { useState } from 'react';
import { X, MessageSquare, Send, Trash2, ShieldAlert, Clock, User } from 'lucide-react';
import { CustomerLookupRecord, LeadRemarkItem } from '../types';

interface RemarksModalProps {
  record: CustomerLookupRecord | null;
  isOpen: boolean;
  onClose: () => void;
  onAddRemark: (recordId: number, text: string) => Promise<void>;
  onDeleteRemark: (recordId: number, remarkId: string) => Promise<void>;
  currentUser?: { name: string; role: string } | null;
}

export const RemarksModal: React.FC<RemarksModalProps> = ({
  record,
  isOpen,
  onClose,
  onAddRemark,
  onDeleteRemark,
  currentUser,
}) => {
  const [newRemarkText, setNewRemarkText] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);

  if (!isOpen || !record) return null;

  const remarks: LeadRemarkItem[] = record.remarks || [];

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newRemarkText.trim() || isSubmitting) return;

    try {
      setIsSubmitting(true);
      await onAddRemark(record.id, newRemarkText.trim());
      setNewRemarkText('');
    } catch (err) {
      console.error('Failed to post remark:', err);
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-hidden flex flex-col max-h-[85vh]">
        {/* Modal Header */}
        <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-2 rounded-xl bg-amber-500/20 text-amber-400 border border-amber-400/30">
              <MessageSquare className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-base font-bold">Internal Lead Remarks</h3>
                <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-500/20 text-amber-300 border border-amber-400/30 uppercase tracking-wider">
                  Internal Only
                </span>
              </div>
              <p className="text-xs text-slate-300">
                Operational notes for {record.customerName || 'Lead'} ({record.phoneNumber})
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

        {/* Existing Remarks List */}
        <div className="p-6 overflow-y-auto flex-1 space-y-3.5 bg-slate-50/50">
          {remarks.length === 0 ? (
            <div className="text-center py-8 text-slate-400">
              <ShieldAlert className="w-9 h-9 mx-auto text-slate-300 mb-2" />
              <p className="text-xs font-semibold">No operational remarks logged yet.</p>
              <p className="text-[11px] text-slate-400 mt-0.5">
                Leave a note below to track follow-ups, document checks, or agent comments.
              </p>
            </div>
          ) : (
            remarks.map((item) => {
              const date = new Date(item.createdAt);
              const formattedDate = !isNaN(date.getTime())
                ? date.toLocaleString('en-IN', {
                    dateStyle: 'short',
                    timeStyle: 'short',
                  })
                : item.createdAt;

              return (
                <div
                  key={item.id}
                  className="bg-white border border-slate-200/90 rounded-xl p-3.5 shadow-2xs space-y-2 group hover:border-slate-300 transition-colors"
                >
                  <div className="flex items-center justify-between text-xs">
                    <div className="flex items-center gap-1.5 font-bold text-slate-700">
                      <User className="w-3.5 h-3.5 text-blue-600" />
                      <span>{item.authorName || item.author || 'Admin'}</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] font-mono text-slate-400 flex items-center gap-1">
                        <Clock className="w-3 h-3 text-slate-300" />
                        {formattedDate}
                      </span>
                      <button
                        type="button"
                        onClick={() => onDeleteRemark(record.id, item.id)}
                        className="opacity-0 group-hover:opacity-100 p-1 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded transition-all cursor-pointer"
                        title="Delete this remark entry"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                  <p className="text-xs text-slate-800 leading-relaxed font-sans whitespace-pre-wrap">
                    {item.text}
                  </p>
                </div>
              );
            })
          )}
        </div>

        {/* New Remark Input Form */}
        <form onSubmit={handleSubmit} className="p-4 bg-white border-t border-slate-200 space-y-2">
          <div className="relative">
            <textarea
              rows={2}
              value={newRemarkText}
              onChange={(e) => setNewRemarkText(e.target.value)}
              placeholder="Add an internal remark or operational note to this lead..."
              className="w-full text-xs p-3 pr-12 rounded-xl border border-slate-200 focus:outline-hidden focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all resize-none"
            />
            <button
              type="submit"
              disabled={!newRemarkText.trim() || isSubmitting}
              className="absolute right-2.5 bottom-3.5 p-2 rounded-lg bg-blue-600 hover:bg-blue-700 text-white disabled:opacity-40 transition-colors cursor-pointer shadow-xs"
              title="Post Remark"
            >
              <Send className="w-3.5 h-3.5" />
            </button>
          </div>
          <div className="flex items-center justify-between text-[11px] text-slate-400">
            <span>Posted as: <strong className="text-slate-600">{currentUser?.name || 'Administrator'}</strong></span>
            <span>Visible to internal staff only</span>
          </div>
        </form>
      </div>
    </div>
  );
};
