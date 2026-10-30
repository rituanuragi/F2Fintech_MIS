import React from 'react';
import { Construction } from 'lucide-react';

/** Temporary stand-in for a screen whose original file was not available. */
export const Placeholder: React.FC<{ name: string; note?: string }> = ({ name, note }) => (
  <div className="bg-white border border-dashed border-amber-300 rounded-2xl p-10 text-center text-slate-600 space-y-2">
    <Construction className="w-8 h-8 mx-auto text-amber-500" />
    <h3 className="text-base font-bold text-slate-900">{name} still needs to be rebuilt</h3>
    <p className="text-xs max-w-md mx-auto">{note || 'The original file was not part of the upload. See HANDOFF.md for what this screen must do.'}</p>
  </div>
);
