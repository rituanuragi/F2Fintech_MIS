import React, { useState } from 'react';
import { Lock, Loader2 } from 'lucide-react';
import { useAuth } from '../auth/AuthContext';
import { F2FintechLogo } from './F2FintechLogo';

const card = 'mx-auto w-full max-w-sm bg-white text-slate-800 rounded-2xl border border-slate-200 shadow-lg p-7 space-y-4';
const input = 'w-full px-3 py-2.5 rounded-xl border border-slate-400 bg-white text-slate-900 placeholder-slate-500 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500';
const btn = 'w-full py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold disabled:bg-slate-300 disabled:text-slate-600 flex items-center justify-center gap-2 cursor-pointer';

export const LoginPage: React.FC = () => {
  const { login } = useAuth();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true); setError('');
    try { await login(username, password); } catch (err: any) { setError(err.message); }
    setBusy(false);
  };

  return (
    <div className="min-h-screen w-full bg-[#f3f6fc] text-slate-800 flex items-center justify-center p-4">
      <form onSubmit={submit} className={card}>
        <div className="flex flex-col items-center gap-2">
          <F2FintechLogo variant="icon" size="sm" />
          <h1 className="text-lg font-black text-slate-900">F2 Lead Intelligence</h1>
          <p className="text-xs text-slate-500">Log in to continue</p>
        </div>
        <input className={input} placeholder="Username" autoFocus autoComplete="username" value={username} onChange={(e) => setUsername(e.target.value)} />
        <input className={input} type="password" placeholder="Password" autoComplete="current-password" value={password} onChange={(e) => setPassword(e.target.value)} />
        {error && <div className="text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-lg p-2">{error}</div>}
        <button className={btn} disabled={busy || !username || !password}>
          {busy ? <Loader2 className="w-4 h-4 animate-spin" /> : <Lock className="w-4 h-4" />} Log in
        </button>
      </form>
    </div>
  );
};

/** Shown right after login when an admin gave a temporary password. */
export const ChangePasswordPage: React.FC = () => {
  const { reload, logout, user } = useAuth();
  const [cur, setCur] = useState('');
  const [next, setNext] = useState('');
  const [again, setAgain] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (next !== again) { setError('The two new passwords are not the same.'); return; }
    setBusy(true); setError('');
    const res = await fetch('/api/auth/change-password', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ currentPassword: cur, newPassword: next }),
    });
    const data = await res.json().catch(() => ({}));
    setBusy(false);
    if (!res.ok) { setError(data.error || 'Could not change password'); return; }
    await reload();
  };

  return (
    <div className="min-h-screen w-full bg-[#f3f6fc] text-slate-800 flex items-center justify-center p-4">
      <form onSubmit={submit} className={card}>
        <h1 className="text-lg font-black text-slate-900">Choose a new password</h1>
        <p className="text-xs text-slate-500">Hello {user?.name}. Please replace your temporary password. Use at least 8 characters.</p>
        <input className={input} type="password" placeholder="Current (temporary) password" value={cur} onChange={(e) => setCur(e.target.value)} />
        <input className={input} type="password" placeholder="New password" value={next} onChange={(e) => setNext(e.target.value)} />
        <input className={input} type="password" placeholder="New password again" value={again} onChange={(e) => setAgain(e.target.value)} />
        {error && <div className="text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-lg p-2">{error}</div>}
        <button className={btn} disabled={busy || !cur || next.length < 8}>
          {busy ? <Loader2 className="w-4 h-4 animate-spin" /> : null} Save password
        </button>
        <button type="button" onClick={logout} className="w-full text-xs text-slate-600 hover:text-slate-800 cursor-pointer">Log out</button>
      </form>
    </div>
  );
};
