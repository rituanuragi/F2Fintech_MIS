import React, { useCallback, useEffect, useState } from 'react';
import { UserPlus, KeyRound, Trash2, Power, Loader2, Pencil } from 'lucide-react';
import { AuthUser } from '../types';
import { useAuth } from '../auth/AuthContext';
import { TAB_LABELS } from './Sidebar';

async function call(url: string, method = 'GET', body?: unknown) {
  const res = await fetch(url, {
    method,
    headers: body ? { 'Content-Type': 'application/json' } : undefined,
    body: body ? JSON.stringify(body) : undefined,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `Request failed (${res.status})`);
  return data;
}

export const UsersAccessView: React.FC = () => {
  const { user: me } = useAuth();
  const [users, setUsers] = useState<AuthUser[]>([]);
  const [grantable, setGrantable] = useState<string[]>([]);
  const [defaults, setDefaults] = useState<string[]>([]);
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);
  const [loading, setLoading] = useState(true);
  const [form, setForm] = useState({ name: '', username: '', password: '', role: 'operator' });

  const load = useCallback(async () => {
    try {
      const d = await call('/api/admin/users');
      setUsers(d.users); setGrantable(d.grantableTabs); setDefaults(d.defaultEmployeeTabs);
    } catch (e: any) { setMsg({ ok: false, text: e.message }); }
    setLoading(false);
  }, []);
  useEffect(() => { load(); }, [load]);

  const run = async (fn: () => Promise<any>, okText?: string) => {
    try { const d = await fn(); setMsg({ ok: true, text: okText || d?.message || 'Saved' }); await load(); }
    catch (e: any) { setMsg({ ok: false, text: e.message }); }
  };

  const toggleTab = (u: AuthUser, tab: string) => {
    const has = (u.allowedTabs || []).includes(tab);
    const tabs = has ? u.allowedTabs!.filter((t) => t !== tab) : [...(u.allowedTabs || []), tab];
    run(() => call(`/api/admin/users/${u.id}/tabs`, 'PUT', { tabs }), 'Access updated');
  };

  const field = 'px-3 py-2 rounded-lg border border-slate-400 bg-white text-slate-900 placeholder-slate-500 text-sm';
  if (loading) return <div className="p-6"><Loader2 className="w-5 h-5 animate-spin" /></div>;

  return (
    <div className="space-y-6 text-slate-800">
      <div>
        <h2 className="text-xl font-black text-slate-900">Users & Access</h2>
        <p className="text-sm text-slate-500">Admins see every tab. For each employee, tick the tabs they may open. Changes apply the next time that person's screen refreshes.</p>
      </div>

      {msg && <div className={`text-xs rounded-lg p-2 border ${msg.ok ? 'bg-emerald-50 border-emerald-200 text-emerald-800' : 'bg-rose-50 border-rose-200 text-rose-800'}`}>{msg.text}</div>}

      <form
        className="bg-white border border-slate-200 rounded-2xl p-4 flex flex-wrap gap-2 items-end"
        onSubmit={(e) => {
          e.preventDefault();
          run(async () => { const d = await call('/api/admin/users', 'POST', { ...form, allowedTabs: defaults }); setForm({ name: '', username: '', password: '', role: 'operator' }); return d; });
        }}
      >
        <input className={field} placeholder="Full name" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
        <input className={field} placeholder="Username" value={form.username} onChange={(e) => setForm({ ...form, username: e.target.value })} />
        <input className={field} type="text" placeholder="Temporary password (8+)" value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} />
        <select className={field} value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>
          <option value="operator">Employee</option>
          <option value="admin">Admin</option>
        </select>
        <button className="px-4 py-2 rounded-lg bg-blue-600 text-white text-sm font-bold flex items-center gap-2 cursor-pointer disabled:opacity-50"
          disabled={!form.name || !form.username || form.password.length < 8}>
          <UserPlus className="w-4 h-4" /> Add user
        </button>
      </form>

      {users.map((u) => (
        <div key={u.id} className="bg-white border border-slate-200 rounded-2xl p-4 space-y-3">
          <div className="flex flex-wrap items-center gap-3 justify-between">
            <div>
              <div className="font-bold text-slate-900">{u.name} <span className="text-xs font-mono text-slate-400">@{u.username}</span></div>
              <div className="text-xs text-slate-500">
                {u.role === 'admin' ? 'Admin' : 'Employee'} · {u.status === 'active' ? 'Active' : 'Switched off'}
                {u.mustChangePassword ? ' · must change password' : ''}
              </div>
            </div>
            <div className="flex flex-wrap gap-2 text-xs font-bold">
              <button className="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 cursor-pointer flex items-center gap-1"
                onClick={() => {
                  const username = window.prompt('Username (no spaces):', u.username);
                  if (username === null) return;
                  const name = window.prompt('Full name:', u.name);
                  if (name === null) return;
                  run(() => call(`/api/admin/users/${u.id}/profile`, 'PUT', { username, name }));
                }}>
                <Pencil className="w-3.5 h-3.5" /> Edit name / username
              </button>
              <button className="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 cursor-pointer"
                onClick={() => run(() => call(`/api/admin/users/${u.id}/role`, 'PUT', { role: u.role === 'admin' ? 'operator' : 'admin' }))}>
                Make {u.role === 'admin' ? 'Employee' : 'Admin'}
              </button>
              <button className="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 cursor-pointer flex items-center gap-1"
                onClick={() => run(() => call(`/api/admin/users/${u.id}/status`, 'PUT', { status: u.status === 'active' ? 'inactive' : 'active' }))}>
                <Power className="w-3.5 h-3.5" /> {u.status === 'active' ? 'Switch off' : 'Switch on'}
              </button>
              <button className="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 cursor-pointer flex items-center gap-1"
                onClick={() => {
                  const pw = window.prompt(`New temporary password for ${u.name} (8+ characters):`);
                  if (pw) run(() => call(`/api/admin/users/${u.id}/reset-password`, 'POST', { newPassword: pw }));
                }}>
                <KeyRound className="w-3.5 h-3.5" /> Reset password
              </button>
              {u.id !== me?.id && (
                <button className="px-2.5 py-1.5 rounded-lg bg-rose-50 text-rose-700 hover:bg-rose-100 cursor-pointer flex items-center gap-1"
                  onClick={() => { if (window.confirm(`Delete ${u.name}? This cannot be undone.`)) run(() => call(`/api/admin/users/${u.id}`, 'DELETE')); }}>
                  <Trash2 className="w-3.5 h-3.5" /> Delete
                </button>
              )}
            </div>
          </div>

          {u.role === 'admin' ? (
            <div className="text-xs text-slate-500">Admins can open every tab.</div>
          ) : (
            <div className="flex flex-wrap gap-1.5">
              {grantable.map((t) => {
                const on = (u.allowedTabs || []).includes(t);
                return (
                  <button key={t} onClick={() => toggleTab(u, t)}
                    className={`px-2.5 py-1 rounded-full text-[11px] font-semibold border cursor-pointer ${on ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-500 border-slate-300 hover:border-slate-400'}`}>
                    {on ? '✓ ' : '+ '}{TAB_LABELS[t] || t}
                  </button>
                );
              })}
            </div>
          )}
        </div>
      ))}
    </div>
  );
};
