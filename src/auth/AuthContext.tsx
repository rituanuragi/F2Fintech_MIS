import React, { createContext, useCallback, useContext, useEffect, useMemo, useState } from 'react';
import { AuthUser } from '../types';

const TOKEN_KEY = 'f2_auth_token';

interface AuthValue {
  user: AuthUser | null;
  loading: boolean;
  isAdmin: boolean;
  can: (tab: string) => boolean;
  login: (username: string, password: string) => Promise<void>;
  logout: () => void;
  reload: () => Promise<void>;
}

const Ctx = createContext<AuthValue | null>(null);
export const useAuth = () => {
  const v = useContext(Ctx);
  if (!v) throw new Error('useAuth must be used inside AuthProvider');
  return v;
};

// Every request to /api carries the login token. A 401 (expired / switched-off account) sends the person back to login.
let patched = false;
function patchFetch(onUnauthorized: () => void) {
  if (patched) return;
  patched = true;
  const original = window.fetch.bind(window);
  window.fetch = async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    const isApi = url.startsWith('/api') || url.startsWith(`${location.origin}/api`);
    const token = localStorage.getItem(TOKEN_KEY);
    if (isApi && token) {
      const headers = new Headers(init?.headers || (typeof input !== 'string' && !(input instanceof URL) ? input.headers : undefined));
      if (!headers.has('Authorization')) headers.set('Authorization', `Bearer ${token}`);
      init = { ...init, headers };
    }
    const res = await original(input, init);
    if (isApi && res.status === 401 && !url.includes('/auth/login')) onUnauthorized();
    return res;
  };
}

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<AuthUser | null>(null);
  const [loading, setLoading] = useState(true);

  const logout = useCallback(() => {
    localStorage.removeItem(TOKEN_KEY);
    setUser(null);
  }, []);

  useEffect(() => { patchFetch(() => { localStorage.removeItem(TOKEN_KEY); setUser(null); }); }, []);

  const reload = useCallback(async () => {
    if (!localStorage.getItem(TOKEN_KEY)) { setUser(null); setLoading(false); return; }
    try {
      const res = await fetch('/api/auth/me');
      if (res.ok) setUser((await res.json()).user || null);
      else if (res.status === 401) logout();
    } catch { /* server not reachable: stay on the login screen */ }
    setLoading(false);
  }, [logout]);

  useEffect(() => { reload(); }, [reload]);

  const login = useCallback(async (username: string, password: string) => {
    const res = await fetch('/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Login failed');
    localStorage.setItem(TOKEN_KEY, data.token);
    setUser(data.user);
  }, []);

  const value = useMemo<AuthValue>(() => ({
    user, loading, login, logout, reload,
    isAdmin: user?.role === 'admin',
    can: (tab: string) => !!user && (user.role === 'admin' || (user.allowedTabs || []).includes(tab)),
  }), [user, loading, login, logout, reload]);

  return <Ctx.Provider value={value}>{children}</Ctx.Provider>;
};
