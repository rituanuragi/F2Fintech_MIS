import crypto from 'crypto';
import fs from 'fs';
import os from 'os';
import path from 'path';
import type { NextFunction, Request, Response } from 'express';
import { getDb, saveDb } from './db.js';

/**
 * Login, roles and tab access.
 *  - admin    : sees every tab, manages users, and chooses which tabs each employee can open.
 *  - operator : "Employee" on screen. Sees only the tabs the admin gave them.
 * Passwords are stored hashed (scrypt). The login token is signed and expires, so it survives a server restart.
 */

// Every tab an admin can give to an employee. 'users' (Users & Access) is admin only and never listed here.
export const GRANTABLE_TABS = [
  'dashboard', 'directory', 'files', 'sheets', 'search', 'batch-search', 'source-finder', 'duplicates',
  'history', 'errors', 'settings', 'customer-lookup', 'google-sheets-sync', 'mis', 'mis-dashboard',
  'mis-search-book', 'f2-tickets', 'priority-config',
] as const;

// What a NEW employee can open until the admin changes it.
export const DEFAULT_EMPLOYEE_TABS = ['dashboard', 'search', 'customer-lookup', 'mis-dashboard'];

export interface AuthUser {
  id: string;
  username: string;
  name: string;
  role: 'admin' | 'operator';
  status: string;
  allowedTabs: string[];
  mustChangePassword: boolean;
}

declare module 'express-serve-static-core' {
  interface Request { authUser?: AuthUser }
}

const TOKEN_HOURS = Number(process.env.F2_SESSION_HOURS || 12);

// ---------- secret ----------
function loadSecret(): string {
  if (process.env.F2_AUTH_SECRET) return process.env.F2_AUTH_SECRET;
  const dir = process.env.F2_CONFIG_DIR ? path.resolve(process.env.F2_CONFIG_DIR) : path.join(os.homedir(), '.f2-lead-intelligence');
  const file = path.join(dir, 'auth-secret.txt');
  try {
    if (fs.existsSync(file)) return fs.readFileSync(file, 'utf-8').trim();
    fs.mkdirSync(dir, { recursive: true });
    const s = crypto.randomBytes(48).toString('hex');
    fs.writeFileSync(file, s, { mode: 0o600 });
    return s;
  } catch {
    return crypto.randomBytes(48).toString('hex'); // sessions then end when the server restarts
  }
}
const SECRET = loadSecret();

// ---------- passwords ----------
export function hashPassword(plain: string): string {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(plain, salt, 64).toString('hex');
  return `scrypt$${salt}$${hash}`;
}

export function verifyPassword(plain: string, stored: string): boolean {
  if (!stored.startsWith('scrypt$')) return false;
  const [, salt, hash] = stored.split('$');
  const test = crypto.scryptSync(plain, salt, 64);
  const real = Buffer.from(hash, 'hex');
  return real.length === test.length && crypto.timingSafeEqual(real, test);
}

// ---------- tokens ----------
const b64 = (s: string) => Buffer.from(s).toString('base64url');
const sign = (body: string) => crypto.createHmac('sha256', SECRET).update(body).digest('base64url');

export function makeToken(userId: string): string {
  const body = b64(JSON.stringify({ u: userId, exp: Date.now() + TOKEN_HOURS * 3600 * 1000 }));
  return `${body}.${sign(body)}`;
}

function readToken(token: string): string | null {
  const [body, sig] = (token || '').split('.');
  if (!body || !sig) return null;
  const good = sign(body);
  if (sig.length !== good.length || !crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(good))) return null;
  try {
    const p = JSON.parse(Buffer.from(body, 'base64url').toString('utf-8'));
    return p.exp > Date.now() ? String(p.u) : null;
  } catch { return null; }
}

// ---------- users ----------
function parseTabs(raw: unknown, role: string): string[] {
  if (role === 'admin') return [...GRANTABLE_TABS, 'users'];
  try {
    const arr = raw ? JSON.parse(String(raw)) : null;
    if (Array.isArray(arr)) return arr.filter((t) => (GRANTABLE_TABS as readonly string[]).includes(t));
  } catch {}
  return [...DEFAULT_EMPLOYEE_TABS];
}

const COLS = 'id, username, name, role, status, allowed_tabs, must_change_password';

function rowToUser(r: any[]): AuthUser {
  const role = r[3] === 'admin' ? 'admin' : 'operator';
  return {
    id: r[0], username: r[1], name: r[2], role, status: r[4],
    allowedTabs: parseTabs(r[5], role), mustChangePassword: !!r[6],
  };
}

export async function getUserById(id: string): Promise<AuthUser | null> {
  const db = await getDb();
  const res = db.exec(`SELECT ${COLS} FROM users WHERE id = ?;`, [id]);
  return res.length && res[0].values.length ? rowToUser(res[0].values[0]) : null;
}

export async function listUsers(): Promise<AuthUser[]> {
  const db = await getDb();
  const res = db.exec(`SELECT ${COLS} FROM users ORDER BY created_at ASC;`);
  return (res[0]?.values || []).map(rowToUser);
}

/** Adds the new columns, hashes any old plain-text password, and makes sure an admin exists. Runs once at start. */
export async function ensureAuthReady(): Promise<void> {
  const db = await getDb();
  for (const ddl of [
    'ALTER TABLE users ADD COLUMN allowed_tabs TEXT;',
    'ALTER TABLE users ADD COLUMN must_change_password INTEGER DEFAULT 0;',
  ]) { try { db.run(ddl); } catch { /* column already there */ } }

  const res = db.exec('SELECT id, password FROM users;');
  for (const [id, pw] of res[0]?.values || []) {
    const stored = String(pw);
    if (stored.startsWith('scrypt$')) continue;
    // The old default password must be changed at first login.
    db.run('UPDATE users SET password = ?, must_change_password = ? WHERE id = ?;',
      [hashPassword(stored), stored === 'admin123' || stored === 'admin' ? 1 : 0, id]);
  }
  saveDb();
}

// ---------- middleware ----------
const OPEN = (req: Request) => req.method === 'POST' && req.path === '/auth/login';

/** Admin-only server routes (an employee gets 403 even if they call the URL by hand). */
function adminOnly(req: Request): boolean {
  const p = req.path;
  const write = req.method !== 'GET' && req.method !== 'HEAD';
  if (p.startsWith('/admin/')) return true;
  if (['/connect', '/disconnect', '/sync', '/records/clear-all'].includes(p)) return true;
  if (write && (p === '/priority-rules' || p === '/priority-rules/reset' || p === '/mis-priority-rules' || p === '/mis-priority-rules/reset')) return true;
  if (write && p === '/mis/sheet-priorities') return true;
  if (write && /^\/mis\/(upload-sheet|connect-cached-sheet|auto-sync-config|crawler-sync-all)$/.test(p)) return true;
  if (write && /^\/mis\/sheets(\/|$)/.test(p) && !/\/sync$/.test(p)) return true;
  if (p === '/mis-search-book/upload-workbook') return true;
  return false;
}

export async function authGate(req: Request, res: Response, next: NextFunction) {
  if (OPEN(req)) return next();
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : String(req.query.access_token || '');
  const userId = readToken(token);
  if (!userId) return res.status(401).json({ error: 'Please log in.' });
  try {
    const user = await getUserById(userId);
    if (!user || user.status !== 'active') return res.status(401).json({ error: 'Your account is not active.' });
    req.authUser = user;
    if (user.mustChangePassword && req.path !== '/auth/me' && req.path !== '/auth/change-password') {
      return res.status(403).json({ error: 'Please change your password first.', mustChangePassword: true });
    }
    if (user.role !== 'admin' && adminOnly(req)) {
      return res.status(403).json({ error: 'Only an admin can do this.' });
    }
    // Audit trail and remarks always carry the real signed-in name, never one sent by the browser.
    if (req.body && typeof req.body === 'object' && /^\/records\/[^/]+\/(check|remarks)$/.test(req.path)) {
      req.body.adminName = user.name;
      req.body.author = user.username;
      req.body.authorName = user.name;
    }
    next();
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Auth check failed' });
  }
}

export const requireAdmin = (req: Request, res: Response, next: NextFunction) =>
  req.authUser?.role === 'admin' ? next() : res.status(403).json({ error: 'Only an admin can do this.' });
