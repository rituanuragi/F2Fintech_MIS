// Closes any old copy of this app that is still holding the ports (3000 server, 5173 web page).
import { execSync } from 'child_process';
import fs from 'fs';

const ports = process.argv.slice(2).map(Number).filter(Boolean);
const isWin = process.platform === 'win32';
const run = (cmd) => { try { return execSync(cmd, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }); } catch { return ''; } };

export function parseNetstat(text, port) {
  const pids = new Set();
  for (const line of text.split(/\r?\n/)) {
    const cols = line.trim().split(/\s+/);
    // Proto  Local  Foreign  State  PID
    if (cols.length >= 5 && /LISTENING/i.test(cols[3]) && cols[1].endsWith(':' + port)) pids.add(cols[4]);
  }
  return pids;
}

function procPids(port) {
  // Linux without lsof/ss: read /proc directly
  const pids = new Set();
  const hex = port.toString(16).toUpperCase().padStart(4, '0');
  const inodes = new Set();
  for (const f of ['/proc/net/tcp', '/proc/net/tcp6']) {
    try {
      for (const l of fs.readFileSync(f, 'utf8').split('\n').slice(1)) {
        const c = l.trim().split(/\s+/);
        if (c[3] === '0A' && c[1]?.endsWith(':' + hex)) inodes.add(c[9]);
      }
    } catch {}
  }
  if (!inodes.size) return pids;
  for (const pid of fs.readdirSync('/proc').filter((d) => /^\d+$/.test(d))) {
    try {
      for (const fd of fs.readdirSync(`/proc/${pid}/fd`)) {
        const t = fs.readlinkSync(`/proc/${pid}/fd/${fd}`);
        const m = t.match(/^socket:\[(\d+)\]$/);
        if (m && inodes.has(m[1])) pids.add(pid);
      }
    } catch {}
  }
  return pids;
}

function pidsOnPort(port) {
  const pids = new Set();
  if (isWin) {
    for (const p of parseNetstat(run('netstat -ano'), port)) pids.add(p);
    if (!pids.size) {
      // Fallback: PowerShell sees both IPv4 and IPv6 listeners
      const out = run(`powershell -NoProfile -Command "(Get-NetTCPConnection -LocalPort ${port} -State Listen -ErrorAction SilentlyContinue).OwningProcess"`);
      for (const p of out.split(/\s+/)) if (/^\d+$/.test(p)) pids.add(p);
    }
  } else {
    for (const p of run(`lsof -t -iTCP:${port} -sTCP:LISTEN`).split(/\s+/)) if (p) pids.add(p);
    if (!pids.size) for (const m of run(`ss -ltnp 'sport = :${port}'`).matchAll(/pid=(\d+)/g)) pids.add(m[1]);
    if (!pids.size) for (const p of procPids(port)) pids.add(p);
  }
  pids.delete(String(process.pid));
  pids.delete('0');
  return [...pids];
}

if (process.argv[1] && process.argv[1].endsWith('free-ports.mjs')) for (const port of ports) {
  for (const pid of pidsOnPort(port)) {
    run(isWin ? `taskkill /PID ${pid} /F /T` : `kill -9 ${pid}`);
    console.log(`Closed old process ${pid} that was using port ${port}.`);
  }
}
