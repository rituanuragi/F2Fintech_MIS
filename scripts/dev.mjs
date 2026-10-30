// Starts the app on free ports. If 3000 / 5173 are busy it simply uses the next free ones.
import net from 'net';
import { spawn, execSync } from 'child_process';

const isFree = (port) => new Promise((resolve) => {
  const s = net.createServer();
  s.once('error', () => resolve(false));
  s.once('listening', () => s.close(() => resolve(true)));
  s.listen(port); // same way Express and Vite listen
});

async function pickPort(start, avoid = []) {
  for (let p = start; p < start + 200; p++) {
    if (!avoid.includes(p) && (await isFree(p))) return p;
  }
  throw new Error(`No free port found from ${start}`);
}

// Best effort: close old copies of this app first (they could overwrite the saved data).
if (!process.env.NO_KILL) {
  try { execSync('node scripts/free-ports.mjs 3005 5173', { stdio: 'inherit' }); } catch {}
}

const apiPort = await pickPort(3005);
const webPort = await pickPort(5173, [apiPort]);

console.log(`\n=========================================`);
console.log(`  Open this address:  http://localhost:${webPort}`);
console.log(`  (server on port ${apiPort})`);
console.log(`=========================================\n`);

const env = { ...process.env, PORT: String(apiPort), API_PORT: String(apiPort), WEB_PORT: String(webPort) };
const cmd = 'npx concurrently -k --kill-others-on-fail -n server,web "tsx watch server/server.ts" "vite"';
const child = spawn(cmd, { shell: true, stdio: 'inherit', env });
child.on('exit', (code) => process.exit(code ?? 0));
