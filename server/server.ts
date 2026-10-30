/**
 * F2 Lead Intelligence - backend entry point.
 *
 * NOTE: the original serverApp.ts was not available, so the directory routes below
 * (status, dashboard, directory, connect, sync, search, export, preview, record,
 * history, errors, priority-rules) were REBUILT from how the
 * frontend and the engine classes use them. Check each one against real behaviour.
 */
import './env.js';
import express, { Request, Response } from 'express';
import * as path from 'path';
import * as fs from 'fs';
import { getDb, saveDb, getStatsSummary, recomputeStatsSummary } from './db.js';
import { DriveCrawler } from './crawler.js';
import { Indexer, IndexerProgress } from './indexer.js';
import { SearchEngine } from './search.js';
import {
  getPriorityRules, savePriorityRules, resetPriorityRules, DEFAULT_PRIORITY_RULES, evaluateSingleRecordPriority,
  getMisPriorityRules, saveMisPriorityRules, resetMisPriorityRules,
  syncPriorityRulesWithDirectoriesAndFiles, syncMisPriorityRulesWithSources,
} from './priority.js';
import { shubhiRouter, misAutoSyncState, startMisAutoSyncScheduler } from './customerLookupRouter.js';
import { batchCustomerLookupRouter } from './batchCustomerLookup.js';
import { f2TicketsRouter, startTicketsAutoUpdateScheduler } from './f2TicketsRouter.js';
import { healthRouter } from './healthRouter.js';
import { misDashboardRouter } from './misDashboardRouter.js';
import { misSearchBookRouter } from './misSearchBook.js';
import { directoryToolsRouter } from './directoryToolsRouter.js';
import { parseRecordKey } from './recordKey.js';

const PORT = Number(process.env.PORT || 3000);
const APP_VERSION: string = JSON.parse(fs.readFileSync('./package.json', 'utf-8')).version;
const AUTO_SYNC_MINUTES = Number(process.env.AUTO_SYNC_MINUTES || 10);
const app = express();
app.use(express.json({ limit: '100mb' }));

// ---------- shared state ----------
const crawler = new DriveCrawler('./data/cache');
const indexer = new Indexer(crawler);
const searchEngine = new SearchEngine();
let isSyncing = false;
let progress: IndexerProgress = { stage: 'completed', message: 'System ready' };

const rows = (res: any[]): any[][] => (res.length ? res[0].values : []);

async function getActiveSource() {
  const db = await getDb();
  const r = rows(db.exec(
    'SELECT id, name, root_url, root_folder_id, updated_at FROM sources WHERE is_active = 1 ORDER BY updated_at DESC LIMIT 1;'
  ))[0];
  return r ? { id: r[0], name: r[1], rootUrl: r[2], rootFolderId: r[3], updatedAt: r[4] } : null;
}

// The folder the app uses when nothing has been connected yet (same one the screen shows).
const DEFAULT_ROOT_FOLDER_ID = '1xFVOGEiQNshKBpA8tQoedy9FfPTP4rjn';

async function ensureActiveSource() {
  const existing = await getActiveSource();
  if (existing) return existing;
  const db = await getDb();
  const now = new Date().toISOString();
  const id = `src_${DEFAULT_ROOT_FOLDER_ID}`;
  db.run('UPDATE sources SET is_active = 0;');
  db.run(
    'INSERT OR REPLACE INTO sources (id, name, root_url, root_folder_id, is_active, created_at, updated_at) VALUES (?, ?, ?, ?, 1, ?, ?);',
    [id, 'Main MIS Drive Directory', `https://drive.google.com/drive/folders/${DEFAULT_ROOT_FOLDER_ID}?usp=sharing`, DEFAULT_ROOT_FOLDER_ID, now, now]
  );
  saveDb();
  return getActiveSource();
}

async function runSync(sourceId: string, rootFolderId: string, full: boolean) {
  if (isSyncing) return;
  isSyncing = true;
  try {
    await indexer.syncSource(sourceId, rootFolderId, full, (p) => { progress = p; });
    progress = { stage: 'completed', message: 'Sync completed' };
    recomputeStatsSummary(await getDb()); // refresh Dashboard totals
    saveDb();
  } catch (err: any) {
    console.error('Sync failed:', err);
    progress = { stage: 'error', message: err?.message || 'Sync failed' };
    try { recomputeStatsSummary(await getDb()); } catch {}
  } finally {
    isSyncing = false;
  }
}

// ---------- status / dashboard / directory ----------
app.get('/api/status', async (_req, res) => {
  const activeSource = await getActiveSource();
  res.json({ version: APP_VERSION, isConnected: !!activeSource, activeSource, isSyncing, progress, misAutoSync: misAutoSyncState });
});

app.get('/api/dashboard', async (_req, res) => {
  try {
    const db = await getDb();
    const activeSource = await getActiveSource();
    const stats = getStatsSummary(db);
    res.json({ hasConnectedSource: !!activeSource, activeSource, stats });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/directory', async (_req, res) => {
  try {
    const db = await getDb();
    const folders = rows(db.exec('SELECT folder_id, folder_name, folder_path, parent_id FROM folders;')).map((r) => ({
      id: r[0], name: r[1], path: r[2], parentId: r[3],
    }));
    const sheetsByFile: Record<string, any[]> = {};
    for (const r of rows(db.exec('SELECT file_id, sheet_name, rows_count, cols_count, is_indexed, status FROM sheets;'))) {
      (sheetsByFile[r[0] as string] ||= []).push({
        name: r[1], rowsCount: r[2], colsCount: r[3], isIndexed: !!r[4], status: r[5],
      });
    }
    const files = rows(db.exec(
      'SELECT file_id, file_name, folder_id, folder_path, file_size, modified_time, sheets_count, total_rows, status FROM files;'
    )).map((r) => ({
      id: r[0], name: r[1], folderId: r[2], folderPath: r[3], size: r[4] || 0, modifiedTime: r[5],
      sheetsCount: r[6], totalRows: r[7] || 0, status: r[8], sheets: sheetsByFile[r[0] as string] || [],
    }));
    res.json({ folders, files });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// ---------- connect / sync / disconnect ----------
app.post('/api/connect', async (req: Request, res: Response) => {
  try {
    const { rootUrl, name } = req.body || {};
    const rootFolderId = DriveCrawler.extractFolderId(String(rootUrl || ''));
    const db = await getDb();
    const now = new Date().toISOString();
    const id = `src_${rootFolderId}`;
    db.run('UPDATE sources SET is_active = 0;');
    db.run(
      'INSERT OR REPLACE INTO sources (id, name, root_url, root_folder_id, is_active, created_at, updated_at) VALUES (?, ?, ?, ?, 1, ?, ?);',
      [id, name || 'Google Drive Directory', rootUrl, rootFolderId, now, now]
    );
    saveDb();
    runSync(id, rootFolderId, false); // fire and forget; the UI polls /api/status
    res.json({ success: true, activeSource: await getActiveSource() });
  } catch (err: any) {
    res.status(400).json({ error: err.message || 'Connection failed' });
  }
});

app.post('/api/sync', async (req, res) => {
  const src = await ensureActiveSource();
  if (!src) return res.status(400).json({ error: 'No Drive directory connected.' });
  runSync(src.id as string, src.rootFolderId as string, !!req.body?.fullResync);
  res.json({ success: true });
});

app.post('/api/disconnect', async (_req, res) => {
  const db = await getDb();
  db.run('UPDATE sources SET is_active = 0;');
  saveDb();
  res.json({ success: true });
});

// ---------- search / export / preview / record ----------
app.get('/api/search', async (req, res) => {
  try {
    const q = String(req.query.q || '');
    const type = String(req.query.type || 'smart');
    const limit = Math.min(500, Math.max(1, Number(req.query.limit) || 100));
    const offset = Math.max(0, Number(req.query.offset) || 0);
    res.json(await searchEngine.search({ query: q, searchType: type, limit, offset } as any));
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/export', async (req, res) => {
  try {
    const { matches, format } = req.body || {};
    if (!Array.isArray(matches)) return res.status(400).json({ error: 'matches must be an array' });
    const fmt = format === 'csv' ? 'csv' : 'xlsx';
    const buf = await searchEngine.exportResults(matches, fmt);
    res.setHeader('Content-Type', fmt === 'csv' ? 'text/csv' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/preview', async (req, res) => {
  try {
    const q = req.query;
    const num = (v: any) => (v === undefined || v === '' ? null : Number(v));
    res.json(await searchEngine.previewSheet(String(q.fileId || ''), String(q.sheetName || ''), {
      limit: Number(q.limit) || 100,
      offset: Number(q.offset) || 0,
      search: q.search ? String(q.search) : undefined,
      rowNumber: num(q.rowNum),
      rowFrom: num(q.rowFrom),
      rowTo: num(q.rowTo),
      columnName: q.columnName ? String(q.columnName) : null,
      columnValue: q.columnValue ? String(q.columnValue) : null,
      columnOperator: q.columnOperator ? String(q.columnOperator) : null,
      sortBy: q.sortBy ? String(q.sortBy) : null,
      sortOrder: q.sortOrder === 'desc' ? 'desc' : 'asc',
    } as any));
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/export-sheet', async (req, res) => {
  try {
    const { fileId, sheetName, format } = req.body || {};
    const fmt = format === 'csv' ? 'csv' : 'xlsx';
    const buf = await searchEngine.exportSheetPreview(fileId, sheetName, fmt);
    res.setHeader('Content-Type', fmt === 'csv' ? 'text/csv' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/record/:id', async (req, res) => {
  try {
    let id = /^\d+$/.test(req.params.id) ? Number(req.params.id) : NaN;

    // Stable key (workbook + tab + row): the numeric id changes every time a sheet is re-synced
    const loc = parseRecordKey(req.params.id);
    if (loc) {
      const db = await getDb();
      const found = db.exec(
        'SELECT id FROM search_index WHERE file_id = ? AND sheet_name = ? AND row_num = ? ORDER BY id DESC LIMIT 1;',
        [loc.fileId, loc.sheetName, loc.rowNum]
      );
      id = found[0]?.values[0]?.[0] as number;
    }

    if (!Number.isFinite(id)) {
      return res.status(404).json({
        error: loc
          ? 'This row is no longer in the synced data (the sheet may have changed). Sync the sheet and try again.'
          : 'Record not found',
      });
    }
    const out = await searchEngine.getFullRow(id);
    if (!out || !out.match) {
      return res.status(404).json({
        error: 'This row is no longer in the synced data. It was probably replaced by a newer sync. Refresh the list and open it again.',
      });
    }
    res.json(out);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// ---------- history / errors ----------
app.get('/api/history', async (_req, res) => {
  const db = await getDb();
  const runs = rows(db.exec(
    'SELECT id, source_id, started_at, completed_at, duration_ms, files_discovered, files_processed, sheets_processed, rows_processed, errors_count, status FROM sync_runs ORDER BY started_at DESC LIMIT 50;'
  )).map((r) => ({
    id: r[0], sourceId: r[1], startedAt: r[2], completedAt: r[3], durationMs: r[4], filesDiscovered: r[5],
    filesProcessed: r[6], sheetsProcessed: r[7], rowsProcessed: r[8], errorsCount: r[9], status: r[10],
  }));
  res.json({ runs });
});

app.get('/api/errors', async (_req, res) => {
  const db = await getDb();
  const errors = rows(db.exec(
    'SELECT id, source_id, file_id, file_name, sheet_name, error_message, timestamp FROM sync_errors ORDER BY id DESC LIMIT 200;'
  )).map((r) => ({
    id: r[0], sourceId: r[1], fileId: r[2], fileName: r[3], sheetName: r[4], errorMessage: r[5], timestamp: r[6],
  }));
  res.json({ errors });
});

// ---------- priority rules ----------
// GET never changes anything. Syncing happens only on POST .../sync.
app.get('/api/priority-rules', (_req, res) => res.json({ rules: getPriorityRules(), defaultRules: DEFAULT_PRIORITY_RULES }));
app.post('/api/priority-rules', (req, res) => res.json({ rules: savePriorityRules(req.body?.rules || []) }));

// Column names found in the real files of each priority rule (feeds the "Combine Columns" screen).
// Returns { columns: { [ruleId]: string[] }, all: string[] }
app.get('/api/priority-rules/columns', async (_req, res) => {
  try {
    const db = await getDb();
    const rules = getPriorityRules();
    const byRule: Record<string, Set<string>> = {};
    const all = new Set<string>();
    const r = db.exec(`
      SELECT f.file_name, f.folder_path, s.sheet_name, s.headers_json
      FROM sheets s LEFT JOIN files f ON f.file_id = s.file_id
      WHERE f.file_name IS NOT NULL AND f.file_name NOT LIKE 'gsheet_%' AND f.folder_path != 'Google Sheets MIS';
    `);
    for (const v of r[0]?.values || []) {
      let headers: string[] = [];
      try { headers = JSON.parse(String(v[3] || '[]')); } catch { headers = []; }
      headers = headers.map((h) => String(h || '').trim()).filter(Boolean);
      const ev = evaluateSingleRecordPriority(
        { fileName: String(v[0] || ''), folderPath: String(v[1] || ''), sheetName: String(v[2] || '') },
        rules
      );
      if (!byRule[ev.ruleId]) byRule[ev.ruleId] = new Set();
      for (const h of headers) { byRule[ev.ruleId].add(h); all.add(h); }
    }
    const columns: Record<string, string[]> = {};
    for (const [id, set] of Object.entries(byRule)) columns[id] = Array.from(set).sort((a, b) => a.localeCompare(b));
    res.json({ columns, all: Array.from(all).sort((a, b) => a.localeCompare(b)) });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});
app.post('/api/priority-rules/reset', (_req, res) => res.json({ rules: resetPriorityRules() }));
app.post('/api/priority-rules/sync', async (_req, res) => {
  const db = await getDb();
  const dirs = rows(db.exec('SELECT DISTINCT folder_path FROM files;')).map((r) => String(r[0]));
  const files = rows(db.exec('SELECT DISTINCT file_name FROM files;')).map((r) => String(r[0]));
  res.json(syncPriorityRulesWithDirectoriesAndFiles(dirs, files));
});

app.get('/api/mis-priority-rules', (_req, res) => res.json({ rules: getMisPriorityRules(), defaultRules: [] }));
app.post('/api/mis-priority-rules', (req, res) => res.json({ rules: saveMisPriorityRules(req.body?.rules || []) }));
app.post('/api/mis-priority-rules/reset', (_req, res) => res.json({ rules: resetMisPriorityRules() }));
app.post('/api/mis-priority-rules/sync', async (_req, res) => {
  const db = await getDb();
  const sheets = rows(db.exec('SELECT name FROM connected_sheets;')).map((r) => String(r[0]));
  res.json(syncMisPriorityRulesWithSources([], sheets, []));
});


// ---------- routers you already had ----------
// directoryToolsRouter first so its read-only /records/:id/checks etc. are always reachable
app.use('/api', directoryToolsRouter);
app.use('/api', batchCustomerLookupRouter);
app.use('/api', shubhiRouter);
app.use('/api', f2TicketsRouter);
app.use('/api', healthRouter);
app.use('/api', misDashboardRouter);
app.use('/api', misSearchBookRouter);

// Unknown /api routes return JSON (not the HTML page) so the UI shows a clear error
app.use('/api', (req, res) => res.status(404).json({ error: `Route not found: ${req.method} ${req.originalUrl}` }));

// Serve the built frontend in production
const dist = path.resolve('./dist');
app.use(express.static(dist));
app.get('*', (_req, res) => res.sendFile(path.join(dist, 'index.html')));

(async () => {
  const db = await getDb();
  recomputeStatsSummary(db);
  const server = app.listen(PORT, () => console.log(`F2 Lead Intelligence v${APP_VERSION} running on http://localhost:${PORT}`));
  server.on('error', (err: any) => {
    if (err?.code === 'EADDRINUSE') {
      console.error(`\nPort ${PORT} is already used by another (old) copy of this app. Close all terminals and run again.\n`);
    } else {
      console.error('Server error:', err);
    }
    process.exit(1);
  });

  // Auto sync: always check Drive on start, then again every few minutes. Unchanged files are skipped.
  try {
    const src = await ensureActiveSource();
    if (src) {
      console.log('Starting Drive sync...');
      runSync(src.id as string, src.rootFolderId as string, false);
      setInterval(async () => {
        try {
          const cur = await getActiveSource();
          if (cur && !isSyncing) runSync(cur.id as string, cur.rootFolderId as string, false);
        } catch (e) { console.warn('Auto sync skipped:', e); }
      }, AUTO_SYNC_MINUTES * 60 * 1000);
    }
  } catch (e) {
    console.warn('Drive sync could not start:', e);
  }
  startMisAutoSyncScheduler();
  startTicketsAutoUpdateScheduler();
})();
