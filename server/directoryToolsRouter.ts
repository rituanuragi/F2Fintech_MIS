/**
 * Routes used by the Sheets, Duplicates, Batch Excel Search and Customer Lookup screens.
 * New file. It only adds routes; nothing existing is changed.
 *   GET  /api/sheets
 *   GET  /api/duplicates            POST /api/duplicates/export
 *   POST /api/batch                 POST /api/batch/export
 *   GET  /api/records/:id/checks    (read-only audit trail, creates no entry)
 */
import { Router, Request, Response } from 'express';
import { getDb } from './db.js';
import { SearchEngine } from './search.js';

export const directoryToolsRouter = Router();
const engine = new SearchEngine();

const MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// ---------- Sheets ----------
directoryToolsRouter.get('/sheets', async (_req: Request, res: Response) => {
  try {
    const db = await getDb();
    const r = db.exec(`
      SELECT s.id, s.file_id, f.file_name, f.folder_path, s.sheet_name, s.rows_count, s.cols_count,
             s.headers_json, s.is_indexed, s.status
      FROM sheets s
      LEFT JOIN files f ON f.file_id = s.file_id
      ORDER BY f.folder_path, f.file_name, s.sheet_index;
    `);
    const sheets = (r[0]?.values || []).map((v) => {
      let headers: string[] = [];
      try { headers = JSON.parse((v[7] as string) || '[]'); } catch { headers = []; }
      return {
        id: v[0], fileId: v[1], fileName: v[2] || '', folderPath: v[3] || '', sheetName: v[4],
        rowsCount: v[5] || 0, colsCount: v[6] || 0, headers, isIndexed: !!v[8], status: v[9] || '',
      };
    });
    res.json({ sheets });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// ---------- Duplicates ----------
directoryToolsRouter.get('/duplicates', async (req: Request, res: Response) => {
  try {
    const q = req.query;
    const section = ['cross', 'within', 'both'].includes(String(q.section)) ? (q.section as any) : 'cross';
    const countFilter = ['all', '2', '3', '4', '5', 'gt5'].includes(String(q.countFilter)) ? (q.countFilter as any) : 'all';
    const out = await engine.findDuplicates({
      search: q.search ? String(q.search) : '',
      page: Number(q.page) || 1,
      limit: Number(q.limit) || 60,
      section,
      countFilter,
    });
    res.json(out);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

directoryToolsRouter.post('/duplicates/export', async (req: Request, res: Response) => {
  try {
    const { countFilter, search } = req.body || {};
    const buf = await engine.exportDuplicatesExcel({
      countFilter: ['all', '2', '3', '4', '5', 'gt5'].includes(countFilter) ? countFilter : 'all',
      search: search ? String(search) : '',
    });
    res.setHeader('Content-Type', MIME_XLSX);
    res.setHeader('Content-Disposition', 'attachment; filename="F2_Duplicates.xlsx"');
    res.send(buf);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// ---------- Batch phone lookup ----------
directoryToolsRouter.post('/batch', async (req: Request, res: Response) => {
  try {
    const { phones, scope } = req.body || {};
    if (!Array.isArray(phones) || phones.length === 0) {
      return res.status(400).json({ error: 'phones must be a non-empty array' });
    }
    if (phones.length > 50000) {
      return res.status(400).json({ error: 'Too many numbers in one batch (limit 50,000).' });
    }
    const list = phones.map((p: any) =>
      typeof p === 'object' && p !== null
        ? { rawPhone: String(p.rawPhone ?? p.phone ?? ''), originalData: p.originalData }
        : { rawPhone: String(p ?? '') }
    );
    res.json(await engine.batchPhoneLookup(list, { scope: scope === 'all' ? 'all' : 'neet_login_lot' }));
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

directoryToolsRouter.post('/batch/export', async (req: Request, res: Response) => {
  try {
    const { results, format, includePivot, pivotTitle, dataHeader, countHeader, blanksLabel, totalLabel, useTitleAsTabName } = req.body || {};
    if (!Array.isArray(results)) return res.status(400).json({ error: 'results must be an array' });
    const fmt = format === 'csv' ? 'csv' : 'xlsx';
    // header texts typed by the user: keep them short and plain
    const clean = (v: any) => (v === undefined || v === null ? undefined : String(v).trim().slice(0, 120) || undefined);
    const buf = await engine.exportBatchLookup(results, fmt, {
      includePivot: includePivot !== false,
      pivotTitle: clean(pivotTitle),
      dataHeader: clean(dataHeader),
      countHeader: clean(countHeader),
      blanksLabel: clean(blanksLabel),
      totalLabel: clean(totalLabel),
      useHeaderAsSheetName: useTitleAsTabName === true,
    });
    res.setHeader('Content-Type', fmt === 'csv' ? 'text/csv' : MIME_XLSX);
    res.send(buf);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// ---------- Read-only audit trail ----------
directoryToolsRouter.get('/records/:id/checks', async (req: Request, res: Response) => {
  const recordId = parseInt(req.params.id, 10);
  if (isNaN(recordId)) return res.status(400).json({ error: 'Invalid record ID' });
  try {
    const db = await getDb();
    const r = db.exec('SELECT id, admin_name, timestamp FROM lead_checks WHERE record_id = ? ORDER BY id DESC;', [recordId]);
    const history = (r[0]?.values || []).map((v) => ({ id: v[0], adminName: v[1], timestamp: v[2] }));
    res.json({ checkCount: history.length, checkHistory: history });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});
