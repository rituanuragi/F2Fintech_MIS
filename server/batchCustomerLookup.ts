/**
 * Batch Customer Lookup
 *
 * Input : many phone numbers (pasted, or read from an Excel file by the browser)
 * Rule  : MIS FIRST.
 *   1. The number must exist in an MIS sheet. If it does not, the entry stops there
 *      ("not_in_mis"), even when it is in the main directory or in the live tickets.
 *   2. If it is in the MIS, we count every MIS entry (per workbook / tab), pick the primary
 *      entry, and then ALSO check the main directory files and the live ticket API.
 *
 * "What is an MIS row" uses the same rule as /records/search (isDriveRecord) so both agree.
 */
import { Router, Request, Response } from 'express';
import { getDb } from './db.js';
import { findColumnValue } from './columnMatch.js';
import { extractSafeCellValue, normalizePhone } from './normalizer.js';
import { evaluateSingleRecordPriority, getPriorityRules } from './priority.js';
import { findMatchingLiveTicket, ensureTicketsLoaded } from './f2TicketsRouter.js';
import { MisService } from './mis.js';

export const batchCustomerLookupRouter = Router();

const MAX_NUMBERS = 5000;
const CHUNK = 300;
const MAX_MIS_ENTRIES_SHOWN = 25;
const MAX_DIR_ENTRIES_SHOWN = 15;

const isMisRow = (sourceId: string, fileId: string, fileName: string) =>
  sourceId === 'mis_sheet' || fileId.startsWith('gsheet_') || fileName.toLowerCase().includes('mis');

const maskId = (v: string) => {
  const s = String(v || '').trim();
  if (!s) return '';
  return s.length <= 4 ? s : `${'X'.repeat(s.length - 4)}${s.slice(-4)}`;
};

const pick = (row: Record<string, any>, names: string[]) => findColumnValue(row, names)?.value || '';

interface InputItem {
  rawPhone: string;
  sourceSheet?: string;
  sourceRow?: number;
}

batchCustomerLookupRouter.post('/records/batch-lookup', async (req: Request, res: Response) => {
  const started = Date.now();
  try {
    const body = req.body?.phones;
    if (!Array.isArray(body) || body.length === 0) {
      return res.status(400).json({ success: false, error: 'Send at least one phone number.' });
    }
    if (body.length > MAX_NUMBERS) {
      return res.status(400).json({
        success: false,
        error: `Too many numbers (${body.length.toLocaleString()}). The limit is ${MAX_NUMBERS.toLocaleString()} per batch.`,
      });
    }

    const items: InputItem[] = body.map((b: any) =>
      typeof b === 'object' && b !== null
        ? { rawPhone: String(b.rawPhone ?? ''), sourceSheet: b.sourceSheet, sourceRow: b.sourceRow }
        : { rawPhone: String(b ?? '') }
    );

    // ---- 1. normalise and de-duplicate the input -------------------------------------------
    type Group = { inputPhone: string; normPhone: string; inputCount: number; sources: string[] };
    const groups = new Map<string, Group>();
    const invalid: Group[] = [];
    for (const it of items) {
      const raw = it.rawPhone.trim();
      if (!raw) continue;
      const norm = normalizePhone(raw).normalized || '';
      const where = it.sourceRow ? `${it.sourceSheet ? it.sourceSheet + ' ' : ''}row ${it.sourceRow}` : '';
      if (norm.length !== 10) {
        invalid.push({ inputPhone: raw, normPhone: '', inputCount: 1, sources: where ? [where] : [] });
        continue;
      }
      const g = groups.get(norm);
      if (g) {
        g.inputCount += 1;
        if (where && g.sources.length < 5) g.sources.push(where);
      } else {
        groups.set(norm, { inputPhone: raw, normPhone: norm, inputCount: 1, sources: where ? [where] : [] });
      }
    }
    const phones = Array.from(groups.keys());

    // ---- 2. pull every matching row from the index -----------------------------------------
    const db = await getDb();
    const cols =
      'id, source_id, folder_path, file_id, file_name, sheet_name, row_num, customer_name, norm_phone, secondary_phone, pan, city, state, lender, loan_amount, status_val, raw_row_json';

    type DbRow = {
      id: number; sourceId: string; folderPath: string; fileId: string; fileName: string; sheetName: string;
      rowNum: number; customerName: string; normPhone: string; secondaryPhone: string; pan: string; city: string;
      state: string; lender: string; loanAmount: string; status: string; raw: string;
    };
    const toRow = (v: any[]): DbRow => ({
      id: v[0] as number, sourceId: String(v[1] || ''), folderPath: String(v[2] || ''), fileId: String(v[3] || ''),
      fileName: String(v[4] || ''), sheetName: String(v[5] || ''), rowNum: Number(v[6]) || 0,
      customerName: String(v[7] || ''), normPhone: String(v[8] || ''), secondaryPhone: String(v[9] || ''),
      pan: String(v[10] || ''), city: String(v[11] || ''), state: String(v[12] || ''), lender: String(v[13] || ''),
      loanAmount: String(v[14] || ''), status: String(v[15] || ''), raw: String(v[16] || '{}'),
    });

    const byPhone = new Map<string, DbRow[]>();
    const add = (phone: string, row: DbRow) => {
      const list = byPhone.get(phone) || [];
      if (list.some((r) => r.id === row.id)) return; // same row matched by primary and secondary phone
      list.push(row);
      byPhone.set(phone, list);
    };
    const phoneSet = new Set(phones);

    for (let i = 0; i < phones.length; i += CHUNK) {
      const chunk = phones.slice(i, i + CHUNK);
      const ph = chunk.map(() => '?').join(',');

      // primary phone: uses the phone index, works for MIS and directory rows
      const main = db.exec(`SELECT ${cols} FROM search_index WHERE norm_phone IN (${ph});`, chunk);
      for (const v of main[0]?.values || []) {
        const r = toRow(v);
        if (phoneSet.has(r.normPhone)) add(r.normPhone, r);
      }
      // secondary phone: MIS rows only (keeps this query small)
      const sec = db.exec(
        `SELECT ${cols} FROM search_index WHERE source_id = 'mis_sheet' AND secondary_phone IN (${ph});`,
        chunk
      );
      for (const v of sec[0]?.values || []) {
        const r = toRow(v);
        if (phoneSet.has(r.secondaryPhone)) add(r.secondaryPhone, r);
      }
    }

    // MIS sheet priority (Priority #1 sheet is the primary one), same table the single lookup uses
    const sheetRank: Record<string, number> = {};
    try {
      const cs = db.exec('SELECT id, COALESCE(priority_rank, 999) FROM connected_sheets;');
      for (const v of cs[0]?.values || []) sheetRank[String(v[0])] = Number(v[1]) || 999;
    } catch {
      /* table may not exist on a fresh database */
    }
    const rankOf = (fileId: string) => sheetRank[fileId] ?? (fileId === 'mis_master_mis' ? 1 : 999);

    await ensureTicketsLoaded();
    const rules = getPriorityRules();

    // ---- 3. build one result per number ----------------------------------------------------
    const results = phones.map((phone) => {
      const g = groups.get(phone)!;
      const rows = byPhone.get(phone) || [];
      const misRows = rows.filter((r) => isMisRow(r.sourceId, r.fileId, r.fileName));
      const dirRows = rows.filter((r) => !isMisRow(r.sourceId, r.fileId, r.fileName));

      const base = {
        inputPhone: g.inputPhone,
        normPhone: phone,
        inputCount: g.inputCount,
        sources: g.sources,
      };

      // MIS FIRST: not in MIS -> stop here, nothing from directory or API is shown
      if (misRows.length === 0) {
        return { ...base, status: 'not_in_mis' as const, mis: null, directory: null, api: null };
      }

      const misEntries = misRows
        .map((r) => {
          let raw: Record<string, any> = {};
          try {
            const parsed = JSON.parse(r.raw);
            for (const [k, v] of Object.entries(parsed)) raw[k] = extractSafeCellValue(v, k);
          } catch {
            raw = {};
          }
          return {
            id: r.id,
            fileId: r.fileId,
            fileName: r.fileId === 'mis_master_mis' ? 'MASTER MIS' : r.fileName,
            sheetName: r.sheetName,
            rowNum: r.rowNum,
            sheetRank: rankOf(r.fileId),
            customerName: r.customerName || pick(raw, ['customer name', 'name', 'applicant name']),
            lender: r.lender || pick(raw, ['lender', 'bank', 'lender name']),
            loanAmount: r.loanAmount || pick(raw, ['loan amount', 'amount', 'approved amount']),
            status: r.status || pick(raw, ['status', 'case status', 'stage']),
            ticketId: pick(raw, ['ticket id', 'ticket no', 'ticket', 'lead id', 'application id']),
            loginDate: pick(raw, ['login date', 'case created date', 'created date', 'date']),
            approvalDate: pick(raw, ['approval date', 'approved date']),
            disbursedDate: pick(raw, ['disbursed date', 'disbursal date', 'disbursement date']),
            teamLeader: pick(raw, ['team leader', 'tl name', 'tl']),
            manager: pick(raw, ['manager', 'manager name']),
            employee: pick(raw, ['employee name', 'emp name', 'agent', 'executive']),
            city: r.city || pick(raw, ['city']),
            state: r.state || pick(raw, ['state']),
            cibil: pick(raw, ['cibil score', 'cibil']),
            pan: maskId(r.pan || pick(raw, ['pan', 'pan number', 'pan no'])),
            opsRemarks: pick(raw, ['ops remarks', 'ops remark', 'remarks']),
            creditRemarks: pick(raw, ['credit remarks', 'credit remark']),
          };
        })
        // Priority #1 workbook first, then the newest row
        .sort((a, b) => a.sheetRank - b.sheetRank || b.id - a.id);

      // Per workbook / tab breakdown
      const bookMap = new Map<string, { fileName: string; sheetName: string; entries: number; rank: number }>();
      for (const e of misEntries) {
        const k = `${e.fileId}|${e.sheetName}`;
        const b = bookMap.get(k) || { fileName: e.fileName, sheetName: e.sheetName, entries: 0, rank: e.sheetRank };
        b.entries += 1;
        bookMap.set(k, b);
      }
      const statuses: Record<string, number> = {};
      const lenders: Record<string, number> = {};
      for (const e of misEntries) {
        const s = (e.status || 'NO STATUS').toUpperCase();
        statuses[s] = (statuses[s] || 0) + 1;
        const l = (e.lender || 'UNKNOWN').toUpperCase();
        lenders[l] = (lenders[l] || 0) + 1;
      }
      const primary = misEntries[0];

      // Main directory (BL / PL / DOCTORS ...) occurrences
      const directoryEntries = dirRows
        .map((r) => {
          let rawRow: Record<string, any> = {};
          try {
            rawRow = JSON.parse(r.raw);
          } catch {
            rawRow = {};
          }
          let priority = '';
          let rank = 999;
          let badge: { bg?: string; text?: string; border?: string } = {};
          try {
            const ev = evaluateSingleRecordPriority(
              { fileName: r.fileName, folderPath: r.folderPath, sheetName: r.sheetName, rawRow },
              rules
            );
            priority = ev.priorityLabel;
            rank = ev.effectiveRank;
            badge = { bg: ev.badgeBg, text: ev.badgeText, border: ev.badgeBorder };
          } catch {
            /* priority is optional */
          }
          return {
            id: r.id,
            folderPath: r.folderPath,
            fileName: r.fileName,
            sheetName: r.sheetName,
            rowNum: r.rowNum,
            customerName: r.customerName,
            priority,
            rank,
            badge,
          };
        })
        .sort((a, b) => a.rank - b.rank || a.fileName.localeCompare(b.fileName));

      // Live tickets: matched on the primary MIS entry (phone is mandatory inside the matcher)
      let ticket: any = null;
      try {
        ticket =
          findMatchingLiveTicket({
            phone,
            lender: primary.lender,
            name: primary.customerName,
            ticketId: primary.ticketId && !primary.ticketId.startsWith('TCK-') ? primary.ticketId : undefined,
            loanAmount: primary.loanAmount,
          }) || null;
      } catch {
        ticket = null;
      }

      return {
        ...base,
        status: 'in_mis' as const,
        mis: {
          entryCount: misEntries.length,
          workbookCount: new Set(misEntries.map((e) => e.fileId)).size,
          tabCount: bookMap.size,
          books: Array.from(bookMap.values()).sort((a, b) => a.rank - b.rank),
          statuses,
          lenders,
          primary,
          entries: misEntries.slice(0, MAX_MIS_ENTRIES_SHOWN),
          entriesHidden: Math.max(0, misEntries.length - MAX_MIS_ENTRIES_SHOWN),
        },
        directory: {
          entryCount: directoryEntries.length,
          fileCount: new Set(directoryEntries.map((d) => d.fileName)).size,
          entries: directoryEntries.slice(0, MAX_DIR_ENTRIES_SHOWN),
          entriesHidden: Math.max(0, directoryEntries.length - MAX_DIR_ENTRIES_SHOWN),
        },
        api: { matched: !!ticket, ticket },
      };
    });

    const invalidResults = invalid.map((g) => ({
      inputPhone: g.inputPhone,
      normPhone: '',
      inputCount: 1,
      sources: g.sources,
      status: 'invalid' as const,
      mis: null,
      directory: null,
      api: null,
    }));

    const all = [...results, ...invalidResults];
    const inMis = results.filter((r) => r.status === 'in_mis');
    const summary = {
      rowsReceived: items.filter((i) => i.rawPhone.trim()).length,
      uniqueNumbers: phones.length,
      duplicatesInInput: items.filter((i) => i.rawPhone.trim()).length - phones.length - invalid.length,
      invalid: invalid.length,
      inMis: inMis.length,
      notInMis: results.length - inMis.length,
      totalMisEntries: inMis.reduce((n, r) => n + (r.mis?.entryCount || 0), 0),
      inMisAndDirectory: inMis.filter((r) => (r.directory?.entryCount || 0) > 0).length,
      inMisAndApi: inMis.filter((r) => r.api?.matched).length,
      inAllThree: inMis.filter((r) => (r.directory?.entryCount || 0) > 0 && r.api?.matched).length,
      inMisOnly: inMis.filter((r) => (r.directory?.entryCount || 0) === 0 && !r.api?.matched).length,
      multipleMisEntries: inMis.filter((r) => (r.mis?.entryCount || 0) > 1).length,
    };

    res.json({ success: true, summary, results: all, durationMs: Date.now() - started });
  } catch (err: any) {
    res.status(500).json({ success: false, error: err.message || 'Batch lookup failed' });
  }
});

/**
 * Where is the header row and the phone column? Uses the same server rule as the MIS sheet sync,
 * so an uploaded Excel file and a Google Sheet are read the same way.
 * Body: { rows: any[][] }  (first ~40 rows of a sheet, header row included)
 */
batchCustomerLookupRouter.post('/phone-column/detect', (req: Request, res: Response) => {
  try {
    const rows: any[][] = Array.isArray(req.body?.rows) ? req.body.rows.slice(0, 60) : [];
    if (rows.length === 0) return res.json({ headerRowIndex: 0, headers: [], index: 0, name: '' });
    const headerRowIndex = MisService.detectHeaderRowIndex(rows);
    const headers = (rows[headerRowIndex] || []).map((h: any, i: number) => String(h || `Column_${i + 1}`).trim());
    const detected = MisService.detectPhoneColumn(headers, rows.slice(headerRowIndex + 1, headerRowIndex + 31));
    res.json({ headerRowIndex, headers, index: detected.index, name: detected.name });
  } catch (err: any) {
    res.status(500).json({ error: err.message || 'Could not detect the phone column' });
  }
});
