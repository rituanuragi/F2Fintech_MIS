import * as fs from 'fs';
import * as XLSX from 'xlsx';
import { getDb, saveDb } from './db.js';
import { DriveCrawler, DiscoveredFile } from './crawler.js';
import {
  normalizePhone,
  normalizeName,
  normalizePan,
  classifyHeader,
  extractSafeCellValue,
  rowHasAnyEntry,
} from './normalizer.js';

const xlsxLib: any = (XLSX as any).default || XLSX;
const xlsxUtils: any = xlsxLib.utils || (XLSX as any).utils;

function readWorkbookFromFile(filePath: string, opts: any): any {
  const fileBuf = fs.readFileSync(filePath);
  const readFn = xlsxLib.read || (XLSX as any).read;
  return readFn(fileBuf, { ...opts, type: 'buffer' });
}

export interface IndexerProgress {
  stage: 'discovering' | 'downloading' | 'indexing' | 'completed' | 'error';
  message: string;
  currentFile?: string;
  currentSheet?: string;
  rowsInSheet?: number;
  rowsProcessedInSheet?: number;
  totalRowsProcessed?: number;
  filesProcessed?: number;
  totalFiles?: number;
}

export class Indexer {
  private crawler: DriveCrawler;

  constructor(crawler: DriveCrawler) {
    this.crawler = crawler;
  }

  // Ensure worksheet range covers all existing cells, expanding !ref if truncated
  private ensureFullWorksheetRange(worksheet: XLSX.WorkSheet): void {
    if (!worksheet) return;

    let minRow = Infinity, maxRow = -1, minCol = Infinity, maxCol = -1;
    let hasCells = false;

    for (const key of Object.keys(worksheet)) {
      if (key.startsWith('!')) continue;
      try {
        const cellAddress = xlsxUtils.decode_cell(key);
        hasCells = true;
        if (cellAddress.r < minRow) minRow = cellAddress.r;
        if (cellAddress.r > maxRow) maxRow = cellAddress.r;
        if (cellAddress.c < minCol) minCol = cellAddress.c;
        if (cellAddress.c > maxCol) maxCol = cellAddress.c;
      } catch {}
    }

    if (hasCells) {
      const calculatedRef = xlsxUtils.encode_range({
        s: { r: minRow, c: minCol },
        e: { r: maxRow, c: maxCol },
      });
      if (!worksheet['!ref']) {
        worksheet['!ref'] = calculatedRef;
      } else {
        try {
          const existingRange = xlsxUtils.decode_range(worksheet['!ref']);
          if (maxRow > existingRange.e.r || maxCol > existingRange.e.c || minRow < existingRange.s.r || minCol < existingRange.s.c) {
            worksheet['!ref'] = xlsxUtils.encode_range({
              s: { r: Math.min(minRow, existingRange.s.r), c: Math.min(minCol, existingRange.s.c) },
              e: { r: Math.max(maxRow, existingRange.e.r), c: Math.max(maxCol, existingRange.e.c) },
            });
          }
        } catch {
          worksheet['!ref'] = calculatedRef;
        }
      }
    }
  }

  // Helper to safely extract cell string/primitive value from dense cell or raw value
  private getCellValue(cell: any): any {
    if (cell === null || cell === undefined) return '';
    if (typeof cell === 'object' && cell.v !== undefined) {
      return cell.v;
    }
    return cell;
  }

  // Detect if a specific row is genuinely a column header row
  private isRowGenuineHeader(row: any[]): boolean {
    if (!row || !Array.isArray(row) || row.length === 0) return false;
    let headerKeywords = 0;
    let dataIndicators = 0;
    let textCells = 0;

    for (const cell of row) {
      const val = this.getCellValue(cell);
      if (val === null || val === undefined || val === '') continue;
      const str = String(val).trim();
      if (!str) continue;
      textCells++;

      // Check if it looks like a phone number
      const digits = str.replace(/\D/g, '');
      if (digits.length >= 10 && /^[5-9]\d{9}$/.test(digits.slice(-10))) {
        dataIndicators += 5;
      }
      // Check PAN
      if (/^[A-Z]{5}[0-9]{4}[A-Z]$/.test(str)) {
        dataIndicators += 5;
      }
      // Check pure numeric value > 100
      if (/^\d+(\.\d+)?$/.test(str) && Number(str) > 100) {
        dataIndicators += 2;
      }
      // Check header keywords
      if (
        classifyHeader(str) ||
        /name|phone|mob|contact|pan|aadhar|city|state|address|lender|bank|amount|status|date|remark|email|sr|sno|serial|roll|rank|category|year/i.test(str)
      ) {
        headerKeywords += 3;
      }
    }

    if (textCells === 0) return false;
    return headerKeywords >= 3 && dataIndicators === 0;
  }

  // Detect the best header row index, or -1 if no row is a header
  private detectHeaderRow(rows: any[][]): { headerIndex: number; headers: string[] } {
    if (!rows || rows.length === 0) return { headerIndex: -1, headers: [] };

    const checkLimit = Math.min(5, rows.length);
    for (let i = 0; i < checkLimit; i++) {
      const row = rows[i] || [];
      if (this.isRowGenuineHeader(row)) {
        const headers = [];
        for (let cIdx = 0; cIdx < row.length; cIdx++) {
          const val = this.getCellValue(row[cIdx]);
          const str = String(val ?? '').trim();
          headers.push(str || `Column_${cIdx + 1}`);
        }
        return { headerIndex: i, headers };
      }
    }

    // Fallback: check if row 0 has text cells without numbers/dates while row 1 has numeric/phone data
    if (rows.length > 1) {
      const row0 = rows[0] || [];
      const hasLabels = row0.some(c => {
        const v = String(this.getCellValue(c) ?? '').trim();
        return v.length > 0 && isNaN(Number(v));
      });
      if (hasLabels) {
        const headers = [];
        for (let cIdx = 0; cIdx < row0.length; cIdx++) {
          const val = this.getCellValue(row0[cIdx]);
          const str = String(val ?? '').trim();
          headers.push(str || `Column_${cIdx + 1}`);
        }
        return { headerIndex: 0, headers };
      }
    }

    return { headerIndex: -1, headers: [] };
  }

  // Index a single worksheet with maximum throughput and accurate row tracking
  async indexSheet(
    sourceId: string,
    folderId: string,
    folderPath: string,
    fileId: string,
    fileName: string,
    sheetName: string,
    worksheet: XLSX.WorkSheet,
    onProgress?: (progress: IndexerProgress) => void
  ): Promise<{ rowsIndexed: number; colsCount: number; headers: string[] }> {
    const db = await getDb();

    // Fast pragmas for bulk database ingestion
    db.run('PRAGMA synchronous = OFF;');
    db.run('PRAGMA journal_mode = MEMORY;');
    db.run('PRAGMA temp_store = MEMORY;');

    // Direct dense mode access: avoid expensive sheet_to_json conversions!
    let rawRows: any[][] = [];
    if (Array.isArray(worksheet) && worksheet.length > 0) {
      rawRows = worksheet;
    } else if (worksheet['!data'] && Array.isArray(worksheet['!data']) && worksheet['!data'].length > 0) {
      rawRows = worksheet['!data'];
    } else {
      this.ensureFullWorksheetRange(worksheet);
      rawRows = xlsxUtils.sheet_to_json(worksheet, { header: 1, blankrows: false, raw: true, defval: '' }) as any[][];
    }

    if (!rawRows || rawRows.length === 0) {
      return { rowsIndexed: 0, colsCount: 0, headers: [] };
    }

    const { headerIndex, headers } = this.detectHeaderRow(rawRows);
    let maxCols = headers.length;
    const checkSample = Math.min(30, rawRows.length);
    for (let i = 0; i < checkSample; i++) {
      if (rawRows[i]?.length > maxCols) maxCols = rawRows[i].length;
    }

    while (headers.length < maxCols) {
      headers.push(`Column_${headers.length + 1}`);
    }
    const colsCount = Math.max(headers.length, maxCols);

    // Identify semantic column indices from headers
    const phoneCols: number[] = [];
    const nameCols: number[] = [];
    const panCols: number[] = [];
    const cityCols: number[] = [];
    const stateCols: number[] = [];
    const lenderCols: number[] = [];
    const amountCols: number[] = [];
    const statusCols: number[] = [];

    headers.forEach((h, idx) => {
      const cls = classifyHeader(h);
      if (cls === 'phone') phoneCols.push(idx);
      else if (cls === 'name') nameCols.push(idx);
      else if (cls === 'pan') panCols.push(idx);
      else if (cls === 'city') cityCols.push(idx);
      else if (cls === 'state') stateCols.push(idx);
      else if (cls === 'lender') lenderCols.push(idx);
      else if (cls === 'amount') amountCols.push(idx);
      else if (cls === 'status') statusCols.push(idx);
    });

    if (nameCols.length === 0) {
      headers.forEach((h, idx) => {
        if (/name|customer|client|candidate|applicant|party|doctor|contact/i.test(h) && !/file|sheet|directory/i.test(h)) nameCols.push(idx);
      });
    }
    if (phoneCols.length === 0) {
      headers.forEach((h, idx) => {
        if (/num|mob|tel|contact|calling|cell|phone/i.test(h)) phoneCols.push(idx);
      });
    }

    // Sample first 10 rows to detect fallback phone/name columns if missing
    let detectedPhoneCol = phoneCols[0] !== undefined ? phoneCols[0] : -1;
    let detectedNameCol = nameCols[0] !== undefined ? nameCols[0] : -1;

    if (detectedPhoneCol === -1) {
      for (let r = 0; r < Math.min(10, rawRows.length); r++) {
        if (r === headerIndex) continue;
        const row = rawRows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
          const val = this.getCellValue(row[c]);
          const str = String(val ?? '').replace(/\D/g, '');
          if (str.length >= 10 && /^[5-9]\d{9}$/.test(str.slice(-10))) {
            detectedPhoneCol = c;
            phoneCols.push(c);
            break;
          }
        }
        if (detectedPhoneCol !== -1) break;
      }
    }

    if (detectedNameCol === -1) {
      for (let r = 0; r < Math.min(10, rawRows.length); r++) {
        if (r === headerIndex) continue;
        const row = rawRows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
          if (c === detectedPhoneCol) continue;
          const val = this.getCellValue(row[c]);
          const str = String(val ?? '').trim();
          if (str.length >= 3 && !/^\d+$/.test(str) && !str.includes('@')) {
            detectedNameCol = c;
            nameCols.push(c);
            break;
          }
        }
        if (detectedNameCol !== -1) break;
      }
    }

    // Prepare insert statement
    const insertSql = `
      INSERT INTO search_index (
        source_id, folder_id, folder_path, file_id, file_name, sheet_name,
        row_num, customer_name, norm_name, phone_raw, norm_phone, secondary_phone,
        pan, city, state, lender, loan_amount, status_val, search_text, raw_row_json
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    `;
    let inTransaction = false;
    let stmt: any = null;
    try {
      db.run('BEGIN TRANSACTION;');
      inTransaction = true;
      stmt = db.prepare(insertSql);

      let rowsIndexed = 0;
      const totalDataRows = rawRows.length - (headerIndex >= 0 ? 1 : 0);

      for (let r = 0; r < rawRows.length; r++) {
        if (r === headerIndex) continue;

        const row = rawRows[r];
        if (!row || row.length === 0) continue;
        if (!rowHasAnyEntry(row, headers)) continue; // Strictly skip any empty rows or table templates

        const excelRowNum = r + 1; // 1-based original Excel row number
        const rowObj: Record<string, any> = {};
        const textParts: string[] = [];

        let rawName = '';
        let rawPhone = '';
        let rawPhone2 = '';
        let rawPan = '';
        let rawCity = '';
        let rawState = '';
        let rawLender = '';
        let rawAmount = '';
        let rawStatus = '';

        const rowCols = Math.max(headers.length, row.length);
        let hasData = false;

        for (let c = 0; c < rowCols; c++) {
          const cell = row[c];
          const colHeader = headers[c] || `Column_${c + 1}`;
          const val = cell !== null && cell !== undefined ? extractSafeCellValue(cell, colHeader) : '';
          rowObj[colHeader] = val;

          if (val !== '' && val !== null && val !== undefined) {
            hasData = true;
            textParts.push(String(val));
          }

          // Semantic column extraction
          if (phoneCols.includes(c)) {
            if (!rawPhone) rawPhone = String(val);
            else if (!rawPhone2) rawPhone2 = String(val);
          }
          if (nameCols.includes(c) && !rawName) rawName = String(val);
          if (panCols.includes(c) && !rawPan) rawPan = String(val);
          if (cityCols.includes(c) && !rawCity) rawCity = String(val);
          if (stateCols.includes(c) && !rawState) rawState = String(val);
          if (lenderCols.includes(c) && !rawLender) rawLender = String(val);
          if (amountCols.includes(c) && !rawAmount) rawAmount = String(val);
          if (statusCols.includes(c) && !rawStatus) rawStatus = String(val);
        }

        if (!hasData || !rowHasAnyEntry(row, headers)) continue;

        // Fallback for phone and name
        if (!rawPhone && detectedPhoneCol >= 0 && row[detectedPhoneCol]) {
          rawPhone = String(this.getCellValue(row[detectedPhoneCol]));
        }
        if (!rawName && detectedNameCol >= 0 && row[detectedNameCol]) {
          rawName = String(this.getCellValue(row[detectedNameCol]));
        }

        // Fast normalization
        const cleanPhoneDigits = rawPhone.replace(/\D/g, '');
        const normPhone = cleanPhoneDigits.length >= 10 ? cleanPhoneDigits.slice(-10) : (cleanPhoneDigits || null);
        const cleanPhone2Digits = rawPhone2.replace(/\D/g, '');
        const normPhone2 = cleanPhone2Digits.length >= 10 ? cleanPhone2Digits.slice(-10) : null;
        
        const cleanName = rawName.trim();
        const normName = cleanName ? cleanName.toLowerCase().replace(/[^a-z0-9\s]/g, '').replace(/\s+/g, ' ').trim() : null;
        const normPanVal = normalizePan(rawPan);

        const searchText = textParts.join(' ');
        const rawRowJson = JSON.stringify(rowObj);

        stmt.run([
          sourceId,
          folderId,
          folderPath,
          fileId,
          fileName,
          sheetName,
          excelRowNum,
          cleanName || null,
          normName || null,
          rawPhone || null,
          normPhone,
          normPhone2,
          normPanVal || null,
          rawCity || null,
          rawState || null,
          rawLender || null,
          rawAmount ? String(rawAmount) : null,
          rawStatus || null,
          searchText,
          rawRowJson,
        ]);

        rowsIndexed++;

        // Commit every 5,000 rows BEFORE yielding to the event loop so no transaction is held open across await
        if (rowsIndexed % 5000 === 0) {
          stmt.free();
          stmt = null;
          db.run('COMMIT;');
          inTransaction = false;
          if (onProgress) {
            onProgress({
              stage: 'indexing',
              message: `Indexing ${fileName} → ${sheetName}: ${rowsIndexed.toLocaleString()} / ${totalDataRows.toLocaleString()} rows`,
              currentFile: fileName,
              currentSheet: sheetName,
              rowsInSheet: totalDataRows,
              rowsProcessedInSheet: rowsIndexed,
            });
          }
          await new Promise((resolve) => setTimeout(resolve, 10));
          db.run('BEGIN TRANSACTION;');
          inTransaction = true;
          stmt = db.prepare(insertSql);
        }
      }

      if (stmt) {
        stmt.free();
        stmt = null;
      }
      if (inTransaction) {
        db.run('COMMIT;');
        inTransaction = false;
      }

      return { rowsIndexed, colsCount, headers };
    } catch (err) {
      if (stmt) {
        try { stmt.free(); } catch {}
      }
      if (inTransaction) {
        try { db.run('ROLLBACK;'); } catch {}
      }
      throw err;
    }
  }

  // Ingestion of a discovered Excel file
  async indexFile(
    sourceId: string,
    file: DiscoveredFile,
    forceResync: boolean = false,
    onProgress?: (progress: IndexerProgress) => void
  ): Promise<{ sheetsIndexed: number; totalRows: number }> {
    const db = await getDb();

    // Check if file is already indexed and unchanged (fast incremental sync)
    if (!forceResync) {
      const existing = db.exec(`
        SELECT file_name, file_size, total_rows, status, modified_time FROM files
        WHERE file_id = ? AND source_id = ?;
      `, [file.id, sourceId]);

      if (existing.length > 0 && existing[0].values.length > 0) {
        const [oldName, oldSize, oldRows, oldStatus, oldModified] = existing[0].values[0];

        // If file was renamed: update name in files and search_index instantly!
        const realRowsForRename = (db.exec(
          'SELECT COUNT(*) FROM search_index WHERE file_id = ? AND source_id = ?;', [file.id, sourceId]
        )[0]?.values[0]?.[0] as number) || 0;
        if (oldName !== file.name && oldStatus === 'indexed' && realRowsForRename > 0) {
          db.run('UPDATE files SET file_name = ?, folder_path = ? WHERE file_id = ? AND source_id = ?;', [file.name, file.folderPath, file.id, sourceId]);
          db.run('UPDATE search_index SET file_name = ?, folder_path = ? WHERE file_id = ? AND source_id = ?;', [file.name, file.folderPath, file.id, sourceId]);
          const sCount = db.exec('SELECT COUNT(*) FROM sheets WHERE file_id = ? AND source_id = ?;', [file.id, sourceId]);
          return { sheetsIndexed: (sCount[0]?.values[0]?.[0] as number) || 1, totalRows: (oldRows as number) || 0 };
        }

        // If file is already indexed and unchanged: skip re-indexing.
        // Only when its rows are really in the search index (the starter data lists files as
        // "indexed" before any row has been read, and those must not be skipped).
        const realRows = (db.exec(
          'SELECT COUNT(*) FROM search_index WHERE file_id = ? AND source_id = ?;', [file.id, sourceId]
        )[0]?.values[0]?.[0] as number) || 0;
        // Auto-sync change detection: the Drive "modified" time changed, or the local copy's size changed
        const driveChanged = !!(file.modifiedTime && oldModified && file.modifiedTime !== oldModified);
        const cachedNow = this.crawler.getCachedFilePath(file.id, file.name);
        let localChanged = false;
        try {
          if (cachedNow && oldSize) localChanged = fs.statSync(cachedNow).size !== Number(oldSize);
        } catch {}
        if (driveChanged || localChanged) {
          console.log(`Change detected in ${file.name}; re-reading it.`);
          forceResync = driveChanged; // only re-download when Drive says it changed
        } else if (oldStatus === 'indexed' && (oldRows as number) > 0 && realRows > 0) {
          const sCount = db.exec('SELECT COUNT(*) FROM sheets WHERE file_id = ? AND source_id = ?;', [file.id, sourceId]);
          return { sheetsIndexed: (sCount[0]?.values[0]?.[0] as number) || 1, totalRows: (oldRows as number) || 0 };
        }
      }
    }

    if (onProgress) {
      onProgress({
        stage: 'downloading',
        message: `Verifying & loading ${file.name}...`,
        currentFile: file.name,
      });
    }

    const { filePath, size } = await this.crawler.downloadAndCacheFile(file.id, file.name, forceResync);

    // Read workbook with SheetJS (dense: true for lightning-fast memory efficiency)
    const workbook = readWorkbookFromFile(filePath, {
      dense: true,
      cellFormula: false,
      cellHTML: false,
      cellStyles: false,
      cellDates: false,
    });

    const sheetNames = workbook.SheetNames || [];
    let fileTotalRows = 0;

    // Clear old search_index and sheets entries for this file before re-indexing
    db.run('DELETE FROM search_index WHERE source_id = ? AND file_id = ?;', [sourceId, file.id]);
    db.run('DELETE FROM sheets WHERE source_id = ? AND file_id = ?;', [sourceId, file.id]);

    const fileDbId = `${sourceId}_${file.id}`;

    // Index every worksheet in workbook
    for (let sIdx = 0; sIdx < sheetNames.length; sIdx++) {
      const sheetName = sheetNames[sIdx];
      const worksheet = workbook.Sheets[sheetName];

      if (onProgress) {
        onProgress({
          stage: 'indexing',
          message: `Reading worksheet "${sheetName}" in ${file.name}...`,
          currentFile: file.name,
          currentSheet: sheetName,
        });
      }

      try {
        const { rowsIndexed, colsCount, headers } = await this.indexSheet(
          sourceId,
          file.folderId,
          file.folderPath,
          file.id,
          file.name,
          sheetName,
          worksheet,
          onProgress
        );

        fileTotalRows += rowsIndexed;

        const sheetDbId = `${fileDbId}_${encodeURIComponent(sheetName)}`;
        db.run(`
          INSERT OR REPLACE INTO sheets (
            id, source_id, file_id, sheet_name, sheet_index, rows_count, cols_count, headers_json, is_indexed, status
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, 'indexed');
        `, [
          sheetDbId,
          sourceId,
          file.id,
          sheetName,
          sIdx,
          rowsIndexed,
          colsCount,
          JSON.stringify(headers),
        ]);
      } catch (err: any) {
        console.error(`Error indexing sheet ${sheetName} in ${file.name}:`, err);
        const sheetDbId = `${fileDbId}_${encodeURIComponent(sheetName)}`;
        db.run(`
          INSERT OR REPLACE INTO sheets (
            id, source_id, file_id, sheet_name, sheet_index, rows_count, cols_count, headers_json, is_indexed, status, error_message
          ) VALUES (?, ?, ?, ?, ?, 0, 0, '[]', 1, 'error', ?);
        `, [sheetDbId, sourceId, file.id, sheetName, sIdx, err.message]);

        db.run(`
          INSERT INTO sync_errors (source_id, file_id, file_name, sheet_name, error_message, timestamp)
          VALUES (?, ?, ?, ?, ?, ?);
        `, [sourceId, file.id, file.name, sheetName, err.message, new Date().toISOString()]);
      }
    }

    // Upsert file record with exact sheets_count, total_rows, and indexed status
    db.run(`
      INSERT OR REPLACE INTO files (
        id, source_id, folder_id, file_id, file_name, folder_path,
        mime_type, file_size, modified_time, sheets_count, total_rows, status, indexed_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'indexed', ?);
    `, [
      fileDbId,
      sourceId,
      file.folderId,
      file.id,
      file.name,
      file.folderPath,
      file.mimeType,
      size,
      file.modifiedTime || null,
      sheetNames.length,
      fileTotalRows,
      new Date().toISOString(),
    ]);

    return { sheetsIndexed: sheetNames.length, totalRows: fileTotalRows };
  }

  // Complete Sync of a Google Drive Root Source
  async syncSource(
    sourceId: string,
    rootFolderId: string,
    forceResync: boolean = false,
    onProgress?: (progress: IndexerProgress) => void
  ): Promise<{
    foldersCount: number;
    filesCount: number;
    sheetsCount: number;
    totalRows: number;
    errorsCount: number;
  }> {
    const db = await getDb();
    const startTime = Date.now();
    const syncRunId = `sync_${Date.now()}`;

    // Clear stale sync errors for this source and close any interrupted running syncs
    db.run('DELETE FROM sync_errors WHERE source_id = ?;', [sourceId]);
    db.run("UPDATE sync_runs SET status = 'completed', completed_at = ? WHERE status = 'running';", [new Date().toISOString()]);

    db.run(`
      INSERT INTO sync_runs (id, source_id, started_at, status)
      VALUES (?, ?, ?, 'running');
    `, [syncRunId, sourceId, new Date().toISOString()]);
    saveDb();

    let errorsCount = 0;
    let totalSheets = 0;
    let totalRows = 0;

    try {
      if (onProgress) {
        onProgress({
          stage: 'discovering',
          message: 'Connecting to Google Drive and scanning folders recursively...',
        });
      }

      // Canonical known folders mapping reflecting new 3-directory hierarchy
      const canonicalFolders = [
        { id: '1ciaMYJ_K8c-JuUoW8im3xsDXhubSNuVm', name: 'BL', path: 'BL', parentId: rootFolderId },
        { id: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2', name: 'DOCTORS', path: 'DOCTORS', parentId: rootFolderId },
        { id: '1FDQ8WbaJOXtUaIf09vt4TfeMePUrPxFJ', name: 'PL', path: 'PL', parentId: rootFolderId },
        { id: '1DiVhmzm9SAd1LfcjCNmzLtHQGjYaM8x9', name: 'Directory', path: 'DOCTORS / Directory', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
        { id: '1i0xlql017GXkqGOyuhLJOdNd0CQTD61z', name: 'f2 Leads', path: 'DOCTORS / f2 Leads', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
        { id: '1jgXm7zV6dpxaHGMrpWwHDuHfl1f0rjX8', name: 'Login', path: 'DOCTORS / Login', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
        { id: '1RceF_gohT-TJwvIAcUmOJjJNIGdRkvjk', name: 'lot', path: 'DOCTORS / lot', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
        { id: '1-zOexLIGnV1MWUeaDbppEVr9BeyJqpOK', name: 'Neet', path: 'DOCTORS / Neet', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
        { id: '1klAcW7m0bBYj999MDLUadqJeePbIcGlW', name: 'Neet support data', path: 'DOCTORS / Neet support data', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
        { id: '1SmI46GKzvzUeH7UbgMJHPgu02fCydEdl', name: 'Tata Points', path: 'DOCTORS / Tata Points', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
        { id: '1f1smdvAA0UZGJZfbOPG0s1mxtRdV-Ok8', name: 'TEAM X', path: 'DOCTORS / TEAM X', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
      ];

      // Canonical 12 Google Sheets across the 3 main directories (BL, DOCTORS, PL)
      const canonicalFilesMap: Record<string, { id: string; name: string; folderId: string; folderPath: string }> = {
        'bl': {
          id: '17iVKXYw2Etso8yVHLH2buzhIPGtAraz5',
          name: 'BL',
          folderId: '1ciaMYJ_K8c-JuUoW8im3xsDXhubSNuVm',
          folderPath: 'BL',
        },
        'icici': {
          id: '1Dm-3MzKF8jicncR-kC2wkt1yckNSEyaz',
          name: 'ICICI BANK LOGIN DATA Updated',
          folderId: '1FDQ8WbaJOXtUaIf09vt4TfeMePUrPxFJ',
          folderPath: 'PL',
        },
        'directory': {
          id: '1o8EzNGtsPwZKyreFgcat3t2mkMtYnVglAFP2XjI7lzk',
          name: 'Consolidated Directories',
          folderId: '1DiVhmzm9SAd1LfcjCNmzLtHQGjYaM8x9',
          folderPath: 'DOCTORS / Directory',
        },
        'jan': {
          id: '1mLuNR_LkQM6o_tDqy0uJtof-sehnwP4QPgxKXIMH79A',
          name: 'Jan',
          folderId: '1i0xlql017GXkqGOyuhLJOdNd0CQTD61z',
          folderPath: 'DOCTORS / f2 Leads',
        },
        'may': {
          id: '1KdhvbpRdDZ9dG4JnWKPhzWKFMXdB9-c8ocCudTUrZuM',
          name: 'MAY',
          folderId: '1i0xlql017GXkqGOyuhLJOdNd0CQTD61z',
          folderPath: 'DOCTORS / f2 Leads',
        },
        'login': {
          id: '1Ww9d89THtVeZ1lepuvIrxxv4Jsqd3y74ADMBOP_cUt8',
          name: 'LOGIN',
          folderId: '1jgXm7zV6dpxaHGMrpWwHDuHfl1f0rjX8',
          folderPath: 'DOCTORS / Login',
        },
        'lot': {
          id: '1N01z1GLjpvwvfR26mUuL6D54OOVD6ekrrT8ug7iUzRM',
          name: 'LOT',
          folderId: '1RceF_gohT-TJwvIAcUmOJjJNIGdRkvjk',
          folderPath: 'DOCTORS / lot',
        },
        'neet': {
          id: '1YlHFRFdO_dKHBDJPLNweOHclw69ruClBEkwZwOSvJDw',
          name: 'NEET',
          folderId: '1-zOexLIGnV1MWUeaDbppEVr9BeyJqpOK',
          folderPath: 'DOCTORS / Neet',
        },
        '21k': {
          id: '1H9mPopKRUy-Y0H2fOhaNsVfekJPebODms_D9ypvEZuI',
          name: 'Fresh Data 21k',
          folderId: '1klAcW7m0bBYj999MDLUadqJeePbIcGlW',
          folderPath: 'DOCTORS / Neet support data',
        },
        '30k': {
          id: '1X32dBcaU1c9ZN761n1DpNjjYQeay0EdyOZwfHO6NOJ0',
          name: 'Fresh Data 30k',
          folderId: '1klAcW7m0bBYj999MDLUadqJeePbIcGlW',
          folderPath: 'DOCTORS / Neet support data',
        },
        'tata': {
          id: '1RdiB8h5tpOxLUDyunf4KKN-yx_MYY47bVgnaoYRhjTI',
          name: 'Tata Points Main',
          folderId: '1SmI46GKzvzUeH7UbgMJHPgu02fCydEdl',
          folderPath: 'DOCTORS / Tata Points',
        },
        'team_x': {
          id: '1bw2pXV6ID7yLi7pLUXxD-3kWqzZwB0bZ',
          name: 'TEAM X',
          folderId: '1f1smdvAA0UZGJZfbOPG0s1mxtRdV-Ok8',
          folderPath: 'DOCTORS / TEAM X',
        },
      };

      let folders: { id: string; name: string; path: string; parentId: string | null }[] = [];
      let files: DiscoveredFile[] = [];

      try {
        const scanRes = await this.crawler.scanRecursive(rootFolderId, (msg) => {
          if (onProgress) onProgress({ stage: 'discovering', message: msg });
        });
        if (scanRes.folders.length > 0) {
          // Normalize and clean discovered folders
          folders = scanRes.folders.map(f => ({
            id: f.id,
            name: f.name.trim(),
            path: f.path.trim(),
            parentId: f.parentId || rootFolderId,
          }));
        }
        if (scanRes.files.length > 0) {
          files = scanRes.files;
        }
      } catch (scanErr: any) {
        console.warn('Drive crawl fallback to canonical list:', scanErr.message);
      }

      // Ensure all canonical folders are present and properly parented
      for (const cf of canonicalFolders) {
        const existing = folders.find(f => f.id === cf.id || f.name.toLowerCase() === cf.name.toLowerCase());
        if (!existing) {
          folders.push(cf);
        } else {
          // Keep correct parentId and path
          existing.parentId = cf.parentId;
          existing.path = cf.path;
          existing.name = cf.name;
        }
      }

      const cleanBaseName = (n: string) =>
        (n || '').replace(/\.(xlsx|xls|csv)$/i, '').replace(/^[a-zA-Z0-9_-]{25,60}_/, '').trim().toLowerCase();

      // Ensure all 12 canonical Google Sheet files are present in the list and strip any .xlsx duplicates
      for (const [key, cFile] of Object.entries(canonicalFilesMap)) {
        const cClean = cleanBaseName(cFile.name);
        const existingIdx = files.findIndex(
          f => f.id === cFile.id || cleanBaseName(f.name) === cClean
        );

        if (existingIdx >= 0) {
          // Fix folder mapping and clean Google Sheet name
          files[existingIdx].id = cFile.id;
          files[existingIdx].name = cFile.name.replace(/\.(xlsx|xls|csv)$/i, '');
          files[existingIdx].folderId = cFile.folderId;
          files[existingIdx].folderPath = cFile.folderPath;
          files[existingIdx].mimeType = 'application/vnd.google-apps.spreadsheet';
        } else {
          files.push({
            id: cFile.id,
            name: cFile.name.replace(/\.(xlsx|xls|csv)$/i, ''),
            folderPath: cFile.folderPath,
            folderId: cFile.folderId,
            mimeType: 'application/vnd.google-apps.spreadsheet',
          });
        }
      }

      // Deduplicate files strictly: only 1 entry per clean base name and per file ID
      const canonicalIdsSet = new Set(Object.values(canonicalFilesMap).map(c => c.id));
      const seenFileIds = new Set<string>();
      const seenNames = new Set<string>();

      const excelFiles = files.filter(f => {
        const cleanN = cleanBaseName(f.name);
        if (seenFileIds.has(f.id) || seenNames.has(cleanN)) return false;
        // Skip stale .xlsx IDs that were replaced by Google Sheets
        if (f.name.endsWith('.xlsx') && canonicalIdsSet.has(f.id) === false) {
          return false;
        }
        seenFileIds.add(f.id);
        seenNames.add(cleanN);
        return true;
      });

      // Remove any stale files or sheets from DB that are no longer part of excelFiles
      const validIdsQuote = excelFiles.map(f => `'${f.id.replace(/'/g, "''")}'`).join(',');
      if (validIdsQuote) {
        db.run(`DELETE FROM files WHERE source_id = ? AND file_id NOT IN (${validIdsQuote});`, [sourceId]);
        db.run(`DELETE FROM sheets WHERE source_id = ? AND file_id NOT IN (${validIdsQuote});`, [sourceId]);
        db.run(`DELETE FROM search_index WHERE source_id = ? AND file_id NOT IN (${validIdsQuote});`, [sourceId]);
      }

      // Save discovered folders in DB
      db.run('DELETE FROM folders WHERE source_id = ?;', [sourceId]);
      for (const folder of folders) {
        const folderDbId = `${sourceId}_${folder.id}`;
        db.run(`
          INSERT OR REPLACE INTO folders (id, source_id, folder_id, folder_name, folder_path, parent_id)
          VALUES (?, ?, ?, ?, ?, ?);
        `, [folderDbId, sourceId, folder.id, folder.name, folder.path, folder.parentId]);
      }

      // Determine whether any file actually needs full re-indexing
      let anyFileNeedsReindex = forceResync;
      if (!forceResync) {
        for (const f of excelFiles) {
          const ex = db.exec(
            'SELECT total_rows, status FROM files WHERE file_id = ? AND source_id = ?;',
            [f.id, sourceId]
          );
          const rows = (ex[0]?.values[0]?.[0] as number) || 0;
          const status = (ex[0]?.values[0]?.[1] as string) || '';

          const inSearch = db.exec(
            'SELECT COUNT(*) FROM search_index WHERE file_id = ? AND source_id = ?;',
            [f.id, sourceId]
          );
          const inSearchCount = (inSearch[0]?.values[0]?.[0] as number) || 0;

          if (status !== 'indexed' || rows <= 0 || inSearchCount <= 0) {
            anyFileNeedsReindex = true;
            break;
          }
        }
      }

      // Only drop search indexes if we are actually bulk-inserting new workbook rows
      if (anyFileNeedsReindex) {
        db.run('DROP INDEX IF EXISTS idx_search_phone;');
        db.run('DROP INDEX IF EXISTS idx_search_name;');
        db.run('DROP INDEX IF EXISTS idx_search_pan;');
        db.run('DROP INDEX IF EXISTS idx_search_source;');
        db.run('DROP INDEX IF EXISTS idx_search_file_sheet;');
        db.run('DROP INDEX IF EXISTS idx_search_row;');
      }

      // Process each workbook file
      db.run('DELETE FROM sync_errors WHERE source_id = ?;', [sourceId]);
      errorsCount = 0;
      const fileErrors: string[] = [];
      for (let i = 0; i < excelFiles.length; i++) {
        const f = excelFiles[i];
        if (onProgress) {
          onProgress({
            stage: 'indexing',
            message: `Verifying workbook ${i + 1} of ${excelFiles.length}: ${f.name}`,
            currentFile: f.name,
            filesProcessed: i,
            totalFiles: excelFiles.length,
            totalRowsProcessed: totalRows,
          });
        }

        try {
          const res = await this.indexFile(sourceId, f, forceResync, onProgress);
          totalSheets += res.sheetsIndexed;
          totalRows += res.totalRows;
          // Synchronize folder metadata in search index
          db.run('UPDATE search_index SET folder_id = ?, folder_path = ? WHERE file_id = ?;', [f.folderId, f.folderPath, f.id]);
        } catch (fileErr: any) {
          console.warn(`Skipped or fallback for workbook ${f.name}:`, fileErr.message);
          errorsCount++;
          fileErrors.push(`${f.name}: ${fileErr.message}`);
          db.run(
            'INSERT INTO sync_errors (source_id, file_id, file_name, sheet_name, error_message, timestamp) VALUES (?, ?, ?, ?, ?, ?);',
            [sourceId, f.id, f.name, null, String(fileErr.message || fileErr), new Date().toISOString()]
          );
        }
      }

      // Ensure all search indexes exist
      if (anyFileNeedsReindex) {
        if (onProgress) {
          onProgress({
            stage: 'indexing',
            message: 'Optimizing search indexes and finalizing directory database...',
            totalFiles: excelFiles.length,
            filesProcessed: excelFiles.length,
            totalRowsProcessed: totalRows,
          });
        }
        db.run('CREATE INDEX IF NOT EXISTS idx_search_phone ON search_index(norm_phone);');
        db.run('CREATE INDEX IF NOT EXISTS idx_search_name ON search_index(norm_name);');
        db.run('CREATE INDEX IF NOT EXISTS idx_search_pan ON search_index(pan);');
        db.run('CREATE INDEX IF NOT EXISTS idx_search_source ON search_index(source_id);');
        db.run('CREATE INDEX IF NOT EXISTS idx_search_file_sheet ON search_index(file_id, sheet_name);');
        db.run('CREATE INDEX IF NOT EXISTS idx_search_row ON search_index(file_id, sheet_name, row_num);');
      }

      // If nothing at all could be read, say so instead of reporting a successful sync.
      if (totalRows === 0 && fileErrors.length > 0) {
        saveDb();
        throw new Error(
          `No workbook could be downloaded (${fileErrors.length} failed). First error: ${fileErrors[0]}`
        );
      }

      // Update sync run record
      const duration = Date.now() - startTime;
      db.run(`
        UPDATE sync_runs SET
          completed_at = ?, duration_ms = ?, files_discovered = ?, files_processed = ?,
          sheets_processed = ?, rows_processed = ?, errors_count = ?, status = 'completed'
        WHERE id = ?;
      `, [
        new Date().toISOString(),
        duration,
        excelFiles.length,
        excelFiles.length,
        totalSheets,
        totalRows,
        errorsCount,
        syncRunId,
      ]);

      // Update source updated_at
      db.run('UPDATE sources SET updated_at = ? WHERE id = ?;', [new Date().toISOString(), sourceId]);

      // Persist database to disk only if files were actually re-indexed
      if (anyFileNeedsReindex) {
        saveDb();
      }

      if (onProgress) {
        onProgress({
          stage: 'completed',
          message: `Indexing completed successfully! Ingested ${totalRows.toLocaleString()} rows from ${excelFiles.length} workbooks (${totalSheets} sheets).`,
          filesProcessed: excelFiles.length,
          totalFiles: excelFiles.length,
          totalRowsProcessed: totalRows,
        });
      }

      return {
        foldersCount: folders.length,
        filesCount: excelFiles.length,
        sheetsCount: totalSheets,
        totalRows,
        errorsCount,
      };
    } catch (err: any) {
      console.error('Sync failed:', err);
      db.run(`
        UPDATE sync_runs SET completed_at = ?, errors_count = ?, status = 'failed'
        WHERE id = ?;
      `, [new Date().toISOString(), errorsCount + 1, syncRunId]);
      saveDb();
      throw err;
    }
  }
}
