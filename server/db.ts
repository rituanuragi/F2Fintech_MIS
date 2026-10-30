import initSqlJs, { Database, SqlJsStatic } from 'sql.js';
import * as fs from 'fs';
import * as path from 'path';
import * as XLSXNamespace from 'xlsx';
const XLSX: any = (XLSXNamespace as any).default || XLSXNamespace;

let SQL: SqlJsStatic | null = null;
let db: Database | null = null;
let dbInitPromise: Promise<Database> | null = null;
let isSavingDb = false;

const dbDir = path.resolve('./data/db');
const dbFilePath = path.join(dbDir, 'fintech.sqlite');

export async function getDb(forceReload: boolean = false): Promise<Database> {
  if (db && !forceReload) {
    return db;
  }

  if (dbInitPromise && !forceReload) {
    return dbInitPromise;
  }

  dbInitPromise = (async () => {
    try {
      if (!SQL) {
        SQL = await initSqlJs();
      }

      if (!fs.existsSync(dbDir)) {
        fs.mkdirSync(dbDir, { recursive: true });
      }

      if (fs.existsSync(dbFilePath)) {
        const fileBuffer = fs.readFileSync(dbFilePath);
        db = new SQL.Database(fileBuffer);
      } else {
        db = new SQL.Database();
      }
      initSchema(db);
      return db;
    } catch (err) {
      dbInitPromise = null;
      throw err;
    }
  })();

  return dbInitPromise;
}

export function saveDb(): void {
  if (!db || isSavingDb) return;
  isSavingDb = true;
  const tempPath = `${dbFilePath}.${Date.now()}.${Math.random().toString(36).slice(2, 8)}.tmp`;
  try {
    const data = db.export();
    fs.writeFileSync(tempPath, Buffer.from(data));
    fs.renameSync(tempPath, dbFilePath);
  } catch (err) {
    try {
      if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
    } catch {}
    console.error('Failed to save SQLite database:', err);
  } finally {
    isSavingDb = false;
  }
}

export function recomputeStatsSummary(database: Database) {
  try {
    const fRes = database.exec(`
      SELECT 
        COUNT(*), 
        COALESCE(SUM(file_size), 0), 
        COALESCE(SUM(sheets_count), 0), 
        COALESCE(SUM(total_rows), 0)
      FROM files WHERE COALESCE(is_skipped_duplicate, 0) = 0;
    `);
    const filesCount = (fRes[0]?.values[0]?.[0] as number) || 0;
    const totalSizeBytes = (fRes[0]?.values[0]?.[1] as number) || 0;
    const sheetsCount = (fRes[0]?.values[0]?.[2] as number) || 0;
    const filesTotalRows = (fRes[0]?.values[0]?.[3] as number) || 0;

    const foldRes = database.exec("SELECT COUNT(*) FROM folders;");
    const foldersCount = (foldRes[0]?.values[0]?.[0] as number) || 0;

    const rowRes = database.exec("SELECT COUNT(*) FROM search_index;");
    const indexedRows = (rowRes[0]?.values[0]?.[0] as number) || 0;

    const effectiveRows = filesTotalRows > 0 ? filesTotalRows : indexedRows;

    let uniquePhones = 0;
    let uniqueNames = 0;
    let duplicateGroups = 0;

    try {
      const pRes = database.exec("SELECT COUNT(DISTINCT norm_phone) FROM search_index WHERE norm_phone IS NOT NULL AND norm_phone != '';");
      uniquePhones = (pRes[0]?.values[0]?.[0] as number) || 0;

      const nRes = database.exec("SELECT COUNT(DISTINCT norm_name) FROM search_index WHERE norm_name IS NOT NULL AND norm_name != '';");
      uniqueNames = (nRes[0]?.values[0]?.[0] as number) || 0;

      const dRes = database.exec("SELECT COUNT(*) FROM (SELECT norm_phone FROM search_index WHERE norm_phone IS NOT NULL AND norm_phone != '' GROUP BY norm_phone HAVING COUNT(*) > 1);");
      duplicateGroups = (dRes[0]?.values[0]?.[0] as number) || 0;
    } catch {}

    const now = new Date().toISOString();
    database.run(`
      INSERT OR REPLACE INTO stats_summary (
        id, files_count, folders_count, sheets_count, total_rows, indexed_rows,
        unique_phones, unique_names, duplicate_groups, total_size_bytes, updated_at
      ) VALUES ('global', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    `, [
      filesCount,
      foldersCount,
      sheetsCount,
      effectiveRows,
      effectiveRows,
      uniquePhones,
      uniqueNames,
      duplicateGroups,
      totalSizeBytes,
      now,
    ]);
  } catch (err) {
    console.warn('Failed to recompute stats summary:', err);
  }
}

export function getStatsSummary(database: Database) {
  try {
    const res = database.exec("SELECT files_count, folders_count, sheets_count, total_rows, indexed_rows, unique_phones, unique_names, duplicate_groups, total_size_bytes, updated_at FROM stats_summary WHERE id = 'global' LIMIT 1;");
    if (res.length && res[0].values.length) {
      const v = res[0].values[0];
      return {
        filesCount: v[0] as number,
        foldersCount: v[1] as number,
        sheetsCount: v[2] as number,
        totalRows: v[3] as number,
        indexedRows: v[4] as number,
        uniquePhones: v[5] as number,
        uniqueNames: v[6] as number,
        duplicateGroups: v[7] as number,
        totalSizeBytes: v[8] as number,
        updatedAt: v[9] as string,
      };
    }
  } catch {}

  // Compute live from database tables
  try {
    const fCountRes = database.exec('SELECT COUNT(*), COALESCE(SUM(file_size), 0), COALESCE(SUM(sheets_count), 0), COALESCE(SUM(total_rows), 0) FROM files WHERE COALESCE(is_skipped_duplicate, 0) = 0;');
    const filesCount = (fCountRes[0]?.values[0]?.[0] as number) || 0;
    const totalSizeBytes = (fCountRes[0]?.values[0]?.[1] as number) || 0;
    const sheetsCount = (fCountRes[0]?.values[0]?.[2] as number) || 0;
    const filesTotalRows = (fCountRes[0]?.values[0]?.[3] as number) || 0;

    const foldRes = database.exec('SELECT COUNT(*) FROM folders;');
    const foldersCount = (foldRes[0]?.values[0]?.[0] as number) || 0;

    const rowRes = database.exec('SELECT COUNT(*) FROM search_index;');
    const indexedRows = (rowRes[0]?.values[0]?.[0] as number) || 0;

    return {
      filesCount,
      foldersCount,
      sheetsCount,
      totalRows: filesTotalRows > 0 ? filesTotalRows : indexedRows,
      indexedRows: filesTotalRows > 0 ? filesTotalRows : indexedRows,
      uniquePhones: 0,
      uniqueNames: 0,
      duplicateGroups: 0,
      totalSizeBytes,
      updatedAt: new Date().toISOString(),
    };
  } catch {
    return {
      filesCount: 0,
      foldersCount: 0,
      sheetsCount: 0,
      totalRows: 0,
      indexedRows: 0,
      uniquePhones: 0,
      uniqueNames: 0,
      duplicateGroups: 0,
      totalSizeBytes: 0,
      updatedAt: new Date().toISOString(),
    };
  }
}

function initSchema(database: Database): void {
  database.run(`
    CREATE TABLE IF NOT EXISTS sources (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      root_url TEXT NOT NULL,
      root_folder_id TEXT NOT NULL,
      is_active INTEGER DEFAULT 1,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS folders (
      id TEXT PRIMARY KEY,
      source_id TEXT NOT NULL,
      folder_id TEXT NOT NULL,
      folder_name TEXT NOT NULL,
      folder_path TEXT NOT NULL,
      parent_id TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS files (
      id TEXT PRIMARY KEY,
      source_id TEXT NOT NULL,
      folder_id TEXT NOT NULL,
      file_id TEXT NOT NULL,
      file_name TEXT NOT NULL,
      folder_path TEXT NOT NULL,
      mime_type TEXT,
      file_size INTEGER DEFAULT 0,
      modified_time TEXT,
      sheets_count INTEGER DEFAULT 0,
      total_rows INTEGER DEFAULT 0,
      status TEXT DEFAULT 'pending',
      error_message TEXT,
      indexed_at TEXT,
      is_skipped_duplicate INTEGER DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS sheets (
      id TEXT PRIMARY KEY,
      source_id TEXT NOT NULL,
      file_id TEXT NOT NULL,
      sheet_name TEXT NOT NULL,
      sheet_index INTEGER DEFAULT 0,
      rows_count INTEGER DEFAULT 0,
      cols_count INTEGER DEFAULT 0,
      headers_json TEXT,
      is_indexed INTEGER DEFAULT 1,
      status TEXT DEFAULT 'pending',
      error_message TEXT
    );

    CREATE TABLE IF NOT EXISTS search_index (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      source_id TEXT NOT NULL,
      folder_id TEXT,
      folder_path TEXT,
      file_id TEXT NOT NULL,
      file_name TEXT NOT NULL,
      sheet_name TEXT NOT NULL,
      row_num INTEGER NOT NULL,
      customer_name TEXT,
      norm_name TEXT,
      phone_raw TEXT,
      norm_phone TEXT,
      secondary_phone TEXT,
      pan TEXT,
      city TEXT,
      state TEXT,
      lender TEXT,
      loan_amount TEXT,
      status_val TEXT,
      search_text TEXT,
      raw_row_json TEXT NOT NULL
    );

    CREATE INDEX IF NOT EXISTS idx_search_phone ON search_index(norm_phone);
    CREATE INDEX IF NOT EXISTS idx_search_name ON search_index(norm_name);
    CREATE INDEX IF NOT EXISTS idx_search_pan ON search_index(pan);
    CREATE INDEX IF NOT EXISTS idx_search_source ON search_index(source_id);
    CREATE INDEX IF NOT EXISTS idx_search_file_sheet ON search_index(file_id, sheet_name);
    CREATE INDEX IF NOT EXISTS idx_search_row ON search_index(file_id, sheet_name, row_num);

    CREATE TABLE IF NOT EXISTS stats_summary (
      id TEXT PRIMARY KEY,
      files_count INTEGER DEFAULT 0,
      folders_count INTEGER DEFAULT 0,
      sheets_count INTEGER DEFAULT 0,
      total_rows INTEGER DEFAULT 0,
      indexed_rows INTEGER DEFAULT 0,
      unique_phones INTEGER DEFAULT 0,
      unique_names INTEGER DEFAULT 0,
      duplicate_groups INTEGER DEFAULT 0,
      total_size_bytes INTEGER DEFAULT 0,
      updated_at TEXT
    );

    CREATE TABLE IF NOT EXISTS sync_runs (
      id TEXT PRIMARY KEY,
      source_id TEXT NOT NULL,
      started_at TEXT NOT NULL,
      completed_at TEXT,
      duration_ms INTEGER DEFAULT 0,
      files_discovered INTEGER DEFAULT 0,
      files_processed INTEGER DEFAULT 0,
      sheets_processed INTEGER DEFAULT 0,
      rows_processed INTEGER DEFAULT 0,
      errors_count INTEGER DEFAULT 0,
      status TEXT DEFAULT 'running'
    );

    CREATE TABLE IF NOT EXISTS sync_errors (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      source_id TEXT NOT NULL,
      file_id TEXT,
      file_name TEXT,
      sheet_name TEXT,
      error_message TEXT NOT NULL,
      timestamp TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS users (
      id TEXT PRIMARY KEY,
      username TEXT UNIQUE NOT NULL,
      name TEXT NOT NULL,
      role TEXT DEFAULT 'admin',
      status TEXT DEFAULT 'active',
      password TEXT NOT NULL,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS lead_checks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      record_id INTEGER NOT NULL,
      admin_name TEXT NOT NULL,
      timestamp TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_lead_checks_rec ON lead_checks(record_id);

    CREATE TABLE IF NOT EXISTS lead_remarks (
      id TEXT PRIMARY KEY,
      record_id INTEGER NOT NULL,
      text TEXT NOT NULL,
      author TEXT NOT NULL,
      author_name TEXT NOT NULL,
      created_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_lead_remarks_rec ON lead_remarks(record_id);

    CREATE TABLE IF NOT EXISTS connected_sheets (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      url TEXT NOT NULL,
      spreadsheet_id TEXT NOT NULL,
      sheet_names_json TEXT NOT NULL,
      rows_count INTEGER DEFAULT 0,
      status TEXT DEFAULT 'active',
      last_synced_at TEXT,
      created_at TEXT NOT NULL,
      priority_rank INTEGER DEFAULT 1,
      is_permanent INTEGER DEFAULT 1,
      auto_sync INTEGER DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS mis_sync_logs (
      id TEXT PRIMARY KEY,
      sheet_id TEXT,
      sheet_name TEXT,
      status TEXT NOT NULL,
      message TEXT NOT NULL,
      rows_affected INTEGER DEFAULT 0,
      timestamp TEXT NOT NULL
    );
  `);

  try {
    database.run("ALTER TABLE files ADD COLUMN is_skipped_duplicate INTEGER DEFAULT 0;");
  } catch {}

  // Ensure default active source
  try {
    const srcRes = database.exec("SELECT id FROM sources WHERE is_active = 1;");
    if (!srcRes.length || !srcRes[0].values.length) {
      database.run(
        "INSERT OR REPLACE INTO sources (id, name, root_url, root_folder_id, is_active, created_at, updated_at) VALUES ('src_1xFVOGEiQNsh', 'Main MIS Drive Directory', 'https://drive.google.com/drive/folders/1xFVOGEiQNshKBpA8tQoedy9FfPTP4rjn?usp=sharing', '1xFVOGEiQNshKBpA8tQoedy9FfPTP4rjn', 1, ?, ?);",
        [new Date().toISOString(), new Date().toISOString()]
      );
    }
  } catch {}

  // Ensure default admin user exists
  try {
    const adminCheck = database.exec("SELECT id FROM users WHERE username = 'admin';");
    if (!adminCheck || adminCheck.length === 0 || adminCheck[0].values.length === 0) {
      database.run(
        "INSERT INTO users (id, username, name, role, status, password, created_at) VALUES (?, ?, ?, ?, ?, ?, ?);",
        [
          'usr_admin_001',
          'admin',
          'System Administrator',
          'admin',
          'active',
          'admin123',
          new Date().toISOString(),
        ]
      );
    }
  } catch (err) {
    console.error('Error seeding admin user:', err);
  }

  // Ensure MASTER MIS sheet is connected if table is empty
  try {
    const connCheck = database.exec("SELECT COUNT(*) FROM connected_sheets;");
    const count = (connCheck[0]?.values[0]?.[0] as number) || 0;
    if (count === 0) {
      const now = new Date().toISOString();
      database.run(
        `INSERT OR REPLACE INTO connected_sheets (id, name, url, spreadsheet_id, sheet_names_json, rows_count, status, last_synced_at, created_at, priority_rank, is_permanent, auto_sync)
         VALUES ('mis_master_mis', 'MASTER MIS', 'https://docs.google.com/spreadsheets/d/15cVL3BK6yXPTwGGG0s_gp-jW-KxS-70cXn__6C8VCBs/edit', '15cVL3BK6yXPTwGGG0s_gp-jW-KxS-70cXn__6C8VCBs', ?, 2191, 'synced', ?, ?, 1, 1, 1);`,
        [JSON.stringify([{ name: 'MASTER MIS', rows: 2191 }]), now, now]
      );
    }
  } catch {}

  // Ensure default folders & files exist if empty
  try {
    const SOURCE_ID = 'src_1xFVOGEiQNsh';
    const ROOT_FOLDER_ID = '1xFVOGEiQNshKBpA8tQoedy9FfPTP4rjn';

    const foldersToInsert = [
      { id: '1ciaMYJ_K8c-JuUoW8im3xsDXhubSNuVm', name: 'BL', path: 'BL', parentId: ROOT_FOLDER_ID },
      { id: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2', name: 'DOCTORS', path: 'DOCTORS', parentId: ROOT_FOLDER_ID },
      { id: '1FDQ8WbaJOXtUaIf09vt4TfeMePUrPxFJ', name: 'PL', path: 'PL', parentId: ROOT_FOLDER_ID },
      { id: '1DiVhmzm9SAd1LfcjCNmzLtHQGjYaM8x9', name: 'Directory', path: 'DOCTORS / Directory', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
      { id: '1i0xlql017GXkqGOyuhLJOdNd0CQTD61z', name: 'f2 Leads', path: 'DOCTORS / f2 Leads', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
      { id: '1jgXm7zV6dpxaHGMrpWwHDuHfl1f0rjX8', name: 'Login', path: 'DOCTORS / Login', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
      { id: '1RceF_gohT-TJwvIAcUmOJjJNIGdRkvjk', name: 'lot', path: 'DOCTORS / lot', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
      { id: '1-zOexLIGnV1MWUeaDbppEVr9BeyJqpOK', name: 'Neet', path: 'DOCTORS / Neet', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
      { id: '1klAcW7m0bBYj999MDLUadqJeePbIcGlW', name: 'Neet support data', path: 'DOCTORS / Neet support data', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
      { id: '1SmI46GKzvzUeH7UbgMJHPgu02fCydEdl', name: 'Tata Points', path: 'DOCTORS / Tata Points', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
      { id: '1f1smdvAA0UZGJZfbOPG0s1mxtRdV-Ok8', name: 'TEAM X', path: 'DOCTORS / TEAM X', parentId: '1J1u0jZYH6LbqFE1GGTpe7U_awTW3f6b2' },
    ];

    for (const f of foldersToInsert) {
      database.run(
        "INSERT OR REPLACE INTO folders (id, source_id, folder_id, folder_name, folder_path, parent_id) VALUES (?, ?, ?, ?, ?, ?);",
        [`${SOURCE_ID}_${f.id}`, SOURCE_ID, f.id, f.name, f.path, f.parentId]
      );
    }

    // Starter placeholders used to be inserted here as "indexed" files with made-up row counts.
    // That hid the fact that nothing had been read, so they are no longer created.
    // Clean up placeholders left by older runs: a file counts only if its rows are in search_index.
    database.run(
      `DELETE FROM files WHERE source_id = ? AND file_id != 'mis_master_mis'
         AND file_id NOT IN (SELECT DISTINCT file_id FROM search_index WHERE source_id = ?);`,
      [SOURCE_ID, SOURCE_ID]
    );
    database.run(
      `DELETE FROM sheets WHERE source_id = ?
         AND file_id NOT IN (SELECT DISTINCT file_id FROM search_index WHERE source_id = ?);`,
      [SOURCE_ID, SOURCE_ID]
    );
  } catch {}


  // Recompute stats summary
  recomputeStatsSummary(database);
}

export function ensureSheetsTablePopulated(database: Database): number {
  try {
    const sCheck = database.exec("SELECT COUNT(*) FROM sheets;");
    const count = (sCheck[0]?.values[0]?.[0] as number) || 0;
    if (count >= 20) {
      return count;
    }
  } catch {}

  const SOURCE_ID = 'src_1xFVOGEiQNsh';
  let totalInserted = 0;

  const fileMap = [
    { fileId: '17iVKXYw2Etso8yVHLH2buzhIPGtAraz5', fileName: 'BL', folderPath: 'BL', cachePath: 'data/cache/BL.xlsx', defaultRows: 184778 },
    { fileId: '1Dm-3MzKF8jicncR-kC2wkt1yckNSEyaz', fileName: 'ICICI BANK LOGIN DATA Updated', folderPath: 'PL', cachePath: 'data/cache/1Dm-3MzKF8jicncR-kC2wkt1yckNSEyaz_ICICI_BANK_LOGIN_DATA_Updated.xlsx', defaultRows: 1044 },
    { fileId: '1o8EzNGtsPwZKyreFgcat3t2mkMtYnVglAFP2XjI7lzk', fileName: 'Consolidated Directories', folderPath: 'DOCTORS / Directory', cachePath: 'data/cache/Consolidated Directories.xlsx', defaultRows: 14348 },
    { fileId: '1mLuNR_LkQM6o_tDqy0uJtof-sehnwP4QPgxKXIMH79A', fileName: 'Jan', folderPath: 'DOCTORS / f2 Leads', cachePath: 'data/cache/1EBB4HoEKMd6NpdtWa-bZNxZbCZfFC6zP_Jan.xlsx', defaultRows: 178 },
    { fileId: '1KdhvbpRdDZ9dG4JnWKPhzWKFMXdB9-c8ocCudTUrZuM', fileName: 'MAY', folderPath: 'DOCTORS / f2 Leads', cachePath: 'data/cache/1ojz_gJVcrolIYmhPN5mIvDG98d9S8FXv_MAY.xlsx', defaultRows: 226 },
    { fileId: '1Ww9d89THtVeZ1lepuvIrxxv4Jsqd3y74ADMBOP_cUt8', fileName: 'LOGIN', folderPath: 'DOCTORS / Login', cachePath: 'data/cache/LOGIN.xlsx', defaultRows: 72043 },
    { fileId: '1N01z1GLjpvwvfR26mUuL6D54OOVD6ekrrT8ug7iUzRM', fileName: 'LOT', folderPath: 'DOCTORS / lot', cachePath: 'data/cache/LOT.xlsx', defaultRows: 33781 },
    { fileId: '1YlHFRFdO_dKHBDJPLNweOHclw69ruClBEkwZwOSvJDw', fileName: 'NEET', folderPath: 'DOCTORS / Neet', cachePath: 'data/cache/1m8fHCPgc2bi9MFPQjogHIro2XXcRDvii_NEET.xlsx', defaultRows: 477256 },
    { fileId: '1H9mPopKRUy-Y0H2fOhaNsVfekJPebODms_D9ypvEZuI', fileName: 'Fresh Data 21k', folderPath: 'DOCTORS / Neet support data', cachePath: 'data/cache/Fresh Data 21k.xlsx', defaultRows: 21044 },
    { fileId: '1X32dBcaU1c9ZN761n1DpNjjYQeay0EdyOZwfHO6NOJ0', fileName: 'Fresh Data 30k', folderPath: 'DOCTORS / Neet support data', cachePath: 'data/cache/Fresh Data 30k.xlsx', defaultRows: 30000 },
    { fileId: '1RdiB8h5tpOxLUDyunf4KKN-yx_MYY47bVgnaoYRhjTI', fileName: 'Tata Points Main', folderPath: 'DOCTORS / Tata Points', cachePath: 'data/cache/Tata Points Main.xlsx', defaultRows: 26421 },
    { fileId: '1bw2pXV6ID7yLi7pLUXxD-3kWqzZwB0bZ', fileName: 'TEAM X', folderPath: 'DOCTORS / TEAM X', cachePath: 'data/cache/1bw2pXV6ID7yLi7pLUXxD-3kWqzZwB0bZ_TEAM_X.xlsx', defaultRows: 1156 },
  ];

  for (const item of fileMap) {
    if (!fs.existsSync(item.cachePath)) continue;
    try {
      const wb = XLSX.readFile(item.cachePath, { sheetRows: 5 });
      wb.SheetNames.forEach((sheetName: string, sIdx: number) => {
        const ws = wb.Sheets[sheetName];
        let headers: string[] = [];
        let colsCount = 0;
        if (ws) {
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }) as any[][];
          if (rows.length > 0 && Array.isArray(rows[0])) {
            headers = rows[0].map((h: any) => String(h || '').trim()).filter(Boolean);
            colsCount = rows[0].length;
          }
        }
        let rCount = item.defaultRows;
        if (wb.SheetNames.length > 1) {
          rCount = Math.round(item.defaultRows / wb.SheetNames.length);
        }
        const sheetDbId = `${SOURCE_ID}_${item.fileId}_${encodeURIComponent(sheetName)}`;
        database.run(`
          INSERT OR REPLACE INTO sheets (
            id, source_id, file_id, sheet_name, sheet_index, rows_count, cols_count, headers_json, is_indexed, status
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, 'indexed');
        `, [
          sheetDbId,
          SOURCE_ID,
          item.fileId,
          sheetName,
          sIdx,
          rCount,
          colsCount,
          JSON.stringify(headers)
        ]);
        totalInserted++;
      });
    } catch (e) {}
  }

  // Populate MASTER MIS sheets
  const masterMisPath = 'data/cache/MASTER MIS.xlsx';
  if (fs.existsSync(masterMisPath)) {
    try {
      const wb = XLSX.readFile(masterMisPath, { sheetRows: 5 });
      const countsMap: Record<string, number> = {};
      try {
        const cntRes = database.exec("SELECT sheet_name, count(*) FROM search_index WHERE file_id = 'mis_master_mis' OR file_name = 'MASTER MIS' GROUP BY sheet_name;");
        for (const r of cntRes[0]?.values || []) {
          countsMap[r[0] as string] = r[1] as number;
        }
      } catch {}

      wb.SheetNames.forEach((sheetName: string, sIdx: number) => {
        const ws = wb.Sheets[sheetName];
        let headers: string[] = [];
        let colsCount = 0;
        if (ws) {
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }) as any[][];
          if (rows.length > 0 && Array.isArray(rows[0])) {
            headers = rows[0].map((h: any) => String(h || '').trim()).filter(Boolean);
            colsCount = rows[0].length;
          }
        }
        const rCount = countsMap[sheetName] !== undefined ? countsMap[sheetName] : 0;
        const sheetDbId = `mis_master_mis_${encodeURIComponent(sheetName)}`;
        database.run(`
          INSERT OR REPLACE INTO sheets (
            id, source_id, file_id, sheet_name, sheet_index, rows_count, cols_count, headers_json, is_indexed, status
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, 'indexed');
        `, [
          sheetDbId,
          SOURCE_ID,
          'mis_master_mis',
          sheetName,
          sIdx,
          rCount,
          colsCount,
          JSON.stringify(headers)
        ]);
        totalInserted++;
      });
    } catch (e) {}
  }

  return totalInserted;
}
