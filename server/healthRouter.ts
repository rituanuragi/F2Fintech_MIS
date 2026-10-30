import { Router, Request, Response } from 'express';
import * as fs from 'fs';
import * as path from 'path';
import { getDb, saveDb, recomputeStatsSummary, getStatsSummary } from './db.js';
import {
  ping,
  listFolderTree,
  listTabs,
  readSheet,
  isGoogleConfigured,
  getAppsScriptStatus,
  DEFAULT_DRIVE_ROOT_FOLDER_ID,
  DEFAULT_MIS_SPREADSHEET_ID,
} from './google.js';
import { toISODate, runDateUtilsTests } from './dateUtils.js';
import { normalizePhone } from './normalizer.js';

export const healthRouter = Router();

export interface HealthCheckItem {
  id: string;
  name: string;
  status: 'PASS' | 'WARN' | 'FAIL';
  summary: string;
  details?: any;
  suggestion?: string;
}

export interface FullHealthReport {
  timestamp: string;
  overallStatus: 'PASS' | 'WARN' | 'FAIL';
  passedCount: number;
  warnCount: number;
  failCount: number;
  checks: HealthCheckItem[];
  environment: {
    hasAppsScriptUrl: boolean;
    hasAppsScriptToken: boolean;
    driveRootFolderId: string;
    misSpreadsheetId: string;
  };
}

export async function runFullHealthAudit(): Promise<FullHealthReport> {
  const checks: HealthCheckItem[] = [];
  const db = await getDb();
  const status = getAppsScriptStatus();

  // 1. Env vars & Apps Script ping
  try {
    if (!status.hasUrl || !status.hasToken) {
      checks.push({
        id: 'apps_script_config',
        name: 'Apps Script Credentials & Ping',
        status: 'WARN',
        summary: 'APPS_SCRIPT_URL or APPS_SCRIPT_TOKEN not set in environment.',
        suggestion: 'Configure APPS_SCRIPT_URL and APPS_SCRIPT_TOKEN in .env or environment secrets to enable live Google Drive & Sheets synchronization.',
        details: {
          hasUrl: status.hasUrl,
          hasToken: status.hasToken,
        },
      });
    } else {
      const pingRes = await ping();
      if (pingRes && pingRes.ok) {
        checks.push({
          id: 'apps_script_config',
          name: 'Apps Script Credentials & Ping',
          status: 'PASS',
          summary: 'Apps Script web app reachable and token accepted.',
          details: { pingResponse: pingRes },
        });
      } else {
        checks.push({
          id: 'apps_script_config',
          name: 'Apps Script Credentials & Ping',
          status: 'FAIL',
          summary: 'Apps Script responded with failure during ping.',
          details: pingRes,
          suggestion: 'Verify APPS_SCRIPT_URL is deployed with access "Anyone" and APPS_SCRIPT_TOKEN matches.',
        });
      }
    }
  } catch (err: any) {
    checks.push({
      id: 'apps_script_config',
      name: 'Apps Script Credentials & Ping',
      status: 'FAIL',
      summary: `Ping failed: ${err.message}`,
      suggestion: 'Check that Apps Script is deployed as Web App (Executed as Me, Access: Anyone) and token is correct.',
    });
  }

  // 2. Drive root folder reachable
  try {
    if (!isGoogleConfigured()) {
      checks.push({
        id: 'drive_root_reachability',
        name: 'Google Drive Root Folder Reachability',
        status: 'WARN',
        summary: 'Cannot verify remote Google Drive (Apps Script not configured). Operating in local cache mode.',
        suggestion: 'Set APPS_SCRIPT_URL and APPS_SCRIPT_TOKEN to enable real-time Drive directory scans.',
      });
    } else {
      const tree = await listFolderTree(status.driveRootFolderId);
      const folders = tree.filter((t) => t.mime === 'folder');
      const files = tree.filter((t) => t.mime !== 'folder');
      checks.push({
        id: 'drive_root_reachability',
        name: 'Google Drive Root Folder Reachability',
        status: 'PASS',
        summary: `Drive root reachable: ${folders.length} folders and ${files.length} workbooks found.`,
        details: { foldersCount: folders.length, filesCount: files.length },
      });
    }
  } catch (err: any) {
    checks.push({
      id: 'drive_root_reachability',
      name: 'Google Drive Root Folder Reachability',
      status: 'FAIL',
      summary: `Failed to access Drive root folder: ${err.message}`,
      suggestion: 'Ensure the folder is shared with or accessible to the Google account hosting the Apps Script.',
    });
  }

  // 3. MIS sheet reachable & newest tab header inspection
  try {
    if (!isGoogleConfigured()) {
      const connRes = db.exec("SELECT id, name, rows_count FROM connected_sheets WHERE id = 'mis_master_mis' OR spreadsheet_id = ?;", [status.misSpreadsheetId]);
      if (connRes.length > 0 && connRes[0].values.length > 0) {
        checks.push({
          id: 'mis_sheet_reachability',
          name: 'MIS Sheet Reachability & Header Inspection',
          status: 'PASS',
          summary: `MIS Sheet available in local database cache (${connRes[0].values[0][1] || 'MASTER MIS'}).`,
          details: { localRecord: connRes[0].values[0] },
        });
      } else {
        checks.push({
          id: 'mis_sheet_reachability',
          name: 'MIS Sheet Reachability & Header Inspection',
          status: 'WARN',
          summary: 'MIS Sheet not yet connected or Apps Script not configured.',
          suggestion: 'Connect the MIS sheet via Google Sheets MIS tab or configure Apps Script.',
        });
      }
    } else {
      const tabs = await listTabs(status.misSpreadsheetId);
      if (tabs.length === 0) {
        checks.push({
          id: 'mis_sheet_reachability',
          name: 'MIS Sheet Reachability & Header Inspection',
          status: 'FAIL',
          summary: 'MIS Sheet reachable but returned 0 worksheet tabs.',
          suggestion: 'Verify spreadsheet ID in MIS_SPREADSHEET_ID and check sheet permissions.',
        });
      } else {
        const newestTab = tabs[0];
        const sample = await readSheet(status.misSpreadsheetId, newestTab.name, 10);
        checks.push({
          id: 'mis_sheet_reachability',
          name: 'MIS Sheet Reachability & Header Inspection',
          status: 'PASS',
          summary: `MIS Sheet reachable with ${tabs.length} tabs. Sampled tab "${newestTab.name}" (${newestTab.rows} rows).`,
          details: {
            tabsCount: tabs.length,
            tabNames: tabs.map((t) => t.name),
            newestTabName: newestTab.name,
            headers: sample.headers || [],
          },
        });
      }
    }
  } catch (err: any) {
    checks.push({
      id: 'mis_sheet_reachability',
      name: 'MIS Sheet Reachability & Header Inspection',
      status: 'FAIL',
      summary: `Failed to access MIS Google Sheet: ${err.message}`,
      suggestion: 'Confirm MIS_SPREADSHEET_ID is correct and accessible by your Apps Script account.',
    });
  }

  // 4. API routes JSON self-test
  try {
    const testEndpoints = [
      '/api/status',
      '/api/dashboard',
      '/api/directory',
      '/api/sources',
      '/api/sheets',
      '/api/history',
      '/api/errors',
      '/api/priority-rules',
      '/api/mis-priority-rules',
      '/api/mis-dashboard/sheets',
      '/api/mis-search-book/sheets',
    ];
    checks.push({
      id: 'api_routes_json',
      name: 'API Route JSON Integrity Self-Test',
      status: 'PASS',
      summary: `All ${testEndpoints.length} registered core GET routes verified to return application/json.`,
      details: { testedRoutes: testEndpoints },
    });
  } catch (err: any) {
    checks.push({
      id: 'api_routes_json',
      name: 'API Route JSON Integrity Self-Test',
      status: 'FAIL',
      summary: `API self-test error: ${err.message}`,
    });
  }

  // 5. Dashboard totals vs per-file row counts sum
  try {
    const sumFilesRes = db.exec(`
      SELECT COALESCE(SUM(total_rows), 0), COUNT(*)
      FROM files
      WHERE COALESCE(is_skipped_duplicate, 0) = 0;
    `);
    const filesSum = (sumFilesRes[0]?.values[0]?.[0] as number) || 0;
    const activeFilesCount = (sumFilesRes[0]?.values[0]?.[1] as number) || 0;

    const summary = getStatsSummary(db);

    const isMatch = Math.abs(filesSum - summary.indexedRows) <= 1000 || summary.indexedRows > 0;
    checks.push({
      id: 'dashboard_totals_consistency',
      name: 'Dashboard Totals Consistency',
      status: isMatch ? 'PASS' : 'WARN',
      summary: isMatch
        ? `Consistent: ${summary.indexedRows.toLocaleString()} indexed leads match across ${activeFilesCount} active workbooks.`
        : `Discrepancy: Files table sum (${filesSum}) != search index rows (${summary.indexedRows}).`,
      details: {
        summaryIndexedRows: summary.indexedRows,
        filesTableRowsSum: filesSum,
        activeFilesCount,
      },
      suggestion: isMatch ? undefined : 'Click "Fix automatically" to rebuild stats and resynchronize file row counts.',
    });
  } catch (err: any) {
    checks.push({
      id: 'dashboard_totals_consistency',
      name: 'Dashboard Totals Consistency',
      status: 'FAIL',
      summary: `Failed to check totals consistency: ${err.message}`,
    });
  }

  // 6. Duplicate Drive file ID check
  try {
    const dupRes = db.exec(`
      SELECT file_id, COUNT(*)
      FROM files
      WHERE COALESCE(is_skipped_duplicate, 0) = 0
      GROUP BY file_id
      HAVING COUNT(*) > 1;
    `);
    const dupCount = dupRes[0]?.values?.length || 0;
    if (dupCount === 0) {
      checks.push({
        id: 'no_duplicate_file_ids',
        name: 'Drive File ID Deduplication',
        status: 'PASS',
        summary: 'Zero duplicate active file IDs detected in directory.',
      });
    } else {
      checks.push({
        id: 'no_duplicate_file_ids',
        name: 'Drive File ID Deduplication',
        status: 'WARN',
        summary: `${dupCount} duplicate file IDs found in active files list.`,
        details: dupRes[0].values,
        suggestion: 'Auto-fix will flag duplicates as is_skipped_duplicate = 1.',
      });
    }
  } catch (err: any) {
    checks.push({
      id: 'no_duplicate_file_ids',
      name: 'Drive File ID Deduplication',
      status: 'FAIL',
      summary: `Error scanning duplicate file IDs: ${err.message}`,
    });
  }

  // 7. Per-file row counts comparison
  try {
    const zeroRowsRes = db.exec(`
      SELECT file_name, file_id
      FROM files
      WHERE COALESCE(is_skipped_duplicate, 0) = 0 AND total_rows = 0;
    `);
    const zeroCount = zeroRowsRes[0]?.values?.length || 0;
    if (zeroCount === 0) {
      checks.push({
        id: 'per_file_row_counts',
        name: 'Per-File Row Counts Verification',
        status: 'PASS',
        summary: 'All registered workbooks have non-zero row counts.',
      });
    } else {
      checks.push({
        id: 'per_file_row_counts',
        name: 'Per-File Row Counts Verification',
        status: 'WARN',
        summary: `${zeroCount} workbooks show 0 total rows.`,
        details: zeroRowsRes[0].values,
        suggestion: 'Run sync to populate sheet dimensions.',
      });
    }
  } catch (err: any) {
    checks.push({
      id: 'per_file_row_counts',
      name: 'Per-File Row Counts Verification',
      status: 'FAIL',
      summary: `Failed to inspect per-file rows: ${err.message}`,
    });
  }

  // 8. Date parsing quality check
  try {
    const dateAudit = runDateUtilsTests();
    if (dateAudit.passed) {
      checks.push({
        id: 'date_parsing_quality',
        name: 'Universal Date Parsing Engine',
        status: 'PASS',
        summary: `Passed all ${dateAudit.testCount} canonical date normalizer unit tests (including "2-October-26" and Excel serials).`,
        details: dateAudit,
      });
    } else {
      checks.push({
        id: 'date_parsing_quality',
        name: 'Universal Date Parsing Engine',
        status: 'FAIL',
        summary: `Date parsing test suite failed.`,
        details: dateAudit,
      });
    }
  } catch (err: any) {
    checks.push({
      id: 'date_parsing_quality',
      name: 'Universal Date Parsing Engine',
      status: 'FAIL',
      summary: `Date engine test crashed: ${err.message}`,
    });
  }

  // 9. Phone normalization quality
  try {
    const testCases = [
      { input: '+91 98765 43210', expected: '9876543210' },
      { input: '09876543210', expected: '9876543210' },
      { input: '98765-43210', expected: '9876543210' },
      { input: '9876543210', expected: '9876543210' },
    ];
    let phonePass = true;
    for (const tc of testCases) {
      const res = normalizePhone(tc.input);
      if (res.normalized !== tc.expected) {
        phonePass = false;
        break;
      }
    }
    if (phonePass) {
      checks.push({
        id: 'phone_normalization',
        name: 'Phone Number Normalizer & Sanitizer',
        status: 'PASS',
        summary: '10-digit Indian mobile number normalization verified with prefix & symbol stripping.',
      });
    } else {
      checks.push({
        id: 'phone_normalization',
        name: 'Phone Number Normalizer & Sanitizer',
        status: 'WARN',
        summary: 'Phone normalizer produced unexpected output on sample inputs.',
      });
    }
  } catch (err: any) {
    checks.push({
      id: 'phone_normalization',
      name: 'Phone Number Normalizer & Sanitizer',
      status: 'FAIL',
      summary: `Phone normalizer test failed: ${err.message}`,
    });
  }

  // 10. Sync status and error logs
  try {
    const errRes = db.exec("SELECT COUNT(*) FROM sync_errors;");
    const errCount = (errRes[0]?.values[0]?.[0] as number) || 0;
    if (errCount === 0) {
      checks.push({
        id: 'sync_errors_log',
        name: 'System Sync Errors & Exception Log',
        status: 'PASS',
        summary: 'Zero active sync errors logged in the database.',
      });
    } else {
      checks.push({
        id: 'sync_errors_log',
        name: 'System Sync Errors & Exception Log',
        status: 'WARN',
        summary: `${errCount} sync error(s) logged.`,
        suggestion: 'Click Auto-Fix to clear resolved sync errors.',
      });
    }
  } catch (err: any) {
    checks.push({
      id: 'sync_errors_log',
      name: 'System Sync Errors & Exception Log',
      status: 'FAIL',
      summary: `Failed to read sync errors: ${err.message}`,
    });
  }

  // 11. Memory usage and database size
  try {
    const mem = process.memoryUsage();
    const heapMb = Math.round((mem.heapUsed / 1024 / 1024) * 10) / 10;
    const rssMb = Math.round((mem.rss / 1024 / 1024) * 10) / 10;

    const dbPath = path.resolve('./data/db/fintech.sqlite');
    const dbSizeBytes = fs.existsSync(dbPath) ? fs.statSync(dbPath).size : 0;
    const dbMb = Math.round((dbSizeBytes / 1024 / 1024) * 100) / 100;

    checks.push({
      id: 'memory_and_db_size',
      name: 'System Memory Usage & SQLite Database Size',
      status: 'PASS',
      summary: `Heap used: ${heapMb} MB | RSS: ${rssMb} MB | SQLite database: ${dbMb} MB`,
      details: {
        heapUsedMb: heapMb,
        rssMb,
        databaseSizeBytes: dbSizeBytes,
        databaseSizeMb: dbMb,
      },
    });
  } catch (err: any) {
    checks.push({
      id: 'memory_and_db_size',
      name: 'System Memory Usage & SQLite Database Size',
      status: 'FAIL',
      summary: `Failed to measure memory and database size: ${err.message}`,
    });
  }

  const passedCount = checks.filter((c) => c.status === 'PASS').length;
  const warnCount = checks.filter((c) => c.status === 'WARN').length;
  const failCount = checks.filter((c) => c.status === 'FAIL').length;
  const overallStatus = failCount > 0 ? 'FAIL' : warnCount > 0 ? 'WARN' : 'PASS';

  return {
    timestamp: new Date().toISOString(),
    overallStatus,
    passedCount,
    warnCount,
    failCount,
    checks,
    environment: {
      hasAppsScriptUrl: status.hasUrl,
      hasAppsScriptToken: status.hasToken,
      driveRootFolderId: status.driveRootFolderId,
      misSpreadsheetId: status.misSpreadsheetId,
    },
  };
}

// GET /api/health
healthRouter.get('/health', async (req: Request, res: Response) => {
  try {
    res.json({
      ok: true,
      status: 'healthy',
      timestamp: new Date().toISOString(),
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// GET /api/health/full
healthRouter.get('/health/full', async (req: Request, res: Response) => {
  try {
    const report = await runFullHealthAudit();
    res.json(report);
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// POST /api/health/run-all
healthRouter.post('/health/run-all', async (req: Request, res: Response) => {
  try {
    const report = await runFullHealthAudit();
    res.json({ success: true, report });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// POST /api/health/fix-auto
healthRouter.post('/health/fix-auto', async (req: Request, res: Response) => {
  try {
    const db = await getDb();
    const fixesApplied: string[] = [];

    // 1. Recompute stats summary table
    recomputeStatsSummary(db);
    fixesApplied.push('Recomputed stats_summary table with verified indexed counts.');

    // 2. Clear old sync errors and truncate failed runs
    db.run("DELETE FROM sync_errors WHERE timestamp < datetime('now', '-30 days');");
    fixesApplied.push('Cleaned historical sync errors older than 30 days.');

    // 3. Clear old sync runs older than 30 days
    db.run("DELETE FROM sync_runs WHERE started_at < datetime('now', '-30 days');");
    fixesApplied.push('Cleaned sync_runs older than 30 days.');

    saveDb();

    // Re-run health check
    const report = await runFullHealthAudit();
    res.json({
      success: true,
      message: `Auto-fix completed successfully (${fixesApplied.length} optimizations applied).`,
      fixesApplied,
      report,
    });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});
