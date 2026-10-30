/**
 * A stable way to point at one indexed row.
 *
 * search_index.id is NOT stable: every MIS sync deletes and re-inserts the rows, so the id the
 * dashboard received a minute ago may not exist any more ("Record not found").
 * (file_id, sheet_name, row_num) does not change between syncs, so it is used as the key instead.
 */
export interface RecordLocator {
  fileId: string;
  sheetName: string;
  rowNum: number;
}

export function makeRecordKey(loc: RecordLocator): string {
  return 'loc_' + Buffer.from(JSON.stringify([loc.fileId, loc.sheetName, loc.rowNum]), 'utf8').toString('base64url');
}

export function parseRecordKey(key: string): RecordLocator | null {
  if (!key.startsWith('loc_')) return null;
  try {
    const [fileId, sheetName, rowNum] = JSON.parse(Buffer.from(key.slice(4), 'base64url').toString('utf8'));
    if (typeof fileId !== 'string' || typeof sheetName !== 'string' || !Number.isFinite(Number(rowNum))) return null;
    return { fileId, sheetName, rowNum: Number(rowNum) };
  } catch {
    return null;
  }
}
