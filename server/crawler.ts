import * as fs from 'fs';
import * as path from 'path';

export interface DriveItem {
  id: string;
  name: string;
  mimeType: string;
  isFolder: boolean;
  size?: number;
  modifiedTime?: string | null;
}

export interface DiscoveredFolder {
  id: string;
  name: string;
  path: string;
  parentId: string | null;
}

export interface DiscoveredFile {
  id: string;
  name: string;
  folderPath: string;
  folderId: string;
  mimeType: string;
  size?: number;
  modifiedTime?: string | null;
}

export class DriveCrawler {
  private cacheDir: string;

  constructor(cacheDir: string = './data/cache') {
    this.cacheDir = cacheDir;
    if (!fs.existsSync(this.cacheDir)) {
      fs.mkdirSync(this.cacheDir, { recursive: true });
    }
  }

  // Extract folder ID from various Google Drive URL formats or return raw ID
  static extractFolderId(input: string): string {
    const trimmed = input.trim();
    // Patterns like https://drive.google.com/drive/folders/1xFVOGEiQNshKBpA8tQoedy9FfPTP4rjn?usp=sharing
    const folderMatch = trimmed.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (folderMatch) return folderMatch[1];

    // Pattern like id=1xFVOGEiQNshKBpA8tQoedy9FfPTP4rjn
    const idParamMatch = trimmed.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (idParamMatch) return idParamMatch[1];

    // Direct folder ID (alphanumeric string with hyphens/underscores, 20-50 chars)
    if (/^[a-zA-Z0-9_-]{20,60}$/.test(trimmed)) {
      return trimmed;
    }

    throw new Error('Invalid Google Drive URL or Folder ID. Please provide a valid Google Drive folder link.');
  }

  // Fetch items inside a public shared Google Drive folder
  async fetchFolderItems(folderId: string): Promise<DriveItem[]> {
    const url = `https://drive.google.com/drive/folders/${folderId}?usp=sharing`;
    const res = await fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept-Language': 'en-US,en;q=0.9',
      },
      signal: AbortSignal.timeout(15000),
    });

    if (!res.ok) {
      throw new Error(`Failed to access Google Drive folder (${folderId}): HTTP ${res.status}`);
    }

    const html = await res.text();
    const match = html.match(/window\['_DRIVE_ivd'\]\s*=\s*'([^']+)'/);
    if (!match) {
      throw new Error(`Could not parse folder items for ID: ${folderId}. Ensure the folder is shared with "Anyone with the link can view".`);
    }

    const decoded = match[1].replace(/\\x([0-9A-Fa-f]{2})/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
    const json = JSON.parse(decoded);
    const items = json[0] || [];

    return items.map((item: any) => ({
      id: item[0],
      name: item[2],
      mimeType: item[3],
      isFolder: item[3] === 'application/vnd.google-apps.folder',
      size: typeof item[6] === 'number' ? item[6] : undefined,
      modifiedTime: item[10] ? new Date(item[10]).toISOString() : null,
    }));
  }

  // Recursively scan all folders and subfolders at any arbitrary nesting depth
  async scanRecursive(
    rootFolderId: string,
    onProgress?: (msg: string) => void
  ): Promise<{ folders: DiscoveredFolder[]; files: DiscoveredFile[] }> {
    const folders: DiscoveredFolder[] = [];
    const files: DiscoveredFile[] = [];
    const visitedFolderIds = new Set<string>();

    const traverse = async (
      folderId: string,
      currentPath: string,
      parentId: string | null,
      depth: number
    ): Promise<void> => {
      if (visitedFolderIds.has(folderId) || depth > 8) return;
      visitedFolderIds.add(folderId);

      if (onProgress) {
        onProgress(`Scanning folder ${currentPath || 'Root'}...`);
      }

      let items: DriveItem[] = [];
      try {
        items = await this.fetchFolderItems(folderId);
      } catch (err: any) {
        console.warn(`Failed to fetch items for folder ${folderId} (${currentPath}):`, err.message);
        return;
      }

      const subfolders: DriveItem[] = [];
      for (const item of items) {
        const itemCleanName = item.name.trim();
        const itemPath = currentPath ? `${currentPath} / ${itemCleanName}` : itemCleanName;

        if (item.isFolder) {
          folders.push({
            id: item.id,
            name: itemCleanName,
            path: itemPath,
            parentId: folderId,
          });
          subfolders.push(item);
        } else {
          files.push({
            id: item.id,
            name: itemCleanName,
            folderPath: currentPath || 'Root',
            folderId: folderId,
            mimeType: item.mimeType,
            size: item.size,
            modifiedTime: item.modifiedTime,
          });
        }
      }

      if (subfolders.length > 0) {
        await Promise.all(
          subfolders.map((sf) => {
            const subCleanName = sf.name.trim();
            const subPath = currentPath ? `${currentPath} / ${subCleanName}` : subCleanName;
            return traverse(sf.id, subPath, folderId, depth + 1);
          })
        );
      }
    };

    if (onProgress) onProgress('Scanning root Google Drive folder...');
    await traverse(rootFolderId, '', null, 0);

    return { folders, files };
  }

  // Find cached file with intelligent fuzzy matching
  findInCache(fileId: string, fileName: string): string | null {
    if (!fs.existsSync(this.cacheDir)) return null;
    const cachedFiles = fs.readdirSync(this.cacheDir);

    // 0. Canonical known file ID mapping
    const CANONICAL_MAP: Record<string, string> = {
      '17iVKXYw2Etso8yVHLH2buzhIPGtAraz5': 'BL.xlsx',
      '1Dm-3MzKF8jicncR-kC2wkt1yckNSEyaz': '1Dm-3MzKF8jicncR-kC2wkt1yckNSEyaz_ICICI_BANK_LOGIN_DATA_Updated.xlsx',
      '1o8EzNGtsPwZKyreFgcat3t2mkMtYnVglAFP2XjI7lzk': 'Consolidated Directories.xlsx',
      '13VezMqZ_Uo5sUnufTwMqwCGUrMSsnKmj': 'Consolidated Directories.xlsx',
      '1mLuNR_LkQM6o_tDqy0uJtof-sehnwP4QPgxKXIMH79A': '1EBB4HoEKMd6NpdtWa-bZNxZbCZfFC6zP_Jan.xlsx',
      '1EBB4HoEKMd6NpdtWa-bZNxZbCZfFC6zP': '1EBB4HoEKMd6NpdtWa-bZNxZbCZfFC6zP_Jan.xlsx',
      '1KdhvbpRdDZ9dG4JnWKPhzWKFMXdB9-c8ocCudTUrZuM': '1ojz_gJVcrolIYmhPN5mIvDG98d9S8FXv_MAY.xlsx',
      '1ojz_gJVcrolIYmhPN5mIvDG98d9S8FXv': '1ojz_gJVcrolIYmhPN5mIvDG98d9S8FXv_MAY.xlsx',
      '1Ww9d89THtVeZ1lepuvIrxxv4Jsqd3y74ADMBOP_cUt8': 'LOGIN.xlsx',
      '15_XL1hPhUcxZoF_QFp0-Q-iXZmPfsa5F': 'LOGIN.xlsx',
      '1N01z1GLjpvwvfR26mUuL6D54OOVD6ekrrT8ug7iUzRM': 'LOT.xlsx',
      '1RNpQjq3v9SAnop4eJ7fKwGwymiPEdcxS': 'LOT.xlsx',
      '1YlHFRFdO_dKHBDJPLNweOHclw69ruClBEkwZwOSvJDw': '1m8fHCPgc2bi9MFPQjogHIro2XXcRDvii_NEET.xlsx',
      '1m8fHCPgc2bi9MFPQjogHIro2XXcRDvii': '1m8fHCPgc2bi9MFPQjogHIro2XXcRDvii_NEET.xlsx',
      '1H9mPopKRUy-Y0H2fOhaNsVfekJPebODms_D9ypvEZuI': 'Fresh Data 21k.xlsx',
      '1KOC7vNp6BCciNBLbDdf8O9TpgxW-e9pl': 'Fresh Data 21k.xlsx',
      '1X32dBcaU1c9ZN761n1DpNjjYQeay0EdyOZwfHO6NOJ0': 'Fresh Data 30k.xlsx',
      '1mHD27YhMe9elz4vqMAwbaATyK-R27Xnk': 'Fresh Data 30k.xlsx',
      '1RdiB8h5tpOxLUDyunf4KKN-yx_MYY47bVgnaoYRhjTI': 'Tata Points Main.xlsx',
      '1q0p0My00kJD4MgDsQFS9YRJvdfXhfDFt': 'Tata Points Main.xlsx',
      '1bw2pXV6ID7yLi7pLUXxD-3kWqzZwB0bZ': '1bw2pXV6ID7yLi7pLUXxD-3kWqzZwB0bZ_TEAM_X.xlsx',
    };
    if (CANONICAL_MAP[fileId]) {
      const canonicalFile = path.join(this.cacheDir, CANONICAL_MAP[fileId]);
      if (fs.existsSync(canonicalFile) && fs.statSync(canonicalFile).size > 0) {
        return canonicalFile;
      }
    }

    // 1. Exact fileId prefix match: ${fileId}_*
    const byIdPrefix = cachedFiles.find(
      (f) => f.startsWith(`${fileId}_`) && fs.statSync(path.join(this.cacheDir, f)).size > 0
    );
    if (byIdPrefix) return path.join(this.cacheDir, byIdPrefix);

    // 2. Exact fileName match
    const exact = cachedFiles.find(
      (f) => f.toLowerCase() === fileName.toLowerCase() && fs.statSync(path.join(this.cacheDir, f)).size > 0
    );
    if (exact) return path.join(this.cacheDir, exact);

    // 3. Exact fileName with .xlsx
    const withExt = cachedFiles.find(
      (f) => f.toLowerCase() === `${fileName.toLowerCase()}.xlsx` && fs.statSync(path.join(this.cacheDir, f)).size > 0
    );
    if (withExt) return path.join(this.cacheDir, withExt);

    // 4. Normalized match (strip non-alphanumeric and compare)
    const normSearch = fileName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    for (const f of cachedFiles) {
      const normF = f.replace(/^[a-zA-Z0-9_-]{25,60}_/, '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
      if (normF.includes(normSearch) || normSearch.includes(normF)) {
        const fullP = path.join(this.cacheDir, f);
        if (fs.statSync(fullP).size > 0) return fullP;
      }
    }

    // 5. Special keyword matching for known workbooks
    const lowerName = fileName.toLowerCase();
    for (const f of cachedFiles) {
      const lowerF = f.toLowerCase();
      if (lowerName.includes('team x') && lowerF.includes('team_x')) return path.join(this.cacheDir, f);
      if ((lowerName === 'bl' || lowerName === 'bl.xlsx' || lowerName.endsWith('_bl.xlsx') || lowerName.endsWith('/bl.xlsx')) &&
          (lowerF === 'bl.xlsx' || lowerF.endsWith('_bl.xlsx') || lowerF.endsWith('/bl.xlsx'))) {
        return path.join(this.cacheDir, f);
      }
      if (lowerName.includes('icici') && lowerF.includes('icici')) return path.join(this.cacheDir, f);
      if (lowerName.includes('neet') && !lowerName.includes('support') && lowerF.includes('neet') && !lowerF.includes('sample')) {
        const fullP = path.join(this.cacheDir, f);
        if (fs.statSync(fullP).size > 10 * 1024 * 1024) return fullP;
      }
      if (lowerName.includes('tata') && lowerF.includes('tata')) return path.join(this.cacheDir, f);
      if (lowerName.includes('lot') && lowerF.includes('lot')) return path.join(this.cacheDir, f);
      if (lowerName.includes('login') && !lowerName.includes('icici') && lowerF.includes('login') && !lowerF.includes('icici')) return path.join(this.cacheDir, f);
      if (lowerName.includes('21k') && lowerF.includes('21k')) return path.join(this.cacheDir, f);
      if (lowerName.includes('30k') && lowerF.includes('30k')) return path.join(this.cacheDir, f);
      if (lowerName.startsWith('jan') && lowerF.includes('jan')) return path.join(this.cacheDir, f);
      if (lowerName.startsWith('may') && lowerF.includes('may')) return path.join(this.cacheDir, f);
    }

    return null;
  }

  // Download and cache an Excel file
  async downloadAndCacheFile(fileId: string, fileName: string, forceResync: boolean = false): Promise<{ filePath: string; size: number }> {
    // Check cache first so valid local workbooks are immediately available
    // (skipped when forceResync is true: the Drive file changed, so fetch a fresh copy;
    //  if that download fails we still fall back to the cached copy below)
    const cached = this.findInCache(fileId, fileName);
    if (!forceResync && cached && fs.existsSync(cached)) {
      const stat = fs.statSync(cached);
      if (stat.size > 0) {
        return { filePath: cached, size: stat.size };
      }
    }

    // Sanitize filename for safe local disk storage
    const safeBase = fileName.replace(/[^a-zA-Z0-9._-]/g, '_');
    const safeName = safeBase.toLowerCase().endsWith('.xlsx') || safeBase.toLowerCase().endsWith('.csv')
      ? `${fileId}_${safeBase}`
      : `${fileId}_${safeBase}.xlsx`;
    const targetPath = path.join(this.cacheDir, safeName);

    const downloadUrls = [
      `https://drive.usercontent.google.com/download?id=${fileId}&export=download&confirm=t`,
      `https://drive.google.com/uc?export=download&id=${fileId}`,
      `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`
    ];

    let lastError: any = null;
    for (const url of downloadUrls) {
      try {
        const res = await fetch(url, {
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          },
          signal: AbortSignal.timeout(20000),
        });
        if (res.ok) {
          const arrayBuf = await res.arrayBuffer();
          const buf = Buffer.from(arrayBuf);
          const headStr = buf.subarray(0, 64).toString('utf8').trim().toLowerCase();
          const isHtml = headStr.startsWith('<!doctype') || headStr.startsWith('<html');
          if (buf.length > 4 && !isHtml) {
            // Replace any older cached copy of the same file so it is not picked up instead
            if (cached && cached !== targetPath && fs.existsSync(cached)) { try { fs.unlinkSync(cached); } catch {} }
            fs.writeFileSync(targetPath, buf);
            return { filePath: targetPath, size: buf.length };
          }
        }
      } catch (err) {
        lastError = err;
      }
    }

    // If download failed but we have ANY matching cache file, use it as fallback
    const fallbackCached = this.findInCache(fileId, fileName);
    if (fallbackCached && fs.existsSync(fallbackCached) && fs.statSync(fallbackCached).size > 0) {
      return { filePath: fallbackCached, size: fs.statSync(fallbackCached).size };
    }

    throw new Error(`Failed to download ${fileName} (ID: ${fileId}): ${lastError?.message || 'Download failed'}`);
  }

  getCachedFilePath(fileId: string, fileName: string): string | null {
    return this.findInCache(fileId, fileName);
  }
}
