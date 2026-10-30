export interface ActiveSource {
  id: string;
  name: string;
  rootUrl: string;
  rootFolderId: string;
  updatedAt: string;
}

export interface SyncProgress {
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

export interface DashboardStats {
  foldersCount: number;
  filesCount: number;
  sheetsCount: number;
  totalRows: number;
  indexedRows: number;
  searchableRecords: number;
  uniquePhones: number;
  uniqueNames: number;
  duplicateGroups: number;
  errorsCount: number;
  totalSizeBytes: number;
  lastSync: {
    startedAt: string;
    completedAt: string | null;
    status: string;
    durationMs: number;
  } | null;
}

export interface DiscoveredFolderItem {
  id: string;
  name: string;
  path: string;
  parentId: string | null;
}

export interface DiscoveredFileItem {
  id: string;
  name: string;
  folderId: string;
  folderPath: string;
  size: number;
  modifiedTime: string | null;
  sheetsCount: number;
  totalRows: number;
  status: string;
  sheets: {
    name: string;
    rowsCount: number;
    colsCount: number;
    isIndexed: boolean;
    status: string;
  }[];
}

export interface SheetItem {
  id: string;
  fileId: string;
  fileName: string;
  folderPath: string;
  sheetName: string;
  rowsCount: number;
  colsCount: number;
  headers: string[];
  isIndexed: boolean;
  status: string;
}

export interface SheetPreviewRowItem {
  rowNum: number;
  id: number;
  values: Record<string, any>;
  priority: string;
  priorityCategory: string;
  priorityRank: number;
  priorityBadge: {
    bg: string;
    text: string;
    border: string;
  };
  occurrencesCount: number;
  matchingFiles: string[];
  matchingCategories?: string[];
}

export interface SearchMatchItem {
  id: number;
  sourceId: string;
  folderId: string;
  folderPath: string;
  fileId: string;
  fileName: string;
  sheetName: string;
  rowNum: number;
  customerName: string;
  phoneRaw: string;
  normPhone: string;
  pan: string;
  city: string;
  state: string;
  lender: string;
  loanAmount: string;
  statusVal: string;
  matchType: 'Exact Phone' | 'Exact Name' | 'Normalized Name' | 'Partial Name' | 'Smart Match';
  matchedColumn: string;
  matchedOriginalValue: string;
  rawRow: Record<string, any>;
  priority?: string;
  priorityCategory?: string;
  priorityRank?: number;
  priorityBadge?: {
    bg: string;
    text: string;
    border: string;
  };
}

export interface DuplicateStatsBreakdown {
  total: number;
  c2: number;
  c3: number;
  c4: number;
  c5: number;
  gt5: number;
}

export interface DuplicateFullStats {
  both: DuplicateStatsBreakdown;
  cross: DuplicateStatsBreakdown;
  within: DuplicateStatsBreakdown;
}

export interface DuplicateGroupItem {
  key: string;
  count: number;
  distinctFilesCount?: number;
  records: {
    id: number;
    customerName: string;
    phone: string;
    fileName: string;
    sheetName: string;
    rowNum: number;
    folderPath: string;
  }[];
}

export interface SyncRunItem {
  id: string;
  sourceId: string;
  startedAt: string;
  completedAt: string | null;
  durationMs: number;
  filesDiscovered: number;
  filesProcessed: number;
  sheetsProcessed: number;
  rowsProcessed: number;
  errorsCount: number;
  status: string;
}

export interface SyncErrorItem {
  id: number;
  sourceId: string;
  fileId: string | null;
  fileName: string | null;
  sheetName: string | null;
  errorMessage: string;
  timestamp: string;
}

export interface BatchLookupRecord {
  id: number;
  customerName: string;
  phone: string;
  fileName: string;
  sheetName: string;
  sheetLabel?: string;
  login?: string;
  lot?: string;
  neet?: string;
  priority?: string;
  rowNum: number;
  folderPath: string;
}

export interface BatchLookupResultItem {
  inputPhone: string;
  normPhone: string;
  found: boolean;
  priority?: string;
  priorityCategory?: string;
  priorityRank?: number;
  badgeBg?: string;
  badgeText?: string;
  badgeBorder?: string;
  sheetsName: string;
  login?: string;
  lot?: string;
  neet?: string;
  workbookNames: string;
  customerName: string;
  rowNumbers: string;
  matchCount: number;
  records: BatchLookupRecord[];
  originalData?: Record<string, any>;
}

export interface BatchLookupSummary {
  totalUploaded: number;
  totalFound: number;
  totalNotFound: number;
  durationMs: number;
  scope?: 'neet_login_lot' | 'all';
}

export interface LeadPriorityRule {
  id: string;
  rank: number;
  name: string;
  categoryKey: string;
  description: string;
  filePatterns: string[];
  folderPatterns: string[];
  neetYearOrder: 'desc' | 'asc';
  color: string;
  badgeBg: string;
  badgeText: string;
  badgeBorder: string;
  /** "Combine columns": values of these columns are added to the Priority label (e.g. NEET-2023). */
  combineLabel?: string;
  combineColumns?: string[];
  combineSeparator?: string;
  /** Added by the user in the Combine Columns screen. */
  isCustom?: boolean;
}

export interface CategoryTypeConfig {
  id: string;
  name: string;
  description?: string;
  fileOrFolderPattern: string;
  columnNames: string[];
  outputFormat: string;
  rank: number;
  color: string;
  badgeBg: string;
  badgeText: string;
  badgeBorder: string;
  isActive: boolean;
  isSystem?: boolean;
}

export interface MisDetectedSheet {
  sheetName: string;
  totalRows: number;
  detectedPhoneColIndex: number;
  detectedPhoneColName: string;
  headers: string[];
  previewRows: Record<string, any>[];
}

export interface MisEnrichedRow {
  sourceSheet: string;
  sourceRowNum: number;
  inputPhone: string;
  normPhone: string;
  found: boolean;
  priority: string;
  priorityCategory: string;
  priorityRank: number;
  badgeBg?: string;
  badgeText?: string;
  badgeBorder?: string;
  foundWorkbook: string;
  foundSheet: string;
  foundCustomerName: string;
  foundLender: string;
  foundLoanAmount: string;
  foundRowNum: string;
  matchCount: number;
  firstRecordId?: number;
  firstFileName?: string;
  firstSheetName?: string;
  originalData: Record<string, any>;
}

export interface MisSummary {
  totalRows: number;
  totalSheets: number;
  totalFound: number;
  totalNotFound: number;
  totalShortOrInvalid?: number;
  priorityBreakdown: Record<string, number>;
  durationMs: number;
}

export type AppMode = 'directory' | 'customer-lookup' | 'f2-tickets';

export interface LeadCheckItem {
  id?: number | string;
  adminName: string;
  timestamp: string;
}

export interface LeadRemarkItem {
  id: string;
  recordId: number;
  text: string;
  author: string;
  authorName: string;
  createdAt: string;
}

export interface DriveOccurrence {
  id: number;
  fileName: string;
  sheetName: string;
  rawSheetName?: string;
  customerName: string;
  phoneNumber?: string;
  rawRowJson?: Record<string, any>;
}

export interface CustomerLookupRecord {
  id: number;
  formattedId?: string;
  customerName: string;
  doctorName?: string;
  phoneNumber: string;
  panNumber: string;
  aadharNumber?: string;
  lenderName?: string;
  manager?: string;
  teamLeader?: string;
  loggedBy?: string;
  loginPoc?: string;
  loginDate?: string;
  dateLogged?: string;
  rowNum?: number;
  ticketId?: string;
  lanNo?: string;
  cibilScore?: string | number;
  loanAmount?: string | number;
  branch?: string;
  city?: string;
  state?: string;
  loanType?: string;
  loanCategory?: string;
  status?: string;
  approvalDate?: string;
  disbursedDate?: string;
  opsRemarks?: string;
  creditRemarks?: string;
  bankerName?: string;
  bankerPhone?: string;
  bankerEmail?: string;
  checkCount?: number;
  checkHistory?: LeadCheckItem[];
  remarks?: LeadRemarkItem[];
  fileName?: string;
  folderPath?: string;
  sheetName?: string;
  rawSheetName?: string;
  sourceId?: string;
  isDriveRecord?: boolean;
  driveOccurrences?: DriveOccurrence[];
  rawRowJson?: Record<string, any>;
  liveTicketMatch?: LiveTicketMatchInfo;
  misPriorityRank?: number;
  isHighestPriority?: boolean;
  chosenPriorityReason?: string;
  otherMisMatches?: {
    id: number;
    fileName: string;
    sheetName: string;
    priorityRank?: number;
    rowNum?: number;
    lenderName?: string;
    status?: string;
    loginDate?: string;
  }[];
}

export interface LiveTicketMatchInfo {
  matched: boolean;
  ticketId: string | number;
  formattedTicketId: string;
  customerName: string;
  customerContact: string;
  customerEmail?: string;
  customerLocation?: string;
  customerState?: string;
  applicationAmount?: string;
  applicationProvider?: string;
  standardizedProvider?: string;
  lookupLender?: string;
  loanCategory?: string;
  loanType?: string;
  ticketStatus?: string;
  loanStatus?: string;
  applicationDate?: string;
  createdAt?: string;
  matchScore: number;
  matchReasons: string[];
  matchedParameters: {
    phoneMatched: boolean;
    lenderMatched: boolean;
    nameMatched: boolean;
    ticketIdMatched: boolean;
    amountMatched: boolean;
  };
}

export interface VariantDetails {
  distinctPanCount: number;
  distinctAadharCount: number;
  isMultiRecord: boolean;
  driveOccurrences?: DriveOccurrence[];
}

export interface CustomerSearchResult {
  searchField: string;
  searchQuery: string;
  records: CustomerLookupRecord[];
  /** true when a phone/Aadhaar number was not found in any MIS sheet (lookup stops, even if it is in the directory or live tickets) */
  notInMis?: boolean;
  variantDetails?: VariantDetails;
}

export interface AuthUser {
  id: string;
  username: string;
  name: string;
  role: 'admin' | 'operator';
  status: 'active' | 'inactive';
  allowedTabs?: string[];
  mustChangePassword?: boolean;
}

export interface ConnectedSheetTabDetail {
  name: string;
  rows: number;
  phones?: number;
  newRows?: number;
  changedRows?: number;
}

export interface ConnectedSheetItem {
  id: string;
  name: string;
  url: string;
  spreadsheetId: string;
  sheetNames: string[];
  tabDetails?: ConnectedSheetTabDetail[];
  rowsCount: number;
  newRowsCount?: number;
  changedRowsCount?: number;
  status: 'active' | 'synced' | 'error' | 'pending' | 'restricted';
  lastSyncedAt: string | null;
  createdAt: string;
  lastError?: string | null;
  priorityRank?: number;
  isPermanent?: boolean;
  autoSync?: boolean;
}

export interface MisSyncLogItem {
  id: string;
  sheetId?: string;
  sheetName?: string;
  status: string;
  message: string;
  rowsAffected: number;
  timestamp: string;
}

export interface CachedSheetFile {
  fileName: string;
  displayName: string;
  sizeBytes: number;
  sizeFormatted: string;
  isMasterMis: boolean;
  isConnected: boolean;
  connectedSheetId: string | null;
  rowsCount: number | null;
  modifiedAt: string;
}

export interface MisSearchBookSheetInfo {
  name: string;
  hasPhoneColumn: boolean;
  phoneColumnName: string | null;
  phoneColIndex: number;
  phoneColLetter?: string;
  totalRows: number;
  totalCols?: number;
  samplePhones: string[];
}

export interface MisSearchBookRow {
  phone: string;
  normPhone: string;
  date?: string;
  priority: string;
  priorityRank: number;
  priorityCategory: string;
  badgeBg?: string;
  badgeText?: string;
  badgeBorder?: string;
  workbook: string;
  sheetName: string;
  customerName?: string;
  matchCount: number;
  found: boolean;
}

export interface MisSearchBookSummary {
  totalSheetsProcessed: number;
  totalPhonesCount: number;
  totalMatchesCount: number;
  totalNotFoundCount: number;
  durationMs: number;
  priorityBreakdown: Record<string, number>;
  perSheetStats: Record<string, { total: number; found: number; notFound: number }>;
}

export interface MisDashboardSheet {
  id: string;
  name: string;
  url: string;
  spreadsheetId: string;
  tabs: { name: string; rows: number; phones?: number }[];
  rowsCount: number;
  status: string;
  lastSyncedAt: string | null;
  priorityRank?: number;
}

export interface MisDashboardDateSummary {
  date: string; // ISO format: YYYY-MM-DD
  displayDate: string; // e.g. 10-Oct-2026
  rawSample: string; // e.g. 10-October-26
  count: number;
  tabCount?: Record<string, number>;
}

export interface MisDashboardEntry {
  id: number;
  /** Stable pointer to the row (workbook + tab + row). Use this to open the full row; `id` changes on every sync. */
  recordKey?: string;
  fileId?: string;
  rowNum: number;
  ticketId: string;
  customerName: string;
  phone: string;
  lender: string;
  loanAmount: string;
  netLoanAmount: string;
  status: string;
  caseCreatedDate: string;
  lastUpdatedDate: string;
  approvalDate: string;
  disbursedDate: string;
  tlName: string;
  manager: string;
  empName: string;
  city: string;
  state: string;
  branch: string;
  cibilScore: string;
  opsRemarks: string;
  creditRemarks: string;
  bankerName: string;
  bankerPhone: string;
  bankerEmail: string;
  rawDateValue: string;
  normalizedDate: string;
  matchedDateColumn: string;
  sheetName: string;
  fileName: string;
  rawRow: Record<string, any>;
  liveTicketMatch?: any;
}

export interface MisDashboardSummary {
  totalCases: number;
  statusBreakdown: Record<string, number>;
  lenderBreakdown: Record<string, number>;
  totalLoanAmount: number;
  totalNetLoanAmount: number;
  uniqueCustomers: number;
  uniquePhones: number;
}

export interface MisDashboardQueryResult {
  success: boolean;
  selectedDate: string;
  selectedDisplayDate: string;
  selectedDateColumn: string;
  selectedSheetId: string;
  selectedTab: string;
  totalCount: number;
  entries: MisDashboardEntry[];
  availableDates: MisDashboardDateSummary[];
  detectedDateColumns: string[];
  summary: MisDashboardSummary;
}



