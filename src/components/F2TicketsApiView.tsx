import React, { useState, useEffect, useMemo } from 'react';
import {
  Search,
  RefreshCw,
  Phone,
  User,
  Hash,
  AlertCircle,
  Loader2,
  ExternalLink,
  X,
  Layers,
  Calendar,
  ArrowLeft,
  ArrowUpDown,
  FileText,
  FileSpreadsheet,
  Eye,
  Image as ImageIcon,
  FileCheck,
  Code,
  CheckCircle2,
  MessageSquare,
  Paperclip,
  Download,
  Zap,
  Settings2,
  Activity,
  Check,
  ChevronDown,
} from 'lucide-react';
import {
  standardizeProvider,
  STANDARDIZED_PROVIDERS_ORDER,
  StandardizedProviderCategory,
} from '../utils/standardizeProvider';
import { PdfDocumentViewer } from './PdfDocumentViewer';
import { OnlineSheetModal } from './OnlineSheetModal';

interface F2TicketsApiResponse {
  success: boolean;
  searchQuery?: string;
  searchField?: string;
  fetchedAt: string | null;
  pagesFetched: number;
  duplicatesRemoved?: number;
  apiReportedCount: number | null;
  apiReportedPages: number | null;
  statusCode?: number | null;
  statusMessage: string | null;
  stopReason?: string | null;
  appliedStartDate?: string;
  appliedEndDate?: string;
  totalInMemory: number;
  matchedCount?: number;
  records: Record<string, any>[];
  error?: string;
}

interface F2TicketDetailApiResponse {
  success: boolean;
  cachedInMemory?: boolean;
  ticketId: string;
  fetchedAt?: string;
  endpointUrl?: string;
  httpStatus?: number;
  statusCode?: number;
  message?: string;
  data?: Record<string, any>;
  documentsCount?: number;
  error?: string;
}

interface F2TicketActivitiesApiResponse {
  success: boolean;
  ticketId: string;
  fetchedAt?: string;
  endpointUrl?: string;
  httpStatus?: number;
  statusCode?: number;
  message?: string;
  activitiesCount?: number;
  activities?: Record<string, any>[];
  error?: string;
}

interface NormalizedDocumentItem {
  index: number;
  id: string | number | null;
  type: string;
  url: string | null;
  fileName: string;
  extension: string;
  isImage: boolean;
  isPdf: boolean;
  isSheet: boolean;
  rawObject: Record<string, any>;
}

function toTitleCase(str: any): string {
  if (str === null || str === undefined) return '';
  const cleaned = String(str).trim().replace(/\s+/g, ' ');
  if (!cleaned) return '';
  return cleaned
    .toLowerCase()
    .split(' ')
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}

function capitalizeFirstChar(str: any): string {
  if (str === null || str === undefined) return '';
  const s = String(str).trim();
  if (!s) return '';
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function formatIndianAmountWithDecimals(val: any): string {
  if (val === null || val === undefined || val === '') return '—';
  const num = Number(String(val).replace(/[^0-9.-]/g, ''));
  if (isNaN(num)) return String(val);
  return `₹${num.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

function formatIndianAmountNoDecimals(val: any): string {
  if (val === null || val === undefined || val === '') return '—';
  const num = Number(String(val).replace(/[^0-9.-]/g, ''));
  if (isNaN(num)) return String(val);
  return `₹ ${num.toLocaleString('en-IN', {
    maximumFractionDigits: 0,
  })}`;
}

function formatCreatedAtTableDate(val: any): string {
  if (!val) return '—';
  const d = new Date(val);
  if (isNaN(d.getTime())) return String(val);
  const day = String(d.getUTCDate()).padStart(2, '0');
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const month = months[d.getUTCMonth()];
  const year = d.getUTCFullYear();
  return `${day} ${month} ${year}`;
}

function formatActivityDateTime(val: any): string {
  if (!val) return '—';
  const d = new Date(val);
  if (isNaN(d.getTime())) return String(val);
  try {
    // Format in Asia/Kolkata (IST) to match F2 Fintech portal ("Oct 01, 2026 at 01:19 PM")
    const parts = new Intl.DateTimeFormat('en-US', {
      timeZone: 'Asia/Kolkata',
      month: 'short',
      day: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true,
    }).formatToParts(d);

    const map: Record<string, string> = {};
    for (const p of parts) {
      map[p.type] = p.value;
    }
    return `${map.month} ${map.day}, ${map.year} at ${map.hour}:${map.minute} ${map.dayPeriod?.toUpperCase() || ''}`.trim();
  } catch {
    return d.toLocaleString();
  }
}

function formatApplicationDateMMDDYYYY(appDate: any, createdAt: any): string {
  const raw = appDate || createdAt;
  if (!raw) return '—';
  const str = String(raw).trim();
  const ymdMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (ymdMatch) {
    return `${ymdMatch[2]}/${ymdMatch[3]}/${ymdMatch[1]}`;
  }
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(d.getUTCDate()).padStart(2, '0');
    const yyyy = d.getUTCFullYear();
    return `${mm}/${dd}/${yyyy}`;
  }
  return str;
}

function formatDesignation(designation: any, loanType: any): string {
  if (
    designation &&
    String(designation).trim() &&
    String(designation).trim().toLowerCase() !== 'n/a'
  ) {
    return toTitleCase(designation);
  }
  if (!loanType) return '—';
  const s = String(loanType).trim().toLowerCase();
  if (s.includes('professional')) return 'Professional';
  if (s.includes('doctor')) return 'Doctor';
  if (s.includes('business')) return 'Business';
  if (s.includes('personal')) return 'Personal';
  if (s.includes('home')) return 'Home Loan';
  return toTitleCase(loanType);
}

function formatContactWithPrefix(contact: any): string {
  if (contact === null || contact === undefined || contact === '') return '—';
  const clean = String(contact).trim();
  const digits = clean.replace(/\D/g, '');
  if (digits.length === 10) return `+91 ${digits}`;
  if (digits.length === 12 && digits.startsWith('91')) return `+91 ${digits.slice(2)}`;
  return clean.startsWith('+') ? clean : `+91 ${clean}`;
}

function extractRecordYMD(rec: Record<string, any>): string | null {
  const rawCreated = rec.createdAt ? String(rec.createdAt).trim() : '';
  if (/^\d{4}-\d{2}-\d{2}/.test(rawCreated)) {
    return rawCreated.slice(0, 10);
  }
  const rawAppDate = rec.applicationDate ? String(rec.applicationDate).trim() : '';
  if (/^\d{4}-\d{2}-\d{2}/.test(rawAppDate)) {
    return rawAppDate.slice(0, 10);
  }
  return null;
}

function extractRecordTimestampMs(rec: Record<string, any>): number {
  if (rec.createdAt) {
    const t = Date.parse(String(rec.createdAt));
    if (!isNaN(t)) return t;
  }
  if (rec.applicationDate) {
    const t = Date.parse(String(rec.applicationDate));
    if (!isNaN(t)) return t;
  }
  return 0;
}

function normalizeUrlToDocumentItem(
  rawUrl: string,
  index: number,
  id: string | number | null,
  typeLabel: string,
  rawObj: Record<string, any>
): NormalizedDocumentItem {
  const url = rawUrl && String(rawUrl).trim() ? String(rawUrl).trim() : null;
  const fileName = url
    ? decodeURIComponent(url.split('/').pop()?.split('?')[0] || `document_${index}`)
    : `document_${index}`;
  const ext = fileName.includes('.') ? fileName.split('.').pop()!.toLowerCase() : '';
  const sheetExts = ['xlsx', 'xls', 'csv', 'tsv', 'ods', 'xlsm', 'xlsb'];
  const isSheet =
    sheetExts.includes(ext) ||
    /\.(xlsx|xls|csv|tsv|ods)(\?|$)/i.test(url || '') ||
    /spreadsheet/i.test(url || '') ||
    /spreadsheet/i.test(typeLabel);

  return {
    index,
    id,
    type: typeLabel,
    url,
    fileName,
    extension: ext.toUpperCase() || (isSheet ? 'XLSX' : 'FILE'),
    isImage: ['jpg', 'jpeg', 'png', 'webp', 'gif', 'bmp'].includes(ext),
    isPdf: ext === 'pdf',
    isSheet,
    rawObject: rawObj,
  };
}

function extractDocumentsFromDetailData(detail: Record<string, any>): NormalizedDocumentItem[] {
  const docs: NormalizedDocumentItem[] = [];
  const rawDocs = Array.isArray(detail?.customerDocuments) ? detail.customerDocuments : [];

  rawDocs.forEach((item: any, idx: number) => {
    if (!item) return;

    if (typeof item === 'string') {
      docs.push(
        normalizeUrlToDocumentItem(item, idx + 1, null, `Document ${idx + 1}`, {
          document_url: item,
        })
      );
      return;
    }

    if (typeof item === 'object') {
      const rawUrl =
        item.document_url ??
        item.documentUrl ??
        item.file_url ??
        item.fileUrl ??
        item.url ??
        null;
      const rawType =
        item.type ??
        item.document_type ??
        item.documentType ??
        item.category ??
        item.document_name ??
        item.name ??
        `Document ${idx + 1}`;
      const docItem = normalizeUrlToDocumentItem(
        rawUrl ? String(rawUrl) : '',
        idx + 1,
        item.id ?? item.documentId ?? null,
        String(rawType),
        item
      );
      if (item.file_name || item.fileName) {
        docItem.fileName = String(item.file_name || item.fileName);
      }
      docs.push(docItem);
    }
  });

  if (detail?.voiceNoteUrl && String(detail.voiceNoteUrl).trim()) {
    docs.push(
      normalizeUrlToDocumentItem(
        String(detail.voiceNoteUrl),
        docs.length + 1,
        'voiceNoteUrl',
        'voice note',
        { type: 'voiceNoteUrl', document_url: detail.voiceNoteUrl }
      )
    );
  }

  return docs;
}

function extractAttachmentsFromActivity(
  activity: Record<string, any>
): NormalizedDocumentItem[] {
  const items: NormalizedDocumentItem[] = [];
  if (!activity || typeof activity !== 'object') return items;

  const rawAttachment =
    activity.attachment ??
    activity.attachments ??
    activity.document_url ??
    activity.documentUrl ??
    activity.file_url ??
    null;

  if (!rawAttachment) return items;

  if (typeof rawAttachment === 'string') {
    const trimmed = rawAttachment.trim();
    if (trimmed && trimmed.toLowerCase() !== 'null') {
      items.push(
        normalizeUrlToDocumentItem(
          trimmed,
          1,
          activity.id ?? null,
          activity.comment ? String(activity.comment).trim().slice(0, 40) : 'Remark Attachment',
          activity
        )
      );
    }
  } else if (Array.isArray(rawAttachment)) {
    rawAttachment.forEach((entry, i) => {
      if (typeof entry === 'string' && entry.trim()) {
        items.push(
          normalizeUrlToDocumentItem(
            entry.trim(),
            i + 1,
            activity.id ?? null,
            `Remark Attachment ${i + 1}`,
            activity
          )
        );
      } else if (entry && typeof entry === 'object') {
        const u = entry.document_url || entry.url || entry.attachment || '';
        if (u) {
          items.push(
            normalizeUrlToDocumentItem(
              String(u),
              i + 1,
              entry.id ?? activity.id ?? null,
              String(entry.type || entry.name || `Remark Attachment ${i + 1}`),
              entry
            )
          );
        }
      }
    });
  }

  return items;
}

export interface F2TicketsApiViewProps {
  initialTicketId?: string | number | null;
  onBackToLookup?: () => void;
  onClearInitialTicketId?: () => void;
}

export const F2TicketsApiView: React.FC<F2TicketsApiViewProps> = ({
  initialTicketId,
  onBackToLookup,
  onClearInitialTicketId,
}) => {
  const [records, setRecords] = useState<Record<string, any>[]>([]);
  const [totalInMemory, setTotalInMemory] = useState<number>(0);
  const [fetchedAt, setFetchedAt] = useState<string | null>(null);
  const [appliedStartDate, setAppliedStartDate] = useState<string>('2000-01-01');
  const [appliedEndDate, setAppliedEndDate] = useState<string>('2099-12-31');

  // Preserve the full list of years & latest month discovered in the dataset
  const [discoveredYears, setDiscoveredYears] = useState<string[]>([]);
  const [latestMonthInfo, setLatestMonthInfo] = useState<{
    label: string;
    start: string;
    end: string;
  }>({
    label: 'Oct 2026',
    start: '2026-10-01',
    end: '2026-10-31',
  });

  // Selected Ticket & Live Detail API State
  const [selectedTicket, setSelectedTicket] = useState<Record<string, any> | null>(null);
  const [ticketDetailResponse, setTicketDetailResponse] =
    useState<F2TicketDetailApiResponse | null>(null);
  const [isDetailLoading, setIsDetailLoading] = useState<boolean>(false);
  const [detailError, setDetailError] = useState<string | null>(null);
  const [showRawDetailJson, setShowRawDetailJson] = useState<boolean>(false);

  // Live Ticket Activities / Remarks API State
  const [ticketActivitiesResponse, setTicketActivitiesResponse] =
    useState<F2TicketActivitiesApiResponse | null>(null);
  const [isActivitiesLoading, setIsActivitiesLoading] = useState<boolean>(false);
  const [activitiesError, setActivitiesError] = useState<string | null>(null);
  const [activitiesSortOrder, setActivitiesSortOrder] = useState<'desc' | 'asc'>('desc');

  // In-app Document Preview Modal State
  const [previewDoc, setPreviewDoc] = useState<NormalizedDocumentItem | null>(null);
  const [failedImageMap, setFailedImageMap] = useState<Record<string, boolean>>({});

  // Online Sheet Viewer Modal State
  const [sheetPreviewTarget, setSheetPreviewTarget] = useState<{
    url: string;
    fileName?: string;
    ticketId?: string | number | null;
    activityAuthor?: string | null;
    comment?: string | null;
  } | null>(null);

  // Search, Date Limiter, Category/Provider Filter & Sort State
  const [searchField, setSearchField] = useState<'all' | 'phone' | 'name' | 'ticketId'>('all');
  const [query, setQuery] = useState<string>('');
  const [startDateInput, setStartDateInput] = useState<string>('');
  const [endDateInput, setEndDateInput] = useState<string>('');
  const [dateSortOrder, setDateSortOrder] = useState<'desc' | 'asc'>('desc');

  // Loan Category/Type & Loan Provider Filter State ('all' resets back to all)
  const [selectedLoanCategory, setSelectedLoanCategory] = useState<string>('all');
  const [selectedLoanProvider, setSelectedLoanProvider] = useState<string>('all');
  const [discoveredCategories, setDiscoveredCategories] = useState<
    { key: string; label: string; group: 'Category' | 'Loan Type'; count: number }[]
  >([]);
  const [discoveredProviders, setDiscoveredProviders] = useState<
    { key: string; label: string; count: number }[]
  >([]);
  const [draggedFilterItem, setDraggedFilterItem] = useState<{
    kind: 'category' | 'provider';
    value: string;
  } | null>(null);

  // Drag-and-Drop Table Columns Order
  const DEFAULT_TABLE_COLUMNS = [
    'tid',
    'source',
    'name',
    'email',
    'amount',
    'provider',
    'loanCategory',
    'status',
    'loanType',
    'leadType',
    'tenure',
    'location',
    'createdAt',
    'actions',
  ] as const;
  const [columnOrder, setColumnOrder] = useState<string[]>([...DEFAULT_TABLE_COLUMNS]);
  const [draggedColumnKey, setDraggedColumnKey] = useState<string | null>(null);

  // Entries per view option: default 15, with options for 50 or 100 at once
  const [batchSize, setBatchSize] = useState<15 | 50 | 100>(15);
  const [visibleRowsCount, setVisibleRowsCount] = useState<number>(15);

  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isRefreshing, setIsRefreshing] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  // Live Auto-Update State
  const [autoUpdateEnabled, setAutoUpdateEnabled] = useState<boolean>(true);
  const [autoUpdateInterval, setAutoUpdateInterval] = useState<number>(60);
  const [isAutoChecking, setIsAutoChecking] = useState<boolean>(false);
  const [secondsUntilNextCheck, setSecondsUntilNextCheck] = useState<number>(60);
  const [lastAutoUpdateNotice, setLastAutoUpdateNotice] = useState<string | null>(null);
  const [showAutoSettings, setShowAutoSettings] = useState<boolean>(false);
  const prevSyncVersionRef = React.useRef<number | null>(null);

  const updateDiscoveredYearsAndMonth = (incomingRecords: Record<string, any>[]) => {
    if (!incomingRecords || incomingRecords.length === 0) return;

    const yearsSet = new Set<string>(discoveredYears);
    let maxYmd = '';

    const catCountMap = new Map<string, { label: string; group: 'Category' | 'Loan Type'; count: number }>();
    const provCountMap = new Map<string, { label: string; count: number }>();

    for (const rec of incomingRecords) {
      const ymd = extractRecordYMD(rec);
      if (ymd && /^\d{4}-\d{2}-\d{2}$/.test(ymd)) {
        const yr = ymd.slice(0, 4);
        const yrNum = Number(yr);
        if (yrNum >= 2000 && yrNum <= 2099) {
          yearsSet.add(yr);
          if (ymd > maxYmd) {
            maxYmd = ymd;
          }
        }
      }

      // Collect all Loan Category values (e.g. unsecured, secured)
      const rawCat = rec.loanCategory !== undefined && rec.loanCategory !== null
        ? String(rec.loanCategory).trim()
        : '';
      if (rawCat) {
        const normKey = `cat:${rawCat.toLowerCase()}`;
        const existing = catCountMap.get(normKey);
        if (existing) {
          existing.count += 1;
        } else {
          catCountMap.set(normKey, {
            label: `${toTitleCase(rawCat)} (Category)`,
            group: 'Category',
            count: 1,
          });
        }
      }

      // Collect all Loan Type values (e.g. professional loan, personal loan, business loan, home loan, lap, etc.)
      const rawType = rec.loanType !== undefined && rec.loanType !== null
        ? String(rec.loanType).trim()
        : '';
      if (rawType) {
        const normKey = `type:${rawType.toLowerCase()}`;
        const existing = catCountMap.get(normKey);
        if (existing) {
          existing.count += 1;
        } else {
          catCountMap.set(normKey, {
            label: toTitleCase(rawType),
            group: 'Loan Type',
            count: 1,
          });
        }
      }

      // Collect all Loan Provider values using the standardized 33 categories + Other/Unmapped
      const stdProv =
        rec.standardizedProvider ||
        standardizeProvider(rec.applicationProvider ?? rec.provider);
      rec.standardizedProvider = stdProv;
      provCountMap.set(stdProv, {
        label: stdProv,
        count: (provCountMap.get(stdProv)?.count || 0) + 1,
      });
    }

    const sortedYears = Array.from(yearsSet).sort((a, b) => Number(b) - Number(a));
    if (sortedYears.length > 0) {
      setDiscoveredYears(sortedYears);
    }

    if (catCountMap.size > 0) {
      const catList = Array.from(catCountMap.entries())
        .map(([key, val]) => ({ key, ...val }))
        .sort((a, b) => {
          if (a.group !== b.group) return a.group === 'Category' ? -1 : 1;
          return b.count - a.count;
        });
      setDiscoveredCategories((prev) => (catList.length >= prev.length ? catList : prev));
    }

    // Build standardized providers list ensuring every standardized category is represented with exact count
    const provList = STANDARDIZED_PROVIDERS_ORDER.map((name) => {
      const entry = provCountMap.get(name);
      return {
        key: name,
        label: name,
        count: entry ? entry.count : 0,
      };
    }).sort((a, b) => b.count - a.count);

    setDiscoveredProviders(provList);

    if (maxYmd) {
      const [yrStr, moStr] = maxYmd.split('-');
      const yr = Number(yrStr);
      const mo = Number(moStr);
      const monthNames = [
        'Jan',
        'Feb',
        'Mar',
        'Apr',
        'May',
        'Jun',
        'Jul',
        'Aug',
        'Sep',
        'Oct',
        'Nov',
        'Dec',
      ];
      const lastDay = new Date(Date.UTC(yr, mo, 0)).getUTCDate();
      setLatestMonthInfo({
        label: `${monthNames[mo - 1]} ${yr}`,
        start: `${yrStr}-${moStr}-01`,
        end: `${yrStr}-${moStr}-${String(lastDay).padStart(2, '0')}`,
      });
    }
  };

  const applyResponseData = (data: F2TicketsApiResponse) => {
    const incoming = data.records || [];
    setRecords(incoming);
    setTotalInMemory(data.totalInMemory ?? incoming.length);
    setFetchedAt(data.fetchedAt || null);
    if (data.appliedStartDate) setAppliedStartDate(data.appliedStartDate);
    if (data.appliedEndDate) setAppliedEndDate(data.appliedEndDate);
    updateDiscoveredYearsAndMonth(incoming);
  };

  const loadTicketsFromBackend = async (
    q: string = query,
    field: 'all' | 'phone' | 'name' | 'ticketId' = searchField,
    forceRefresh: boolean = false,
    overrideStartDate?: string,
    overrideEndDate?: string
  ) => {
    const targetStart =
      overrideStartDate !== undefined
        ? overrideStartDate
        : startDateInput.trim() || '2000-01-01';
    const targetEnd =
      overrideEndDate !== undefined
        ? overrideEndDate
        : endDateInput.trim() || '2099-12-31';

    if (forceRefresh) {
      setIsRefreshing(true);
    } else {
      setIsLoading(true);
    }
    setError(null);

    try {
      if (forceRefresh) {
        const res = await fetch('/api/f2-tickets/fetch', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Accept: 'application/json',
          },
          body: JSON.stringify({
            startDate: targetStart,
            endDate: targetEnd,
          }),
        });
        const data: F2TicketsApiResponse = await res.json();
        if (!res.ok || !data.success) {
          throw new Error(data.error || `API fetch failed with status ${res.status}`);
        }
        applyResponseData(data);
        setVisibleRowsCount(batchSize);
        return;
      }

      const params = new URLSearchParams({
        field,
        q: q.trim(),
      });
      if (startDateInput.trim()) params.set('startDate', startDateInput.trim());
      if (endDateInput.trim()) params.set('endDate', endDateInput.trim());

      const res = await fetch(`/api/f2-tickets?${params.toString()}`, {
        method: 'GET',
        headers: { Accept: 'application/json' },
      });

      const data: F2TicketsApiResponse = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || `Request failed with status ${res.status}`);
      }

      applyResponseData(data);
      setVisibleRowsCount(batchSize);
    } catch (err: any) {
      setError(err.message || 'Failed to fetch tickets from F2 Fintech API');
    } finally {
      setIsLoading(false);
      setIsRefreshing(false);
    }
  };

  // Fetch live ticket activities & remarks from /api/f2-tickets/activities/{TICKET_ID}
  const fetchTicketActivitiesById = async (ticketIdVal: string | number) => {
    const cleanId = String(ticketIdVal)
      .replace(/^(f2fin-|tck-|#|ticket\s*)/i, '')
      .trim();
    if (!cleanId) return;

    setIsActivitiesLoading(true);
    setActivitiesError(null);

    try {
      const res = await fetch(`/api/f2-tickets/activities/${encodeURIComponent(cleanId)}`, {
        method: 'GET',
        headers: { Accept: 'application/json' },
      });
      const data: F2TicketActivitiesApiResponse = await res.json();

      if (!res.ok || !data.success) {
        throw new Error(
          data.error || `Activities API returned status ${res.status} for Ticket ID ${cleanId}`
        );
      }

      setTicketActivitiesResponse(data);
    } catch (err: any) {
      setActivitiesError(
        err.message || `Could not retrieve ticket activities for Ticket ID ${cleanId}`
      );
    } finally {
      setIsActivitiesLoading(false);
    }
  };

  // Fetch complete ticket details + customerDocuments AND ticket activities/remarks
  const openTicketDetailAndActivities = async (
    ticketIdVal: string | number,
    baseRecord?: Record<string, any>,
    forceRefresh: boolean = false
  ) => {
    const cleanId = String(ticketIdVal)
      .replace(/^(f2fin-|tck-|#|ticket\s*)/i, '')
      .trim();
    if (!cleanId) return;

    setSelectedTicket(baseRecord || { ticketId: cleanId });
    setIsDetailLoading(true);
    setDetailError(null);
    if (!forceRefresh) {
      setTicketDetailResponse(null);
      setTicketActivitiesResponse(null);
    }

    // Trigger activities fetch concurrently so remarks load without blocking or breaking details
    fetchTicketActivitiesById(cleanId);

    try {
      const url = `/api/f2-tickets/detail/${encodeURIComponent(cleanId)}${
        forceRefresh ? '?refresh=true' : ''
      }`;
      const res = await fetch(url, {
        method: 'GET',
        headers: { Accept: 'application/json' },
      });
      const data: F2TicketDetailApiResponse = await res.json();

      if (!res.ok || !data.success) {
        throw new Error(
          data.error || `Detail API returned status ${res.status} for Ticket ID ${cleanId}`
        );
      }

      setTicketDetailResponse(data);
      if (data.data && typeof data.data === 'object') {
        setSelectedTicket((prev) => ({
          ...(prev || {}),
          ...data.data,
        }));
      }
    } catch (err: any) {
      setDetailError(
        err.message || `Could not retrieve ticket details for Ticket ID ${cleanId}`
      );
    } finally {
      setIsDetailLoading(false);
    }
  };

  useEffect(() => {
    loadTicketsFromBackend('', 'all', false);
  }, []);

  useEffect(() => {
    if (initialTicketId) {
      const cleanId = String(initialTicketId)
        .replace(/^(f2fin-|tck-|#|ticket\s*)/i, '')
        .trim();
      if (cleanId) {
        setSearchField('ticketId');
        setQuery(cleanId);
        openTicketDetailAndActivities(cleanId);
      }
    }
  }, [initialTicketId]);

  useEffect(() => {
    setVisibleRowsCount(batchSize);
  }, [
    query,
    searchField,
    startDateInput,
    endDateInput,
    dateSortOrder,
    batchSize,
    selectedLoanCategory,
    selectedLoanProvider,
  ]);

  // Real-time Auto-Update polling hook: runs in background and syncs seamlessly
  useEffect(() => {
    let timer: any = null;
    if (autoUpdateEnabled) {
      timer = setInterval(() => {
        setSecondsUntilNextCheck((prev) => {
          if (prev <= 1) {
            triggerBackgroundAutoCheck();
            return autoUpdateInterval;
          }
          return prev - 1;
        });
      }, 1000);
    }
    return () => {
      if (timer) clearInterval(timer);
    };
  }, [autoUpdateEnabled, autoUpdateInterval, query, searchField]);

  const triggerBackgroundAutoCheck = async () => {
    if (isAutoChecking || isLoading || isRefreshing) return;
    setIsAutoChecking(true);
    try {
      const res = await fetch('/api/f2-tickets/auto-update-status');
      const data = await res.json();
      if (data && data.success) {
        const remoteVersion = data.syncVersion;
        if (prevSyncVersionRef.current === null) {
          prevSyncVersionRef.current = remoteVersion;
        } else if (remoteVersion > prevSyncVersionRef.current) {
          prevSyncVersionRef.current = remoteVersion;
          // New tickets were detected by the background worker! Reload silently
          await loadTicketsFromBackend(query, searchField, false);
          setLastAutoUpdateNotice(`Auto-updated: ${data.totalInMemory?.toLocaleString()} live tickets synced`);
          setTimeout(() => setLastAutoUpdateNotice(null), 4000);
        }
      }
    } catch {
      // Ignore background transient network blips
    } finally {
      setIsAutoChecking(false);
    }
  };

  const handleToggleAutoUpdate = async (enabled: boolean) => {
    setAutoUpdateEnabled(enabled);
    if (enabled) {
      setSecondsUntilNextCheck(autoUpdateInterval);
    }
    try {
      await fetch('/api/f2-tickets/auto-update-config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ enabled, intervalSeconds: autoUpdateInterval }),
      });
    } catch {}
  };

  const handleChangeInterval = async (sec: number) => {
    setAutoUpdateInterval(sec);
    setSecondsUntilNextCheck(sec);
    setShowAutoSettings(false);
    try {
      await fetch('/api/f2-tickets/auto-update-config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ intervalSeconds: sec }),
      });
    } catch {}
  };

  const handleManualAutoCheck = async () => {
    setIsAutoChecking(true);
    try {
      const res = await fetch('/api/f2-tickets/auto-update-config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ checkNow: true }),
      });
      const data = await res.json();
      if (data?.checkResult?.updated) {
        setLastAutoUpdateNotice(`Auto-updated: ${data.checkResult.newTicketsCount} new ticket(s) retrieved!`);
      } else {
        setLastAutoUpdateNotice('Auto-update checked: Already up to date with live OMS.');
      }
      setTimeout(() => setLastAutoUpdateNotice(null), 4000);
      await loadTicketsFromBackend(query, searchField, false);
    } catch {} finally {
      setIsAutoChecking(false);
      setSecondsUntilNextCheck(autoUpdateInterval);
    }
  };

  // Instant filtering & date sorting over the complete fetched dataset in memory
  const filteredRecords = useMemo(() => {
    const trimmed = query.trim();
    const cleanStart = startDateInput.trim();
    const cleanEnd = endDateInput.trim();
    const hasDateFilter = Boolean(cleanStart || cleanEnd);

    const queryLower = trimmed.toLowerCase().replace(/\s+/g, ' ');
    const queryDigits = trimmed.replace(/\D/g, '');
    const queryLast10 = queryDigits.length >= 10 ? queryDigits.slice(-10) : queryDigits;
    const ticketIdClean = trimmed.replace(/^(f2fin-|tck-|#|ticket\s*)/i, '').trim().toLowerCase();

    const matched = records.filter((rec) => {
      if (hasDateFilter) {
        const recYmd = extractRecordYMD(rec);
        if (recYmd) {
          if (cleanStart && recYmd < cleanStart) return false;
          if (cleanEnd && recYmd > cleanEnd) return false;
        }
      }

      // Filter by Loan Category / Loan Type
      if (selectedLoanCategory && selectedLoanCategory !== 'all') {
        const recCat =
          rec.loanCategory !== undefined && rec.loanCategory !== null
            ? String(rec.loanCategory).trim().toLowerCase()
            : '';
        const recType =
          rec.loanType !== undefined && rec.loanType !== null
            ? String(rec.loanType).trim().toLowerCase()
            : '';
        if (selectedLoanCategory.startsWith('cat:')) {
          const targetCat = selectedLoanCategory.slice(4);
          if (recCat !== targetCat) return false;
        } else if (selectedLoanCategory.startsWith('type:')) {
          const targetType = selectedLoanCategory.slice(5);
          if (recType !== targetType) return false;
        } else {
          const target = selectedLoanCategory.toLowerCase();
          if (recCat !== target && recType !== target) return false;
        }
      }

      // Filter by Standardized Loan Provider
      if (selectedLoanProvider && selectedLoanProvider !== 'all') {
        const stdProv =
          rec.standardizedProvider ||
          standardizeProvider(rec.applicationProvider ?? rec.provider);
        if (stdProv !== selectedLoanProvider) return false;
      }

      if (!trimmed) return true;

      const rawTicketId =
        rec.ticketId !== undefined && rec.ticketId !== null ? String(rec.ticketId).trim() : '';
      const rawCustomerName =
        rec.customerName !== undefined && rec.customerName !== null
          ? String(rec.customerName).trim()
          : '';
      const rawCustomerContact =
        rec.customerContact !== undefined && rec.customerContact !== null
          ? String(rec.customerContact).trim()
          : '';

      const normName = rawCustomerName.toLowerCase().replace(/\s+/g, ' ');
      const contactDigits = rawCustomerContact.replace(/\D/g, '');
      const contactLast10 = contactDigits.length >= 10 ? contactDigits.slice(-10) : contactDigits;

      if (searchField === 'phone') {
        if (queryDigits) {
          return (
            contactDigits.includes(queryDigits) ||
            (queryLast10.length === 10 && contactLast10 === queryLast10) ||
            rawCustomerContact.toLowerCase().includes(queryLower)
          );
        }
        return rawCustomerContact.toLowerCase().includes(queryLower);
      }

      if (searchField === 'name') {
        return (
          normName.includes(queryLower) ||
          rawCustomerName.toLowerCase().includes(trimmed.toLowerCase())
        );
      }

      if (searchField === 'ticketId') {
        return (
          rawTicketId.toLowerCase() === ticketIdClean ||
          rawTicketId.toLowerCase().includes(ticketIdClean) ||
          (queryDigits.length > 0 && rawTicketId.includes(queryDigits))
        );
      }

      const matchesPhone =
        (queryDigits.length > 0 &&
          (contactDigits.includes(queryDigits) ||
            (queryLast10.length === 10 && contactLast10 === queryLast10))) ||
        rawCustomerContact.toLowerCase().includes(queryLower);

      const matchesName =
        normName.includes(queryLower) ||
        rawCustomerName.toLowerCase().includes(trimmed.toLowerCase());

      const matchesTicketId =
        rawTicketId.toLowerCase() === ticketIdClean ||
        rawTicketId.toLowerCase().includes(ticketIdClean) ||
        (queryDigits.length > 0 && rawTicketId.includes(queryDigits));

      const rawProv = String(rec.applicationProvider ?? rec.provider ?? '').toLowerCase();
      const stdProv = String(
        rec.standardizedProvider || standardizeProvider(rawProv)
      ).toLowerCase();
      const matchesProvider =
        rawProv.includes(queryLower) || stdProv.includes(queryLower);

      return matchesPhone || matchesName || matchesTicketId || matchesProvider;
    });

    // Sort by Date (Latest to Oldest = desc, Oldest to Latest = asc)
    const sorted = [...matched].sort((a, b) => {
      const timeA = extractRecordTimestampMs(a);
      const timeB = extractRecordTimestampMs(b);
      if (timeA !== timeB) {
        return dateSortOrder === 'desc' ? timeB - timeA : timeA - timeB;
      }
      const idA = Number(a.ticketId) || 0;
      const idB = Number(b.ticketId) || 0;
      return dateSortOrder === 'desc' ? idB - idA : idA - idB;
    });

    return sorted;
  }, [
    records,
    query,
    searchField,
    startDateInput,
    endDateInput,
    dateSortOrder,
    selectedLoanCategory,
    selectedLoanProvider,
  ]);

  const displayedRecords = useMemo(() => {
    return filteredRecords.slice(0, visibleRowsCount);
  }, [filteredRecords, visibleRowsCount]);

  const handleApplyDatePreset = (start: string, end: string, fetchFromApi: boolean = false) => {
    setStartDateInput(start);
    setEndDateInput(end);
    if (fetchFromApi) {
      loadTicketsFromBackend(query, searchField, true, start || '2000-01-01', end || '2099-12-31');
    }
  };

  const handleSelectBatchSize = (size: 15 | 50 | 100) => {
    setBatchSize(size);
    setVisibleRowsCount(size);
  };

  const handleColumnDrop = (targetColKey: string) => {
    if (!draggedColumnKey || draggedColumnKey === targetColKey) return;
    setColumnOrder((prev) => {
      const next = [...prev];
      const fromIdx = next.indexOf(draggedColumnKey);
      const toIdx = next.indexOf(targetColKey);
      if (fromIdx === -1 || toIdx === -1) return prev;
      next.splice(fromIdx, 1);
      next.splice(toIdx, 0, draggedColumnKey);
      return next;
    });
    setDraggedColumnKey(null);
  };

  // ============================================================================
  // DETAIL VIEW (WHEN USER CLICKS "Visit" ON ANY TICKET)
  // Step 1-3: Calls /get-ticket-with-detail/{TICKET_ID} -> Customer Details + Uploaded Documents
  // Step 4-5: Immediately below Uploaded Documents -> Ticket Remarks & Activities (/get-ticket-activities/{TICKET_ID})
  // ============================================================================
  if (selectedTicket) {
    const activeTicketId = String(selectedTicket.ticketId ?? '');
    const detailData = ticketDetailResponse?.data || selectedTicket;
    const documentsList = extractDocumentsFromDetailData(detailData);

    const rawActivitiesList = Array.isArray(ticketActivitiesResponse?.activities)
      ? ticketActivitiesResponse!.activities!
      : [];

    const sortedActivitiesList = [...rawActivitiesList].sort((a, b) => {
      const tA = a?.created_at ? Date.parse(String(a.created_at)) : 0;
      const tB = b?.created_at ? Date.parse(String(b.created_at)) : 0;
      if (!isNaN(tA) && !isNaN(tB) && tA !== tB) {
        return activitiesSortOrder === 'desc' ? tB - tA : tA - tB;
      }
      const idA = Number(a?.id) || 0;
      const idB = Number(b?.id) || 0;
      return activitiesSortOrder === 'desc' ? idB - idA : idA - idB;
    });

    const primaryHandledKeys = new Set([
      'ticketId',
      'customerName',
      'customerContact',
      'customerEmail',
      'customerDesignation',
      'loanType',
      'customerLocation',
      'applicationAmount',
      'applicationTenure',
      'applicationDate',
      'provider',
      'applicationProvider',
      'loanCategory',
      'customerDocuments',
      'customerProfileImage',
    ]);

    const additionalEntries = Object.entries(detailData).filter(
      ([k]) => !primaryHandledKeys.has(k)
    );

    const tenureNum = Number(detailData.applicationTenure);
    const formattedTenure =
      detailData.applicationTenure !== null &&
      detailData.applicationTenure !== undefined &&
      !isNaN(tenureNum)
        ? `${tenureNum.toFixed(1)} years`
        : String(detailData.applicationTenure ?? '—');

    const providerDisplay =
      detailData.provider || detailData.applicationProvider || '—';

    return (
      <div className="py-4 px-2 sm:px-6 max-w-5xl mx-auto space-y-6">
        {/* Top Back & Refresh Detail Navigation Row */}
        <div className="flex flex-wrap items-center justify-between gap-3">
          <button
            type="button"
            onClick={() => {
              setSelectedTicket(null);
              setTicketDetailResponse(null);
              setTicketActivitiesResponse(null);
              setDetailError(null);
              setActivitiesError(null);
              setPreviewDoc(null);
              if (onClearInitialTicketId) onClearInitialTicketId();
              if (onBackToLookup) onBackToLookup();
            }}
            className="inline-flex items-center gap-2 text-slate-600 hover:text-slate-900 font-bold text-sm sm:text-base transition cursor-pointer py-1.5 px-2.5 rounded-lg hover:bg-slate-200/60"
          >
            <ArrowLeft className="w-4 h-4 stroke-[2.5]" />
            <span>{onBackToLookup ? 'Back to Customer Lookup' : 'Back'}</span>
          </button>

          <div className="flex items-center gap-2">
            {ticketDetailResponse?.endpointUrl && (
              <span className="hidden md:inline-flex items-center gap-1.5 px-3 py-1 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-700 text-xs font-mono">
                <CheckCircle2 className="w-3.5 h-3.5" />
                Ticket #{activeTicketId} Live
              </span>
            )}
            <button
              type="button"
              onClick={() =>
                openTicketDetailAndActivities(activeTicketId, selectedTicket, true)
              }
              disabled={isDetailLoading || isActivitiesLoading}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl bg-[#3f4bb8] hover:bg-[#323d9c] text-white font-bold text-xs transition cursor-pointer disabled:opacity-50 shadow-2xs"
            >
              <RefreshCw
                className={`w-3.5 h-3.5 ${
                  isDetailLoading || isActivitiesLoading ? 'animate-spin' : ''
                }`}
              />
              <span>
                {isDetailLoading || isActivitiesLoading
                  ? 'Refreshing Live Data...'
                  : 'Refresh Ticket & Activities'}
              </span>
            </button>
          </div>
        </div>

        {/* Centered Ticket ID Pill Header */}
        <div className="flex justify-center">
          <div className="px-8 sm:px-12 py-3 rounded-xl bg-[#e8eefc] border border-[#c5d4f7] text-[#35469c] font-extrabold text-lg sm:text-xl tracking-wider uppercase shadow-2xs">
            TICKET ID: F2FIN-{activeTicketId || '—'}
          </div>
        </div>

        {/* Detail API Loading / Error Banner */}
        {isDetailLoading && (
          <div className="bg-blue-50 border border-blue-200 rounded-2xl p-4 flex items-center justify-center gap-3 text-[#35469c]">
            <Loader2 className="w-5 h-5 animate-spin shrink-0" />
            <span className="text-xs sm:text-sm font-bold">
              Fetching live ticket details and uploaded documents from{' '}
              <code className="font-mono">/api/v1/get-ticket-with-detail/{activeTicketId}</code>...
            </span>
          </div>
        )}

        {detailError && (
          <div className="bg-rose-50 border border-rose-200 rounded-2xl p-4 flex flex-wrap items-center justify-between gap-3 text-rose-800">
            <div className="flex items-start gap-2.5">
              <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
              <div className="text-xs space-y-0.5">
                <p className="font-bold text-sm">Detail API Error</p>
                <p>{detailError}</p>
              </div>
            </div>
            <button
              type="button"
              onClick={() =>
                openTicketDetailAndActivities(activeTicketId, selectedTicket, true)
              }
              className="px-3 py-1.5 rounded-lg bg-rose-700 hover:bg-rose-800 text-white text-xs font-bold cursor-pointer"
            >
              Retry
            </button>
          </div>
        )}

        {/* Main White Detail Card */}
        <div className="bg-white rounded-[28px] shadow-sm border border-slate-200/80 p-8 sm:p-12 space-y-10">
          {/* ====================================================================
              STEP 2: CUSTOMER & TICKET DETAILS
             ==================================================================== */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-12 gap-y-8">
            {/* Row 1: NAME & CONTACT */}
            <div>
              <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1.5">
                NAME
              </span>
              <div className="text-lg sm:text-xl font-extrabold text-[#0b1f47]">
                {toTitleCase(detailData.customerName) || '—'}
              </div>
            </div>

            <div>
              <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1.5">
                CONTACT
              </span>
              <div className="text-lg sm:text-xl font-extrabold text-[#0b1f47]">
                {formatContactWithPrefix(detailData.customerContact)}
              </div>
            </div>

            {/* Row 2: EMAIL & DESIGNATION */}
            <div>
              <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1.5">
                EMAIL
              </span>
              <div className="text-lg sm:text-xl font-extrabold text-[#0b1f47] break-all">
                {detailData.customerEmail || '—'}
              </div>
            </div>

            <div>
              <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1.5">
                DESIGNATION
              </span>
              <div className="text-lg sm:text-xl font-extrabold text-[#0b1f47]">
                {formatDesignation(detailData.customerDesignation, detailData.loanType)}
              </div>
            </div>

            {/* Row 3: LOCATION & AMOUNT */}
            <div>
              <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1.5">
                LOCATION
              </span>
              <div className="text-lg sm:text-xl font-extrabold text-[#0b1f47]">
                {toTitleCase(detailData.customerLocation) || '—'}
              </div>
            </div>

            <div>
              <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1.5">
                AMOUNT
              </span>
              <div className="text-lg sm:text-xl font-extrabold text-[#0b1f47]">
                {formatIndianAmountNoDecimals(detailData.applicationAmount)}
              </div>
            </div>

            {/* Row 4: TENURE & APPLICATION DATE */}
            <div>
              <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1.5">
                TENURE
              </span>
              <div className="text-lg sm:text-xl font-extrabold text-[#0b1f47]">
                {formattedTenure}
              </div>
            </div>

            <div>
              <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1.5">
                APPLICATION DATE
              </span>
              <div className="text-lg sm:text-xl font-extrabold text-[#0b1f47]">
                {formatApplicationDateMMDDYYYY(
                  detailData.applicationDate,
                  detailData.createdAt
                )}
              </div>
            </div>

            {/* Row 5: LOAN PROVIDER & LOAN CATEGORY */}
            <div>
              <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1.5">
                LOAN PROVIDER (STANDARDIZED)
              </span>
              <div className="text-lg sm:text-xl font-extrabold text-[#0b1f47]">
                {detailData.standardizedProvider ||
                  standardizeProvider(providerDisplay)}
              </div>
              {providerDisplay && (
                <div className="text-xs text-slate-500 font-medium mt-1">
                  Original API Provider:{' '}
                  <span className="font-mono font-bold text-slate-700">
                    {String(providerDisplay)}
                  </span>
                </div>
              )}
            </div>

            <div>
              <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1.5">
                LOAN CATEGORY
              </span>
              <div className="text-lg sm:text-xl font-extrabold text-[#0b1f47]">
                {toTitleCase(detailData.loanCategory) || '—'}
              </div>
            </div>
          </div>

          {/* ====================================================================
              STEP 3: UPLOADED DOCUMENTS SECTION (customerDocuments from Detail API)
             ==================================================================== */}
          <div className="border-t border-slate-200 pt-8 space-y-5">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-xl bg-[#e8eefc] text-[#35469c] flex items-center justify-center">
                  <FileCheck className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base sm:text-lg font-extrabold text-[#0b1f47] tracking-tight">
                    Uploaded Documents
                  </h3>
                  <p className="text-xs text-[#4b5c82]">
                    Live documents retrieved from{' '}
                    <code className="font-mono">
                      /api/v1/get-ticket-with-detail/{activeTicketId}
                    </code>{' '}
                    (<code className="font-mono">customerDocuments</code>)
                  </p>
                </div>
              </div>

              <span className="px-3.5 py-1.5 rounded-xl bg-[#e8eefc] text-[#35469c] font-extrabold text-xs border border-[#c5d4f7]">
                {documentsList.length} {documentsList.length === 1 ? 'Document' : 'Documents'}
              </span>
            </div>

            {isDetailLoading && documentsList.length === 0 ? (
              <div className="p-8 rounded-2xl bg-slate-50 border border-slate-200 text-center space-y-2">
                <Loader2 className="w-6 h-6 animate-spin text-[#3f4bb8] mx-auto" />
                <p className="text-xs font-bold text-slate-600">
                  Loading uploaded customer documents...
                </p>
              </div>
            ) : documentsList.length === 0 ? (
              <div className="p-8 rounded-2xl bg-slate-50 border border-slate-200 text-center space-y-1">
                <FileText className="w-7 h-7 text-slate-400 mx-auto mb-1" />
                <p className="text-sm font-bold text-slate-700">
                  No Uploaded Documents Found for This Ticket
                </p>
                <p className="text-xs text-slate-500">
                  The detail API returned an empty <code className="font-mono">customerDocuments</code> array for Ticket #{activeTicketId}.
                </p>
              </div>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                {documentsList.map((doc) => {
                  const docKey = `${doc.id ?? doc.index}_${doc.url ?? doc.fileName}`;
                  const imgFailed = Boolean(failedImageMap[docKey]);

                  return (
                    <div
                      key={docKey}
                      className="rounded-2xl border border-slate-200 bg-slate-50/70 hover:bg-white hover:shadow-md transition-all overflow-hidden flex flex-col justify-between"
                    >
                      <div>
                        {/* Document Header */}
                        <div className="p-3.5 bg-white border-b border-slate-200/80 flex items-center justify-between gap-2">
                          <div className="flex items-center gap-2 min-w-0">
                            <span className="px-2 py-0.5 rounded-md bg-[#e8eefc] text-[#35469c] font-extrabold text-[11px] shrink-0">
                              #{doc.index}
                            </span>
                            <span className="font-extrabold text-xs sm:text-sm text-[#0b1f47] uppercase truncate">
                              {toTitleCase(doc.type)}
                            </span>
                          </div>
                          <span className="px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 font-mono font-bold text-[10px] shrink-0 border border-slate-200">
                            {doc.extension}
                          </span>
                        </div>

                        {/* Visual Thumbnail / Preview Area */}
                        <div className="h-40 bg-slate-100/80 flex items-center justify-center overflow-hidden relative border-b border-slate-200/60">
                          {!doc.url ? (
                            <div className="text-center p-4 text-slate-400">
                              <AlertCircle className="w-7 h-7 mx-auto mb-1 text-amber-500" />
                              <span className="text-xs font-bold block text-slate-600">
                                Document URL Missing
                              </span>
                            </div>
                          ) : doc.isImage && !imgFailed ? (
                            <img
                              src={doc.url}
                              alt={doc.type}
                              onClick={() => setPreviewDoc(doc)}
                              onError={() =>
                                setFailedImageMap((prev) => ({ ...prev, [docKey]: true }))
                              }
                              className="w-full h-full object-cover cursor-pointer hover:scale-105 transition-transform"
                            />
                          ) : doc.isPdf ? (
                            <div
                              onClick={() => setPreviewDoc(doc)}
                              className="text-center p-4 cursor-pointer group"
                            >
                              <div className="w-12 h-12 rounded-2xl bg-rose-100 text-rose-700 flex items-center justify-center mx-auto mb-2 group-hover:scale-105 transition-transform">
                                <FileText className="w-6 h-6" />
                              </div>
                              <span className="text-xs font-extrabold text-slate-700 block">
                                PDF Document
                              </span>
                              <span className="text-[11px] text-slate-500">
                                Click Preview or Open
                              </span>
                            </div>
                          ) : doc.isSheet ? (
                            <div
                              onClick={() =>
                                setSheetPreviewTarget({
                                  url: doc.url!,
                                  fileName: doc.fileName,
                                  ticketId: activeTicketId,
                                })
                              }
                              className="text-center p-4 cursor-pointer group"
                            >
                              <div className="w-12 h-12 rounded-2xl bg-emerald-100 text-emerald-700 flex items-center justify-center mx-auto mb-2 group-hover:scale-105 transition-transform">
                                <FileSpreadsheet className="w-6 h-6 text-emerald-600" />
                              </div>
                              <span className="text-xs font-extrabold text-slate-800 block">
                                Excel / Spreadsheet
                              </span>
                              <span className="text-[11px] text-emerald-600 font-bold">
                                Click to Preview Online
                              </span>
                            </div>
                          ) : (
                            <div className="text-center p-4">
                              <div className="w-12 h-12 rounded-2xl bg-blue-100 text-[#3f4bb8] flex items-center justify-center mx-auto mb-2">
                                <ImageIcon className="w-6 h-6" />
                              </div>
                              <span className="text-xs font-bold text-slate-700 block">
                                {imgFailed ? 'Direct File Link' : `${doc.extension} Document`}
                              </span>
                            </div>
                          )}
                        </div>

                        {/* Document Metadata Details */}
                        <div className="p-3.5 space-y-1 text-xs">
                          <div className="flex items-center justify-between text-slate-500">
                            <span className="font-bold uppercase text-[10px]">Document ID</span>
                            <span className="font-mono font-bold text-slate-800">
                              {doc.id !== null ? String(doc.id) : '—'}
                            </span>
                          </div>
                          <div className="flex items-center justify-between text-slate-500">
                            <span className="font-bold uppercase text-[10px]">Category / Type</span>
                            <span className="font-bold text-slate-800">
                              {toTitleCase(doc.type)}
                            </span>
                          </div>
                          <div className="pt-1">
                            <span className="font-bold uppercase text-[10px] text-slate-400 block">
                              File Name
                            </span>
                            <span
                              className="font-mono text-[11px] text-slate-700 truncate block"
                              title={doc.fileName}
                            >
                              {doc.fileName}
                            </span>
                          </div>
                        </div>
                      </div>

                      {/* Preview & Open Buttons */}
                      <div className="p-3.5 pt-0 flex items-center gap-2">
                        {doc.url ? (
                          <>
                            {doc.isSheet ? (
                              <button
                                type="button"
                                onClick={() =>
                                  setSheetPreviewTarget({
                                    url: doc.url!,
                                    fileName: doc.fileName,
                                    ticketId: activeTicketId,
                                  })
                                }
                                className="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shadow-xs transition cursor-pointer"
                                title="Preview spreadsheet online without downloading"
                              >
                                <FileSpreadsheet className="w-3.5 h-3.5" />
                                <span>Preview Sheet Online</span>
                              </button>
                            ) : (
                              <button
                                type="button"
                                onClick={() => setPreviewDoc(doc)}
                                className="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-2 rounded-xl bg-[#e8eefc] hover:bg-[#d8e3fc] text-[#35469c] font-bold text-xs transition cursor-pointer"
                              >
                                <Eye className="w-3.5 h-3.5" />
                                <span>Preview</span>
                              </button>
                            )}
                            <a
                              href={`/api/f2-tickets/proxy-document?url=${encodeURIComponent(
                                doc.url
                              )}&download=1&filename=${encodeURIComponent(
                                doc.fileName || 'document.pdf'
                              )}`}
                              download={doc.fileName || 'document.pdf'}
                              title="Download document"
                              className="p-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 transition cursor-pointer shrink-0"
                            >
                              <Download className="w-3.5 h-3.5" />
                            </a>
                            <a
                              href={doc.url}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-2 rounded-xl bg-[#3f4bb8] hover:bg-[#323d9c] text-white font-bold text-xs transition cursor-pointer"
                            >
                              <ExternalLink className="w-3.5 h-3.5" />
                              <span>Open</span>
                            </a>
                          </>
                        ) : (
                          <span className="w-full py-2 text-center rounded-xl bg-slate-200 text-slate-500 font-bold text-xs">
                            Unavailable
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* ====================================================================
              STEP 4 & 5: TICKET REMARKS & ACTIVITIES SECTION
              Placed immediately BELOW Uploaded Documents
              Calls https://admin.f2fintech.in/api/v1/get-ticket-activities/{TICKET_ID}
             ==================================================================== */}
          <div className="border-t border-slate-200 pt-8 space-y-6">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-xl bg-[#eef3ff] text-[#2557d6] flex items-center justify-center">
                  <MessageSquare className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base sm:text-lg font-extrabold text-[#0b1f47] tracking-tight">
                    Ticket Remarks &amp; Activities
                  </h3>
                  <p className="text-xs text-[#4b5c82]">
                    Complete remark &amp; activity timeline from{' '}
                    <code className="font-mono">
                      /api/v1/get-ticket-activities/{activeTicketId}
                    </code>
                  </p>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                {sortedActivitiesList.length > 1 && (
                  <button
                    type="button"
                    onClick={() =>
                      setActivitiesSortOrder((prev) => (prev === 'desc' ? 'asc' : 'desc'))
                    }
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs transition cursor-pointer"
                  >
                    <ArrowUpDown className="w-3.5 h-3.5" />
                    <span>
                      {activitiesSortOrder === 'desc' ? 'Newest First' : 'Oldest First'}
                    </span>
                  </button>
                )}

                <span className="px-3.5 py-1.5 rounded-xl bg-[#eef3ff] text-[#2557d6] font-extrabold text-xs border border-[#c5d4f7]">
                  {sortedActivitiesList.length}{' '}
                  {sortedActivitiesList.length === 1 ? 'Remark / Activity' : 'Remarks & Activities'}
                </span>
              </div>
            </div>

            {/* Activities Error State (does not break ticket details or documents) */}
            {activitiesError && (
              <div className="bg-rose-50 border border-rose-200 rounded-2xl p-4 flex flex-wrap items-center justify-between gap-3 text-rose-800">
                <div className="flex items-start gap-2.5">
                  <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
                  <div className="text-xs space-y-0.5">
                    <p className="font-bold text-sm">Activities API Error</p>
                    <p>{activitiesError}</p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => fetchTicketActivitiesById(activeTicketId)}
                  className="px-3 py-1.5 rounded-lg bg-rose-700 hover:bg-rose-800 text-white text-xs font-bold cursor-pointer"
                >
                  Retry Activities
                </button>
              </div>
            )}

            {/* Activities Loading State */}
            {isActivitiesLoading && sortedActivitiesList.length === 0 ? (
              <div className="p-8 rounded-2xl bg-slate-50 border border-slate-200 text-center space-y-2">
                <Loader2 className="w-6 h-6 animate-spin text-[#3f4bb8] mx-auto" />
                <p className="text-xs font-bold text-slate-600">
                  Loading ticket remarks and activity history...
                </p>
              </div>
            ) : !activitiesError && sortedActivitiesList.length === 0 ? (
              <div className="p-8 rounded-2xl bg-slate-50 border border-slate-200 text-center space-y-1">
                <MessageSquare className="w-7 h-7 text-slate-400 mx-auto mb-1" />
                <p className="text-sm font-bold text-slate-700">
                  No remarks or activities found for this ticket.
                </p>
              </div>
            ) : (
              <div className="space-y-5">
                {sortedActivitiesList.map((activity, idx) => {
                  const actId = activity.id ?? `act_${idx}`;
                  const userObj =
                    activity.user && typeof activity.user === 'object' ? activity.user : {};
                  const rawUsername =
                    userObj.username ||
                    activity.username ||
                    activity.createdBy ||
                    activity.employeeName ||
                    (activity.user_id ? `User #${activity.user_id}` : 'Employee');
                  const displayEmployeeName = capitalizeFirstChar(rawUsername);
                  const avatarInitial = displayEmployeeName.charAt(0).toUpperCase() || 'U';

                  const rawComment =
                    activity.comment ??
                    activity.remark ??
                    activity.activity ??
                    activity.description ??
                    '';

                  const createdAtFormatted = formatActivityDateTime(
                    activity.created_at || activity.createdAt || activity.updated_at
                  );

                  const remarkAttachments = extractAttachmentsFromActivity(activity);

                  const userRole = userObj.role || activity.role || '';
                  const userStatus = userObj.status || activity.status || '';
                  const userEmail = userObj.email || '';
                  const userNumber = userObj.number || '';

                  return (
                    <div
                      key={String(actId)}
                      className="rounded-2xl bg-white border border-slate-200/90 shadow-2xs hover:shadow-md transition-all p-5 sm:p-6"
                    >
                      <div className="flex items-start gap-4">
                        {/* Circular Initial Avatar matching uploaded screenshot */}
                        <div className="w-11 h-11 rounded-full bg-[#eef3ff] text-[#1d4ed8] font-extrabold text-base flex items-center justify-center shrink-0 select-none">
                          {avatarInitial}
                        </div>

                        <div className="flex-1 min-w-0 space-y-2.5">
                          {/* Employee Name & Formatted Timestamp Header */}
                          <div className="flex flex-wrap items-baseline justify-between gap-2">
                            <div className="flex flex-wrap items-baseline gap-x-3 gap-y-1">
                              <span className="text-base sm:text-lg font-extrabold text-slate-900">
                                {displayEmployeeName}
                              </span>
                              <span className="text-xs sm:text-sm font-medium text-[#3b5284]">
                                {createdAtFormatted}
                              </span>
                            </div>

                            <div className="flex flex-wrap items-center gap-1.5">
                              {userRole && (
                                <span className="px-2.5 py-0.5 rounded-md bg-slate-100 text-slate-700 font-bold text-[11px] uppercase">
                                  {String(userRole)}
                                </span>
                              )}
                              {userStatus && (
                                <span className="px-2.5 py-0.5 rounded-md bg-emerald-50 text-emerald-700 border border-emerald-200 font-bold text-[11px] uppercase">
                                  {String(userStatus)}
                                </span>
                              )}
                              {activity.id !== undefined && activity.id !== null && (
                                <span className="px-2 py-0.5 rounded-md bg-slate-50 text-slate-500 font-mono text-[11px] border border-slate-200">
                                  #{String(activity.id)}
                                </span>
                              )}
                            </div>
                          </div>

                          {/* Remark / Comment Text */}
                          <div className="text-sm sm:text-base text-slate-700 whitespace-pre-line leading-relaxed break-words">
                            {rawComment !== null &&
                            rawComment !== undefined &&
                            String(rawComment).trim() !== ''
                              ? String(rawComment)
                              : '—'}
                          </div>

                          {/* Attached Documents inside the Remark (if present) */}
                          {remarkAttachments.length > 0 && (
                            <div className="pt-2 space-y-2">
                              <div className="text-[11px] font-extrabold uppercase tracking-wider text-[#35469c] flex items-center gap-1.5">
                                <Paperclip className="w-3.5 h-3.5" />
                                <span>
                                  Attached Remark{' '}
                                  {remarkAttachments.length === 1 ? 'Document' : 'Documents'} (
                                  {remarkAttachments.length})
                                </span>
                              </div>

                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                {remarkAttachments.map((att, attIdx) => {
                                  const attKey = `att_${actId}_${attIdx}_${att.url}`;
                                  const imgFailed = Boolean(failedImageMap[attKey]);

                                  return (
                                    <div
                                      key={attKey}
                                      className="p-3 rounded-xl bg-slate-50 border border-slate-200 flex flex-col justify-between gap-3"
                                    >
                                      <div className="flex items-center gap-3 min-w-0">
                                        {att.url && att.isImage && !imgFailed ? (
                                          <img
                                            src={att.url}
                                            alt={att.fileName}
                                            onClick={() => setPreviewDoc(att)}
                                            onError={() =>
                                              setFailedImageMap((prev) => ({
                                                ...prev,
                                                [attKey]: true,
                                              }))
                                            }
                                            className="w-12 h-12 rounded-lg object-cover border border-slate-200 shrink-0 cursor-pointer"
                                          />
                                        ) : att.isSheet ? (
                                          <div
                                            onClick={() =>
                                              setSheetPreviewTarget({
                                                url: att.url!,
                                                fileName: att.fileName,
                                                ticketId: activeTicketId,
                                                activityAuthor: displayEmployeeName,
                                                comment: rawComment ? String(rawComment).trim() : null,
                                              })
                                            }
                                            className="w-11 h-11 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center shrink-0 cursor-pointer hover:bg-emerald-200 transition"
                                            title="Click to preview sheet online"
                                          >
                                            <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
                                          </div>
                                        ) : (
                                          <div className="w-11 h-11 rounded-xl bg-[#e8eefc] text-[#35469c] flex items-center justify-center shrink-0">
                                            <FileText className="w-5 h-5" />
                                          </div>
                                        )}

                                        <div className="min-w-0 flex-1">
                                          <div className="flex items-center gap-1.5">
                                            <span
                                              className={`px-1.5 py-0.5 rounded text-[10px] font-mono font-bold ${
                                                att.isSheet
                                                  ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                                  : 'bg-white border border-slate-200 text-slate-700'
                                              }`}
                                            >
                                              {att.isSheet ? 'XLSX / SHEET' : att.extension}
                                            </span>
                                            <span
                                              className="text-xs font-extrabold text-slate-800 truncate"
                                              title={att.fileName}
                                            >
                                              {att.fileName}
                                            </span>
                                          </div>
                                          <p
                                            className="text-[11px] text-slate-500 font-mono truncate mt-0.5"
                                            title={att.url || ''}
                                          >
                                            {att.url || 'Missing URL'}
                                          </p>
                                        </div>
                                      </div>

                                      {att.url && (
                                        <div className="flex items-center gap-2">
                                          {att.isSheet ? (
                                            <>
                                              <button
                                                type="button"
                                                onClick={() =>
                                                  setSheetPreviewTarget({
                                                    url: att.url!,
                                                    fileName: att.fileName,
                                                    ticketId: activeTicketId,
                                                    activityAuthor: displayEmployeeName,
                                                    comment: rawComment ? String(rawComment).trim() : null,
                                                  })
                                                }
                                                className="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shadow-xs transition cursor-pointer"
                                                title="Preview spreadsheet online without downloading"
                                              >
                                                <FileSpreadsheet className="w-3.5 h-3.5" />
                                                <span>Preview Sheet Online</span>
                                              </button>
                                              <a
                                                href={`/api/f2-tickets/proxy-document?url=${encodeURIComponent(
                                                  att.url
                                                )}&download=1&filename=${encodeURIComponent(
                                                  att.fileName || 'sheet.xlsx'
                                                )}`}
                                                download={att.fileName || 'sheet.xlsx'}
                                                className="p-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs transition cursor-pointer"
                                                title="Download sheet (.xlsx)"
                                              >
                                                <Download className="w-3.5 h-3.5" />
                                              </a>
                                              <a
                                                href={att.url}
                                                target="_blank"
                                                rel="noopener noreferrer"
                                                className="p-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs transition cursor-pointer"
                                                title="Open raw file URL"
                                              >
                                                <ExternalLink className="w-3.5 h-3.5" />
                                              </a>
                                            </>
                                          ) : (
                                            <>
                                              {(att.isImage || att.isPdf) && (
                                                <button
                                                  type="button"
                                                  onClick={() => setPreviewDoc(att)}
                                                  className="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-1.5 rounded-lg bg-[#e8eefc] hover:bg-[#d8e3fc] text-[#35469c] font-bold text-xs transition cursor-pointer"
                                                >
                                                  <Eye className="w-3.5 h-3.5" />
                                                  <span>Preview</span>
                                                </button>
                                              )}
                                              <a
                                                href={att.url}
                                                target="_blank"
                                                rel="noopener noreferrer"
                                                className="flex-1 inline-flex items-center justify-center gap-1.5 px-3 py-1.5 rounded-lg bg-[#3f4bb8] hover:bg-[#323d9c] text-white font-bold text-xs transition cursor-pointer"
                                              >
                                                <ExternalLink className="w-3.5 h-3.5" />
                                                <span>Open Document</span>
                                              </a>
                                            </>
                                          )}
                                        </div>
                                      )}
                                    </div>
                                  );
                                })}
                              </div>
                            </div>
                          )}

                          {/* Employee & Timestamp Metadata Footer */}
                          {(userEmail || userNumber || activity.created_at) && (
                            <div className="pt-2 border-t border-slate-100 flex flex-wrap items-center gap-x-4 gap-y-1 text-[11px] text-slate-500">
                              {userEmail && (
                                <span>
                                  <strong className="text-slate-600">Email:</strong> {userEmail}
                                </span>
                              )}
                              {userNumber && (
                                <span>
                                  <strong className="text-slate-600">Contact:</strong> {userNumber}
                                </span>
                              )}
                              {activity.created_at && (
                                <span className="font-mono">
                                  <strong className="font-sans text-slate-600">Created:</strong>{' '}
                                  {String(activity.created_at)}
                                </span>
                              )}
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* ====================================================================
              ALL REMAINING DETAIL API FIELDS (Customer, Loan, Status, Dates, etc.)
             ==================================================================== */}
          {additionalEntries.length > 0 && (
            <div className="border-t border-slate-200 pt-8 space-y-6">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <h3 className="text-sm sm:text-base font-extrabold uppercase tracking-wider text-[#0b1f47]">
                  Additional Ticket, Customer &amp; Loan Details ({additionalEntries.length} Fields)
                </h3>

                <button
                  type="button"
                  onClick={() => setShowRawDetailJson((prev) => !prev)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs transition cursor-pointer"
                >
                  <Code className="w-3.5 h-3.5" />
                  <span>{showRawDetailJson ? 'Hide Raw JSON' : 'View Raw Detail JSON'}</span>
                </button>
              </div>

              {showRawDetailJson && (
                <pre className="p-4 rounded-2xl bg-slate-900 text-emerald-300 font-mono text-xs overflow-x-auto max-h-96">
                  {JSON.stringify(
                    {
                      ticketDetail: detailData,
                      ticketActivities: sortedActivitiesList,
                    },
                    null,
                    2
                  )}
                </pre>
              )}

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-12 gap-y-6">
                {additionalEntries.map(([key, val]) => {
                  const label = key
                    .replace(/([A-Z])/g, ' $1')
                    .replace(/_/g, ' ')
                    .trim()
                    .toUpperCase();

                  let displayVal: React.ReactNode = '—';
                  if (val !== null && val !== undefined && val !== '') {
                    if (Array.isArray(val)) {
                      displayVal =
                        val.length === 0
                          ? '—'
                          : val.map((item, idx) => {
                              const s =
                                typeof item === 'object' ? JSON.stringify(item) : String(item);
                              return /^https?:\/\//i.test(s) ? (
                                <a
                                  key={idx}
                                  href={s}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  className="text-blue-600 hover:underline text-sm block break-all"
                                >
                                  Open Link ({idx + 1})
                                </a>
                              ) : (
                                <span key={idx} className="block">
                                  {s}
                                </span>
                              );
                            });
                    } else if (typeof val === 'object') {
                      displayVal = JSON.stringify(val);
                    } else if (key === 'createdAt') {
                      displayVal = formatCreatedAtTableDate(val);
                    } else if (/^https?:\/\//i.test(String(val))) {
                      displayVal = (
                        <a
                          href={String(val)}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="text-blue-600 hover:underline text-sm break-all inline-flex items-center gap-1"
                        >
                          <span>Open URL</span>
                          <ExternalLink className="w-3.5 h-3.5" />
                        </a>
                      );
                    } else {
                      displayVal = String(val);
                    }
                  }

                  return (
                    <div key={key}>
                      <span className="block text-xs font-bold uppercase tracking-wider text-[#4b5c82] mb-1">
                        {label}
                      </span>
                      <div className="text-base sm:text-lg font-extrabold text-[#0b1f47] break-words">
                        {displayVal}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>

        {/* In-App Document Preview Modal (works for both Uploaded Documents & Remark Attachments) */}
        {previewDoc && (
          <div
            className="fixed inset-0 z-50 bg-slate-900/75 backdrop-blur-xs flex items-center justify-center p-4"
            onClick={() => setPreviewDoc(null)}
          >
            <div
              className={`bg-white rounded-3xl w-full flex flex-col overflow-hidden shadow-2xl border border-slate-200 ${
                previewDoc.isPdf
                  ? 'max-w-5xl h-[88vh]'
                  : 'max-w-4xl max-h-[90vh]'
              }`}
              onClick={(e) => e.stopPropagation()}
            >
              <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-3">
                <div className="min-w-0">
                  <div className="flex items-center gap-2">
                    <span className="px-2.5 py-0.5 rounded-md bg-[#e8eefc] text-[#35469c] font-extrabold text-xs">
                      Document #{previewDoc.index}
                    </span>
                    <h4 className="font-extrabold text-slate-900 text-sm sm:text-base uppercase truncate">
                      {toTitleCase(previewDoc.type)}
                    </h4>
                    {previewDoc.isPdf && (
                      <span className="px-2 py-0.5 rounded-md bg-rose-50 text-rose-700 border border-rose-200 font-mono font-bold text-[10px]">
                        PDF
                      </span>
                    )}
                  </div>
                  <p className="text-xs text-slate-500 font-mono truncate mt-0.5">
                    {previewDoc.fileName}
                  </p>
                </div>

                <div className="flex items-center gap-2 shrink-0">
                  {previewDoc.url && (
                    <>
                      <a
                        href={`/api/f2-tickets/proxy-document?url=${encodeURIComponent(
                          previewDoc.url
                        )}&download=1&filename=${encodeURIComponent(
                          previewDoc.fileName || 'document.pdf'
                        )}`}
                        download={previewDoc.fileName || 'document.pdf'}
                        className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs transition"
                        title="Download Document"
                      >
                        <Download className="w-3.5 h-3.5" />
                        <span className="hidden sm:inline">Download</span>
                      </a>
                      <a
                        href={previewDoc.url}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl bg-[#3f4bb8] hover:bg-[#323d9c] text-white font-bold text-xs transition"
                      >
                        <ExternalLink className="w-3.5 h-3.5" />
                        <span className="hidden sm:inline">Open in New Tab</span>
                      </a>
                    </>
                  )}
                  <button
                    type="button"
                    onClick={() => setPreviewDoc(null)}
                    className="p-2 rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-700 cursor-pointer"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>
              </div>

              <div
                className={`flex-1 overflow-auto flex items-center justify-center ${
                  previewDoc.isPdf ? 'p-0 bg-slate-900' : 'p-4 bg-slate-100 min-h-[380px]'
                }`}
              >
                {!previewDoc.url ? (
                  <p className="text-sm font-bold text-slate-600">Document URL is missing.</p>
                ) : previewDoc.isImage ? (
                  <img
                    src={previewDoc.url}
                    alt={previewDoc.type}
                    className="max-w-full max-h-[72vh] object-contain rounded-xl shadow-md bg-white"
                  />
                ) : previewDoc.isPdf ? (
                  <PdfDocumentViewer
                    url={previewDoc.url}
                    fileName={previewDoc.fileName}
                    title={previewDoc.type}
                  />
                ) : previewDoc.isSheet ? (
                  <div className="text-center space-y-4 p-8 bg-white rounded-3xl border border-emerald-200 shadow-md max-w-md mx-auto">
                    <div className="w-16 h-16 rounded-2xl bg-emerald-100 text-emerald-600 flex items-center justify-center mx-auto shadow-xs">
                      <FileSpreadsheet className="w-8 h-8" />
                    </div>
                    <div>
                      <span className="px-2.5 py-0.5 rounded-full text-[10px] font-extrabold uppercase bg-emerald-50 text-emerald-700 border border-emerald-200">
                        Spreadsheet Document
                      </span>
                      <h4 className="font-extrabold text-slate-900 text-base mt-2 truncate" title={previewDoc.fileName}>
                        {previewDoc.fileName}
                      </h4>
                      <p className="text-xs text-slate-500 mt-1">
                        Format: <strong className="font-mono text-slate-700">{previewDoc.extension}</strong> • Ready for instant interactive preview
                      </p>
                    </div>
                    <div className="flex flex-col sm:flex-row items-center gap-2.5 justify-center pt-2">
                      <button
                        type="button"
                        onClick={() => {
                          const target = {
                            url: previewDoc.url!,
                            fileName: previewDoc.fileName,
                            ticketId: activeTicketId,
                          };
                          setPreviewDoc(null);
                          setSheetPreviewTarget(target);
                        }}
                        className="w-full sm:w-auto inline-flex items-center justify-center gap-2 px-5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold text-xs shadow-md transition cursor-pointer"
                      >
                        <FileSpreadsheet className="w-4 h-4" />
                        <span>Preview Sheet Online</span>
                      </button>
                      {previewDoc.url && (
                        <a
                          href={`/api/f2-tickets/proxy-document?url=${encodeURIComponent(
                            previewDoc.url
                          )}&download=1&filename=${encodeURIComponent(
                            previewDoc.fileName || 'sheet.xlsx'
                          )}`}
                          download={previewDoc.fileName || 'sheet.xlsx'}
                          className="w-full sm:w-auto inline-flex items-center justify-center gap-1.5 px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs transition cursor-pointer"
                        >
                          <Download className="w-4 h-4" />
                          <span>Download</span>
                        </a>
                      )}
                    </div>
                  </div>
                ) : (
                  <div className="text-center space-y-3 p-8 bg-white rounded-2xl border border-slate-200">
                    <FileText className="w-10 h-10 text-[#3f4bb8] mx-auto" />
                    <p className="text-sm font-bold text-slate-800">{previewDoc.fileName}</p>
                    <a
                      href={previewDoc.url}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-[#3f4bb8] text-white font-bold text-xs"
                    >
                      <ExternalLink className="w-3.5 h-3.5" />
                      <span>Open Document Directly</span>
                    </a>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* Online Sheet Preview Modal */}
        <OnlineSheetModal
          isOpen={Boolean(sheetPreviewTarget)}
          onClose={() => setSheetPreviewTarget(null)}
          documentUrl={sheetPreviewTarget?.url || null}
          fileName={sheetPreviewTarget?.fileName}
          ticketId={sheetPreviewTarget?.ticketId}
          activityAuthor={sheetPreviewTarget?.activityAuthor}
          comment={sheetPreviewTarget?.comment}
        />
      </div>
    );
  }

  // ============================================================================
  // MAIN TABLE VIEW
  // ============================================================================
  return (
    <div className="space-y-5">
      {/* Search, Date Limiter & Sort Controls Card */}
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2.5">
            <span className="px-3 py-1 rounded-xl text-xs font-extrabold bg-blue-50 text-blue-700 border border-blue-200 shadow-2xs">
              Showing {displayedRecords.length.toLocaleString()} of{' '}
              {filteredRecords.length.toLocaleString()} Matching Tickets (Total in Memory:{' '}
              {totalInMemory.toLocaleString()})
            </span>
            {fetchedAt && (
              <span className="text-xs text-slate-500 font-mono">
                Synced at {new Date(fetchedAt).toLocaleTimeString()}
              </span>
            )}

            {/* Live Auto-Update Status Indicator */}
            {autoUpdateEnabled ? (
              <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-xl text-xs font-semibold bg-emerald-50 text-emerald-800 border border-emerald-200 shadow-2xs">
                <span className="relative flex h-2 w-2">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
                </span>
                <span>Auto-Update: ON</span>
                <span className="text-[10px] font-mono text-emerald-700 bg-emerald-100/70 px-1.5 py-0.2 rounded font-bold">
                  {isAutoChecking ? 'syncing...' : `${secondsUntilNextCheck}s`}
                </span>
              </div>
            ) : (
              <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-xl text-xs font-semibold bg-slate-100 text-slate-600 border border-slate-200">
                <span className="h-2 w-2 rounded-full bg-slate-400"></span>
                <span>Auto-Update Paused</span>
              </div>
            )}

            {/* Notice toast when update occurs */}
            {lastAutoUpdateNotice && (
              <span className="text-xs font-medium text-emerald-700 bg-emerald-50 border border-emerald-300 px-2.5 py-1 rounded-xl shadow-2xs flex items-center gap-1">
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                <span>{lastAutoUpdateNotice}</span>
              </span>
            )}
          </div>

          <div className="flex items-center gap-2">
            {/* Auto-Update Quick Setting Dropdown/Toggle */}
            <div className="relative">
              <button
                type="button"
                onClick={() => setShowAutoSettings(!showAutoSettings)}
                className={`px-3 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 border shadow-2xs cursor-pointer ${
                  autoUpdateEnabled
                    ? 'bg-emerald-50 text-emerald-800 border-emerald-300 hover:bg-emerald-100'
                    : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'
                }`}
                title="Configure Auto-Update frequency"
              >
                <Zap className={`w-3.5 h-3.5 ${autoUpdateEnabled ? 'text-emerald-600 fill-emerald-500' : 'text-slate-400'}`} />
                <span>Auto: {autoUpdateEnabled ? `${autoUpdateInterval}s` : 'OFF'}</span>
                <ChevronDown className="w-3 h-3 text-slate-400" />
              </button>

              {/* Popover settings menu */}
              {showAutoSettings && (
                <div className="absolute right-0 top-full mt-1.5 z-30 w-56 bg-white rounded-xl shadow-lg border border-slate-200 p-3 space-y-2.5 text-xs">
                  <div className="flex items-center justify-between pb-2 border-b border-slate-100">
                    <span className="font-bold text-slate-800">Auto-Update</span>
                    <button
                      type="button"
                      onClick={() => handleToggleAutoUpdate(!autoUpdateEnabled)}
                      className={`relative inline-flex h-5 w-9 shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out ${
                        autoUpdateEnabled ? 'bg-emerald-600' : 'bg-slate-300'
                      }`}
                    >
                      <span
                        className={`pointer-events-none inline-block h-4 w-4 transform rounded-full bg-white shadow-sm ring-0 transition duration-200 ease-in-out ${
                          autoUpdateEnabled ? 'translate-x-4' : 'translate-x-0'
                        }`}
                      />
                    </button>
                  </div>

                  <div>
                    <span className="text-[11px] font-semibold text-slate-500 block mb-1.5">
                      Check Frequency
                    </span>
                    <div className="grid grid-cols-2 gap-1.5">
                      {[
                        { sec: 30, label: '30 sec' },
                        { sec: 60, label: '1 min' },
                        { sec: 120, label: '2 min' },
                        { sec: 300, label: '5 min' },
                      ].map((opt) => (
                        <button
                          key={opt.sec}
                          type="button"
                          onClick={() => handleChangeInterval(opt.sec)}
                          className={`px-2 py-1 text-center rounded-lg text-xs font-semibold transition-colors border ${
                            autoUpdateInterval === opt.sec
                              ? 'bg-emerald-600 text-white border-emerald-700'
                              : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                          }`}
                        >
                          {opt.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="pt-2 border-t border-slate-100 flex justify-between items-center">
                    <button
                      type="button"
                      onClick={() => {
                        setShowAutoSettings(false);
                        handleManualAutoCheck();
                      }}
                      className="w-full text-center py-1.5 px-2 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-lg text-xs font-bold transition-colors"
                    >
                      {isAutoChecking ? 'Checking Live API...' : 'Check OMS Now'}
                    </button>
                  </div>
                </div>
              )}
            </div>

            <button
              type="button"
              onClick={() =>
                loadTicketsFromBackend(
                  query,
                  searchField,
                  true,
                  startDateInput.trim() || '2000-01-01',
                  endDateInput.trim() || '2099-12-31'
                )
              }
              disabled={isLoading || isRefreshing}
              className="px-3.5 py-2 rounded-xl bg-[#3f4bb8] hover:bg-[#323d9c] text-white font-bold text-xs transition-all flex items-center gap-2 cursor-pointer shadow-xs disabled:opacity-50"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isRefreshing ? 'animate-spin' : ''}`} />
              <span>{isRefreshing ? 'Fetching All Tickets...' : 'Refresh All Tickets'}</span>
            </button>
          </div>
        </div>

        {/* Date Limiter & Date Sort Section */}
        <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200/90 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <Calendar className="w-4 h-4 text-[#3f4bb8]" />
              <span className="text-xs font-extrabold uppercase tracking-wider text-slate-700">
                Date Limiter &amp; Date Sort
              </span>
            </div>

            {/* Dynamic Date Presets: All Dates + Current Month (e.g. Oct 2026) + ONLY Years Present in Data */}
            <div className="flex flex-wrap items-center gap-1.5">
              <button
                type="button"
                onClick={() => handleApplyDatePreset('', '', false)}
                className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition cursor-pointer ${
                  !startDateInput && !endDateInput
                    ? 'bg-[#3f4bb8] text-white'
                    : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-200'
                }`}
              >
                All Dates ({totalInMemory.toLocaleString()})
              </button>

              <button
                type="button"
                onClick={() =>
                  handleApplyDatePreset(latestMonthInfo.start, latestMonthInfo.end, false)
                }
                className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition cursor-pointer ${
                  startDateInput === latestMonthInfo.start &&
                  endDateInput === latestMonthInfo.end
                    ? 'bg-[#3f4bb8] text-white'
                    : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-200'
                }`}
              >
                {latestMonthInfo.label}
              </button>

              {discoveredYears.map((yr) => {
                const yrStart = `${yr}-01-01`;
                const yrEnd = `${yr}-12-31`;
                const isSelected = startDateInput === yrStart && endDateInput === yrEnd;
                return (
                  <button
                    key={yr}
                    type="button"
                    onClick={() => handleApplyDatePreset(yrStart, yrEnd, false)}
                    className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition cursor-pointer ${
                      isSelected
                        ? 'bg-[#3f4bb8] text-white'
                        : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-200'
                    }`}
                  >
                    Year {yr}
                  </button>
                );
              })}
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 items-end">
            <div>
              <label className="block text-[10px] font-bold uppercase text-slate-500 mb-1">
                From Date
              </label>
              <input
                type="date"
                value={startDateInput}
                onChange={(e) => setStartDateInput(e.target.value)}
                className="w-full px-3 py-2 rounded-xl border border-slate-300 bg-white text-xs font-mono focus:outline-hidden focus:ring-2 focus:ring-blue-500/20 focus:border-blue-600"
              />
            </div>

            <div>
              <label className="block text-[10px] font-bold uppercase text-slate-500 mb-1">
                To Date
              </label>
              <input
                type="date"
                value={endDateInput}
                onChange={(e) => setEndDateInput(e.target.value)}
                className="w-full px-3 py-2 rounded-xl border border-slate-300 bg-white text-xs font-mono focus:outline-hidden focus:ring-2 focus:ring-blue-500/20 focus:border-blue-600"
              />
            </div>

            {/* Sort By Date Asc / Desc Selector */}
            <div>
              <label className="block text-[10px] font-bold uppercase text-slate-500 mb-1">
                Sort By Date
              </label>
              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => setDateSortOrder('desc')}
                  className={`flex-1 inline-flex items-center justify-center gap-1 px-2.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer border ${
                    dateSortOrder === 'desc'
                      ? 'bg-[#3f4bb8] text-white border-[#3f4bb8]'
                      : 'bg-white hover:bg-slate-100 text-slate-700 border-slate-300'
                  }`}
                >
                  <ArrowUpDown className="w-3.5 h-3.5" />
                  <span>Latest to Oldest</span>
                </button>
                <button
                  type="button"
                  onClick={() => setDateSortOrder('asc')}
                  className={`flex-1 inline-flex items-center justify-center gap-1 px-2.5 py-2 rounded-xl text-xs font-bold transition cursor-pointer border ${
                    dateSortOrder === 'asc'
                      ? 'bg-[#3f4bb8] text-white border-[#3f4bb8]'
                      : 'bg-white hover:bg-slate-100 text-slate-700 border-slate-300'
                  }`}
                >
                  <ArrowUpDown className="w-3.5 h-3.5" />
                  <span>Oldest to Latest</span>
                </button>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() =>
                  loadTicketsFromBackend(
                    query,
                    searchField,
                    true,
                    startDateInput.trim() || '2000-01-01',
                    endDateInput.trim() || '2099-12-31'
                  )
                }
                disabled={isLoading || isRefreshing}
                className="flex-1 px-3.5 py-2 rounded-xl bg-[#3f4bb8] hover:bg-[#323d9c] text-white font-bold text-xs transition cursor-pointer disabled:opacity-50"
              >
                Apply Date Range
              </button>

              {(startDateInput || endDateInput) && (
                <button
                  type="button"
                  onClick={() => {
                    setStartDateInput('');
                    setEndDateInput('');
                    if (appliedStartDate !== '2000-01-01' || appliedEndDate !== '2099-12-31') {
                      loadTicketsFromBackend(query, searchField, true, '2000-01-01', '2099-12-31');
                    }
                  }}
                  className="px-3 py-2 rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs transition cursor-pointer"
                >
                  Clear
                </button>
              )}
            </div>
          </div>
        </div>

        {/* Loan Category Type & Loan Provider Type Filter (Dropdowns + Drag-and-Drop Columns + All Reset) */}
        <div
          onDragOver={(e) => e.preventDefault()}
          onDrop={() => {
            if (!draggedFilterItem) return;
            if (draggedFilterItem.kind === 'category') {
              setSelectedLoanCategory(draggedFilterItem.value);
            } else if (draggedFilterItem.kind === 'provider') {
              setSelectedLoanProvider(draggedFilterItem.value);
            }
            setDraggedFilterItem(null);
          }}
          className="p-3.5 rounded-xl bg-slate-50 border border-slate-200/90 space-y-3"
        >
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex flex-wrap items-center gap-2">
              <Layers className="w-4 h-4 text-[#3f4bb8]" />
              <span className="text-xs font-extrabold uppercase tracking-wider text-slate-700">
                Loan Category Type &amp; Loan Provider Filter
              </span>
              <span className="text-[11px] text-slate-500">
                (Select from dropdown, click, or drag &amp; drop items)
              </span>
            </div>

            <div className="flex flex-wrap items-center gap-1.5">
              <button
                type="button"
                onClick={() => {
                  setSelectedLoanCategory('all');
                  setSelectedLoanProvider('all');
                }}
                className={`px-3 py-1 rounded-lg text-xs font-extrabold transition cursor-pointer ${
                  selectedLoanCategory === 'all' && selectedLoanProvider === 'all'
                    ? 'bg-[#3f4bb8] text-white shadow-2xs'
                    : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-200'
                }`}
              >
                All Categories &amp; Providers ({totalInMemory.toLocaleString()})
              </button>

              {(selectedLoanCategory !== 'all' || selectedLoanProvider !== 'all') && (
                <button
                  type="button"
                  onClick={() => {
                    setSelectedLoanCategory('all');
                    setSelectedLoanProvider('all');
                  }}
                  className="px-2.5 py-1 rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-bold transition cursor-pointer inline-flex items-center gap-1"
                >
                  <X className="w-3 h-3" />
                  <span>Reset to All</span>
                </button>
              )}
            </div>
          </div>

          {/* Two Drag-and-Drop Columns: Column 1 = Loan Category / Type, Column 2 = Loan Provider Type */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
            {/* Column 1: Loan Category Type */}
            <div
              onDragOver={(e) => e.preventDefault()}
              onDrop={() => {
                if (draggedFilterItem?.kind === 'category') {
                  setSelectedLoanCategory(draggedFilterItem.value);
                  setDraggedFilterItem(null);
                }
              }}
              className="p-3 rounded-xl bg-white border border-slate-200/90 space-y-2.5"
            >
              <div className="flex items-center justify-between gap-2">
                <label className="block text-[11px] font-extrabold uppercase tracking-wider text-[#35469c]">
                  Loan Category / Loan Type
                </label>
                {selectedLoanCategory !== 'all' && (
                  <button
                    type="button"
                    onClick={() => setSelectedLoanCategory('all')}
                    className="text-[11px] font-bold text-blue-700 hover:underline cursor-pointer"
                  >
                    All Categories
                  </button>
                )}
              </div>

              <select
                value={selectedLoanCategory}
                onChange={(e) => setSelectedLoanCategory(e.target.value)}
                className="w-full px-3 py-2 rounded-xl border border-slate-300 bg-slate-50/70 text-xs font-bold text-slate-800 focus:outline-hidden focus:ring-2 focus:ring-blue-500/20 focus:border-blue-600 cursor-pointer"
              >
                <option value="all">All Loan Categories &amp; Types ({totalInMemory.toLocaleString()})</option>
                <optgroup label="Loan Category (Secured / Unsecured)">
                  {discoveredCategories
                    .filter((c) => c.group === 'Category')
                    .map((c) => (
                      <option key={c.key} value={c.key}>
                        {c.label} ({c.count.toLocaleString()})
                      </option>
                    ))}
                </optgroup>
                <optgroup label="Loan Type (All Fields)">
                  {discoveredCategories
                    .filter((c) => c.group === 'Loan Type')
                    .map((c) => (
                      <option key={c.key} value={c.key}>
                        {c.label} ({c.count.toLocaleString()})
                      </option>
                    ))}
                </optgroup>
              </select>

              {/* Draggable & Clickable Chips for All Loan Categories & Loan Types */}
              <div className="flex flex-wrap items-center gap-1.5 max-h-28 overflow-y-auto pt-0.5 pr-1">
                <button
                  type="button"
                  draggable
                  onDragStart={() =>
                    setDraggedFilterItem({ kind: 'category', value: 'all' })
                  }
                  onClick={() => setSelectedLoanCategory('all')}
                  className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition cursor-pointer border ${
                    selectedLoanCategory === 'all'
                      ? 'bg-[#3f4bb8] text-white border-[#3f4bb8]'
                      : 'bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200'
                  }`}
                >
                  All
                </button>
                {discoveredCategories.map((c) => {
                  const active = selectedLoanCategory === c.key;
                  return (
                    <button
                      key={c.key}
                      type="button"
                      draggable
                      onDragStart={() =>
                        setDraggedFilterItem({ kind: 'category', value: c.key })
                      }
                      onClick={() =>
                        setSelectedLoanCategory(active ? 'all' : c.key)
                      }
                      title="Click or Drag & Drop to filter by this Loan Category / Type"
                      className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition cursor-grab active:cursor-grabbing border ${
                        active
                          ? 'bg-[#3f4bb8] text-white border-[#3f4bb8]'
                          : 'bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200'
                      }`}
                    >
                      {c.label} ({c.count.toLocaleString()})
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Column 2: Loan Provider Type */}
            <div
              onDragOver={(e) => e.preventDefault()}
              onDrop={() => {
                if (draggedFilterItem?.kind === 'provider') {
                  setSelectedLoanProvider(draggedFilterItem.value);
                  setDraggedFilterItem(null);
                }
              }}
              className="p-3 rounded-xl bg-white border border-slate-200/90 space-y-2.5"
            >
              <div className="flex items-center justify-between gap-2">
                <label className="block text-[11px] font-extrabold uppercase tracking-wider text-[#35469c]">
                  Loan Provider Type ({discoveredProviders.length} Standardized Categories)
                </label>
                {selectedLoanProvider !== 'all' && (
                  <button
                    type="button"
                    onClick={() => setSelectedLoanProvider('all')}
                    className="text-[11px] font-bold text-blue-700 hover:underline cursor-pointer"
                  >
                    All Providers
                  </button>
                )}
              </div>

              <select
                value={selectedLoanProvider}
                onChange={(e) => setSelectedLoanProvider(e.target.value)}
                className="w-full px-3 py-2 rounded-xl border border-slate-300 bg-slate-50/70 text-xs font-bold text-slate-800 focus:outline-hidden focus:ring-2 focus:ring-blue-500/20 focus:border-blue-600 cursor-pointer"
              >
                <option value="all">All Standardized Loan Providers ({totalInMemory.toLocaleString()})</option>
                {discoveredProviders.map((p) => (
                  <option key={p.key} value={p.key}>
                    {p.label} ({p.count.toLocaleString()})
                  </option>
                ))}
              </select>

              {/* Draggable & Clickable Chips for All Loan Providers */}
              <div className="flex flex-wrap items-center gap-1.5 max-h-28 overflow-y-auto pt-0.5 pr-1">
                <button
                  type="button"
                  draggable
                  onDragStart={() =>
                    setDraggedFilterItem({ kind: 'provider', value: 'all' })
                  }
                  onClick={() => setSelectedLoanProvider('all')}
                  className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition cursor-pointer border ${
                    selectedLoanProvider === 'all'
                      ? 'bg-[#3f4bb8] text-white border-[#3f4bb8]'
                      : 'bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200'
                  }`}
                >
                  All
                </button>
                {discoveredProviders.map((p) => {
                  const active = selectedLoanProvider === p.key;
                  return (
                    <button
                      key={p.key}
                      type="button"
                      draggable
                      onDragStart={() =>
                        setDraggedFilterItem({ kind: 'provider', value: p.key })
                      }
                      onClick={() =>
                        setSelectedLoanProvider(active ? 'all' : p.key)
                      }
                      title="Click or Drag & Drop to filter by this Loan Provider"
                      className={`px-2.5 py-1 rounded-lg text-[11px] font-bold transition cursor-grab active:cursor-grabbing border ${
                        active
                          ? 'bg-[#3f4bb8] text-white border-[#3f4bb8]'
                          : 'bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200'
                      }`}
                    >
                      {p.label} ({p.count.toLocaleString()})
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
        </div>

        {/* Search Field Filter Pills & Input */}
        <div className="flex flex-wrap items-center gap-2 pt-1">
          <button
            type="button"
            onClick={() => setSearchField('all')}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              searchField === 'all'
                ? 'bg-[#3f4bb8] text-white'
                : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
            }`}
          >
            <Layers className="w-3.5 h-3.5" />
            <span>All Fields</span>
          </button>

          <button
            type="button"
            onClick={() => setSearchField('phone')}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              searchField === 'phone'
                ? 'bg-[#3f4bb8] text-white'
                : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
            }`}
          >
            <Phone className="w-3.5 h-3.5" />
            <span>Mobile / Phone</span>
          </button>

          <button
            type="button"
            onClick={() => setSearchField('name')}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              searchField === 'name'
                ? 'bg-[#3f4bb8] text-white'
                : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
            }`}
          >
            <User className="w-3.5 h-3.5" />
            <span>Customer Name</span>
          </button>

          <button
            type="button"
            onClick={() => setSearchField('ticketId')}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              searchField === 'ticketId'
                ? 'bg-[#3f4bb8] text-white'
                : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
            }`}
          >
            <Hash className="w-3.5 h-3.5" />
            <span>Ticket ID</span>
          </button>
        </div>

        <form
          onSubmit={(e) => {
            e.preventDefault();
            loadTicketsFromBackend(query, searchField, false);
          }}
          className="flex flex-col sm:flex-row gap-2.5"
        >
          <div className="relative flex-1">
            <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder={
                searchField === 'phone'
                  ? 'Search by customer mobile/phone number...'
                  : searchField === 'name'
                  ? 'Search by customer name...'
                  : searchField === 'ticketId'
                  ? 'Search by ticket ID (e.g. 7268 or 7498)...'
                  : 'Search by Customer Phone, Customer Name, or Ticket ID...'
              }
              className="w-full pl-10 pr-10 py-2.5 rounded-xl border border-slate-300 text-sm focus:outline-hidden focus:ring-2 focus:ring-blue-500/20 focus:border-blue-600"
            />
            {query && (
              <button
                type="button"
                onClick={() => {
                  setQuery('');
                  loadTicketsFromBackend('', searchField, false);
                }}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700 p-1 cursor-pointer"
                title="Clear search"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>

          <button
            type="submit"
            disabled={isLoading}
            className="px-6 py-2.5 rounded-xl bg-[#3f4bb8] hover:bg-[#323d9c] text-white font-bold text-xs uppercase tracking-wider transition-all disabled:opacity-50 shadow-xs cursor-pointer flex items-center justify-center gap-2"
          >
            <Search className="w-4 h-4" />
            <span>Search</span>
          </button>
        </form>
      </div>

      {/* Error Banner */}
      {error && (
        <div className="bg-rose-50 border border-rose-200 rounded-2xl p-4 flex items-start gap-3 text-rose-800">
          <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
          <div className="text-xs space-y-1">
            <p className="font-bold text-sm">API Fetch Error</p>
            <p>{error}</p>
          </div>
        </div>
      )}

      {/* Loading State */}
      {(isLoading || isRefreshing) && (
        <div className="bg-white rounded-2xl border border-slate-200 p-12 text-center space-y-3">
          <Loader2 className="w-8 h-8 animate-spin text-[#3f4bb8] mx-auto" />
          <p className="text-sm font-bold text-slate-800">
            Loading all tickets from F2 Fintech API...
          </p>
        </div>
      )}

      {/* Tickets Table with Top Row Count Controls (15 default, 50, 100, Show Next, Show All) */}
      {!isLoading && !isRefreshing && (
        <div className="bg-white rounded-t-2xl rounded-b-xl shadow-sm border border-slate-200 overflow-hidden">
          {/* TOP ROW DISPLAY & LOAD MORE BAR (ABOVE THE TABLE) */}
          <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-xs text-slate-600 font-medium">
                Showing <strong>{displayedRecords.length.toLocaleString()}</strong> of{' '}
                <strong>{filteredRecords.length.toLocaleString()}</strong> matching tickets
              </span>

              <div className="flex items-center gap-1.5 ml-1">
                <span className="text-[11px] font-bold uppercase text-slate-400">
                  Show at once:
                </span>
                {([15, 50, 100] as const).map((sz) => (
                  <button
                    key={sz}
                    type="button"
                    onClick={() => handleSelectBatchSize(sz)}
                    className={`px-2.5 py-1 rounded-lg text-xs font-extrabold transition cursor-pointer ${
                      batchSize === sz && visibleRowsCount === sz
                        ? 'bg-[#3f4bb8] text-white shadow-2xs'
                        : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-200'
                    }`}
                  >
                    {sz}
                  </button>
                ))}
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {displayedRecords.length < filteredRecords.length && (
                <>
                  <button
                    type="button"
                    onClick={() => setVisibleRowsCount((prev) => prev + batchSize)}
                    className="px-4 py-1.5 rounded-xl bg-[#3f4bb8] hover:bg-[#323d9c] text-white font-bold text-xs cursor-pointer transition"
                  >
                    Show Next {batchSize} Rows
                  </button>
                  <button
                    type="button"
                    onClick={() => setVisibleRowsCount(filteredRecords.length)}
                    className="px-4 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs cursor-pointer transition"
                  >
                    Show All ({filteredRecords.length.toLocaleString()}) Rows
                  </button>
                </>
              )}
              {visibleRowsCount > 15 && (
                <button
                  type="button"
                  onClick={() => {
                    setBatchSize(15);
                    setVisibleRowsCount(15);
                  }}
                  className="px-3 py-1.5 rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs cursor-pointer transition"
                >
                  Reset to 15
                </button>
              )}
              {columnOrder.join(',') !== DEFAULT_TABLE_COLUMNS.join(',') && (
                <button
                  type="button"
                  onClick={() => setColumnOrder([...DEFAULT_TABLE_COLUMNS])}
                  className="px-3 py-1.5 rounded-xl bg-indigo-50 hover:bg-indigo-100 text-[#3f4bb8] border border-indigo-200 font-bold text-xs cursor-pointer transition"
                >
                  Reset Column Order
                </button>
              )}
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-left">
              <thead>
                <tr className="bg-[#3f4bb8] text-white text-sm font-extrabold">
                  {columnOrder.map((colKey, colIdx) => {
                    const isLast = colIdx === columnOrder.length - 1;
                    const borderClass = isLast ? '' : 'border-r border-white/30 pr-3';
                    const isDraggingThis = draggedColumnKey === colKey;

                    const headerLabelMap: Record<string, React.ReactNode> = {
                      tid: 'T.ID',
                      source: 'Source',
                      name: 'Name',
                      email: 'Email',
                      amount: 'Amount',
                      provider: 'Provider',
                      loanCategory: (
                        <span>
                          Loan
                          <br />
                          Category
                        </span>
                      ),
                      status: 'Status',
                      loanType: (
                        <span>
                          Loan
                          <br />
                          Type
                        </span>
                      ),
                      leadType: (
                        <span>
                          Lead
                          <br />
                          Type
                        </span>
                      ),
                      tenure: 'Tenure',
                      location: 'Location',
                      createdAt: (
                        <div className="flex items-center gap-1">
                          <span>
                            Created
                            <br />
                            At
                          </span>
                          <ArrowUpDown className="w-3.5 h-3.5 opacity-90 shrink-0" />
                        </div>
                      ),
                      actions: 'Actions',
                    };

                    return (
                      <th
                        key={colKey}
                        draggable
                        onDragStart={() => setDraggedColumnKey(colKey)}
                        onDragOver={(e) => e.preventDefault()}
                        onDrop={() => handleColumnDrop(colKey)}
                        onClick={
                          colKey === 'createdAt'
                            ? () =>
                                setDateSortOrder((prev) => (prev === 'desc' ? 'asc' : 'desc'))
                            : undefined
                        }
                        title={
                          colKey === 'createdAt'
                            ? 'Drag to reorder column or Click to toggle Date Sort'
                            : 'Drag & Drop to reorder column'
                        }
                        className={`py-4 px-4 select-none transition cursor-grab active:cursor-grabbing hover:bg-white/10 ${
                          colKey === 'actions' ? 'text-center px-5 whitespace-nowrap' : ''
                        } ${colKey === 'name' ? 'min-w-[160px]' : ''} ${
                          colKey === 'email' ? 'min-w-[180px]' : ''
                        } ${colKey === 'loanType' || colKey === 'loanCategory' ? 'min-w-[115px]' : ''} ${
                          colKey === 'location' ? 'min-w-[130px]' : ''
                        } ${isDraggingThis ? 'opacity-50 bg-white/20' : ''}`}
                      >
                        <div className={`${borderClass} leading-tight`}>
                          {headerLabelMap[colKey] || colKey}
                        </div>
                      </th>
                    );
                  })}
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-200/80 text-xs sm:text-[13px] text-slate-800">
                {displayedRecords.length === 0 ? (
                  <tr>
                    <td colSpan={columnOrder.length} className="py-12 text-center text-slate-500">
                      No matching tickets found.
                    </td>
                  </tr>
                ) : (
                  displayedRecords.map((ticket, index) => {
                    const ticketKey =
                      ticket.ticketId !== undefined && ticket.ticketId !== null
                        ? String(ticket.ticketId)
                        : `row_${index}`;

                    const tenureVal = Number(ticket.applicationTenure);
                    const tenureFormatted =
                      ticket.applicationTenure !== null &&
                      ticket.applicationTenure !== undefined &&
                      !isNaN(tenureVal)
                        ? `${tenureVal.toFixed(1)} years`
                        : '—';

                    const locParts = [
                      toTitleCase(ticket.customerLocation),
                      toTitleCase(ticket.customerState),
                    ].filter(Boolean);

                    return (
                      <tr
                        key={ticketKey}
                        className="hover:bg-slate-50/90 transition-colors"
                      >
                        {columnOrder.map((colKey) => {
                          switch (colKey) {
                            case 'tid':
                              return (
                                <td
                                  key={colKey}
                                  className="py-4 px-4 font-extrabold text-slate-900 whitespace-nowrap"
                                >
                                  {String(ticket.ticketId ?? '—')}
                                </td>
                              );
                            case 'source':
                              return (
                                <td key={colKey} className="py-4 px-4 whitespace-nowrap">
                                  {ticket.applicationSource ? (
                                    <span className="inline-block px-2.5 py-0.5 rounded-md bg-blue-50 text-[#3f4bb8] border border-blue-200/80 text-[11px] font-bold uppercase tracking-wide">
                                      {String(ticket.applicationSource)}
                                    </span>
                                  ) : (
                                    '—'
                                  )}
                                </td>
                              );
                            case 'name':
                              return (
                                <td
                                  key={colKey}
                                  className="py-4 px-4 font-extrabold text-slate-900 uppercase leading-snug"
                                >
                                  {String(ticket.customerName ?? '—')
                                    .trim()
                                    .replace(/\s+/g, ' ')}
                                </td>
                              );
                            case 'email':
                              return (
                                <td
                                  key={colKey}
                                  className="py-4 px-4 text-slate-700 break-all leading-snug"
                                >
                                  {ticket.customerEmail || '—'}
                                </td>
                              );
                            case 'amount':
                              return (
                                <td
                                  key={colKey}
                                  className="py-4 px-4 font-semibold text-[#15803d] whitespace-nowrap text-sm"
                                >
                                  {formatIndianAmountWithDecimals(ticket.applicationAmount)}
                                </td>
                              );
                            case 'provider': {
                              const rawP =
                                ticket.applicationProvider ||
                                ticket.provider ||
                                '';
                              const stdP =
                                ticket.standardizedProvider ||
                                standardizeProvider(rawP);
                              return (
                                <td
                                  key={colKey}
                                  className="py-4 px-4 font-extrabold text-[#6b21a8] leading-tight"
                                >
                                  <div>{stdP}</div>
                                  {rawP && rawP.trim() !== stdP && (
                                    <div
                                      className="text-[10px] text-slate-500 font-normal font-mono truncate max-w-[170px]"
                                      title={`Original API value: ${rawP}`}
                                    >
                                      Orig: {rawP}
                                    </div>
                                  )}
                                </td>
                              );
                            }
                            case 'loanCategory':
                              return (
                                <td key={colKey} className="py-4 px-4 whitespace-nowrap">
                                  {ticket.loanCategory ? (
                                    <span className="inline-block px-2.5 py-0.5 rounded-md bg-slate-100 text-slate-800 border border-slate-200 text-xs font-bold">
                                      {toTitleCase(ticket.loanCategory)}
                                    </span>
                                  ) : (
                                    '—'
                                  )}
                                </td>
                              );
                            case 'status':
                              return (
                                <td key={colKey} className="py-4 px-4 whitespace-nowrap">
                                  <span className="inline-block px-3.5 py-1 rounded-full bg-[#e8f0fe] text-[#1d4ed8] font-bold text-xs">
                                    {toTitleCase(
                                      ticket.ticketStatus || ticket.loanStatus || 'Active'
                                    )}
                                  </span>
                                </td>
                              );
                            case 'loanType':
                              return (
                                <td
                                  key={colKey}
                                  className="py-4 px-4 text-slate-900 font-medium leading-snug"
                                >
                                  {toTitleCase(ticket.loanType) || '—'}
                                </td>
                              );
                            case 'leadType':
                              return (
                                <td
                                  key={colKey}
                                  className="py-4 px-4 text-slate-900 font-medium whitespace-nowrap"
                                >
                                  {toTitleCase(ticket.leadType) || '—'}
                                </td>
                              );
                            case 'tenure':
                              return (
                                <td
                                  key={colKey}
                                  className="py-4 px-4 text-slate-900 font-medium whitespace-nowrap"
                                >
                                  {tenureFormatted}
                                </td>
                              );
                            case 'location':
                              return (
                                <td
                                  key={colKey}
                                  className="py-4 px-4 text-slate-900 font-medium leading-snug"
                                >
                                  {locParts.length > 0 ? locParts.join(' , ') : '—'}
                                </td>
                              );
                            case 'createdAt':
                              return (
                                <td
                                  key={colKey}
                                  className="py-4 px-4 text-slate-900 font-medium whitespace-nowrap"
                                >
                                  {formatCreatedAtTableDate(ticket.createdAt)}
                                </td>
                              );
                            case 'actions':
                              return (
                                <td key={colKey} className="py-4 px-5 text-center whitespace-nowrap">
                                  <button
                                    type="button"
                                    onClick={() =>
                                      openTicketDetailAndActivities(
                                        ticket.ticketId,
                                        ticket,
                                        false
                                      )
                                    }
                                    className="inline-flex items-center justify-center gap-1.5 px-4 py-1.5 rounded-full bg-[#eaf2ff] hover:bg-[#dbe8ff] text-[#1d4ed8] font-bold text-xs transition cursor-pointer shadow-2xs"
                                  >
                                    <ExternalLink className="w-3.5 h-3.5" />
                                    <span>Visit</span>
                                  </button>
                                </td>
                              );
                            default:
                              return <td key={colKey} className="py-4 px-4">—</td>;
                          }
                        })}
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Bottom Load More Footer as well */}
          {displayedRecords.length < filteredRecords.length && (
            <div className="px-5 py-3.5 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3">
              <span className="text-xs text-slate-600 font-medium">
                Showing <strong>{displayedRecords.length.toLocaleString()}</strong> of{' '}
                <strong>{filteredRecords.length.toLocaleString()}</strong> matching tickets
              </span>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setVisibleRowsCount((prev) => prev + batchSize)}
                  className="px-4 py-1.5 rounded-xl bg-[#3f4bb8] hover:bg-[#323d9c] text-white font-bold text-xs cursor-pointer transition"
                >
                  Show Next {batchSize} Rows
                </button>
                <button
                  type="button"
                  onClick={() => setVisibleRowsCount(filteredRecords.length)}
                  className="px-4 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-900 text-white font-bold text-xs cursor-pointer transition"
                >
                  Show All ({filteredRecords.length.toLocaleString()}) Rows
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Online Sheet Preview Modal */}
      <OnlineSheetModal
        isOpen={Boolean(sheetPreviewTarget)}
        onClose={() => setSheetPreviewTarget(null)}
        documentUrl={sheetPreviewTarget?.url || null}
        fileName={sheetPreviewTarget?.fileName}
        ticketId={sheetPreviewTarget?.ticketId}
        activityAuthor={sheetPreviewTarget?.activityAuthor}
        comment={sheetPreviewTarget?.comment}
      />
    </div>
  );
};
