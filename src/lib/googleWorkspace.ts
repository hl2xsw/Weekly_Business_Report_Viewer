import { parseExcelBufferToGrid } from './xlsxParser';

export const GOOGLE_SHEET_MIME = 'application/vnd.google-apps.spreadsheet';
export const EXCEL_XLSX_MIME =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export const EXCEL_XLS_MIME = 'application/vnd.ms-excel';
export const EXCEL_XLSM_MIME = 'application/vnd.ms-excel.sheet.macroEnabled.12';
export const SHORTCUT_MIME = 'application/vnd.google-apps.shortcut';

const SUPPORTED_SPREADSHEET_MIMES = new Set([
  GOOGLE_SHEET_MIME,
  EXCEL_XLSX_MIME,
  EXCEL_XLS_MIME,
  EXCEL_XLSM_MIME,
]);

export type SpreadsheetFileKind = 'google-sheet' | 'excel';

export interface DriveSpreadsheetFile {
  id: string;
  name: string;
  mimeType: string;
  fileKind: SpreadsheetFileKind;
  modifiedTime: string;
  webViewLink?: string;
  ownerName?: string;
  shared?: boolean;
}

export interface SheetTabInfo {
  sheetId: number;
  title: string;
  index: number;
  hidden: boolean;
}

export interface TextFormatRun {
  startIndex?: number;
  format?: {
    strikethrough?: boolean;
  };
}

export interface SheetCellData {
  formattedValue?: string;
  effectiveFormat?: {
    textFormat?: {
      strikethrough?: boolean;
    };
  };
  textFormatRuns?: TextFormatRun[];
}

export interface GridRange {
  sheetId?: number;
  startRowIndex?: number;
  endRowIndex?: number;
  startColumnIndex?: number;
  endColumnIndex?: number;
}

export interface SheetGridPayload {
  spreadsheetTitle: string;
  fileKind?: SpreadsheetFileKind;
  webViewLink?: string;
  tabs: SheetTabInfo[];
  activeTab: SheetTabInfo;
  merges: GridRange[];
  columnHidden: boolean[];
  rowHidden?: boolean[];
  rows: SheetCellData[][];
}

export class WorkspaceApiError extends Error {
  status: number;

  constructor(message: string, status: number) {
    super(message);
    this.name = 'WorkspaceApiError';
    this.status = status;
  }
}

interface DriveFileMetadata {
  id: string;
  name: string;
  mimeType: string;
  webViewLink?: string;
  shortcutDetails?: {
    targetId?: string;
    targetMimeType?: string;
  };
}

// In-memory cache for downloaded Excel files so switching tabs is instantaneous
const excelBufferCache = new Map<
  string,
  { buffer: ArrayBuffer; name: string; webViewLink?: string }
>();

export function clearSpreadsheetCache(spreadsheetId?: string) {
  if (spreadsheetId) {
    excelBufferCache.delete(spreadsheetId);
  } else {
    excelBufferCache.clear();
  }
}

export function extractSpreadsheetId(input: string): string | null {
  const trimmed = input.trim();
  if (!trimmed) return null;

  const sheetsMatch = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (sheetsMatch?.[1]) {
    return sheetsMatch[1];
  }

  const driveFileMatch = trimmed.match(/\/file\/d\/([a-zA-Z0-9-_]+)/);
  if (driveFileMatch?.[1]) {
    return driveFileMatch[1];
  }

  const queryIdMatch = trimmed.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (queryIdMatch?.[1]) {
    return queryIdMatch[1];
  }

  if (/^[a-zA-Z0-9-_]{20,}$/.test(trimmed)) {
    return trimmed;
  }

  return null;
}

async function fetchGoogleJson<T>(url: string, accessToken: string): Promise<T> {
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!res.ok) {
    let errorMessage = `요청 처리 중 오류가 발생했습니다. (${res.status})`;
    try {
      const body = await res.json();
      if (body?.error?.message) {
        errorMessage = body.error.message;
      }
    } catch {
      // ignore json parse failure
    }
    throw new WorkspaceApiError(errorMessage, res.status);
  }

  return res.json() as Promise<T>;
}

async function fetchGoogleBinary(
  url: string,
  accessToken: string
): Promise<ArrayBuffer> {
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!res.ok) {
    let errorMessage = `파일 다운로드 중 오류가 발생했습니다. (${res.status})`;
    try {
      const body = await res.json();
      if (body?.error?.message) {
        errorMessage = body.error.message;
      }
    } catch {
      // ignore json parse failure
    }
    throw new WorkspaceApiError(errorMessage, res.status);
  }

  return res.arrayBuffer();
}

function classifyFileKind(mimeType: string, name?: string): SpreadsheetFileKind {
  if (mimeType === GOOGLE_SHEET_MIME) return 'google-sheet';
  if (
    mimeType === EXCEL_XLSX_MIME ||
    mimeType === EXCEL_XLS_MIME ||
    mimeType === EXCEL_XLSM_MIME
  ) {
    return 'excel';
  }
  if (name && /\.(xlsx|xlsm|xls)$/i.test(name)) {
    return 'excel';
  }
  return 'google-sheet';
}

export async function listDriveSpreadsheets(
  accessToken: string,
  searchQuery = '',
  scope: 'all' | 'sharedWithMe' = 'all'
): Promise<DriveSpreadsheetFile[]> {
  const mimeFilter = [
    `mimeType='${GOOGLE_SHEET_MIME}'`,
    `mimeType='${EXCEL_XLSX_MIME}'`,
    `mimeType='${EXCEL_XLS_MIME}'`,
    `mimeType='${EXCEL_XLSM_MIME}'`,
    `mimeType='${SHORTCUT_MIME}'`,
  ].join(' or ');

  const qParts = [`(${mimeFilter})`, 'trashed=false'];

  if (scope === 'sharedWithMe') {
    qParts.push('sharedWithMe=true');
  }

  const cleanQuery = searchQuery.trim().replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  if (cleanQuery) {
    // Combine name and fullText search so Korean filenames with underscores (e.g. 경영위원회_기술연구소.xlsx) match reliably
    qParts.push(`(name contains '${cleanQuery}' or fullText contains '${cleanQuery}')`);
  }

  const params = new URLSearchParams({
    q: qParts.join(' and '),
    orderBy: 'modifiedTime desc',
    pageSize: '100',
    supportsAllDrives: 'true',
    includeItemsFromAllDrives: 'true',
    fields:
      'files(id,name,mimeType,modifiedTime,webViewLink,shared,owners(displayName,emailAddress),shortcutDetails(targetId,targetMimeType))',
  });

  const data = await fetchGoogleJson<{
    files?: Array<{
      id: string;
      name: string;
      mimeType: string;
      modifiedTime: string;
      webViewLink?: string;
      shared?: boolean;
      owners?: Array<{ displayName?: string; emailAddress?: string }>;
      shortcutDetails?: { targetId?: string; targetMimeType?: string };
    }>;
  }>(`https://www.googleapis.com/drive/v3/files?${params.toString()}`, accessToken);

  const results: DriveSpreadsheetFile[] = [];
  const seenIds = new Set<string>();

  for (const f of data.files ?? []) {
    let resolvedId = f.id;
    let resolvedMime = f.mimeType;

    if (f.mimeType === SHORTCUT_MIME) {
      const targetMime = f.shortcutDetails?.targetMimeType ?? '';
      const targetId = f.shortcutDetails?.targetId;
      if (!targetId || !SUPPORTED_SPREADSHEET_MIMES.has(targetMime)) {
        continue;
      }
      resolvedId = targetId;
      resolvedMime = targetMime;
    }

    if (seenIds.has(resolvedId)) continue;
    seenIds.add(resolvedId);

    results.push({
      id: resolvedId,
      name: f.name,
      mimeType: resolvedMime,
      fileKind: classifyFileKind(resolvedMime, f.name),
      modifiedTime: f.modifiedTime,
      webViewLink: f.webViewLink,
      ownerName: f.owners?.[0]?.displayName || f.owners?.[0]?.emailAddress,
      shared: f.shared,
    });
  }

  return results;
}

async function fetchDriveFileMetadata(
  accessToken: string,
  fileId: string
): Promise<DriveFileMetadata | null> {
  try {
    const params = new URLSearchParams({
      supportsAllDrives: 'true',
      fields: 'id,name,mimeType,webViewLink,shortcutDetails(targetId,targetMimeType)',
    });
    const meta = await fetchGoogleJson<DriveFileMetadata>(
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?${params.toString()}`,
      accessToken
    );

    if (meta.mimeType === SHORTCUT_MIME && meta.shortcutDetails?.targetId) {
      return {
        ...meta,
        id: meta.shortcutDetails.targetId,
        mimeType: meta.shortcutDetails.targetMimeType || meta.mimeType,
      };
    }

    return meta;
  } catch (err) {
    if (err instanceof WorkspaceApiError && err.status === 401) {
      throw err;
    }
    return null;
  }
}

export async function fetchSpreadsheetTabs(
  accessToken: string,
  spreadsheetId: string
): Promise<{ title: string; tabs: SheetTabInfo[] }> {
  const params = new URLSearchParams({
    fields: 'properties.title,sheets(properties(sheetId,title,index,hidden))',
  });

  const data = await fetchGoogleJson<{
    properties?: { title?: string };
    sheets?: Array<{
      properties?: {
        sheetId?: number;
        title?: string;
        index?: number;
        hidden?: boolean;
      };
    }>;
  }>(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}?${params.toString()}`,
    accessToken
  );

  const tabs: SheetTabInfo[] = (data.sheets ?? []).map((s, idx) => ({
    sheetId: s.properties?.sheetId ?? idx,
    title: s.properties?.title ?? `Sheet ${idx + 1}`,
    index: s.properties?.index ?? idx,
    hidden: Boolean(s.properties?.hidden),
  }));

  return {
    title: data.properties?.title || '주간업무보고',
    tabs,
  };
}

export async function fetchSheetGrid(
  accessToken: string,
  spreadsheetId: string,
  sheetTitle: string,
  options?: { forceRefresh?: boolean }
): Promise<SheetGridPayload> {
  const driveMeta = await fetchDriveFileMetadata(accessToken, spreadsheetId);
  const resolvedId = driveMeta?.id ?? spreadsheetId;
  const isExcelFile =
    driveMeta !== null &&
    classifyFileKind(driveMeta.mimeType, driveMeta.name) === 'excel';

  if (isExcelFile) {
    let cached = options?.forceRefresh
      ? undefined
      : excelBufferCache.get(resolvedId);

    if (!cached) {
      const buffer = await fetchGoogleBinary(
        `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(resolvedId)}?alt=media&supportsAllDrives=true`,
        accessToken
      );
      cached = {
        buffer,
        name: driveMeta.name || '주간업무보고.xlsx',
        webViewLink: driveMeta.webViewLink,
      };
      excelBufferCache.set(resolvedId, cached);
    }

    const parsedGrid = await parseExcelBufferToGrid(
      cached.buffer,
      cached.name,
      sheetTitle || undefined
    );

    return {
      ...parsedGrid,
      fileKind: 'excel',
      webViewLink: cached.webViewLink,
    };
  }

  // Native Google Spreadsheet via Sheets API v4
  const meta = await fetchSpreadsheetTabs(accessToken, resolvedId);
  const targetTab =
    meta.tabs.find((t) => t.title === sheetTitle) ??
    meta.tabs.find((t) => !t.hidden) ??
    meta.tabs[0];

  if (!targetTab) {
    throw new WorkspaceApiError('시트 탭을 찾을 수 없습니다.', 404);
  }

  const escapedRange = `'${targetTab.title.replace(/'/g, "''")}'`;
  const params = new URLSearchParams({
    ranges: escapedRange,
    includeGridData: 'true',
    fields:
      'properties.title,sheets(properties(sheetId,title,index,hidden),merges,data(columnMetadata(hiddenByUser),rowMetadata(hiddenByUser),rowData(values(formattedValue,effectiveFormat/textFormat/strikethrough,textFormatRuns(startIndex,format/strikethrough)))))',
  });

  const data = await fetchGoogleJson<{
    properties?: { title?: string };
    sheets?: Array<{
      merges?: GridRange[];
      data?: Array<{
        columnMetadata?: Array<{ hiddenByUser?: boolean }>;
        rowMetadata?: Array<{ hiddenByUser?: boolean }>;
        rowData?: Array<{ values?: SheetCellData[] }>;
      }>;
    }>;
  }>(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(resolvedId)}?${params.toString()}`,
    accessToken
  );

  const sheetObj = data.sheets?.[0];
  const gridData = sheetObj?.data?.[0];

  const columnHidden = (gridData?.columnMetadata ?? []).map((col) =>
    Boolean(col.hiddenByUser)
  );
  const rowHidden = (gridData?.rowMetadata ?? []).map((row) =>
    Boolean(row.hiddenByUser)
  );
  const rows = (gridData?.rowData ?? []).map((row) => row.values ?? []);

  return {
    spreadsheetTitle:
      data.properties?.title || driveMeta?.name || meta.title,
    fileKind: 'google-sheet',
    webViewLink: driveMeta?.webViewLink,
    tabs: meta.tabs,
    activeTab: targetTab,
    merges: sheetObj?.merges ?? [],
    columnHidden,
    rowHidden,
    rows,
  };
}
