export interface DriveSpreadsheetFile {
  id: string;
  name: string;
  modifiedTime: string;
  webViewLink?: string;
  ownerName?: string;
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
  tabs: SheetTabInfo[];
  activeTab: SheetTabInfo;
  merges: GridRange[];
  columnHidden: boolean[];
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

export function extractSpreadsheetId(input: string): string | null {
  const trimmed = input.trim();
  if (!trimmed) return null;

  const urlMatch = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch?.[1]) {
    return urlMatch[1];
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

export async function listDriveSpreadsheets(
  accessToken: string,
  searchQuery = ''
): Promise<DriveSpreadsheetFile[]> {
  const qParts = [
    "mimeType='application/vnd.google-apps.spreadsheet'",
    'trashed=false',
  ];

  const cleanQuery = searchQuery.trim().replace(/'/g, "\\'");
  if (cleanQuery) {
    qParts.push(`name contains '${cleanQuery}'`);
  }

  const params = new URLSearchParams({
    q: qParts.join(' and '),
    orderBy: 'modifiedTime desc',
    pageSize: '40',
    supportsAllDrives: 'true',
    includeItemsFromAllDrives: 'true',
    fields: 'files(id,name,modifiedTime,webViewLink,owners(displayName,emailAddress))',
  });

  const data = await fetchGoogleJson<{
    files?: Array<{
      id: string;
      name: string;
      modifiedTime: string;
      webViewLink?: string;
      owners?: Array<{ displayName?: string; emailAddress?: string }>;
    }>;
  }>(`https://www.googleapis.com/drive/v3/files?${params.toString()}`, accessToken);

  return (data.files ?? []).map((f) => ({
    id: f.id,
    name: f.name,
    modifiedTime: f.modifiedTime,
    webViewLink: f.webViewLink,
    ownerName: f.owners?.[0]?.displayName || f.owners?.[0]?.emailAddress,
  }));
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
  sheetTitle: string
): Promise<SheetGridPayload> {
  const meta = await fetchSpreadsheetTabs(accessToken, spreadsheetId);
  const targetTab =
    meta.tabs.find((t) => t.title === sheetTitle) ?? meta.tabs[0];

  if (!targetTab) {
    throw new WorkspaceApiError('시트 탭을 찾을 수 없습니다.', 404);
  }

  const escapedRange = `'${targetTab.title.replace(/'/g, "''")}'`;
  const params = new URLSearchParams({
    ranges: escapedRange,
    includeGridData: 'true',
    fields:
      'properties.title,sheets(properties(sheetId,title,index,hidden),merges,data(columnMetadata(hiddenByUser),rowData(values(formattedValue,effectiveFormat/textFormat/strikethrough,textFormatRuns(startIndex,format/strikethrough)))))',
  });

  const data = await fetchGoogleJson<{
    properties?: { title?: string };
    sheets?: Array<{
      merges?: GridRange[];
      data?: Array<{
        columnMetadata?: Array<{ hiddenByUser?: boolean }>;
        rowData?: Array<{ values?: SheetCellData[] }>;
      }>;
    }>;
  }>(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}?${params.toString()}`,
    accessToken
  );

  const sheetObj = data.sheets?.[0];
  const gridData = sheetObj?.data?.[0];

  const columnHidden = (gridData?.columnMetadata ?? []).map((col) =>
    Boolean(col.hiddenByUser)
  );
  const rows = (gridData?.rowData ?? []).map((row) => row.values ?? []);

  return {
    spreadsheetTitle: data.properties?.title || meta.title,
    tabs: meta.tabs,
    activeTab: targetTab,
    merges: sheetObj?.merges ?? [],
    columnHidden,
    rows,
  };
}
