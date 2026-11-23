import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import type {
  GridRange,
  SheetCellData,
  SheetGridPayload,
  SheetTabInfo,
  TextFormatRun,
} from './googleWorkspace';

interface ParsedSharedString {
  text: string;
  runs?: TextFormatRun[];
}

interface ParsedStyleTable {
  cellStrikeByXf: boolean[];
  cellDateByXf: boolean[];
}

interface WorkbookSheetRef {
  tab: SheetTabInfo;
  targetPath: string;
}

const BUILTIN_DATE_FMT_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
  45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
]);

function isExcelDateFormatCode(formatCode: string): boolean {
  const stripped = formatCode
    .replace(/"[^"]*"/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\\./g, '')
    .toLowerCase();
  return /[ymd]/.test(stripped);
}

function formatExcelSerialDate(serial: number): string {
  // Excel 1900 date system (accounting for 1900 leap year bug)
  const utcDays = Math.floor(serial - 25569);
  const utcMs = utcDays * 86400 * 1000;
  const d = new Date(utcMs);
  if (Number.isNaN(d.getTime())) return String(serial);
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function colLettersToIndex(letters: string): number {
  let idx = 0;
  for (let i = 0; i < letters.length; i++) {
    idx = idx * 26 + (letters.charCodeAt(i) - 64);
  }
  return idx - 1;
}

function parseCellRef(ref: string): { row: number; col: number } | null {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  return {
    col: colLettersToIndex(match[1].toUpperCase()),
    row: Number(match[2]) - 1,
  };
}

function parseRangeRef(ref: string): GridRange | null {
  const parts = ref.split(':');
  const start = parseCellRef(parts[0]);
  if (!start) return null;
  const end = parts[1] ? parseCellRef(parts[1]) : start;
  if (!end) return null;

  return {
    startRowIndex: Math.min(start.row, end.row),
    endRowIndex: Math.max(start.row, end.row) + 1,
    startColumnIndex: Math.min(start.col, end.col),
    endColumnIndex: Math.max(start.col, end.col) + 1,
  };
}

function getDirectChildrenByLocalName(
  parent: Element,
  localName: string
): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.children.length; i++) {
    const child = parent.children[i];
    if (child.localName === localName) {
      out.push(child);
    }
  }
  return out;
}

function getFirstDirectChild(
  parent: Element,
  localName: string
): Element | null {
  for (let i = 0; i < parent.children.length; i++) {
    const child = parent.children[i];
    if (child.localName === localName) {
      return child;
    }
  }
  return null;
}

function hasEnabledStrike(styleParent: Element | null): boolean {
  if (!styleParent) return false;
  const strikeEl = getFirstDirectChild(styleParent, 'strike');
  if (!strikeEl) return false;
  const val = strikeEl.getAttribute('val');
  if (val === '0' || val === 'false') return false;
  return true;
}

function parseRichStringElement(el: Element): ParsedSharedString {
  const runElements = getDirectChildrenByLocalName(el, 'r');
  if (runElements.length > 0) {
    let fullText = '';
    const runs: TextFormatRun[] = [];
    for (const runEl of runElements) {
      const rPr = getFirstDirectChild(runEl, 'rPr');
      const strike = hasEnabledStrike(rPr);
      const tEls = getDirectChildrenByLocalName(runEl, 't');
      const chunk = tEls.map((t) => t.textContent ?? '').join('');
      if (chunk.length > 0) {
        runs.push({
          startIndex: fullText.length,
          format: { strikethrough: strike },
        });
        fullText += chunk;
      }
    }
    return { text: fullText, runs: runs.length > 0 ? runs : undefined };
  }

  const tElements = getDirectChildrenByLocalName(el, 't');
  const text = tElements.map((t) => t.textContent ?? '').join('');
  return { text };
}

function parseXml(xmlText: string): Document {
  const parser = new DOMParser();
  return parser.parseFromString(xmlText, 'application/xml');
}

async function parseStyles(zip: JSZip): Promise<ParsedStyleTable> {
  const stylesFile = zip.file('xl/styles.xml');
  if (!stylesFile) {
    return { cellStrikeByXf: [], cellDateByXf: [] };
  }

  const doc = parseXml(await stylesFile.async('string'));

  // Parse custom number formats
  const customDateFmtIds = new Set<number>();
  const numFmtsEl = doc.getElementsByTagNameNS('*', 'numFmts')[0];
  if (numFmtsEl) {
    for (const numFmt of getDirectChildrenByLocalName(numFmtsEl, 'numFmt')) {
      const id = Number(numFmt.getAttribute('numFmtId'));
      const code = numFmt.getAttribute('formatCode') ?? '';
      if (!Number.isNaN(id) && isExcelDateFormatCode(code)) {
        customDateFmtIds.add(id);
      }
    }
  }

  // Parse fonts for strikethrough
  const fontStrikeFlags: boolean[] = [];
  const fontsEl = doc.getElementsByTagNameNS('*', 'fonts')[0];
  if (fontsEl) {
    for (const fontEl of getDirectChildrenByLocalName(fontsEl, 'font')) {
      fontStrikeFlags.push(hasEnabledStrike(fontEl));
    }
  }

  // Parse cellXfs mapping style index -> font strike & date format
  const cellStrikeByXf: boolean[] = [];
  const cellDateByXf: boolean[] = [];
  const cellXfsEl = doc.getElementsByTagNameNS('*', 'cellXfs')[0];
  if (cellXfsEl) {
    for (const xfEl of getDirectChildrenByLocalName(cellXfsEl, 'xf')) {
      const fontId = Number(xfEl.getAttribute('fontId') ?? '0');
      const numFmtId = Number(xfEl.getAttribute('numFmtId') ?? '0');
      cellStrikeByXf.push(Boolean(fontStrikeFlags[fontId]));
      cellDateByXf.push(
        BUILTIN_DATE_FMT_IDS.has(numFmtId) || customDateFmtIds.has(numFmtId)
      );
    }
  }

  return { cellStrikeByXf, cellDateByXf };
}

async function parseSharedStrings(zip: JSZip): Promise<ParsedSharedString[]> {
  const sstFile = zip.file('xl/sharedStrings.xml');
  if (!sstFile) return [];

  const doc = parseXml(await sstFile.async('string'));
  const root = doc.documentElement;
  if (!root) return [];

  const siList = getDirectChildrenByLocalName(root, 'si');
  return siList.map(parseRichStringElement);
}

function resolveWorkbookTargetPath(target: string): string {
  const cleaned = target.replace(/^\/+/, '');
  if (cleaned.startsWith('xl/')) return cleaned;
  if (cleaned.startsWith('../')) return cleaned.replace(/^(\.\.\/)+/, '');
  return `xl/${cleaned}`;
}

async function parseWorkbookSheetRefs(zip: JSZip): Promise<WorkbookSheetRef[]> {
  const wbFile = zip.file('xl/workbook.xml');
  const relsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (!wbFile || !relsFile) {
    throw new Error('올바른 Excel(.xlsx) 워크북 구조가 아닙니다.');
  }

  const wbDoc = parseXml(await wbFile.async('string'));
  const relsDoc = parseXml(await relsFile.async('string'));

  const relMap = new Map<string, string>();
  const relNodes = relsDoc.getElementsByTagNameNS('*', 'Relationship');
  for (let i = 0; i < relNodes.length; i++) {
    const rel = relNodes[i];
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && target) {
      relMap.set(id, resolveWorkbookTargetPath(target));
    }
  }

  const sheetNodes = wbDoc.getElementsByTagNameNS('*', 'sheet');
  const refs: WorkbookSheetRef[] = [];

  for (let i = 0; i < sheetNodes.length; i++) {
    const s = sheetNodes[i];
    const title = s.getAttribute('name') || `Sheet ${i + 1}`;
    const sheetId = Number(s.getAttribute('sheetId') ?? i + 1);
    const state = s.getAttribute('state') ?? 'visible';
    const rId =
      s.getAttribute('r:id') ||
      s.getAttributeNS(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'id'
      ) ||
      '';
    const targetPath = relMap.get(rId);
    if (!targetPath) continue;

    refs.push({
      tab: {
        sheetId: Number.isNaN(sheetId) ? i + 1 : sheetId,
        title,
        index: i,
        hidden: state === 'hidden' || state === 'veryHidden',
      },
      targetPath,
    });
  }

  return refs;
}

async function parseOoxmlBuffer(
  buffer: ArrayBuffer,
  spreadsheetTitle: string,
  requestedSheetTitle?: string
): Promise<SheetGridPayload> {
  const zip = await JSZip.loadAsync(buffer);
  const sheetRefs = await parseWorkbookSheetRefs(zip);
  if (sheetRefs.length === 0) {
    throw new Error('엑셀 파일에 시트 탭이 없습니다.');
  }

  const tabs = sheetRefs.map((r) => r.tab);
  const chosenRef =
    sheetRefs.find((r) => r.tab.title === requestedSheetTitle) ??
    sheetRefs.find((r) => !r.tab.hidden) ??
    sheetRefs[0];

  const sheetFile = zip.file(chosenRef.targetPath);
  if (!sheetFile) {
    throw new Error(`시트 데이터를 찾을 수 없습니다: ${chosenRef.tab.title}`);
  }

  const [styles, sharedStrings, sheetXml] = await Promise.all([
    parseStyles(zip),
    parseSharedStrings(zip),
    sheetFile.async('string'),
  ]);

  const doc = parseXml(sheetXml);

  // Hidden columns
  const columnHidden: boolean[] = [];
  const colsNodes = doc.getElementsByTagNameNS('*', 'cols');
  for (let i = 0; i < colsNodes.length; i++) {
    for (const colEl of getDirectChildrenByLocalName(colsNodes[i], 'col')) {
      const hiddenAttr = colEl.getAttribute('hidden');
      const isHidden = hiddenAttr === '1' || hiddenAttr === 'true';
      if (!isHidden) continue;

      const min = Math.max(1, Number(colEl.getAttribute('min') ?? '1'));
      const max = Math.min(
        500,
        Math.max(min, Number(colEl.getAttribute('max') ?? String(min)))
      );
      for (let c = min - 1; c <= max - 1; c++) {
        columnHidden[c] = true;
      }
    }
  }

  // Merged cells
  const merges: GridRange[] = [];
  const mergeNodes = doc.getElementsByTagNameNS('*', 'mergeCell');
  for (let i = 0; i < mergeNodes.length; i++) {
    const ref = mergeNodes[i].getAttribute('ref');
    if (!ref) continue;
    const range = parseRangeRef(ref);
    if (range) merges.push(range);
  }

  // Cell rows
  const rows: SheetCellData[][] = [];
  const sheetDataEl = doc.getElementsByTagNameNS('*', 'sheetData')[0];
  if (sheetDataEl) {
    const rowNodes = getDirectChildrenByLocalName(sheetDataEl, 'row');
    let implicitRowIdx = 0;

    for (const rowEl of rowNodes) {
      const rAttr = rowEl.getAttribute('r');
      const rowIdx = rAttr ? Number(rAttr) - 1 : implicitRowIdx;
      implicitRowIdx = rowIdx + 1;
      if (rowIdx < 0 || rowIdx > 5000) continue;

      if (!rows[rowIdx]) rows[rowIdx] = [];

      const cellNodes = getDirectChildrenByLocalName(rowEl, 'c');
      let implicitColIdx = 0;

      for (const cEl of cellNodes) {
        const refAttr = cEl.getAttribute('r');
        const parsedRef = refAttr ? parseCellRef(refAttr) : null;
        const colIdx = parsedRef ? parsedRef.col : implicitColIdx;
        implicitColIdx = colIdx + 1;
        if (colIdx < 0 || colIdx > 1000) continue;

        const cellType = cEl.getAttribute('t') ?? '';
        const styleIdxAttr = cEl.getAttribute('s');
        const styleIdx = styleIdxAttr !== null ? Number(styleIdxAttr) : -1;
        const cellStrike =
          styleIdx >= 0 ? Boolean(styles.cellStrikeByXf[styleIdx]) : false;
        const isDateStyle =
          styleIdx >= 0 ? Boolean(styles.cellDateByXf[styleIdx]) : false;

        let formattedValue = '';
        let textFormatRuns: TextFormatRun[] | undefined;

        if (cellType === 's') {
          const vEl = getFirstDirectChild(cEl, 'v');
          const sstIdx = Number(vEl?.textContent ?? '-1');
          const sst = sharedStrings[sstIdx];
          if (sst) {
            formattedValue = sst.text;
            textFormatRuns = sst.runs;
          }
        } else if (cellType === 'inlineStr') {
          const isEl = getFirstDirectChild(cEl, 'is');
          if (isEl) {
            const parsedInline = parseRichStringElement(isEl);
            formattedValue = parsedInline.text;
            textFormatRuns = parsedInline.runs;
          }
        } else if (cellType === 'b') {
          const vEl = getFirstDirectChild(cEl, 'v');
          formattedValue = vEl?.textContent === '1' ? 'TRUE' : 'FALSE';
        } else {
          const vEl = getFirstDirectChild(cEl, 'v');
          const rawVal = vEl?.textContent ?? '';
          if (rawVal !== '') {
            if (isDateStyle && cellType !== 'str' && cellType !== 'e') {
              const numVal = Number(rawVal);
              formattedValue = Number.isFinite(numVal)
                ? formatExcelSerialDate(numVal)
                : rawVal;
            } else {
              formattedValue = rawVal;
            }
          }
        }

        if (formattedValue !== '' || cellStrike) {
          const cellObj: SheetCellData = { formattedValue };
          if (cellStrike) {
            cellObj.effectiveFormat = { textFormat: { strikethrough: true } };
          }
          if (textFormatRuns && textFormatRuns.length > 0) {
            cellObj.textFormatRuns = textFormatRuns;
          }
          rows[rowIdx][colIdx] = cellObj;
        }
      }
    }
  }

  // Ensure dense row and cell arrays with zero undefined holes
  for (let r = 0; r < rows.length; r++) {
    if (!rows[r]) {
      rows[r] = [];
      continue;
    }
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      if (!row[c]) {
        row[c] = {};
      }
    }
  }

  return {
    spreadsheetTitle,
    tabs,
    activeTab: chosenRef.tab,
    merges,
    columnHidden,
    rows,
  };
}

function parseLegacyExcelWithSheetJS(
  buffer: ArrayBuffer,
  spreadsheetTitle: string,
  requestedSheetTitle?: string
): SheetGridPayload {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: false });
  const sheetProps = wb.Workbook?.Sheets ?? [];

  const tabs: SheetTabInfo[] = wb.SheetNames.map((title, idx) => ({
    sheetId: idx + 1,
    title,
    index: idx,
    hidden: Boolean(sheetProps[idx]?.Hidden),
  }));

  if (tabs.length === 0) {
    throw new Error('엑셀 파일에 시트 탭이 없습니다.');
  }

  const activeTab =
    tabs.find((t) => t.title === requestedSheetTitle) ??
    tabs.find((t) => !t.hidden) ??
    tabs[0];

  const ws = wb.Sheets[activeTab.title];
  const merges: GridRange[] = (ws['!merges'] ?? []).map((m) => ({
    startRowIndex: m.s.r,
    endRowIndex: m.e.r + 1,
    startColumnIndex: m.s.c,
    endColumnIndex: m.e.c + 1,
  }));

  const columnHidden: boolean[] = (ws['!cols'] ?? []).map((col) =>
    Boolean(col?.hidden)
  );

  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(ws, {
    header: 1,
    defval: '',
    raw: false,
  });

  const rows: SheetCellData[][] = rawRows.map((row) =>
    (row ?? []).map((val) => ({
      formattedValue:
        val === null || val === undefined ? '' : String(val),
    }))
  );

  return {
    spreadsheetTitle,
    tabs,
    activeTab,
    merges,
    columnHidden,
    rows,
  };
}

export async function parseExcelBufferToGrid(
  buffer: ArrayBuffer,
  spreadsheetTitle: string,
  requestedSheetTitle?: string
): Promise<SheetGridPayload> {
  try {
    return await parseOoxmlBuffer(
      buffer,
      spreadsheetTitle,
      requestedSheetTitle
    );
  } catch {
    return parseLegacyExcelWithSheetJS(
      buffer,
      spreadsheetTitle,
      requestedSheetTitle
    );
  }
}
