import { GridRange, SheetCellData, SheetGridPayload } from './googleWorkspace';

export interface RichTextLine {
  text: string;
  strikethrough: boolean;
}

export interface ExecutionWeekColumn {
  colIndex: number;
  colLetter: string;
  headerLabel: string;
  dateLabel: string;
  hidden: boolean;
}

export function indexToColLetters(colIndex: number): string {
  let n = colIndex + 1;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

export interface ProjectEntry {
  key: string;
  numberPrefix: string;
  projectTitle: string;
  owners: RichTextLine[];
  plannedTasks: RichTextLine[];
  issues: RichTextLine[];
  executionByCol: Record<number, RichTextLine[]>;
}

export interface TeamReport {
  id: string;
  rawTeamLabel: string;
  teamName: string;
  leaderName: string;
  projects: ProjectEntry[];
  extraOwners: RichTextLine[];
  extraPlanned: RichTextLine[];
  extraIssues: RichTextLine[];
  extraExecutionByCol: Record<number, RichTextLine[]>;
}

export interface ParsedWeeklyReport {
  reportTitle: string;
  organization: string;
  reportSchedule: string;
  headerRowIndex: number;
  executionColumns: ExecutionWeekColumn[];
  defaultExecutionColIndex: number | null;
  teams: TeamReport[];
}

function normalizeHeader(text: string): string {
  return text.replace(/\s+/g, '').toLowerCase();
}

function isPlausibleHeaderLabel(raw: string): boolean {
  const trimmed = raw.trim();
  if (!trimmed) return false;
  if (trimmed.length > 45) return false;
  if (/[\r\n]/.test(trimmed)) return false;
  // Numbered body items like "1) 480kW..." are never table column headers
  if (/^\d{1,2}\)\s*/.test(trimmed)) return false;
  return true;
}

function extractRichLines(cell?: SheetCellData): RichTextLine[] {
  if (!cell || !cell.formattedValue) return [];
  const raw = cell.formattedValue;
  const cellStrike = Boolean(cell.effectiveFormat?.textFormat?.strikethrough);
  const runs = cell.textFormatRuns;

  if (!runs || runs.length === 0) {
    return raw
      .split(/\r?\n/)
      .map((line) => line.trimEnd())
      .filter((line) => line.trim().length > 0)
      .map((text) => ({ text, strikethrough: cellStrike }));
  }

  const lines: RichTextLine[] = [];
  let charOffset = 0;
  const rawLines = raw.split(/\r?\n/);

  for (const lineText of rawLines) {
    const lineStart = charOffset;
    const lineEnd = charOffset + lineText.length;
    // Advance accurately past either \r\n (2 chars) or \n (1 char)
    if (raw[lineEnd] === '\r' && raw[lineEnd + 1] === '\n') {
      charOffset = lineEnd + 2;
    } else {
      charOffset = lineEnd + 1;
    }

    if (lineText.trim().length === 0) continue;

    let struckChars = 0;
    let totalNonSpace = 0;

    for (let i = lineStart; i < lineEnd; i++) {
      const ch = raw[i];
      if (!ch || /\s/.test(ch)) continue;
      totalNonSpace++;

      let activeRunStrike = cellStrike;
      for (let r = 0; r < runs.length; r++) {
        const runStart = runs[r].startIndex ?? 0;
        const nextRunStart =
          r + 1 < runs.length ? (runs[r + 1].startIndex ?? Infinity) : Infinity;
        if (i >= runStart && i < nextRunStart) {
          if (runs[r].format?.strikethrough !== undefined) {
            activeRunStrike = Boolean(runs[r].format?.strikethrough);
          }
          break;
        }
      }
      if (activeRunStrike) struckChars++;
    }

    const isLineStruck =
      totalNonSpace > 0 ? struckChars / totalNonSpace >= 0.7 : cellStrike;

    lines.push({
      text: lineText.trimEnd(),
      strikethrough: isLineStruck,
    });
  }

  return lines;
}

interface NumberedChunk {
  num: string | null;
  subPrefix: string | null;
  lines: RichTextLine[];
}

const TOP_LEVEL_NUM_RE = /^(\d{1,2})(?:[)）]\s*|\.\s+)(.*)$/;
const SUB_LEVEL_NUM_RE = /^(\d{1,2})\s*[-–—]\s*(\d{1,2})(?:[)）]\s*|\.\s+)(.*)$/;

function splitIntoNumberedChunks(
  lines: RichTextLine[],
  options?: { allowSubNumber?: boolean }
): NumberedChunk[] {
  const allowSubNumber = options?.allowSubNumber ?? true;
  const chunks: NumberedChunk[] = [];
  let current: NumberedChunk | null = null;

  for (const line of lines) {
    const trimmed = line.text.trim();

    if (allowSubNumber) {
      const subMatch = trimmed.match(SUB_LEVEL_NUM_RE);
      if (subMatch) {
        if (current) chunks.push(current);
        const majorNum = subMatch[1];
        const minorNum = subMatch[2];
        const rest = subMatch[3].trim();
        const normalizedPrefix = `${majorNum}-${minorNum})`;
        current = {
          num: majorNum,
          subPrefix: normalizedPrefix,
          lines: [
            {
              text: rest ? `${normalizedPrefix} ${rest}` : normalizedPrefix,
              strikethrough: line.strikethrough,
            },
          ],
        };
        continue;
      }
    }

    const topMatch = trimmed.match(TOP_LEVEL_NUM_RE);
    if (topMatch) {
      if (current) chunks.push(current);
      const majorNum = topMatch[1];
      const rest = topMatch[2].trim();
      current = {
        num: majorNum,
        subPrefix: null,
        lines: [
          {
            text: rest ? `${majorNum}) ${rest}` : trimmed,
            strikethrough: line.strikethrough,
          },
        ],
      };
    } else {
      if (!current) {
        current = { num: null, subPrefix: null, lines: [line] };
      } else {
        current.lines.push(line);
      }
    }
  }

  if (current) chunks.push(current);
  return chunks;
}

function parseTeamAndLeader(raw: string): { teamName: string; leaderName: string } {
  const cleaned = raw.replace(/\s+/g, ' ').trim();
  const parenMatch = cleaned.match(/^(.+?)\s*\(([^)]+)\)\s*$/);
  if (parenMatch) {
    return {
      teamName: parenMatch[1].trim(),
      leaderName: parenMatch[2].trim(),
    };
  }
  return {
    teamName: cleaned,
    leaderName: '',
  };
}

function extractDateFromExecHeader(header: string): string {
  const match = header.match(/\(([^)]+)\)/);
  if (match?.[1]) return match[1].trim();
  return header.replace(/업무\s*실행/i, '').trim() || header;
}

function isExecutionHeader(raw: string): boolean {
  if (!isPlausibleHeaderLabel(raw)) return false;
  const norm = normalizeHeader(raw);
  if (norm.includes('업무실행') || norm.includes('실행')) return true;
  // Also match date-only weekly execution columns like "11/20" or "(11/20)"
  if (/^\(?\d{1,2}[./-]\d{1,2}\)?$/.test(norm)) return true;
  return false;
}

function findHeaderRow(rows: SheetCellData[][]): number {
  const maxScan = Math.min(rows.length, 25);
  let bestRow = -1;
  let bestScore = 0;

  for (let r = 0; r < maxScan; r++) {
    const row = rows[r] ?? [];

    // Score by DISTINCT header column roles (never let multi-line body paragraphs inflate score)
    let hasTeam = false;
    let hasProject = false;
    let hasOwner = false;
    let hasPlanned = false;
    let hasIssue = false;
    let hasExec = false;
    let positionalBonus = 0;

    for (let c = 0; c < row.length; c++) {
      const raw = row[c]?.formattedValue?.trim() ?? '';
      if (!isPlausibleHeaderLabel(raw)) continue;
      const norm = normalizeHeader(raw);

      if (!hasTeam && (norm === '팀' || norm.startsWith('팀(') || norm.includes('팀명') || norm.includes('조직'))) {
        hasTeam = true;
        if (c === 0) positionalBonus += 3;
      } else if (!hasProject && norm.includes('프로젝트')) {
        hasProject = true;
        if (c === 1) positionalBonus += 3;
      } else if (!hasOwner && norm.includes('담당자')) {
        hasOwner = true;
        if (c === 2) positionalBonus += 2;
      } else if (
        !hasPlanned &&
        (norm.includes('업무진행예정') || norm.includes('진행예정') || norm.includes('추진일정'))
      ) {
        hasPlanned = true;
        if (c === 3) positionalBonus += 2;
      } else if (!hasIssue && norm.includes('이슈')) {
        hasIssue = true;
        if (c === 4) positionalBonus += 2;
      } else if (c >= 4 && isExecutionHeader(raw)) {
        hasExec = true;
      }
    }

    let score = positionalBonus;
    if (hasTeam) score += 5;
    if (hasProject) score += 5;
    if (hasOwner) score += 4;
    if (hasPlanned) score += 4;
    if (hasIssue) score += 3;
    if (hasExec) score += 5;

    // Require at least 2 distinct core header columns (score >= 10)
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }

  return bestScore >= 10 ? bestRow : -1;
}

function getMergeEndRow(
  merges: GridRange[],
  rowIndex: number,
  colIndex: number
): number {
  for (const m of merges) {
    const sr = m.startRowIndex ?? -1;
    const er = m.endRowIndex ?? -1;
    const sc = m.startColumnIndex ?? -1;
    const ec = m.endColumnIndex ?? -1;
    if (rowIndex >= sr && rowIndex < er && colIndex >= sc && colIndex < ec) {
      return er - 1;
    }
  }
  return rowIndex;
}

export function parseWeeklyReportGrid(grid: SheetGridPayload): ParsedWeeklyReport {
  const { rows, merges, columnHidden, rowHidden = [], spreadsheetTitle, activeTab } = grid;
  const headerRowIndex = findHeaderRow(rows);

  // Extract metadata from rows above the table header
  let reportTitle = `${spreadsheetTitle} · ${activeTab.title}`;
  let organization = '';
  let reportSchedule = '';

  const metaEndRow = headerRowIndex > 0 ? headerRowIndex : Math.min(rows.length, 8);
  for (let r = 0; r < metaEndRow; r++) {
    const rowTexts = (rows[r] ?? [])
      .map((c) => c?.formattedValue?.trim() ?? '')
      .filter(Boolean);
    if (rowTexts.length === 0) continue;

    const joined = rowTexts.join(' ').trim();
    const compact = joined.replace(/\s+/g, '');

    if (compact.includes('조직명')) {
      const m = joined.match(/조\s*직\s*명\s*[:：]?\s*(.+)/);
      organization = m?.[1]?.trim() || joined;
    } else if (compact.includes('보고일시') || compact.includes('일시')) {
      const m = joined.match(/보고\s*일시\s*[:：]?\s*(.+)/);
      reportSchedule = m?.[1]?.trim() || joined;
    } else if (
      (compact.includes('주간업무보고') || compact.includes('주간보고')) &&
      joined.length > 3
    ) {
      reportTitle = joined;
    }
  }

  if (headerRowIndex === -1) {
    return {
      reportTitle,
      organization,
      reportSchedule,
      headerRowIndex: -1,
      executionColumns: [],
      defaultExecutionColIndex: null,
      teams: [],
    };
  }

  const headerRow = rows[headerRowIndex] ?? [];
  let teamCol = -1;
  let projectCol = -1;
  let ownerCol = -1;
  let plannedCol = -1;
  let issueCol = -1;
  const executionColumns: ExecutionWeekColumn[] = [];

  for (let c = 0; c < headerRow.length; c++) {
    const rawHeader = headerRow[c]?.formattedValue?.trim() ?? '';
    if (!isPlausibleHeaderLabel(rawHeader)) continue;
    const norm = normalizeHeader(rawHeader);

    if (
      teamCol === -1 &&
      c <= 2 &&
      (norm === '팀' || norm.startsWith('팀(') || norm.includes('팀명') || norm.includes('조직'))
    ) {
      teamCol = c;
    } else if (projectCol === -1 && c <= 3 && norm.includes('프로젝트')) {
      projectCol = c;
    } else if (ownerCol === -1 && c <= 4 && norm.includes('담당자')) {
      ownerCol = c;
    } else if (
      plannedCol === -1 &&
      c <= 5 &&
      (norm.includes('업무진행예정') || norm.includes('진행예정') || norm.includes('추진일정') || norm.includes('예정'))
    ) {
      plannedCol = c;
    } else if (issueCol === -1 && c <= 6 && norm.includes('이슈')) {
      issueCol = c;
    } else if (isExecutionHeader(rawHeader)) {
      executionColumns.push({
        colIndex: c,
        colLetter: indexToColLetters(c),
        headerLabel: rawHeader,
        dateLabel: extractDateFromExecHeader(rawHeader),
        hidden: Boolean(columnHidden[c]),
      });
    }
  }

  // Enforce standard A/B/C/D/E defaults when headers are standard or partially blank
  if (teamCol === -1) teamCol = 0;
  if (projectCol === -1) projectCol = 1;
  if (ownerCol === -1 && headerRow.length > 2) ownerCol = 2;
  if (plannedCol === -1 && headerRow.length > 3) plannedCol = 3;
  if (issueCol === -1 && headerRow.length > 4) issueCol = 4;

  // Fallback: if no execution column header matched explicitly, scan columns after E (col >= 5) that have data
  if (executionColumns.length === 0) {
    let maxCol = headerRow.length - 1;
    for (let r = headerRowIndex + 1; r < rows.length; r++) {
      maxCol = Math.max(maxCol, (rows[r]?.length ?? 0) - 1);
    }
    for (let c = Math.max(5, issueCol + 1); c <= maxCol; c++) {
      let colHasData = false;
      for (let r = headerRowIndex + 1; r < rows.length; r++) {
        if (rows[r]?.[c]?.formattedValue?.trim()) {
          colHasData = true;
          break;
        }
      }
      const hdr = headerRow[c]?.formattedValue?.trim() || `업무 실행(${indexToColLetters(c)})`;
      if (colHasData || headerRow[c]?.formattedValue?.trim()) {
        executionColumns.push({
          colIndex: c,
          colLetter: indexToColLetters(c),
          headerLabel: hdr,
          dateLabel: extractDateFromExecHeader(hdr),
          hidden: Boolean(columnHidden[c]),
        });
      }
    }
  }

  // Pick the latest (rightmost) visible execution column with data (e.g. AU: 업무 실행(11/20))
  const visibleExecCols = executionColumns.filter((col) => !col.hidden);
  let defaultExecutionColIndex: number | null = null;

  const hasDataInCol = (colIdx: number) => {
    for (let r = headerRowIndex + 1; r < rows.length; r++) {
      if (rows[r]?.[colIdx]?.formattedValue?.trim()) return true;
    }
    return false;
  };

  for (let i = visibleExecCols.length - 1; i >= 0; i--) {
    if (hasDataInCol(visibleExecCols[i].colIndex)) {
      defaultExecutionColIndex = visibleExecCols[i].colIndex;
      break;
    }
  }
  if (defaultExecutionColIndex === null && executionColumns.length > 0) {
    for (let i = executionColumns.length - 1; i >= 0; i--) {
      if (hasDataInCol(executionColumns[i].colIndex)) {
        defaultExecutionColIndex = executionColumns[i].colIndex;
        break;
      }
    }
  }
  if (defaultExecutionColIndex === null && visibleExecCols.length > 0) {
    defaultExecutionColIndex = visibleExecCols[visibleExecCols.length - 1].colIndex;
  }
  if (defaultExecutionColIndex === null && executionColumns.length > 0) {
    defaultExecutionColIndex =
      executionColumns[executionColumns.length - 1].colIndex;
  }

  // Group data rows into teams starting right after headerRowIndex (e.g. Row 8 -> r = 7)
  const teams: TeamReport[] = [];
  let r = headerRowIndex + 1;

  while (r < rows.length) {
    if (rowHidden[r]) {
      r++;
      continue;
    }

    const row = rows[r] ?? [];
    const teamText = row[teamCol]?.formattedValue?.trim() ?? '';

    if (!teamText) {
      r++;
      continue;
    }

    // Consume merged rows under this team cell in A column
    const mergeEnd = getMergeEndRow(merges, r, teamCol);
    let endRow = mergeEnd;

    // Also consume immediately contiguous rows where A column is empty but B/C/D/E/exec has content
    let probe = Math.max(r, mergeEnd) + 1;
    while (probe < rows.length) {
      const probeRow = rows[probe] ?? [];
      const probeTeam = probeRow[teamCol]?.formattedValue?.trim() ?? '';
      if (probeTeam) break;

      const probeHasRelevantContent =
        Boolean(probeRow[projectCol]?.formattedValue?.trim()) ||
        (ownerCol >= 0 && Boolean(probeRow[ownerCol]?.formattedValue?.trim())) ||
        (plannedCol >= 0 && Boolean(probeRow[plannedCol]?.formattedValue?.trim())) ||
        (issueCol >= 0 && Boolean(probeRow[issueCol]?.formattedValue?.trim())) ||
        (defaultExecutionColIndex !== null &&
          Boolean(probeRow[defaultExecutionColIndex]?.formattedValue?.trim()));

      if (!probeHasRelevantContent) break;
      endRow = probe;
      probe++;
    }

    const rawTeamLabel = teamText
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    const { teamName, leaderName } = parseTeamAndLeader(rawTeamLabel);

    const projectLinesAll: RichTextLine[] = [];
    const ownerLinesAll: RichTextLine[] = [];
    const plannedLinesAll: RichTextLine[] = [];
    const issueLinesAll: RichTextLine[] = [];
    const execLinesByCol: Record<number, RichTextLine[]> = {};
    for (const ec of executionColumns) {
      execLinesByCol[ec.colIndex] = [];
    }

    for (let rr = r; rr <= endRow; rr++) {
      const curRow = rows[rr] ?? [];
      projectLinesAll.push(...extractRichLines(curRow[projectCol]));
      if (ownerCol >= 0) ownerLinesAll.push(...extractRichLines(curRow[ownerCol]));
      if (plannedCol >= 0) plannedLinesAll.push(...extractRichLines(curRow[plannedCol]));
      if (issueCol >= 0) issueLinesAll.push(...extractRichLines(curRow[issueCol]));
      for (const ec of executionColumns) {
        execLinesByCol[ec.colIndex].push(...extractRichLines(curRow[ec.colIndex]));
      }
    }

    // Split B column by top-level project numbers (1), 2), ...),
    // and split C/D/E/AU columns by both top-level (1), 2)) and sub-numbers (1-1), 2-1), ...)
    const projectChunks = splitIntoNumberedChunks(projectLinesAll, {
      allowSubNumber: false,
    });
    const ownerChunks = splitIntoNumberedChunks(ownerLinesAll, {
      allowSubNumber: true,
    });
    const plannedChunks = splitIntoNumberedChunks(plannedLinesAll, {
      allowSubNumber: true,
    });
    const issueChunks = splitIntoNumberedChunks(issueLinesAll, {
      allowSubNumber: true,
    });
    const execChunksByCol: Record<number, NumberedChunk[]> = {};
    for (const ec of executionColumns) {
      execChunksByCol[ec.colIndex] = splitIntoNumberedChunks(
        execLinesByCol[ec.colIndex],
        { allowSubNumber: true }
      );
    }

    const projects: ProjectEntry[] = [];
    const extraOwners: RichTextLine[] = [];
    const extraPlanned: RichTextLine[] = [];
    const extraIssues: RichTextLine[] = [];
    const extraExecutionByCol: Record<number, RichTextLine[]> = {};
    for (const ec of executionColumns) {
      extraExecutionByCol[ec.colIndex] = [];
    }

    const numberedProjects = projectChunks.filter((ch) => ch.num !== null);

    if (numberedProjects.length > 0) {
      const usedNums = new Set<string>();

      for (const pChunk of projectChunks) {
        if (pChunk.num === null) {
          const title = pChunk.lines.map((l) => l.text).join(' ').trim();
          if (title) {
            projects.push({
              key: `${r}-unnumbered-${projects.length}`,
              numberPrefix: '-',
              projectTitle: title,
              owners: [],
              plannedTasks: [],
              issues: [],
              executionByCol: {},
            });
          }
          continue;
        }

        const num = pChunk.num;
        usedNums.add(num);

        const rawProjText = pChunk.lines.map((l) => l.text).join(' ').trim();
        const cleanTitle = rawProjText
          .replace(/^\d{1,2}(?:[)）]\s*|\.\s+)/, '')
          .trim();

        const collectMatchedLines = (chunks: NumberedChunk[]): RichTextLine[] => {
          const matched = chunks.filter((c) => c.num === num);
          const result: RichTextLine[] = [];
          for (const m of matched) {
            m.lines.forEach((line, idx) => {
              if (idx === 0 && !m.subPrefix) {
                const cleaned = line.text
                  .replace(/^\d{1,2}(?:[)）]\s*|\.\s+)/, '')
                  .trim();
                if (cleaned) result.push({ ...line, text: cleaned });
              } else {
                // Keep sub-number prefixes like "1-1)", "1-2)", "2-1)" intact
                result.push(line);
              }
            });
          }
          return result;
        };

        const execMap: Record<number, RichTextLine[]> = {};
        for (const ec of executionColumns) {
          execMap[ec.colIndex] = collectMatchedLines(execChunksByCol[ec.colIndex]);
        }

        projects.push({
          key: `${r}-proj-${num}-${projects.length}`,
          numberPrefix: num,
          projectTitle: cleanTitle || rawProjText,
          owners: collectMatchedLines(ownerChunks),
          plannedTasks: collectMatchedLines(plannedChunks),
          issues: collectMatchedLines(issueChunks),
          executionByCol: execMap,
        });
      }

      for (const ch of ownerChunks) {
        if (ch.num === null || !usedNums.has(ch.num)) {
          extraOwners.push(...ch.lines);
        }
      }
      for (const ch of plannedChunks) {
        if (ch.num === null || !usedNums.has(ch.num)) {
          extraPlanned.push(...ch.lines);
        }
      }
      for (const ch of issueChunks) {
        if (ch.num === null || !usedNums.has(ch.num)) {
          extraIssues.push(...ch.lines);
        }
      }
      for (const ec of executionColumns) {
        for (const ch of execChunksByCol[ec.colIndex]) {
          if (ch.num === null || !usedNums.has(ch.num)) {
            extraExecutionByCol[ec.colIndex].push(...ch.lines);
          }
        }
      }
    } else {
      const projTitle =
        projectLinesAll.map((l) => l.text).join(' / ').trim() || '주요 업무 내용';
      const execMap: Record<number, RichTextLine[]> = {};
      for (const ec of executionColumns) {
        execMap[ec.colIndex] = execLinesByCol[ec.colIndex];
      }
      projects.push({
        key: `${r}-single`,
        numberPrefix: '1',
        projectTitle: projTitle,
        owners: ownerLinesAll,
        plannedTasks: plannedLinesAll,
        issues: issueLinesAll,
        executionByCol: execMap,
      });
    }

    teams.push({
      id: `team-${teams.length}-${r}`,
      rawTeamLabel,
      teamName,
      leaderName,
      projects,
      extraOwners,
      extraPlanned,
      extraIssues,
      extraExecutionByCol,
    });

    r = endRow + 1;
  }

  return {
    reportTitle,
    organization,
    reportSchedule,
    headerRowIndex,
    executionColumns,
    defaultExecutionColIndex,
    teams,
  };
}
