import { GridRange, SheetCellData, SheetGridPayload } from './googleWorkspace';

export interface RichTextLine {
  text: string;
  strikethrough: boolean;
}

export interface ExecutionWeekColumn {
  colIndex: number;
  headerLabel: string;
  dateLabel: string;
  hidden: boolean;
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
    charOffset = lineEnd + 1;

    if (lineText.trim().length === 0) continue;

    // Determine whether the non-whitespace characters of this line are predominantly struck through
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
  lines: RichTextLine[];
}

function splitIntoNumberedChunks(lines: RichTextLine[]): NumberedChunk[] {
  const chunks: NumberedChunk[] = [];
  let current: NumberedChunk | null = null;

  for (const line of lines) {
    const trimmed = line.text.trim();
    const match = trimmed.match(/^(\d{1,2})\s*[).]\s*(.*)$/);
    if (match) {
      if (current) chunks.push(current);
      current = {
        num: match[1],
        lines: [
          {
            text: match[2].trim() ? `${match[1]}) ${match[2].trim()}` : trimmed,
            strikethrough: line.strikethrough,
          },
        ],
      };
    } else {
      if (!current) {
        current = { num: null, lines: [line] };
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

function findHeaderRow(rows: SheetCellData[][]): number {
  const maxScan = Math.min(rows.length, 25);
  let bestRow = -1;
  let bestScore = 0;

  for (let r = 0; r < maxScan; r++) {
    const row = rows[r] ?? [];
    let score = 0;
    for (const cell of row) {
      const norm = normalizeHeader(cell.formattedValue ?? '');
      if (!norm) continue;
      if (norm.includes('팀')) score += 2;
      if (norm.includes('프로젝트')) score += 3;
      if (norm.includes('담당자')) score += 2;
      if (norm.includes('업무진행예정') || norm.includes('진행예정')) score += 2;
      if (norm.includes('이슈')) score += 2;
      if (norm.includes('업무실행') || norm.includes('실행')) score += 3;
    }
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }

  return bestScore >= 4 ? bestRow : -1;
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
  const { rows, merges, columnHidden, spreadsheetTitle, activeTab } = grid;
  const headerRowIndex = findHeaderRow(rows);

  // Extract metadata from rows above the table header
  let reportTitle = `${spreadsheetTitle} · ${activeTab.title}`;
  let organization = '';
  let reportSchedule = '';

  const metaEndRow = headerRowIndex > 0 ? headerRowIndex : Math.min(rows.length, 8);
  for (let r = 0; r < metaEndRow; r++) {
    const rowTexts = (rows[r] ?? [])
      .map((c) => c.formattedValue?.trim() ?? '')
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
    if (!rawHeader) continue;
    const norm = normalizeHeader(rawHeader);

    if (teamCol === -1 && norm.includes('팀')) {
      teamCol = c;
    } else if (projectCol === -1 && norm.includes('프로젝트')) {
      projectCol = c;
    } else if (ownerCol === -1 && norm.includes('담당자')) {
      ownerCol = c;
    } else if (
      plannedCol === -1 &&
      (norm.includes('업무진행예정') || norm.includes('진행예정') || norm.includes('예정'))
    ) {
      plannedCol = c;
    } else if (issueCol === -1 && norm.includes('이슈')) {
      issueCol = c;
    } else if (norm.includes('업무실행') || norm.includes('실행')) {
      executionColumns.push({
        colIndex: c,
        headerLabel: rawHeader,
        dateLabel: extractDateFromExecHeader(rawHeader),
        hidden: Boolean(columnHidden[c]),
      });
    }
  }

  if (teamCol === -1) teamCol = 0;
  if (projectCol === -1) projectCol = 1;

  // Determine default execution column: prefer visible rightmost column with content, or rightmost visible, or rightmost overall
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
  if (defaultExecutionColIndex === null && visibleExecCols.length > 0) {
    defaultExecutionColIndex = visibleExecCols[visibleExecCols.length - 1].colIndex;
  }
  if (defaultExecutionColIndex === null && executionColumns.length > 0) {
    for (let i = executionColumns.length - 1; i >= 0; i--) {
      if (hasDataInCol(executionColumns[i].colIndex)) {
        defaultExecutionColIndex = executionColumns[i].colIndex;
        break;
      }
    }
    if (defaultExecutionColIndex === null) {
      defaultExecutionColIndex =
        executionColumns[executionColumns.length - 1].colIndex;
    }
  }

  // Group data rows into teams
  const teams: TeamReport[] = [];
  let r = headerRowIndex + 1;

  while (r < rows.length) {
    const row = rows[r] ?? [];
    const teamText = row[teamCol]?.formattedValue?.trim() ?? '';
    const hasAnyRowContent = row.some((c) => Boolean(c?.formattedValue?.trim()));

    if (!teamText && !hasAnyRowContent) {
      r++;
      continue;
    }

    if (!teamText && teams.length === 0) {
      r++;
      continue;
    }

    // If this row starts a new team (or merges across multiple rows)
    const mergeEnd = getMergeEndRow(merges, r, teamCol);
    let endRow = mergeEnd;

    // Also consume subsequent rows if teamCol is empty while other columns have content (unmerged continuation rows)
    if (teamText) {
      let probe = Math.max(r, mergeEnd) + 1;
      while (probe < rows.length) {
        const probeRow = rows[probe] ?? [];
        const probeTeam = probeRow[teamCol]?.formattedValue?.trim() ?? '';
        const probeHasContent = probeRow.some((c) =>
          Boolean(c?.formattedValue?.trim())
        );
        if (probeTeam || !probeHasContent) break;
        endRow = probe;
        probe++;
      }
    }

    const rawTeamLabel =
      teamText || (teams.length > 0 ? teams[teams.length - 1].rawTeamLabel : '미지정 팀');
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

    // Split each column into numbered chunks so 1), 2), 3)... align into structured project rows
    const projectChunks = splitIntoNumberedChunks(projectLinesAll);
    const ownerChunks = splitIntoNumberedChunks(ownerLinesAll);
    const plannedChunks = splitIntoNumberedChunks(plannedLinesAll);
    const issueChunks = splitIntoNumberedChunks(issueLinesAll);
    const execChunksByCol: Record<number, NumberedChunk[]> = {};
    for (const ec of executionColumns) {
      execChunksByCol[ec.colIndex] = splitIntoNumberedChunks(
        execLinesByCol[ec.colIndex]
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
          // Non-numbered project text: create an entry
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

        // Strip leading "1)" from project title for crisp display while keeping numberPrefix
        const rawProjText = pChunk.lines.map((l) => l.text).join(' ').trim();
        const cleanTitle = rawProjText.replace(/^\d{1,2}\s*[).]\s*/, '').trim();

        const stripNumFromFirstLine = (chunks: NumberedChunk[]): RichTextLine[] => {
          const matched = chunks.filter((c) => c.num === num);
          const result: RichTextLine[] = [];
          for (const m of matched) {
            m.lines.forEach((line, idx) => {
              if (idx === 0) {
                const cleaned = line.text.replace(/^\d{1,2}\s*[).]\s*/, '').trim();
                if (cleaned) result.push({ ...line, text: cleaned });
              } else {
                result.push(line);
              }
            });
          }
          return result;
        };

        const execMap: Record<number, RichTextLine[]> = {};
        for (const ec of executionColumns) {
          execMap[ec.colIndex] = stripNumFromFirstLine(execChunksByCol[ec.colIndex]);
        }

        projects.push({
          key: `${r}-proj-${num}-${projects.length}`,
          numberPrefix: num,
          projectTitle: cleanTitle || rawProjText,
          owners: stripNumFromFirstLine(ownerChunks),
          plannedTasks: stripNumFromFirstLine(plannedChunks),
          issues: stripNumFromFirstLine(issueChunks),
          executionByCol: execMap,
        });
      }

      // Collect any unnumbered or unmatched chunks into extra sections so zero data is lost
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
      // Fallback: sheet row does not use 1) 2) numbering — preserve everything in a single project/summary entry
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
