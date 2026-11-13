import React, { useEffect, useMemo, useState } from 'react';
import {
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  ExternalLink,
  FolderOpen,
  RefreshCw,
  Search,
} from 'lucide-react';
import { User } from 'firebase/auth';
import {
  fetchSheetGrid,
  SheetGridPayload,
  SheetTabInfo,
  WorkspaceApiError,
} from '../lib/googleWorkspace';
import {
  ParsedWeeklyReport,
  parseWeeklyReportGrid,
  ProjectEntry,
  RichTextLine,
  TeamReport,
} from '../lib/reportParser';
import { RichTaskLines } from './RichTaskLines';

interface ReportViewerProps {
  accessToken: string;
  user: User;
  spreadsheetId: string;
  spreadsheetName: string;
  onChangeSpreadsheet: () => void;
  onLogout: () => void;
  onAuthExpired: () => void;
  onResolvedTitle: (id: string, title: string) => void;
}

export const ReportViewer: React.FC<ReportViewerProps> = ({
  accessToken,
  user,
  spreadsheetId,
  spreadsheetName,
  onChangeSpreadsheet,
  onLogout,
  onAuthExpired,
  onResolvedTitle,
}) => {
  const [tabs, setTabs] = useState<SheetTabInfo[]>([]);
  const [activeTabTitle, setActiveTabTitle] = useState<string>('');
  const [gridPayload, setGridPayload] = useState<SheetGridPayload | null>(null);
  const [parsedReport, setParsedReport] = useState<ParsedWeeklyReport | null>(
    null
  );

  const [selectedExecCol, setSelectedExecCol] = useState<number | null>(null);
  const [selectedTeamIndex, setSelectedTeamIndex] = useState<number>(0);

  // Requirement 4: Collapsible sections (collapsed by default per Requirement 2)
  const [showOwners, setShowOwners] = useState(false);
  const [showPlanned, setShowPlanned] = useState(false);
  const [showIssues, setShowIssues] = useState(false);
  const [hideStrikethrough, setHideStrikethrough] = useState(false);

  const [teamSearch, setTeamSearch] = useState('');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const loadSheet = async (targetTab?: string) => {
    setLoading(true);
    setError(null);
    try {
      const payload = await fetchSheetGrid(
        accessToken,
        spreadsheetId,
        targetTab ?? activeTabTitle
      );
      setGridPayload(payload);
      setTabs(payload.tabs);
      setActiveTabTitle(payload.activeTab.title);
      onResolvedTitle(spreadsheetId, payload.spreadsheetTitle);

      const parsed = parseWeeklyReportGrid(payload);
      setParsedReport(parsed);
      setSelectedExecCol(parsed.defaultExecutionColIndex);
      setSelectedTeamIndex(0);
      setTeamSearch('');
    } catch (err) {
      if (err instanceof WorkspaceApiError && err.status === 401) {
        onAuthExpired();
        return;
      }
      setError(
        err instanceof Error
          ? err.message
          : '스프레드시트 데이터를 불러오는 중 오류가 발생했습니다.'
      );
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    setActiveTabTitle('');
    loadSheet('');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [spreadsheetId, accessToken]);

  const teams = parsedReport?.teams ?? [];
  const currentTeam: TeamReport | null = teams[selectedTeamIndex] ?? null;

  // Keyboard Left/Right navigation between teams when not typing in inputs
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const tag = (e.target as HTMLElement | null)?.tagName;
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
      if (teams.length <= 1) return;

      if (e.key === 'ArrowLeft') {
        setSelectedTeamIndex((prev) => Math.max(0, prev - 1));
      } else if (e.key === 'ArrowRight') {
        setSelectedTeamIndex((prev) => Math.min(teams.length - 1, prev + 1));
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [teams.length]);

  const activeExecColMeta = useMemo(() => {
    if (!parsedReport || selectedExecCol === null) return null;
    return (
      parsedReport.executionColumns.find(
        (c) => c.colIndex === selectedExecCol
      ) ?? null
    );
  }, [parsedReport, selectedExecCol]);

  const activeExecHeaderLabel =
    activeExecColMeta?.headerLabel ?? '업무 실행';

  const currentTeamIssueCount = useMemo(() => {
    if (!currentTeam) return 0;
    const projectIssues = currentTeam.projects.filter(
      (p) => p.issues.length > 0
    ).length;
    return projectIssues + (currentTeam.extraIssues.length > 0 ? 1 : 0);
  }, [currentTeam]);

  const currentTeamHasStrikethrough = useMemo(() => {
    if (!currentTeam) return false;
    return currentTeam.projects.some(
      (p) =>
        p.plannedTasks.some((l) => l.strikethrough) ||
        (selectedExecCol !== null &&
          (p.executionByCol[selectedExecCol] ?? []).some((l) => l.strikethrough))
    );
  }, [currentTeam, selectedExecCol]);

  const filteredProjects: ProjectEntry[] = useMemo(() => {
    if (!currentTeam) return [];
    const q = teamSearch.trim().toLowerCase();
    if (!q) return currentTeam.projects;

    return currentTeam.projects.filter((p) => {
      const execLines =
        selectedExecCol !== null ? (p.executionByCol[selectedExecCol] ?? []) : [];
      const searchable = [
        p.projectTitle,
        ...p.owners.map((l) => l.text),
        ...p.plannedTasks.map((l) => l.text),
        ...p.issues.map((l) => l.text),
        ...execLines.map((l) => l.text),
      ]
        .join(' ')
        .toLowerCase();
      return searchable.includes(q);
    });
  }, [currentTeam, teamSearch, selectedExecCol]);

  const allExpanded = showOwners && showPlanned && showIssues;
  const handleToggleAllDetails = () => {
    const next = !allExpanded;
    setShowOwners(next);
    setShowPlanned(next);
    setShowIssues(next);
  };

  // Execution column stepping
  const execCols = parsedReport?.executionColumns ?? [];
  const currentExecPosition = execCols.findIndex(
    (c) => c.colIndex === selectedExecCol
  );
  const canStepPrevWeek = currentExecPosition > 0;
  const canStepNextWeek =
    currentExecPosition >= 0 && currentExecPosition < execCols.length - 1;

  // Sheet tab stepping
  const currentTabIndex = tabs.findIndex((t) => t.title === activeTabTitle);
  const canStepPrevTab = currentTabIndex > 0;
  const canStepNextTab =
    currentTabIndex >= 0 && currentTabIndex < tabs.length - 1;

  const extraExecLines: RichTextLine[] =
    currentTeam && selectedExecCol !== null
      ? (currentTeam.extraExecutionByCol[selectedExecCol] ?? [])
      : [];

  const hasVisibleExtras =
    Boolean(currentTeam) &&
    (extraExecLines.length > 0 ||
      (showOwners && (currentTeam?.extraOwners.length ?? 0) > 0) ||
      (showPlanned && (currentTeam?.extraPlanned.length ?? 0) > 0) ||
      (showIssues && (currentTeam?.extraIssues.length ?? 0) > 0));

  const sheetUrl = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(
    spreadsheetId
  )}`;

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      {/* Top Bar: 3-zone contract */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-30">
        <div className="max-w-[1440px] mx-auto px-4 sm:px-6 h-14 flex items-center justify-between gap-8">
          <span className="text-base font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0">
            주간업무보고 뷰어
          </span>

          <nav className="hidden md:flex items-center gap-6 text-sm font-medium text-slate-600">
            <button
              type="button"
              onClick={onChangeSpreadsheet}
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1.5"
            >
              <FolderOpen className="w-4 h-4" />
              <span>문서 변경</span>
            </button>
            <button
              type="button"
              onClick={() => loadSheet(activeTabTitle)}
              disabled={loading}
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1.5 disabled:opacity-50"
            >
              <RefreshCw
                className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`}
              />
              <span>새로고침</span>
            </button>
            <a
              href={sheetUrl}
              target="_blank"
              rel="noopener noreferrer"
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1.5"
            >
              <span>원본 시트</span>
              <ExternalLink className="w-3.5 h-3.5" />
            </a>
          </nav>

          <div className="flex items-center gap-4 shrink-0">
            <span className="hidden sm:inline text-xs text-slate-500 truncate max-w-[200px]">
              {user.displayName || user.email}
            </span>
            <button
              type="button"
              onClick={onLogout}
              className="px-3.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg transition-colors whitespace-nowrap shrink-0"
            >
              로그아웃
            </button>
          </div>
        </div>
      </header>

      {/* Workspace Control Toolbar */}
      <div className="bg-white border-b border-slate-200">
        <div className="max-w-[1440px] mx-auto px-4 sm:px-6 py-3 flex flex-col xl:flex-row xl:items-center xl:justify-between gap-3">
          {/* Left: Document & History Selectors (Sheet Tabs + Execution Week Columns) */}
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex md:hidden items-center gap-2">
              <button
                type="button"
                onClick={onChangeSpreadsheet}
                className="h-9 px-3 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg whitespace-nowrap"
              >
                문서 변경
              </button>
              <button
                type="button"
                onClick={() => loadSheet(activeTabTitle)}
                disabled={loading}
                className="h-9 px-3 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg whitespace-nowrap"
              >
                새로고침
              </button>
            </div>

            {tabs.length > 0 && (
              <div className="flex items-center gap-1.5">
                <span className="text-xs font-semibold text-slate-500 whitespace-nowrap mr-1">
                  시트 탭
                </span>
                <button
                  type="button"
                  onClick={() => {
                    if (canStepPrevTab) {
                      loadSheet(tabs[currentTabIndex - 1].title);
                    }
                  }}
                  disabled={!canStepPrevTab || loading}
                  aria-label="이전 시트 탭"
                  className="h-9 w-8 inline-flex items-center justify-center rounded-lg border border-slate-300 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-40 transition-colors"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <select
                  aria-label="시트 탭 선택"
                  value={activeTabTitle}
                  onChange={(e) => loadSheet(e.target.value)}
                  disabled={loading}
                  className="h-9 pl-3 pr-8 text-sm font-medium text-slate-800 bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 max-w-[220px]"
                >
                  {tabs.map((tab) => (
                    <option key={tab.sheetId} value={tab.title}>
                      {tab.title}
                      {tab.hidden ? ' (숨김)' : ''}
                    </option>
                  ))}
                </select>
                <button
                  type="button"
                  onClick={() => {
                    if (canStepNextTab) {
                      loadSheet(tabs[currentTabIndex + 1].title);
                    }
                  }}
                  disabled={!canStepNextTab || loading}
                  aria-label="다음 시트 탭"
                  className="h-9 w-8 inline-flex items-center justify-center rounded-lg border border-slate-300 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-40 transition-colors"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            )}

            {execCols.length > 0 && (
              <div className="flex items-center gap-1.5">
                <span className="text-xs font-semibold text-slate-500 whitespace-nowrap mr-1">
                  실행 주차
                </span>
                <button
                  type="button"
                  onClick={() => {
                    if (canStepPrevWeek) {
                      setSelectedExecCol(
                        execCols[currentExecPosition - 1].colIndex
                      );
                    }
                  }}
                  disabled={!canStepPrevWeek}
                  aria-label="이전 주 업무 실행"
                  className="h-9 w-8 inline-flex items-center justify-center rounded-lg border border-slate-300 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-40 transition-colors"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <select
                  aria-label="실행 주차 선택"
                  value={selectedExecCol ?? ''}
                  onChange={(e) => setSelectedExecCol(Number(e.target.value))}
                  className="h-9 pl-3 pr-8 text-sm font-medium text-blue-700 bg-blue-50/70 border border-blue-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 font-mono-tabular"
                >
                  {execCols.map((col) => (
                    <option key={col.colIndex} value={col.colIndex}>
                      {col.headerLabel}
                      {col.hidden ? ' (이전/숨김)' : ''}
                    </option>
                  ))}
                </select>
                <button
                  type="button"
                  onClick={() => {
                    if (canStepNextWeek) {
                      setSelectedExecCol(
                        execCols[currentExecPosition + 1].colIndex
                      );
                    }
                  }}
                  disabled={!canStepNextWeek}
                  aria-label="다음 주 업무 실행"
                  className="h-9 w-8 inline-flex items-center justify-center rounded-lg border border-slate-300 bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-40 transition-colors"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            )}
          </div>

          {/* Right: Collapsible Column Toggles (Requirement 4) */}
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-xs font-semibold text-slate-500 mr-1 whitespace-nowrap">
              추가 항목 보기
            </span>

            <button
              type="button"
              onClick={() => setShowOwners((v) => !v)}
              aria-pressed={showOwners}
              className={`h-9 px-3 text-xs font-semibold rounded-lg border inline-flex items-center gap-1.5 transition-colors whitespace-nowrap shrink-0 ${
                showOwners
                  ? 'bg-blue-700 text-white border-blue-700'
                  : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
              }`}
            >
              <span>담당자</span>
              {showOwners ? (
                <ChevronUp className="w-3.5 h-3.5" />
              ) : (
                <ChevronDown className="w-3.5 h-3.5" />
              )}
            </button>

            <button
              type="button"
              onClick={() => setShowPlanned((v) => !v)}
              aria-pressed={showPlanned}
              className={`h-9 px-3 text-xs font-semibold rounded-lg border inline-flex items-center gap-1.5 transition-colors whitespace-nowrap shrink-0 ${
                showPlanned
                  ? 'bg-blue-700 text-white border-blue-700'
                  : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
              }`}
            >
              <span>업무 진행 예정</span>
              {showPlanned ? (
                <ChevronUp className="w-3.5 h-3.5" />
              ) : (
                <ChevronDown className="w-3.5 h-3.5" />
              )}
            </button>

            <button
              type="button"
              onClick={() => setShowIssues((v) => !v)}
              aria-pressed={showIssues}
              className={`h-9 px-3 text-xs font-semibold rounded-lg border inline-flex items-center gap-1.5 transition-colors whitespace-nowrap shrink-0 ${
                showIssues
                  ? 'bg-red-700 text-white border-red-700'
                  : currentTeamIssueCount > 0
                    ? 'bg-white text-red-700 border-red-200 hover:bg-red-50'
                    : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
              }`}
            >
              <span>
                이슈사항
                {currentTeamIssueCount > 0 ? ` (${currentTeamIssueCount})` : ''}
              </span>
              {showIssues ? (
                <ChevronUp className="w-3.5 h-3.5" />
              ) : (
                <ChevronDown className="w-3.5 h-3.5" />
              )}
            </button>

            <button
              type="button"
              onClick={handleToggleAllDetails}
              className="h-9 px-3 text-xs font-medium text-slate-600 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors whitespace-nowrap shrink-0"
            >
              {allExpanded ? '모두 접기' : '모두 펼치기'}
            </button>
          </div>
        </div>
      </div>

      {/* Main Content Workspace */}
      <div className="flex-1 max-w-[1440px] w-full mx-auto px-4 sm:px-6 py-6 flex flex-col lg:flex-row gap-6 items-start">
        {/* Left Sidebar: Report Context & Team List */}
        <aside className="w-full lg:w-72 shrink-0 space-y-4">
          <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-2.5">
            <p className="text-xs font-medium text-slate-500 truncate">
              {spreadsheetName}
            </p>
            <h1 className="text-base font-bold text-slate-900 leading-snug">
              {parsedReport?.reportTitle || spreadsheetName}
            </h1>
            {(parsedReport?.organization || parsedReport?.reportSchedule) && (
              <div className="pt-2 border-t border-slate-100 text-xs text-slate-600 space-y-1">
                {parsedReport.organization && (
                  <p>조직명: {parsedReport.organization}</p>
                )}
                {parsedReport.reportSchedule && (
                  <p>보고 일시: {parsedReport.reportSchedule}</p>
                )}
              </div>
            )}
          </div>

          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
            <div className="px-4 py-3 border-b border-slate-200 flex items-center justify-between">
              <span className="text-xs font-semibold text-slate-700">
                팀 목록 (한 페이지에 한 팀)
              </span>
              <span className="text-xs font-mono-tabular text-slate-500">
                {teams.length}개 팀
              </span>
            </div>

            {loading ? (
              <div className="p-3 space-y-2 animate-pulse">
                {[1, 2, 3, 4].map((n) => (
                  <div key={n} className="h-11 bg-slate-100 rounded-lg" />
                ))}
              </div>
            ) : teams.length === 0 ? (
              <div className="p-4 text-xs text-slate-500">
                팀 데이터를 찾지 못했습니다.
              </div>
            ) : (
              <>
                {/* Mobile Dropdown */}
                <div className="p-3 lg:hidden">
                  <select
                    aria-label="팀 선택"
                    value={selectedTeamIndex}
                    onChange={(e) =>
                      setSelectedTeamIndex(Number(e.target.value))
                    }
                    className="w-full h-10 px-3 text-sm font-medium text-slate-900 bg-white border border-slate-300 rounded-lg"
                  >
                    {teams.map((t, idx) => (
                      <option key={t.id} value={idx}>
                        {idx + 1}. {t.teamName}
                        {t.leaderName ? ` (${t.leaderName})` : ''} · 프로젝트{' '}
                        {t.projects.length}건
                      </option>
                    ))}
                  </select>
                </div>

                {/* Desktop Vertical Team Navigation */}
                <nav
                  aria-label="팀 목록"
                  className="hidden lg:block p-2 space-y-1 max-h-[calc(100vh-340px)] overflow-y-auto"
                >
                  {teams.map((team, idx) => {
                    const isSelected = idx === selectedTeamIndex;
                    return (
                      <button
                        key={team.id}
                        type="button"
                        onClick={() => setSelectedTeamIndex(idx)}
                        className={`w-full text-left px-3 py-2.5 rounded-lg border transition-colors flex items-center justify-between gap-2 ${
                          isSelected
                            ? 'bg-blue-50 text-blue-800 border-blue-200'
                            : 'bg-transparent text-slate-700 border-transparent hover:bg-slate-50'
                        }`}
                      >
                        <div className="min-w-0">
                          <p
                            className={`text-sm truncate ${
                              isSelected ? 'font-bold' : 'font-medium'
                            }`}
                          >
                            {team.teamName}
                          </p>
                          {team.leaderName && (
                            <p className="text-xs text-slate-500 truncate mt-0.5">
                              리더 {team.leaderName}
                            </p>
                          )}
                        </div>
                        <span
                          className={`text-xs font-mono-tabular shrink-0 ${
                            isSelected ? 'text-blue-700 font-semibold' : 'text-slate-400'
                          }`}
                        >
                          {team.projects.length}
                        </span>
                      </button>
                    );
                  })}
                </nav>
              </>
            )}
          </div>
        </aside>

        {/* Right Main Viewport: Single Team Per Page */}
        <main className="flex-1 min-w-0 w-full space-y-5">
          {error ? (
            <div
              role="alert"
              className="bg-white border border-red-200 rounded-xl p-8 text-center space-y-3"
            >
              <p className="text-base font-semibold text-red-700">
                시트를 불러오지 못했습니다
              </p>
              <p className="text-sm text-slate-600">{error}</p>
              <div className="pt-2 flex items-center justify-center gap-3">
                <button
                  type="button"
                  onClick={() => loadSheet(activeTabTitle)}
                  className="px-4 py-2 text-sm font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors"
                >
                  다시 시도
                </button>
                <button
                  type="button"
                  onClick={onChangeSpreadsheet}
                  className="px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-lg transition-colors"
                >
                  다른 문서 선택
                </button>
              </div>
            </div>
          ) : loading ? (
            <div className="bg-white border border-slate-200 rounded-xl p-6 space-y-5 animate-pulse">
              <div className="flex items-center justify-between pb-4 border-b border-slate-100">
                <div className="space-y-2">
                  <div className="h-6 bg-slate-200 rounded w-56" />
                  <div className="h-4 bg-slate-100 rounded w-40" />
                </div>
                <div className="h-9 bg-slate-100 rounded w-36" />
              </div>
              {[1, 2, 3].map((row) => (
                <div
                  key={row}
                  className="grid grid-cols-12 gap-6 py-4 border-b border-slate-100"
                >
                  <div className="col-span-4 h-5 bg-slate-200 rounded w-3/4" />
                  <div className="col-span-8 space-y-2">
                    <div className="h-4 bg-slate-100 rounded w-full" />
                    <div className="h-4 bg-slate-100 rounded w-5/6" />
                    <div className="h-4 bg-slate-100 rounded w-2/3" />
                  </div>
                </div>
              ))}
            </div>
          ) : !currentTeam ? (
            <div className="bg-white border border-slate-200 rounded-xl p-10 text-center space-y-3">
              <h2 className="text-base font-bold text-slate-900">
                현재 탭({gridPayload?.activeTab.title || '-'})에서 주간업무보고 표를
                찾지 못했습니다
              </h2>
              <p className="text-sm text-slate-600 max-w-lg mx-auto leading-relaxed">
                &lsquo;팀(리더)&rsquo;, &lsquo;프로젝트&rsquo;, &lsquo;업무
                실행&rsquo; 열이 포함된 시트 탭을 선택하거나 다른 문서를 열어보세요.
              </p>
            </div>
          ) : (
            <>
              {/* Single Team Header + Pagination */}
              <section className="bg-white border border-slate-200 rounded-xl p-5 sm:p-6 space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <div className="space-y-1">
                    <div className="flex flex-wrap items-baseline gap-x-3 gap-y-1">
                      <h2 className="text-2xl font-bold tracking-tight text-slate-900">
                        {currentTeam.teamName}
                      </h2>
                      {currentTeam.leaderName && (
                        <span className="text-sm font-medium text-slate-600">
                          리더 {currentTeam.leaderName}
                        </span>
                      )}
                    </div>
                    <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500">
                      <span>프로젝트 {currentTeam.projects.length}건</span>
                      <span aria-hidden="true">·</span>
                      <span className="font-mono-tabular text-blue-700 font-medium">
                        {activeExecHeaderLabel}
                      </span>
                      {activeTabTitle && (
                        <>
                          <span aria-hidden="true">·</span>
                          <span>시트: {activeTabTitle}</span>
                        </>
                      )}
                    </div>
                  </div>

                  {/* Team Page Stepper */}
                  <div className="flex items-center gap-2 shrink-0">
                    <button
                      type="button"
                      onClick={() =>
                        setSelectedTeamIndex((i) => Math.max(0, i - 1))
                      }
                      disabled={selectedTeamIndex === 0}
                      className="h-9 px-3 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg disabled:opacity-40 transition-colors inline-flex items-center gap-1 whitespace-nowrap"
                    >
                      <ChevronLeft className="w-4 h-4" />
                      <span>이전 팀</span>
                    </button>
                    <span className="px-2.5 text-xs font-mono-tabular font-semibold text-slate-700 whitespace-nowrap">
                      {selectedTeamIndex + 1} / {teams.length}
                    </span>
                    <button
                      type="button"
                      onClick={() =>
                        setSelectedTeamIndex((i) =>
                          Math.min(teams.length - 1, i + 1)
                        )
                      }
                      disabled={selectedTeamIndex >= teams.length - 1}
                      className="h-9 px-3 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg disabled:opacity-40 transition-colors inline-flex items-center gap-1 whitespace-nowrap"
                    >
                      <span>다음 팀</span>
                      <ChevronRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                {/* In-team Search & Strikethrough Filter */}
                <div className="pt-3 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div className="relative flex-1 max-w-md">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                    <input
                      type="text"
                      value={teamSearch}
                      onChange={(e) => setTeamSearch(e.target.value)}
                      placeholder="현재 팀 내 프로젝트 또는 업무 내용 검색"
                      className="w-full h-9 pl-9 pr-3 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                    />
                  </div>

                  {currentTeamHasStrikethrough && (
                    <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-600 cursor-pointer select-none whitespace-nowrap">
                      <input
                        type="checkbox"
                        checked={hideStrikethrough}
                        onChange={(e) => setHideStrikethrough(e.target.checked)}
                        className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                      />
                      <span>완료(취소선) 항목 숨기기</span>
                    </label>
                  )}
                </div>
              </section>

              {/* Team Projects Table / Structured View */}
              <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
                {/* Column Header */}
                <div className="hidden md:grid grid-cols-12 gap-6 px-6 py-3 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
                  <div className="col-span-4">
                    프로젝트{showOwners ? ' · 담당자' : ''}
                  </div>
                  <div className="col-span-8">{activeExecHeaderLabel}</div>
                </div>

                {filteredProjects.length === 0 ? (
                  <div className="px-6 py-12 text-center text-sm text-slate-500">
                    검색어에 해당하는 프로젝트가 없습니다.
                  </div>
                ) : (
                  <div className="divide-y divide-slate-200">
                    {filteredProjects.map((proj) => {
                      const execLines =
                        selectedExecCol !== null
                          ? (proj.executionByCol[selectedExecCol] ?? [])
                          : [];

                      const showExpansionRow = showPlanned || showIssues;

                      return (
                        <article
                          key={proj.key}
                          className="px-5 sm:px-6 py-5 hover:bg-slate-50/40 transition-colors"
                        >
                          <div className="grid grid-cols-1 md:grid-cols-12 gap-4 md:gap-6 items-start">
                            {/* Project Title (+ Optional Owners) */}
                            <div className="md:col-span-4 space-y-2.5">
                              <div className="flex items-start gap-2.5">
                                <span className="font-mono-tabular text-xs font-semibold text-blue-700 mt-1 shrink-0 w-5">
                                  {proj.numberPrefix !== '-'
                                    ? `${proj.numberPrefix.padStart(2, '0')}.`
                                    : '·'}
                                </span>
                                <h3 className="text-[15px] font-bold text-slate-900 leading-snug break-words">
                                  {proj.projectTitle}
                                </h3>
                              </div>

                              {showOwners && (
                                <div className="pl-7 pt-1">
                                  <p className="text-[11px] font-semibold text-slate-400 mb-0.5">
                                    담당자
                                  </p>
                                  {proj.owners.length > 0 ? (
                                    <p className="text-sm text-slate-700 leading-relaxed">
                                      {proj.owners.map((o) => o.text).join(', ')}
                                    </p>
                                  ) : (
                                    <p className="text-xs text-slate-400">
                                      미기재
                                    </p>
                                  )}
                                </div>
                              )}
                            </div>

                            {/* Weekly Execution Content (Default Core Column) */}
                            <div className="md:col-span-8 space-y-4">
                              <div>
                                <p className="md:hidden text-xs font-semibold text-blue-700 mb-1.5">
                                  {activeExecHeaderLabel}
                                </p>
                                <RichTaskLines
                                  lines={execLines}
                                  hideStrikethrough={hideStrikethrough}
                                  emptyText="이번 주 업무 실행 내용이 기재되지 않았습니다."
                                />
                              </div>

                              {/* Optional Collapsible Sections: 업무 진행 예정 / 이슈사항 */}
                              {showExpansionRow && (
                                <div
                                  className={`pt-4 border-t border-slate-100 grid grid-cols-1 ${
                                    showPlanned && showIssues
                                      ? 'xl:grid-cols-2'
                                      : ''
                                  } gap-5`}
                                >
                                  {showPlanned && (
                                    <div className="space-y-1.5">
                                      <h4 className="text-xs font-semibold text-slate-500">
                                        업무 진행 예정
                                      </h4>
                                      <RichTaskLines
                                        lines={proj.plannedTasks}
                                        hideStrikethrough={hideStrikethrough}
                                        emptyText="기재된 진행 예정 사항이 없습니다."
                                      />
                                    </div>
                                  )}

                                  {showIssues && (
                                    <div className="space-y-1.5">
                                      <h4
                                        className={`text-xs font-semibold ${
                                          proj.issues.length > 0
                                            ? 'text-red-700'
                                            : 'text-slate-500'
                                        }`}
                                      >
                                        이슈사항
                                      </h4>
                                      <RichTaskLines
                                        lines={proj.issues}
                                        hideStrikethrough={hideStrikethrough}
                                        emphasisTone={
                                          proj.issues.length > 0
                                            ? 'issue'
                                            : 'default'
                                        }
                                        emptyText="특이 이슈사항 없음"
                                      />
                                    </div>
                                  )}
                                </div>
                              )}
                            </div>
                          </div>
                        </article>
                      );
                    })}

                    {/* Unnumbered / Common Team Notes if any exist */}
                    {hasVisibleExtras && (
                      <article className="px-5 sm:px-6 py-5 bg-slate-50/70">
                        <div className="grid grid-cols-1 md:grid-cols-12 gap-4 md:gap-6 items-start">
                          <div className="md:col-span-4 space-y-2">
                            <h3 className="text-sm font-bold text-slate-800">
                              기타 · 공통 사항
                            </h3>
                            {showOwners &&
                              currentTeam.extraOwners.length > 0 && (
                                <div>
                                  <p className="text-[11px] font-semibold text-slate-400">
                                    담당자
                                  </p>
                                  <RichTaskLines
                                    lines={currentTeam.extraOwners}
                                  />
                                </div>
                              )}
                          </div>
                          <div className="md:col-span-8 space-y-4">
                            {extraExecLines.length > 0 && (
                              <RichTaskLines
                                lines={extraExecLines}
                                hideStrikethrough={hideStrikethrough}
                              />
                            )}
                            {showPlanned &&
                              currentTeam.extraPlanned.length > 0 && (
                                <div className="pt-3 border-t border-slate-200 space-y-1.5">
                                  <h4 className="text-xs font-semibold text-slate-500">
                                    업무 진행 예정 (기타)
                                  </h4>
                                  <RichTaskLines
                                    lines={currentTeam.extraPlanned}
                                    hideStrikethrough={hideStrikethrough}
                                  />
                                </div>
                              )}
                            {showIssues &&
                              currentTeam.extraIssues.length > 0 && (
                                <div className="pt-3 border-t border-slate-200 space-y-1.5">
                                  <h4 className="text-xs font-semibold text-red-700">
                                    이슈사항 (공통)
                                  </h4>
                                  <RichTaskLines
                                    lines={currentTeam.extraIssues}
                                    emphasisTone="issue"
                                  />
                                </div>
                              )}
                          </div>
                        </div>
                      </article>
                    )}
                  </div>
                )}
              </section>

              {/* Bottom Team Pagination */}
              {teams.length > 1 && (
                <div className="flex items-center justify-between pt-1 pb-6">
                  <button
                    type="button"
                    onClick={() => {
                      setSelectedTeamIndex((i) => Math.max(0, i - 1));
                      window.scrollTo({ top: 0, behavior: 'smooth' });
                    }}
                    disabled={selectedTeamIndex === 0}
                    className="h-10 px-4 text-sm font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg disabled:opacity-40 transition-colors inline-flex items-center gap-1.5"
                  >
                    <ChevronLeft className="w-4 h-4" />
                    <span>
                      {selectedTeamIndex > 0
                        ? teams[selectedTeamIndex - 1].teamName
                        : '이전 팀'}
                    </span>
                  </button>

                  <span className="hidden sm:inline text-xs text-slate-400">
                    키보드 ← / → 로 팀 전환
                  </span>

                  <button
                    type="button"
                    onClick={() => {
                      setSelectedTeamIndex((i) =>
                        Math.min(teams.length - 1, i + 1)
                      );
                      window.scrollTo({ top: 0, behavior: 'smooth' });
                    }}
                    disabled={selectedTeamIndex >= teams.length - 1}
                    className="h-10 px-4 text-sm font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg disabled:opacity-40 transition-colors inline-flex items-center gap-1.5"
                  >
                    <span>
                      {selectedTeamIndex < teams.length - 1
                        ? teams[selectedTeamIndex + 1].teamName
                        : '다음 팀'}
                    </span>
                    <ChevronRight className="w-4 h-4" />
                  </button>
                </div>
              )}
            </>
          )}
        </main>
      </div>
    </div>
  );
};
