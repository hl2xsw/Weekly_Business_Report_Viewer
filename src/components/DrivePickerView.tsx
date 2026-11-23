import React, { useEffect, useMemo, useState } from 'react';
import {
  ArrowRight,
  FileSpreadsheet,
  Link as LinkIcon,
  RefreshCw,
  Search,
  X,
} from 'lucide-react';
import {
  DriveSpreadsheetFile,
  extractSpreadsheetId,
  listDriveSpreadsheets,
  WorkspaceApiError,
} from '../lib/googleWorkspace';

interface DrivePickerViewProps {
  accessToken: string;
  recentFiles: Array<{ id: string; name: string }>;
  onSelectFile: (file: { id: string; name: string }) => void;
  onAuthExpired: () => void;
  onClose?: () => void;
  isModal?: boolean;
}

function formatModifiedDate(iso: string): string {
  if (!iso) return '-';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const hh = String(d.getHours()).padStart(2, '0');
  const min = String(d.getMinutes()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
}

export const DrivePickerView: React.FC<DrivePickerViewProps> = ({
  accessToken,
  recentFiles,
  onSelectFile,
  onAuthExpired,
  onClose,
  isModal = false,
}) => {
  const [searchInput, setSearchInput] = useState('');
  const [activeQuery, setActiveQuery] = useState('');
  const [scope, setScope] = useState<'all' | 'sharedWithMe'>('all');
  const [files, setFiles] = useState<DriveSpreadsheetFile[]>([]);
  const [cachedAllFiles, setCachedAllFiles] = useState<DriveSpreadsheetFile[]>(
    []
  );
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [urlInput, setUrlInput] = useState('');
  const [urlError, setUrlError] = useState<string | null>(null);

  const loadFiles = async (
    query: string,
    targetScope: 'all' | 'sharedWithMe'
  ) => {
    setLoading(true);
    setError(null);
    try {
      const list = await listDriveSpreadsheets(accessToken, query, targetScope);
      setFiles(list);
      if (!query && targetScope === 'all') {
        setCachedAllFiles(list);
      }
    } catch (err) {
      if (err instanceof WorkspaceApiError && err.status === 401) {
        onAuthExpired();
        return;
      }
      setError(
        err instanceof Error
          ? err.message
          : 'Google Drive 파일 목록을 불러오지 못했습니다.'
      );
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadFiles(activeQuery, scope);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [accessToken, activeQuery, scope]);

  // Combine server search results with client-side substring matches on cached files
  // so filenames with underscores (e.g., 경영위원회_기술연구소.xlsx) match reliably
  const displayedFiles = useMemo(() => {
    const localQuery = searchInput.trim().toLowerCase();
    if (!localQuery) return files;

    const merged = new Map<string, DriveSpreadsheetFile>();
    for (const f of files) {
      merged.set(f.id, f);
    }
    for (const f of cachedAllFiles) {
      if (scope === 'sharedWithMe' && !f.shared) continue;
      if (
        f.name.toLowerCase().includes(localQuery) ||
        (f.ownerName && f.ownerName.toLowerCase().includes(localQuery))
      ) {
        if (!merged.has(f.id)) {
          merged.set(f.id, f);
        }
      }
    }

    const list = Array.from(merged.values());
    // When user is typing before clicking 검색, filter current list live
    if (localQuery !== activeQuery.trim().toLowerCase()) {
      return list.filter(
        (f) =>
          f.name.toLowerCase().includes(localQuery) ||
          (f.ownerName && f.ownerName.toLowerCase().includes(localQuery))
      );
    }
    return list;
  }, [files, cachedAllFiles, searchInput, activeQuery, scope]);

  const handleSearchSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setActiveQuery(searchInput.trim());
  };

  const handleQuickFilter = (term: string) => {
    setSearchInput(term);
    setActiveQuery(term);
  };

  const handleOpenByUrl = (e: React.FormEvent) => {
    e.preventDefault();
    setUrlError(null);
    const id = extractSpreadsheetId(urlInput);
    if (!id) {
      setUrlError(
        '올바른 Google Sheets 또는 Google Drive 파일 링크/ID를 입력해 주세요.'
      );
      return;
    }
    onSelectFile({ id, name: '선택한 스프레드시트' });
  };

  const content = (
    <div className="bg-white border border-slate-200 rounded-xl overflow-hidden flex flex-col max-h-[85vh]">
      <div className="px-6 py-5 border-b border-slate-200 flex items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-bold text-slate-900">
            Google Drive에서 주간업무보고 선택
          </h2>
          <p className="text-sm text-slate-500 mt-0.5">
            내 드라이브·공유 문서함의 Google Sheets 및 Excel(.xlsx, .xls) 파일을
            모두 불러올 수 있습니다.
          </p>
        </div>
        {onClose && (
          <button
            type="button"
            onClick={onClose}
            aria-label="닫기"
            className="p-2 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-100 transition-colors shrink-0"
          >
            <X className="w-5 h-5" />
          </button>
        )}
      </div>

      <div className="p-6 space-y-6 overflow-y-auto">
        {recentFiles.length > 0 && (
          <div className="space-y-2.5">
            <h3 className="text-xs font-semibold text-slate-500">
              최근 열어본 문서
            </h3>
            <div className="flex flex-wrap gap-2">
              {recentFiles.map((item) => (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => onSelectFile(item)}
                  className="inline-flex items-center gap-2 px-3 py-2 text-sm font-medium text-slate-700 bg-slate-50 hover:bg-blue-50 hover:text-blue-700 border border-slate-200 hover:border-blue-200 rounded-lg transition-colors whitespace-nowrap"
                >
                  <FileSpreadsheet className="w-4 h-4 text-emerald-700 shrink-0" />
                  <span className="truncate max-w-[260px]">{item.name}</span>
                </button>
              ))}
            </div>
          </div>
        )}

        <div className="space-y-3">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <form
              onSubmit={handleSearchSubmit}
              className="flex-1 flex items-center gap-2"
            >
              <div className="relative flex-1">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                <input
                  type="text"
                  value={searchInput}
                  onChange={(e) => setSearchInput(e.target.value)}
                  placeholder="파일명으로 검색 (예: 경영위원회, 기술연구소, 주간)"
                  className="w-full h-10 pl-9 pr-3 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
                />
              </div>
              <button
                type="submit"
                className="h-10 px-4 text-sm font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-lg transition-colors whitespace-nowrap shrink-0"
              >
                검색
              </button>
              <button
                type="button"
                onClick={() => loadFiles(activeQuery, scope)}
                disabled={loading}
                aria-label="새로고침"
                className="h-10 px-3 text-sm font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg transition-colors shrink-0 inline-flex items-center gap-1.5"
              >
                <RefreshCw className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
                <span className="hidden sm:inline">새로고침</span>
              </button>
            </form>
          </div>

          {/* Scope & Quick Filters */}
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="inline-flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
              <button
                type="button"
                onClick={() => setScope('all')}
                className={`px-3 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
                  scope === 'all'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                전체 문서
              </button>
              <button
                type="button"
                onClick={() => setScope('sharedWithMe')}
                className={`px-3 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap ${
                  scope === 'sharedWithMe'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                공유 문서함
              </button>
            </div>

            <div className="flex items-center gap-1.5 shrink-0">
              <button
                type="button"
                onClick={() => handleQuickFilter('주간')}
                className={`px-3 py-1.5 text-xs font-medium rounded-lg border transition-colors whitespace-nowrap ${
                  activeQuery === '주간'
                    ? 'bg-blue-50 text-blue-700 border-blue-200'
                    : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'
                }`}
              >
                &quot;주간&quot; 필터
              </button>
              <button
                type="button"
                onClick={() => handleQuickFilter('기술연구소')}
                className={`px-3 py-1.5 text-xs font-medium rounded-lg border transition-colors whitespace-nowrap ${
                  activeQuery === '기술연구소'
                    ? 'bg-blue-50 text-blue-700 border-blue-200'
                    : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'
                }`}
              >
                &quot;기술연구소&quot; 필터
              </button>
              {(activeQuery || searchInput) && (
                <button
                  type="button"
                  onClick={() => handleQuickFilter('')}
                  className="px-3 py-1.5 text-xs font-medium rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 transition-colors whitespace-nowrap"
                >
                  초기화
                </button>
              )}
            </div>
          </div>

          {error && (
            <div
              role="alert"
              className="p-3.5 rounded-lg bg-red-50 border border-red-200 text-sm text-red-700"
            >
              {error}
            </div>
          )}

          <div className="border border-slate-200 rounded-lg overflow-hidden">
            <div className="grid grid-cols-12 gap-3 px-4 py-2.5 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
              <div className="col-span-6 sm:col-span-6">스프레드시트 이름</div>
              <div className="col-span-3 hidden sm:block">소유자</div>
              <div className="col-span-4 sm:col-span-2 text-right">수정 일시</div>
              <div className="col-span-2 sm:col-span-1 text-right">열기</div>
            </div>

            {loading ? (
              <div className="divide-y divide-slate-100">
                {[1, 2, 3, 4, 5].map((n) => (
                  <div
                    key={n}
                    className="grid grid-cols-12 gap-3 px-4 py-3.5 items-center animate-pulse"
                  >
                    <div className="col-span-6 h-4 bg-slate-200 rounded w-3/4" />
                    <div className="col-span-3 hidden sm:block h-4 bg-slate-100 rounded w-1/2" />
                    <div className="col-span-4 sm:col-span-2 h-4 bg-slate-100 rounded ml-auto w-24" />
                    <div className="col-span-2 sm:col-span-1 h-4 bg-slate-100 rounded ml-auto w-8" />
                  </div>
                ))}
              </div>
            ) : displayedFiles.length === 0 ? (
              <div className="px-6 py-10 text-center space-y-2">
                <p className="text-sm font-medium text-slate-700">
                  {searchInput.trim()
                    ? `"${searchInput.trim()}"에 해당하는 스프레드시트가 없습니다.`
                    : '조회된 Google Sheets 또는 Excel 문서가 없습니다.'}
                </p>
                <p className="text-xs text-slate-500">
                  검색 버튼을 눌러 드라이브 전체를 조회하거나, 아래에 공유 링크를 직접
                  붙여넣어 열어보세요.
                </p>
              </div>
            ) : (
              <div className="divide-y divide-slate-100 max-h-[340px] overflow-y-auto">
                {displayedFiles.map((file) => (
                  <button
                    key={file.id}
                    type="button"
                    onClick={() => onSelectFile({ id: file.id, name: file.name })}
                    className="w-full text-left grid grid-cols-12 gap-3 px-4 py-3 items-center hover:bg-blue-50/60 transition-colors group"
                  >
                    <div className="col-span-6 sm:col-span-6 flex items-center gap-2.5 min-w-0">
                      <FileSpreadsheet
                        className={`w-4 h-4 shrink-0 ${
                          file.fileKind === 'excel'
                            ? 'text-teal-700'
                            : 'text-emerald-700'
                        }`}
                      />
                      <div className="min-w-0">
                        <p className="text-sm font-medium text-slate-900 truncate group-hover:text-blue-700">
                          {file.name}
                        </p>
                        <p className="text-[11px] text-slate-400">
                          {file.fileKind === 'excel' ? 'Excel' : 'Google Sheets'}
                          {file.shared ? ' · 공유됨' : ''}
                        </p>
                      </div>
                    </div>
                    <div className="col-span-3 hidden sm:block text-xs text-slate-500 truncate">
                      {file.ownerName || '-'}
                    </div>
                    <div className="col-span-4 sm:col-span-2 text-right text-xs text-slate-500 font-mono-tabular whitespace-nowrap">
                      {formatModifiedDate(file.modifiedTime)}
                    </div>
                    <div className="col-span-2 sm:col-span-1 flex justify-end">
                      <ArrowRight className="w-4 h-4 text-slate-400 group-hover:text-blue-700 transition-colors" />
                    </div>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>

        <form
          onSubmit={handleOpenByUrl}
          className="pt-4 border-t border-slate-200 space-y-2.5"
        >
          <label
            htmlFor="sheet-url-input"
            className="block text-xs font-semibold text-slate-700"
          >
            또는 Google Sheets / Drive 링크·ID 직접 입력
          </label>
          <div className="flex items-center gap-2">
            <div className="relative flex-1">
              <LinkIcon className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
              <input
                id="sheet-url-input"
                type="text"
                value={urlInput}
                onChange={(e) => {
                  setUrlInput(e.target.value);
                  setUrlError(null);
                }}
                placeholder="https://docs.google.com/spreadsheets/d/... 또는 https://drive.google.com/file/d/..."
                className="w-full h-10 pl-9 pr-3 text-sm bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-600 focus:border-blue-600"
              />
            </div>
            <button
              type="submit"
              className="h-10 px-4 text-sm font-medium text-slate-800 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded-lg transition-colors whitespace-nowrap shrink-0"
            >
              직접 열기
            </button>
          </div>
          {urlError && <p className="text-xs text-red-600">{urlError}</p>}
        </form>
      </div>
    </div>
  );

  if (isModal) {
    return (
      <div
        className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-[1px] flex items-center justify-center p-4"
        role="dialog"
        aria-modal="true"
      >
        <div className="max-w-3xl w-full">{content}</div>
      </div>
    );
  }

  return <div className="max-w-4xl w-full mx-auto">{content}</div>;
};
