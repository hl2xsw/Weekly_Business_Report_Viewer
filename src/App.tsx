/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useEffect, useState } from 'react';
import { User } from 'firebase/auth';
import {
  clearAccessToken,
  googleSignIn,
  initAuth,
  logout,
} from './lib/firebaseAuth';
import { LoginPage } from './components/LoginPage';
import { DrivePickerView } from './components/DrivePickerView';
import { ReportViewer } from './components/ReportViewer';
import { clearSpreadsheetCache } from './lib/googleWorkspace';

interface SelectedFile {
  id: string;
  name: string;
}

const RECENT_FILES_KEY = 'weekly_report_recent_files_v1';
const LAST_FILE_KEY = 'weekly_report_last_file_v1';

function readStoredRecentFiles(): SelectedFile[] {
  try {
    const raw = localStorage.getItem(RECENT_FILES_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function readStoredLastFile(): SelectedFile | null {
  try {
    const raw = localStorage.getItem(LAST_FILE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed?.id ? parsed : null;
  } catch {
    return null;
  }
}

export default function App() {
  const [needsAuth, setNeedsAuth] = useState(true);
  const [authReady, setAuthReady] = useState(false);
  const [token, setToken] = useState<string | null>(null);
  const [user, setUser] = useState<User | null>(null);
  const [isSigningIn, setIsSigningIn] = useState(false);
  const [authError, setAuthError] = useState<string | null>(null);

  const [selectedFile, setSelectedFile] = useState<SelectedFile | null>(() =>
    readStoredLastFile()
  );
  const [recentFiles, setRecentFiles] = useState<SelectedFile[]>(() =>
    readStoredRecentFiles()
  );
  const [showPickerModal, setShowPickerModal] = useState(false);

  useEffect(() => {
    const unsubscribe = initAuth(
      (authedUser, accessToken) => {
        setUser(authedUser);
        setToken(accessToken);
        setNeedsAuth(false);
        setAuthReady(true);
      },
      () => {
        setToken(null);
        setNeedsAuth(true);
        setAuthReady(true);
      }
    );
    return () => unsubscribe();
  }, []);

  const handleSignIn = async () => {
    setIsSigningIn(true);
    setAuthError(null);
    try {
      const result = await googleSignIn();
      setUser(result.user);
      setToken(result.accessToken);
      setNeedsAuth(false);
    } catch (err) {
      const msg =
        err instanceof Error
          ? err.message
          : 'Google 로그인 중 오류가 발생했습니다.';
      setAuthError(msg);
      setNeedsAuth(true);
    } finally {
      setIsSigningIn(false);
    }
  };

  const handleLogout = async () => {
    await logout();
    clearSpreadsheetCache();
    setToken(null);
    setUser(null);
    setNeedsAuth(true);
    setShowPickerModal(false);
  };

  const handleAuthExpired = () => {
    clearAccessToken();
    setToken(null);
    setNeedsAuth(true);
    setAuthError('Google 인증이 만료되었습니다. 다시 로그인해 주세요.');
  };

  const saveRecentFile = (file: SelectedFile) => {
    setRecentFiles((prev) => {
      const next = [file, ...prev.filter((item) => item.id !== file.id)].slice(
        0,
        5
      );
      try {
        localStorage.setItem(RECENT_FILES_KEY, JSON.stringify(next));
      } catch {
        // ignore storage errors
      }
      return next;
    });
    try {
      localStorage.setItem(LAST_FILE_KEY, JSON.stringify(file));
    } catch {
      // ignore storage errors
    }
  };

  const handleSelectFile = (file: SelectedFile) => {
    setSelectedFile(file);
    saveRecentFile(file);
    setShowPickerModal(false);
  };

  const handleResolvedTitle = (id: string, title: string) => {
    if (!title) return;
    setSelectedFile((prev) => (prev?.id === id ? { id, name: title } : prev));
    saveRecentFile({ id, name: title });
  };

  if (!authReady) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center">
        <div className="text-sm text-slate-500">연결 상태 확인 중...</div>
      </div>
    );
  }

  if (needsAuth || !token || !user) {
    return (
      <LoginPage
        onSignIn={handleSignIn}
        isSigningIn={isSigningIn}
        error={authError}
      />
    );
  }

  if (!selectedFile) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col">
        <header className="bg-white border-b border-slate-200">
          <div className="max-w-[1200px] mx-auto px-6 h-14 flex items-center justify-between gap-8">
            <span className="text-base font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0">
              주간업무보고 뷰어
            </span>
            <div className="flex items-center gap-4 shrink-0">
              <span className="hidden sm:inline text-xs text-slate-500">
                {user.displayName || user.email}
              </span>
              <button
                type="button"
                onClick={handleLogout}
                className="px-3.5 py-1.5 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-lg transition-colors whitespace-nowrap shrink-0"
              >
                로그아웃
              </button>
            </div>
          </div>
        </header>

        <main className="flex-1 px-4 sm:px-6 py-10 flex items-start justify-center">
          <DrivePickerView
            accessToken={token}
            recentFiles={recentFiles}
            onSelectFile={handleSelectFile}
            onAuthExpired={handleAuthExpired}
          />
        </main>
      </div>
    );
  }

  return (
    <>
      <ReportViewer
        accessToken={token}
        user={user}
        spreadsheetId={selectedFile.id}
        spreadsheetName={selectedFile.name}
        onChangeSpreadsheet={() => setShowPickerModal(true)}
        onLogout={handleLogout}
        onAuthExpired={handleAuthExpired}
        onResolvedTitle={handleResolvedTitle}
      />

      {showPickerModal && (
        <DrivePickerView
          isModal
          accessToken={token}
          recentFiles={recentFiles}
          onSelectFile={handleSelectFile}
          onAuthExpired={handleAuthExpired}
          onClose={() => setShowPickerModal(false)}
        />
      )}
    </>
  );
}
