import React, { useState } from 'react';
import {
  Check,
  Copy,
  ExternalLink,
  FileSpreadsheet,
  Layers,
  SlidersHorizontal,
  Users,
} from 'lucide-react';
import { AuthErrorInfo, FIREBASE_PROJECT_ID } from '../lib/firebaseAuth';

interface LoginPageProps {
  onSignIn: () => void;
  isSigningIn: boolean;
  error: AuthErrorInfo | null;
}

export const LoginPage: React.FC<LoginPageProps> = ({
  onSignIn,
  isSigningIn,
  error,
}) => {
  const [copied, setCopied] = useState(false);
  const currentHost =
    typeof window !== 'undefined' ? window.location.hostname : '';
  const firebaseAuthSettingsUrl = `https://console.firebase.google.com/project/${FIREBASE_PROJECT_ID}/authentication/settings`;

  const handleCopyHost = async () => {
    if (!currentHost) return;
    try {
      await navigator.clipboard.writeText(currentHost);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1800);
    } catch {
      // ignore clipboard errors
    }
  };
  return (
    <div className="min-h-screen bg-slate-50 flex flex-col">
      <header className="border-b border-slate-200 bg-white">
        <div className="max-w-[1200px] mx-auto px-6 h-16 flex items-center justify-between gap-8">
          <span className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0">
            주간업무보고 뷰어
          </span>
          <span className="text-xs font-medium text-slate-500 whitespace-nowrap shrink-0">
            Google Drive · Google Sheets 연동
          </span>
        </div>
      </header>

      <main className="flex-1 flex items-center justify-center px-6 py-12">
        <div className="max-w-[1040px] w-full grid grid-cols-1 lg:grid-cols-12 gap-12 items-center">
          <div className="lg:col-span-7 space-y-8">
            <div className="space-y-4">
              <p className="text-xs font-semibold tracking-wide text-blue-700">
                기술연구소 주간업무 전용 뷰어
              </p>
              <h1
                className="text-3xl sm:text-4xl font-bold tracking-tight text-slate-900 leading-tight"
                style={{ textWrap: 'balance' }}
              >
                복잡한 주간업무보고 시트를
                <br />
                팀별 한 페이지씩 깔끔하게 확인하세요
              </h1>
              <p className="text-base text-slate-600 leading-relaxed max-w-xl">
                Google Drive에 저장된 주간업무보고 스프레드시트를 불러와, 이번 주{' '}
                <span className="font-semibold text-slate-800">업무 실행</span> 내용만
                집중해서 볼 수 있습니다. 필요한 경우에만 담당자·예정·이슈를 펼쳐서
                확인하세요.
              </p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 pt-2 border-t border-slate-200">
              <div className="flex items-start gap-3">
                <div className="mt-0.5 text-blue-700 shrink-0">
                  <Users className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-sm font-semibold text-slate-900">
                    한 페이지에 한 팀씩
                  </h2>
                  <p className="text-sm text-slate-600 mt-1 leading-relaxed">
                    팀별 업무가 길어도 한 번에 한 팀만 집중해 읽고 이전·다음 팀으로
                    바로 넘깁니다.
                  </p>
                </div>
              </div>

              <div className="flex items-start gap-3">
                <div className="mt-0.5 text-blue-700 shrink-0">
                  <SlidersHorizontal className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-sm font-semibold text-slate-900">
                    필요한 항목만 접기/펼치기
                  </h2>
                  <p className="text-sm text-slate-600 mt-1 leading-relaxed">
                    기본으로 팀·프로젝트·이번 주 업무 실행만 표시하며, 담당자·업무 진행
                    예정·이슈사항을 선택해 펼칩니다.
                  </p>
                </div>
              </div>

              <div className="flex items-start gap-3">
                <div className="mt-0.5 text-blue-700 shrink-0">
                  <Layers className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-sm font-semibold text-slate-900">
                    매주 누적되는 이전 기록 탐색
                  </h2>
                  <p className="text-sm text-slate-600 mt-1 leading-relaxed">
                    매주 추가되는 시트 탭과 업무 실행 주차 열을 모두 인식해 과거 보고
                    자료도 쉽게 조회합니다.
                  </p>
                </div>
              </div>

              <div className="flex items-start gap-3">
                <div className="mt-0.5 text-blue-700 shrink-0">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-sm font-semibold text-slate-900">
                    Google Drive 직접 연결
                  </h2>
                  <p className="text-sm text-slate-600 mt-1 leading-relaxed">
                    Google 계정으로 로그인해 내 드라이브의 주간업무보고 파일을 바로
                    선택해 열람합니다.
                  </p>
                </div>
              </div>
            </div>
          </div>

          <div className="lg:col-span-5">
            <div className="bg-white border border-slate-200 rounded-xl p-8 space-y-6">
              <div className="space-y-2">
                <h2 className="text-xl font-bold text-slate-900">Google 계정 로그인</h2>
                <p className="text-sm text-slate-600 leading-relaxed">
                  Google Drive 및 Google Sheets 읽기 권한을 승인하면 스프레드시트 목록을
                  불러올 수 있습니다. 원본 시트 데이터는 수정되지 않습니다.
                </p>
              </div>

              {error && (
                <div
                  role="alert"
                  className="p-4 rounded-lg bg-red-50 border border-red-200 text-sm text-red-800 space-y-3 leading-relaxed"
                >
                  <p className="font-medium">{error.message}</p>

                  {error.code === 'auth/unauthorized-domain' && currentHost && (
                    <div className="pt-2.5 border-t border-red-200/80 text-xs text-red-900 space-y-2.5">
                      <p className="font-semibold">
                        해결 방법 (1회만 설정): Firebase 콘솔에 아래 도메인을 추가해 주세요.
                      </p>
                      <div className="flex items-center justify-between gap-2 bg-white border border-red-200 rounded-md px-3 py-2">
                        <code className="font-mono-tabular text-slate-900 font-semibold truncate">
                          {currentHost}
                        </code>
                        <button
                          type="button"
                          onClick={handleCopyHost}
                          className="inline-flex items-center gap-1 text-xs font-semibold text-blue-700 hover:text-blue-800 whitespace-nowrap shrink-0"
                        >
                          {copied ? (
                            <>
                              <Check className="w-3.5 h-3.5" />
                              <span>복사됨</span>
                            </>
                          ) : (
                            <>
                              <Copy className="w-3.5 h-3.5" />
                              <span>도메인 복사</span>
                            </>
                          )}
                        </button>
                      </div>
                      <ol className="list-decimal list-inside space-y-1 text-red-800">
                        <li>아래 링크로 Firebase Authentication 설정 페이지 이동</li>
                        <li>승인된 도메인(Authorized domains) → 도메인 추가 클릭</li>
                        <li>
                          <code className="font-mono-tabular font-semibold">
                            {currentHost}
                          </code>{' '}
                          입력 후 추가 완료
                        </li>
                      </ol>
                      <a
                        href={firebaseAuthSettingsUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="inline-flex items-center gap-1.5 font-semibold text-blue-700 hover:underline"
                      >
                        <span>Firebase 승인된 도메인 설정 열기</span>
                        <ExternalLink className="w-3.5 h-3.5" />
                      </a>
                    </div>
                  )}
                </div>
              )}

              <div className="pt-1">
                <button
                  type="button"
                  onClick={onSignIn}
                  disabled={isSigningIn}
                  className="gsi-material-button w-full"
                >
                  <div className="gsi-material-button-content-wrapper">
                    <div className="gsi-material-button-icon">
                      <svg
                        version="1.1"
                        xmlns="http://www.w3.org/2000/svg"
                        viewBox="0 0 48 48"
                        style={{ display: 'block', width: '100%', height: '100%' }}
                      >
                        <path
                          fill="#EA4335"
                          d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"
                        />
                        <path
                          fill="#4285F4"
                          d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"
                        />
                        <path
                          fill="#FBBC05"
                          d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"
                        />
                        <path
                          fill="#34A853"
                          d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"
                        />
                        <path fill="none" d="M0 0h48v48H0z" />
                      </svg>
                    </div>
                    <span>
                      {isSigningIn ? 'Google 계정 연결 중...' : 'Sign in with Google'}
                    </span>
                  </div>
                </button>
              </div>

              <div className="pt-4 border-t border-slate-100 text-xs text-slate-500 space-y-1">
                <p>읽기 전용(Read-only) 권한만 요청하며 별도 서버에 문서를 저장하지 않습니다.</p>
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
};
