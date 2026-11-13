import React, { useState } from 'react';
import { usePWAInstall } from '../hooks/usePWAInstall.ts';

interface PWAInstallButtonProps {
  variant?: 'header' | 'login';
}

export const PWAInstallButton: React.FC<PWAInstallButtonProps> = ({ variant = 'header' }) => {
  const { isInstallable, isInstalled, isIOS, isDesktop, install } = usePWAInstall();
  const [showIOSGuide, setShowIOSGuide] = useState(false);

  if (isInstalled) {
    return null;
  }

  if (isInstallable) {
    const label = isDesktop ? 'Install di PC' : 'Install App';
    return (
      <button
        type="button"
        onClick={install}
        title="Instal aplikasi ke perangkat agar cepat dan ringan"
        className={
          variant === 'login'
            ? 'inline-flex items-center gap-2 rounded-xl bg-blue-600/20 border border-blue-400/30 px-3.5 py-2 text-xs font-semibold text-blue-200 hover:bg-blue-600/30 hover:text-white transition-colors'
            : 'inline-flex items-center gap-1.5 rounded-xl bg-blue-600 px-3 py-1.5 text-xs sm:text-sm font-medium text-white shadow-sm hover:bg-blue-700 transition-colors'
        }
      >
        <svg className="w-4 h-4 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
        </svg>
        <span className={variant === 'header' ? 'hidden sm:inline' : ''}>{label}</span>
      </button>
    );
  }

  if (isIOS) {
    return (
      <>
        <button
          type="button"
          onClick={() => setShowIOSGuide(true)}
          className={
            variant === 'login'
              ? 'inline-flex items-center gap-2 rounded-xl border border-slate-700 px-3.5 py-2 text-xs font-medium text-slate-200 hover:bg-slate-800 transition-colors'
              : 'inline-flex items-center gap-1.5 rounded-xl border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 transition-colors'
          }
        >
          Install di iOS
        </button>

        {showIOSGuide && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 p-4">
            <div className="w-full max-w-sm rounded-2xl bg-white p-6 shadow-xl border border-slate-200 text-left">
              <h3 className="text-lg font-bold text-slate-800">Install di iPhone / iPad</h3>
              <p className="mt-2 text-sm text-slate-600 leading-relaxed">
                1. Ketuk tombol <strong>Bagikan (Share)</strong> di bilah Safari.<br />
                2. Gulir ke bawah lalu ketuk <strong>Tambah ke Layar Utama (Add to Home Screen)</strong>.
              </p>
              <button
                type="button"
                onClick={() => setShowIOSGuide(false)}
                className="mt-5 w-full rounded-xl bg-slate-100 py-2.5 text-sm font-semibold text-slate-800 hover:bg-slate-200 transition-colors"
              >
                Tutup
              </button>
            </div>
          </div>
        )}
      </>
    );
  }

  return null;
};
