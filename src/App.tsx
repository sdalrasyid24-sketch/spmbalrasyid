import React, { useState, useEffect, useRef } from 'react';
import { 
  LayoutDashboard, Users, UserCheck, FileText, Settings, Menu, X, 
  Plus, Search, CheckCircle, XCircle, AlertCircle, Save, UserPlus, 
  BookOpen, ChevronRight, FileArchive, ChevronLeft, Download, Check, 
  ChevronDown, Layers, FileCheck, Clock, ArrowRight, Activity, Trash2, Edit, CalendarDays,
  Lock, User, Eye, EyeOff, LogOut, ShieldCheck, Sparkles
} from 'lucide-react';
import { firestoreService, initialDefaultSettings } from './services/firestoreService.ts';
import { PWAInstallButton } from './components/PWAInstallButton.tsx';
import { OfflineIndicator } from './components/OfflineIndicator.tsx';

// --- UI COMPONENTS ---
const Card = ({ children, className = "" }: any) => (
  <div className={`bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden ${className}`}>
    {children}
  </div>
);

const Button = ({ children, variant = "primary", size = "md", className = "", ...props }: any) => {
  const baseStyle = "inline-flex items-center justify-center font-medium transition-all duration-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-offset-2 disabled:opacity-50 disabled:cursor-not-allowed";
  const sizes: any = { sm: "px-3 py-1.5 text-sm", md: "px-4 py-2.5 text-sm", lg: "px-6 py-3 text-base" };
  const variants: any = {
    primary: "bg-blue-600 text-white hover:bg-blue-700 shadow-sm focus:ring-blue-500",
    secondary: "bg-slate-100 text-slate-700 hover:bg-slate-200 focus:ring-slate-500",
    outline: "border border-slate-200 text-slate-700 hover:bg-slate-50 focus:ring-slate-500",
    ghost: "text-slate-600 hover:bg-slate-100 focus:ring-slate-500",
    danger: "bg-rose-500 text-white hover:bg-rose-600 shadow-sm focus:ring-rose-500"
  };
  return (
    <button className={`${baseStyle} ${sizes[size]} ${variants[variant]} ${className}`} {...props}>
      {children}
    </button>
  );
};

const Input = ({ label, type = "text", required, className = "", ...props }: any) => (
  <div className={`flex flex-col space-y-1.5 ${className}`}>
    {label && <label className="text-sm font-semibold text-slate-700">{label} {required && <span className="text-rose-500">*</span>}</label>}
    <input
      type={type}
      required={required}
      className="px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 focus:bg-white transition-all duration-200 text-sm text-slate-800 placeholder-slate-400"
      {...props}
    />
  </div>
);

const Select = ({ label, options, value, onChange, required, className = "", placeholder = "Pilih salah satu..." }: any) => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: any) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const selectedOption = options.find((opt: any) => opt.value === value);

  return (
    <div className={`flex flex-col space-y-1.5 relative ${className}`} ref={dropdownRef}>
      {label && <label className="text-sm font-semibold text-slate-700">{label} {required && <span className="text-rose-500">*</span>}</label>}
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 focus:bg-white transition-all duration-200 text-sm text-left flex items-center justify-between"
      >
        <span className={selectedOption ? "text-slate-800 font-medium" : "text-slate-400"}>
          {selectedOption ? selectedOption.label : placeholder}
        </span>
        <ChevronDown size={16} className={`text-slate-400 transition-transform duration-200 ${isOpen ? 'rotate-180' : ''}`} />
      </button>

      {isOpen && (
        <div className="absolute top-[calc(100%+4px)] left-0 right-0 z-50 bg-white border border-slate-200 rounded-xl shadow-lg max-h-60 overflow-y-auto animate-in fade-in slide-in-from-top-2 duration-200 py-1">
          {options.map((opt: any, i: number) => (
            <button
              key={i}
              type="button"
              onClick={() => {
                onChange({ target: { value: opt.value } });
                setIsOpen(false);
              }}
              className={`w-full text-left px-4 py-2.5 text-sm transition-colors duration-150 ${
                value === opt.value 
                  ? 'bg-blue-50 text-blue-700 font-semibold' 
                  : 'text-slate-700 hover:bg-slate-50 hover:text-blue-600'
              }`}
            >
              {opt.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
};

const Textarea = ({ label, required, className = "", rows = 3, ...props }: any) => (
  <div className={`flex flex-col space-y-1.5 ${className}`}>
    {label && <label className="text-sm font-semibold text-slate-700">{label} {required && <span className="text-rose-500">*</span>}</label>}
    <textarea
      required={required}
      rows={rows}
      className="px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 focus:bg-white transition-all duration-200 text-sm text-slate-800 placeholder-slate-400 resize-y"
      {...props}
    />
  </div>
);

// --- EXPORT UTILITY ---
const exportToExcel = (title: string, columns: string[], data: any[][], filename: string) => {
  const tableHtml = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="utf-8" />
      <style>
        body { font-family: Arial, sans-serif; font-size: 12pt; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #000000; padding: 6px; text-align: left; vertical-align: middle; }
        .header-title { font-size: 16pt; font-weight: bold; text-align: center; border: none; background-color: #ffffff; height: 40px; }
        .spacer { border: none; background-color: #ffffff; height: 20px; }
        th { background-color: #d1d5db; font-weight: bold; }
      </style>
    </head>
    <body>
      <table>
        <tr><td colspan="${columns.length}" class="header-title">${title}</td></tr>
        <tr><td colspan="${columns.length}" class="spacer"></td></tr>
        <tr>${columns.map(col => `<th>${col}</th>`).join('')}</tr>
        ${data.map(row => `<tr>${row.map(cell => `<td>${cell !== null && cell !== undefined ? cell : ''}</td>`).join('')}</tr>`).join('')}
      </table>
    </body>
    </html>
  `;
  const blob = new Blob([tableHtml], { type: 'application/vnd.ms-excel' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${filename}.xls`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

// --- LOGIN COMPONENT ---
function LoginView({ onLoginSuccess, sessionMessage }: { onLoginSuccess: () => void; sessionMessage?: string | null }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (username.trim() === 'admin' && password === 'spmbalrasyid') {
      setError('');
      onLoginSuccess();
    } else {
      setError('Username atau password salah. Silakan coba lagi.');
    }
  };

  return (
    <div className="min-h-screen w-full bg-slate-100 flex items-center justify-center p-4 sm:p-6 lg:p-10 font-sans text-slate-800">
      {/* Main Container mengikuti posisi gambar referensi dengan warna sistem utama */}
      <div className="w-full max-w-4xl bg-slate-900 rounded-3xl shadow-2xl overflow-hidden relative p-6 sm:p-10 lg:p-12 grid grid-cols-1 md:grid-cols-12 gap-8 items-center animate-in fade-in zoom-in-95 duration-300">
        {/* Ornamen Latar Halus (seperti lengkungan pada gambar referensi namun dengan warna sistem) */}
        <div className="absolute -bottom-24 -left-24 w-72 h-72 rounded-full bg-slate-800/80 pointer-events-none" />
        <div className="absolute -top-32 left-1/3 w-96 h-96 rounded-full bg-blue-600/10 pointer-events-none" />

        {/* Sisi Kiri: Informasi & Sambutan SPMB SD Al-Rasyid */}
        <div className="md:col-span-7 relative z-10 text-white space-y-4 pr-0 md:pr-4 text-center md:text-left">
          <img
            src="https://iili.io/Fr7hdzb.png"
            alt="Logo Al-Rasyid"
            className="w-24 h-24 object-contain mx-auto md:mx-0 mb-2"
          />
          <div className="space-y-1.5">
            <h1 className="text-2xl sm:text-3xl lg:text-4xl font-extrabold tracking-tight text-white">
              SPMB SD AL-RASYID
            </h1>
            <p className="text-xs sm:text-sm font-semibold tracking-widest uppercase text-blue-400">
              Sistem Penerimaan Murid Baru
            </p>
          </div>
          <p className="text-sm text-slate-300 leading-relaxed max-w-md mx-auto md:mx-0">
            Selamat datang di portal administrasi Penerimaan Murid Baru SD Al-Rasyid. Kelola data pendaftar, verifikasi kelengkapan berkas fisik, seleksi penerimaan siswa, hingga pembagian rombongan belajar secara terpadu dan efisien.
          </p>
          <div className="pt-1 flex justify-center md:justify-start">
            <PWAInstallButton variant="login" />
          </div>
        </div>

        {/* Sisi Kanan: Kartu Putih Form Login */}
        <div className="md:col-span-5 relative z-10">
          <Card className="p-6 sm:p-8 border-slate-200 shadow-xl rounded-2xl bg-white">
            {sessionMessage && !error && (
              <div className="mb-5 p-3.5 rounded-xl bg-amber-50 border border-amber-200 flex items-start space-x-2.5 text-amber-800 text-sm animate-in fade-in duration-200">
                <Clock size={18} className="flex-shrink-0 mt-0.5 text-amber-600" />
                <span className="font-medium leading-snug">{sessionMessage}</span>
              </div>
            )}
            {error && (
              <div className="mb-5 p-3.5 rounded-xl bg-rose-50 border border-rose-200 flex items-center space-x-2.5 text-rose-600 text-sm animate-in fade-in duration-200">
                <AlertCircle size={18} className="flex-shrink-0" />
                <span className="font-medium">{error}</span>
              </div>
            )}

            <form onSubmit={handleLogin} className="space-y-4">
              <Input
                label="Username"
                type="text"
                value={username}
                onChange={(e: any) => {
                  setUsername(e.target.value);
                  if (error) setError('');
                }}
                placeholder="Masukkan username"
                required
                autoComplete="username"
              />

              <div className="flex flex-col space-y-1.5">
                <label className="text-sm font-semibold text-slate-700">
                  Password <span className="text-rose-500">*</span>
                </label>
                <div className="relative">
                  <input
                    type={showPassword ? 'text' : 'password'}
                    value={password}
                    onChange={(e) => {
                      setPassword(e.target.value);
                      if (error) setError('');
                    }}
                    placeholder="Masukkan password"
                    required
                    autoComplete="current-password"
                    className="w-full px-4 py-2.5 pr-11 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 focus:bg-white transition-all duration-200 text-sm text-slate-800 placeholder-slate-400"
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword(!showPassword)}
                    className="absolute inset-y-0 right-2.5 flex items-center justify-center px-1.5 text-slate-400 hover:text-slate-600 rounded-lg transition-colors"
                    title={showPassword ? 'Sembunyikan Password' : 'Lihat Password'}
                  >
                    {showPassword ? <EyeOff size={18} /> : <Eye size={18} />}
                  </button>
                </div>
              </div>

              <div className="pt-2">
                <Button type="submit" className="w-full py-2.5">
                  Masuk ke Sistem
                </Button>
              </div>
            </form>
          </Card>
        </div>
      </div>
    </div>
  );
}

// --- SESSION & INACTIVITY CONSTANTS ---
const AUTH_KEY = 'spmb_alrasyid_auth';
const LAST_ACTIVITY_KEY = 'spmb_alrasyid_last_activity';
const INACTIVITY_TIMEOUT_MS = 60 * 60 * 1000; // 1 jam (60 menit)
const WARNING_BEFORE_LOGOUT_MS = 60 * 1000; // Peringatan 60 detik sebelum logout otomatis

// --- MAIN APPLICATION ---
export default function App() {
  const [sessionMessage, setSessionMessage] = useState<string | null>(() => {
    const isAuth = localStorage.getItem(AUTH_KEY) === 'true';
    const lastAct = Number(localStorage.getItem(LAST_ACTIVITY_KEY) || 0);
    if (isAuth && lastAct > 0 && Date.now() - lastAct >= INACTIVITY_TIMEOUT_MS) {
      localStorage.removeItem(AUTH_KEY);
      localStorage.removeItem(LAST_ACTIVITY_KEY);
      return 'Sesi Anda telah berakhir otomatis karena tidak ada aktivitas selama 1 jam. Silakan login kembali.';
    }
    return null;
  });

  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(() => {
    return localStorage.getItem(AUTH_KEY) === 'true';
  });
  const [activeTab, setActiveTab] = useState('dashboard');
  const [isSidebarOpen, setIsSidebarOpen] = useState<boolean>(() => {
    if (typeof window !== 'undefined') {
      return window.innerWidth >= 768;
    }
    return true;
  });
  const [idleWarningSeconds, setIdleWarningSeconds] = useState<number | null>(null);
  const lastActivityRef = useRef<number>(
    Number(localStorage.getItem(LAST_ACTIVITY_KEY) || 0) || Date.now()
  );

  const handleLoginSuccess = () => {
    const now = Date.now();
    lastActivityRef.current = now;
    localStorage.setItem(AUTH_KEY, 'true');
    localStorage.setItem(LAST_ACTIVITY_KEY, String(now));
    setSessionMessage(null);
    setIdleWarningSeconds(null);
    setIsAuthenticated(true);
  };

  const handleLogout = (isAutoTimeout = false) => {
    localStorage.removeItem(AUTH_KEY);
    localStorage.removeItem(LAST_ACTIVITY_KEY);
    setIdleWarningSeconds(null);
    setIsAuthenticated(false);
    if (isAutoTimeout) {
      setSessionMessage('Sesi Anda telah berakhir otomatis karena tidak ada aktivitas selama 1 jam. Silakan login kembali.');
    } else {
      setSessionMessage(null);
    }
  };

  // Auto-logout saat tidak ada aktivitas selama 1 jam
  useEffect(() => {
    if (!isAuthenticated) return;

    const storedLast = Number(localStorage.getItem(LAST_ACTIVITY_KEY) || 0);
    if (storedLast > 0) {
      lastActivityRef.current = storedLast;
    } else {
      const now = Date.now();
      lastActivityRef.current = now;
      localStorage.setItem(LAST_ACTIVITY_KEY, String(now));
    }

    let lastSavedAt = 0;

    const resetActivity = (force = false) => {
      const now = Date.now();
      lastActivityRef.current = now;
      setIdleWarningSeconds((prev) => (prev !== null ? null : prev));
      if (force || now - lastSavedAt >= 10000) {
        lastSavedAt = now;
        localStorage.setItem(LAST_ACTIVITY_KEY, String(now));
      }
    };

    const handleUserActivity = () => {
      // Saat modal peringatan muncul, user harus menekan tombol "Tetap Masuk" agar tidak tertutup tidak sengaja
      if (Date.now() - lastActivityRef.current >= INACTIVITY_TIMEOUT_MS - WARNING_BEFORE_LOGOUT_MS) {
        return;
      }
      resetActivity(false);
    };

    const checkInactivity = () => {
      const latestStored = Number(localStorage.getItem(LAST_ACTIVITY_KEY) || 0);
      const effectiveLast = Math.max(lastActivityRef.current, latestStored);
      const elapsed = Date.now() - effectiveLast;

      if (elapsed >= INACTIVITY_TIMEOUT_MS) {
        handleLogout(true);
        return;
      }

      if (elapsed >= INACTIVITY_TIMEOUT_MS - WARNING_BEFORE_LOGOUT_MS) {
        const remaining = Math.max(1, Math.ceil((INACTIVITY_TIMEOUT_MS - elapsed) / 1000));
        setIdleWarningSeconds(remaining);
      } else {
        setIdleWarningSeconds((prev) => (prev !== null ? null : prev));
      }
    };

    const handleStorageSync = (e: StorageEvent) => {
      if (e.key === AUTH_KEY && e.newValue !== 'true') {
        setIsAuthenticated(false);
        setIdleWarningSeconds(null);
      } else if (e.key === LAST_ACTIVITY_KEY && e.newValue) {
        const val = Number(e.newValue);
        if (!isNaN(val) && val > lastActivityRef.current) {
          lastActivityRef.current = val;
          setIdleWarningSeconds(null);
        }
      }
    };

    const handleVisibilityOrFocus = () => {
      if (document.visibilityState === 'visible') {
        checkInactivity();
      }
    };

    const events: (keyof WindowEventMap)[] = ['mousemove', 'mousedown', 'keydown', 'scroll', 'touchstart', 'click'];
    events.forEach((ev) => window.addEventListener(ev, handleUserActivity, { passive: true }));
    window.addEventListener('storage', handleStorageSync);
    window.addEventListener('focus', handleVisibilityOrFocus);
    document.addEventListener('visibilitychange', handleVisibilityOrFocus);

    checkInactivity();
    const intervalId = window.setInterval(checkInactivity, 5000);

    return () => {
      window.clearInterval(intervalId);
      events.forEach((ev) => window.removeEventListener(ev, handleUserActivity));
      window.removeEventListener('storage', handleStorageSync);
      window.removeEventListener('focus', handleVisibilityOrFocus);
      document.removeEventListener('visibilitychange', handleVisibilityOrFocus);
    };
  }, [isAuthenticated]);
  
  // Data State for Application - Synchronized with Firebase Firestore
  const [pendaftar, setPendaftar] = useState<any[]>([]);
  const [rombels, setRombels] = useState<any[]>([]);
  const [settings, setSettings] = useState<any>(initialDefaultSettings);
  const [isSyncing, setIsSyncing] = useState<boolean>(false);
  const [syncError, setSyncError] = useState<string | null>(null);

  const pendaftarRef = useRef<any[]>([]);
  const rombelsRef = useRef<any[]>([]);
  const pendingDeleteDocIdsRef = useRef<Set<string>>(new Set());

  useEffect(() => {
    pendaftarRef.current = pendaftar;
  }, [pendaftar]);

  useEffect(() => {
    rombelsRef.current = rombels;
  }, [rombels]);

  useEffect(() => {
    if (!isAuthenticated) return;

    // 1. Subscribe to Registrations in Firestore
    const unsubRegistrations = firestoreService.subscribeRegistrations(
      (items) => {
        const filtered = items.filter(
          (item: any) => !pendingDeleteDocIdsRef.current.has(String(item._docId ?? item.id))
        );
        pendaftarRef.current = filtered;
        setPendaftar(filtered);
        setSyncError(null);
      },
      (error) => {
        console.warn('Real-time connection listener note:', error);
        setSyncError('Gagal menyinkronkan data dengan Firebase Firestore.');
      }
    );

    // 2. Subscribe to Settings in Firestore
    const unsubSettings = firestoreService.subscribeSettings((savedSettings) => {
      if (savedSettings) {
        setSettings(savedSettings);
      }
    });

    // 3. Subscribe to Rombels in Firestore
    const unsubRombels = firestoreService.subscribeRombels((savedRombels) => {
      rombelsRef.current = savedRombels;
      setRombels(savedRombels);
    });

    return () => {
      unsubRegistrations();
      unsubSettings();
      unsubRombels();
    };
  }, [isAuthenticated]);

  // Persistent handler for registrations (CRUD)
  const handleSetPendaftar = async (newValOrFn: any) => {
    setIsSyncing(true);
    setSyncError(null);
    const prevList = pendaftarRef.current;
    const updated = typeof newValOrFn === 'function' ? newValOrFn(prevList) : newValOrFn;
    pendaftarRef.current = updated;
    setPendaftar(updated);

    try {
      if (Array.isArray(updated)) {
        // Detect deletions
        const updatedIds = new Set(updated.map((item: any) => String(item.id)));
        const deletedItems = prevList.filter((oldItem: any) => !updatedIds.has(String(oldItem.id)));

        for (const oldItem of deletedItems) {
          const docId = String(oldItem._docId ?? oldItem.id);
          pendingDeleteDocIdsRef.current.add(docId);
          try {
            await firestoreService.deleteRegistration(docId);
          } finally {
            pendingDeleteDocIdsRef.current.delete(docId);
          }
        }

        // Remove deleted student IDs from any rombel
        if (deletedItems.length > 0) {
          const deletedIdStrings = new Set(deletedItems.map((d: any) => String(d.id)));
          const currentRombels = rombelsRef.current;
          const updatedRombels = currentRombels.map((r: any) => ({
            ...r,
            students: (r.students || []).filter((sid: any) => !deletedIdStrings.has(String(sid)))
          }));
          const hasRombelChanges = updatedRombels.some(
            (r: any, idx: number) => (r.students?.length || 0) !== (currentRombels[idx]?.students?.length || 0)
          );
          if (hasRombelChanges) {
            rombelsRef.current = updatedRombels;
            setRombels(updatedRombels);
            for (const r of updatedRombels) {
              await firestoreService.saveRombel(r);
            }
          }
        }

        // Detect additions or modifications
        for (const item of updated) {
          const prev = prevList.find((p: any) => String(p.id) === String(item.id));
          if (!prev || JSON.stringify(prev) !== JSON.stringify(item)) {
            await firestoreService.saveRegistration(item);
          }
        }
      }
    } catch (e) {
      console.error('Error syncing registration to Firebase:', e);
      setSyncError('Terjadi kesalahan saat menyimpan atau menghapus data di Firebase Firestore.');
    } finally {
      setIsSyncing(false);
    }
  };

  // Explicit permanent delete handler for a single registration
  const handleDeletePendaftar = async (targetReg: any) => {
    setIsSyncing(true);
    setSyncError(null);
    const targetId = targetReg?.id ?? targetReg;
    const docId = String(targetReg?._docId ?? targetId);

    pendingDeleteDocIdsRef.current.add(docId);
    const updated = pendaftarRef.current.filter(
      (item: any) => String(item._docId ?? item.id) !== docId && String(item.id) !== String(targetId)
    );
    pendaftarRef.current = updated;
    setPendaftar(updated);

    try {
      await firestoreService.deleteRegistration(docId);

      // Remove deleted student from any rombel
      const currentRombels = rombelsRef.current;
      const updatedRombels = currentRombels.map((r: any) => ({
        ...r,
        students: (r.students || []).filter((sid: any) => String(sid) !== String(targetId))
      }));
      const hasRombelChanges = updatedRombels.some(
        (r: any, idx: number) => (r.students?.length || 0) !== (currentRombels[idx]?.students?.length || 0)
      );
      if (hasRombelChanges) {
        rombelsRef.current = updatedRombels;
        setRombels(updatedRombels);
        for (const r of updatedRombels) {
          await firestoreService.saveRombel(r);
        }
      }
    } catch (e) {
      console.error('Error deleting registration from Firebase:', e);
      setSyncError('Gagal menghapus data pendaftar secara permanen dari Firebase Firestore.');
      throw e;
    } finally {
      pendingDeleteDocIdsRef.current.delete(docId);
      setIsSyncing(false);
    }
  };

  // Persistent handler for settings
  const handleSetSettings = async (newValOrFn: any) => {
    setIsSyncing(true);
    setSyncError(null);
    const updated = typeof newValOrFn === 'function' ? newValOrFn(settings) : newValOrFn;
    setSettings(updated);
    try {
      await firestoreService.saveSettings(updated);
    } catch (e) {
      console.error('Error syncing settings to Firebase:', e);
      setSyncError('Gagal menyimpan pengaturan ke Firebase Firestore.');
    } finally {
      setIsSyncing(false);
    }
  };

  // Persistent handler for rombels
  const handleSetRombels = async (newValOrFn: any) => {
    setIsSyncing(true);
    setSyncError(null);
    const prevRombels = rombelsRef.current;
    const updated = typeof newValOrFn === 'function' ? newValOrFn(prevRombels) : newValOrFn;
    rombelsRef.current = updated;
    setRombels(updated);
    try {
      if (Array.isArray(updated)) {
        const updatedIds = new Set(updated.map((item: any) => String(item.id)));
        for (const oldItem of prevRombels) {
          if (!updatedIds.has(String(oldItem.id))) {
            await firestoreService.deleteRombel(String(oldItem._docId ?? oldItem.id));
          }
        }
        for (const item of updated) {
          const prev = prevRombels.find((r: any) => String(r.id) === String(item.id));
          if (!prev || JSON.stringify(prev) !== JSON.stringify(item)) {
            await firestoreService.saveRombel(item);
          }
        }
      }
    } catch (e) {
      console.error('Error syncing rombels to Firebase:', e);
      setSyncError('Gagal memperbarui data rombel di Firebase Firestore.');
    } finally {
      setIsSyncing(false);
    }
  };

  const menuItems = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'pendaftar', label: 'Data Pendaftar', icon: Users },
    { id: 'seleksi', label: 'Seleksi & Verifikasi', icon: UserCheck },
    { id: 'siswabaru', label: 'Siswa Diterima', icon: FileArchive },
    { id: 'rombel', label: 'Pembagian Rombel', icon: Layers },
  ];

  if (!isAuthenticated) {
    return (
      <>
        <LoginView onLoginSuccess={handleLoginSuccess} sessionMessage={sessionMessage} />
        <OfflineIndicator />
      </>
    );
  }

  return (
    <div className="flex app-shell-height w-full overflow-hidden bg-slate-50 font-sans text-slate-800 relative">
      {/* Backdrop overlay untuk mobile drawer */}
      {isSidebarOpen && (
        <div
          onClick={() => setIsSidebarOpen(false)}
          className="fixed inset-0 bg-slate-900/50 z-30 md:hidden"
        />
      )}

      {/* Sidebar */}
      <aside
        className={`bg-slate-900 text-slate-300 flex-shrink-0 transition-all duration-300 ease-in-out flex flex-col fixed inset-y-0 left-0 z-40 md:static md:z-auto ${
          isSidebarOpen ? 'w-64 flex' : 'hidden md:flex md:w-20'
        }`}
      >
        <div className="h-16 flex items-center justify-between md:justify-center border-b border-slate-800 px-4">
          <div className="flex items-center min-w-0">
            <img src="https://iili.io/Fr7hdzb.png" alt="Logo Al-Rasyid" className="h-8 w-8 object-contain mr-3 flex-shrink-0" />
            {isSidebarOpen && <span className="font-bold text-lg tracking-tight text-white whitespace-nowrap truncate">SPMB Al-Rasyid</span>}
          </div>
          <button
            type="button"
            onClick={() => setIsSidebarOpen(false)}
            className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 md:hidden"
            aria-label="Tutup Menu"
          >
            <X size={18} />
          </button>
        </div>
        <div className="flex-1 overflow-y-auto py-4 scrollbar-thin">
          <nav className="space-y-1 px-3">
            {menuItems.map(item => (
              <button
                key={item.id}
                onClick={() => {
                  setActiveTab(item.id);
                  if (typeof window !== 'undefined' && window.innerWidth < 768) {
                    setIsSidebarOpen(false);
                  }
                }}
                className={`w-full flex items-center px-3 py-3 rounded-xl transition-all duration-200 ${
                  activeTab === item.id 
                    ? 'bg-blue-600 text-white shadow-md shadow-blue-900/50' 
                    : 'hover:bg-slate-800 hover:text-white'
                } ${!isSidebarOpen && 'justify-center'}`}
                title={!isSidebarOpen ? item.label : ''}
              >
                <item.icon size={20} className={activeTab === item.id ? 'text-white flex-shrink-0' : 'text-slate-400 flex-shrink-0'} />
                {isSidebarOpen && <span className="ml-3 text-sm font-medium whitespace-nowrap">{item.label}</span>}
              </button>
            ))}
          </nav>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col min-w-0 overflow-hidden bg-slate-50/50">
        <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-3 sm:px-6 z-10 flex-shrink-0">
          <div className="flex items-center min-w-0">
            <button onClick={() => setIsSidebarOpen(!isSidebarOpen)} className="p-2 mr-2 sm:mr-3 rounded-lg hover:bg-slate-100 text-slate-500 transition-colors flex-shrink-0" aria-label="Toggle Menu">
              <Menu size={20} />
            </button>
            <h1 className="text-base sm:text-xl font-bold text-slate-800 truncate">
              {menuItems.find(m => m.id === activeTab)?.label}
            </h1>
          </div>
          <div className="flex items-center space-x-2 sm:space-x-3">
            <PWAInstallButton variant="header" />
            <div className="h-8 w-8 rounded-full bg-blue-100 border border-blue-200 flex items-center justify-center text-blue-700 font-bold shadow-sm">A</div>
            <span className="text-sm font-medium text-slate-600 hidden sm:block">Admin</span>
            <Button
              variant="danger"
              size="sm"
              onClick={() => handleLogout(false)}
              className="ml-1 sm:ml-2 bg-red-600 hover:bg-red-700 text-white focus:ring-red-500"
              title="Keluar"
            >
              <LogOut size={16} className="sm:mr-1.5" />
              <span className="hidden sm:inline">Keluar</span>
            </Button>
          </div>
        </header>

        <div className="flex-1 overflow-y-auto p-4 sm:p-6 lg:p-8">
          {syncError && (
            <div className="mb-4 p-3.5 rounded-xl bg-rose-50 border border-rose-200 flex items-center justify-between text-rose-700 text-sm">
              <div className="flex items-center space-x-2.5">
                <AlertCircle size={18} className="flex-shrink-0" />
                <span className="font-medium">{syncError}</span>
              </div>
              <button onClick={() => setSyncError(null)} className="text-rose-500 hover:text-rose-700">
                <X size={16} />
              </button>
            </div>
          )}
          {activeTab === 'dashboard' && <DashboardView data={pendaftar} rombels={rombels} setActiveTab={setActiveTab} settings={settings} />}
          {activeTab === 'pendaftar' && <PendaftarView data={pendaftar} setData={handleSetPendaftar} onDelete={handleDeletePendaftar} />}
          {activeTab === 'seleksi' && <SeleksiView data={pendaftar} setData={handleSetPendaftar} settings={settings} setSettings={handleSetSettings} />}
          {activeTab === 'siswabaru' && <SiswaDiterimaView data={pendaftar} settings={settings} setSettings={handleSetSettings} />}
          {activeTab === 'rombel' && <RombelView data={pendaftar} rombels={rombels} setRombels={handleSetRombels} />}
        </div>
      </main>

      {/* Modal Peringatan Auto-Logout Inaktivitas 1 Jam */}
      {idleWarningSeconds !== null && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-4 animate-in fade-in">
          <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-6 border border-slate-200 animate-in zoom-in-95 duration-200">
            <div className="w-12 h-12 rounded-full bg-amber-100 flex items-center justify-center mb-4 text-amber-600">
              <Clock size={24} />
            </div>
            <h3 className="text-xl font-bold text-slate-800 mb-2">Sesi Akan Berakhir Otomatis</h3>
            <p className="text-slate-600 mb-6 text-sm leading-relaxed">
              Tidak ada aktivitas terdeteksi selama hampir 1 jam. Demi keamanan data, sistem akan keluar secara otomatis dalam{' '}
              <strong className="text-rose-600">{idleWarningSeconds} detik</strong>.
            </p>
            <div className="flex justify-end space-x-3">
              <Button
                variant="danger"
                className="bg-red-600 hover:bg-red-700 text-white focus:ring-red-500"
                onClick={() => handleLogout(false)}
              >
                Keluar Sekarang
              </Button>
              <Button
                onClick={() => {
                  const now = Date.now();
                  lastActivityRef.current = now;
                  localStorage.setItem(LAST_ACTIVITY_KEY, String(now));
                  setIdleWarningSeconds(null);
                }}
              >
                Tetap Masuk
              </Button>
            </div>
          </div>
        </div>
      )}
      <OfflineIndicator />
    </div>
  );
}

// --- VIEWS ---

function DashboardView({ data, rombels, setActiveTab, settings }: any) {
  const total = data.length;
  const diterima = data.filter((d: any) => d.status === 'Diterima').length;
  const cadangan = data.filter((d: any) => d.status === 'Cadangan').length;
  const menunggu = data.filter((d: any) => d.status === 'Menunggu Keputusan').length;
  const ditolak = data.filter((d: any) => d.status === 'Tidak Diterima').length;
  const laki = data.filter((d: any) => d.gender === 'Laki-laki').length;
  const perempuan = data.filter((d: any) => d.gender === 'Perempuan').length;
  const berkasLengkap = data.filter((d: any) => Object.values(d.docs).filter(Boolean).length === 4).length;
  
  const kuotaMaks = Number(settings?.kuotaMaksimal) || 0;
  const isKuotaPenuh = kuotaMaks > 0 && diterima >= kuotaMaks;

  const unassignedCount = data.filter((d: any) => d.status === 'Diterima' && !rombels.some((r: any) => r.students?.includes(d.id))).length;

  const pctLaki = total > 0 ? Math.round((laki / total) * 100) : 0;
  const pctPerempuan = total > 0 ? Math.round((perempuan / total) * 100) : 0;

  const recentData = [...data].reverse().slice(0, 5); // 5 data terbaru

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      {isKuotaPenuh && (
        <div className="p-4 rounded-2xl bg-amber-50 border border-amber-200 flex items-start space-x-3 text-amber-800 shadow-sm">
          <AlertCircle size={22} className="text-amber-600 flex-shrink-0 mt-0.5" />
          <div className="text-sm">
            <p className="font-bold text-amber-900">Pemberitahuan Sistem: Kuota Penerimaan Siswa Telah Terpenuhi ({diterima} / {kuotaMaks} Siswa)</p>
            <p className="text-amber-700 mt-0.5">
              Jumlah siswa yang diterima telah mencapai batas maksimal penerimaan. Pendaftar selanjutnya yang memenuhi kriteria disarankan untuk ditetapkan sebagai <strong>Cadangan</strong>.
            </p>
          </div>
        </div>
      )}

      {/* Baris 1: Kartu Metrik Utama */}
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-6">
        <Card className="p-5 border-l-4 border-l-blue-500 hover:shadow-md transition-shadow">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-slate-500">Total Pendaftar</p>
              <h3 className="text-2xl font-bold text-slate-800 mt-1">{total}</h3>
            </div>
            <div className="p-3 bg-blue-50 rounded-xl text-blue-600"><Users size={24} /></div>
          </div>
        </Card>
        <Card className="p-5 border-l-4 border-l-amber-500 hover:shadow-md transition-shadow">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-slate-500">Berkas Lengkap</p>
              <h3 className="text-2xl font-bold text-slate-800 mt-1">{berkasLengkap}</h3>
            </div>
            <div className="p-3 bg-amber-50 rounded-xl text-amber-600"><FileCheck size={24} /></div>
          </div>
        </Card>
        <Card className="p-5 border-l-4 border-l-emerald-500 hover:shadow-md transition-shadow">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-slate-500">Siswa Diterima (Kuota)</p>
              <h3 className="text-2xl font-bold text-slate-800 mt-1">
                {diterima} <span className="text-sm font-medium text-slate-400">/ {kuotaMaks}</span>
              </h3>
            </div>
            <div className="p-3 bg-emerald-50 rounded-xl text-emerald-600"><UserCheck size={24} /></div>
          </div>
        </Card>
        <Card className="p-5 border-l-4 border-l-purple-500 hover:shadow-md transition-shadow">
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-slate-500">Belum Dapat Kelas</p>
              <h3 className="text-2xl font-bold text-slate-800 mt-1">{unassignedCount}</h3>
            </div>
            <div className="p-3 bg-purple-50 rounded-xl text-purple-600"><Layers size={24} /></div>
          </div>
        </Card>
      </div>

      {/* Baris 2: Detail Laporan */}
      <div className="grid grid-cols-1 xl:grid-cols-3 gap-6">
        
        {/* Kolom Kiri (Lebar) */}
        <div className="xl:col-span-2 space-y-6">
          <Card className="p-6">
            <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center">
              <Activity className="mr-2 text-slate-500" size={18}/> Status Seleksi
            </h3>
            <div className="space-y-4 mt-2">
              {[
                { label: 'Menunggu Keputusan', count: menunggu, color: 'bg-amber-500', bg: 'bg-amber-100' },
                { label: 'Diterima', count: diterima, color: 'bg-emerald-500', bg: 'bg-emerald-100' },
                { label: 'Cadangan', count: cadangan, color: 'bg-blue-500', bg: 'bg-blue-100' },
                { label: 'Tidak Diterima', count: ditolak, color: 'bg-rose-500', bg: 'bg-rose-100' }
              ].map(stat => {
                const pct = total > 0 ? Math.round((stat.count / total) * 100) : 0;
                return (
                  <div key={stat.label}>
                    <div className="flex justify-between text-sm mb-1">
                      <span className="font-medium text-slate-700">{stat.label}</span>
                      <span className="text-slate-500">{stat.count} Siswa ({pct}%)</span>
                    </div>
                    <div className={`w-full ${stat.bg} rounded-full h-2.5`}>
                      <div className={`${stat.color} h-2.5 rounded-full transition-all duration-500`} style={{ width: `${pct}%` }}></div>
                    </div>
                  </div>
                )
              })}
            </div>
          </Card>

          <Card className="overflow-hidden">
            <div className="p-5 border-b border-slate-100 flex justify-between items-center">
              <h3 className="text-base font-bold text-slate-800 flex items-center">
                <Clock className="mr-2 text-slate-500" size={18}/> Pendaftar Terbaru
              </h3>
              <button onClick={() => setActiveTab('pendaftar')} className="text-sm text-blue-600 hover:text-blue-700 font-medium flex items-center">
                Lihat Semua <ArrowRight size={14} className="ml-1" />
              </button>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm text-slate-600">
                <thead className="bg-slate-50 text-slate-700 font-medium">
                  <tr>
                    <th className="px-5 py-3">Nama Pendaftar</th>
                    <th className="px-5 py-3">No. Daftar</th>
                    <th className="px-5 py-3">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {recentData.map((reg: any) => (
                    <tr key={reg.id} className="hover:bg-slate-50 transition-colors">
                      <td className="px-5 py-3 font-medium text-slate-800">{reg.name}</td>
                      <td className="px-5 py-3">{reg.regNumber}</td>
                      <td className="px-5 py-3">
                        <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium 
                          ${reg.status === 'Diterima' ? 'bg-emerald-100 text-emerald-700' : 
                            reg.status === 'Cadangan' ? 'bg-blue-100 text-blue-700' :
                            reg.status === 'Tidak Diterima' ? 'bg-rose-100 text-rose-700' : 'bg-slate-100 text-slate-700'}`}>
                          {reg.status}
                        </span>
                      </td>
                    </tr>
                  ))}
                  {recentData.length === 0 && <tr><td colSpan={3} className="px-5 py-8 text-center text-slate-500">Belum ada data</td></tr>}
                </tbody>
              </table>
            </div>
          </Card>
        </div>

        {/* Kolom Kanan (Kecil) */}
        <div className="space-y-6">
          <Card className="p-6">
            <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center">
              <Users className="mr-2 text-slate-500" size={18}/> Komposisi Gender
            </h3>
            <div className="space-y-5">
              <div className="relative pt-2">
                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-sm font-bold text-blue-700">Laki-laki</p>
                    <p className="text-xs text-slate-500 mt-0.5">{laki} Siswa</p>
                  </div>
                  <span className="text-lg font-bold text-blue-600">{pctLaki}%</span>
                </div>
                <div className="w-full bg-blue-50 rounded-full h-2 mt-2 overflow-hidden">
                  <div className="bg-blue-500 h-2 rounded-full transition-all duration-500" style={{ width: `${pctLaki}%` }}></div>
                </div>
              </div>
              
              <div className="relative pt-2">
                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-sm font-bold text-rose-600">Perempuan</p>
                    <p className="text-xs text-slate-500 mt-0.5">{perempuan} Siswa</p>
                  </div>
                  <span className="text-lg font-bold text-rose-500">{pctPerempuan}%</span>
                </div>
                <div className="w-full bg-rose-50 rounded-full h-2 mt-2 overflow-hidden">
                  <div className="bg-rose-400 h-2 rounded-full transition-all duration-500" style={{ width: `${pctPerempuan}%` }}></div>
                </div>
              </div>
            </div>
          </Card>
        </div>

      </div>
    </div>
  );
}

function PendaftarView({ data, setData, onDelete }: any) {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [step, setStep] = useState(1);
  const [searchTerm, setSearchTerm] = useState('');
  const [deleteTarget, setDeleteTarget] = useState<any>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  
  const initialForm: any = {
    regNumber: '', name: '', nickname: '', gender: '', birthPlace: '', birthDate: '',
    religion: '', nationality: '', childOrder: '', siblingsCount: '', familyStatus: '', dailyLanguage: '',
    siblingsDetail: { kandung: '', sebapak: '', seibu: '', tiri: '', angkat: '' },
    livingWith: '', distanceToSchool: '',
    physical: { weight: '', height: '', bloodType: '', chronicDisease: '' },
    address: '', phone: '',
    previousSchool: { name: '', address: '' },
    parent: { fatherName: '', fatherEdu: '', fatherJob: '', motherName: '', motherEdu: '', motherJob: '', parentAddress: '', income: '' },
    guardian: { hasGuardian: false, relation: '', edu: '', job: '', phone: '', address: '' },
    emergencyContact: { name: '', relation: '', phone: '', address: '' },
    specialNeeds: '', achievements: '', latinReading: '', quranReading: '', transport: '', suggestions: '',
    docs: { akta: false, kk: false, materai: false, foto: false },
    status: 'Menunggu Keputusan'
  };

  const [formData, setFormData] = useState<any>(initialForm);

  const eduOptions = [{value:'SD', label:'SD'}, {value:'SMP', label:'SMP'}, {value:'SMA/SMK', label:'SMA/SMK'}, {value:'D3', label:'D3'}, {value:'S1', label:'S1'}, {value:'S2', label:'S2'}, {value:'S3', label:'S3'}];
  const incomeOptions = [{value:'< Rp 1.000.000', label:'< Rp 1.000.000'}, {value:'Rp 1.000.000 - Rp 3.000.000', label:'Rp 1.000.000 - Rp 3.000.000'}, {value:'Rp 3.000.000 - Rp 5.000.000', label:'Rp 3.000.000 - Rp 5.000.000'}, {value:'> Rp 5.000.000', label:'> Rp 5.000.000'}];

  const filteredData = data.filter((d: any) => d.name.toLowerCase().includes(searchTerm.toLowerCase()) || d.regNumber.toLowerCase().includes(searchTerm.toLowerCase()));

  const handleSubmit = (e: any) => {
    e.preventDefault();
    if (step < 3) {
      setStep(step + 1);
      return;
    }
    // Final Save or Update
    if (formData.id !== undefined && formData.id !== null && formData.id !== '') {
      setData(data.map((d: any) => String(d.id) === String(formData.id) ? formData : d));
    } else {
      const numericIds = data.map((d: any) => Number(d.id)).filter((n: number) => !isNaN(n));
      const newId = numericIds.length ? Math.max(...numericIds) + 1 : Date.now();
      setData([...data, { ...formData, id: newId }]);
    }
    setIsFormOpen(false);
    setStep(1);
    setFormData(initialForm);
  };

  const handleEdit = (reg: any) => {
    // Merge standard form structure with existing data to prevent undefined errors on nested fields
    setFormData({
      ...initialForm,
      ...reg,
      siblingsDetail: { ...initialForm.siblingsDetail, ...(reg.siblingsDetail || {}) },
      physical: { ...initialForm.physical, ...(reg.physical || {}) },
      previousSchool: { ...initialForm.previousSchool, ...(reg.previousSchool || {}) },
      parent: { ...initialForm.parent, ...(reg.parent || {}) },
      guardian: { ...initialForm.guardian, ...(reg.guardian || {}) },
      emergencyContact: { ...initialForm.emergencyContact, ...(reg.emergencyContact || {}) },
      docs: { ...initialForm.docs, ...(reg.docs || {}) }
    });
    setStep(1);
    setIsFormOpen(true);
  };

  const confirmDelete = async () => {
    if (!deleteTarget || isDeleting) return;
    setIsDeleting(true);
    try {
      if (onDelete) {
        await onDelete(deleteTarget);
      } else {
        await setData((prevData: any[]) => prevData.filter((d: any) => String(d.id) !== String(deleteTarget.id)));
      }
      setDeleteTarget(null);
    } catch (err) {
      console.error('Delete failed:', err);
    } finally {
      setIsDeleting(false);
    }
  };

  const handleExport = () => {
    const cols = ['No', 'No. Daftar', 'Nama Lengkap', 'Jenis Kelamin', 'Tempat Lahir', 'Tanggal Lahir', 'Nama Ayah', 'Nama Ibu', 'Status Berkas'];
    const rows = filteredData.map((d: any, i: number) => {
      const isLengkap = Object.values(d.docs).filter(Boolean).length === 4;
      return [
        i + 1, d.regNumber, d.name, d.gender, d.birthPlace, d.birthDate, d.parent.fatherName, d.parent.motherName,
        isLengkap ? 'Persyaratan Lengkap' : 'Persyaratan Belum Lengkap'
      ];
    });
    exportToExcel('DATA PENDAFTAR SPMB SD AL-RASYID', cols, rows, 'Data_Pendaftar');
  };

  if (isFormOpen) {
    return (
      <div className="space-y-6 animate-in slide-in-from-right-8 duration-300">
        <div className="flex justify-between items-center bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <h2 className="text-xl font-bold text-slate-800">Pendaftar Baru</h2>
          <Button variant="ghost" onClick={() => { setIsFormOpen(false); setStep(1); setFormData(initialForm); }} className="text-rose-600 hover:bg-rose-50">
            <X size={20} /> Batal
          </Button>
        </div>

        <div className="flex gap-2">
          {[1, 2, 3].map(s => (
            <div key={s} className={`flex-1 h-2 rounded-full transition-all duration-300 ${s <= step ? 'bg-blue-600' : 'bg-slate-200'}`} />
          ))}
        </div>

        <Card>
          <form onSubmit={handleSubmit}>
            {step === 1 && (
              <div className="p-6 space-y-6">
                <h3 className="font-bold text-lg text-slate-800 flex items-center border-b pb-2"><UserPlus size={20} className="mr-2 text-blue-600"/> A. Data Peserta Didik</h3>
                
                <div className="bg-blue-50/50 p-4 rounded-xl border border-blue-100">
                  <Input label="Nomor Pendaftaran" value={formData.regNumber} onChange={(e: any) => setFormData({...formData, regNumber: e.target.value})} required placeholder="Ketik nomor pendaftaran manual di sini..." />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <Input label="1. a. Nama Lengkap" value={formData.name} onChange={(e: any) => setFormData({...formData, name: e.target.value})} required />
                  <Input label="b. Nama Panggilan" value={formData.nickname} onChange={(e: any) => setFormData({...formData, nickname: e.target.value})} />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <Select label="2. Jenis Kelamin" options={[{value:'Laki-laki', label:'Laki-laki'}, {value:'Perempuan', label:'Perempuan'}]} value={formData.gender} onChange={(e: any) => setFormData({...formData, gender: e.target.value})} required />
                  <Input label="3. Tempat Lahir" value={formData.birthPlace} onChange={(e: any) => setFormData({...formData, birthPlace: e.target.value})} required />
                  <Input label="4. Tanggal Lahir" type="date" value={formData.birthDate} onChange={(e: any) => setFormData({...formData, birthDate: e.target.value})} required />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <Select label="5. Agama" options={[{value:'Islam', label:'Islam'}]} value={formData.religion} onChange={(e: any) => setFormData({...formData, religion: e.target.value})} />
                  <Input label="6. Kewarganegaraan" value={formData.nationality} onChange={(e: any) => setFormData({...formData, nationality: e.target.value})} />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <Input label="7. Anak Ke" type="number" value={formData.childOrder} onChange={(e: any) => setFormData({...formData, childOrder: e.target.value})} />
                  <Input label="Dari (Bersaudara)" type="number" value={formData.siblingsCount} onChange={(e: any) => setFormData({...formData, siblingsCount: e.target.value})} />
                </div>

                <div className="p-4 border border-slate-100 bg-slate-50/50 rounded-xl space-y-4">
                  <label className="text-sm font-semibold text-slate-700">8. Jumlah Saudara</label>
                  <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                    <Input label="a. Kandung" type="number" value={formData.siblingsDetail.kandung} onChange={(e: any) => setFormData({...formData, siblingsDetail: {...formData.siblingsDetail, kandung: e.target.value}})} />
                    <Input label="b. Sebapak" type="number" value={formData.siblingsDetail.sebapak} onChange={(e: any) => setFormData({...formData, siblingsDetail: {...formData.siblingsDetail, sebapak: e.target.value}})} />
                    <Input label="c. Seibu" type="number" value={formData.siblingsDetail.seibu} onChange={(e: any) => setFormData({...formData, siblingsDetail: {...formData.siblingsDetail, seibu: e.target.value}})} />
                    <Input label="d. Tiri" type="number" value={formData.siblingsDetail.tiri} onChange={(e: any) => setFormData({...formData, siblingsDetail: {...formData.siblingsDetail, tiri: e.target.value}})} />
                    <Input label="e. Angkat" type="number" value={formData.siblingsDetail.angkat} onChange={(e: any) => setFormData({...formData, siblingsDetail: {...formData.siblingsDetail, angkat: e.target.value}})} />
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <Input label="9. Status dalam keluarga" value={formData.familyStatus} onChange={(e: any) => setFormData({...formData, familyStatus: e.target.value})} placeholder="Misal: Anak Kandung" />
                  <Input label="10. Bahasa sehari-hari" value={formData.dailyLanguage} onChange={(e: any) => setFormData({...formData, dailyLanguage: e.target.value})} />
                </div>

                <div className="grid grid-cols-1 gap-4">
                  <Select label="11. Bertempat Tinggal pada" options={[{value:'Orangtua',label:'Orangtua'},{value:'Menumpang',label:'Menumpang'},{value:'Asrama',label:'Asrama'}]} value={formData.livingWith} onChange={(e: any) => setFormData({...formData, livingWith: e.target.value})} />
                </div>

                <div className="p-4 border border-slate-100 bg-slate-50/50 rounded-xl space-y-4">
                  <label className="text-sm font-semibold text-slate-700 block">12. Keadaan Jasmani</label>
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    <Input label="a. Berat Badan (kg)" type="number" value={formData.physical.weight} onChange={(e: any) => setFormData({...formData, physical: {...formData.physical, weight: e.target.value}})} />
                    <Input label="b. Tinggi Badan (cm)" type="number" value={formData.physical.height} onChange={(e: any) => setFormData({...formData, physical: {...formData.physical, height: e.target.value}})} />
                    <Select label="c. Golongan Darah" options={[{value:'A',label:'A'},{value:'B',label:'B'},{value:'AB',label:'AB'},{value:'O',label:'O'},{value:'-',label:'Tidak Tahu'}]} value={formData.physical.bloodType} onChange={(e: any) => setFormData({...formData, physical: {...formData.physical, bloodType: e.target.value}})} />
                    <Input label="d. Penyakit (Kronis)" placeholder="Kosongkan jika tidak ada" value={formData.physical.chronicDisease} onChange={(e: any) => setFormData({...formData, physical: {...formData.physical, chronicDisease: e.target.value}})} />
                  </div>
                </div>

                <div className="space-y-4 p-4 border border-slate-100 bg-slate-50/50 rounded-xl">
                  <Textarea label="13. Alamat Tempat Tinggal" value={formData.address} onChange={(e: any) => setFormData({...formData, address: e.target.value})} required />
                  <Input label="Telepon" value={formData.phone} onChange={(e: any) => setFormData({...formData, phone: e.target.value})} />
                </div>

                <div className="grid grid-cols-1 gap-4">
                  <Input label="14. Jarak Tempat Tinggal ke Sekolah (km)" type="number" step="0.1" value={formData.distanceToSchool} onChange={(e: any) => setFormData({...formData, distanceToSchool: e.target.value})} />
                </div>

                <div className="space-y-4 p-4 border border-slate-100 bg-slate-50/50 rounded-xl">
                  <Input label="15. a. Nama Sekolah Asal" value={formData.previousSchool.name} onChange={(e: any) => setFormData({...formData, previousSchool: {...formData.previousSchool, name: e.target.value}})} />
                  <Textarea label="b. Alamat Sekolah Asal" value={formData.previousSchool.address} onChange={(e: any) => setFormData({...formData, previousSchool: {...formData.previousSchool, address: e.target.value}})} />
                </div>
              </div>
            )}

            {step === 2 && (
              <div className="p-6 space-y-6">
                <h3 className="font-bold text-lg text-slate-800 flex items-center border-b pb-2"><Users size={20} className="mr-2 text-blue-600"/> B. Data Orangtua / Wali / Keluarga</h3>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 p-4 border border-slate-100 rounded-xl bg-slate-50/50">
                  <Input label="1. a. Nama Ayah Kandung" value={formData.parent.fatherName} onChange={(e: any) => setFormData({...formData, parent: {...formData.parent, fatherName: e.target.value}})} required />
                  <Input label="b. Nama Ibu Kandung" value={formData.parent.motherName} onChange={(e: any) => setFormData({...formData, parent: {...formData.parent, motherName: e.target.value}})} required />
                </div>

                <div className="grid grid-cols-1 gap-4">
                  <Textarea label="2. Alamat Orangtua" value={formData.parent.parentAddress} onChange={(e: any) => setFormData({...formData, parent: {...formData.parent, parentAddress: e.target.value}})} required />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 p-4 border border-slate-100 rounded-xl bg-slate-50/50">
                  <Select label="3. a. Pendidikan Tertinggi Ayah" options={eduOptions} value={formData.parent.fatherEdu} onChange={(e: any) => setFormData({...formData, parent: {...formData.parent, fatherEdu: e.target.value}})} />
                  <Select label="b. Pendidikan Tertinggi Ibu" options={eduOptions} value={formData.parent.motherEdu} onChange={(e: any) => setFormData({...formData, parent: {...formData.parent, motherEdu: e.target.value}})} />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 p-4 border border-slate-100 rounded-xl bg-slate-50/50">
                  <Input label="4. a. Pekerjaan Ayah" value={formData.parent.fatherJob} onChange={(e: any) => setFormData({...formData, parent: {...formData.parent, fatherJob: e.target.value}})} />
                  <Input label="b. Pekerjaan Ibu" value={formData.parent.motherJob} onChange={(e: any) => setFormData({...formData, parent: {...formData.parent, motherJob: e.target.value}})} />
                </div>

                <div className="grid grid-cols-1 gap-4">
                  <Select label="5. Penghasilan per bulan" options={incomeOptions} value={formData.parent.income} onChange={(e: any) => setFormData({...formData, parent: {...formData.parent, income: e.target.value}})} />
                </div>

                <div className="space-y-4 p-4 border border-slate-100 rounded-xl bg-slate-50/50 mt-4">
                  <div className="flex items-center justify-between border-b pb-2">
                    <h4 className="font-semibold text-slate-700">6. Wali Peserta Didik (Jika Mempunyai)</h4>
                    <label className="flex items-center space-x-2 text-sm cursor-pointer bg-white px-3 py-1.5 rounded-lg border shadow-sm">
                      <input type="checkbox" checked={formData.guardian.hasGuardian} onChange={(e: any) => setFormData({...formData, guardian: {...formData.guardian, hasGuardian: e.target.checked}})} className="w-4 h-4 text-blue-600 rounded" />
                      <span className="font-medium text-slate-700">Isi Data Wali</span>
                    </label>
                  </div>
                  {formData.guardian.hasGuardian && (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-2">
                      <Input label="a. Hubungan Keluarga" value={formData.guardian.relation} onChange={(e: any) => setFormData({...formData, guardian: {...formData.guardian, relation: e.target.value}})} />
                      <Select label="b. Pendidikan Tertinggi" options={eduOptions} value={formData.guardian.edu} onChange={(e: any) => setFormData({...formData, guardian: {...formData.guardian, edu: e.target.value}})} />
                      <Input label="c. Pekerjaan/Jabatan" value={formData.guardian.job} onChange={(e: any) => setFormData({...formData, guardian: {...formData.guardian, job: e.target.value}})} />
                      <Textarea label="d. Alamat" rows={2} value={formData.guardian.address} onChange={(e: any) => setFormData({...formData, guardian: {...formData.guardian, address: e.target.value}})} />
                      <Input label="e. Telepon" className="md:col-span-2" value={formData.guardian.phone} onChange={(e: any) => setFormData({...formData, guardian: {...formData.guardian, phone: e.target.value}})} />
                    </div>
                  )}
                </div>

                <div className="space-y-4 p-4 border border-slate-100 rounded-xl bg-slate-50/50">
                  <h4 className="font-semibold text-slate-700 border-b pb-2">7. Keluarga/Saudara Terdekat (Bisa Dihubungi)</h4>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <Input label="a. Nama" value={formData.emergencyContact.name} onChange={(e: any) => setFormData({...formData, emergencyContact: {...formData.emergencyContact, name: e.target.value}})} required />
                    <Textarea label="b. Alamat" rows={2} value={formData.emergencyContact.address} onChange={(e: any) => setFormData({...formData, emergencyContact: {...formData.emergencyContact, address: e.target.value}})} required />
                    <Input label="c. Hubungan Keluarga" value={formData.emergencyContact.relation} onChange={(e: any) => setFormData({...formData, emergencyContact: {...formData.emergencyContact, relation: e.target.value}})} required />
                    <Input label="d. Telepon" value={formData.emergencyContact.phone} onChange={(e: any) => setFormData({...formData, emergencyContact: {...formData.emergencyContact, phone: e.target.value}})} required />
                  </div>
                </div>
              </div>
            )}

            {step === 3 && (
              <div className="p-6 space-y-6">
                <h3 className="font-bold text-lg text-slate-800 flex items-center border-b pb-2"><BookOpen size={20} className="mr-2 text-blue-600"/> C. Perkembangan & Berkas</h3>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-6">
                    <div className="space-y-4 p-4 border border-slate-100 rounded-xl bg-slate-50/50">
                      <h4 className="font-semibold text-slate-700 border-b pb-2">C. Perkembangan Peserta Didik</h4>
                      <Input label="1. Kondisi anak (perlu perhatian khusus)" value={formData.specialNeeds} onChange={(e: any) => setFormData({...formData, specialNeeds: e.target.value})} placeholder="Kosongkan jika tidak ada" />
                      <Input label="2. Prestasi yang pernah diraih" value={formData.achievements} onChange={(e: any) => setFormData({...formData, achievements: e.target.value})} placeholder="Kosongkan jika tidak ada" />
                      <Select label="3. Kemampuan membaca huruf latin" options={[{value:'Lancar membaca',label:'Lancar membaca'}, {value:'Belum lancar',label:'Belum lancar'}, {value:'Sudah mengenal huruf A-Z',label:'Sudah mengenal huruf A-Z'}]} value={formData.latinReading} onChange={(e: any) => setFormData({...formData, latinReading: e.target.value})} />
                      <Select label="4. Kemampuan membaca Al-Qur'an" options={[{value:"Iqro' 1-3",label:"Iqro' 1-3"}, {value:"Iqro' 4-6",label:"Iqro' 4-6"}, {value:'Al-Qur\'an',label:'Al-Qur\'an'}]} value={formData.quranReading} onChange={(e: any) => setFormData({...formData, quranReading: e.target.value})} />
                      <Select label="5. Layanan kendaraan antar jemput" options={[{value:'Memerlukan',label:'Memerlukan'}, {value:'Tidak',label:'Tidak'}]} value={formData.transport} onChange={(e: any) => setFormData({...formData, transport: e.target.value})} />
                    </div>

                    <div className="space-y-4 p-4 border border-slate-100 rounded-xl bg-slate-50/50">
                      <h4 className="font-semibold text-slate-700 border-b pb-2">D. Saran & Harapan</h4>
                      <Textarea label="Saran dan Harapan Orangtua terhadap SD Al-Rasyid" rows={4} value={formData.suggestions} onChange={(e: any) => setFormData({...formData, suggestions: e.target.value})} placeholder="Tuliskan saran dan harapan Anda di sini..." />
                    </div>
                  </div>

                  <div className="space-y-4 p-4 border border-slate-100 rounded-xl bg-slate-50/50 h-fit">
                    <h4 className="font-semibold text-slate-700 border-b pb-2">Checklist Berkas Fisik Pendaftaran</h4>
                    <div className="space-y-3">
                      {[
                        { id: 'akta', label: 'Fotocopy Akta Kelahiran' },
                        { id: 'kk', label: 'Fotocopy Kartu Keluarga' },
                        { id: 'materai', label: 'Surat Pernyataan Bermaterai 10.000' },
                        { id: 'foto', label: 'Pas Foto Berwarna 3x4 (2 Lembar)' }
                      ].map(doc => (
                        <label key={doc.id} className="flex items-center space-x-3 p-3 border rounded-xl bg-white hover:bg-slate-50 cursor-pointer transition-colors shadow-sm">
                          <input 
                            type="checkbox" 
                            checked={formData.docs[doc.id]} 
                            onChange={(e: any) => setFormData({...formData, docs: {...formData.docs, [doc.id]: e.target.checked}})}
                            className="w-5 h-5 text-blue-600 rounded border-slate-300 focus:ring-blue-500"
                          />
                          <span className="text-sm font-medium text-slate-700">{doc.label}</span>
                        </label>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            )}

            <div className="p-6 bg-slate-50 border-t border-slate-100 flex justify-between items-center rounded-b-xl">
              {step > 1 ? (
                <Button type="button" variant="outline" onClick={() => setStep(step - 1)} className="text-slate-600"><ChevronLeft size={18} className="mr-1" /> Sebelumnya</Button>
              ) : <div></div>}
              
              {step === 3 ? (
                <Button type="submit" className="bg-blue-600 hover:bg-blue-700 ml-auto text-white shadow-md"><Save size={18} className="mr-2" /> Simpan Pendaftar</Button>
              ) : (
                <Button type="submit" className="ml-auto bg-slate-800 hover:bg-slate-900 text-white">Selanjutnya <ChevronRight size={18} className="ml-1" /></Button>
              )}
            </div>
          </form>
        </Card>
      </div>
    );
  }

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
        <div className="relative w-full sm:w-80">
          <div className="pointer-events-none absolute inset-y-0 left-3 flex items-center text-slate-400">
            <Search size={18} />
          </div>
          <input 
            type="text" placeholder="Cari nama atau no pendaftaran..." 
            className="w-full pl-10 pr-4 py-2.5 bg-white border border-slate-200 rounded-xl text-sm focus:outline-none focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10 transition-all shadow-sm"
            value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        <div className="flex space-x-3 w-full sm:w-auto">
          <Button variant="outline" onClick={handleExport} className="w-full sm:w-auto text-emerald-700 border-emerald-200 hover:bg-emerald-50">
            <Download size={18} className="mr-2" /> Export Excel
          </Button>
          <Button onClick={() => setIsFormOpen(true)} className="w-full sm:w-auto">
            <Plus size={18} className="mr-2" /> Tambah Pendaftar
          </Button>
        </div>
      </div>

      <Card className="overflow-hidden p-0">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm text-slate-600">
            <thead className="bg-slate-50 text-slate-700 font-semibold border-b border-slate-200">
              <tr>
                <th className="px-4 py-3">No. Daftar</th>
                <th className="px-4 py-3">Nama Pendaftar</th>
                <th className="px-4 py-3">L/P</th>
                <th className="px-4 py-3">Berkas Fisik</th>
                <th className="px-4 py-3 text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredData.length > 0 ? filteredData.map((reg: any) => {
                const isLengkap = Object.values(reg.docs || {}).filter(Boolean).length === 4;
                return (
                  <tr key={reg._docId ?? reg.id} className="hover:bg-slate-50/80 transition-colors">
                    <td className="px-4 py-3 font-medium text-slate-800">{reg.regNumber}</td>
                    <td className="px-4 py-3">{reg.name}</td>
                    <td className="px-4 py-3">{reg.gender === 'Laki-laki' ? 'L' : 'P'}</td>
                    <td className="px-4 py-3">
                      <span className={`inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium ${isLengkap ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700'}`}>
                        {isLengkap ? 'Persyaratan Lengkap' : 'Persyaratan Belum Lengkap'}
                      </span>
                    </td>
                    <td className="px-4 py-3 text-right space-x-2">
                      <button onClick={() => handleEdit(reg)} className="p-1.5 text-blue-600 hover:bg-blue-50 rounded-md transition-colors" title="Edit Data">
                        <Edit size={18} />
                      </button>
                      <button onClick={() => setDeleteTarget(reg)} className="p-1.5 text-rose-600 hover:bg-rose-50 rounded-md transition-colors" title="Hapus Data">
                        <Trash2 size={18} />
                      </button>
                    </td>
                  </tr>
                )
              }) : (
                <tr><td colSpan={5} className="px-4 py-12 text-center text-slate-500">Tidak ada data pendaftar yang ditemukan.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>

      {/* Modal Konfirmasi Hapus */}
      {deleteTarget && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center z-50 p-4 animate-in fade-in">
          <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-6 animate-in zoom-in-95 duration-200">
            <div className="w-12 h-12 rounded-full bg-rose-100 flex items-center justify-center mb-4 text-rose-600">
              <AlertCircle size={24} />
            </div>
            <h3 className="text-xl font-bold text-slate-800 mb-2">Hapus Data Pendaftar?</h3>
            <p className="text-slate-500 mb-6 text-sm leading-relaxed">
              Data pendaftar <strong>{deleteTarget.name}</strong> ({deleteTarget.regNumber}) akan dihapus secara permanen dari Firebase Firestore dan tidak dapat dikembalikan. Apakah Anda yakin?
            </p>
            <div className="flex justify-end space-x-3">
              <Button variant="ghost" disabled={isDeleting} onClick={() => setDeleteTarget(null)}>Batal</Button>
              <Button variant="danger" disabled={isDeleting} onClick={confirmDelete}>
                {isDeleting ? 'Menghapus...' : 'Ya, Hapus Permanen'}
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function SeleksiView({ data, setData, settings }: any) {
  const [quotaAlert, setQuotaAlert] = useState<string | null>(null);
  const diterimaCount = data.filter((d: any) => d.status === 'Diterima').length;
  const cadanganCount = data.filter((d: any) => d.status === 'Cadangan').length;
  const kuotaMaks = Number(settings?.kuotaMaksimal) || 0;
  const isKuotaPenuh = kuotaMaks > 0 && diterimaCount >= kuotaMaks;

  const updateStatus = (id: any, newStatus: string) => {
    const currentStudent = data.find((d: any) => d.id === id);
    if (
      newStatus === 'Diterima' &&
      currentStudent?.status !== 'Diterima' &&
      kuotaMaks > 0 &&
      diterimaCount >= kuotaMaks
    ) {
      setQuotaAlert(
        `Batas maksimal penerimaan siswa (${kuotaMaks} siswa) telah tercapai! Silakan ubah status calon siswa ini menjadi "Cadangan".`
      );
      return;
    }
    setQuotaAlert(null);
    setData(data.map((d: any) => (d.id === id ? { ...d, status: newStatus } : d)));
  };

  return (
    <div className="space-y-5 animate-in fade-in duration-300">
      {/* Info Kuota & Pemberitahuan Sistem */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <Card className="p-4 flex items-center justify-between border-l-4 border-l-emerald-500">
          <div>
            <p className="text-xs font-semibold uppercase tracking-wider text-slate-400">Kuota Diterima Terisi</p>
            <p className="text-xl font-bold text-slate-800 mt-0.5">
              {diterimaCount} <span className="text-sm font-normal text-slate-500">/ {kuotaMaks} Siswa</span>
            </p>
          </div>
          <div className="p-2.5 bg-emerald-50 text-emerald-600 rounded-xl">
            <UserCheck size={20} />
          </div>
        </Card>
        <Card className="p-4 flex items-center justify-between border-l-4 border-l-blue-500">
          <div>
            <p className="text-xs font-semibold uppercase tracking-wider text-slate-400">Daftar Cadangan</p>
            <p className="text-xl font-bold text-slate-800 mt-0.5">
              {cadanganCount} <span className="text-sm font-normal text-slate-500">Siswa</span>
            </p>
          </div>
          <div className="p-2.5 bg-blue-50 text-blue-600 rounded-xl">
            <Clock size={20} />
          </div>
        </Card>
        <Card className="p-4 flex items-center justify-between border-l-4 border-l-amber-500">
          <div>
            <p className="text-xs font-semibold uppercase tracking-wider text-slate-400">Sisa Kuota Penerimaan</p>
            <p className="text-xl font-bold text-slate-800 mt-0.5">
              {Math.max(0, kuotaMaks - diterimaCount)} <span className="text-sm font-normal text-slate-500">Kursi Tersedia</span>
            </p>
          </div>
          <div className="p-2.5 bg-amber-50 text-amber-600 rounded-xl">
            <Users size={20} />
          </div>
        </Card>
      </div>

      {isKuotaPenuh && (
        <div className="p-4 rounded-2xl bg-amber-50 border border-amber-200 flex items-start space-x-3 text-amber-800 shadow-sm">
          <AlertCircle size={20} className="text-amber-600 flex-shrink-0 mt-0.5" />
          <div className="text-sm">
            <p className="font-bold text-amber-900">Pemberitahuan Batas Kuota Penerimaan ({diterimaCount} / {kuotaMaks} Siswa)</p>
            <p className="text-amber-700 mt-0.5">
              Jumlah siswa dengan status <strong>Diterima</strong> sudah mencapai batas maksimal penerimaan. Gunakan opsi <strong>Cadangan</strong> untuk calon siswa pengganti apabila ada siswa yang mengundurkan diri.
            </p>
          </div>
        </div>
      )}

      {quotaAlert && (
        <div className="p-4 rounded-2xl bg-rose-50 border border-rose-200 flex items-center justify-between text-rose-800 shadow-sm animate-in fade-in duration-200">
          <div className="flex items-start space-x-3">
            <XCircle size={20} className="text-rose-600 flex-shrink-0 mt-0.5" />
            <p className="text-sm font-medium">{quotaAlert}</p>
          </div>
          <button onClick={() => setQuotaAlert(null)} className="text-rose-400 hover:text-rose-600 p-1">
            <X size={16} />
          </button>
        </div>
      )}

      <Card className="overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm text-slate-600">
            <thead className="bg-slate-50 text-slate-700 font-semibold border-b border-slate-200">
              <tr>
                <th className="px-4 py-3">Nama Pendaftar</th>
                <th className="px-4 py-3">Kelengkapan Berkas</th>
                <th className="px-4 py-3">Status Seleksi</th>
                <th className="px-4 py-3 text-right">Aksi Keputusan</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {data.map((reg: any) => {
                const isLengkap = Object.values(reg.docs).filter(Boolean).length === 4;
                return (
                  <tr key={reg.id} className="hover:bg-slate-50/80 transition-colors">
                    <td className="px-4 py-3 font-medium text-slate-800">{reg.name}</td>
                    <td className="px-4 py-3 text-xs">{isLengkap ? 'Lengkap (4/4)' : 'Belum Lengkap'}</td>
                    <td className="px-4 py-3">
                      <span className={`inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium 
                        ${reg.status === 'Diterima' ? 'bg-emerald-100 text-emerald-700' : 
                          reg.status === 'Cadangan' ? 'bg-blue-100 text-blue-700' :
                          reg.status === 'Tidak Diterima' ? 'bg-rose-100 text-rose-700' : 'bg-slate-100 text-slate-700'}`}>
                        {reg.status}
                      </span>
                    </td>
                    <td className="px-4 py-3 text-right">
                      <div className="inline-block relative min-w-[190px]">
                        <select
                          value={reg.status}
                          onChange={(e) => updateStatus(reg.id, e.target.value)}
                          className={`w-full appearance-none pl-3.5 pr-9 py-2 text-xs font-semibold rounded-xl border transition-all duration-200 cursor-pointer focus:outline-none focus:ring-4 ${
                            reg.status === 'Diterima'
                              ? 'bg-emerald-50/70 border-emerald-200 text-emerald-700 focus:ring-emerald-500/10 focus:border-emerald-500'
                              : reg.status === 'Cadangan'
                              ? 'bg-blue-50/70 border-blue-200 text-blue-700 focus:ring-blue-500/10 focus:border-blue-500'
                              : reg.status === 'Tidak Diterima'
                              ? 'bg-rose-50/70 border-rose-200 text-rose-700 focus:ring-rose-500/10 focus:border-rose-500'
                              : 'bg-slate-50 border-slate-200 text-slate-700 focus:ring-blue-500/10 focus:border-blue-500'
                          }`}
                        >
                          <option value="Menunggu Keputusan">Menunggu Keputusan</option>
                          <option value="Diterima">Diterima</option>
                          <option value="Cadangan">Cadangan</option>
                          <option value="Tidak Diterima">Tidak Diterima</option>
                        </select>
                        <div className="pointer-events-none absolute inset-y-0 right-3 flex items-center text-slate-400">
                          <ChevronDown size={14} />
                        </div>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

function SiswaDiterimaView({ data, settings, setSettings }: any) {
  const diterimaData = data.filter((d: any) => d.status === 'Diterima');
  const kuotaMaks = Number(settings?.kuotaMaksimal) || 0;
  const isKuotaPenuh = kuotaMaks > 0 && diterimaData.length >= kuotaMaks;
  
  const handleExport = () => {
    const cols = [
      'No',
      'Nomor Pendaftaran',
      'NIS Lokal',
      'NISN Nasional',
      'Nama Lengkap',
      'Nama Panggilan',
      'Jenis Kelamin',
      'Tempat Lahir',
      'Tanggal Lahir',
      'Agama',
      'Kewarganegaraan',
      'Anak Ke',
      'Dari (Bersaudara)',
      'Saudara Kandung',
      'Saudara Sebapak',
      'Saudara Seibu',
      'Saudara Tiri',
      'Saudara Angkat',
      'Status dalam Keluarga',
      'Bahasa Sehari-hari',
      'Bertempat Tinggal Pada',
      'Berat Badan (kg)',
      'Tinggi Badan (cm)',
      'Golongan Darah',
      'Penyakit (Kronis)',
      'Alamat Tempat Tinggal',
      'Telepon',
      'Jarak Tempat Tinggal ke Sekolah (km)',
      'Nama Sekolah Asal',
      'Alamat Sekolah Asal',
      'Nama Ayah Kandung',
      'Nama Ibu Kandung',
      'Alamat Orangtua',
      'Pendidikan Tertinggi Ayah',
      'Pendidikan Tertinggi Ibu',
      'Pekerjaan Ayah',
      'Pekerjaan Ibu',
      'Penghasilan per Bulan',
      'Hubungan Keluarga Wali',
      'Pendidikan Tertinggi Wali',
      'Pekerjaan/Jabatan Wali',
      'Alamat Wali',
      'Telepon Wali',
      'Nama Keluarga Terdekat',
      'Alamat Keluarga Terdekat',
      'Hubungan Keluarga Terdekat',
      'Telepon Keluarga Terdekat',
      'Kondisi Anak (Perhatian Khusus)',
      'Prestasi yang Pernah Diraih',
      'Kemampuan Membaca Huruf Latin',
      'Kemampuan Membaca Al-Qur\'an',
      'Layanan Kendaraan Antar Jemput',
      'Saran dan Harapan Orangtua',
      'Fotocopy Akta Kelahiran',
      'Fotocopy Kartu Keluarga',
      'Surat Pernyataan Bermaterai 10.000',
      'Pas Foto Berwarna 3x4 (2 Lembar)',
      'Status Seleksi'
    ];
    const rows = diterimaData.map((d: any, i: number) => [
      i + 1,
      d.regNumber || '-',
      '-',
      '-',
      d.name || '-',
      d.nickname || '-',
      d.gender || '-',
      d.birthPlace || '-',
      d.birthDate || '-',
      d.religion || '-',
      d.nationality || '-',
      d.childOrder || '-',
      d.siblingsCount || '-',
      d.siblingsDetail?.kandung || '0',
      d.siblingsDetail?.sebapak || '0',
      d.siblingsDetail?.seibu || '0',
      d.siblingsDetail?.tiri || '0',
      d.siblingsDetail?.angkat || '0',
      d.familyStatus || '-',
      d.dailyLanguage || '-',
      d.livingWith || '-',
      d.physical?.weight || '-',
      d.physical?.height || '-',
      d.physical?.bloodType || '-',
      d.physical?.chronicDisease || '-',
      d.address || '-',
      d.phone || '-',
      d.distanceToSchool || '-',
      d.previousSchool?.name || '-',
      d.previousSchool?.address || '-',
      d.parent?.fatherName || '-',
      d.parent?.motherName || '-',
      d.parent?.parentAddress || '-',
      d.parent?.fatherEdu || '-',
      d.parent?.motherEdu || '-',
      d.parent?.fatherJob || '-',
      d.parent?.motherJob || '-',
      d.parent?.income || '-',
      d.guardian?.hasGuardian ? (d.guardian?.relation || '-') : '-',
      d.guardian?.hasGuardian ? (d.guardian?.edu || '-') : '-',
      d.guardian?.hasGuardian ? (d.guardian?.job || '-') : '-',
      d.guardian?.hasGuardian ? (d.guardian?.address || '-') : '-',
      d.guardian?.hasGuardian ? (d.guardian?.phone || '-') : '-',
      d.emergencyContact?.name || '-',
      d.emergencyContact?.address || '-',
      d.emergencyContact?.relation || '-',
      d.emergencyContact?.phone || '-',
      d.specialNeeds || '-',
      d.achievements || '-',
      d.latinReading || '-',
      d.quranReading || '-',
      d.transport || '-',
      d.suggestions || '-',
      d.docs?.akta ? 'Ada' : 'Tidak Ada',
      d.docs?.kk ? 'Ada' : 'Tidak Ada',
      d.docs?.materai ? 'Ada' : 'Tidak Ada',
      d.docs?.foto ? 'Ada' : 'Tidak Ada',
      d.status || '-'
    ]);
    exportToExcel('DATA SISWA DITERIMA (TU) SD AL-RASYID', cols, rows, 'Data_Siswa_Diterima');
  };

  return (
    <div className="space-y-4 animate-in fade-in duration-300">
      <div className="flex flex-col lg:flex-row justify-between lg:items-center gap-4">
        <div className="bg-blue-50 border border-blue-100 rounded-xl p-4 flex items-start space-x-3 flex-1">
          <FileArchive className="text-blue-600 mt-0.5 flex-shrink-0" size={20} />
          <div>
            <h4 className="font-semibold text-blue-900">Database Tata Usaha (Siswa Diterima)</h4>
            <p className="text-sm text-blue-700 mt-1">Hanya pendaftar yang dinyatakan "Diterima" yang akan masuk ke daftar ini untuk dilengkapi datanya oleh TU.</p>
          </div>
        </div>

        <Card className="p-4 flex flex-col sm:flex-row sm:items-center gap-4 border-emerald-200 bg-white">
          <div>
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1">
              Jumlah Kuota yang Akan Diterima
            </label>
            <div className="flex items-center space-x-3">
              <input
                type="number"
                min="1"
                value={settings?.kuotaMaksimal ?? ''}
                onChange={(e: any) => setSettings({ ...settings, kuotaMaksimal: e.target.value })}
                placeholder="Contoh: 50"
                className="w-28 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 focus:bg-white transition-all text-sm font-bold text-slate-800"
              />
              <span className="text-xs font-medium text-slate-500 whitespace-nowrap">
                Terisi: <strong className="text-emerald-600">{diterimaData.length}</strong> / {kuotaMaks} Siswa
              </span>
            </div>
          </div>
          <Button variant="outline" onClick={handleExport} className="text-emerald-700 border-emerald-200 hover:bg-emerald-50 whitespace-nowrap">
            <Download size={18} className="mr-2" /> Export Excel
          </Button>
        </Card>
      </div>

      {isKuotaPenuh && (
        <div className="p-4 rounded-2xl bg-amber-50 border border-amber-200 flex items-start space-x-3 text-amber-800 shadow-sm">
          <AlertCircle size={20} className="text-amber-600 flex-shrink-0 mt-0.5" />
          <div className="text-sm">
            <p className="font-bold text-amber-900">Pemberitahuan Sistem: Kuota Siswa Diterima Telah Mencapai Batas ({diterimaData.length} / {kuotaMaks} Siswa)</p>
            <p className="text-amber-700 mt-0.5">
              Jumlah siswa diterima telah mencapai batas kuota yang ditetapkan. Anda dapat mengubah angka pada field <strong>Jumlah Kuota yang Akan Diterima</strong> di atas jika ingin menyesuaikan kapasitas.
            </p>
          </div>
        </div>
      )}

      <Card>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm text-slate-600">
            <thead className="bg-slate-50 text-slate-700 font-semibold border-b border-slate-200">
              <tr>
                <th className="px-4 py-3">No. Pendaftaran</th>
                <th className="px-4 py-3">Nama Siswa</th>
                <th className="px-4 py-3">L/P</th>
                <th className="px-4 py-3">NIS Lokal</th>
                <th className="px-4 py-3">NISN Nasional</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {diterimaData.length > 0 ? diterimaData.map((student: any) => (
                <tr key={student.id} className="hover:bg-slate-50/80 transition-colors">
                  <td className="px-4 py-3 font-medium text-slate-800">{student.regNumber}</td>
                  <td className="px-4 py-3">{student.name}</td>
                  <td className="px-4 py-3">{student.gender === 'Laki-laki' ? 'L' : 'P'}</td>
                  <td className="px-4 py-3 text-slate-400 italic">Belum diisi</td>
                  <td className="px-4 py-3 text-slate-400 italic">Belum diisi</td>
                </tr>
              )) : (
                <tr><td colSpan={5} className="px-4 py-12 text-center text-slate-500">Belum ada siswa diterima.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

function RombelView({ data, rombels, setRombels }: any) {
  const [isAdding, setIsAdding] = useState(false);
  const [newRombel, setNewRombel] = useState({ name: '', capacity: 28, malePct: 50 });
  
  const diterimaData = data.filter((d: any) => d.status === 'Diterima');
  const unassigned = diterimaData.filter((d: any) => !rombels.some((r: any) => r.students?.includes(d.id)));
  const unassignedLaki = unassigned.filter((d: any) => d.gender === 'Laki-laki').length;
  const unassignedPerempuan = unassigned.filter((d: any) => d.gender === 'Perempuan').length;

  const handleAddRombel = (e: any) => {
    e.preventDefault();
    const maleTarget = Math.round((newRombel.malePct / 100) * newRombel.capacity);
    const femaleTarget = newRombel.capacity - maleTarget;
    setRombels([...rombels, { ...newRombel, id: Date.now(), students: [], maleTarget, femaleTarget }]);
    setIsAdding(false);
    setNewRombel({ name: '', capacity: 28, malePct: 50 });
  };

  const handleBagiOtomatis = () => {
    // Siapkan antrean siswa yang belum dapat kelas
    let poolLaki = [...unassigned.filter((d: any) => d.gender === 'Laki-laki')];
    let poolPerempuan = [...unassigned.filter((d: any) => d.gender === 'Perempuan')];
    
    if (poolLaki.length === 0 && poolPerempuan.length === 0) return;

    // Lakukan pembagian berdasarkan sisa target tiap rombel
    const updatedRombels = rombels.map((rombel: any) => {
      const currentStudents = rombel.students.map((id: any) => data.find((d: any) => d.id === id)).filter(Boolean);
      const currentLakiCount = currentStudents.filter((s: any) => s.gender === 'Laki-laki').length;
      const currentPerempuanCount = currentStudents.filter((s: any) => s.gender === 'Perempuan').length;

      let neededLaki = Math.max(0, rombel.maleTarget - currentLakiCount);
      let neededPerempuan = Math.max(0, rombel.femaleTarget - currentPerempuanCount);

      const newStudents = [...rombel.students];

      // Masukkan siswa laki-laki sampai target kelas terpenuhi (atau antrean habis)
      while(neededLaki > 0 && poolLaki.length > 0) {
        const student = poolLaki.pop();
        newStudents.push(student.id);
        neededLaki--;
      }

      // Masukkan siswa perempuan sampai target kelas terpenuhi (atau antrean habis)
      while(neededPerempuan > 0 && poolPerempuan.length > 0) {
        const student = poolPerempuan.pop();
        newStudents.push(student.id);
        neededPerempuan--;
      }

      return { ...rombel, students: newStudents };
    });

    setRombels(updatedRombels);
  };

  const handleExportRombel = () => {
    if (rombels.length === 0) return;

    const escapeXml = (val: any) =>
      String(val ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');

    const usedNames = new Set<string>();
    const worksheetsXml = rombels.map((rombel: any, idx: number) => {
      let baseName = String(rombel.name || `Kelas ${idx + 1}`).replace(/[\\/?*:[\]]/g, ' ').trim().slice(0, 28) || `Kelas ${idx + 1}`;
      let sheetName = baseName;
      let counter = 1;
      while (usedNames.has(sheetName.toLowerCase())) {
        sheetName = `${baseName} (${counter++})`;
      }
      usedNames.add(sheetName.toLowerCase());

      const stds = (rombel.students || [])
        .map((id: any) => data.find((d: any) => d.id === id))
        .filter(Boolean)
        .sort((a: any, b: any) => String(a.name || '').localeCompare(String(b.name || ''), 'id'));

      const totalRows = Math.max(Number(rombel.capacity) || 25, stds.length);
      let rowsXml = '';

      for (let i = 0; i < totalRows; i++) {
        const s = stds[i];
        const noInduk = s ? (s.nisLokal || '') : '';
        const nisn = s ? (s.nisn || '') : '';
        const nama = s ? (s.name || '') : '';
        const lp = s ? (s.gender === 'Laki-laki' ? 'L' : s.gender === 'Perempuan' ? 'P' : '') : '';

        rowsXml += `
        <Row ss:Height="20">
          <Cell ss:StyleID="sCenter"><Data ss:Type="Number">${i + 1}</Data></Cell>
          <Cell ss:StyleID="sCenter"><Data ss:Type="String">${escapeXml(noInduk)}</Data></Cell>
          <Cell ss:StyleID="sCenter"><Data ss:Type="String">${escapeXml(nisn)}</Data></Cell>
          <Cell ss:StyleID="sLeft"><Data ss:Type="String">${escapeXml(nama)}</Data></Cell>
          <Cell ss:StyleID="sCenter"><Data ss:Type="String">${escapeXml(lp)}</Data></Cell>
        </Row>`;
      }

      return `
      <Worksheet ss:Name="${escapeXml(sheetName)}">
        <Table>
          <Column ss:Width="35"/>
          <Column ss:Width="80"/>
          <Column ss:Width="80"/>
          <Column ss:Width="210"/>
          <Column ss:Width="50"/>
          <Row ss:Height="28">
            <Cell ss:StyleID="sHeader"><Data ss:Type="String">No</Data></Cell>
            <Cell ss:StyleID="sHeader"><Data ss:Type="String">No. Induk</Data></Cell>
            <Cell ss:StyleID="sHeader"><Data ss:Type="String">NISN</Data></Cell>
            <Cell ss:StyleID="sHeader"><Data ss:Type="String">Nama</Data></Cell>
            <Cell ss:StyleID="sHeader"><Data ss:Type="String">L/P</Data></Cell>
          </Row>${rowsXml}
        </Table>
      </Worksheet>`;
    }).join('');

    const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
  <Styles>
    <Style ss:ID="sHeader">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
      </Borders>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#000000"/>
    </Style>
    <Style ss:ID="sCenter">
      <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
      </Borders>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Color="#000000"/>
    </Style>
    <Style ss:ID="sLeft">
      <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#000000"/>
      </Borders>
      <Font ss:FontName="Calibri" ss:Size="11" ss:Color="#000000"/>
    </Style>
  </Styles>${worksheetsXml}
</Workbook>`;

    const blob = new Blob([workbookXml], { type: 'application/vnd.ms-excel;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Data_Pembagian_Rombel.xls';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const targetLaki = Math.round((newRombel.malePct / 100) * newRombel.capacity);
  const targetPerempuan = newRombel.capacity - targetLaki;

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
        <div className="bg-slate-100 p-4 rounded-xl border border-slate-200 text-center">
          <div className="text-sm text-slate-500 font-medium">Sisa Siswa Belum Dapat Kelas</div>
          <div className="text-2xl font-bold text-slate-800">{unassigned.length}</div>
        </div>
        <div className="bg-blue-50 p-4 rounded-xl border border-blue-100 text-center">
          <div className="text-sm text-blue-600 font-medium">Laki-laki (Belum Dapat Kelas)</div>
          <div className="text-2xl font-bold text-blue-800">{unassignedLaki}</div>
        </div>
        <div className="bg-rose-50 p-4 rounded-xl border border-rose-100 text-center">
          <div className="text-sm text-rose-600 font-medium">Perempuan (Belum Dapat Kelas)</div>
          <div className="text-2xl font-bold text-rose-800">{unassignedPerempuan}</div>
        </div>
      </div>

      <div className="flex flex-col md:flex-row justify-between md:items-center gap-4 mb-4">
        <h3 className="text-lg font-bold text-slate-800">Daftar Rombongan Belajar (Kelas)</h3>
        <div className="flex flex-wrap gap-3 w-full md:w-auto">
          <Button 
            variant="outline" 
            onClick={handleExportRombel} 
            disabled={rombels.length === 0} 
            className="flex-1 md:flex-none text-emerald-700 border-emerald-200 hover:bg-emerald-50"
          >
            <Download size={18} className="mr-2"/> Export Excel
          </Button>
          <Button 
            variant="outline" 
            onClick={handleBagiOtomatis} 
            disabled={unassigned.length === 0 || rombels.length === 0} 
            className="flex-1 md:flex-none text-blue-700 border-blue-200 hover:bg-blue-50"
          >
            <Users size={18} className="mr-2"/> Bagi Otomatis
          </Button>
          <Button onClick={() => setIsAdding(true)} className="flex-1 md:flex-none">
            <Plus size={18} className="mr-2"/> Tambah Rombel
          </Button>
        </div>
      </div>

      {isAdding && (
        <Card className="p-6 bg-slate-50">
          <form onSubmit={handleAddRombel} className="space-y-4">
            <h4 className="font-semibold text-slate-800 border-b pb-2">Buat Rombel Baru</h4>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <Input label="Nama Rombel (Kelas)" value={newRombel.name} onChange={(e: any) => setNewRombel({...newRombel, name: e.target.value})} required placeholder="Contoh: 1 Abu Bakar" />
              <Input label="Total Kapasitas" type="number" value={newRombel.capacity} onChange={(e: any) => setNewRombel({...newRombel, capacity: Number(e.target.value)})} required />
              
              <div className="flex flex-col space-y-2">
                <label className="text-sm font-semibold text-slate-700">Rasio Gender: {newRombel.malePct}% Laki-laki</label>
                <input type="range" min="0" max="100" step="5" value={newRombel.malePct} onChange={(e: any) => setNewRombel({...newRombel, malePct: Number(e.target.value)})} className="w-full accent-blue-600" />
                <div className="flex justify-between text-xs text-slate-500 font-medium">
                  <span className="text-blue-600">Target: {targetLaki} L</span>
                  <span className="text-rose-600">Target: {targetPerempuan} P</span>
                </div>
              </div>
            </div>
            <div className="flex justify-end space-x-3 pt-4 border-t border-slate-200">
              <Button type="button" variant="ghost" onClick={() => setIsAdding(false)}>Batal</Button>
              <Button type="submit">Simpan Rombel</Button>
            </div>
          </form>
        </Card>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        {rombels.length === 0 ? (
          <div className="col-span-full p-8 text-center text-slate-500 bg-white border border-slate-200 rounded-xl border-dashed">Belum ada rombel yang dibuat.</div>
        ) : rombels.map((rombel: any) => {
          const stds = rombel.students.map((id: any) => data.find((d: any) => d.id === id)).filter(Boolean);
          const L = stds.filter((s: any) => s.gender === 'Laki-laki').length;
          const P = stds.filter((s: any) => s.gender === 'Perempuan').length;
          
          return (
            <Card key={rombel.id} className="p-5">
              <div className="flex justify-between items-center mb-3">
                <h4 className="font-bold text-lg text-slate-800">{rombel.name}</h4>
                <div className="flex items-center space-x-2">
                  <span className="text-xs bg-slate-100 text-slate-600 px-2 py-1 rounded-md font-medium border border-slate-200">Kapasitas: {rombel.capacity}</span>
                  <button
                    onClick={() => setRombels(rombels.filter((r: any) => String(r.id) !== String(rombel.id)))}
                    className="p-1.5 text-rose-600 hover:bg-rose-50 rounded-md transition-colors"
                    title="Hapus Rombel"
                  >
                    <Trash2 size={16} />
                  </button>
                </div>
              </div>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between p-2 bg-slate-50 rounded">
                  <span className="text-slate-600">Siswa Terisi</span>
                  <span className="font-semibold">{rombel.students.length} / {rombel.capacity}</span>
                </div>
                <div className="flex justify-between p-2 bg-blue-50/50 rounded">
                  <span className="text-blue-700">Laki-laki (Terisi / Target)</span>
                  <span className="font-semibold text-blue-700">{L} / {rombel.maleTarget}</span>
                </div>
                <div className="flex justify-between p-2 bg-rose-50/50 rounded">
                  <span className="text-rose-700">Perempuan (Terisi / Target)</span>
                  <span className="font-semibold text-rose-700">{P} / {rombel.femaleTarget}</span>
                </div>
              </div>
            </Card>
          );
        })}
      </div>
    </div>
  );
}
