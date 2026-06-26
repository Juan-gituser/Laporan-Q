import React, { useState, useEffect, useRef } from 'react';
import './App.css';
import { PlusCircle, Printer, LogOut, Loader2, Trash2, X, CheckCircle } from 'lucide-react';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, HeadingLevel, AlignmentType, ImageRun } from 'docx';
import { initializeApp } from 'firebase/app';
import { getAuth, signInAnonymously, onAuthStateChanged, signOut } from 'firebase/auth';
import { getFirestore, collection, query, onSnapshot, orderBy, addDoc, deleteDoc, doc, setDoc, getDoc } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: 'AIzaSyDt2Ur0oVFbnooONeq3DMNqIMr1qXqcSsg',
  authDomain: 'laporan-q.firebaseapp.com',
  projectId: 'laporan-q',
  storageBucket: 'laporan-q.firebasestorage.app',
  messagingSenderId: '949459194513',
  appId: '1:949459194513:web:e4fec30cb78e27a72a6530',
  measurementId: 'G-3FV5SCG8Y1',
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const projectId = firebaseConfig.projectId;

const defaultFormData = {
  judulLaporan: '',
  jenis: 'WFH',
  tanggal: new Date().toISOString().split('T')[0],
  nama: '',
  nip: '',
  unit: '',
  organisasi: 'Badan Karantina Indonesia',
};

export default function App() {
  const [user, setUser] = useState(null);
  const [reports, setReports] = useState([]);
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [loading, setLoading] = useState(true);
  const [inputNip, setInputNip] = useState('');
  const [needsProfile, setNeedsProfile] = useState(false);
  const [formData, setFormData] = useState(defaultFormData);
  const [tugasList, setTugasList] = useState([{ tugas: '', hasil: '', documentation: null }]);
  const [signatureDataUrl, setSignatureDataUrl] = useState('');
  const [formError, setFormError] = useState('');
  const [authError, setAuthError] = useState('');
  const [authInitializing, setAuthInitializing] = useState(true);
  const [authReady, setAuthReady] = useState(false);
  const [offlineMode, setOfflineMode] = useState(false);
  const [isOnline, setIsOnline] = useState(typeof navigator !== 'undefined' ? navigator.onLine : true);
  const [syncStatus, setSyncStatus] = useState('idle');
  const [syncError, setSyncError] = useState('');
  const [localReports, setLocalReports] = useState([]);
  const [previewReport, setPreviewReport] = useState(null);
  const isNipValid = /^[0-9]{18}$/.test(inputNip);
  const showOfflineFallback =
    !offlineMode &&
    (authError || /firestore|izin|permission|permissions|offline/i.test(formError));
  const canvasRef = useRef(null);
  const drawing = useRef(false);

  useEffect(() => {
    if (typeof window !== 'undefined') {
      window.__appState = {
        isLoggedIn,
        needsProfile,
        offlineMode,
        authReady,
        isOnline,
        inputNip,
        formData,
      };
    }
  });

  useEffect(() => {
    setAuthInitializing(true);
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      if (currentUser) {
        setAuthReady(true);
      }
      setLoading(false);
      setAuthInitializing(false);
    });

    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);

    if (typeof window !== 'undefined') {
      window.addEventListener('online', handleOnline);
      window.addEventListener('offline', handleOffline);
    }

    signInAnonymously(auth)
      .then(() => setAuthReady(true))
      .catch((error) => {
        console.error('Firebase auth error:', error);
        setAuthError(
          'Gagal terhubung ke Firebase Auth. Periksa konfigurasi Firebase, pastikan project ID benar, dan aktifkan Anonymous Authentication di Firebase Console.'
        );
        setOfflineMode(true);
      })
      .finally(() => {
        setLoading(false);
        setAuthInitializing(false);
      });

    return () => {
      unsubscribe();
      if (typeof window !== 'undefined') {
        window.removeEventListener('online', handleOnline);
        window.removeEventListener('offline', handleOffline);
      }
    };
  }, []);

  useEffect(() => {
    if (!isLoggedIn || !user || offlineMode) return;

    const reportsQuery = query(
      collection(db, 'artifacts', projectId, 'users', user.uid, 'reports'),
      orderBy('createdAt', 'desc'),
    );

    const unsubscribe = onSnapshot(
      reportsQuery,
      (snapshot) => {
        const data = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        setReports(dedupeReports(data));
      },
      (error) => {
        console.error('Firestore snapshot error:', error);
      },
    );

    return unsubscribe;
  }, [isLoggedIn, user]);

  const getCurrentUser = () => (offlineMode ? { uid: 'offline' } : auth.currentUser || user);

  const getLocalProfile = (nip) => {
    if (typeof window === 'undefined') return null;
    const stored = window.localStorage.getItem(`laporan-q-profile-${nip}`);
    return stored ? JSON.parse(stored) : null;
  };

  const saveLocalProfile = (profile) => {
    if (typeof window === 'undefined') return;
    window.localStorage.setItem(`laporan-q-profile-${profile.nip}`, JSON.stringify(profile));
  };

  const generateLocalId = () => {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) {
      return crypto.randomUUID();
    }
    return `local-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
  };

  const loadLocalReports = (nip) => {
    if (typeof window === 'undefined') return [];
    const stored = window.localStorage.getItem(`laporan-q-reports-${nip}`);
    const reports = stored ? JSON.parse(stored) : [];
    return reports.map((report) => ({
      ...report,
      id: report.id || generateLocalId(),
    }));
  };

  const dedupeReports = (items) => {
    const seen = new Set();
    return items.filter((item) => {
      if (!item?.id) return true;
      if (seen.has(item.id)) return false;
      seen.add(item.id);
      return true;
    });
  };

  const saveLocalReports = (nip, reportsToSave) => {
    if (typeof window === 'undefined') return;
    window.localStorage.setItem(`laporan-q-reports-${nip}`, JSON.stringify(reportsToSave));
  };

  const getAllLocalProfiles = () => {
    const profiles = [];
    if (typeof window === 'undefined') return profiles;

    for (let i = 0; i < window.localStorage.length; i += 1) {
      const key = window.localStorage.key(i);
      if (!key || !key.startsWith('laporan-q-profile-')) continue;
      const item = window.localStorage.getItem(key);
      if (!item) continue;
      try {
        const profile = JSON.parse(item);
        profiles.push(profile);
      } catch (error) {
        console.warn('Invalid local profile data', key, error);
      }
    }

    return profiles;
  };

  const getAllLocalReports = () => {
    const reports = [];
    if (typeof window === 'undefined') return reports;

    for (let i = 0; i < window.localStorage.length; i += 1) {
      const key = window.localStorage.key(i);
      if (!key || !key.startsWith('laporan-q-reports-')) continue;
      const item = window.localStorage.getItem(key);
      if (!item) continue;
      try {
        const storedReports = JSON.parse(item);
        const nip = key.replace('laporan-q-reports-', '');
        storedReports.forEach((report) => {
          reports.push({
            ...report,
            id: report.id || generateLocalId(),
            nip,
          });
        });
      } catch (error) {
        console.warn('Invalid local reports data', key, error);
      }
    }

    return reports;
  };

  const handleLogin = async (event) => {
    event.preventDefault();
    if (typeof window !== 'undefined') {
      window.__debugLogin = window.__debugLogin || [];
      window.__debugLogin.push({ step: 'start', authReady, offlineMode, inputNip, userId: user?.uid || null });
    }
    console.log('handleLogin start', { authReady, offlineMode, inputNip, user });
    setFormError('');

    if (!offlineMode && !authReady) {
      if (typeof window !== 'undefined') {
        window.__debugLogin.push({ step: 'waiting-auth' });
      }
      setFormError('Sedang memulai autentikasi. Tunggu beberapa saat lalu coba lagi.');
      return;
    }

    if (!/^[0-9]{18}$/.test(inputNip.trim())) {
      setFormError('NIP harus terdiri dari tepat 18 angka.');
      return;
    }

    if (offlineMode) {
      if (typeof window !== 'undefined') {
        window.__debugLogin.push({ step: 'offline-login' });
      }
      const profile = getLocalProfile(inputNip.trim());
      if (profile) {
        setFormData({ ...defaultFormData, ...profile });
        setReports(loadLocalReports(inputNip.trim()));
        setIsLoggedIn(true);
        setNeedsProfile(false);
      } else {
        setFormData({ ...defaultFormData, nip: inputNip.trim() });
        setNeedsProfile(true);
      }
      return;
    }

    const currentUser = getCurrentUser();
    window.__handleLoginRan = true;
    window.__handleLoginCurrentUser = currentUser?.uid || null;
    if (!currentUser) {
      setFormError('Autentikasi Firebase belum siap. Silakan muat ulang halaman.');
      window.__handleLoginError = 'no-user';
      return;
    }

    try {
      const profileRef = doc(db, 'artifacts', projectId, 'users', currentUser.uid, 'profile', inputNip.trim());
      const profileSnap = await getDoc(profileRef);
      window.__handleLoginProfileExists = profileSnap.exists();

      if (profileSnap.exists()) {
        if (typeof window !== 'undefined') {
          window.__debugLogin.push({ step: 'profile-exists' });
        }
        setFormData({ ...defaultFormData, ...profileSnap.data() });
        setIsLoggedIn(true);
        setNeedsProfile(false);
      } else {
        if (typeof window !== 'undefined') {
          window.__debugLogin.push({ step: 'profile-missing' });
        }
        setFormData({ ...defaultFormData, nip: inputNip.trim() });
        setNeedsProfile(true);
      }
    } catch (error) {
      console.error('handleLogin getDoc error:', error);
      if (typeof window !== 'undefined') {
        window.__handleLoginLastFirebaseError = {
          message: error?.message,
          code: error?.code,
        };
        window.__debugLogin.push({
          step: 'getDoc-catch',
          message: error?.message,
          code: error?.code,
          offlineError: error?.message?.toLowerCase().includes('client is offline') ||
            error?.message?.toLowerCase().includes('offline') ||
            error?.code?.includes('unavailable'),
          networkOffline: typeof navigator !== 'undefined' ? !navigator.onLine : false,
        });
      }

      const offlineError =
        error?.message?.toLowerCase().includes('client is offline') ||
        error?.message?.toLowerCase().includes('offline') ||
        error?.code?.includes('unavailable');
      const networkOffline = typeof navigator !== 'undefined' ? !navigator.onLine : false;
      const shouldUseOffline = offlineError || networkOffline;

      if (shouldUseOffline) {
        if (typeof window !== 'undefined') {
          window.__debugLogin.push({ step: 'enter-offline-mode', offlineError, networkOffline });
        }
        setOfflineMode(true);
        const profile = getLocalProfile(inputNip.trim());
        if (profile) {
          setFormData({ ...defaultFormData, ...profile });
          setReports(loadLocalReports(inputNip.trim()));
          setIsLoggedIn(true);
          setNeedsProfile(false);
          setFormError('Mode offline aktif. Data akan disimpan secara lokal.');
        } else {
          setFormData({ ...defaultFormData, nip: inputNip.trim() });
          setNeedsProfile(true);
          setFormError('Mode offline aktif. Isi profil untuk menyimpan laporan lokal.');
        }
        return;
      }

      const permissionMessage =
        error?.code === 'permission-denied' ||
        error?.code === 'unauthenticated' ||
        error?.code === 'missing-permission';

      if (permissionMessage) {
        setOfflineMode(true);
        setFormError('Akses Firebase ditolak oleh aturan Firestore. Aplikasi akan beralih ke mode offline.');
        const profile = getLocalProfile(inputNip.trim());
        if (profile) {
          setFormData({ ...defaultFormData, ...profile });
          setReports(loadLocalReports(inputNip.trim()));
          setIsLoggedIn(true);
          setNeedsProfile(false);
        } else {
          setFormData({ ...defaultFormData, nip: inputNip.trim() });
          setNeedsProfile(true);
        }
        return;
      }

      setFormError('Gagal memeriksa profil. Periksa koneksi atau izin Firestore.');
    }
  };

  const syncLocalForProfile = async (profile) => {
    const currentUser = getCurrentUser();
    if (!currentUser) return;

    try {
      const profileRef = doc(db, 'artifacts', projectId, 'users', currentUser.uid, 'profile', profile.nip);
      await setDoc(profileRef, { ...profile, updatedAt: Date.now() });
    } catch (error) {
      console.error('syncLocalForProfile error:', error);
      throw error;
    }
  };

  const syncLocalReportsToFirebase = async (nip, reportsToSync) => {
    const currentUser = getCurrentUser();
    if (!currentUser) return;

    try {
      for (const report of reportsToSync) {
        const reportCopy = { ...report };
        delete reportCopy.id;
        await addDoc(collection(db, 'artifacts', projectId, 'users', currentUser.uid, 'reports'), reportCopy);
      }
    } catch (error) {
      console.error('syncLocalReportsToFirebase error:', error);
      throw error;
    }
  };

  const handleRegister = async (event) => {
    event.preventDefault();
    setFormError('');

    if (!formData.nama.trim() || !formData.unit.trim()) {
      setFormError('Nama dan unit kerja harus diisi.');
      return;
    }

    if (offlineMode) {
      saveLocalProfile(formData);
      setIsLoggedIn(true);
      setNeedsProfile(false);
      setFormError('');
      return;
    }

    const currentUser = getCurrentUser();
    if (!currentUser) {
      setFormError('Autentikasi Firebase belum siap. Silakan muat ulang halaman.');
      return;
    }

    const profileRef = doc(db, 'artifacts', projectId, 'users', currentUser.uid, 'profile', formData.nip);
    await setDoc(profileRef, { ...formData, updatedAt: Date.now() });
    setIsLoggedIn(true);
    setNeedsProfile(false);
  };

  const fileToDataUrl = (file) =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error('Gagal membaca file'));
      reader.readAsDataURL(file);
    });

  const handleTaskDocumentationChange = async (index, event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      setFormError('Dokumentasi harus berupa gambar (foto atau screenshot).');
      return;
    }

    try {
      const documentation = {
        id: generateLocalId(),
        name: file.name,
        type: file.type,
        dataUrl: await fileToDataUrl(file),
      };

      const nextTasks = [...tugasList];
      nextTasks[index] = { ...nextTasks[index], documentation };
      setTugasList(nextTasks);
      setFormError('');
    } catch (error) {
      console.error('handleTaskDocumentationChange error:', error);
      setFormError('Gagal membaca file dokumentasi. Coba lagi.');
    }
  };

  const dataUriToUint8Array = (dataUri) => {
    const base64 = dataUri.split(',')[1];
    const raw = atob(base64);
    const uint8Array = new Uint8Array(raw.length);
    for (let i = 0; i < raw.length; i += 1) {
      uint8Array[i] = raw.charCodeAt(i);
    }
    return uint8Array;
  };

  const formatIndonesianDate = (value) => {
    const date = new Date(value);
    return date.toLocaleDateString('id-ID', {
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric',
    });
  };

  const buildReportDocument = (report) => {
    const createTextRun = (text, options = {}) =>
      new TextRun({
        text,
        font: 'Arial',
        size: 24,
        ...options,
      });

    const rows = [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({ width: { size: 8, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [createTextRun('No', { bold: true })] })] }),
          new TableCell({ width: { size: 35, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [createTextRun('Uraian Kegiatan', { bold: true })] })] }),
          new TableCell({ width: { size: 27, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [createTextRun('Hasil', { bold: true })] })] }),
          new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [createTextRun('Dokumentasi', { bold: true })] })] }),
        ],
      }),
      ...report.items.map((item, index) => {
        const documentationChildren = [];
        if (item.documentation?.name) {
          documentationChildren.push(new Paragraph({ children: [createTextRun(item.documentation.name || '')] }));
        }
        if (item.documentation?.dataUrl) {
          const documentationData = dataUriToUint8Array(item.documentation.dataUrl);
          documentationChildren.push(
            new Paragraph({
              children: [
                new ImageRun({
                  data: documentationData,
                  transformation: { width: 250, height: 120 },
                }),
              ],
            }),
          );
        }

        return new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [createTextRun(String(index + 1))] })] }),
            new TableCell({ children: [new Paragraph({ children: [createTextRun(item.tugas || '')] })] }),
            new TableCell({ children: [new Paragraph({ children: [createTextRun(item.hasil || '')] })] }),
            new TableCell({ children: documentationChildren.length ? documentationChildren : [new Paragraph({ children: [createTextRun('')] })] }),
          ],
        });
      }),
    ];

    const children = [
      new Paragraph({
        children: [createTextRun('LAPORAN PELAKSANAAN', { bold: true, size: 28 })],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        children: [createTextRun(`WORK FROM HOME (${report.jenis || 'WFH'})`, { bold: true, size: 28 })],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({ children: [createTextRun('BIRO ORGANISASI DAN SUMBER DAYA MANUSIA')] , alignment: AlignmentType.CENTER }),
      new Paragraph({ text: ' ', spacing: { after: 200 } }),
      new Paragraph({
        children: [
          createTextRun('Nama', { bold: true }),
          createTextRun(` : ${report.nama || ''}`),
        ],
      }),
      new Paragraph({
        children: [
          createTextRun('NIP', { bold: true }),
          createTextRun(` : ${report.nip || ''}`),
        ],
      }),
      new Paragraph({
        children: [
          createTextRun('Hari/Tanggal', { bold: true }),
          createTextRun(` : ${formatIndonesianDate(report.tanggal || report.createdAt)}`),
        ],
      }),
      new Paragraph({ text: ' ' }),
      new Table({
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        rows,
      }),
      new Paragraph({ text: ' ' }),
    ];

    if (report.signature) {
      const imageData = dataUriToUint8Array(report.signature);
      children.push(
        new Paragraph({
          children: [createTextRun(`Jakarta, ${formatIndonesianDate(report.tanggal || report.createdAt)}`)],
          alignment: AlignmentType.RIGHT,
        }),
      );
      children.push(new Paragraph({ children: [createTextRun('Yang Melaksanakan,')] , alignment: AlignmentType.RIGHT }));
      children.push(new Paragraph({ text: ' ' }));
      children.push(
        new Paragraph({
          children: [
            new ImageRun({
              data: imageData,
              transformation: { width: 200, height: 90 },
            }),
          ],
          alignment: AlignmentType.RIGHT,
        }),
      );
      children.push(new Paragraph({ text: ' ' }));
    }

    children.push(
      new Paragraph({
        children: [createTextRun(report.nama || '')],
        alignment: AlignmentType.RIGHT,
      }),
    );

    return new Document({ sections: [{ properties: {}, children }] });
  };

  const handleExportWord = async (report = null) => {
    const exportReport = report || {
      ...formData,
      items: tugasList,
      signature: signatureDataUrl,
      createdAt: Date.now(),
    };

    try {
      const doc = buildReportDocument(exportReport);
      const blob = await Packer.toBlob(doc);
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `Laporan-${exportReport.nip || 'data'}-${exportReport.tanggal || new Date().toISOString().split('T')[0]}.docx`;
      anchor.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Export Word error:', error);
      setFormError('Gagal membuat file Word. Coba lagi.');
    }
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    setFormError('');

    const currentUser = getCurrentUser();
    if (!offlineMode && !currentUser) {
      setFormError('Autentikasi Firebase belum siap. Silakan muat ulang halaman.');
      return;
    }

    if (tugasList.some((task) => !task.tugas.trim() || !task.hasil.trim())) {
      setFormError('Semua tugas dan hasil harus diisi.');
      return;
    }

    if (tugasList.some((task) => !task.documentation || !task.documentation.dataUrl)) {
      setFormError('Dokumentasi berupa foto atau screenshot diperlukan untuk setiap tugas.');
      return;
    }

    if (!signatureDataUrl) {
      setFormError('Tanda tangan digital harus diisi sebelum menyimpan laporan.');
      return;
    }

    const reportData = {
      ...formData,
      items: tugasList,
      signature: signatureDataUrl,
      createdAt: Date.now(),
    };

    if (offlineMode) {
      const reportWithId = { ...reportData, id: generateLocalId() };
      const existing = loadLocalReports(formData.nip);
      const updated = [reportWithId, ...existing];
      setLocalReports(updated);
      setReports(updated);
      saveLocalReports(formData.nip, updated);
      setTugasList([{ tugas: '', hasil: '', documentation: null }]);
      clearSignature();
      return;
    }

    try {
      const ref = await addDoc(collection(db, 'artifacts', projectId, 'users', currentUser.uid, 'reports'), reportData);
      const savedReport = { id: ref.id, ...reportData };
      setReports((current) => dedupeReports([savedReport, ...current]));
      setTugasList([{ tugas: '', hasil: '', documentation: null }]);
      clearSignature();
    } catch (error) {
      console.error('handleSubmit addDoc error:', error);
      setFormError('Gagal menyimpan laporan. Periksa koneksi atau izin Firebase.');
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
    } catch (error) {
      console.error('Logout error:', error);
    }

    setIsLoggedIn(false);
    setNeedsProfile(false);
    setReports([]);
    setFormData(defaultFormData);
    setTugasList([{ tugas: '', hasil: '', documentation: null }]);
    setSignatureDataUrl('');
    setInputNip('');
  };

  const updateTask = (index, field, value) => {
    const next = [...tugasList];
    next[index][field] = value;
    setTugasList(next);
  };

  const addTask = () => setTugasList([...tugasList, { tugas: '', hasil: '', documentation: null }]);

  const removeTask = (index) => {
    if (tugasList.length === 1) return;
    setTugasList(tugasList.filter((_, i) => i !== index));
  };

  const drawStart = (event) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    ctx.strokeStyle = '#1d4ed8';
    ctx.lineWidth = 3;
    ctx.lineCap = 'round';
    ctx.beginPath();
    ctx.moveTo(x, y);
    drawing.current = true;
    canvas.setPointerCapture(event.pointerId);
  };

  const drawMove = (event) => {
    if (!drawing.current || !canvasRef.current) return;
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    ctx.lineTo(x, y);
    ctx.stroke();
  };

  const drawEnd = (event) => {
    if (!drawing.current || !canvasRef.current) return;
    drawing.current = false;
    const canvas = canvasRef.current;
    canvas.releasePointerCapture(event.pointerId);
    setSignatureDataUrl(canvas.toDataURL());
  };

  const clearSignature = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setSignatureDataUrl('');
  };

  const handleSync = async () => {
    if (offlineMode) {
      setSyncError('Matikan mode offline untuk melakukan sinkronisasi.');
      return;
    }

    const currentUser = getCurrentUser();
    if (!currentUser) {
      setSyncError('Autentikasi Firebase belum siap.');
      return;
    }

    setSyncStatus('running');
    setSyncError('');

    try {
      const profiles = getAllLocalProfiles();
      const localReportsToSync = getAllLocalReports();

      for (const profile of profiles) {
        await syncLocalForProfile(profile);
      }

      const groupedByNip = localReportsToSync.reduce((acc, report) => {
        if (!acc[report.nip]) acc[report.nip] = [];
        acc[report.nip].push(report);
        return acc;
      }, {});

      for (const nip of Object.keys(groupedByNip)) {
        await syncLocalReportsToFirebase(nip, groupedByNip[nip]);
      }

      setSyncStatus('completed');
      setSyncError('');
    } catch (error) {
      setSyncStatus('failed');
      setSyncError('Sinkronisasi gagal. Periksa koneksi dan coba lagi.');
    } finally {
      setTimeout(() => setSyncStatus('idle'), 2500);
    }
  };

  const deleteReport = async (reportId) => {
    const currentUser = getCurrentUser();
    if (!currentUser) return;

    const confirmed = window.confirm('Hapus laporan ini secara permanen?');
    if (!confirmed) return;

    if (offlineMode) {
      const updated = localReports.filter((report) => report.id !== reportId);
      setLocalReports(updated);
      setReports(updated);
      saveLocalReports(formData.nip, updated);
      return;
    }

    await deleteDoc(doc(db, 'artifacts', projectId, 'users', currentUser.uid, 'reports', reportId));
  };

  const openReportPreview = (report) => {
    setPreviewReport(report);
  };

  const closeReportPreview = () => {
    setPreviewReport(null);
  };

  if (loading || authInitializing) {
    return (
      <div className="app-loading">
        <Loader2 className="loader" />
        <p>Memuat aplikasi...</p>
      </div>
    );
  }


  if (!isLoggedIn) {
    return (
      <div className="app-shell app-login">
        <div className="app-card form-card">
          <div className="form-header login-header-centered">
            <div>
              <h1>Masuk ke LAPORAN‑Q</h1>
              <p className="login-subtitle">Jangan lupa buat laporan biar gak kena omel atasan</p>
            </div>
          </div>

          {authError && (
            <div className="error-message">
              {authError} Anda dapat terus menggunakan aplikasi tanpa Firebase jika perlu.
            </div>
          )}
          {showOfflineFallback && !offlineMode && (
            <button
              type="button"
              className="button secondary-button"
              onClick={() => setOfflineMode(true)}
            >
              Gunakan mode offline
            </button>
          )}
          {formError && <div className="error-message">{formError}</div>}
          {syncError && <div className="error-message">{syncError}</div>}
          {offlineMode && (
            <div className="info-message">
              Mode offline aktif. Profil dan laporan akan tersimpan lokal.
            </div>
          )}

          {!needsProfile ? (
            <form onSubmit={handleLogin} className="form-grid login-grid">
              <input
                type="text"
                className="input"
                placeholder="NIP (18 angka)"
                inputMode="numeric"
                pattern="[0-9]*"
                maxLength={18}
                value={inputNip}
                onChange={(e) => {
                  const digits = e.target.value.replace(/\D/g, '').slice(0, 18);
                  setInputNip(digits);
                }}
              />
              <div className="login-actions">
                <button type="submit" className="button primary-button" disabled={!isNipValid}>
                  Lanjutkan
                </button>
              </div>
              <div className="login-note">
                Hanya angka, tepat 18 digit.
              </div>
            </form>
          ) : (
            <form onSubmit={handleRegister} className="form-grid">
              <input
                type="text"
                className="input"
                placeholder="Nama Lengkap"
                value={formData.nama}
                onChange={(e) => setFormData({ ...formData, nama: e.target.value })}
              />
              <input
                type="text"
                className="input"
                placeholder="Unit Kerja"
                value={formData.unit}
                onChange={(e) => setFormData({ ...formData, unit: e.target.value })}
              />
              <div className="register-actions">
                <button type="button" onClick={() => setNeedsProfile(false)} className="button secondary-button">
                  Batal
                </button>
                <button type="submit" className="button primary-button">
                  Simpan Profil
                </button>
              </div>
            </form>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="app-shell">
      <div className="app-container">
        <header className="app-header no-print">
          <div>
            <p className="subtitle">Selamat datang di</p>
            <h1 className="brand">LAPORAN-Q</h1>
          </div>

          <div className="header-actions">
            {!offlineMode && isOnline && (
              <button type="button" className="button secondary-button" onClick={handleSync}>
                <Loader2 size={18} /> {syncStatus === 'running' ? 'Sinkronisasi...' : 'Sinkronkan Lokal'}
              </button>
            )}
            <button type="button" className="button secondary-button" onClick={handleLogout}>
              <LogOut size={18} /> Keluar
            </button>
          </div>
        </header>

        <section className="app-card profile-card">
          <div className="section-header section-header-center">
            <div>
              <h2>Profil Pegawai</h2>
              <p>Data pegawai yang terdaftar untuk laporan ini.</p>
            </div>
          </div>

          <div className="profile-info">
            <div className="profile-row">
              <span className="profile-label">NIP</span>
              <span className="profile-separator">:</span>
              <span className="profile-value">{formData.nip}</span>
            </div>
            <div className="profile-row">
              <span className="profile-label">Nama</span>
              <span className="profile-separator">:</span>
              <span className="profile-value">{formData.nama}</span>
            </div>
            <div className="profile-row">
              <span className="profile-label">Unit Kerja</span>
              <span className="profile-separator">:</span>
              <span className="profile-value">{formData.unit}</span>
            </div>
          </div>
        </section>

        <form onSubmit={handleSubmit}>
          <section className="app-card">
            <div className="section-header section-header-center">
              <div>
                <h2>Form Laporan Harian</h2>
                <p>Isi jenis dan tanggal laporan terlebih dahulu.</p>
              </div>
            </div>
            <div className="form-grid">
              <label className="field-group">
                <span>Jenis</span>
                <select
                  className="select"
                  value={formData.jenis}
                  onChange={(e) => setFormData({ ...formData, jenis: e.target.value })}
                >
                  <option value="WFH">WFH</option>
                  <option value="WFA">WFA</option>
                  <option value="WFO">WFO</option>
                </select>
              </label>
              <label className="field-group">
                <span>Tanggal</span>
                <input
                  type="date"
                  className="input"
                  value={formData.tanggal}
                  onChange={(e) => setFormData({ ...formData, tanggal: e.target.value })}
                />
              </label>
            </div>
          </section>

          <section className="app-card">
            <div className="section-header section-header-center">
              <div>
                <h2>Detail Laporan</h2>
                <p>Isi judul, tugas, dokumentasi, dan tanda tangan di bawah.</p>
              </div>
            </div>

            <div className="form-grid">
              <label className="field-group full-width">
                <span>Judul Laporan</span>
                <input
                  type="text"
                  className="input"
                  placeholder="Judul laporan"
                  value={formData.judulLaporan}
                  onChange={(e) => setFormData({ ...formData, judulLaporan: e.target.value })}
                />
              </label>
            </div>

            <div className="task-list">
              {tugasList.map((task, index) => (
                <div key={index} className="task-row">
                  <div className="task-fields">
                    <div className="task-row-top">
                      <label className="field-group">
                        <span>Tugas {index + 1}</span>
                        <input
                          className="input"
                          placeholder="Deskripsi tugas"
                          value={task.tugas}
                          onChange={(e) => updateTask(index, 'tugas', e.target.value)}
                        />
                      </label>
                      <label className="field-group">
                        <span>Hasil</span>
                        <input
                          className="input"
                          placeholder="Hasil pekerjaan"
                          value={task.hasil}
                          onChange={(e) => updateTask(index, 'hasil', e.target.value)}
                        />
                      </label>
                    </div>
                    <label className="field-group">
                      <span>Dokumentasi</span>
                      <input
                        type="file"
                        className="input file-input"
                        accept="image/*"
                        onChange={(e) => handleTaskDocumentationChange(index, e)}
                      />
                      {task.documentation?.name && (
                        <div className="document-name">{task.documentation.name}</div>
                      )}
                    </label>
                  </div>
                  <button type="button" className="icon-button" onClick={() => removeTask(index)} title="Hapus tugas">
                    <Trash2 size={18} />
                  </button>
                </div>
              ))}
            </div>

            <button type="button" className="button secondary-button add-task-button" onClick={addTask}>
              <PlusCircle size={16} /> Tambah Tugas
            </button>

            {formError && <div className="error-message">{formError}</div>}

            <div className="signature-card">
              <div className="field-group">
                <span>Tanda Tangan Digital</span>
                <canvas
                  ref={canvasRef}
                  className="signature-canvas"
                  width="760"
                  height="160"
                  onPointerDown={drawStart}
                  onPointerMove={drawMove}
                  onPointerUp={drawEnd}
                  onPointerLeave={drawEnd}
                />
              </div>
              <div className="signature-actions">
                <button type="button" className="button secondary-button" onClick={clearSignature}>
                  <X size={16} /> Bersihkan
                </button>
                {signatureDataUrl && (
                  <div className="signature-preview">
                    <CheckCircle size={16} /> Tanda tangan siap.
                  </div>
                )}
              </div>
            </div>

            <div className="form-actions-row">
              <button type="button" className="button secondary-button" onClick={handleExportWord}>
                Ekspor Word
              </button>
              <button type="submit" className="button primary-button submit-button">
                Simpan Laporan
              </button>
            </div>
          </section>
        </form>

        <section className="app-card">
          <div className="section-header-centered">
            <div>
              <h2>Riwayat Laporan</h2>
              <p>Daftar ringkasan laporan harian. Klik preview untuk melihat detail laporan.</p>
            </div>
          </div>

          {reports.length === 0 ? (
            <div className="empty-state">
              Belum ada laporan. Isi form lalu klik Simpan Laporan.
            </div>
          ) : (
            <div className="report-summary-table-wrapper">
              <table className="report-summary-table">
                <thead>
                  <tr>
                    <th>No</th>
                    <th>Judul Laporan</th>
                    <th>Jenis</th>
                    <th>Tanggal</th>
                    <th>Jumlah Tugas</th>
                    <th>Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  {reports.map((report, index) => (
                    <tr key={report.id}>
                      <td>{index + 1}</td>
                      <td>{report.judulLaporan || '-'}</td>
                      <td>{report.jenis}</td>
                      <td>{new Date(report.tanggal || report.createdAt).toLocaleDateString('id-ID')}</td>
                      <td>{report.items?.length || 0}</td>
                      <td className="report-summary-actions">
                        <button type="button" className="button secondary-button" onClick={() => openReportPreview(report)}>
                          Preview
                        </button>
                        <button type="button" className="button secondary-button" onClick={() => deleteReport(report.id)}>
                          <Trash2 size={16} /> Hapus
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {previewReport && (
            <div className="preview-modal">
              <div className="preview-panel">
                <div className="preview-header no-print">
                  <div>
                    <h3>Preview Laporan</h3>
                    <p>{previewReport.nama} • {previewReport.nip}</p>
                  </div>
                  <div className="preview-actions">
                    <button type="button" className="button secondary-button" onClick={() => handleExportWord(previewReport)}>
                      Ekspor Word
                    </button>
                    <button type="button" className="button secondary-button" onClick={() => window.print()}>
                      <Printer size={16} /> Cetak
                    </button>
                    <button type="button" className="button secondary-button" onClick={closeReportPreview}>
                      Tutup
                    </button>
                  </div>
                </div>
                <div className="preview-body printable-report">
                  <article className="report-card a4-page">
                    <div className="report-header">
                      <div>
                        <h3 className="report-title">{previewReport.judulLaporan || 'Laporan Kerja'}</h3>
                        <p className="report-meta">{previewReport.jenis} | {previewReport.nip} | {previewReport.unit}</p>
                        <p className="report-meta">{new Date(previewReport.createdAt).toLocaleString('id-ID')}</p>
                      </div>
                    </div>
                    <div className="report-table-wrapper">
                      <table className="report-detail-table">
                        <thead>
                          <tr>
                            <th>No</th>
                            <th>Uraian Tugas</th>
                            <th>Hasil</th>
                            <th>Dokumentasi</th>
                          </tr>
                        </thead>
                        <tbody>
                          {previewReport.items?.map((item, index) => (
                            <tr key={index}>
                              <td>{index + 1}</td>
                              <td>{item.tugas}</td>
                              <td>{item.hasil}</td>
                              <td>
                                {item.documentation?.dataUrl ? (
                                  <img src={item.documentation.dataUrl} alt={`Dokumentasi tugas ${index + 1}`} />
                                ) : (
                                  '-'
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    {previewReport.signature && (
                      <div className="signature-preview-card">
                        <img src={previewReport.signature} alt="Tanda tangan" />
                      </div>
                    )}
                  </article>
                </div>
              </div>
            </div>
          )}
        </section>
      </div>
    </div>
  );
}
