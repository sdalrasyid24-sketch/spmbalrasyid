import {
  collection,
  doc,
  setDoc,
  deleteDoc,
  getDocs,
  onSnapshot,
  query,
  orderBy,
  serverTimestamp,
} from 'firebase/firestore';
import { db } from '../lib/firebase.ts';
import { handleFirestoreError, OperationType } from '../lib/firestoreErrors.ts';

const REGISTRATIONS_COLLECTION = 'registrations';
const ROMBELS_COLLECTION = 'rombels';
const SETTINGS_COLLECTION = 'school_settings';
const SETTINGS_DOC_ID = 'config';

export const initialDefaultPendaftar = [
  { 
    id: 1,
    regNumber: 'SPMB-2024-001',
    name: 'Ahmad Rafiq',
    nickname: 'Rafiq',
    gender: 'Laki-laki',
    birthPlace: 'Pekanbaru',
    birthDate: '2015-05-12', 
    religion: 'Islam',
    nationality: 'WNI',
    childOrder: '1',
    siblingsCount: '2',
    familyStatus: 'Anak Kandung',
    dailyLanguage: 'Indonesia',
    siblingsDetail: { kandung: '1', sebapak: '0', seibu: '0', tiri: '0', angkat: '0' },
    livingWith: 'Orangtua',
    distanceToSchool: '2',
    physical: { weight: '25', height: '120', bloodType: 'O', chronicDisease: '' },
    address: 'Jl. Sudirman No. 10, Pekanbaru',
    phone: '08123456789',
    previousSchool: { name: 'TK IT Al-Rasyid', address: 'Jl. Sudirman' },
    parent: {
      fatherName: 'Budi Santoso',
      fatherEdu: 'S1',
      fatherJob: 'PNS',
      motherName: 'Siti Aminah',
      motherEdu: 'S1',
      motherJob: 'Guru',
      parentAddress: 'Jl. Sudirman No. 10',
      income: 'Rp 3.000.000 - Rp 5.000.000'
    },
    guardian: { hasGuardian: false, relation: '', edu: '', job: '', phone: '', address: '' },
    emergencyContact: { name: 'Joko', relation: 'Paman', phone: '08129876543', address: 'Jl. Ahmad Yani' },
    specialNeeds: '',
    achievements: '',
    latinReading: 'Lancar membaca',
    quranReading: 'Iqro\' 4-6',
    transport: 'Tidak',
    suggestions: '',
    docs: { akta: true, kk: true, materai: true, foto: true },
    status: 'Diterima' 
  },
  { 
    id: 2,
    regNumber: 'SPMB-2024-002',
    name: 'Nisa Sabyan',
    nickname: 'Nisa',
    gender: 'Perempuan',
    birthPlace: 'Padang',
    birthDate: '2015-08-20', 
    religion: 'Islam',
    nationality: 'WNI',
    childOrder: '2',
    siblingsCount: '2',
    familyStatus: 'Anak Kandung',
    dailyLanguage: 'Indonesia',
    siblingsDetail: { kandung: '1', sebapak: '0', seibu: '0', tiri: '0', angkat: '0' },
    livingWith: 'Orangtua',
    distanceToSchool: '5',
    physical: { weight: '22', height: '115', bloodType: 'A', chronicDisease: '' },
    address: 'Jl. Nangka No. 5, Pekanbaru',
    phone: '08521234567',
    previousSchool: { name: 'TK Pertiwi', address: 'Jl. Nangka' },
    parent: {
      fatherName: 'Hasan',
      fatherEdu: 'SMA/SMK',
      fatherJob: 'Wiraswasta',
      motherName: 'Maimunah',
      motherEdu: 'SMA/SMK',
      motherJob: 'Ibu Rumah Tangga',
      parentAddress: 'Jl. Nangka No. 5',
      income: 'Rp 1.000.000 - Rp 3.000.000'
    },
    guardian: { hasGuardian: false, relation: '', edu: '', job: '', phone: '', address: '' },
    emergencyContact: { name: 'Ali', relation: 'Kakek', phone: '082133445566', address: 'Jl. Rambutan' },
    specialNeeds: '',
    achievements: 'Juara 1 Lomba Mewarnai',
    latinReading: 'Belum lancar',
    quranReading: 'Iqro\' 1-3',
    transport: 'Memerlukan',
    suggestions: '',
    docs: { akta: true, kk: false, materai: false, foto: true },
    status: 'Menunggu Keputusan' 
  }
];

export const initialDefaultSettings = {
  tahunAjaran: '2024/2025',
  gelombang: 'Gelombang 1',
  status: 'Buka',
  kuotaMaksimal: '50',
  tanggalMulai: '2024-01-15',
  tanggalAkhir: '2024-06-30',
  usiaMinimal: '6',
  biayaPendaftaran: 'Gratis',
  ketuaPanitia: 'Ustadz Ridwan, S.Pd',
  kontakPanitia: '0812-3456-7890',
  catatanSeleksi: 'Pendaftaran ulang wajib membawa berkas fisik asli untuk diverifikasi oleh panitia SPMB SD Al-Rasyid.'
};

export const firestoreService = {
  // --- REGISTRATIONS ---
  subscribeRegistrations(onData: (data: any[]) => void, onError?: (err: any) => void) {
    const colRef = collection(db, REGISTRATIONS_COLLECTION);
    return onSnapshot(
      colRef,
      async (snapshot) => {
        if (snapshot.empty) {
          // Auto-seed initial demo registrations to Firestore if currently empty
          try {
            for (const item of initialDefaultPendaftar) {
              const docRef = doc(db, REGISTRATIONS_COLLECTION, String(item.id));
              await setDoc(docRef, { ...item, _syncedAt: new Date().toISOString() });
            }
          } catch (seedErr) {
            console.error('Initial seeding error:', seedErr);
          }
          onData(initialDefaultPendaftar);
        } else {
          const list: any[] = [];
          snapshot.forEach((d) => {
            list.push({ ...d.data(), id: isNaN(Number(d.id)) ? d.id : Number(d.id) });
          });
          // Sort by id or regNumber
          list.sort((a, b) => (Number(a.id) || 0) - (Number(b.id) || 0));
          onData(list);
        }
      },
      (error) => {
        console.warn('Firestore snapshot error, falling back to local memory:', error);
        if (onError) onError(error);
        try {
          handleFirestoreError(error, OperationType.LIST, REGISTRATIONS_COLLECTION);
        } catch {
          // Handled and logged
        }
      }
    );
  },

  async saveRegistration(item: any): Promise<void> {
    const id = String(item.id || Date.now());
    const docRef = doc(db, REGISTRATIONS_COLLECTION, id);
    try {
      await setDoc(docRef, {
        ...item,
        id: isNaN(Number(id)) ? id : Number(id),
        updatedAt: new Date().toISOString(),
      }, { merge: true });
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `${REGISTRATIONS_COLLECTION}/${id}`);
    }
  },

  async deleteRegistration(id: any): Promise<void> {
    const docRef = doc(db, REGISTRATIONS_COLLECTION, String(id));
    try {
      await deleteDoc(docRef);
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, `${REGISTRATIONS_COLLECTION}/${id}`);
    }
  },

  // --- SETTINGS ---
  subscribeSettings(onData: (data: any) => void) {
    const docRef = doc(db, SETTINGS_COLLECTION, SETTINGS_DOC_ID);
    return onSnapshot(
      docRef,
      async (snapshot) => {
        if (!snapshot.exists()) {
          try {
            await setDoc(docRef, { ...initialDefaultSettings, updatedAt: new Date().toISOString() });
          } catch (err) {
            console.error('Error seeding settings:', err);
          }
          onData(initialDefaultSettings);
        } else {
          onData(snapshot.data());
        }
      },
      (error) => {
        console.warn('Firestore settings listener error:', error);
        try {
          handleFirestoreError(error, OperationType.GET, `${SETTINGS_COLLECTION}/${SETTINGS_DOC_ID}`);
        } catch {}
      }
    );
  },

  async saveSettings(settings: any): Promise<void> {
    const docRef = doc(db, SETTINGS_COLLECTION, SETTINGS_DOC_ID);
    try {
      await setDoc(docRef, { ...settings, updatedAt: new Date().toISOString() }, { merge: true });
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `${SETTINGS_COLLECTION}/${SETTINGS_DOC_ID}`);
    }
  },

  // --- ROMBELS ---
  subscribeRombels(onData: (data: any[]) => void) {
    const colRef = collection(db, ROMBELS_COLLECTION);
    return onSnapshot(
      colRef,
      (snapshot) => {
        const list: any[] = [];
        snapshot.forEach((d) => list.push({ ...d.data(), id: d.id }));
        onData(list);
      },
      (error) => {
        console.warn('Firestore rombels listener error:', error);
        try {
          handleFirestoreError(error, OperationType.LIST, ROMBELS_COLLECTION);
        } catch {}
      }
    );
  },

  async saveRombel(rombel: any): Promise<void> {
    const id = String(rombel.id || rombel.name || Date.now());
    const docRef = doc(db, ROMBELS_COLLECTION, id);
    try {
      await setDoc(docRef, { ...rombel, id, updatedAt: new Date().toISOString() }, { merge: true });
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `${ROMBELS_COLLECTION}/${id}`);
    }
  },

  async deleteRombel(id: string): Promise<void> {
    const docRef = doc(db, ROMBELS_COLLECTION, id);
    try {
      await deleteDoc(docRef);
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, `${ROMBELS_COLLECTION}/${id}`);
    }
  }
};
