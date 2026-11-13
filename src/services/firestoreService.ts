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
      (snapshot) => {
        const list: any[] = [];
        snapshot.forEach((d) => {
          const raw = d.data();
          const normalizedId = raw.id !== undefined && raw.id !== null
            ? raw.id
            : (isNaN(Number(d.id)) ? d.id : Number(d.id));
          list.push({
            ...raw,
            id: normalizedId,
            _docId: d.id,
          });
        });
        list.sort((a, b) => (Number(a.id) || 0) - (Number(b.id) || 0));
        onData(list);
      },
      (error) => {
        console.warn('Firestore snapshot error:', error);
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
    const { _docId, ...cleanItem } = item;
    const docId = String(_docId || cleanItem.id || Date.now());
    const docRef = doc(db, REGISTRATIONS_COLLECTION, docId);
    try {
      await setDoc(docRef, {
        ...cleanItem,
        id: cleanItem.id ?? (isNaN(Number(docId)) ? docId : Number(docId)),
        updatedAt: new Date().toISOString(),
      }, { merge: true });
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `${REGISTRATIONS_COLLECTION}/${docId}`);
    }
  },

  async deleteRegistration(docId: any): Promise<void> {
    const targetDocId = String(docId);
    const docRef = doc(db, REGISTRATIONS_COLLECTION, targetDocId);
    try {
      await deleteDoc(docRef);
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, `${REGISTRATIONS_COLLECTION}/${targetDocId}`);
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
            console.error('Error initializing settings:', err);
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
        } catch {
          // Handled and logged
        }
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
        snapshot.forEach((d) => {
          const raw = d.data();
          list.push({
            ...raw,
            id: raw.id !== undefined && raw.id !== null ? raw.id : d.id,
            _docId: d.id,
          });
        });
        onData(list);
      },
      (error) => {
        console.warn('Firestore rombels listener error:', error);
        try {
          handleFirestoreError(error, OperationType.LIST, ROMBELS_COLLECTION);
        } catch {
          // Handled and logged
        }
      }
    );
  },

  async saveRombel(rombel: any): Promise<void> {
    const { _docId, ...cleanRombel } = rombel;
    const docId = String(_docId || cleanRombel.id || cleanRombel.name || Date.now());
    const docRef = doc(db, ROMBELS_COLLECTION, docId);
    try {
      await setDoc(docRef, { ...cleanRombel, id: cleanRombel.id ?? docId, updatedAt: new Date().toISOString() }, { merge: true });
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `${ROMBELS_COLLECTION}/${docId}`);
    }
  },

  async deleteRombel(docId: string): Promise<void> {
    const targetDocId = String(docId);
    const docRef = doc(db, ROMBELS_COLLECTION, targetDocId);
    try {
      await deleteDoc(docRef);
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, `${ROMBELS_COLLECTION}/${targetDocId}`);
    }
  }
};
