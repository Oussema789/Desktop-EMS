import JSZip from 'jszip';
import { buildSamplePdfBlob, INITIAL_DATABASE } from '../data/seedData';
import { AtlasDatabaseState, DocumentAttachmentMeta } from '../types/database';
import { buildOrganizedAttachmentRecords, GEREP_DOCS_FOLDER } from './docPaths';

const DB_NAME = 'AtlasBaseLocalDB';
const DB_VERSION = 1;
const STORE_STATE = 'state';
const STORE_FILES = 'files';
const STATE_KEY = 'main';
const DIR_HANDLE_KEY = 'desktop_dir_handle';
const FALLBACK_LS_KEY = 'atlasbase_state_v1';

export interface LocalStorageDiagnostics {
  persisted: boolean;
  usageBytes: number;
  quotaBytes: number;
  backend: 'indexeddb' | 'localstorage';
  fileCount: number;
}

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') {
      reject(new Error('IndexedDB unavailable'));
      return;
    }
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_STATE)) {
        db.createObjectStore(STORE_STATE);
      }
      if (!db.objectStoreNames.contains(STORE_FILES)) {
        db.createObjectStore(STORE_FILES);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

const EMPTY_DATABASE: AtlasDatabaseState = {
  experts: [],
  subcontractors: [],
  partners: [],
};

const MIGRATION_FLAG_KEY = 'atlasbase_empty_default_migrated_v1';

function isUntouchedSeed(state: AtlasDatabaseState): boolean {
  return (
    state.experts.length === INITIAL_DATABASE.experts.length &&
    state.subcontractors.length === INITIAL_DATABASE.subcontractors.length &&
    state.partners.length === INITIAL_DATABASE.partners.length &&
    state.experts.every(
      (e, idx) =>
        e.id === INITIAL_DATABASE.experts[idx]?.id &&
        e.updatedAt === INITIAL_DATABASE.experts[idx]?.updatedAt
    )
  );
}

export async function requestPersistentStorage(): Promise<boolean> {
  try {
    if (typeof navigator !== 'undefined' && navigator.storage?.persist) {
      const already = (await navigator.storage.persisted?.()) ?? false;
      if (already) return true;
      return await navigator.storage.persist();
    }
  } catch {
    // ignore
  }
  return false;
}

export async function getLocalStorageDiagnostics(): Promise<LocalStorageDiagnostics> {
  let persisted = false;
  let usageBytes = 0;
  let quotaBytes = 0;
  let backend: 'indexeddb' | 'localstorage' = 'localstorage';
  let fileCount = 0;

  try {
    if (navigator.storage?.persisted) {
      persisted = await navigator.storage.persisted();
    }
    if (navigator.storage?.estimate) {
      const est = await navigator.storage.estimate();
      usageBytes = est.usage ?? 0;
      quotaBytes = est.quota ?? 0;
    }
  } catch {
    // ignore
  }

  try {
    const db = await openDatabase();
    backend = 'indexeddb';
    fileCount = await new Promise<number>((resolve, reject) => {
      const tx = db.transaction(STORE_FILES, 'readonly');
      const req = tx.objectStore(STORE_FILES).count();
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  } catch {
    backend = 'localstorage';
  }

  return { persisted, usageBytes, quotaBytes, backend, fileCount };
}

export async function loadDatabaseState(): Promise<AtlasDatabaseState> {
  // Proactively request persistent local storage for desktop reliability
  void requestPersistentStorage();

  try {
    const db = await openDatabase();
    const existing = await new Promise<AtlasDatabaseState | undefined>((resolve, reject) => {
      const tx = db.transaction(STORE_STATE, 'readonly');
      const req = tx.objectStore(STORE_STATE).get(STATE_KEY);
      req.onsuccess = () => resolve(req.result as AtlasDatabaseState | undefined);
      req.onerror = () => reject(req.error);
    });

    if (existing && Array.isArray(existing.experts)) {
      if (!localStorage.getItem(MIGRATION_FLAG_KEY)) {
        localStorage.setItem(MIGRATION_FLAG_KEY, '1');
        if (isUntouchedSeed(existing)) {
          await clearAllDatabase();
          return EMPTY_DATABASE;
        }
      }
      return existing;
    }

    localStorage.setItem(MIGRATION_FLAG_KEY, '1');
    await saveDatabaseState(EMPTY_DATABASE);
    return EMPTY_DATABASE;
  } catch {
    const raw = localStorage.getItem(FALLBACK_LS_KEY);
    if (raw) {
      try {
        const parsed = JSON.parse(raw) as AtlasDatabaseState;
        if (!localStorage.getItem(MIGRATION_FLAG_KEY)) {
          localStorage.setItem(MIGRATION_FLAG_KEY, '1');
          if (isUntouchedSeed(parsed)) {
            localStorage.setItem(FALLBACK_LS_KEY, JSON.stringify(EMPTY_DATABASE));
            return EMPTY_DATABASE;
          }
        }
        return parsed;
      } catch {
        // ignore
      }
    }
    localStorage.setItem(MIGRATION_FLAG_KEY, '1');
    localStorage.setItem(FALLBACK_LS_KEY, JSON.stringify(EMPTY_DATABASE));
    return EMPTY_DATABASE;
  }
}

export async function saveDatabaseState(state: AtlasDatabaseState): Promise<void> {
  try {
    const db = await openDatabase();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_STATE, 'readwrite');
      tx.objectStore(STORE_STATE).put(state, STATE_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    localStorage.setItem(FALLBACK_LS_KEY, JSON.stringify(state));
  }
}

export async function saveAttachmentBlob(attachmentId: string, blob: Blob): Promise<void> {
  try {
    const db = await openDatabase();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_FILES, 'readwrite');
      tx.objectStore(STORE_FILES).put(blob, attachmentId);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    // Fallback ignored if IndexedDB is disabled
  }
}

export async function getAttachmentBlob(attachmentId: string): Promise<Blob | null> {
  try {
    const db = await openDatabase();
    const result = await new Promise<Blob | undefined>((resolve, reject) => {
      const tx = db.transaction(STORE_FILES, 'readonly');
      const req = tx.objectStore(STORE_FILES).get(attachmentId);
      req.onsuccess = () => resolve(req.result as Blob | undefined);
      req.onerror = () => reject(req.error);
    });
    return result ?? null;
  } catch {
    return null;
  }
}

export async function deleteAttachmentBlob(attachmentId: string): Promise<void> {
  try {
    const db = await openDatabase();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_FILES, 'readwrite');
      tx.objectStore(STORE_FILES).delete(attachmentId);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    // ignore
  }
}

export async function getAllAttachmentEntries(): Promise<Array<{ id: string; blob: Blob }>> {
  try {
    const db = await openDatabase();
    return await new Promise<Array<{ id: string; blob: Blob }>>((resolve, reject) => {
      const tx = db.transaction(STORE_FILES, 'readonly');
      const store = tx.objectStore(STORE_FILES);
      const keysReq = store.getAllKeys();
      const valsReq = store.getAll();
      tx.oncomplete = () => {
        const keys = keysReq.result as string[];
        const vals = valsReq.result as Blob[];
        const entries: Array<{ id: string; blob: Blob }> = [];
        for (let i = 0; i < keys.length; i++) {
          if (vals[i] instanceof Blob) {
            entries.push({ id: String(keys[i]), blob: vals[i] });
          }
        }
        resolve(entries);
      };
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    return [];
  }
}

export async function saveDirectoryHandle(handle: FileSystemDirectoryHandle): Promise<void> {
  try {
    const db = await openDatabase();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_STATE, 'readwrite');
      tx.objectStore(STORE_STATE).put(handle, DIR_HANDLE_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    // ignore
  }
}

export async function loadDirectoryHandle(): Promise<FileSystemDirectoryHandle | null> {
  try {
    const db = await openDatabase();
    const result = await new Promise<FileSystemDirectoryHandle | undefined>(
      (resolve, reject) => {
        const tx = db.transaction(STORE_STATE, 'readonly');
        const req = tx.objectStore(STORE_STATE).get(DIR_HANDLE_KEY);
        req.onsuccess = () =>
          resolve(req.result as FileSystemDirectoryHandle | undefined);
        req.onerror = () => reject(req.error);
      }
    );
    return result ?? null;
  } catch {
    return null;
  }
}

export async function clearDirectoryHandle(): Promise<void> {
  try {
    const db = await openDatabase();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE_STATE, 'readwrite');
      tx.objectStore(STORE_STATE).delete(DIR_HANDLE_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    // ignore
  }
}

export function collectAllAttachmentMeta(state: AtlasDatabaseState): DocumentAttachmentMeta[] {
  return [
    ...state.experts.flatMap((e) => e.attachments),
    ...state.subcontractors.flatMap((s) => s.adminDocuments ?? []),
    ...state.partners.flatMap((p) => p.adminDocuments ?? []),
  ];
}

export async function exportFullLocalBackupZip(
  state: AtlasDatabaseState
): Promise<{ blob: Blob; fileCount: number }> {
  const zip = new JSZip();
  const organized = buildOrganizedAttachmentRecords(state);
  const recordById = new Map(organized.map((r) => [r.id, r]));
  const entries = await getAllAttachmentEntries();
  const entryById = new Map(entries.map((e) => [e.id, e.blob]));

  const attachmentIndex: Record<string, string> = {};
  let fileCount = 0;

  for (const record of organized) {
    const blob = entryById.get(record.id);
    if (blob) {
      zip.file(record.relativePath, blob);
      attachmentIndex[record.relativePath] = record.id;
      fileCount += 1;
    }
  }

  // Include any unreferenced blobs safely inside GEREP-EMS-Docs/Autres
  for (const entry of entries) {
    if (!recordById.has(entry.id)) {
      const fallbackPath = `${GEREP_DOCS_FOLDER}/Autres/${entry.id}.bin`;
      zip.file(fallbackPath, entry.blob);
      attachmentIndex[fallbackPath] = entry.id;
      fileCount += 1;
    }
  }

  zip.file(
    'gerepbase-backup.json',
    JSON.stringify(
      {
        app: 'GerepBase',
        schemaVersion: 2,
        docsFolder: GEREP_DOCS_FOLDER,
        exportedAt: new Date().toISOString(),
        attachmentIndex,
        state,
      },
      null,
      2
    )
  );

  const blob = await zip.generateAsync({ type: 'blob' });
  return { blob, fileCount };
}

export async function importFullLocalBackupZip(
  file: File
): Promise<{ state: AtlasDatabaseState; restoredFiles: number }> {
  const buffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const jsonEntry = zip.file('gerepbase-backup.json');
  if (!jsonEntry) {
    throw new Error('Archive invalide : gerepbase-backup.json introuvable.');
  }

  const parsed = JSON.parse(await jsonEntry.async('text')) as {
    state?: AtlasDatabaseState;
    attachmentIndex?: Record<string, string>;
  };
  const nextState = parsed.state;
  if (
    !nextState ||
    !Array.isArray(nextState.experts) ||
    !Array.isArray(nextState.subcontractors) ||
    !Array.isArray(nextState.partners)
  ) {
    throw new Error('Structure de sauvegarde JSON invalide.');
  }

  const metas = collectAllAttachmentMeta(nextState);
  const mimeById = new Map(metas.map((m) => [m.id, m.mimeType]));
  const attachmentIndex = parsed.attachmentIndex ?? {};

  await saveDatabaseState(nextState);

  let restoredFiles = 0;
  const attachmentEntries = Object.entries(zip.files).filter(
    ([path, item]) =>
      !item.dir &&
      (path.startsWith(`${GEREP_DOCS_FOLDER}/`) || path.startsWith('attachments/'))
  );

  for (const [path, item] of attachmentEntries) {
    let attachmentId = attachmentIndex[path];
    if (!attachmentId && path.startsWith('attachments/')) {
      const fileName = path.replace(/^attachments\//, '');
      const sepIdx = fileName.indexOf('__');
      attachmentId = sepIdx > 0 ? fileName.slice(0, sepIdx) : fileName;
    }
    if (!attachmentId) continue;

    const rawBytes = await item.async('uint8array');
    const copy = new Uint8Array(rawBytes.byteLength);
    copy.set(rawBytes);
    const mimeType = mimeById.get(attachmentId) || 'application/octet-stream';
    await saveAttachmentBlob(attachmentId, new Blob([copy.buffer], { type: mimeType }));
    restoredFiles += 1;
  }

  return { state: nextState, restoredFiles };
}

async function seedSampleAttachmentBlobs(state: AtlasDatabaseState): Promise<void> {
  for (const expert of state.experts) {
    const fullName = `${expert.firstName} ${expert.lastName}`;
    for (const att of expert.attachments) {
      const blob = buildSamplePdfBlob(fullName, att);
      await saveAttachmentBlob(att.id, blob);
    }
  }
}

export async function resetToSampleDatabase(): Promise<AtlasDatabaseState> {
  const db = await openDatabase();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction([STORE_STATE, STORE_FILES], 'readwrite');
    tx.objectStore(STORE_STATE).delete(STATE_KEY);
    tx.objectStore(STORE_FILES).clear();
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  await saveDatabaseState(INITIAL_DATABASE);
  await seedSampleAttachmentBlobs(INITIAL_DATABASE);
  return INITIAL_DATABASE;
}

export async function clearAllDatabase(): Promise<AtlasDatabaseState> {
  const empty: AtlasDatabaseState = { experts: [], subcontractors: [], partners: [] };
  try {
    const db = await openDatabase();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction([STORE_STATE, STORE_FILES], 'readwrite');
      tx.objectStore(STORE_STATE).put(empty, STATE_KEY);
      tx.objectStore(STORE_FILES).clear();
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    localStorage.setItem(FALLBACK_LS_KEY, JSON.stringify(empty));
  }
  return empty;
}
