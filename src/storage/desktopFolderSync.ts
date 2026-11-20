import { AtlasDatabaseState } from '../types/database';
import { buildOrganizedAttachmentRecords, GEREP_DOCS_FOLDER } from './docPaths';
import {
  clearDirectoryHandle,
  collectAllAttachmentMeta,
  getAllAttachmentEntries,
  loadDirectoryHandle,
  saveAttachmentBlob,
  saveDatabaseState,
  saveDirectoryHandle,
} from './localDb';

type PermissionMode = 'read' | 'readwrite';
type PermissionStateValue = 'granted' | 'denied' | 'prompt';

interface DirectoryHandleWithPermissions extends FileSystemDirectoryHandle {
  queryPermission?: (descriptor?: { mode?: PermissionMode }) => Promise<PermissionStateValue>;
  requestPermission?: (descriptor?: { mode?: PermissionMode }) => Promise<PermissionStateValue>;
}

interface WindowWithDirectoryPicker extends Window {
  showDirectoryPicker?: (options?: {
    id?: string;
    mode?: PermissionMode;
  }) => Promise<FileSystemDirectoryHandle>;
}

const DATA_FILENAME = 'gerepbase-database.json';
const INDEX_FILENAME = '.attachment-index.json';
const LEGACY_ATTACHMENTS_DIR = 'attachments';

export function isRunningInIframe(): boolean {
  try {
    return window.self !== window.top;
  } catch {
    return true;
  }
}

export function isDesktopFolderApiSupported(): boolean {
  if (typeof window === 'undefined') return false;
  const win = window as WindowWithDirectoryPicker;
  return typeof win.showDirectoryPicker === 'function' && !isRunningInIframe();
}

export async function getLinkedFolderStatus(): Promise<{
  handle: FileSystemDirectoryHandle | null;
  folderName: string | null;
  permission: PermissionStateValue | 'unsupported';
}> {
  const handle = (await loadDirectoryHandle()) as DirectoryHandleWithPermissions | null;
  if (!handle) {
    return { handle: null, folderName: null, permission: 'prompt' };
  }
  try {
    const perm = (await handle.queryPermission?.({ mode: 'readwrite' })) ?? 'prompt';
    return { handle, folderName: handle.name, permission: perm };
  } catch {
    return { handle, folderName: handle.name, permission: 'prompt' };
  }
}

export async function requestLinkedFolderPermission(
  handle: FileSystemDirectoryHandle
): Promise<boolean> {
  const ext = handle as DirectoryHandleWithPermissions;
  try {
    const current = await ext.queryPermission?.({ mode: 'readwrite' });
    if (current === 'granted') return true;
    const requested = await ext.requestPermission?.({ mode: 'readwrite' });
    return requested === 'granted';
  } catch {
    return false;
  }
}

export async function pickDesktopFolder(): Promise<FileSystemDirectoryHandle> {
  const win = window as WindowWithDirectoryPicker;
  if (!win.showDirectoryPicker) {
    throw new Error(
      'L’accès direct aux dossiers locaux nécessite Chrome/Edge hors iframe (ou l’application bureau installée).'
    );
  }
  const handle = await win.showDirectoryPicker({
    id: 'gerepbase-local-data',
    mode: 'readwrite',
  });
  await saveDirectoryHandle(handle);
  return handle;
}

export async function unlinkDesktopFolder(): Promise<void> {
  await clearDirectoryHandle();
}

async function writeFileAtRelativePath(
  rootDir: FileSystemDirectoryHandle,
  relativePath: string,
  data: Blob | string
): Promise<void> {
  const segments = relativePath.split('/').filter(Boolean);
  const fileName = segments.pop();
  if (!fileName) return;

  let currentDir = rootDir;
  for (const segment of segments) {
    currentDir = await currentDir.getDirectoryHandle(segment, { create: true });
  }

  const fileHandle = await currentDir.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(data);
  await writable.close();
}

async function readAllFilesRecursively(
  dir: FileSystemDirectoryHandle,
  prefix: string,
  out: Array<{ relativePath: string; fileHandle: FileSystemFileHandle }>
): Promise<void> {
  for await (const [name, entry] of dir.entries()) {
    const nextPath = prefix ? `${prefix}/${name}` : name;
    if (entry.kind === 'file') {
      out.push({ relativePath: nextPath, fileHandle: entry });
    } else if (entry.kind === 'directory') {
      await readAllFilesRecursively(entry, nextPath, out);
    }
  }
}

export async function syncStateToDesktopFolder(
  state: AtlasDatabaseState,
  explicitHandle?: FileSystemDirectoryHandle | null
): Promise<{ synced: boolean; folderName?: string; filesWritten?: number }> {
  const handle = (explicitHandle ??
    (await loadDirectoryHandle())) as DirectoryHandleWithPermissions | null;
  if (!handle) return { synced: false };

  try {
    const perm = await handle.queryPermission?.({ mode: 'readwrite' });
    if (perm !== 'granted') {
      return { synced: false, folderName: handle.name };
    }

    const organized = buildOrganizedAttachmentRecords(state);
    const recordById = new Map(organized.map((r) => [r.id, r]));
    const entries = await getAllAttachmentEntries();
    const entryById = new Map(entries.map((e) => [e.id, e.blob]));

    const attachmentIndex: Record<string, string> = {};
    let filesWritten = 0;

    // Organize documents into GEREP-EMS-Docs/Experts/..., GEREP-EMS-Docs/Sous-traitants/..., GEREP-EMS-Docs/Partenaires/...
    for (const record of organized) {
      const blob = entryById.get(record.id);
      if (blob) {
        await writeFileAtRelativePath(handle, record.relativePath, blob);
        attachmentIndex[record.relativePath] = record.id;
        filesWritten += 1;
      }
    }

    for (const entry of entries) {
      if (!recordById.has(entry.id)) {
        const fallbackPath = `${GEREP_DOCS_FOLDER}/Autres/${entry.id}.bin`;
        await writeFileAtRelativePath(handle, fallbackPath, entry.blob);
        attachmentIndex[fallbackPath] = entry.id;
        filesWritten += 1;
      }
    }

    // Write attachment path index inside GEREP-EMS-Docs/
    const docsDir = await handle.getDirectoryHandle(GEREP_DOCS_FOLDER, { create: true });
    const indexFile = await docsDir.getFileHandle(INDEX_FILENAME, { create: true });
    const indexWritable = await indexFile.createWritable();
    await indexWritable.write(JSON.stringify(attachmentIndex, null, 2));
    await indexWritable.close();

    // Write main JSON database
    const dataFile = await handle.getFileHandle(DATA_FILENAME, { create: true });
    const writable = await dataFile.createWritable();
    await writable.write(
      JSON.stringify(
        {
          app: 'GerepBase',
          schemaVersion: 2,
          docsFolder: GEREP_DOCS_FOLDER,
          updatedAt: new Date().toISOString(),
          attachmentIndex,
          state,
        },
        null,
        2
      )
    );
    await writable.close();

    return { synced: true, folderName: handle.name, filesWritten };
  } catch {
    return { synced: false, folderName: handle.name };
  }
}

export async function loadStateFromDesktopFolder(
  explicitHandle?: FileSystemDirectoryHandle | null
): Promise<{ state: AtlasDatabaseState; restoredFiles: number } | null> {
  const handle = (explicitHandle ??
    (await loadDirectoryHandle())) as DirectoryHandleWithPermissions | null;
  if (!handle) return null;

  const granted = await requestLinkedFolderPermission(handle);
  if (!granted) {
    throw new Error('Permission d’accès au dossier refusée.');
  }

  let dataFile: FileSystemFileHandle;
  try {
    dataFile = await handle.getFileHandle(DATA_FILENAME, { create: false });
  } catch {
    return null;
  }

  const file = await dataFile.getFile();
  const parsed = JSON.parse(await file.text()) as {
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
    throw new Error('Le fichier gerepbase-database.json du dossier est invalide.');
  }

  const metas = collectAllAttachmentMeta(nextState);
  const mimeById = new Map(metas.map((m) => [m.id, m.mimeType]));

  await saveDatabaseState(nextState);

  let attachmentIndex: Record<string, string> = parsed.attachmentIndex ?? {};
  let restoredFiles = 0;

  // Prefer reading organized GEREP-EMS-Docs/ tree
  try {
    const docsDir = await handle.getDirectoryHandle(GEREP_DOCS_FOLDER, { create: false });
    try {
      const idxHandle = await docsDir.getFileHandle(INDEX_FILENAME, { create: false });
      const idxFile = await idxHandle.getFile();
      const idxParsed = JSON.parse(await idxFile.text()) as Record<string, string>;
      attachmentIndex = { ...attachmentIndex, ...idxParsed };
    } catch {
      // Index file optional if already in database json
    }

    const expectedRecords = buildOrganizedAttachmentRecords(nextState);
    const idByExpectedPath = new Map(
      expectedRecords.map((r) => [r.relativePath, r.id])
    );

    const diskFiles: Array<{ relativePath: string; fileHandle: FileSystemFileHandle }> =
      [];
    await readAllFilesRecursively(docsDir, GEREP_DOCS_FOLDER, diskFiles);

    for (const item of diskFiles) {
      if (item.relativePath.endsWith(`/${INDEX_FILENAME}`)) continue;
      const attachmentId =
        attachmentIndex[item.relativePath] ?? idByExpectedPath.get(item.relativePath);
      if (!attachmentId) continue;

      const fileObj = await item.fileHandle.getFile();
      const mimeType =
        mimeById.get(attachmentId) || fileObj.type || 'application/octet-stream';
      const buffer = await fileObj.arrayBuffer();
      await saveAttachmentBlob(attachmentId, new Blob([buffer], { type: mimeType }));
      restoredFiles += 1;
    }
  } catch {
    // Fallback to legacy attachments/ directory if present
    try {
      const attDir = await handle.getDirectoryHandle(LEGACY_ATTACHMENTS_DIR, {
        create: false,
      });
      for await (const [name, entry] of attDir.entries()) {
        if (entry.kind === 'file') {
          const fileObj = await entry.getFile();
          const sepIdx = name.indexOf('__');
          const attachmentId = sepIdx > 0 ? name.slice(0, sepIdx) : name;
          const mimeType =
            mimeById.get(attachmentId) || fileObj.type || 'application/octet-stream';
          const buffer = await fileObj.arrayBuffer();
          await saveAttachmentBlob(attachmentId, new Blob([buffer], { type: mimeType }));
          restoredFiles += 1;
        }
      }
    } catch {
      // No attachments folder yet
    }
  }

  return { state: nextState, restoredFiles };
}
