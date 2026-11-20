import {
  CheckCircle2,
  Database,
  Download,
  FolderOpen,
  FolderSync,
  HardDrive,
  MonitorCheck,
  RefreshCw,
  ShieldCheck,
  Trash2,
  Unlink,
  Upload,
  Wifi,
  WifiOff,
} from 'lucide-react';
import { ChangeEvent, useEffect, useState } from 'react';
import { useOnlineStatus } from '../hooks/useOnlineStatus';
import { usePWAInstall } from '../hooks/usePWAInstall';
import {
  getLinkedFolderStatus,
  isDesktopFolderApiSupported,
  isRunningInIframe,
  loadStateFromDesktopFolder,
  pickDesktopFolder,
  requestLinkedFolderPermission,
  syncStateToDesktopFolder,
  unlinkDesktopFolder,
} from '../storage/desktopFolderSync';
import {
  exportFullLocalBackupZip,
  getLocalStorageDiagnostics,
  importFullLocalBackupZip,
  LocalStorageDiagnostics,
  requestPersistentStorage,
} from '../storage/localDb';
import { AtlasDatabaseState } from '../types/database';
import { PWAInstallButton } from './PWAInstallButton';

interface DesktopStorageViewProps {
  state: AtlasDatabaseState;
  onReplaceState: (next: AtlasDatabaseState) => Promise<void>;
  onResetBlank: () => void;
  onNotify: (text: string, type?: 'success' | 'error') => void;
  onFolderStatusChange?: (folderName: string | null) => void;
}

function formatBytes(bytes: number): string {
  if (!bytes || bytes <= 0) return '0 o';
  const units = ['o', 'Ko', 'Mo', 'Go'];
  const idx = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / Math.pow(1024, idx);
  return `${value.toFixed(idx === 0 ? 0 : 1)} ${units[idx]}`;
}

export function DesktopStorageView({
  state,
  onReplaceState,
  onResetBlank,
  onNotify,
  onFolderStatusChange,
}: DesktopStorageViewProps) {
  const isOnline = useOnlineStatus();
  const { isInstalled } = usePWAInstall();
  const [diag, setDiag] = useState<LocalStorageDiagnostics | null>(null);
  const [folderName, setFolderName] = useState<string | null>(null);
  const [folderPermission, setFolderPermission] = useState<string>('prompt');
  const [folderHandle, setFolderHandle] = useState<FileSystemDirectoryHandle | null>(null);
  const [busy, setBusy] = useState(false);

  const folderApiSupported = isDesktopFolderApiSupported();
  const inIframe = isRunningInIframe();

  const refreshDiagnostics = async () => {
    const [nextDiag, folderStatus] = await Promise.all([
      getLocalStorageDiagnostics(),
      getLinkedFolderStatus(),
    ]);
    setDiag(nextDiag);
    setFolderHandle(folderStatus.handle);
    setFolderName(folderStatus.folderName);
    setFolderPermission(folderStatus.permission);
    onFolderStatusChange?.(folderStatus.folderName);
  };

  useEffect(() => {
    void refreshDiagnostics();
  }, [state]);

  const handleRequestPersist = async () => {
    setBusy(true);
    try {
      const granted = await requestPersistentStorage();
      await refreshDiagnostics();
      onNotify(
        granted
          ? 'Stockage persistant garanti par le système.'
          : 'Le navigateur gère déjà ce stockage localement.',
        'success'
      );
    } finally {
      setBusy(false);
    }
  };

  const handleLinkFolder = async () => {
    setBusy(true);
    try {
      const handle = await pickDesktopFolder();
      const loaded = await loadStateFromDesktopFolder(handle).catch(() => null);
      if (loaded && (loaded.state.experts.length > 0 || loaded.state.subcontractors.length > 0 || loaded.state.partners.length > 0) && state.experts.length === 0 && state.subcontractors.length === 0 && state.partners.length === 0) {
        await onReplaceState(loaded.state);
        onNotify(
          `Dossier « ${handle.name} » lié et base locale chargée (${loaded.restoredFiles} fichier(s)).`
        );
      } else {
        const res = await syncStateToDesktopFolder(state, handle);
        onNotify(
          res.synced
            ? `Dossier « ${handle.name} » lié et synchronisé (${res.filesWritten ?? 0} pièce(s) jointe(s)).`
            : `Dossier « ${handle.name} » lié.`
        );
      }
      await refreshDiagnostics();
    } catch (err) {
      if ((err as Error)?.name !== 'AbortError') {
        onNotify(
          (err as Error)?.message || 'Impossible de lier ce dossier local.',
          'error'
        );
      }
    } finally {
      setBusy(false);
    }
  };

  const handleSyncNow = async () => {
    if (!folderHandle) return;
    setBusy(true);
    try {
      const granted = await requestLinkedFolderPermission(folderHandle);
      if (!granted) {
        onNotify('Autorisation d’écriture refusée pour ce dossier.', 'error');
        return;
      }
      const res = await syncStateToDesktopFolder(state, folderHandle);
      await refreshDiagnostics();
      if (res.synced) {
        onNotify(
          `Données enregistrées dans « ${res.folderName} » (${res.filesWritten ?? 0} pièce(s) jointe(s)).`
        );
      }
    } finally {
      setBusy(false);
    }
  };

  const handleLoadFromFolder = async () => {
    if (!folderHandle) return;
    setBusy(true);
    try {
      const loaded = await loadStateFromDesktopFolder(folderHandle);
      if (!loaded) {
        onNotify(
          'Aucun fichier gerepbase-database.json trouvé dans ce dossier.',
          'error'
        );
        return;
      }
      await onReplaceState(loaded.state);
      await refreshDiagnostics();
      onNotify(
        `Base restaurée depuis « ${folderHandle.name} » (${loaded.restoredFiles} pièce(s) jointe(s)).`
      );
    } catch (err) {
      onNotify(
        (err as Error)?.message || 'Erreur de lecture du dossier local.',
        'error'
      );
    } finally {
      setBusy(false);
    }
  };

  const handleUnlinkFolder = async () => {
    await unlinkDesktopFolder();
    await refreshDiagnostics();
    onNotify('Dossier local dissocié. Les données restent dans IndexedDB.');
  };

  const handleExportBackupZip = async () => {
    setBusy(true);
    try {
      const { blob, fileCount } = await exportFullLocalBackupZip(state);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const stamp = new Date().toISOString().slice(0, 10);
      a.href = url;
      a.download = `GerepBase_Sauvegarde_Locale_${stamp}.zip`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      onNotify(
        `Sauvegarde locale (.zip) enregistrée avec ${fileCount} pièce(s) jointe(s).`
      );
    } catch {
      onNotify('Erreur lors de la création de la sauvegarde locale.', 'error');
    } finally {
      setBusy(false);
    }
  };

  const handleImportBackupZip = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    try {
      const { state: nextState, restoredFiles } = await importFullLocalBackupZip(file);
      await onReplaceState(nextState);
      await refreshDiagnostics();
      onNotify(
        `Sauvegarde restaurée : ${nextState.experts.length} experts, ${restoredFiles} pièce(s) jointe(s).`
      );
    } catch (err) {
      onNotify(
        (err as Error)?.message || 'Archive de sauvegarde invalide.',
        'error'
      );
    } finally {
      setBusy(false);
      e.target.value = '';
    }
  };

  const totalRecords =
    state.experts.length + state.subcontractors.length + state.partners.length;

  return (
    <div className="space-y-6">
      {/* Header Card */}
      <div className="bg-white border border-slate-200 rounded-xl p-6 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-5">
        <div className="space-y-1.5 max-w-2xl">
          <div className="flex items-center gap-2 text-xs font-semibold text-gerep-forest">
            <HardDrive className="w-4 h-4" />
            Application Bureau &amp; Stockage 100% Local
          </div>
          <h1 className="text-xl font-bold text-slate-900">
            Stockage local sur cet ordinateur &amp; Mode hors ligne
          </h1>
          <p className="text-xs text-slate-600 leading-relaxed">
            Toutes vos fiches (Experts, Sous-traitants, Partenaires, Projets) et vos
            pièces jointes PDF/images sont conservées localement sur votre appareil sans
            aucun envoi vers un serveur externe.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5 shrink-0">
          {totalRecords > 0 && (
            <button
              type="button"
              onClick={onResetBlank}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg bg-white border border-slate-200 text-xs font-medium text-rose-700 hover:bg-rose-50 transition-colors"
            >
              <Trash2 className="w-3.5 h-3.5" />
              Version vierge (tout vider)
            </button>
          )}
          {isInstalled ? (
            <span className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg bg-emerald-50 border border-emerald-200 text-xs font-semibold text-gerep-forest">
              <MonitorCheck className="w-4 h-4" />
              Application bureau active
            </span>
          ) : (
            <PWAInstallButton variant="header" />
          )}
        </div>
      </div>

      {/* Status Grid */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-1.5">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>Mode d’exécution</span>
            <MonitorCheck className="w-4 h-4 text-gerep-blue" />
          </div>
          <div className="text-sm font-semibold text-slate-900">
            {isInstalled ? 'Fenêtre Bureau autonome' : 'Navigateur (PWA prête)'}
          </div>
          <div className="text-[11px] text-slate-500">
            Cache applicatif hors ligne actif
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-1.5">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>Connectivité</span>
            {isOnline ? (
              <Wifi className="w-4 h-4 text-gerep-forest" />
            ) : (
              <WifiOff className="w-4 h-4 text-amber-600" />
            )}
          </div>
          <div className="text-sm font-semibold text-slate-900">
            {isOnline ? 'En ligne (prêt hors ligne)' : 'Hors ligne actif'}
          </div>
          <div className="text-[11px] text-slate-500">
            Fonctionne sans connexion Internet
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-1.5">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>Base locale</span>
            <Database className="w-4 h-4 text-gerep-blue" />
          </div>
          <div className="text-sm font-semibold text-slate-900 font-mono tabular-nums">
            {totalRecords} fiches · {diag?.fileCount ?? 0} fichiers
          </div>
          <div className="text-[11px] text-slate-500">
            Moteur : {diag?.backend === 'indexeddb' ? 'IndexedDB local' : 'LocalStorage'}
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-1.5">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>Espace &amp; Persistance</span>
            <ShieldCheck className="w-4 h-4 text-gerep-forest" />
          </div>
          <div className="text-sm font-semibold text-slate-900 font-mono tabular-nums">
            {diag ? formatBytes(diag.usageBytes) : '—'}
          </div>
          <div className="flex items-center justify-between gap-2 text-[11px]">
            <span className={diag?.persisted ? 'text-gerep-forest font-medium' : 'text-slate-500'}>
              {diag?.persisted ? 'Stockage persistant activé' : 'Stockage standard'}
            </span>
            {!diag?.persisted && (
              <button
                type="button"
                disabled={busy}
                onClick={handleRequestPersist}
                className="text-gerep-blue hover:underline font-medium"
              >
                Activer
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Desktop Disk Folder Sync & Full Archive Backup */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Local Disk Folder Sync */}
        <div className="bg-white border border-slate-200 rounded-xl p-6 flex flex-col justify-between gap-5">
          <div className="space-y-2.5">
            <div className="flex items-center gap-2 text-xs font-semibold text-gerep-blue">
              <FolderOpen className="w-4 h-4" />
              Synchronisation avec un dossier de votre ordinateur
            </div>
            <h2 className="text-base font-semibold text-slate-900">
              Enregistrer directement sur votre disque dur
            </h2>
            <p className="text-xs text-slate-600 leading-relaxed">
              Liez un dossier local (ex. <span className="font-mono">Documents/GerepBase</span>)
              pour enregistrer automatiquement{' '}
              <span className="font-mono">gerepbase-database.json</span> et classer toutes vos
              pièces jointes dans le répertoire structuré{' '}
              <span className="font-mono font-semibold text-gerep-blue">GEREP-EMS-Docs/</span> :
            </p>
            <div className="rounded-lg bg-slate-50 border border-slate-200 px-3 py-2 font-mono text-[11px] text-slate-600 space-y-0.5">
              <div className="text-slate-900 font-semibold">GEREP-EMS-Docs/</div>
              <div className="pl-3">├── Experts/&lt;Nom_Prenom_Code&gt;/&lt;Catégorie&gt;/</div>
              <div className="pl-3">├── Sous-traitants/&lt;Nom_Code&gt;/</div>
              <div className="pl-3">└── Partenaires/&lt;Nom_Code&gt;/</div>
            </div>

            {folderName && (
              <div className="mt-2 p-3 rounded-lg bg-slate-50 border border-slate-200 flex items-center justify-between gap-3">
                <div className="min-w-0">
                  <div className="text-xs font-semibold text-slate-900 flex items-center gap-1.5">
                    <CheckCircle2 className="w-3.5 h-3.5 text-gerep-forest shrink-0" />
                    <span className="truncate">Dossier lié : {folderName}</span>
                  </div>
                  <div className="text-[11px] text-slate-500 mt-0.5">
                    {folderPermission === 'granted'
                      ? 'Synchronisation automatique active'
                      : 'Cliquez sur « Synchroniser » pour réactiver l’écriture après redémarrage'}
                  </div>
                </div>
                <button
                  type="button"
                  onClick={handleUnlinkFolder}
                  className="p-1.5 text-slate-400 hover:text-rose-600 rounded-md"
                  title="Dissocier ce dossier"
                >
                  <Unlink className="w-4 h-4" />
                </button>
              </div>
            )}

            {!folderApiSupported && (
              <div className="p-3 rounded-lg bg-amber-50 border border-amber-200 text-xs text-amber-900 leading-relaxed">
                {inIframe
                  ? 'L’accès direct à un dossier du disque est bloqué dans un aperçu intégré (iframe). Installez GerepBase comme application bureau ou ouvrez-le dans un onglet dédié Chrome/Edge pour activer la liaison directe de dossier.'
                  : 'Votre navigateur actuel ne prend pas en charge la liaison directe de dossier. Utilisez Chrome ou Edge sur ordinateur, ou la sauvegarde complète (.zip) ci-contre.'}
              </div>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <button
              type="button"
              disabled={busy || !folderApiSupported}
              onClick={handleLinkFolder}
              className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-gerep-blue rounded-lg hover:bg-gerep-blue-dark transition-colors disabled:opacity-50"
            >
              <FolderOpen className="w-3.5 h-3.5" />
              {folderName ? 'Changer de dossier' : 'Choisir un dossier local'}
            </button>

            {folderName && folderApiSupported && (
              <>
                <button
                  type="button"
                  disabled={busy}
                  onClick={handleSyncNow}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-gerep-blue bg-white border border-slate-200 rounded-lg hover:bg-slate-50 disabled:opacity-50"
                >
                  <FolderSync className="w-3.5 h-3.5 text-gerep-forest" />
                  Synchroniser
                </button>
                <button
                  type="button"
                  disabled={busy}
                  onClick={handleLoadFromFolder}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 disabled:opacity-50"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  Charger depuis le dossier
                </button>
              </>
            )}
          </div>
        </div>

        {/* Full Local Backup / Restore ZIP */}
        <div className="bg-white border border-slate-200 rounded-xl p-6 flex flex-col justify-between gap-5">
          <div className="space-y-2.5">
            <div className="flex items-center gap-2 text-xs font-semibold text-gerep-forest">
              <Download className="w-4 h-4" />
              Sauvegarde &amp; Transfert de poste (.zip)
            </div>
            <h2 className="text-base font-semibold text-slate-900">
              Archive complète de la base locale
            </h2>
            <p className="text-xs text-slate-600 leading-relaxed">
              Exportez l’intégralité de votre base (fiches, projets et toutes les pièces
              jointes binaires) dans un fichier <span className="font-mono">.zip</span> à
              conserver sur votre disque dur ou clé USB, et restaurez-le à tout moment sur
              n’importe quel poste.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <button
              type="button"
              disabled={busy}
              onClick={handleExportBackupZip}
              className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-gerep-forest rounded-lg hover:bg-emerald-800 transition-colors disabled:opacity-50"
            >
              <Download className="w-3.5 h-3.5" />
              Créer une sauvegarde (.zip)
            </button>

            <label className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-slate-700 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 cursor-pointer transition-colors">
              <Upload className="w-3.5 h-3.5 text-gerep-blue" />
              <span>Restaurer une sauvegarde (.zip)</span>
              <input
                type="file"
                accept=".zip,application/zip"
                className="hidden"
                disabled={busy}
                onChange={handleImportBackupZip}
              />
            </label>
          </div>
        </div>
      </div>
    </div>
  );
}
