import { MonitorDown, Share, X } from 'lucide-react';
import { useState } from 'react';
import { usePWAInstall } from '../hooks/usePWAInstall';

interface PWAInstallButtonProps {
  variant?: 'header' | 'sidebar';
}

export function PWAInstallButton({ variant = 'header' }: PWAInstallButtonProps) {
  const { isInstallable, isInstalled, isIOS, install } = usePWAInstall();
  const [showGuide, setShowGuide] = useState(false);

  // Hide install button once running as an installed standalone desktop app
  if (isInstalled) {
    return null;
  }

  const handleClick = async () => {
    if (isInstallable) {
      await install();
      return;
    }
    setShowGuide(true);
  };

  const buttonClasses =
    variant === 'sidebar'
      ? 'w-full inline-flex items-center justify-center gap-1.5 px-3 py-2 text-xs font-semibold text-white bg-gerep-blue rounded-lg hover:bg-gerep-blue-dark transition-colors whitespace-nowrap'
      : 'inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold text-gerep-blue bg-sky-50 border border-sky-200 rounded-lg hover:bg-sky-100 transition-colors whitespace-nowrap shrink-0';

  return (
    <>
      <button
        type="button"
        onClick={handleClick}
        className={buttonClasses}
        title="Installer GerepBase en tant qu’application bureau autonome"
      >
        <MonitorDown className="w-3.5 h-3.5 shrink-0" />
        <span>Installer l’app bureau</span>
      </button>

      {showGuide && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-xs p-4"
          role="dialog"
          aria-modal="true"
        >
          <div className="w-full max-w-md rounded-xl bg-white border border-slate-200 p-6 shadow-xl space-y-4">
            <div className="flex items-start justify-between gap-3">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-lg bg-gerep-blue/10 text-gerep-blue flex items-center justify-center shrink-0">
                  <MonitorDown className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-semibold text-slate-900">
                    {isIOS
                      ? 'Installer sur iPhone / iPad'
                      : 'Installer GerepBase sur votre ordinateur'}
                  </h3>
                  <p className="text-xs text-slate-500">
                    Application autonome hors ligne · Stockage local
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowGuide(false)}
                className="p-1 text-slate-400 hover:text-slate-700 rounded-md"
                aria-label="Fermer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {isIOS ? (
              <ol className="space-y-2.5 text-xs text-slate-700 leading-relaxed bg-slate-50 border border-slate-200 rounded-lg p-4">
                <li className="flex items-start gap-2">
                  <span className="font-mono font-semibold text-gerep-blue">1.</span>
                  <span>
                    Appuyez sur le bouton{' '}
                    <strong className="inline-flex items-center gap-1">
                      Partager <Share className="w-3.5 h-3.5" />
                    </strong>{' '}
                    dans la barre d’outils Safari.
                  </span>
                </li>
                <li className="flex items-start gap-2">
                  <span className="font-mono font-semibold text-gerep-blue">2.</span>
                  <span>
                    Faites défiler puis sélectionnez{' '}
                    <strong>Sur l’écran d’accueil</strong>.
                  </span>
                </li>
              </ol>
            ) : (
              <div className="space-y-3 text-xs text-slate-600 leading-relaxed">
                <p>
                  GerepBase est configuré comme une <strong>application bureau installable (PWA)</strong>{' '}
                  avec stockage local persistant (IndexedDB) et mode hors ligne complet.
                </p>
                <ol className="space-y-2 bg-slate-50 border border-slate-200 rounded-lg p-3.5 text-slate-700">
                  <li>
                    <span className="font-semibold text-gerep-blue">1.</span> Si vous
                    consultez l’aperçu intégré, ouvrez d’abord l’application dans un{' '}
                    <strong>onglet dédié du navigateur</strong> (Chrome ou Edge).
                  </li>
                  <li>
                    <span className="font-semibold text-gerep-blue">2.</span> Cliquez sur
                    l’icône <strong>Installer</strong> dans la barre d’adresse ou dans le
                    menu du navigateur (<strong>Applications → Installer GerepBase</strong>).
                  </li>
                  <li>
                    <span className="font-semibold text-gerep-blue">3.</span> Une fois
                    installée, l’application s’ouvre dans sa propre fenêtre bureau et
                    conserve toutes vos données localement sur cet ordinateur.
                  </li>
                </ol>
              </div>
            )}

            <div className="flex justify-end pt-1">
              <button
                type="button"
                onClick={() => setShowGuide(false)}
                className="px-4 py-2 text-xs font-semibold text-white bg-gerep-blue rounded-lg hover:bg-gerep-blue-dark transition-colors"
              >
                Compris
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
